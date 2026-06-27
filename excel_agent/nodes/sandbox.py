"""
nodes/sandbox.py — 沙盒执行节点（双轨执行架构）

职责：
    1. Track 1: 运行命中缓存的独立子表代码 (提取 cache 里的数据)。
    2. Track 2: 运行 LLM 生成的新代码 (提取 missed_subtables 的数据)。
    3. 合并两次执行的结果，返回给后端的 restore / quality 节点校验。
    4. 将执行成功的数据回写进 CacheState 的第四阶段 (extracted_data)，供 SA 节点打包。
"""
import os
import re
import math
import json
import builtins
import sys
import traceback
from typing import Any, Dict, List, Optional

import openpyxl

from Excel_Agent.excel_agent.state import AgentState, CacheState


_SAFE_BUILTINS = {
    name: getattr(builtins, name)
    for name in [
        "int", "float", "str", "bool", "list", "dict", "tuple", "set",
        "bytes", "bytearray", "type", "object",
        "len", "range", "enumerate", "zip", "map", "filter", "sorted",
        "reversed", "sum", "min", "max", "abs", "round", "divmod", "pow",
        "print", "repr", "format", "chr", "ord",
        "isinstance", "issubclass", "hasattr", "getattr", "setattr",
        "callable", "iter", "next", "all", "any", "open",
        "Exception", "ValueError", "TypeError", "KeyError", "IndexError",
        "StopIteration", "RuntimeError",
        "True", "False", "None",
    ]
    if hasattr(builtins, name)
}


def _build_merge_map(ws) -> Dict[tuple, Any]:
    """合并单元格填充图：{(row, col): 包含边界和值的字典}"""
    m: Dict[tuple, Any] = {}
    for rng in ws.merged_cells.ranges:
        val = ws.cell(row=rng.min_row, column=rng.min_col).value
        for r in range(rng.min_row, rng.max_row + 1):
            for c in range(rng.min_col, rng.max_col + 1):
                m[(r, c)] = {
                    "value": val,
                    "min_row": rng.min_row,
                    "max_row": rng.max_row,
                    "min_col": rng.min_col,
                    "max_col": rng.max_col
                }
    return m


def _run_with_timeout(fn, timeout_sec: int = 30):
    try:
        import signal
        def _handler(signum, frame):
            raise TimeoutError(f"代码执行超时（>{timeout_sec}s）")
        old = signal.signal(signal.SIGALRM, _handler)
        signal.alarm(timeout_sec)
        try:
            return fn()
        finally:
            signal.alarm(0)
            signal.signal(signal.SIGALRM, old)
    except (ImportError, AttributeError):
        # 兼容 Windows 系统无 SIGALRM 的情况
        return fn()


def _validate_result(result: Any) -> Dict[str, List[List[Any]]]:
    """校验 extract() 返回值格式。"""
    if not isinstance(result, dict):
        raise ValueError(f"extract() 必须返回 dict，实际返回 {type(result).__name__}")

    validated = {}
    for key, table_data in result.items():
        if not isinstance(table_data, list):
            raise ValueError(f"字典的 value 必须是 list，键 '{key}' 对应的是 {type(table_data).__name__}")
        for i, row in enumerate(table_data):
            if not isinstance(row, (list, tuple)):
                raise ValueError(f"键 '{key}' 的第 {i} 行不是 list/tuple，而是 {type(row).__name__}")
        validated[key] = [list(r) for r in table_data]
    return validated


def sandbox_node(state: AgentState) -> dict:
    config = state.get("config", {})
    excel_path = config.get("excel_path") or state.get("excel_path", "")
    sheet_name = config.get("sheet_name", "")

    # —————————————————— KV Agent 抽取开发
    # 获取 kv_list
    kv_list: List[str] = config.get("kv_list", [])
    is_kv_model = False
    if kv_list: is_kv_model = True
    if is_kv_model:
        kv_state = state.get("kv_state")
        # KV Agent 生成的代码
        kv_generated_code = kv_state.get("generated_code")
        # 沙箱报错
        kv_sandbox_error = None
        bboxes = kv_state.get("sheet_structure").get("bboxes")


    cache_state: CacheState = state.get("cache", {})
    entries = cache_state.get("entries", {})

    # 兼容处理配置字段
    target_configs = config.get("subtable_configs") or config.get("target_columns")
    hints = config.get("hints")

    try:
        wb = openpyxl.load_workbook(excel_path, data_only=True)
        ws = wb[sheet_name]
    except Exception as e:
        return {
            "raw_result": None,
            "sandbox_error": f"沙盒无法打开 Excel 文件：{e}",
        }

    merged_map = _build_merge_map(ws)

    final_raw_result: Dict[str, List[List[Any]]] = {}
    errors: List[str] = []

    # ==========================================================
    # ── 双轨执行 Track 1: 运行命中缓存的独立子表代码 ──
    # ==========================================================
    for title, entry in entries.items():
        if entry.get("cache_hit") and entry.get("code"):
            try:
                # 为每个缓存代码提供独立的纯净命名空间
                namespace = {
                    "__builtins__": _SAFE_BUILTINS,
                    "re": re, "math": math, "json": json
                }
                exec(compile(entry["code"], f"<cached_{title}>", "exec"), namespace)
                extract_fn = namespace.get("extract")

                if callable(extract_fn):
                    sub_res = _run_with_timeout(lambda: extract_fn(ws, merged_map, entry.get("start_row"), entry.get("start_col")))

                    # 鲁棒性兼容：SA节点拆分出的代码可能返回二维数组，也可能返回字典 {title: 二维数组}
                    if isinstance(sub_res, dict):
                        sub_res = sub_res.get(title) or (list(sub_res.values())[0] if sub_res else [])

                    final_raw_result[title] = sub_res
                    entry["extracted_data"] = sub_res  # 填入 Stage 4，供后续 SA 打包
                else:
                    errors.append(f"缓存代码 [{title}] 中未找到 extract 函数。")
            except Exception:
                errors.append(f"缓存代码 [{title}] 执行失败:\n{traceback.format_exc()}")


    # ==========================================================
    # ── 双轨执行 Track 2: 运行 LLM 新生成的代码 (针对 missed_subtables) ──
    # ==========================================================
    # new_code_list = state.get("generated_code", [])
    # ———————— KV Agent 抽取开发
    new_code_list = kv_generated_code if is_kv_model else state.get("generated_code", [])


    for new_code in new_code_list:

        if is_kv_model:
            namespace = {
                "__builtins__": _SAFE_BUILTINS,
                "re": re, "math": math, "json": json,
                "sheet_structure": state.get("kv_state", {}).get("sheet_structure"),
                "kv_list": kv_list
            }
        else:
            missed_subtables = cache_state.get("missed_subtables", [])

            # if new_code and missed_subtables:
            namespace = {
                "__builtins__": _SAFE_BUILTINS,
                "re": re, "math": math, "json": json,
                "sheet_structure": state.get("sheet_structure"),
                "subtable_titles": missed_subtables,  # LLM 只需知道它该负责哪些表
                "target_columns": target_configs,
                "hints": hints,
            }

        try:
            # 1. 剥离 markdown 代码块标记
            md_pattern = r'``' + r'`(?:python)?(.*?)``' + r'`'
            match = re.search(md_pattern, new_code, re.DOTALL)

            if match:
                new_code = match.group(1)

            # 2. 去除首尾多余的空白符以及可能引起解析报错的外层引号
            new_code = new_code.strip('"\n\r\' ')

            # 3. 处理 \n 在 JSON 传输中变成真实换行符导致的 SynataxError
            new_code = new_code.replace("'\n'", r"'\n'").replace("'\r'", r"'\r'")


            exec(compile(new_code, "<llm_generated>", "exec"), namespace)
            extract_fn = namespace.get("extract")

            if callable(extract_fn):
                if is_kv_model:
                    raw_llm = _run_with_timeout(lambda: extract_fn(ws, merged_map, bboxes))
                else:
                    raw_llm = _run_with_timeout(lambda: extract_fn(ws, merged_map))
                llm_dict = _validate_result(raw_llm) if not is_kv_model else raw_llm
                if is_kv_model:
                    for title, data in llm_dict.items():
                        final_raw_result[title] = data
                    # 不再 early return，继续走后续流程让 restore/quality 设置 result



                # 将 LLM 跑出来的数据合并到最终结果中
                for title, data in llm_dict.items():
                    final_raw_result[title] = data
                    # 同步更新到 cache entries 里，准备给 SA 节点存库用
                    if title in entries:
                        entries[title]["extracted_data"] = data
            else:
                errors.append("新生成的代码中未找到 extract(ws, merged_map) 函数。")
        except TimeoutError as e:
            errors.append(f"新生成的代码执行超时: {str(e)}")
        except Exception:
            errors.append(f"新生成的代码执行报错:\n{traceback.format_exc()}")


    # ==========================================================
    # ── 3. 结果合并与校验 ──
    # ==========================================================
    sandbox_error_str = "\n".join(errors) if errors else None

    # 如果所有表的结果都是空字典或空列表，视同执行失败，交给 Quality 节点扣分打回
    if not final_raw_result or all(not v for v in final_raw_result.values()):
        if not sandbox_error_str:
            sandbox_error_str = "提取失败：所有执行轨道均未返回有效数据 (空字典或全空列表)。请检查硬编码行号或逻辑是否错误。"

    kv_result_payload = {
        "raw_result": final_raw_result,
        "sandbox_error": sandbox_error_str,
        "cache": cache_state
    }
    if is_kv_model and final_raw_result:
        kv_result_payload["result"] = final_raw_result
        kv_result_payload["final_output"] = final_raw_result
    return kv_result_payload
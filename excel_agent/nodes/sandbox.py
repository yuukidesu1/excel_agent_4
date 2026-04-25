"""
nodes/sandbox.py — 沙盒执行节点（双轨执行架构）

职责：
    1. Track 1: 运行命中缓存的独立子表代码 (提取 cache 里的数据)。
    2. Track 2: 运行 LLM 生成的新代码 (提取 missed_subtables 的数据)。
    3. 合并两次执行的结果，返回给后端的 restore / quality 节点校验。
    4. 将执行成功的数据回写进 CacheState 的第四阶段 (extracted_data)，供 SA 节点打包。

KV 模式支持：
    - 当 config.extract_type == "kv" 时，执行 KV 抽取代码
    - KV 代码签名：extract(ws, merged_map, kv_list) -> dict
    - 返回结果存储到 kv_result 字段
"""

import re
import math
import json
import builtins
import traceback
from typing import Any, Dict, List, Optional

import openpyxl

from excel_agent.state import AgentState, CacheState


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
    """合并单元格填充图：{(row, col): 左上角值}"""
    m: Dict[tuple, Any] = {}
    for rng in ws.merged_cells.ranges:
        val = ws.cell(row=rng.min_row, column=rng.min_col).value
        for r in range(rng.min_row, rng.max_row + 1):
            for c in range(rng.min_col, rng.max_col + 1):
                m[(r, c)] = val
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
    """校验 extract() 返回值格式（表格模式）。"""
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


def _validate_kv_result(result: Any) -> Dict[str, str]:
    """校验 KV 模式 extract() 返回值格式。"""
    if not isinstance(result, dict):
        raise ValueError(f"KV extract() 必须返回 dict，实际返回 {type(result).__name__}")

    validated = {}
    for key, value in result.items():
        if not isinstance(key, str):
            raise ValueError(f"KV 结果的键必须是 str，键 '{key}' 类型是 {type(key).__name__}")
        validated[key] = str(value) if value is not None else ""
    return validated


def sandbox_node(state: AgentState) -> dict:
    config = state.get("config", {})
    excel_path = config.get("excel_path") or state.get("excel_path", "")
    sheet_name = config.get("sheet_name", "")

    # 判断抽取模式
    extract_type = config.get("extract_type", "table")  # "table" | "kv"

    cache_state: CacheState = state.get("cache", {})
    entries = cache_state.get("entries", {})

    # 获取配置
    subtable_configs = config.get("subtable_configs")
    hints = config.get("hints")

    # KV 模式专用
    kv_list = config.get("kv_list", [])

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
    kv_result: Dict[str, str] = {}
    errors: List[str] = []

    # ==========================================================
    # ── 双轨执行 Track 1: 运行命中缓存的独立子表代码 ──
    # ==========================================================
    if extract_type == "kv":
        # KV 模式：运行缓存代码
        for title, entry in entries.items():
            if entry.get("cache_hit") and entry.get("code"):
                try:
                    namespace = {
                        "__builtins__": _SAFE_BUILTINS,
                        "re": re, "math": math, "json": json,
                        "kv_list": kv_list
                    }
                    exec(compile(entry["code"], f"<cached_{title}>", "exec"), namespace)
                    extract_fn = namespace.get("extract")

                    if callable(extract_fn):
                        sub_res = _run_with_timeout(lambda: extract_fn(ws, merged_map, kv_list))
                        kv_result = _validate_kv_result(sub_res)
                    else:
                        errors.append(f"缓存代码 [{title}] 中未找到 extract 函数。")
                except Exception:
                    errors.append(f"缓存代码 [{title}] 执行失败:\n{traceback.format_exc()}")
    else:
        # 表格模式：运行命中缓存的独立子表代码
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

                        # 鲁棒性兼容：SA 节点拆分出的代码可能返回二维数组，也可能返回字典 {title: 二维数组}
                        if isinstance(sub_res, dict):
                            sub_res = sub_res.get(title) or (list(sub_res.values())[0] if sub_res else [])

                        final_raw_result[title] = sub_res
                        entry["extracted_data"] = sub_res  # 填入 Stage 4，供后续 SA 打包
                    else:
                        errors.append(f"缓存代码 [{title}] 中未找到 extract 函数。")
                except Exception:
                    errors.append(f"缓存代码 [{title}] 执行失败:\n{traceback.format_exc()}")


    # ==========================================================
    # ── 双轨执行 Track 2: 运行 LLM 新生成的代码 ──
    # ==========================================================
    new_code_list = state.get("generated_code", [])

    for new_code in new_code_list:
        if extract_type == "kv":
            # KV 模式：执行新生成的 KV 抽取代码
            if new_code and kv_list:
                namespace = {
                    "__builtins__": _SAFE_BUILTINS,
                    "re": re, "math": math, "json": json,
                    "kv_list": kv_list
                }

                try:
                    exec(compile(new_code, "<llm_generated>", "exec"), namespace)
                    extract_fn = namespace.get("extract")

                    if callable(extract_fn):
                        raw_kv = _run_with_timeout(lambda: extract_fn(ws, merged_map, kv_list))
                        kv_result = _validate_kv_result(raw_kv)
                    else:
                        errors.append("新生成的代码中未找到 extract(ws, merged_map, kv_list) 函数。")
                except TimeoutError as e:
                    errors.append(f"新生成的代码执行超时：{str(e)}")
                except Exception:
                    errors.append(f"新生成的代码执行报错:\n{traceback.format_exc()}")
        else:
            # 表格模式：执行新生成的代码
            missed_subtables = cache_state.get("missed_subtables", [])

            if new_code and missed_subtables:
                namespace = {
                    "__builtins__": _SAFE_BUILTINS,
                    "re": re, "math": math, "json": json,
                    "sheet_structure": state.get("sheet_structure"),
                    "subtable_titles": missed_subtables,
                    "hints": hints,
                }

                try:
                    exec(compile(new_code, "<llm_generated>", "exec"), namespace)
                    extract_fn = namespace.get("extract")

                    if callable(extract_fn):
                        raw_llm = _run_with_timeout(lambda: extract_fn(ws, merged_map))
                        llm_dict = _validate_result(raw_llm)

                        # 将 LLM 跑出来的数据合并到最终结果中
                        for title, data in llm_dict.items():
                            final_raw_result[title] = data
                            if title in entries:
                                entries[title]["extracted_data"] = data
                    else:
                        errors.append("新生成的代码中未找到 extract(ws, merged_map) 函数。")
                except TimeoutError as e:
                    errors.append(f"新生成的代码执行超时：{str(e)}")
                except Exception:
                    errors.append(f"新生成的代码执行报错:\n{traceback.format_exc()}")


    # ==========================================================
    # ── 3. 结果合并与校验 ──
    # ==========================================================
    sandbox_error_str = "\n".join(errors) if errors else None

    # KV 模式返回
    if extract_type == "kv":
        if not kv_result:
            if not sandbox_error_str:
                sandbox_error_str = "KV 提取失败：未返回任何有效数据。请检查 Key 定位逻辑或扫描方向。"
            return {
                "kv_result": {},
                "sandbox_error": sandbox_error_str,
                "cache": cache_state
            }
        return {
            "kv_result": kv_result,
            "sandbox_error": sandbox_error_str,
            "cache": cache_state
        }

    # 表格模式返回
    if not final_raw_result or all(not v for v in final_raw_result.values()):
        if not sandbox_error_str:
            sandbox_error_str = "提取失败：所有执行轨道均未返回有效数据 (空字典或全空列表)。请检查硬编码行号或逻辑是否错误。"

    return {
        "raw_result": final_raw_result,
        "sandbox_error": sandbox_error_str,
        "cache": cache_state
    }

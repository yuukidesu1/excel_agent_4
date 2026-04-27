"""
nodes/kv_sandbox.py — KV 模式专用沙盒执行节点

职责：
    1. 执行 KV 抽取代码：extract(ws, merged_map, kv_list, scope=None) -> Dict[str, str]
    2. 支持多子表同 Keys 场景（不同 scope 执行同一份代码）
    3. 返回结果到 kv_result 字段
    4. 不与其他模式共享状态更新
"""

import re
import math
import json
import builtins
import traceback
from typing import Any, Dict, List, Optional

import openpyxl

from excel_agent.state import AgentState


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


def kv_sandbox_node(state: AgentState) -> dict:
    """KV 模式专用沙盒节点（支持 scope 参数 + 多子表执行）"""
    config = state.get("config", {})
    excel_path = config.get("excel_path", "")
    sheet_name = config.get("sheet_name", "")
    kv_list = config.get("kv_list", [])

    # 检查是否有缓存代码（KV 缓存命中时）
    cache_state = state.get("cache", {})
    if cache_state.get("hit", False) and "code" in cache_state:
        # 使用缓存代码
        code_to_execute = cache_state["code"]
        print("📦 使用缓存的 KV 代码执行...")
    else:
        # 使用新生成的代码
        code_to_execute = None
        print("🔨 执行新生成的 KV 代码...")

    try:
        wb = openpyxl.load_workbook(excel_path, data_only=True)
        ws = wb[sheet_name]
    except Exception as e:
        return {
            "kv_result": {},
            "sandbox_error": f"沙盒无法打开 Excel 文件：{e}",
        }

    merged_map = _build_merge_map(ws)

    kv_result: Dict[str, Any] = {}
    errors: List[str] = []

    # 执行代码（缓存代码或新生成的代码）
    if code_to_execute:
        code_list = [code_to_execute]
    else:
        code_list = state.get("generated_code", [])

    # 检查是否有多子表 KV 场景（entries 中有 scope）
    entries = cache_state.get("entries", {})
    kv_entries = {
        title: entry for title, entry in entries.items()
        if entry.get("extract_mode") == "kv" or entry.get("scope")
    }
    is_multi_subtable = len(kv_entries) > 0

    for code in code_list:
        # if code and kv_list:
        if code:
            # 过滤掉 import 语句（沙盒已预置常用模块）
            filtered_code_lines = []
            for line in code.split('\n'):
                stripped = line.strip()
                # 跳过 import 开头的行
                if stripped.startswith('import ') or stripped.startswith('from '):
                    continue
                filtered_code_lines.append(line)
            filtered_code = '\n'.join(filtered_code_lines)

            namespace = {
                "__builtins__": _SAFE_BUILTINS,
                "re": re, "math": math, "json": json,
                "openpyxl": openpyxl,
                "kv_list": kv_list
            }

            try:
                exec(compile(filtered_code, "<llm_generated>", "exec"), namespace)
                extract_fn = namespace.get("extract")

                if callable(extract_fn):
                    if is_multi_subtable:
                        # 多子表场景：每个子表用不同 scope 执行
                        for title, entry in kv_entries.items():
                            scope = entry.get("scope")
                            raw_kv = _run_with_timeout(lambda s=scope: extract_fn(ws, merged_map, kv_list, s))
                            kv_result[title] = _validate_kv_result(raw_kv)
                    else:
                        # 全局 KV 场景：scope=None
                        raw_kv = _run_with_timeout(lambda: extract_fn(ws, merged_map, kv_list, None))
                        kv_result = _validate_kv_result(raw_kv)
                else:
                    errors.append("新生成的代码中未找到 extract(ws, merged_map, kv_list, scope=None) 函数。")
            except TimeoutError as e:
                errors.append(f"新生成的代码执行超时：{str(e)}")
            except Exception:
                errors.append(f"新生成的代码执行报错:\n{traceback.format_exc()}")

    sandbox_error_str = "\n".join(errors) if errors else None

    if not kv_result:
        if not sandbox_error_str:
            sandbox_error_str = "KV 提取失败：未返回任何有效数据。请检查 Key 定位逻辑或扫描方向。"

    return {
        "kv_result": kv_result,
        "sandbox_error": sandbox_error_str,
    }

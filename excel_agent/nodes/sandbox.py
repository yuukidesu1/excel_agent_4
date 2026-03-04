"""
nodes/sandbox.py — 沙盒执行节点（纯代码）
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
        return fn()


def _validate_result(result: Any) -> Dict[str, List[List[Any]]]:
    """
    校验 extract() 返回值格式。
    空列表不抛异常：让 raw_result=[] 传出，由 quality_node 扣分并触发重试，
    同时把"返回空列表"写入 sandbox_error 告知 LLM，比崩掉报错信息更清晰。
    """
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
    code       = state.get("generated_code", "")
    config     = state.get("config", {})
    excel_path = config.get("excel_path") or state.get("excel_path", "")
    sheet_name = state["sheet_structure"]["sheet_name"]

    if not code:
        return {
            "raw_result":    None,
            "sandbox_error": "generated_code 为空。",
        }

    try:
        wb = openpyxl.load_workbook(excel_path)
        ws = wb[sheet_name]
    except Exception as e:
        return {
            "raw_result":    None,
            "sandbox_error": f"无法打开 Excel 文件：{e}",
        }

    merged_map = _build_merge_map(ws)

    # ── 沙盒命名空间：注入所有 LLM 代码可能引用的变量 ──────────
    # 兼容 config 嵌套结构和平铺结构两种 state 设计
    subtable_titles  = config.get("subtable_titles")  or state.get("subtable_titles", [])
    target_columns  = config.get("target_columns")  or state.get("target_columns")
    hints           = config.get("hints")           or state.get("hints")

    namespace = {
        "__builtins__":  _SAFE_BUILTINS,
        "re":            re,
        "math":          math,
        "json":          json,
        # LLM 代码可直接使用的上下文变量
        "sheet_structure": state["sheet_structure"],
        "subtable_titles":  subtable_titles,
        "target_columns":  target_columns,
        "hints":           hints,
    }

    try:
        exec(compile(code, "<llm_generated>", "exec"), namespace)
    except Exception:
        return {
            "raw_result":    None,
            "sandbox_error": f"代码编译/定义阶段报错：\n{traceback.format_exc()}",
        }

    extract_fn = namespace.get("extract")
    if not callable(extract_fn):
        return {
            "raw_result":    None,
            "sandbox_error": "代码中未找到 extract 函数，请确保定义了 def extract(ws, merged_map)。",
        }

    try:
        raw    = _run_with_timeout(lambda: extract_fn(ws, merged_map))
        result = _validate_result(raw)

        # 检查字典是否完全为空，或者内部所有列表为空
        if not result or all(len(v) == 0 for v in result.values()):
            return {
                "raw_result": {},
                "sandbox_error": (
                    "extract() 返回了空字典或全空列表。"
                    "请检查所有目标子表的硬编码行号列好是否计算正确。"
                ),
            }

        return {"raw_result": result, "sandbox_error": None}

    except TimeoutError as e:
        return {"raw_result": None, "sandbox_error": str(e)}
    except Exception:
        return {
            "raw_result":    None,
            "sandbox_error": f"extract() 执行时报错：\n{traceback.format_exc()}",
        }
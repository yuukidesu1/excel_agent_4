"""
nodes/restore.py — 结果组装节点（纯代码）

职责：
    1. 从 header_map.all_columns 生成表头行（取 child 字段作为列名）
    2. 将 raw_data 中的 None 统一转为 ""（保留空单元格位置）
    3. 拼合 [表头行] + [数据行] → 标准 JSON 二维数组

输出格式示例：
    [
      ["SYSTEM MODULE", "CELL", "Type", "NEW/SWAP/EXIST", "Antenna Type", ...],
      ["Module-A", "Cell-1", "AAU", "NEW", "Kathrein", ...],
      ["Module-A", "Cell-2", "",    "SWAP", "",          ...],   ← SYSTEM MODULE 已填充
      ...
    ]
"""

from typing import Any, List
from excel_agent.state import AgentState


def _normalize(v: Any) -> str:
    """统一单元格值格式"""
    if v is None:
        return ""
    if isinstance(v, int):
        return str(v)
    if isinstance(v, float):
        # 整数值的浮点数去掉小数点（如 1.0 → "1"）
        return str(int(v)) if v == int(v) else str(v)
    return str(v).strip()


def restore_node(state: AgentState) -> dict:
    all_cols   = state["header_map"]["all_columns"]
    header_row = [col.get("child", "") for col in all_cols]

    data_rows = [
        [_normalize(cell) for cell in row]
        for row in (state["raw_data"] or [[]])
    ]

    result = [header_row] + data_rows
    return {"result": result, "final_output": result}

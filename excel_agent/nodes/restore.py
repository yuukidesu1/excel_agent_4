"""
nodes/restore.py — 结果组装 + 列过滤（纯代码）

新增功能：列过滤
  若 state["target_columns"] 不为 None，则只输出指定的列。

  target_columns 格式：
    [
      {"parent": "ANTENNAS", "child": "NEW/SWAP/EXIST"},
      {"parent": None,       "child": "CELL"},
      {"parent": "RF MODULE","child": "TYPE"},
    ]

  匹配规则：
    - child  名称：忽略大小写 + 忽略空格/换行符
    - parent 若指定（非 None）：也做同样的模糊匹配
    - parent 为 None：不限父级，只靠 child 匹配（适合单级表头列）
"""

from typing import Any, List, Optional, Dict
from excel_agent.state import AgentState


def _fmt(v: Any) -> str:
    """统一单元格值格式"""
    if v is None:
        return ""
    if isinstance(v, float) and v == int(v):
        return str(int(v))
    return str(v).strip()


def _norm(s: Optional[str]) -> str:
    """模糊匹配归一化：小写 + 去空格 + 去换行"""
    if s is None:
        return ""
    return s.lower().replace(" ", "").replace("\n", "").replace("\r", "").replace("\t", "")

def _find_col_index(header: List[str], target: Dict) -> int:
    """
    在表头行中找到 target 对应的列索引（-1 表示未找到）。

    target 格式：{"parent": str|None, "child": str}
    header 中的列名可能是：
      - "child"           （单级）
      - "parent||child"   （双级，LLM 可能用这种格式）
    """
    t_child  = _norm(target.get("child", ""))
    t_parent = target.get("parent")

    for i, h in enumerate(header):
        if "||" in h:
            parts = h.split("||", 1)
            col_parent, col_child = parts[0], parts[1]
        else:
            col_parent, col_child = None, h

        if _norm(col_child) != t_child:
            continue
        if t_parent is not None and _norm(col_parent) != _norm(t_parent):
            continue
        return i

    return -1  # 未找到

def restore_node(state: AgentState) -> dict:
    raw_result = state.get("raw_result") or {}  # 🚀 变成了字典
    config = state.get("config", {})
    target_columns = config.get("target_columns") or state.get("target_columns")

    if not raw_result:
        return {"result": {}, "final_output": {}}

    final_res = {}

    # 🚀 遍历字典处理每一个子表
    for title, table_data in raw_result.items():
        if not table_data:
            final_res[title] = []
            continue

        # ── 格式化所有值 ──
        formatted = [[_fmt(cell) for cell in row] for row in table_data]
        header = formatted[0] if formatted else []
        data_rows = formatted[1:] if len(formatted) > 1 else []

        # 列过滤规则
        current_targets = None
        if isinstance(target_columns, dict):
            current_targets = target_columns.get(title)
        elif isinstance(target_columns, list):
            current_targets = target_columns

        if current_targets:
            # 根据当前子表规则进行过滤
            keep_indices: List[int] = []
            keep_labels: List[str] = []

            for target in current_targets:
                idx = _find_col_index(header, target)
                keep_indices.append(idx)
                keep_labels.append(
                    header[idx] if idx >= 0 else f"[未找到]{target.get('child','?')}"
                )
            new_header = keep_labels
            new_data = [
                [row[i] if (0 <= i < len(row)) else "" for i in keep_indices]
                for row in data_rows
            ]
            final_res[title] = [new_header] + new_data
        else:
            final_res[title] = [header] + data_rows

    # 确保 final_output 的名称与你的 AgentState 中定义的名称对齐
    return {"result": final_res, "final_output": final_res}

def _match(col: Dict, target: Dict) -> bool:
    """
    判断 all_columns 中的一列是否匹配用户 target。
    col    = {"parent": str|None, "child": str, "col_idx": int}
    target = {"parent": str|None, "child": str}
    """
    # child 必须匹配
    if _norm(col.get("child")) != _norm(target.get("child")):
        return False
    # parent：target 指定了才校验；target.parent=None 则不限父级
    t_parent = target.get("parent")
    if t_parent is not None:
        return _norm(col.get("parent")) == _norm(t_parent)
    return True


def restore_node_(state: AgentState) -> dict:
    all_cols       = state["header_map"]["all_columns"]
    raw_data       = state.get("raw_data") or [[]]
    target_columns: Optional[List[Dict]] = state["config"].get("target_columns")

    if target_columns:
        # ── 列过滤模式 ─────────────────────────────────────────
        # 为每个 target 找到 all_columns 中对应的位置索引
        keep_indices: List[int]  = []   # -1 表示未找到
        keep_labels:  List[str]  = []

        for target in target_columns:
            found_idx = -1
            for i, col in enumerate(all_cols):
                if _match(col, target):
                    found_idx = i
                    break
            keep_indices.append(found_idx)
            # 列名优先用 all_columns 里的 child，找不到就用 target 里的 child
            if found_idx >= 0:
                keep_labels.append(all_cols[found_idx].get("child", ""))
            else:
                keep_labels.append(f"[未找到]{target.get('child','?')}")

        header_row = keep_labels
        data_rows  = [
            [
                _fmt(row[i]) if (0 <= i < len(row)) else ""
                for i in keep_indices
            ]
            for row in raw_data
        ]
    else:
        # ── 全列模式 ──────────────────────────────────────────
        header_row = [col.get("child", "") for col in all_cols]
        data_rows  = [[_fmt(cell) for cell in row] for row in raw_data]

    result = [header_row] + data_rows
    return {"result": result, "final_output": result}
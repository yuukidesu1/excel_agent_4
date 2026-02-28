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


def restore_node(state: AgentState) -> dict:
    all_cols       = state["header_map"]["all_columns"]
    raw_data       = state.get("raw_data") or [[]]
    target_columns: Optional[List[Dict]] = state.get("target_columns")

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
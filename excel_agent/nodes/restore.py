"""
nodes/restore.py — 结果组装 + 列过滤（纯代码）

新增功能：列过滤
  若 state["config"] 中存在 subtable_configs / target_columns，则只输出指定的列。

  匹配规则：
    - 兼容新版配置：提取 col_headers 或 row_headers 列表中的 "A||B" 字符串
    - 兼容旧版配置：提取 {"parent": "A", "child": "B"} 字典
    - 忽略大小写 + 忽略空格/换行符
"""

from typing import Any, List, Optional, Dict, Union
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


def _parse_target(target: Union[str, Dict]) -> tuple:
    """
    统一解析 target，返回 (parent, child) 元组。
    支持旧版 dict: {"parent": "A", "child": "B"}
    支持新版 str:  "A||B" 或 "B"
    """
    if isinstance(target, dict):
        return target.get("parent"), str(target.get("child", ""))
    elif isinstance(target, str):
        if "||" in target:
            parts = target.split("||", 1)
            return parts[0], parts[1]
        return None, target
    return None, str(target)


def _find_col_index(header: List[str], target: Union[str, Dict]) -> int:
    """
    在表头行中找到 target 对应的列索引（-1 表示未找到）。
    """
    t_parent, t_child_raw = _parse_target(target)
    t_child = _norm(t_child_raw)

    for i, h in enumerate(header):
        if "||" in h:
            parts = h.split("||", 1)
            col_parent, col_child = parts[0], parts[1]
        else:
            col_parent, col_child = None, h

        # 子类必须匹配
        if _norm(col_child) != t_child:
            continue

        # 父类如果指定了，也必须匹配
        if t_parent is not None and _norm(col_parent) != _norm(t_parent):
            continue

        return i

    return -1  # 未找到


def restore_node(state: AgentState) -> dict:
    raw_result = state.get("raw_result") or {}
    config = state.get("config", {})

    target_columns_config = config.get("subtable_configs") or config.get("target_columns")
    subtable_titles = list(target_columns_config.keys())
    if not raw_result:
        return {"result": {}, "final_output": {}}

    final_res = {}

    for title, table_data in raw_result.items():
        target_columns = target_columns_config[title].get("headers")
        set_target_columns = set(target_columns)
        if not table_data:
            final_res[title] = []
            continue

        if title not in subtable_titles:
            raise ValueError(f"Title '{title}' not found in user configs")

        formatted = [[_fmt(cell) for cell in row] for row in table_data]
        headers_tmp = formatted[0] if formatted else []
        data_rows = formatted[1:] if len(formatted) > 1 else []

        header = []
        for header_tmp in headers_tmp:
            if "||" in header_tmp:
                parts = header_tmp.split("||")
                if set(parts).issubset(set_target_columns) and (len(set(parts))) == 1:
                    header.append(parts[0])
                else:
                    header.append(header_tmp)
            else:
                header.append(header_tmp)

        keep_indices: List[int] = []
        keep_labels: List[str] = []

        for target in target_columns:
            idx = _find_col_index(header, target)
            keep_indices.append(idx)

            if idx >= 0:
                keep_labels.append(header[idx])
            else:
                keep_labels.append(target)

        new_header = keep_labels
        new_data = [
            [row[i] if (0 <= i < len(row)) else "" for i in keep_indices]
            for row in data_rows
        ]

        final_res[title] = [new_header] + new_data
    return {"result": final_res, "final_output": final_res}

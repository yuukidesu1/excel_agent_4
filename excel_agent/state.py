"""
state.py — 全局状态定义

输入：
    必填：excel_path, sheet_name, subtable_titles
    可选：target_columns  → 只保留指定列（None = 保留全部）
          hints           → 额外定位提示
"""

from typing import TypedDict, Optional, List, Dict, Any

class ConfigState(TypedDict):
    excel_path: str
    sheet_name: str
    subtable_titles: List[str] # str -> List[str] 以适配多子表抽取
    hints: Optional[str]
    target_columns: Optional[List[Dict[str, Any]]]

class AgentState(TypedDict):

    # ── 用户输入 ───────────────────────────────────────────────
    config: ConfigState

    # ── 列过滤配置（新增）─────────────────────────────────────
    # target_columns: Optional[List[Dict[str, Any]]]
    # None  → 保留子表所有列（全自动）
    # list  → 只保留指定列，格式：
    #   [
    #     {"parent": None,        "child": "CELL"},
    #     {"parent": "ANTENNAS",  "child": "NEW/SWAP/EXIST"},
    #     {"parent": "RF MODULE", "child": "TYPE"},
    #   ]
    # parent=None 时不限父级，只按 child 名称匹配（忽略大小写+空格）

    # ── parse_node 输出 ───────────────────────────────────────────────
    sheet_structure: Optional[Dict[str, Any]]

    # ── code_gen_node 输出 ───────────────────────────────────────────────
    generated_code: Optional[str]

    # ── locate_node 输出──
    header_map: Optional[Dict[str, Any]]

    # ── extract_node 输出──
    raw_data: Optional[List[List[Any]]]

    # ── sandbox_node 输出 ───────────────────────────────────────────────
    raw_result: Optional[Dict[str, List[List[Any]]]]   # Optional[List[List[Any]]] -> Optional[Dict[str, List[List[Any]]]]
    # 代码执行后返回的原始二维数组(含表头行)

    # ── restore_node 输出 ───────────────────────────────────────────────
    result:     Optional[Dict[str, List[List[str]]]]              # Optional[List[List[str]]] -> Optional[Dict[str, List[List[str]]]]
    final_result: Optional[Dict[str, List[List[str]]]]            # Optional[List[List[str]]] -> Optional[Dict[str, List[List[str]]]]

    # ── 质量控制 ───────────────────────────────────────────────
    quality_score: float
    retry_count:   int
    errors:        List[str]
    sandbox_error: Optional[str] # 代码执行异常信息，重试时传给 LLM
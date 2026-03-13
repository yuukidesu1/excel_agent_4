"""
state.py — 全局状态定义

输入：
    必填：excel_path, sheet_name, subtable_titles
    可选：target_columns  → 只保留指定列（None = 保留全部）
          hints           → 额外定位提示
"""

from typing import TypedDict, Optional, List, Dict, Any, Union


class SubtableConfig(TypedDict, total=False):
    """
    单个子表的详细抽取配置
    """
    layout: str     # “仅行” ｜ “仅列” ｜ “交叉”
    col_headers: List[str]
    row_headers: List[str]

class ConfigState(TypedDict):
    excel_path: str
    sheet_name: str
    subtable_titles: List[str] # str -> List[str] 以适配多子表抽取
    hints: Optional[str]
    # target_columns: Optional[Dict[str, List[Dict[str, Any]]]]
    subtable_configs: Optional[Dict[str, Union[SubtableConfig, List[Any]]]]

class SubtableCacheEntry(TypedDict):
    """
    单个子表的缓存流转对象。
    在流水线中按阶段被不同的节点逐步填充。
    """
    # 一阶段，PSA节点纯代码写入
    target_title: str       # 子表名称
    signature: str          # 该子表的 L2 结构指纹
    start_row: int          # 起始行
    start_col: int          # 起始列

    # 增加一个 PSA 识别出的物理布局类型
    layout_type: str

    # 二阶段，由 cache_query 节点查询数据库后写入
    cache_hit: bool         # 是否命中缓存
    cache_level: str        # "l1" | "l2" | "miss"
    l2_row_offset: int      # L2 命中时的行偏移量
    l2_col_offset: int      # L2 命中时的列偏移量

    # 三阶段，命中时从库中读取，或未命中时由 SA 拆解大段代码后写入
    code: Optional[str]     # 专门针对该字表的、可独立运行的纯净 Python 提取代码
    header_map: Optional[Dict[str, Any]]    # 该子表的表头映射关系

    # 四阶段：★ 新增！执行与合并期 (为 Sandbox/Extract 准备) ──
    # 当 Sandbox 节点执行完后，或者在 Cache Query 节点直接执行后，
    # 属于该子表的、最终提取出来的 2D 数组数据存放在这里。
    # 这样后续节点就能把 cached_data 和 LLM 新提取的数据拼到一起！
    extracted_data: Optional[List[List[Any]]]


class CacheState(TypedDict):
    """全局缓存状态管理器"""
    entries: Dict[str, SubtableCacheEntry]

    # 宏观调度
    all_cached: bool
    partial_cached: bool
    missed_subtables: List[str]  # code_gen 节点现在的唯一输入源！

    # 后置分析控制
    analyzer_skipped: bool
    analyzer_skip_reason: Optional[str]

class AgentState(TypedDict):

    # ── 用户输入 ───────────────────────────────────────────────
    config: ConfigState

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

    # ── 缓存系统 ───────────────────────────────────────────────
    cache: CacheState
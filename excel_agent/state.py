"""
state.py — 全局状态定义

输入：
    必填：excel_path, sheet_name, subtable_titles
    可选：target_columns  → 只保留指定列（None = 保留全部）
          hints           → 额外定位提示
"""
from typing import TypedDict, Optional, List, Dict, Any, Union, Annotated, Tuple


class SubtableConfig(TypedDict, total=False):
    """
    单个子表的详细抽取配置
    """
    layout: str     # “仅行” ｜ “仅列” ｜ “交叉”
    headers: List[str]
    keys: List[str]

class ConfigState(TypedDict):
    excel_path: str
    sheet_name: str
    subtable_titles: List[str] # str -> List[str] 以适配多子表抽取
    hints: Optional[str]
    extract_type: Optional[str]
    kv_list: Optional[List[str]]
    target_columns: Optional[Dict[str, List[Dict[str, Any]]]]
    # target_columns: Optional[Dict[str, List[Dict[str, Any]]]]
    subtable_configs: Optional[Dict[str, Union[SubtableConfig, List[Any]]]]


# ============================================================================
# 新增与优化：KV 抽取模式专属的状态容器
# ============================================================================
class KVSheetStructure(TypedDict):
    """KV 模式专属的紧凑 Excel 物理结构视图"""
    max_row: int
    max_col: int
    merged_cells_info: List[str]  # 格式如 "R1C1~R1C6:\"Cell A\""
    sample_cells: List[str]  # 格式如 "R5C1:\"Azmuith\""
    # ★ 新增：为了适配上一版逻辑，记录各个目标区域的包围盒
    bboxes: Dict[str, Dict[str, int]]


class KVCacheState(TypedDict, total=False):
    """
    KV 专属的拓扑指纹缓存状态。
    与 Table 的独立子表级缓存不同，KV 采用全量函数级缓存 (All-or-Nothing)。
    """
    # 一阶段：在 kv_preprocess_node (或随后的结构分析节点) 写入
    signature: str  # 基于 bboxes 和合并单元格特征构建的 16 位 L2 拓扑指纹
    fingerprint_data: Dict[str, Any]  # 记录用于哈希的原始拓扑特征（极度推荐保留，便于日后排查指纹突变原因）

    # 二阶段：在 kv_cache_query_node 查库后写入
    cache_hit: bool  # 是否命中缓存

    # 三阶段：命中时从库中读取写入 (准备传给 Sandbox)
    code: Optional[str]  # 针对该拓扑结构缓存的统一 extract() 提取代码


class KVState(TypedDict, total=False):
    """KV 抽取模式的工作流状态 (隔离于传统 Table 抽取)"""
    # 1. 前处理节点 (kv_preprocess_node) 输出
    sheet_structure: Optional[KVSheetStructure]
    parsed_region_paths: Optional[List[List[str]]]

    # 新增：缓存管理域
    cache: Optional[KVCacheState]

    # 2. 生成节点 (kv_generation_node) 输出
    # 修改：为了兼容 LangGraph 的 Annotated append 操作，保持与表格 generated_code 类型一致
    generated_code: Optional[List[str]]

    # 3. 沙盒执行节点 (kv_executor) 输出
    extracted_data: Optional[Dict[str, Union[str, List[Dict[str, str]]]]]


def _update_kv_state(old_state: Optional[KVState], new_state: KVState) -> KVState:
    if not old_state:
        return new_state
    return {**old_state, **new_state}
# ============================================================================

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
    missed_subtables: Optional[List[Tuple[str, str]]]

    # 后置分析控制
    analyzer_skipped: bool
    analyzer_skip_reason: Optional[str]

class AgentState(TypedDict):

    # ── 用户输入 ───────────────────────────────────────────────
    config: ConfigState

    # ── parse_node 输出 ───────────────────────────────────────────────
    sheet_structure: Optional[Dict[str, Any]]

    # ── code_gen_node 输出 ───────────────────────────────────────────────
    generated_code: Optional[Any]

    # ── extract_node 输出──
    raw_data: Optional[List[List[Any]]]

    # ── sandbox_node 输出 ───────────────────────────────────────────────
    raw_result: Optional[Dict[str, List[List[Any]]]]   # Optional[List[List[Any]]] -> Optional[Dict[str, List[List[Any]]]]
    # 代码执行后返回的原始二维数组(含表头行)

    # ── restore_node 输出 ───────────────────────────────────────────────
    result:     Optional[Dict[str, List[List[str]]]]              # Optional[List[List[str]]] -> Optional[Dict[str, List[List[str]]]]
    final_output: Optional[Any]

    # ── 新增：KV 抽取专用状态域 ──────────────────────────────
    # 将 KV 的所有上下文封存在这里，不污染常规表的 sheet_structure 和 result
    kv_state: Annotated[KVState, _update_kv_state]

    # ── 质量控制 ───────────────────────────────────────────────
    quality_score: float
    retry_count:   int
    errors:        List[str]
    sandbox_error: Optional[str] # 代码执行异常信息，重试时传给 LLM

    # ── 缓存系统 ───────────────────────────────────────────────
    cache: CacheState

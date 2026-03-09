"""
agent.py — LangGraph 图组装 + 对外入口（Agent B：LLM 写代码版）

流程：
  parse_node
      ↓
  code_gen_node   [LLM]  看 parse 数据，写 openpyxl 抽取函数
      ↓
  sandbox_node    [代码]  在受限命名空间执行 LLM 生成的代码
      ↓
  restore_node    [代码]  格式化 + 列过滤
      ↓
  quality_node    [代码]  质量打分
      ↓
  route ──→ 通过 → END ✅
       └──→ 失败（代码报错 or 质量不达标）→ retry → code_gen（携带错误上下文）

重试策略：
  - sandbox_error 不为空（代码执行失败）→ 告知 LLM 具体报错信息
  - quality_score 不达标              → 告知 LLM 质量问题
  - 最多重试 MAX_RETRY 次后强制输出
"""

import sys
import os
from pathlib import Path
from typing import Optional, List, Dict, Any, AsyncGenerator

from langgraph.graph import StateGraph, END

from excel_agent.state            import AgentState
from excel_agent.nodes.parse      import parse_node
from excel_agent.nodes.code_gen   import code_gen_node
from excel_agent.nodes.sandbox    import sandbox_node
from excel_agent.nodes.restore    import restore_node
from excel_agent.nodes.quality    import quality_node, route, QUALITY_THRESHOLD
from excel_agent.memory import get_reference_code, save_reference_code
from excel_agent.cache import (
    get_level1_cache, set_level1_cache,
    get_level2_cache, set_level2_cache,
    apply_code_offset,
)

# 开启 LangSmith 追踪
os.environ["LANGCHAIN_TRACING_V2"] = "true"
os.environ["LANGCHAIN_PROJECT"] = "LangGraph_Debug_Test"
os.environ["LANGCHAIN_API_KEY"] = "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"

def _retry_node(state: AgentState) -> dict:
    return {"retry_count": state.get("retry_count", 0) + 1}


def build_agent():
    g = StateGraph(AgentState)

    g.add_node("parse",    parse_node)
    g.add_node("code_gen", code_gen_node)
    g.add_node("sandbox",  sandbox_node)
    g.add_node("restore",  restore_node)
    g.add_node("quality",  quality_node)
    g.add_node("retry",    _retry_node)

    g.set_entry_point("parse")
    g.add_edge("parse",    "code_gen")
    g.add_edge("code_gen", "sandbox")
    g.add_edge("sandbox",  "restore")
    g.add_edge("restore",  "quality")
    g.add_edge("retry",    "code_gen")   # 重试跳过 parse，直接重新生成代码

    g.add_conditional_edges("quality", route, {"end": END, "retry": "retry"})

    return g.compile()


def _make_initial(
    excel_path:     str,
    sheet_name:     str,
    subtable_titles: List[str],    # 修改 str -> List[str] 以适配多子表抽取
    hints:          Optional[str],
    target_columns: Optional[Dict[str, List[Dict[str, Any]]]],   # 更新为字典
    reference_code: Optional[str] = None,   # 新增入参数
    reference_similarity: float = 0.0,  # 记忆匹配置信度
    reference_matched_titles: Optional[List[str]] = None,  # 匹配的子表名称
) -> AgentState:
    return {
        "config": {
            "excel_path": excel_path,
            "sheet_name": sheet_name,
            "subtable_titles": subtable_titles,
            "hints": hints,
            "target_columns": target_columns,
        },
        "sheet_structure": None,
        "reference_code": reference_code,
        "reference_similarity": reference_similarity,  # 新增：记忆匹配置信度
        "reference_matched_titles": reference_matched_titles or [],  # 新增：匹配的子表名称
        "generate_code": None,
        "generated_code":  None,
        "raw_result":      None,
        "raw_data":        None,
        "result":          None,
        "final_output":    None,
        "quality_score":   0.0,
        "retry_count":     0,
        "errors":          [],
        "sandbox_error":   None,
    }


def _run_extraction_with_cache(
    excel_path:     str,
    sheet_name:     str,
    subtable_titles: List[str],
    hints:          Optional[str]        = None,
    target_columns: Optional[Dict[str, List[Dict[str, Any]]]] = None,
) -> dict:
    """
    带 Level 2 缓存检查的提取入口。

    流程：
    1. 先执行 parse_node 获取 sheet_structure（纯代码，速度快）
    2. 检查 Level 2 结构指纹缓存
    3. 如果命中，计算坐标偏移并调整代码，直接执行
    4. 如果未命中，走完整 LangGraph 流程
    """
    from excel_agent.nodes.parse import parse_node
    from excel_agent.cache import get_level2_cache
    from excel_agent.state import AgentState

    # ── 步骤 1: 执行 parse_node 获取结构（纯代码，约 100-500ms）─────────
    initial_state: AgentState = {
        "config": {
            "excel_path": excel_path,
            "sheet_name": sheet_name,
            "subtable_titles": subtable_titles,
            "hints": hints,
            "target_columns": target_columns,
        },
        "sheet_structure": None,
        "reference_code": None,
        "reference_similarity": 0.0,
        "reference_matched_titles": [],
        "generated_code": None,
        "raw_result": None,
        "raw_data": None,
        "result": None,
        "final_output": None,
        "final_result": None,
        "quality_score": 0.0,
        "retry_count": 0,
        "errors": [],
        "sandbox_error": None,
        "header_map": None,
    }

    parse_result = parse_node(initial_state)
    sheet_structure = parse_result.get("sheet_structure")

    if not sheet_structure:
        # parse 失败，走完整流程
        return _run_full_agent_flow(
            excel_path, sheet_name, subtable_titles, hints, target_columns
        )

    # ── 步骤 2: 检查 Level 2 结构指纹缓存 ─────────
    l2_result = get_level2_cache(sheet_name, sheet_structure, subtable_titles)

    if l2_result:
        cached_code, cached_header_map, row_offset = l2_result

        # 修复：cached_header_map 可能是空字典，布尔值为 False
        if cached_code:
            # 应用坐标偏移
            if row_offset != 0:
                adjusted_code, adjusted_header_map = apply_code_offset(
                    cached_code, row_offset, cached_header_map or {}
                )
            else:
                adjusted_code, adjusted_header_map = cached_code, cached_header_map or {}

            # 直接执行调整后的代码（传入当前请求的 subtable_titles）
            sandbox_result = execute_cached_code(
                excel_path, sheet_name, adjusted_code, adjusted_header_map,
                subtable_titles, target_columns
            )

            if sandbox_result.get("success"):
                result_data = sandbox_result.get("data", {})
                quality = _compute_simple_quality(result_data)

                return {
                    "success":        quality >= 0.75,
                    "data":           result_data,
                    "quality_score":  quality,
                    "retry_count":    0,
                    "errors":         [],
                    "generated_code": adjusted_code,
                    "cache_hit":      True,
                    "cache_level":    "level2",
                    "row_offset":     row_offset,
                }

    # ── 步骤 3: Level 2 未命中，走完整 LangGraph 流程 ─────────
    return _run_full_agent_flow(
        excel_path, sheet_name, subtable_titles, hints, target_columns,
        sheet_structure  # 传入已解析的结构，避免重复 parse
    )


def _run_full_agent_flow(
    excel_path:     str,
    sheet_name:     str,
    subtable_titles: List[str],
    hints:          Optional[str]        = None,
    target_columns: Optional[Dict[str, List[Dict[str, Any]]]] = None,
    sheet_structure: Optional[Dict[str, Any]] = None,  # 可选：预解析的结构
) -> dict:
    """
    完整的 LangGraph Agent 流程（原有逻辑）
    """
    agent = build_agent()

    # 从记忆库读取参考代码
    ref_code, similarity, matched_titles = get_reference_code(sheet_name, subtable_titles)

    initial = _make_initial(
        excel_path, sheet_name, subtable_titles, hints, target_columns,
        ref_code, similarity, matched_titles
    )

    # 如果有预解析的结构，直接注入（跳过 parse_node）
    if sheet_structure:
        initial["sheet_structure"] = sheet_structure

    final = agent.invoke(initial)
    success = final["quality_score"] >= QUALITY_THRESHOLD

    # 成功后缓存
    if success and final["quality_score"] >= 0.95 and final.get("generated_code"):
        generated_code = final["generated_code"]

        save_reference_code(
            sheet_name,
            subtable_titles,
            generated_code,
            quality_score=final["quality_score"]
        )

        # 保存到 Level 1 和 Level 2 缓存
        # 注意：当前流程不生成 header_map，只缓存代码
        set_level1_cache(
            excel_path,
            sheet_name,
            subtable_titles,
            header_map={},  # 暂不需要
            generated_code=generated_code,
            sheet_structure=final.get("sheet_structure", {})
        )

        set_level2_cache(
            sheet_name,
            final.get("sheet_structure", {}),
            subtable_titles,
            header_map={},  # 暂不需要
            generated_code=generated_code
        )

    return {
        "success":        final["quality_score"] >= QUALITY_THRESHOLD,
        "data":           final.get("final_output") or final.get("result"),
        "quality_score":  final["quality_score"],
        "retry_count":    final["retry_count"],
        "errors":         final["errors"],
        "generated_code": final.get("generated_code", ""),
        "sandbox_error":  final.get("sandbox_error"),
        "cache_hit":      False,
    }


def run_extraction(
    excel_path:     str,
    sheet_name:     str,
    subtable_titles: List[str],    # str -> List[str]
    hints:          Optional[str]        = None,
    target_columns: Optional[Dict[str, List[Dict[str, Any]]]] = None,    # 更新为字典
) -> dict:
    """
    同步调用入口（优化版）。

    优化策略（分层缓存）：
        Level 1: 完全匹配缓存 - 文件 hash + sheet + titles 完全相同时直接复用代码
        Level 2: 结构指纹缓存 - 结构相同时计算坐标偏移并调整代码
        Level 3: LLM 生成 - 兜底方案

    参数：
        excel_path      : Excel 文件路径
        sheet_name      : Sheet 名，如 "CONFIGURATION"
        subtable_titles  : 子表标题关键词，如 "4G Configuration"
        hints           : 可选额外提示
        target_columns  : 列过滤，None=全部。示例：
                          [
                            {"parent": None,       "child": "CELL"},
                            {"parent": "ANTENNAS", "child": "Antenna Qty."},
                          ]

    返回：
        {
            "success":        bool,
            "data":           List[List[str]],
            "quality_score":  float,
            "retry_count":    int,
            "errors":         List[str],
            "generated_code": str,   # LLM 生成的代码（供调试）
            "cache_hit":      bool,  # 是否命中缓存
            "cache_level":    str,   # 命中的缓存层级（level1/level2）
        }
    """

    # ═══════════════════════════════════════════════════════════
    # Level 1: 检查完全匹配缓存（最快，毫秒级）
    # ═══════════════════════════════════════════════════════════
    l1_result = get_level1_cache(excel_path, sheet_name, subtable_titles)

    if l1_result:
        cached_code = l1_result.get("generated_code")
        cached_header_map = l1_result.get("header_map")

        # 修复：cached_header_map 可能是空字典，布尔值为 False
        if cached_code:  # 只要有代码就执行
            sandbox_result = execute_cached_code(
                excel_path, sheet_name, cached_code, cached_header_map or {},
                subtable_titles, target_columns
            )

            if sandbox_result.get("success"):
                result_data = sandbox_result.get("data", {})
                quality = _compute_simple_quality(result_data)

                return {
                    "success":        quality >= 0.75,
                    "data":           result_data,
                    "quality_score":  quality,
                    "retry_count":    0,
                    "errors":         [],
                    "generated_code": cached_code,
                    "cache_hit":      True,
                    "cache_level":    "level1",
                }

    # ═══════════════════════════════════════════════════════════
    # Level 2/3: 执行提取（自动检查 L2 缓存）
    # ═══════════════════════════════════════════════════════════
    return _run_extraction_with_cache(
        excel_path, sheet_name, subtable_titles, hints, target_columns
    )


def _compute_simple_quality(result: Dict) -> float:
    """简单质量评分（用于缓存代码执行后的快速验证）"""
    if not result:
        return 0.0

    total_rows = 0
    empty_tables = 0

    for title, data in result.items():
        if not data or len(data) <= 1:
            empty_tables += 1
        else:
            total_rows += len(data) - 1  # 减去表头

    if empty_tables > 0 and total_rows == 0:
        return 0.3

    if total_rows > 0:
        return 1.0

    return 0.5


def execute_cached_code(
    excel_path: str,
    sheet_name: str,
    cached_code: str,
    cached_header_map: Dict[str, Any],
    subtable_titles: List[str],
    target_columns: Optional[Dict[str, List[Dict[str, Any]]]] = None,
) -> Dict[str, Any]:
    """
    直接执行缓存的代码（不经过 LangGraph 流程）

    返回：
        {
            "success": bool,
            "data": Dict[str, List[List[str]]],
            "error": Optional[str]
        }
    """
    from excel_agent.nodes.sandbox import sandbox_node
    from excel_agent.state import AgentState

    # 构造最小化 state（适配 sandbox_node）
    state: AgentState = {
        "config": {
            "excel_path": excel_path,
            "sheet_name": sheet_name,
            "subtable_titles": subtable_titles,  # 使用传入的子表标题
            "hints": None,
            "target_columns": target_columns,  # 使用传入的列配置
        },
        "sheet_structure": {"sheet_name": sheet_name},  # 最小化结构
        "reference_code": None,
        "reference_similarity": 0.0,
        "reference_matched_titles": [],
        "header_map": cached_header_map or {},  # 可选
        "generated_code": cached_code,
        "raw_result": None,
        "raw_data": None,
        "result": None,
        "final_output": None,
        "final_result": None,
        "quality_score": 0.0,
        "retry_count": 0,
        "errors": [],
        "sandbox_error": None,
    }

    try:
        # 执行 sandbox_node
        sandbox_result = sandbox_node(state)

        if sandbox_result.get("sandbox_error"):
            return {
                "success": False,
                "data": {},
                "error": sandbox_result["sandbox_error"],
            }

        # 执行 restore_node（格式化 + 列过滤）
        from excel_agent.nodes.restore import restore_node
        state["raw_result"] = sandbox_result["raw_result"]
        restore_result = restore_node(state)

        return {
            "success": True,
            "data": restore_result.get("result", {}),
            "error": None,
        }
    except Exception as e:
        import traceback
        return {
            "success": False,
            "data": {},
            "error": f"{str(e)}\n{traceback.format_exc()}",
        }


async def run_extraction_deep_stream(
    excel_path:     str,
    sheet_name:     str,
    subtable_titles: List[str],
    hints:          Optional[str]        = None,
    target_columns: Optional[List[Dict]] = None,
) -> AsyncGenerator[dict, None]:
    """
    异步流式入口（调试用）。
    捕获 Token 吐字、节点完成事件。
    """
    agent   = build_agent()
    initial = _make_initial(excel_path, sheet_name, subtable_titles, hints, target_columns)

    async for event in agent.astream_events(initial, version="v2"):
        kind = event["event"]
        name = event["name"]

        if kind == "on_chat_model_stream":
            content = event["data"]["chunk"].content
            if content:
                yield {"type": "token", "content": content}

        elif kind == "on_chain_end" and name in (
            "parse", "code_gen", "sandbox", "restore", "quality"
        ):
            yield {
                "type": "node_end",
                "name": name,
                "data": event["data"].get("output", {}),
            }
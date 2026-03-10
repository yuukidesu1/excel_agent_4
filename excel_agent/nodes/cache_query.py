"""
nodes/cache_query.py — 缓存查询节点

职责：
    在 parse 之后、code_gen 之前查询三级缓存系统。

    - 命中 L1 → 直接使用缓存代码，跳过 LLM
    - 命中 L2 → 调整代码坐标后使用，跳过 LLM
    - 未命中 → 继续 code_gen 流程（LLM 生成）

    注意：L2 缓存查询使用 light_structure_fingerprint（轻量级结构指纹）

输入：
    - config.excel_path
    - config.sheet_name
    - config.subtable_titles
    - sheet_structure (parse_node 输出)
    - light_structure_fingerprint (light_structure_analyzer_node 输出)

输出：
    - cache_hit: bool
    - cache_level: "l1" | "l2" | "l3"
    - generated_code: str (命中时直接从缓存获取)
    - header_map: dict (命中时直接从缓存获取)
    - l2_row_offset: int (L2 命中时的行偏移量)
"""

from typing import Optional, Dict, Any
from pathlib import Path

from excel_agent.state import AgentState
from excel_agent.cache_manager import l1_get, apply_code_offset, l2_search_by_light_signature


def cache_query_node(state: AgentState) -> dict:
    """
    缓存查询节点（优先查询 L1，然后 L2）

    流程：
    1. 读取 excel_path, sheet_name, subtable_titles
    2. 查询 L1 缓存（完全匹配）
    3. 如果未命中，使用 light_structure_fingerprint 查询 L2 缓存
    4. 如果命中，返回缓存的代码和 header_map（L2 命中时计算偏移）
    5. 如果未命中，返回空（继续 LLM 生成流程）
    """
    config = state.get("config", {})
    excel_path = config.get("excel_path")
    sheet_name = config.get("sheet_name")
    subtable_titles = config.get("subtable_titles", [])

    if not excel_path or not sheet_name or not subtable_titles:
        return {
            "cache_hit": False,
            "cache_level": "l3",
            "cache_skip_reason": "缺少必要参数"
        }

    # 查询 L1 缓存（完全匹配）
    l1_data = l1_get(excel_path, sheet_name, subtable_titles)

    if l1_data:
        # 命中 L1 缓存
        return {
            "cache_hit": True,
            "cache_level": "l1",
            "generated_code": l1_data.get("generated_code", ""),
            "header_map": l1_data.get("header_map", {}),
        }

    # 未命中 L1，使用 light_structure_fingerprint 查询 L2 缓存
    light_fp = state.get("light_structure_fingerprint")
    if light_fp:
        l2_result = l2_search_by_light_signature(sheet_name, light_fp)
        if l2_result:
            l2_data, offset = l2_result
            generated_code = l2_data.get("generated_code", "")
            header_map = l2_data.get("header_map", {})

            result = {
                "cache_hit": True,
                "cache_level": "l2",
                "generated_code": generated_code,
                "header_map": header_map,
            }

            # L2 命中需要调整代码坐标
            if offset != 0:
                adjusted_code, adjusted_header_map = apply_code_offset(
                    generated_code, offset, 0, header_map
                )
                result["generated_code"] = adjusted_code
                result["header_map"] = adjusted_header_map
                result["l2_row_offset"] = offset

            return result

    # 未命中缓存
    return {
        "cache_hit": False,
        "cache_level": "l3",
    }

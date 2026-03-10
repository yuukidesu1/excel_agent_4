"""
nodes/cache_save.py — 缓存保存节点

职责：
    当 LLM 成功生成高质量代码后，将结果保存到三级缓存系统。

输入：
    - generated_code: LLM 生成的代码
    - header_map: 表头映射
    - structure_fingerprint: 结构指纹（来自 structure_analyzer_node）
    - light_structure_fingerprint: 轻量级结构指纹（来自 light_structure_analyzer_node）
    - config.excel_path, sheet_name, subtable_titles

输出：
    - cache_saved: bool
"""

from typing import Optional, Dict, Any
from pathlib import Path

from excel_agent.state import AgentState
from excel_agent.cache_manager import save_to_cache


def cache_save_node(state: AgentState) -> dict:
    """
    缓存保存节点

    保存条件：
    1. 代码生成成功（generated_code 非空）
    2. 质量达标（quality_score >= 0.75）
    3. 非缓存命中（cache_hit == False，避免重复保存）
    """
    config = state.get("config", {})
    excel_path = config.get("excel_path")
    sheet_name = config.get("sheet_name")
    subtable_titles = config.get("subtable_titles", [])

    generated_code = state.get("generated_code", "")
    quality_score = state.get("quality_score", 0.0)
    cache_hit = state.get("cache_hit", False)
    header_map = state.get("header_map", {})
    structure_fingerprint = state.get("structure_fingerprint")
    light_structure_fingerprint = state.get("light_structure_fingerprint")

    # 不满足保存条件
    if not generated_code:
        return {"cache_saved": False, "cache_save_reason": "generated_code 为空"}

    if quality_score < 0.75:
        return {"cache_saved": False, "cache_save_reason": f"质量评分 {quality_score} < 0.75"}

    if cache_hit:
        return {"cache_saved": False, "cache_save_reason": "数据来自缓存，无需重复保存"}

    if not excel_path or not sheet_name or not subtable_titles:
        return {"cache_saved": False, "cache_save_reason": "缺少必要参数"}

    # 保存到缓存
    try:
        save_to_cache(
            excel_path=excel_path,
            sheet_name=sheet_name,
            subtable_titles=subtable_titles,
            data={
                "generated_code": generated_code,
                "header_map": header_map,
            },
            structure_fingerprint=structure_fingerprint,
            light_structure_fingerprint=light_structure_fingerprint,
        )
        return {
            "cache_saved": True,
            "cache_save_reason": "LLM 生成成功，已保存到 L1 和 L2 缓存"
        }
    except Exception as e:
        return {
            "cache_saved": False,
            "cache_save_reason": f"保存失败：{str(e)}"
        }

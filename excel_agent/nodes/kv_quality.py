"""
nodes/kv_quality.py - quality checks and routing for KV extraction.
"""

from excel_agent.state import AgentState


MAX_RETRY = 3
QUALITY_THRESHOLD = 0.75


def kv_quality_node(state: AgentState) -> dict:
    config = state.get("config", {}) or {}
    kv_list = config.get("kv_list") or config.get("subtable_titles") or []
    result = state.get("result") or (state.get("kv_state", {}) or {}).get("extracted_data") or {}
    errors = list(state.get("errors", []))
    score = 1.0

    if not isinstance(result, dict) or not result:
        return {
            "quality_score": 0.0,
            "errors": errors + ["[KV质量] 返回结果为空或不是字典。"],
        }

    missing_keys = [key for key in kv_list if key not in result]
    if missing_keys:
        score -= 0.35
        errors.append(f"[KV质量] 缺少配置中的 Key：{missing_keys}")

    empty_keys = [key for key in kv_list if not str(result.get(key, "")).strip()]
    if kv_list and empty_keys:
        empty_rate = len(empty_keys) / len(kv_list)
        score -= min(0.60, 0.60 * empty_rate)
        errors.append(f"[KV质量] 空值 Key 占比过高：{empty_rate:.0%}，空值：{empty_keys}")

    return {
        "quality_score": round(max(0.0, min(1.0, score)), 3),
        "errors": errors,
    }


def kv_route(state: AgentState) -> str:
    if state.get("sandbox_error"):
        if state.get("retry_count", 0) >= MAX_RETRY:
            return "end"
        return "retry"
    if state.get("quality_score", 0.0) >= QUALITY_THRESHOLD:
        return "end"
    if state.get("retry_count", 0) >= MAX_RETRY:
        return "end"
    return "retry"

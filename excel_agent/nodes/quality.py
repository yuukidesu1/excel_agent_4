"""
nodes/quality.py — 质量打分 + 条件路由

质量校验维度（累计扣分）：
    1. 结果非空        (-0.50)：无数据行
    2. 列发现非空      (-0.30)：LLM 未发现任何列（all_columns 为空）
    3. 数据空值率      (-0.20)：超过 90% 的数据格为空（可能定位偏了）
    4. 行数合理性      (-0.20)：数据行数为 0（仅表头）

通过阈值：quality_score >= 0.75
重试上限：MAX_RETRY = 3（超过后强制输出现有结果）
"""

from excel_agent.state import AgentState

MAX_RETRY         = 3
QUALITY_THRESHOLD = 0.75


def quality_node(state: AgentState) -> dict:
    result  = state.get("result", [])
    hmap    = state.get("header_map", {})
    errors  = list(state.get("errors", []))
    score   = 1.0

    # ── 校验 1：结果非空 ─────────────────────────────────────
    if not result or len(result) < 1:
        score -= 0.5
        errors.append("[质量] 结果完全为空，子表定位可能失败。")

    # ── 校验 2：列发现非空 ────────────────────────────────────
    all_cols = hmap.get("all_columns", [])
    if not all_cols:
        score -= 0.3
        errors.append("[质量] 未发现任何列，请检查子表标题是否正确。")

    # ── 校验 3：数据行空值率 ──────────────────────────────────
    if result and len(result) > 1:
        data_rows   = result[1:]
        total_cells = sum(len(r) for r in data_rows)
        empty_cells = sum(1 for r in data_rows for v in r if v == "" or v is None)
        empty_rate  = empty_cells / total_cells if total_cells else 1.0
        if empty_rate > 0.9 and len(data_rows) > 1:
            score -= 0.2
            errors.append(
                f"[质量] 数据行空值率过高：{empty_rate:.0%}，"
                f"可能定位到了错误区域，请检查 subtable_title 是否正确。"
            )

    # ── 校验 4：数据行数合理性 ────────────────────────────────
    if result and len(result) <= 1:
        score -= 0.2
        errors.append("[质量] 无数据行（仅有表头），子表可能为空或数据行未被识别。")

    score = round(max(0.0, min(1.0, score)), 3)
    return {"quality_score": score, "errors": errors}


def route(state: AgentState) -> str:
    """LangGraph 条件路由：通过 → end，不通过 → retry"""
    if state["quality_score"] >= QUALITY_THRESHOLD:
        return "end"
    if state.get("retry_count", 0) >= MAX_RETRY:
        return "end"   # 达到上限，强制输出现有最佳结果
    return "retry"

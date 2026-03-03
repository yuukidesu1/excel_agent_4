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

"""
nodes/quality.py — 质量打分 + 条件路由

与原版的区别：
  - 沙盒执行失败（sandbox_error 不为空）直接重试，不扣分（已在 agent 路由层处理）
  - 其余校验维度与原版一致
"""


def quality_node(state: AgentState) -> dict:
    result  = state.get("result", [])
    errors  = list(state.get("errors", []))
    score   = 1.0

    # 校验 1：结果非空
    if not result:
        score -= 0.50
        errors.append("[质量] 结果完全为空。")

    # 校验 2：无数据行
    if result and len(result) <= 1:
        score -= 0.20
        errors.append("[质量] 无数据行（仅有表头），子表可能为空或数据起始行判断有误。")

    # 校验 3：数据行空值率过高
    if result and len(result) > 1:
        data_rows = result[1:]
        total     = sum(len(r) for r in data_rows)
        empty     = sum(1 for r in data_rows for v in r if not v)
        rate      = empty / total if total else 1.0
        if rate > 0.9 and len(data_rows) > 1:
            score -= 0.20
            errors.append(f"[质量] 数据空值率过高：{rate:.0%}，可能定位偏移或表头行数判断有误。")

    # 校验 4：target_columns 中有未找到的列
    target_columns = state.get("target_columns")
    if target_columns and result and result[0]:
        missing = [h for h in result[0] if h.startswith("[未找到]")]
        if missing:
            score -= 0.25 * len(missing) / len(target_columns)
            errors.append(f"[质量] 以下列未找到：{missing}")

    return {
        "quality_score": round(max(0.0, min(1.0, score)), 3),
        "errors": errors,
    }


def route(state: AgentState) -> str:
    """路由逻辑：沙盒失败 or 质量不达标 → 重试"""
    if state.get("sandbox_error"):
        # 代码执行失败，必须重试
        if state.get("retry_count", 0) >= MAX_RETRY:
            return "end"
        return "retry"
    if state["quality_score"] >= QUALITY_THRESHOLD:
        return "end"
    if state.get("retry_count", 0) >= MAX_RETRY:
        return "end"
    return "retry"

def quality_node_(state: AgentState) -> dict:
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


def route_(state: AgentState) -> str:
    """LangGraph 条件路由：通过 → end，不通过 → retry"""
    if state["quality_score"] >= QUALITY_THRESHOLD:
        return "end"
    if state.get("retry_count", 0) >= MAX_RETRY:
        return "end"   # 达到上限，强制输出现有最佳结果
    return "retry"

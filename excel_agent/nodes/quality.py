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
    result = state.get("result", {})
    errors = list(state.get("errors", []))
    score = 1.0

    config = state.get("config", {})
    target_titles = (
        config.get("subtable_titles") or state.get("subtable_titles") or
        config.get("target_title") or state.get("target_title", [])
    ) # 修复读取不到config的问题
    if isinstance(target_titles, str):
        target_titles = [target_titles]
    target_columns = config.get("target_columns") or state.get("target_columns")

    if not result:
        score -= 0.50
        errors.append("[质量] 返回结果为空字典。")
        return {"quality_score": max(0.0, score), "errors": errors}

    empty_tables = 0
    total_data_cells = 0
    empty_data_cells = 0
    missing_cols_count = 0
    expected_cols_total = 0  # 🚀 新增：用于按比例公平扣分

    for title in target_titles:
        table_data = result.get(title)

        if not table_data or len(table_data) <= 1:
            empty_tables += 1
            errors.append(f"[质量] 子表 '{title}' 无数据或定位失败。")
            continue

        data_rows = table_data[1:]
        total_data_cells += sum(len(r) for r in data_rows)
        empty_data_cells += sum(1 for r in data_rows for v in r if not v)

        # 🚀 获取当前子表的列过滤规则
        current_targets = None
        if isinstance(target_columns, dict):
            current_targets = target_columns.get(title)
        elif isinstance(target_columns, list):
            current_targets = target_columns

        # 检查丢失的列
        if current_targets:
            expected_cols_total += len(current_targets)  # 累加预期总列数
            if table_data[0]:
                missing = [h for h in table_data[0] if str(h).startswith("[未找到]")]
                missing_cols_count += len(missing)
                if missing:
                    errors.append(f"[质量] 子表 '{title}' 中以下列未找到：{missing}")

    if empty_tables > 0:
        score -= (0.30 * (empty_tables / len(target_titles)))

    if total_data_cells > 0:
        rate = empty_data_cells / total_data_cells
        if rate > 0.85:
            score -= 0.10
            errors.append(f"[质量] 总体数据空值率过高：{rate:.0%}。")

    # 🚀 修改扣分逻辑：基于预期总列数扣分
    if expected_cols_total > 0 and missing_cols_count > 0:
        score -= 0.50 * (missing_cols_count / expected_cols_total)

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
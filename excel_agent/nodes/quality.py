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

    if not result:
        return {"quality_score": 0.0, "errors": errors + ["[质量] 返回结果为空字典。"]}

    # 分离 table 子表和 KV 子表
    table_subtables = {t: d for t, d in result.items() if isinstance(d, list)}
    kv_subtables = {t: d for t, d in result.items() if isinstance(d, dict)}

    # 分别打分然后加权
    table_score = _check_table_quality(state, table_subtables, errors) if table_subtables else 1.0
    kv_score = _check_kv_quality(state, kv_subtables, errors) if kv_subtables else 1.0

    # 加权平均（按子表数量加权）
    table_count = len(table_subtables)
    kv_count = len(kv_subtables)
    total = table_count + kv_count

    if total == 0:
        combined = 1.0
    elif table_count == 0:
        combined = kv_score
    elif kv_count == 0:
        combined = table_score
    else:
        combined = (table_score * table_count + kv_score * kv_count) / total

    return {
        "quality_score": round(max(0.0, min(1.0, combined)), 3),
        "errors": errors,
    }


def _check_table_quality(state: AgentState, table_subtables: dict, errors: list) -> float:
    """检查 Table 格式子表的质量"""
    score = 1.0

    config = state.get("config", {})
    target_titles = (
        config.get("subtable_titles") or state.get("subtable_titles") or
        config.get("target_title") or state.get("target_title", [])
    )
    if isinstance(target_titles, str):
        target_titles = [target_titles]
    subtable_configs = config.get("subtable_configs") or state.get("subtable_configs")

    if not table_subtables:
        score -= 0.50
        errors.append("[质量] 返回结果为空字典。")
        return max(0.0, score)

    empty_tables = 0
    total_data_cells = 0
    empty_data_cells = 0
    missing_cols_count = 0
    expected_cols_total = 0

    for title in target_titles:
        if title not in table_subtables:
            continue
        table_data = table_subtables[title]

        if not table_data or len(table_data) <= 1:
            empty_tables += 1
            errors.append(f"[质量] 子表 '{title}' 无数据或定位失败。")
            continue

        data_rows = table_data[1:]
        total_data_cells += sum(len(r) for r in data_rows)
        empty_data_cells += sum(1 for r in data_rows for v in r if not v)

        current_targets = None
        if isinstance(subtable_configs, dict):
            current_targets = subtable_configs.get(title)
        elif isinstance(subtable_configs, list):
            current_targets = subtable_configs

        if current_targets:
            expected_cols_total += len(current_targets)
            if table_data[0]:
                missing = [h for h in table_data[0] if str(h).startswith("[未找到]")]
                missing_cols_count += len(missing)
                if missing:
                    errors.append(f"[质量] 子表 '{title}' 中以下列未找到：{missing}")

    if empty_tables > 0:
        score -= (0.30 * (empty_tables / max(1, len(target_titles))))

    if total_data_cells > 0:
        rate = empty_data_cells / total_data_cells
        if rate > 0.85:
            score -= 0.10
            errors.append(f"[质量] 总体数据空值率过高：{rate:.0%}。")

    if expected_cols_total > 0 and missing_cols_count > 0:
        score -= 0.50 * (missing_cols_count / expected_cols_total)

    return max(0.0, min(1.0, score))


def _check_kv_quality(state: AgentState, kv_subtables: dict, errors: list) -> float:
    """检查 KV 格式子表的质量"""
    score = 1.0
    config = state.get("config", {})
    kv_list = config.get("kv_list", [])

    total_found = 0
    total_keys = 0

    for title, kv_data in kv_subtables.items():
        if not kv_data:
            errors.append(f"[质量] KV 子表 '{title}' 结果为空。")
            score -= 0.30
            continue

        found = sum(1 for v in kv_data.values() if v)
        total = len(kv_data)
        total_found += found
        total_keys += total

        if total > 0 and found / total < 0.5:
            missing_keys = [k for k, v in kv_data.items() if not v]
            errors.append(f"[质量] KV 子表 '{title}' Key 发现率过低：{found}/{total}，缺失：{missing_keys}")

    if total_keys > 0:
        discovery_rate = total_found / total_keys
        if discovery_rate < 0.5:
            score -= 0.50
        elif discovery_rate < 1.0:
            score -= 0.10 * (1 - discovery_rate)

    return max(0.0, min(1.0, score))


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
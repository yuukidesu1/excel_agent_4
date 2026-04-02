"""
nodes/restore.py — 结果组装 + 列过滤（纯代码）

新增功能：列过滤
  若 state["config"] 中存在 subtable_configs / target_columns，则只输出指定的列。

  匹配规则：
    - 兼容新版配置：提取 col_headers 或 row_headers 列表中的 "A||B" 字符串
    - 兼容旧版配置：提取 {"parent": "A", "child": "B"} 字典
    - 忽略大小写 + 忽略空格/换行符
"""

from typing import Any, List, Optional, Dict, Union
from excel_agent.state import AgentState


def _fmt(v: Any) -> str:
    """统一单元格值格式"""
    if v is None:
        return ""
    if isinstance(v, float) and v == int(v):
        return str(int(v))
    return str(v).strip()


def _norm(s: Optional[str]) -> str:
    """模糊匹配归一化：小写 + 去空格 + 去换行"""
    if s is None:
        return ""
    return s.lower().replace(" ", "").replace("\n", "").replace("\r", "").replace("\t", "")


def _parse_target(target: Union[str, Dict]) -> tuple:
    """
    统一解析 target，返回 (parent, child) 元组。
    支持旧版 dict: {"parent": "A", "child": "B"}
    支持新版 str:  "A||B" 或 "B"
    """
    if isinstance(target, dict):
        return target.get("parent"), str(target.get("child", ""))
    elif isinstance(target, str):
        if "||" in target:
            parts = target.split("||", 1)
            return parts[0], parts[1]
        return None, target
    return None, str(target)


def _find_col_index(header: List[str], target: Union[str, Dict]) -> int:
    """
    在表头行中找到 target 对应的列索引（-1 表示未找到）。
    """
    t_parent, t_child_raw = _parse_target(target)
    t_child = _norm(t_child_raw)

    for i, h in enumerate(header):
        if "||" in h:
            parts = h.split("||", 1)
            col_parent, col_child = parts[0], parts[1]
        else:
            col_parent, col_child = None, h

        # 子类必须匹配
        if _norm(col_child) != t_child:
            continue

        # 父类如果指定了，也必须匹配
        if t_parent is not None and _norm(col_parent) != _norm(t_parent):
            continue

        return i

    return -1  # 未找到


def _find_row_index(first_col: List[str], target: Union[str, Dict]) -> int:
    """
    在首列（行头）中找到 target 对应的行索引（-1 表示未找到）。
    用于"仅行"布局的行过滤。
    """
    t_parent, t_child_raw = _parse_target(target)
    t_child = _norm(t_child_raw)

    for i, h in enumerate(first_col):
        if "||" in h:
            parts = h.split("||", 1)
            row_parent, row_child = parts[0], parts[1]
        else:
            row_parent, row_child = None, h

        # 子类必须匹配
        if _norm(row_child) != t_child:
            continue

        # 父类如果指定了，也必须匹配
        if t_parent is not None and _norm(row_parent) != _norm(t_parent):
            continue

        return i

    return -1  # 未找到


def restore_node(state: AgentState) -> dict:
    raw_result = state.get("raw_result") or {}
    config = state.get("config", {})
    sheet_structure = state.get("sheet_structure", {})
    cache_state = state.get("cache", {})

    # 获取配置 (优先新版 subtable_configs)
    target_columns_config = config.get("subtable_configs") or config.get("target_columns")
    # 获取子表标题列表（用于过滤逻辑）
    subtable_titles = config.get("subtable_titles", [])

    if not raw_result:
        return {"result": {}, "final_output": {}}

    final_res = {}

    # ── 第 1 步：先格式化所有数据 ──
    formatted_result: Dict[str, List[List[str]]] = {}
    for title, table_data in raw_result.items():
        if not table_data:
            formatted_result[title] = []
            continue
        formatted = [[_fmt(cell) for cell in row] for row in table_data]
        formatted_result[title] = formatted

    # ── 新增：混合表类型边界过滤 ──
    # 当存在 kv 表和普通表混合时，需要防止普通表向下穿透到 kv 表区域
    if subtable_titles and sheet_structure:
        # 获取所有子表的配置信息（如果有）
        target_configs_dict = config.get("subtable_configs") or config.get("target_columns") or {}

        # 找出所有 kv 表的位置（根据配置或 cache_state 中的 layout_type 判断）
        kv_table_positions = []
        for title in subtable_titles:
            tc_def = target_configs_dict.get(title, {})
            layout_type = tc_def.get("layout") if isinstance(tc_def, dict) else None

            # 也检查 cache_state 中的 layout_type
            cache_entry = cache_state.get("entries", {}).get(title, {})
            cache_layout = cache_entry.get("layout_type")

            if layout_type == "kv" or cache_layout == "kv":
                # 尝试从 cache_state 中获取 kv 表的 start_row
                kv_start_row = cache_entry.get("start_row")
                if kv_start_row:
                    kv_table_positions.append({
                        "title": title,
                        "start_row": kv_start_row
                    })

        # 如果有 kv 表，检查前面的普通表是否穿透了
        if kv_table_positions:
            for kv_pos in kv_table_positions:
                kv_start_row = kv_pos["start_row"]
                kv_title = kv_pos["title"]

                # 找到 kv 表前面的那个子表
                kv_index = subtable_titles.index(kv_title)
                if kv_index > 0:
                    prev_title = subtable_titles[kv_index - 1]
                    if prev_title in formatted_result and formatted_result[prev_title]:
                        # 检查 prev_title 的抽取代码中记录的 start_row
                        prev_entry = cache_state.get("entries", {}).get(prev_title, {})
                        prev_start_row = prev_entry.get("start_row", 0)

                        # 计算 prev_title 最多能占多少行（到 kv 表之前的行数）
                        max_allowed_rows = kv_start_row - prev_start_row - 1  # -1 是给标题行留的空间

                        # 如果 prev_title 的抽取结果超过了这个行数，截断它
                        if len(formatted_result[prev_title]) > max_allowed_rows:
                            formatted_result[prev_title] = formatted_result[prev_title][:max_allowed_rows]

    # ── 第 2 步：方法一 - 相邻子表值过滤 ──
    # 检查每个子表（除了最后一个）的最后一行是否等于下一个子表的标题
    for i, title in enumerate(subtable_titles[:-1]):
        if title not in formatted_result or len(formatted_result[title]) <= 1:
            continue
        next_title = subtable_titles[i + 1]
        last_row = formatted_result[title][-1]
        # 如果最后一行的所有单元格值都等于下一个子表标题，则删除该行
        if last_row and all(cell == next_title for cell in last_row):
            formatted_result[title].pop()

    # ── 第 3 步：方法二 - 合并单元格截断 ──
    # 找到最左侧合并单元格 row_span 最大的那个，用它来判断是否需要截断最后一个子表
    use_fallback = False
    if subtable_titles and sheet_structure:
        merged_cells = sheet_structure.get("merged_cells_info", [])
        if merged_cells:
            # 按 min_col 升序排序，找到最左侧的合并单元格
            # 在同 min_col 的情况下，选择 row_span 最大的
            sorted_merged = sorted(merged_cells, key=lambda x: (x["min_col"], -x["row_span"]))
            if sorted_merged:
                max_row_span_cell = sorted_merged[0]
                expected_total_rows = max_row_span_cell["row_span"]
                # 计算当前所有子表的总行数（包括表头）
                current_total_rows = sum(len(rows)+1 for rows in formatted_result.values()) # + 1 是表头所占的行数
                # 如果总行数超出预期，在最后一个子表进行截断
                if current_total_rows > expected_total_rows:
                    last_title = subtable_titles[-1]
                    excess_rows = current_total_rows - expected_total_rows
                    if last_title in formatted_result and len(formatted_result[last_title]) > excess_rows:
                        formatted_result[last_title] = formatted_result[last_title][:-excess_rows]
                else:
                    # 总行数没有超出，但最后一个子表仍然可能抽多了，使用众数方案
                    use_fallback = True
            else:
                use_fallback = True
        else:
            use_fallback = True
    elif subtable_titles:
        use_fallback = True

    # ── 备用方案：众数截断 ──
    # 当合并单元格信息不可用或无效时，使用众数判断合理行数
    if use_fallback and subtable_titles:
        # 找到最可能的行数（众数）
        row_counts = [len(rows) for rows in formatted_result.values() if rows]
        if row_counts:
            from collections import Counter
            count_counter = Counter(row_counts)
            # 找到最常见的行数（排除异常值）
            most_common = count_counter.most_common()
            if most_common:
                mode_row_count = most_common[0][0]
                # 截断所有超过这个行数的子表
                for title in subtable_titles:
                    if title in formatted_result and len(formatted_result[title]) > mode_row_count:
                        formatted_result[title] = formatted_result[title][:mode_row_count]

    # ── 第 4 步：列过滤逻辑 ──
    for title, formatted in formatted_result.items():
        if not formatted:
            final_res[title] = []
            continue

        header = formatted[0] if formatted else []
        data_rows = formatted[1:] if len(formatted) > 1 else []

        # ── 1. 安全提取当前子表的 target 列表 ──
        current_targets_raw = None
        if isinstance(target_columns_config, dict):
            current_targets_raw = target_columns_config.get(title)
        elif isinstance(target_columns_config, list):
            current_targets_raw = target_columns_config

        # ── 2. 剥离字典结构，拿到真正的 headers 列表 ──
        current_targets = []
        if isinstance(current_targets_raw, dict):
            # 如果是新版配置字典，根据 layout 提取对应的 headers
            layout = current_targets_raw.get("layout", "仅列")
            if layout in ["仅行", "kv"]:
                current_targets = current_targets_raw.get("row_headers", [])
            else:
                current_targets = current_targets_raw.get("col_headers", [])
        elif isinstance(current_targets_raw, list):
            # 如果是旧版，本身就是列表
            current_targets = current_targets_raw

        # ── 3. 执行过滤 ──
        if current_targets:
            keep_indices: List[int] = []
            keep_labels: List[str] = []

            for target in current_targets:
                idx = _find_col_index(header, target)
                keep_indices.append(idx)

                if idx >= 0:
                    keep_labels.append(header[idx])
                else:
                    _, t_child = _parse_target(target)
                    keep_labels.append(f"[未找到]{t_child}")

            new_header = keep_labels
            new_data = [
                [row[i] if (0 <= i < len(row)) else "" for i in keep_indices]
                for row in data_rows
            ]
            final_res[title] = [new_header] + new_data
        else:
            # 如果没有配置 target，直接返回完整表
            final_res[title] = [header] + data_rows

    return {"result": final_res, "final_output": final_res}
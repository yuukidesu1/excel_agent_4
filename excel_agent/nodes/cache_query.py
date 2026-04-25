"""
nodes/cache_query.py — 细粒度缓存查询节点（支持表格模式和 KV 模式）

职责：
    【表格模式】
    基于 PSA (前置结构分析器) 节点构建的各子表指纹，逐个进行缓存检索。

    工作流：
    1. 遍历 cache.entries 中的每一个子表建档。
    2. 使用 signature 查询 L2 缓存库。
    3. 若命中：动态计算行列偏移量 (row_offset, col_offset)，并调整 Python 代码中的绝对坐标。
    4. 若未命中：将子表名称加入 missed_subtables 列表，供下游 LLM 处理。
    5. 更新宏观调度标志 (all_cached, partial_cached)。

    【KV 模式】
    1. 基于 kv_list 和 sheet 特征构建缓存 Key。
    2. 查询 L1 缓存（完全匹配：同一个 Excel 文件）。
    3. 查询 L2 缓存（结构匹配：不同 Excel 但结构相同）。
    4. 若命中：将缓存代码填充到 state，供下游 kv_sandbox 直接使用。
    5. 若未命中：标记 missed=True，交由 kv_code_gen 节点生成新代码。

注意：本节点纯代码执行，不调用 LLM。
"""

from typing import Dict, Any

from excel_agent.state import AgentState, CacheState
from excel_agent.cache_manager import l2_get, apply_code_offset, l1_get
from excel_agent.nodes.kv_cache_save import build_kv_cache_key, compute_kv_structure_signature


def cache_query_node(state: AgentState) -> dict:
    """按子表粒度查询缓存，并执行代码坐标偏移修正（支持表格模式和 KV 模式）"""
    config = state.get("config", {})
    sheet_name = config.get("sheet_name")
    # 通过 kv_list 判断是否是 KV 模式（用户配置或 PSA 检测到 KV 布局后设置）
    kv_list = config.get("kv_list")
    extract_type = config.get("extract_type", "table")
    cache_state = state.get("cache", {})

    # KV 模式判断：有 kv_list 或 extract_type="kv"
    # 注意：PSA 节点检测到 KV 布局后会返回 cache.kv_auto_detected=True
    is_kv_mode = (kv_list is not None and len(kv_list) > 0) or (extract_type and extract_type.lower() == "kv") or cache_state.get("kv_auto_detected", False)

    # ==================== KV 模式缓存查询 ====================
    if is_kv_mode:
        return _kv_cache_query(state)

    # ==================== 表格模式缓存查询 ====================
    cache_state: CacheState = state.get("cache", {})
    entries = cache_state.get("entries", {})

    # 如果没有建档数据或缺少表名，直接放行交由 LLM 处理全量
    if not sheet_name or not entries:
        return {"cache": cache_state}

    missed_subtables = []
    hit_count = 0

    for title, entry in entries.items():
        signature = entry.get("signature")
        if not signature:
            missed_subtables.append((title, entry.get("layout_type")))
            continue

        # ── 1. 按子表独有特征查询缓存库 ──
        # 在新架构中，我们统一使用极智 Hash (signature) 查询，
        # 只要 Hash 一致，结构就绝对一致。
        cached_data = l2_get(sheet_name, signature)

        if cached_data:
            # ── 2. 缓存命中：计算 2D 偏移量 ──
            # 从缓存数据中读取当年存入这套代码时的“历史锚点”
            cached_start_row = cached_data.get("start_row", entry["start_row"])
            cached_start_col = cached_data.get("start_col", entry["start_col"])

            # 计算偏移量 (当前实际位置 - 历史编写代码时的位置)
            row_offset = entry["start_row"] - cached_start_row
            col_offset = entry["start_col"] - cached_start_col

            raw_code = cached_data.get("code", "")
            raw_header_map = cached_data.get("header_map") or {}

            # ── 3. 动态修正代码中的坐标 ──
            if row_offset != 0 or col_offset != 0:
                adjusted_code, adjusted_header_map = apply_code_offset(
                    raw_code, row_offset, col_offset, raw_header_map
                )
                cache_level = "l2"  # 存在偏移，定义为结构复用 (L2)
            else:
                adjusted_code = raw_code
                adjusted_header_map = raw_header_map
                cache_level = "l1"  # 零偏移，定义为精准复用 (相当于原先的 L1)

            # ── 4. 填充阶段二/阶段三状态 (查库成功) ──
            entry["cache_hit"] = True
            entry["cache_level"] = cache_level
            entry["l2_row_offset"] = row_offset
            entry["l2_col_offset"] = col_offset
            entry["code"] = adjusted_code
            entry["header_map"] = adjusted_header_map

            hit_count += 1
            print(f"✅ 缓存命中 [{cache_level.upper()}]: 子表 '{title}' (偏移: 行{row_offset}, 列{col_offset})")
        else:
            # ── 5. 缓存未命中 ──
            entry["cache_hit"] = False
            entry["cache_level"] = "miss"
            missed_subtables.append((title, entry.get("layout_type")))
            print(f"❌ 缓存未命中: 子表 '{title}'，将交由 LLM 生成代码。")

    # ── 6. 更新全局宏观调度标志 ──
    # 这些标志将决定 agent.py 里的路由是跳过 LLM 还是进入 LLM
    total_tables = len(entries)
    cache_state["all_cached"] = (hit_count == total_tables) and (total_tables > 0)
    cache_state["partial_cached"] = (hit_count > 0)
    cache_state["missed_subtables"] = missed_subtables

    return {"cache": cache_state}


def _kv_cache_query(state: AgentState) -> dict:
    """
    KV 模式缓存查询逻辑

    查询策略：
    1. 先查 L1 缓存（完全匹配，同一个 Excel 文件）
    2. 再查 L2 缓存（结构匹配，不同 Excel 但结构相同）
    3. 若命中，直接返回缓存代码；若未命中，交由 LLM 生成
    """
    config = state.get("config", {})
    sheet_name = config.get("sheet_name", "")
    # 优先从 config 读取 kv_list，如果 PSA 检测到 KV 布局并返回了 cache.kv_list，也一并读取
    kv_list = config.get("kv_list") or state.get("cache", {}).get("kv_list", [])
    excel_path = config.get("excel_path", "")

    # 如果没有 kv_list，直接返回未命中
    if not kv_list:
        return {
            "cache": {
                "hit": False,
                "cache_level": "miss",
                "missed": True
            }
        }

    # 获取 sheet 结构信息（用于构建缓存 Key）
    sheet_structure = state.get("sheet_structure", {})
    max_row = sheet_structure.get("max_row", 0)
    max_col = sheet_structure.get("max_col", 0)

    # ==================== 1. 查询 L1 缓存（完全匹配）====================
    l1_data = l1_get(
        excel_path=excel_path,
        sheet_name=sheet_name,
        subtable_titles=kv_list
    )

    if l1_data:
        # L1 命中：直接返回缓存代码
        code = l1_data.get("generated_code", "")
        print(f"✅ KV 缓存 L1 命中：key={build_kv_cache_key(sheet_name, kv_list, max_row, max_col)}")
        return {
            "cache": {
                "hit": True,
                "cache_level": "l1",
                "missed": False,
                "code": code,
                "l1_data": l1_data
            }
        }

    # ==================== 2. 查询 L2 缓存（结构匹配）====================
    structure_signature = compute_kv_structure_signature(kv_list, sheet_structure)

    if structure_signature:
        l2_data = l2_get(
            sheet_name=sheet_name,
            structure_signature=structure_signature
        )

        if l2_data:
            # L2 命中：返回缓存代码
            code = l2_data.get("generated_code", "")
            print(f"✅ KV 缓存 L2 命中：signature={structure_signature}")
            return {
                "cache": {
                    "hit": True,
                    "cache_level": "l2",
                    "missed": False,
                    "code": code,
                    "l2_data": l2_data,
                    "structure_signature": structure_signature
                }
            }

    # ==================== 3. 未命中，交由 LLM 生成 ====================
    print(f"❌ KV 缓存未命中，将交由 LLM 生成代码")
    return {
        "cache": {
            "hit": False,
            "cache_level": "miss",
            "missed": True,
            "structure_signature": structure_signature
        }
    }
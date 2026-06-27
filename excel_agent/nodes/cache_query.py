from excel_agent.state import AgentState, CacheState
from excel_agent.cache_manager import l2_get, apply_code_offset


def cache_query_node(state: AgentState) -> dict:
    """KV 缓存查询"""
    config = state.get("config", {})
    sheet_name = config.get("sheet_name")

    # =====================================================================
    # ── 1. KV 专属缓存查询轨道 ──
    # =====================================================================
    # 使用 config.get("kv_list") 或检查 kv_state 是否被前处理节点实质性填充
    is_kv_mode = bool(config.get("kv_list"))

    if is_kv_mode:
        kv_state = state.get("kv_state", {})
        kv_cache = kv_state.get("cache", {})
        signature = kv_cache.get("signature")

        # 异常兜底：如果没有算出版式指纹或缺少表名，直接交由 LLM 处理
        if not sheet_name or not signature:
            kv_cache["cache_hit"] = False
            return {"kv_state": {"cache": kv_cache}}

        # 查库：复用现有的底层接口 l2_get
        cached_data = l2_get(sheet_name, signature)

        if cached_data and cached_data.get("code"):
            # 🎯 缓存命中：Zero-Offset 机制，无需修正，原版代码直接可用！
            raw_code = cached_data.get("code")
            kv_cache["cache_hit"] = True
            kv_cache["code"] = raw_code

            print(f"✅ KV 缓存极速命中 [Signature: {signature}]! 准备直达沙盒执行。")

            # 直接将代码置入 generated_code 数组，方便下游 sandbox 无缝读取
            return {
                "kv_state": {
                    "cache": kv_cache,
                    "generated_code": [raw_code]
                }
            }
        else:
            # ❌ 缓存未命中：打回 LLM 重造
            kv_cache["cache_hit"] = False
            print(f"❌ KV 缓存未命中 [Signature: {signature}]，流水线将转入大模型生成节点。")

            return {"kv_state": {"cache": kv_cache}}
    """按子表粒度查询缓存，并执行代码坐标偏移修正"""
    config = state.get("config", {})
    sheet_name = config.get("sheet_name")

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
                cache_level = "l2"  # 存在偏移，定义为结构 复用 (L2)
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
            print(f"缓存命中 [{cache_level.upper()}]: 子表 '{title}' (偏移: 行{row_offset}, 列{col_offset})")
        else:
            # ── 5. 缓存未命中 ──
            entry["cache_hit"] = False
            entry["cache_level"] = "miss"
            missed_subtables.append((title, entry.get("layout_type")))
            print(f"缓存未命中: 子表 '{title}'，将交由 LLM 生成代码。")

    # ── 6. 更新全局宏观调度标志 ──
    # 这些标志将决定 agent.py 里的路由是跳过 LLM 还是进入 LLM
    total_tables = len(entries)
    cache_state["all_cached"] = (hit_count == total_tables) and (total_tables > 0)
    cache_state["partial_cached"] = (hit_count > 0)
    cache_state["missed_subtables"] = missed_subtables

    return {"cache": cache_state}

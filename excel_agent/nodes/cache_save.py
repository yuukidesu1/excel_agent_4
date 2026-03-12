"""
nodes/cache_save.py — 细粒度缓存保存节点

职责：
    当 LLM 成功生成代码，且 SA 节点成功将大块代码拆分为子表专属代码后，
    将这些新鲜出炉的专属代码按子表粒度保存到 L2 缓存系统中。

保存条件：
    1. 质量评分达标 (quality_score >= QUALITY_THRESHOLD，即 0.75)
    2. 子表原本未命中缓存 (cache_hit == False)
    3. 子表档案中已被填入了拆分后的纯净代码 (code 非空)
"""

from excel_agent.state import AgentState, CacheState
from excel_agent.cache_manager import l2_set
from excel_agent.nodes.quality import QUALITY_THRESHOLD


def cache_save_node(state: AgentState) -> dict:
    """按子表粒度，将新生成的专属代码保存至 L2 缓存"""
    config = state.get("config", {})
    sheet_name = config.get("sheet_name")

    cache_state: CacheState = state.get("cache", {})
    entries = cache_state.get("entries", {})
    quality_score = state.get("quality_score", 0.0)

    # ── 1. 宏观质量与参数校验 ──
    if quality_score < QUALITY_THRESHOLD:
        return {"cache_saved": False, "cache_save_reason": f"沙盒质量评分 {quality_score} 偏低，为防止污染缓存，取消保存。"}

    if not sheet_name or not entries:
        return {"cache_saved": False, "cache_save_reason": "缺少 sheet_name 或子表档案数据。"}

    # ── 2. 遍历档案，精准入库 ──
    saved_count = 0
    saved_titles = []

    for title, entry in entries.items():
        # 核心逻辑：只保存以前没命中，且现在有了独立代码的表
        if not entry.get("cache_hit") and entry.get("code"):
            signature = entry.get("signature")

            if signature:
                # 构建需要存入 L2 的纯净 Payload
                # 这里必须存入 start_row/col，因为将来查缓存时要靠它们算偏移量！
                payload = {
                    "code": entry["code"],
                    "start_row": entry.get("start_row", 1),
                    "start_col": entry.get("start_col", 1),
                    "header_map": entry.get("header_map", {})
                }

                try:
                    # 直接调用底层的 l2_set，以子表 Hash 为 Key 写入
                    l2_set(sheet_name, signature, payload)
                    saved_count += 1
                    saved_titles.append(title)
                except Exception as e:
                    print(f"❌ 保存子表 '{title}' 缓存失败：{e}")

    # ── 3. 返回结果 ──
    if saved_count > 0:
        reason = f"成功将 {saved_count} 个子表的专属代码写入 L2 缓存：{', '.join(saved_titles)}"
        print(f"💾 {reason}")
        return {
            "cache_saved": True,
            "cache_save_reason": reason
        }
    else:
        # 可能是全部命中了 L1/L2 所以没有新代码需要保存，或者 SA 拆分失败
        return {
            "cache_saved": False,
            "cache_save_reason": "没有检测到需要新保存的未命中子表代码。"
        }

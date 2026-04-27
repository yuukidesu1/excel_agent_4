"""
nodes/kv_cache_save.py — KV 缓存保存节点

职责：
    1. 构建 KV 模型缓存 Key：Hash(kv_list + 模板结构特征)
    2. 保存验证通过的抽取代码到 L1 和 L2 缓存
    3. 存储内容包含：kv_list, generated_code, structure_signature, confidence_score

缓存 Key 设计：
    - 引入 sheet 维度特征（max_row, max_col）避免 Hash 碰撞
    - signature = sha256(sorted(kv_list) + sheet_name:max_row:max_col)[:16]
"""

import hashlib
import json
from datetime import datetime
from typing import Dict, Any, Optional

from excel_agent.state import AgentState
from excel_agent.cache_manager import l1_set, l2_set


# ==================== KV 缓存 Key 构建 ====================

def build_kv_cache_key(sheet_name: str, kv_list: list, max_row: int, max_col: int,
                        subtable_title: str = None, scope: Dict = None) -> str:
    """
    构建 KV 模型缓存 Key

    设计原则：
    - Hash(kv_list + 模板结构特征)
    - 避免 Hash 碰撞：引入 sheet 维度特征
    - 子表 KV 场景：增加 subtable_title + scope hash
    """
    key_str = "|".join(sorted(kv_list))
    structure_hint = f"{sheet_name}:{max_row}:{max_col}"
    scope_tag = ""
    if subtable_title:
        scope_str = f"{scope.get('start_row', 0)}-{scope.get('end_row', 0)}-{scope.get('start_col', 0)}-{scope.get('end_col', 0)}" if scope else ""
        scope_tag = f"|{subtable_title}|{hashlib.sha256(scope_str.encode()).hexdigest()[:8]}"
    raw_key = f"{key_str}||{structure_hint}{scope_tag}"
    return hashlib.sha256(raw_key.encode()).hexdigest()[:16]


def compute_kv_structure_signature(kv_list: list, sheet_structure: Dict[str, Any],
                                     subtable_title: str = None, scope: Dict = None) -> str:
    """
    计算 KV 结构指纹签名（用于 L2 缓存）

    签名应满足：
    - 相同的表格结构 → 相同的签名
    - 不同的表格结构 → 不同的签名
    - 子表 KV 场景：增加 subtable_title 和 scope 信息
    """
    signature_data = {
        "kv_list_sorted": sorted(kv_list),
        "merge_cell_count": len(sheet_structure.get("merged_cells_info", [])),
        "non_empty_cell_count": len(sheet_structure.get("non_empty_cells", [])),
    }

    if subtable_title:
        signature_data["subtable_title"] = subtable_title
        if scope:
            signature_data["scope"] = scope

    serialized = json.dumps(signature_data, sort_keys=True, ensure_ascii=False)
    return hashlib.sha256(serialized.encode()).hexdigest()[:16]


# ==================== KV 缓存保存节点 ====================

def kv_cache_save_node(state: AgentState) -> dict:
    """
    KV 模式缓存保存节点

    保存验证通过的抽取代码到 L1 和 L2 缓存
    """
    config = state.get("config", {})
    sheet_name = config.get("sheet_name", "")
    kv_list = config.get("kv_list", [])
    excel_path = config.get("excel_path", "")

    generated_code = state.get("generated_code", [])
    quality_score = state.get("quality_score", 0.0)
    sheet_structure = state.get("sheet_structure", {})

    if not generated_code:
        return {"cache_saved": False, "reason": "没有生成代码"}

    # 取第一段代码（当前是单代码模式）
    code = generated_code[0] if isinstance(generated_code, list) else generated_code

    # 构建缓存 Key
    max_row = sheet_structure.get("max_row", 0)
    max_col = sheet_structure.get("max_col", 0)

    # 从缓存 entries 中提取子表信息（子表 KV 场景）
    cache_state = state.get("cache", {})
    entries = cache_state.get("entries", {})
    subtable_title = None
    scope = None
    for title, entry in entries.items():
        if entry.get("extract_mode") == "kv" or entry.get("scope"):
            subtable_title = title
            scope = entry.get("scope")
            break

    l1_key = build_kv_cache_key(sheet_name, kv_list, max_row, max_col,
                                 subtable_title=subtable_title, scope=scope)
    structure_signature = compute_kv_structure_signature(kv_list, sheet_structure,
                                                          subtable_title=subtable_title, scope=scope)

    # 准备缓存数据
    cache_data = {
        "kv_list": kv_list,
        "generated_code": code,
        "structure_signature": structure_signature,
        "confidence_score": quality_score,
        "created_at": datetime.now().isoformat(),
        "template_hash": l1_key
    }

    # 保存到 L1 缓存（完全匹配）
    try:
        l1_set(
            excel_path=excel_path,
            sheet_name=sheet_name,
            subtable_titles=kv_list,  # 复用 subtable_titles 字段
            data=cache_data
        )
        print(f"✅ KV 缓存 L1 保存成功：key={l1_key}")
    except Exception as e:
        print(f"⚠ KV 缓存 L1 保存失败：{e}")

    # 保存到 L2 缓存（结构指纹匹配）
    try:
        l2_set(
            sheet_name=sheet_name,
            structure_signature=structure_signature,
            data=cache_data
        )
        print(f"✅ KV 缓存 L2 保存成功：signature={structure_signature}")
    except Exception as e:
        print(f"⚠ KV 缓存 L2 保存失败：{e}")

    return {
        "cache_saved": True,
        "l1_key": l1_key,
        "structure_signature": structure_signature
    }

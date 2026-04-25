"""
nodes/pre_structure_analyzer.py — 前置结构分析器 (PSA)

职责：
    1. 基于 parse_node 输出的非空单元格和合并单元格字典，构建 O(1) 查询的 Map。
    2. 基于子表标题，精确定位物理锚点 (start_row, start_col)。
    3. 支持解析双行/多级表头 (如 "RF Module||Type") 的纵深结构。
    4. 仅提取目标列/行表头的相对位置和跨度作为“结构指纹”(Cache Key)，极大增强抗噪能力。
    5. 组装并初始化 CacheState。

注意：本节点纯代码执行，不调用 LLM。直接复用内存数据，速度极快。
"""

import hashlib
import json
import re
from typing import Dict, List, Any, Optional

from excel_agent.state import AgentState, CacheState, SubtableCacheEntry


# ==================== 1. 基础工具函数 ====================

def _normalize(text: Any) -> str:
    """归一化字符串：转小写、去标点、去空白"""
    if text is None: return ""
    return re.sub(r'[^\w\u4e00-\u9fff]', '', str(text).lower())

def _normalize_header(text: str) -> str:
    """针对包含层级符号 || 的表头进行归一化"""
    if not text: return ""
    return "||".join([_normalize(p) for p in str(text).split("||")])

def _detect_layout(found_headers: List[Dict]) -> str:
    """基于表头几何区域跨度区分： 仅列(横表) vs 仅行(纵表)"""
    if not found_headers: return "仅列"

    rel_rows = [h["rel_row"] for h in found_headers]
    rel_cols = [h["rel_col"] for h in found_headers]

    row_span = max(rel_rows) - min(rel_rows) + 1
    col_span = max(rel_cols) - min(rel_cols) + 1

    # 行跨度 > 列跨度 -> 表头竖着排 -> 仅行(纵表)
    if row_span > col_span:
        return "仅行"
    else:
        return "仅列"


# ==================== 2. 核心结构分析引擎 ====================

def _build_maps(non_empty_cells: List[Dict], merged_cells_info: List[Dict]) -> tuple[Dict, Dict]:
    """将 parse_node 的输出转化为 O(1) 的查询字典"""
    cell_map = {(c["row"], c["col"]): c["value"] for c in non_empty_cells}

    merged_map = {}
    for m in merged_cells_info:
        top_r, top_c = m["min_row"], m["min_col"]
        val = m["value"]
        r_span = m["max_row"] - m["min_row"] + 1
        c_span = m["max_col"] - m["min_col"] + 1

        for r in range(m["min_row"], m["max_row"] + 1):
            for c in range(m["min_col"], m["max_col"] + 1):
                merged_map[(r, c)] = {
                    "val": val,
                    "top_left": (top_r, top_c),
                    "r_span": r_span,
                    "c_span": c_span,
                    "is_top_left": (r == top_r and c == top_c)
                }
    return cell_map, merged_map


def _get_header_path(cell_map: Dict, merged_map: Dict, start_r: int, start_c: int, depth: int, direction: str) -> str:
    """向下或向右探索多级表头路径"""
    path = []
    seen_merge_ids = set()

    for i in range(depth):
        r = start_r + i if direction == "col" else start_r
        c = start_c if direction == "col" else start_c + i

        m_info = merged_map.get((r, c))
        if m_info:
            mid = m_info["top_left"]
            val = m_info["val"]
            if mid not in seen_merge_ids and val:
                path.append(_normalize(val))
                seen_merge_ids.add(mid)
        else:
            val = cell_map.get((r, c))
            if val:
                path.append(_normalize(val))
                seen_merge_ids.add((r, c))

    return "||".join(path)


def _scan_headers(cell_map: Dict, merged_map: Dict, s_row: int, e_row: int, s_col: int, e_col: int,
                  target_headers: List[str], direction: str, origin_r: int, origin_c: int) -> List[Dict]:
    """在指定范围内扫描目标表头，返回它们的相对坐标和合并跨度"""
    if not target_headers: return []

    max_depth = max([len(th.split("||")) for th in target_headers])
    norm_targets = [_normalize_header(th) for th in target_headers]

    if direction == "col":
        for r in range(s_row, e_row + 1):
            found = []
            seen_top_lefts = set()  # 用于去重同一合并单元格

            for c in range(s_col, e_col + 1):
                path = _get_header_path(cell_map, merged_map, r, c, max_depth, "col")
                if not path: continue

                for th in norm_targets:
                    if path == th or path.startswith(th + "||"):
                        m_info = merged_map.get((r, c), {"r_span": 1, "c_span": 1, "is_top_left": True, "top_left": (r, c)})
                        top_left = m_info.get("top_left", (r, c))

                        # 使用 top_left 去重，避免同一合并单元格被多次记录
                        if top_left in seen_top_lefts:
                            continue

                        found.append({
                            "target": th,
                            "rel_row": r - origin_r,
                            "rel_col": c - origin_c,
                            "row_span": m_info["r_span"],
                            "col_span": m_info["c_span"]
                        })
                        seen_top_lefts.add(top_left)
                        break
            if found:
                # 按列排序，确保表头顺序正确
                found.sort(key=lambda x: x["rel_col"])
                return found

    elif direction == "row":
        for c in range(s_col, e_col + 1):
            found = []
            seen_top_lefts = set()  # 用于去重同一合并单元格

            for r in range(s_row, e_row + 1):
                path = _get_header_path(cell_map, merged_map, r, c, max_depth, "row")
                if not path: continue

                for th in norm_targets:
                    if path == th or path.startswith(th + "||"):
                        m_info = merged_map.get((r, c), {"r_span": 1, "c_span": 1, "is_top_left": True, "top_left": (r, c)})
                        top_left = m_info.get("top_left", (r, c))

                        # 使用 top_left 去重，避免同一合并单元格被多次记录
                        if top_left in seen_top_lefts:
                            continue

                        found.append({
                            "target": th,
                            "rel_row": r - origin_r,
                            "rel_col": c - origin_c,
                            "row_span": m_info["r_span"],
                            "col_span": m_info["c_span"]
                        })
                        seen_top_lefts.add(top_left)
                        break
            if found:
                found.sort(key=lambda x: x["rel_row"])
                return found

    return []


# ==================== 3. 主流程 ====================


def _check_headers_kv_layout(
    cell_map: Dict,
    merged_map: Dict,
    start_row: int,
    start_col: int,
    headers: List[str],
    search_range: int = 30
) -> tuple[bool, List[str]]:
    """
    检查用户配置的 headers 是否呈现 KV 布局特征

    核心逻辑：
    1. 在 Excel 中查找每个 header 的位置
    2. 检查这些 header 是否都在同一列或相邻列（Key 列）
    3. **关键区分**：表头行 vs KV 数据行

    Args:
        cell_map: 单元格值映射
        merged_map: 合并单元格映射
        start_row, start_col: 子表起始位置
        headers: 用户配置的 headers
        search_range: 向下搜索范围

    Returns:
        (is_kv_layout, matched_headers)
        - is_kv_layout: 是否为 KV 布局
        - matchedHeaders: 匹配到的 headers（用于确认哪些 Keys 存在）
    """
    if not headers:
        return False, []

    # 用于记录找到的 header 及其位置
    found_headers = []

    for header in headers:
        header_norm = _normalize(header)
        if not header_norm:
            continue

        # 在区域内搜索该 header
        found = False
        for r in range(start_row, start_row + search_range):
            for c in range(start_col, start_col + 10):  # 假设 Key 在左侧 10 列内
                val = cell_map.get((r, c))
                if val and _normalize(val) == header_norm:
                    # 检查同一行是否有其他单元格
                    same_row_cells = []
                    for check_c in range(start_col, start_col + 25):  # 扫描更宽的范围
                        if check_c != c:
                            neighbor_val = cell_map.get((r, check_c))
                            if neighbor_val:
                                same_row_cells.append(neighbor_val)

                    # 检查下方是否有其他行（KV 邻居特征）
                    has_row_below = any(
                        cell_map.get((r + offset, c)) is not None
                        for offset in range(1, 4)
                    )

                    # 🚀 关键区分：表头行 vs KV 数据行
                    is_like_table_header = False

                    # 判断 1：如果 header 本身很短（<20 字符）且是全大写，则是表头
                    if len(header) < 20 and header.upper() == header:
                        is_like_table_header = True

                    # 判断 2：检查同一行的其他单元格
                    if len(same_row_cells) >= 1:
                        # 统计同一行中有多少个"长文本 Key 特征"的单元格
                        # KV 布局：同一行的其他单元格也是长文本（其他 Key）或短 Value（Yes/No/数字）
                        # 表头行：同一行的其他单元格也是短表头（<20 字符，不是 Value 特征）

                        long_keys_in_row = [
                            cell for cell in same_row_cells
                            if len(str(cell)) > 15  # 长文本，像 Key
                        ]

                        short_header_like = [
                            cell for cell in same_row_cells
                            if len(str(cell)) < 20 and
                               str(cell).upper() == str(cell) and  # 全大写
                               not str(cell).lower() in ['yes', 'no', 'true', 'false', 'n/a', ''] and
                               not str(cell).isdigit()
                        ]

                        # 🚀 修改：如果同一行有≥3 个全大写短表头，则是表头行（优先于长文本判断）
                        if len(short_header_like) >= 3:
                            is_like_table_header = True
                        # 如果同一行有多个长文本 Key，且短表头<3 个，说明这是 KV 数据行（多列 KV 布局）
                        elif len(long_keys_in_row) >= 1:
                            is_like_table_header = False  # 这是 KV 布局的其他 Key

                    found_headers.append({
                        "header": header,
                        "row": r,
                        "col": c,
                        "row_has_other_cells": len(same_row_cells) > 0,
                        "has_row_below": has_row_below,
                        "is_like_table_header": is_like_table_header
                    })
                    found = True
                    break
            if found:
                break

    if len(found_headers) == 0:
        return False, []

    # 判断是否符合 KV 布局特征
    kv_like_count = sum(
        1 for h in found_headers
        if (h["row_has_other_cells"] or h["has_row_below"]) and not h["is_like_table_header"]
    )

    # 2. 找到的 headers 的列位置应该比较集中（都在同一列或相邻列）
    if len(found_headers) >= 2:
        cols = [h["col"] for h in found_headers]
        col_span = max(cols) - min(cols) + 1
        is_concentrated = col_span <= 3
    else:
        is_concentrated = True

    # 🚀 如果超过 50% 的 headers 像表头，则不是 KV 布局
    header_like_count = sum(1 for h in found_headers if h["is_like_table_header"])
    if len(found_headers) > 0 and header_like_count / len(found_headers) >= 0.5:
        return False, []

    # 判定：找到至少 1 个 header，且符合 KV 特征，且不像表头
    is_kv = (
        len(found_headers) >= 1 and
        (kv_like_count / len(found_headers) >= 0.5 or len(found_headers) == 1) and
        is_concentrated
    )

    if is_kv:
        return True, [h["header"] for h in found_headers]
    return False, []


def pre_structure_analyzer_node(state: AgentState) -> dict:
    """PSA 主节点逻辑（支持多子表 KV 布局自动检测）"""
    config = state.get("config", {})

    # 🚀 用户明确意图优先：如果配置了 kv_list，直接使用，不进行任何检测
    global_kv_list = config.get("kv_list", [])
    extract_type = config.get("extract_type", "table")

    # 如果用户已经配置了 kv_list（全局或子表级别），直接切换到 KV 模式
    if global_kv_list:
        config["extract_type"] = "kv"
        config["kv_list"] = global_kv_list  # 确保 config 中也设置，供 downstream 节点使用
        # 将 kv_list 传递给 downstream 节点
        return {
            "cache": {
                "hit": False,
                "missed": True,
                "kv_auto_detected": False,  # 用户明确配置，非自动检测
                "kv_list": global_kv_list
            },
            "config": config
        }

    # KV 模式（配置文件指定 extract_type="kv"）直接跳过 PSA 分析，KV 缓存不依赖 subtable_title 建档
    if extract_type and extract_type.lower() == "kv":
        return {"cache": state.get("cache", {})}

    subtable_titles = config.get("subtable_titles", [])
    st = state.get("sheet_structure", {})

    # 初始化 CacheState
    cache_state: CacheState = state.get("cache") or {
        "entries": {},
        "all_cached": False,
        "partial_cached": False,
        "missed_subtables": [],
        "analyzer_skipped": False,
        "analyzer_skip_reason": None
    }

    if not st or not subtable_titles:
        return {"cache": cache_state}

    max_row = st.get("max_row", 1000)
    max_col = st.get("max_col", 100)
    cell_map, merged_map = _build_maps(st.get("non_empty_cells", []), st.get("merged_cells_info", []))

    entries: Dict[str, SubtableCacheEntry] = {}

    # 用于收集所有子表中检测到的 KV 配置
    kv_subtables: Dict[str, List[str]] = {}  # {title: kv_list}

    for target_title in subtable_titles:
        norm_target_title = _normalize(target_title)
        sub_info = None

        # ── 1. 定位物理锚点 (寻找标题单元格) ──
        for c_info in st.get("non_empty_cells", []):
            if norm_target_title in _normalize(c_info["value"]):
                r, c = c_info["row"], c_info["col"]
                m_info = merged_map.get((r, c), {"r_span": 1, "c_span": 1})
                sub_info = {
                    "start_row": r,
                    "start_col": c,
                    "title_row_span": m_info["r_span"],
                    "title_col_span": m_info["c_span"]
                }
                break

        if not sub_info:
            continue

        sr, sc = sub_info["start_row"], sub_info["start_col"]
        tr_span = sub_info["title_row_span"]

        # ── 2. 解析用户配置的目标字段 ──
        headers = []

        target_configs_dict = config.get("subtable_configs") or {}
        tc_def = target_configs_dict.get(target_title)

        if isinstance(tc_def, dict):
            headers = tc_def.get("headers", [])
        elif isinstance(tc_def, list):
            # 列表结构直接作为 headers
            headers = [str(item) for item in tc_def]

        # ── 3. 扫描目标表头提取特征 ──
        col_h = []
        row_h = []

        # 扫描列表头
        if headers:
            col_h = _scan_headers(
                cell_map, merged_map,
                sr, min(sr + tr_span + 15, max_row),
                sc, max_col,
                headers, "col", sr, sc
            )

        # ── 3.5 KV 布局自动检测 ──
        # 即使用户配置成了"表格模式"，如果实际是 KV 布局，自动切换到 KV 模式
        kv_layout_detected = False
        extracted_kv_list: List[str] = []

        # 尝试从用户配置的 headers 中检测 KV 布局
        if headers:
            # 使用新的检查逻辑：基于用户配置的 headers 判断
            is_kv, matched_headers = _check_headers_kv_layout(
                cell_map, merged_map,
                sr, sc,
                headers,
                search_range=30
            )

            if is_kv:
                kv_layout_detected = True
                # 使用用户配置的 headers（匹配到的）作为 kv_list
                extracted_kv_list = headers  # 使用全部配置的 headers，LLM 会处理找不到的情况

                # 记录该子表的 KV 配置
                kv_subtables[target_title] = extracted_kv_list

                # 记录日志
                print(f"🔄 [PSA 节点] 检测到子表 '{target_title}' 为 KV 布局")
                print(f"   用户配置的 Headers: {len(headers)} 个")
                print(f"   匹配到的 Keys: {len(matched_headers)} 个")

        # 如果检测到 KV 布局，跳过该子表的表格模式建档
        if kv_layout_detected:
            continue

        # ── 4. 计算精准结构指纹 (Signature) ──

        fingerprint_data = {
            "title_pattern": norm_target_title,
            "headers": sorted(col_h, key=lambda x: (x["rel_row"], x["rel_col"]))
        }

        signature = hashlib.sha256(json.dumps(fingerprint_data, sort_keys=True).encode()).hexdigest()[:16]

        # ── 5. 组装建档 ──
        layout_en = "vertical"

        entries[target_title] = {
            "target_title": target_title,
            "signature": signature,
            "start_row": sr,
            "start_col": sc,
            "layout_type": layout_en,

            "cache_hit": False,
            "cache_level": "miss",
            "l2_row_offset": 0,
            "l2_col_offset": 0,
            "code": None,
            "header_map": {  # 填入 PSA 识别的表头信息，供 code_gen 节点使用
                "headers": col_h,  # List[Dict: {target, rel_row, rel_col, row_span, col_span}]
            },
            "extracted_data": None
        }

    # ── 处理 KV 子表 ──
    # 如果有任何子表被检测为 KV 布局，切换到 KV 模式
    if kv_subtables:
        config["extract_type"] = "kv"
        # 合并所有 KV 子表的 kv_list
        all_kv_keys = []
        for title, keys in kv_subtables.items():
            all_kv_keys.extend(keys)
        config["kv_list"] = all_kv_keys

        # 如果是多子表 KV 模式，需要特殊处理
        if len(kv_subtables) > 1:
            print(f"🔄 [PSA 节点] 检测到 {len(kv_subtables)} 个 KV 子表，将合并处理")

        print(f"🔄 [PSA 节点] 切换到 KV 模式，共 {len(all_kv_keys)} 个 Keys")

        return {
            "cache": {
                "hit": False,
                "missed": True,
                "kv_auto_detected": True,
                "kv_list": all_kv_keys,
                "kv_subtables": kv_subtables  # 记录哪些子表是 KV 模式
            },
            "config": config
        }

    cache_state["entries"] = entries
    cache_state["missed_subtables"] = list(entries.keys())

    return {"cache": cache_state}
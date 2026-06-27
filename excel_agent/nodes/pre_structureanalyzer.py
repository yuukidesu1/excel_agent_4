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

from Excel_Agent.excel_agent.state import AgentState, CacheState, SubtableCacheEntry


# ==================== 1. 基础工具函数 ====================

def _normalize(text: Any) -> str:
    """归一化字符串：转小写、去标点、去空白"""
    if text is None: return ""
    return re.sub(r'[^\w\u4e00-\u9fff]', '', str(text).lower())

def _normalize_header(text: str) -> str:
    """针对包含层级符号 || 的表头进行归一化"""
    if not text: return ""
    return "||".join([_normalize(p) for p in str(text).split("||")])


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

def _filter_valid_headers(raw_headers: List[Dict], max_span: int = 5) -> List[Dict]:
    """表头空间清洗：去重并剔除越界离群点"""
    if not raw_headers:
        return []

    unique_map = {}
    for h in raw_headers:
        target = h["target"]
        distance = h["rel_row"] + h["rel_col"]
        if target not in unique_map or distance < unique_map[target][0]:
            unique_map[target] = (distance, h)

    deduped_headers = [item[1] for item in unique_map.values()]
    if not deduped_headers:
        return []

    min_r = min(h["rel_row"] for h in deduped_headers)
    min_c = min(h["rel_col"] for h in deduped_headers)

    valid_headers = []
    for h in deduped_headers:
        if h["rel_row"] <= min_r + max_span and h["rel_col"] <= min_c + max_span:
            valid_headers.append(h)

    return valid_headers

def _detect_layout(found_headers: List[Dict]) -> str:
    """基于表头几何跨度区分: 仅列(横表) vs 仅行(纵表)"""
    if not found_headers:
        return "仅列"  # 兜底默认横表

    rel_rows = [h["rel_row"] for h in found_headers]
    rel_cols = [h["rel_col"] for h in found_headers]

    row_span = max(rel_rows) - min(rel_rows) + 1
    col_span = max(rel_cols) - min(rel_cols) + 1

    # 行跨度 > 列跨度 -> 表头竖着排 -> 仅行 (纵表)
    if row_span > col_span:
        return "仅行"
    else:
        return "仅列"

# ==================== 3. 主流程 ====================

def pre_structure_analyzer_node(state: AgentState) -> dict:
    """PSA 主节点逻辑"""
    config = state.get("config", {})
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

    for target_title in subtable_titles:
        norm_target_title = _normalize(target_title)
        subtable_info = config.get('subtable_configs')
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

        headers = subtable_info[target_title].get("headers")
        if sub_info:
            sr, sc = sub_info["start_row"], sub_info["start_col"]
            tr_span = sub_info["title_row_span"]
            tc_span = sub_info["title_col_span"]

            # 双向扫描表头
            col_h = _scan_headers(
                cell_map, merged_map,
                sr, min(sr + tr_span + 15, max_row),
                sc, max_col,
                headers, "col", sr, sc
            )
            row_h = _scan_headers(
                cell_map, merged_map,
                sr, max_row,
                sc, min(sc + tc_span + 15, max_col),
                headers, "row", sr, sc
            )

            # 合并去重并过滤离群点
            raw_combined_headers = col_h + row_h
            all_found_headers = _filter_valid_headers(raw_combined_headers, max_span=4)

            # 执行动态版式推断
            final_layout = _detect_layout(all_found_headers)

            # 生成 Hash 结构指纹
            fingerprint_data = {
                "title_pattern": norm_target_title,
                "layout": final_layout,  # 这里的 layout 已经动态算出了 "仅行" 或 "仅列"
                "col_headers": sorted(all_found_headers, key=lambda x: (x["rel_row"], x["rel_col"])),
                "row_headers": []
            }

            signature = hashlib.sha256(json.dumps(fingerprint_data, sort_keys=True).encode()).hexdigest()[:16]

            # ── 5. 组装建档 ──
            layout_en = "vertical" if final_layout == "仅列" else ("horizontal" if final_layout == "仅行" else "cross")

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
                "col_headers": col_h,  # List[Dict: {target, rel_row, rel_col, row_span, col_span}]
                "row_headers": row_h,
            },
            "extracted_data": None
        }

    cache_state["entries"] = entries
    cache_state["missed_subtables"] = list(entries.keys())

    return {"cache": cache_state}
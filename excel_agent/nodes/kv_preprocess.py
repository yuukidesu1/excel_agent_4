import difflib

import openpyxl
import re
import hashlib
import json

from Excel_Agent.excel_agent.state import AgentState


def _find_sheet(wb, requested: str) -> str:
    """大小写不敏感匹配 Sheet 名"""
    if requested in wb.sheetnames:
        return requested
    req = requested.strip().upper()
    for name in wb.sheetnames:
        if name.strip().upper() == req:
            return name
    return requested

# 新增加结构指纹构建
def _generate_kv_signature(bboxes_dict, merged_cells_info):
    """
    生成 16 为 KV 抽取缓存哈希
    :param bboxes_dict: kv_preprocess 产生的 named_bboxes
    :param merged_cells_info: kv_preprocess 产生的合并单元格信息
    :return:
    """

    # 1. 构建宏观拓扑序列 (提取所有合法包围和的相对顺序特征)
    # 不记录绝对的 R（行号），只记录 C（列号宽度）和它们出现的先后顺序
    topology_sequence = []

    # 将 bboxes 按照从上到下的顺序排列
    # 过滤掉 fallback
    valid_boxes = {k: v for k, v in bboxes_dict.items() if k != "fallback"}
    sorted_items = sorted(valid_boxes.items(), key=lambda item: item[1]["min_r"])

    for key, box in sorted_items:
        # 特征：包围盒跨越了多少列？这影响横向/纵向代码逻辑
        width_span = box["max_c"] - box["min_c"]
        # 特征：这是一个多行的区域表头，还是单行的全局 Key？
        height_span = box["max_r"] - box["min_r"]

        topology_sequence.append({
            "key": key.lower(),  # 统一小写，忽略配置时的大小写差异
            "w_span": width_span,
            "is_multi_row": height_span > 0
        })

    # 2. 构建合并单元格特征 (防格式突变)
    # 取前 5 个有意义的合并单元格的跨度，作为模板验证的防伪标签
    merge_features = []
    # 从 "R5C1~R10C2:"[空合并]"" 这样的字符串中提取行列跨度
    import re
    for m_str in merged_cells_info[:5]:
        if m_str == "...": continue
        match = re.search(r'R(\d+)C(\d+)~R(\d+)C(\d+)', m_str)
        if match:
            min_r, min_c, max_r, max_c = map(int, match.groups())
            merge_features.append(f"W{max_c - min_c}_H{max_r - min_r}")

    # 3. 组装最终指纹载体
    fingerprint_data = {
        "topology": topology_sequence,
        "merge_layout": merge_features
    }

    # 4. SHA256 哈希计算
    signature = hashlib.sha256(
        json.dumps(fingerprint_data, sort_keys=True).encode('utf-8')
    ).hexdigest()[:16]

    return signature, fingerprint_data

def kv_preprocess_node(state: AgentState) -> dict:
    config = state.get("config", {})
    excel_path = config.get("excel_path")
    sheet_name_req = config.get("sheet_name")
    kv_list = config.get("kv_list", [])

    # 1. 拆解：使用字典映射保存原始大小写
    global_keywords_map = {}  # {小写: 原始大小写}
    local_headers_map = {}  # {小写: 原始大小写}
    region_paths_set = set()

    for key in kv_list:
        if "||" in key:
            parts = [p.strip() for p in key.split("||")]
            region_paths_set.add(tuple(parts[:-1]))
            # 局部表头：只取最左边的作为严格雷达词
            original_header = parts[0]
            local_headers_map[original_header.lower()] = original_header
        else:
            # 全局 KV：作为模糊雷达词
            original_global = key.strip()
            global_keywords_map[original_global.lower()] = original_global

    parsed_region_paths = [list(path) for path in region_paths_set]

    try:
        wb = openpyxl.load_workbook(excel_path, data_only=True)
        sheet_name = _find_sheet(wb, sheet_name_req)
        if sheet_name not in wb.sheetnames:
            return {"errors": [f"Can not find sheet: {sheet_name_req}"]}
        ws = wb[sheet_name]
    except Exception as e:
        return {"errors": [f"preprocessing failed: {str(e)}"]}

    max_row = ws.max_row
    max_col = ws.max_column

    merged_ranges = list(ws.merged_cells.ranges)
    merged_map = {}
    merged_coords = set()
    for m in merged_ranges:
        for r in range(m.min_row, m.max_row + 1):
            for c in range(m.min_col, m.max_col + 1):
                merged_coords.add((r, c))
        merged_map[(m.min_row, m.min_col)] = m

    def is_col_empty(min_r, max_r, col_idx):
        for r in range(min_r, max_r + 1):
            val = ws.cell(row=r, column=col_idx).value
            if val is not None and str(val).strip() != "":
                return False
        return True

    # =====================================================================
    # 【核心压缩 Step 1：包围盒探测雷达 (区分严格与模糊)】
    # =====================================================================
    bboxes = []
    named_bboxes = {}

    for r in range(1, max_row + 1):
        for c in range(1, max_col + 1):
            if (r, c) in merged_coords and (r, c) not in merged_map:
                continue

            cell_val = ws.cell(row=r, column=c).value
            if cell_val is None or str(cell_val).strip() == "":
                continue

            clean_val_lower = str(cell_val).strip().lower()

            # ---------------------------------------------------------
            # 情况 1：命中全局 KV (开启 difflib 滑动窗口容错)
            # ---------------------------------------------------------
            matched_gw_lower = None
            for gw in global_keywords_map:
                if global_keywords_map[gw] in named_bboxes:
                    continue

                pattern = rf'(?<![a-z0-9]){re.escape(gw)}(?![a-z0-9])'

                if re.search(pattern, clean_val_lower):
                    matched_gw_lower = gw
                    break
                else:
                    gw_len = len(gw)
                    val_len = len(clean_val_lower)

                    if gw_len >= 3 and val_len >= gw_len and val_len <= gw_len *1.5:
                        for i in range(val_len - gw_len + 1):
                            sub = clean_val_lower[i:i + gw_len]
                        if difflib.SequenceMatcher(None, gw, sub).ratio() > 0.85:
                            matched_gw_lower = gw
                            break


                if matched_gw_lower:
                    break

            if matched_gw_lower:
                original_header_key = global_keywords_map[matched_gw_lower]
                if original_header_key not in named_bboxes:
                    box = {"min_r": max(1, r - 5), "max_r": min(max_row, r + 5), "min_c": 1, "max_c": max_col}
                    bboxes.append(box)
                    named_bboxes[original_header_key] = box

            # ---------------------------------------------------------
            # 情况 2：命中局部 KV 表头 (绝对严格匹配，绝不容错)
            # ---------------------------------------------------------
            matched_lh_lower = next((lh for lh in local_headers_map if lh == clean_val_lower), None)
            if matched_lh_lower:
                m_info = merged_map.get((r, c))
                if m_info and m_info.max_row > m_info.min_row:
                    stop_c = m_info.max_col
                    while stop_c < max_col:
                        if is_col_empty(m_info.min_row, m_info.max_row, stop_c + 1):
                            break
                        stop_c += 1
                    box = {"min_r": m_info.min_row, "max_r": m_info.max_row, "min_c": c, "max_c": stop_c}
                else:
                    box = {"min_r": r, "max_r": min(max_row, r + 15), "min_c": c, "max_c": max_col}

                bboxes.append(box)
                original_local_header = local_headers_map[matched_lh_lower]
                named_bboxes[original_local_header] = box  # 绑定原始局部表头 Key

    # 兜底
    if not bboxes:
        box = {"min_r": 1, "max_r": min(30, max_row), "min_c": 1, "max_c": max_col}
        bboxes.append(box)
        named_bboxes["fallback"] = box

    # =====================================================================
    # 【核心压缩 Step 1.5：紧凑化包围盒 (Tight Fit Algorithm)】
    # =====================================================================
    for b in bboxes:
        actual_max_r = b["min_r"]
        actual_max_c = b["min_c"]
        for r in range(b["min_r"], b["max_r"] + 1):
            for c in range(b["min_c"], b["max_c"] + 1):
                val = ws.cell(row=r, column=c).value
                has_data = val is not None and str(val).strip() != ""
                if has_data or (r, c) in merged_coords:
                    actual_max_r = max(actual_max_r, r)
                    actual_max_c = max(actual_max_c, c)
        b["max_r"] = actual_max_r
        b["max_c"] = actual_max_c

    def in_any_bbox(r, c):
        for b in bboxes:
            if b["min_r"] <= r <= b["max_r"] and b["min_c"] <= c <= b["max_c"]:
                return True
        return False

    def is_rect_intersect_bboxes(min_r, max_r, min_c, max_c):
        for b in bboxes:
            if not (max_r < b["min_r"] or min_r > b["max_r"] or max_c < b["min_c"] or min_c > b["max_c"]):
                return True
        return False

    # =====================================================================
    # 【核心压缩 Step 2：收集与压缩 合并单元格】
    # =====================================================================
    merged_cells_info = []
    merged_ranges.sort(key=lambda m: (m.min_row, m.min_col))
    last_m_row = -1
    for m in merged_ranges:
        if is_rect_intersect_bboxes(m.min_row, m.max_row, m.min_col, m.max_col):
            if last_m_row != -1 and m.min_row > last_m_row + 1:
                merged_cells_info.append("...")
            top_val = ws.cell(row=m.min_row, column=m.min_col).value
            clean_val = str(top_val).replace("\n", " ").replace("\r", "").strip()[:100] if top_val else ""
            display_val = clean_val if clean_val else "[空合并]"
            merged_cells_info.append(f"R{m.min_row}C{m.min_col}~R{m.max_row}C{m.max_col}:\"{display_val}\"")
            last_m_row = max(last_m_row, m.max_row)

    # =====================================================================
    # 【核心压缩 Step 3：收集与压缩 详细内容单元格】
    # =====================================================================
    sample_cells = []
    last_r = -1
    for r in range(1, max_row + 1):
        row_has_data_in_bbox = False
        for c in range(1, max_col + 1):
            if in_any_bbox(r, c):
                if (r, c) in merged_coords:
                    continue
                cell_val = ws.cell(row=r, column=c).value
                if cell_val is None or str(cell_val).strip() == "":
                    clean_val = ""
                else:
                    clean_val = str(cell_val).replace("\n", " ").replace("\r", "").strip()
                    clean_val = clean_val[:50] + "..." if len(clean_val) > 50 else clean_val

                if not row_has_data_in_bbox and last_r != -1 and r > last_r + 1:
                    sample_cells.append("...")

                sample_cells.append(f"R{r}C{c}:\"{clean_val}\"")
                row_has_data_in_bbox = True
        if row_has_data_in_bbox:
            last_r = r

    sample_cells = sample_cells[:300]

    # 获取结构特征
    signature, fingerprint_data = _generate_kv_signature(named_bboxes, merged_cells_info)

    return {
        "kv_state": {
            "sheet_structure": {
                "max_row": max_row,
                "max_col": max_col,
                "merged_cells_info": merged_cells_info,
                "sample_cells": sample_cells,
                "bboxes": named_bboxes
            },
            "parsed_region_paths": parsed_region_paths,
            "cache": {
                "signature": signature,
                "fingerprint_data": fingerprint_data,
                "cache_hit" : False,
                "code": None
            }
        }
    }


# if __name__ == "__main__":
#     # 模拟环境测试
#     test_excel_path = r"D:\Projects\y00957025\Excel_Agent\ALX3090-WL-DBS-TDD2600-EXP-ALX3090-14-Region_B-Range_3-ONOFF-1_TSSR_2022 WL Site Survey Template(ALX3090) 标注样例.xlsx"
#     # test_excel_path = r"D:\Projects\y00957025\Excel_Agent\111017.xlsx"
#     # test_sheet_name = "Public"
#     test_sheet_name = "RF Data"
#
#     mock_state = {
#         "config": {
#             "excel_path": test_excel_path,
#             "sheet_name": test_sheet_name,
#             "kv_list": [
#                 "Cell C||Data||Antenna type /  model",
#                 "Cell C||Data||Shared With",
#                 "Cell C||Antenna Height (From Ground)",
#                 "Cell B'||Data||Antenna Sharing (Yes/No)",
#                 "Cell B'||Data||Shared With",
#                 "Cell B'||Data||Azmuith",
#                 "Cell B'||Antenna Height (From Ground)",
#                 "Cell B'||No Of Antenna Port ( total )",
#                 "Cell B'||RRU type/model (Jumper) 1800",
#                 "Cell B'||Data||NO of Free Feeders "
#             ]
#             # "kv_list": [
#             #     "Site Name",
#             #     "Site ID",
#             #     "ET Region",
#             #     "Latitude (N)"
#             # ]
#         }
#     }
#
#     print("\n🚀 开始执行 kv_preprocess_node...")
#     result = kv_preprocess_node(mock_state)
#     print("\n📊 节点输出结果：")
#     print(json.dumps(result, ensure_ascii=False, indent=4))
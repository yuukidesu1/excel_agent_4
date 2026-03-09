"""
excel_agent/cache.py — 分层缓存系统

目标：
  在已经见过一次场景的情况下，大幅提升提取速度，避免每次都调用 LLM。

缓存层级：
  Level 1: 完全匹配缓存 - (excel_file_hash, sheet_name, subtable_titles)
           命中后直接返回 cached generated_code，跳过 LLM

  Level 2: 结构指纹缓存 - (sheet_name, structure_signature)
           当子表结构相同但起始行偏移时，计算偏移量并调整代码

  Level 3: 子表模板缓存 - (sheet_name, normalized_title_pattern)
           抽象出通用提取模板，支持跨文件复用

使用策略：
  1. 每次提取前依次检查 L1 → L2 → L3
  2. 提取成功后，将结果缓存到所有适用层级
  3. 支持缓存淘汰策略（LRU）
"""

import json
import os
import hashlib
import threading
import re
from typing import Dict, Any, Optional, List, Tuple
from datetime import datetime
from pathlib import Path

CACHE_FILE = "extraction_cache.json"
_cache_lock = threading.Lock()

# 缓存配置
MAX_CACHE_ENTRIES = 500  # 最大缓存条目数
CACHE_TTL_DAYS = 30  # 缓存有效期（天）


def _compute_file_hash(file_path: str) -> str:
    """
    计算文件内容的 SHA256 哈希值
    只计算文件前 1MB + 后 10KB（平衡速度与准确性）
    """
    if not os.path.exists(file_path):
        return "file_not_found"

    hasher = hashlib.sha256()
    file_size = os.path.getsize(file_path)

    with open(file_path, 'rb') as f:
        # 读取文件开头 1MB
        head_size = min(1024 * 1024, file_size)
        hasher.update(f.read(head_size))

        # 如果文件大于 1MB + 10KB，再读取末尾 10KB
        if file_size > 1024 * 1024 + 1024 * 10:
            f.seek(-1024 * 10, 2)
            hasher.update(f.read(1024 * 10))

    return hasher.hexdigest()[:16]  # 取前 16 位缩短 key 长度

# 🚀 1. 新增辅助函数：对目标列配置进行哈希，防止列配置变更时误命中旧缓存
def _hash_target_columns(target_columns: Optional[Dict[str, List[Dict[str, Any]]]]) -> str:
    if not target_columns:
        return "none"
    # 将字典转为稳定排序的 JSON 字符串进行 hash
    col_str = json.dumps(target_columns, sort_keys=True)
    return hashlib.md5(col_str.encode('utf-8')).hexdigest()[:8]

def _compute_structure_signature(
    sheet_structure: Dict[str, Any],
    subtable_titles: List[str]
) -> str:
    """
    计算子表结构的指纹（纯结构特征，与具体内容无关，支持跨文件复用）

    重要设计变更：
    - 指纹计算包含 Sheet 中所有子表的结构，与 subtable_titles 无关
    - 只关注结构特征：行数、列数、合并模式、表头位置
    - 不关注具体单元格内容（列名值），因为相同结构的文件内容可能不同

    指纹包含：
    - 所有子表的标题（归一化）
    - 每个子表的行数
    - 每个子表的列数
    - 合并单元格的相对模式（位置、跨度）
    """
    signature_data = []

    # 处理所有子表，不只限于 subtable_titles
    for subtable in sheet_structure.get("potential_subtables", []):
        title = subtable.get("title", "")
        if not title:
            continue

        # 提取结构特征（与绝对行号无关）
        start_row = subtable.get("start_row", 0)
        end_row = subtable.get("end_row", 0)
        row_count = end_row - start_row + 1

        # 列特征：从 sample_cells 中提取列模式
        cells_in_subtable = [
            c for c in sheet_structure.get("non_empty_cells", [])
            if start_row <= c["row"] <= end_row
        ]

        # 提取所有出现的列号
        columns = sorted(set(c["col"] for c in cells_in_subtable))
        if not columns:
            continue

        min_col = columns[0]  # 起始列
        max_col = columns[-1]  # 结束列
        col_count = len(columns)
        col_span = max_col - min_col + 1  # 列跨度

        # 合并单元格模式（相对于子表起始行和起始列）
        # 这是最关键的结构特征
        merge_patterns = []
        for merged in sheet_structure.get("merged_cells_info", []):
            if start_row <= merged["min_row"] <= end_row:
                rel_row = merged["min_row"] - start_row
                rel_col = merged["min_col"] - min_col
                merge_patterns.append({
                    "rel_row": rel_row,
                    "rel_col": rel_col,
                    "col_span": merged["col_span"],
                    "row_span": merged["row_span"],
                    # 不存储 value，因为相同结构的文件合并单元格的值可能不同
                })

        # 排序以确保相同结构产生相同指纹
        merge_patterns.sort(key=lambda x: (x["rel_row"], x["rel_col"]))

        signature_data.append({
            "title_pattern": _normalize_string(title),
            "row_count": row_count,
            "col_count": col_count,
            "col_span": col_span,
            "merge_patterns": merge_patterns,
        })

    # 排序以确保相同结构产生相同指纹
    signature_data.sort(key=lambda x: x["title_pattern"])

    signature_str = json.dumps(signature_data, sort_keys=True)
    return hashlib.md5(signature_str.encode()).hexdigest()


def _normalize_string(s: str) -> str:
    """归一化字符串：转小写，移除数字前缀和特殊字符"""
    import re
    # 移除数字前缀（如 "4G Configuration" → "configuration"）
    s = re.sub(r'^\d+\s*', '', str(s))
    # 转小写并移除非字母字符
    return re.sub(r'[\W_]+', '', s.lower())


def _match_subtable_title(requested_title: str, cached_title: str) -> bool:
    """
    判断请求的子表标题是否与缓存的标题匹配。

    匹配规则：
    - 归一化后完全相同则匹配
    - 但如果数字前缀不同（如 "2G" vs "3G"），则不匹配

    例如：
    - "2G" 匹配 "2G Configuration" ✓
    - "3G" 匹配 "3G Configuration" ✓
    - "2G" 不匹配 "3G Configuration" ✗
    """
    # 提取数字前缀
    req_num_match = re.match(r'^(\d+)\s*', requested_title)
    cached_num_match = re.match(r'^(\d+)\s*', cached_title)

    req_num = req_num_match.group(1) if req_num_match else None
    cached_num = cached_num_match.group(1) if cached_num_match else None

    # 如果都有数字前缀，数字必须相同
    if req_num and cached_num and req_num != cached_num:
        return False

    # 归一化后比较（移除数字前缀后比较剩余部分）
    req_norm = _normalize_string(requested_title)
    cached_norm = _normalize_string(cached_title)

    # 请求标题归一化后应该包含在缓存标题归一化结果中，或是其子串
    return req_norm in cached_norm or cached_norm in req_norm or req_norm == cached_norm


def _calculate_row_offset(
    cached_structure: Dict[str, Any],
    current_structure: Dict[str, Any],
    subtable_titles: List[str]
) -> int:
    """
    计算新旧结构之间的行偏移量

    通过比较匹配子表的起始行差值来确定偏移
    参数:
        cached_structure: 缓存中的结构快照
        current_structure: 当前文件的结构
        subtable_titles: 请求的子表名称列表
    """
    for title in subtable_titles:
        # 在缓存结构中查找匹配的子表
        old_start = None
        for sub in cached_structure.get("potential_subtables", []):
            if _match_subtable_title(title, sub.get("title", "")):
                old_start = sub.get("start_row")
                break

        # 在当前结构中查找匹配的子表
        new_start = None
        for sub in current_structure.get("potential_subtables", []):
            if _match_subtable_title(title, sub.get("title", "")):
                new_start = sub.get("start_row")
                break

        if old_start is not None and new_start is not None:
            return new_start - old_start

    return 0  # 默认无偏移


def _load_cache() -> Dict[str, Any]:
    """安全加载缓存文件"""
    if not os.path.exists(CACHE_FILE):
        return {}

    try:
        with open(CACHE_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    except (json.JSONDecodeError, IOError):
        return {}


def _save_cache(cache: Dict[str, Any]):
    """安全保存缓存"""
    with open(CACHE_FILE, 'w', encoding='utf-8') as f:
        json.dump(cache, f, ensure_ascii=False, indent=2)


def _enforce_cache_limit(cache: Dict[str, Any], max_count: int = MAX_CACHE_ENTRIES):
    """LRU 缓存淘汰：淘汰最久未使用的条目"""
    if len(cache) <= max_count:
        return cache

    # 按 last_used 排序
    items = list(cache.items())
    items.sort(key=lambda x: x[1].get('last_used', ''))

    # 淘汰最旧的
    for key, _ in items[:len(cache) - max_count]:
        del cache[key]

    return cache


def _expire_old_cache(cache: Dict[str, Any], ttl_days: int = CACHE_TTL_DAYS) -> Dict[str, Any]:
    """删除过期的缓存条目"""
    now = datetime.now()
    to_remove = []

    for key, entry in cache.items():
        created_at = entry.get('created_at', '')
        if created_at:
            try:
                created_time = datetime.fromisoformat(created_at)
                age_days = (now - created_time).days
                if age_days > ttl_days:
                    to_remove.append(key)
            except ValueError:
                pass

    for key in to_remove:
        del cache[key]

    return cache


# ==================== Level 1: 完全匹配缓存 ====================

def get_level1_cache(
    excel_path: str,
    sheet_name: str,
    subtable_titles: List[str]
) -> Optional[Dict[str, Any]]:
    """
    获取 Level 1 完全匹配缓存

    返回：{
        "header_map": {...},
        "generated_code": "...",
        "structure_snapshot": {...}  # 用于验证结构是否仍匹配
    } 或 None
    """
    with _cache_lock:
        cache = _load_cache()

        file_hash = _compute_file_hash(excel_path)
        key = f"l1:{file_hash}:{sheet_name}:{'|'.join(sorted(subtable_titles))}"

        if key in cache:
            entry = cache[key]
            # 更新最后使用时间
            entry['last_used'] = datetime.now().isoformat()
            _save_cache(cache)
            return entry.get('data')

        return None


def set_level1_cache(
    excel_path: str,
    sheet_name: str,
    subtable_titles: List[str],
    header_map: Dict[str, Any],
    generated_code: str,
    sheet_structure: Dict[str, Any]
):
    """设置 Level 1 完全匹配缓存"""
    with _cache_lock:
        cache = _load_cache()

        file_hash = _compute_file_hash(excel_path)
        key = f"l1:{file_hash}:{sheet_name}:{'|'.join(sorted(subtable_titles))}"

        # 存储结构快照，用于后续验证
        structure_snapshot = {
            "potential_subtables": [
                {
                    "title": sub.get("title"),
                    "start_row": sub.get("start_row"),
                    "row_count": sub.get("row_count"),
                }
                for sub in sheet_structure.get("potential_subtables", [])
                if any(_normalize_string(t) in _normalize_string(sub.get("title", ""))
                       for t in subtable_titles)
            ],
        }

        cache[key] = {
            'data': {
                'header_map': header_map,
                'generated_code': generated_code,
                'structure_snapshot': structure_snapshot,
            },
            'created_at': datetime.now().isoformat(),
            'last_used': datetime.now().isoformat(),
            'type': 'level1',
        }

        cache = _enforce_cache_limit(cache)
        _save_cache(cache)


# ==================== Level 2: 结构指纹缓存 ====================

def get_level2_cache(
    sheet_name: str,
    current_structure: Dict[str, Any],
    subtable_titles: List[str]
) -> Optional[Tuple[str, Dict[str, Any], int]]:
    """
    获取 Level 2 结构指纹缓存

    返回：(generated_code, cached_header_map, row_offset) 或 None

    注意：只有当请求的 subtable_titles 与缓存时完全一致才命中，
    因为生成的代码中硬编码了子表名称，不同的 subtable_titles 需要不同的代码。
    """
    with _cache_lock:
        cache = _load_cache()

        signature = _compute_structure_signature(current_structure, subtable_titles)
        key = f"l2:{sheet_name}:{signature}"

        if key in cache:
            entry = cache[key]
            data = entry.get('data', {})

            # 获取缓存中的结构快照
            cached_structure = data.get('structure_snapshot', {})
            cached_subtables = cached_structure.get('potential_subtables', [])

            # 验证每个请求的子表标题是否都能在缓存中找到匹配
            cached_titles = [sub.get('title', '') for sub in cached_subtables]

            # 建立请求子表与缓存子表的一一映射关系
            used_cached_indices = set()

            for requested_title in subtable_titles:
                matched = False
                for i, cached_title in enumerate(cached_titles):
                    if i in used_cached_indices:
                        continue  # 已经匹配过的缓存标题不能再使用
                    if _match_subtable_title(requested_title, cached_title):
                        used_cached_indices.add(i)
                        matched = True
                        break
                if not matched:
                    # 有子表标题不匹配，缓存不命中
                    return None

            # 不再检查数量完全一致，只要请求的子表都能在缓存中找到即可

            # 计算行偏移量
            offset = _calculate_row_offset(
                cached_structure,
                current_structure,
                subtable_titles
            )

            entry['last_used'] = datetime.now().isoformat()
            _save_cache(cache)

            return (data.get('generated_code'), data.get('header_map'), offset)

        return None


def set_level2_cache(
    sheet_name: str,
    sheet_structure: Dict[str, Any],
    subtable_titles: List[str],
    header_map: Dict[str, Any],
    generated_code: str
):
    """设置 Level 2 结构指纹缓存"""
    with _cache_lock:
        cache = _load_cache()

        signature = _compute_structure_signature(sheet_structure, subtable_titles)
        key = f"l2:{sheet_name}:{signature}"

        # 存储结构快照
        structure_snapshot = {
            "potential_subtables": [
                {
                    "title": sub.get("title"),
                    "start_row": sub.get("start_row"),
                    "row_count": sub.get("row_count"),
                }
                for sub in sheet_structure.get("potential_subtables", [])
                if any(_normalize_string(t) in _normalize_string(sub.get("title", ""))
                       for t in subtable_titles)
            ],
            "sheet_name": sheet_name,
        }

        cache[key] = {
            'data': {
                'generated_code': generated_code,
                'header_map': header_map,
                'structure_snapshot': structure_snapshot,
            },
            'created_at': datetime.now().isoformat(),
            'last_used': datetime.now().isoformat(),
            'type': 'level2',
        }

        cache = _enforce_cache_limit(cache)
        _save_cache(cache)


# ==================== Level 3: 子表模板缓存 ====================

def get_level3_cache(
    sheet_name: str,
    subtable_title: str
) -> Optional[Dict[str, Any]]:
    """
    获取 Level 3 子表模板缓存

    返回：{
        "header_pattern": "single" | "double",
        "column_patterns": [...],
        "template_code": "..."
    } 或 None
    """
    with _cache_lock:
        cache = _load_cache()

        # 归一化子表标题（移除数字前缀）
        norm_title = _normalize_string(subtable_title)
        key = f"l3:{sheet_name}:{norm_title}"

        if key in cache:
            entry = cache[key]
            entry['last_used'] = datetime.now().isoformat()
            _save_cache(cache)
            return entry.get('data')

        return None


def set_level3_cache(
    sheet_name: str,
    subtable_title: str,
    template_data: Dict[str, Any]
):
    """设置 Level 3 子表模板缓存"""
    with _cache_lock:
        cache = _load_cache()

        norm_title = _normalize_string(subtable_title)
        key = f"l3:{sheet_name}:{norm_title}"

        cache[key] = {
            'data': template_data,
            'created_at': datetime.now().isoformat(),
            'last_used': datetime.now().isoformat(),
            'type': 'level3',
        }

        cache = _enforce_cache_limit(cache)
        _save_cache(cache)


# ==================== 坐标偏移代码调整 ====================

def apply_code_offset(
    generated_code: str,
    row_offset: int,
    header_map: Dict[str, Any]
) -> Tuple[str, Dict[str, Any]]:
    """
    当结构匹配但起始行有偏移时，调整代码中的行号

    参数:
        generated_code: LLM 生成的原始代码
        row_offset: 行偏移量（正数=向下偏移，负数=向上偏移）
        header_map: 原始 header_map

    返回:
        (adjusted_code, adjusted_header_map)
    """
    if row_offset == 0:
        return generated_code, header_map

    adjusted_code = generated_code

    # 替换策略：一次性替换所有行号，避免重复替换
    # 使用单个正则表达式匹配所有数字，然后根据上下文判断是否是行号

    def replace_all_numbers(match):
        """替换代码中的所有行号数字"""
        full_match = match.group(0)
        number = int(match.group(1))

        # 跳过太小的数字（可能是列号或常数）
        if number < 3:  # 列号从 1 开始，行号通常从 3 开始
            return full_match

        # 替换行号
        new_number = number + row_offset
        return full_match.replace(str(number), str(new_number))

    # 匹配函数调用中的数字参数：cell_val(6, c) 或 cell_val(6,c)
    # 只替换第一个参数（行号）
    def replace_cell_val(match):
        prefix = match.group(1)
        row = int(match.group(2))
        rest = match.group(3)
        return f"{prefix}{row + row_offset}{rest}"

    adjusted_code = re.sub(
        r'(cell_val\()(\d+)(,\s*[a-zA-Z_]\w*\))',
        replace_cell_val,
        adjusted_code
    )

    # 也匹配 cell_val(6, 2) 两个都是数字的情况
    adjusted_code = re.sub(
        r'(cell_val\()(\d+)(,\s*\d+\))',
        replace_cell_val,
        adjusted_code
    )

    # 匹配 ws.cell(row=6, column=2)
    def replace_ws_cell(match):
        prefix = match.group(1)
        row = int(match.group(2))
        rest = match.group(3)
        return f"{prefix}{row + row_offset}{rest}"

    adjusted_code = re.sub(
        r'(ws\.cell\(row=)(\d+)(,\s*column=\d+\))',
        replace_ws_cell,
        adjusted_code
    )

    # 匹配 range(X, Y) 但不匹配 for r in range（已经在上面处理了）
    # 只匹配直接调用：range(6, 14)
    def replace_range(match):
        start = int(match.group(1))
        end = int(match.group(2))
        # 跳过太小的数字（可能是列号）
        if start < 3:
            return match.group(0)
        return f"range({start + row_offset}, {end + row_offset})"

    adjusted_code = re.sub(
        r'range\((\d+),\s*(\d+)\)',
        replace_range,
        adjusted_code
    )

    # 调整 header_map 中的行号
    adjusted_header_map = header_map.copy()
    if 'subtable_start_row' in adjusted_header_map:
        adjusted_header_map['subtable_start_row'] += row_offset
    if 'subtable_end_row' in adjusted_header_map:
        adjusted_header_map['subtable_end_row'] += row_offset

    return adjusted_code, adjusted_header_map


# ==================== 缓存管理接口 ====================

def clear_cache(level: Optional[str] = None):
    """
    清空缓存

    参数:
        level: "level1" | "level2" | "level3" | None (全部)
    """
    with _cache_lock:
        if level is None:
            if os.path.exists(CACHE_FILE):
                os.remove(CACHE_FILE)
        else:
            cache = _load_cache()
            cache = {k: v for k, v in cache.items()
                     if v.get('type') != level}
            _save_cache(cache)


def get_cache_stats() -> Dict[str, Any]:
    """获取缓存统计信息"""
    with _cache_lock:
        cache = _load_cache()

        stats = {
            "total_entries": len(cache),
            "level1_count": 0,
            "level2_count": 0,
            "level3_count": 0,
        }

        for entry in cache.values():
            cache_type = entry.get('type', 'unknown')
            if cache_type == 'level1':
                stats["level1_count"] += 1
            elif cache_type == 'level2':
                stats["level2_count"] += 1
            elif cache_type == 'level3':
                stats["level3_count"] += 1

        # 计算文件大小
        if os.path.exists(CACHE_FILE):
            stats["cache_file_size"] = os.path.getsize(CACHE_FILE)

        return stats


def list_cache_entries(level: Optional[str] = None) -> List[Dict[str, Any]]:
    """列出所有缓存条目（用于调试）"""
    with _cache_lock:
        cache = _load_cache()
        result = []

        for key, entry in cache.items():
            if level is None or entry.get('type') == level:
                result.append({
                    'key': key,
                    'type': entry.get('type'),
                    'created_at': entry.get('created_at'),
                    'last_used': entry.get('last_used'),
                })

        return result

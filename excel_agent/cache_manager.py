"""
excel_agent/cache_manager.py — 三级缓存系统管理器

三级缓存架构：
┌─────────────────────────────────────────────────────────────┐
│                    Level 1: 完全匹配缓存                      │
│   Key: (excel_file_hash, sheet_name, subtable_titles)       │
│   Value: { generated_code, structure_fingerprint }          │
│   命中后：跳过 LLM，直接执行缓存的代码                         │
│   适用：同一个 Excel 文件                                     │
└─────────────────────────────────────────────────────────────┘
                          ↓ 未命中
┌─────────────────────────────────────────────────────────────┐
│                   Level 2: 结构指纹缓存                      │
│   Key: (sheet_name, structure_signature)                    │
│   Value: { generated_code, structure_fingerprint, ... }     │
│   命中后：计算坐标偏移量 → 调整代码 → 执行                     │
│   适用：不同 Excel 文件，但子表结构相同                         │
└─────────────────────────────────────────────────────────────┘
                          ↓ 未命中
┌─────────────────────────────────────────────────────────────┐
│                   Level 3: LLM 生成 (兜底)                   │
│   执行完整流程：parse → code_gen → sandbox → restore        │
│   成功后自动缓存到 Level 1 和 Level 2                         │
└─────────────────────────────────────────────────────────────┘
"""

import json
import hashlib
import os
import threading
from pathlib import Path
from typing import Dict, List, Any, Optional, Tuple
from datetime import datetime, timedelta
from collections import OrderedDict
import re


# ==================== 缓存文件路径 ====================
DEFAULT_CACHE_FILE = Path(__file__).parent.parent / "extraction_cache.json"
_cache_lock = threading.Lock()
_memory_cache: Optional[OrderedDict] = None  # 内存缓存（LRU）


# ==================== 配置常量 ====================
MAX_CACHE_ENTRIES = 500  # 最大缓存条目数
CACHE_TTL_DAYS = 30      # 缓存有效期（天）


# ==================== 工具函数 ====================

def _load_cache(file_path: Path) -> OrderedDict:
    """从文件加载缓存（LRU  OrderedDict）"""
    global _memory_cache

    if _memory_cache is not None:
        return _memory_cache

    if file_path.exists():
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
                # 转为 OrderedDict 并按 last_used 排序
                sorted_data = OrderedDict(
                    sorted(data.items(), key=lambda x: x[1].get("last_used", ""))
                )
                _memory_cache = sorted_data
                return _memory_cache
        except (json.JSONDecodeError, Exception):
            pass

    _memory_cache = OrderedDict()
    return _memory_cache


def _save_cache(file_path: Path, data: OrderedDict) -> None:
    """保存缓存到文件"""
    with open(file_path, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def _compute_file_hash(file_path: str) -> str:
    """
    计算文件哈希（平衡速度与准确性）
    读取文件头 1MB + 末尾 10KB
    """
    hasher = hashlib.sha256()
    file_size = os.path.getsize(file_path)

    with open(file_path, 'rb') as f:
        # 读取文件开头 1MB
        head_size = min(1024 * 1024, file_size)
        hasher.update(f.read(head_size))

        # 如果文件较大，再读取末尾 10KB
        if file_size > 1024 * 1024 + 1024 * 10:
            f.seek(-1024 * 10, 2)
            hasher.update(f.read(1024 * 10))

    return hasher.hexdigest()[:16]


def _normalize_titles(titles: List[str]) -> str:
    """归一化子表标题列表（用于缓存 Key）"""
    return "|".join(sorted(t.lower().strip() for t in titles))


def _is_expired(entry: Dict) -> bool:
    """检查缓存条目是否过期"""
    created_at = entry.get("created_at")
    if not created_at:
        return False

    try:
        created_time = datetime.fromisoformat(created_at)
        return datetime.now() - created_time > timedelta(days=CACHE_TTL_DAYS)
    except (ValueError, TypeError):
        return False


def _evict_if_needed(cache: OrderedDict, file_path: Path) -> None:
    """如果缓存超出限制，淘汰最久未使用的条目"""
    while len(cache) > MAX_CACHE_ENTRIES:
        # 删除最旧的条目（OrderedDict 保持插入顺序）
        cache.popitem(last=False)
        _save_cache(file_path, cache)


def _update_last_used(cache: OrderedDict, key: str, file_path: Path) -> None:
    """更新条目的最后使用时间（移动到末尾）"""
    if key in cache:
        cache.move_to_end(key)
        cache[key]["last_used"] = datetime.now().isoformat()
        _save_cache(file_path, cache)


# ==================== Level 1 缓存 ====================

def _build_l1_key(excel_path: str, sheet_name: str, subtable_titles: List[str]) -> str:
    """构建 Level 1 缓存 Key"""
    file_hash = _compute_file_hash(excel_path)
    titles_key = _normalize_titles(subtable_titles)
    return f"l1:{file_hash}:{sheet_name}:{titles_key}"


def l1_get(
    excel_path: str,
    sheet_name: str,
    subtable_titles: List[str],
    cache_file: Optional[Path] = None
) -> Optional[Dict]:
    """
    查询 Level 1 缓存（完全匹配）

    返回：
        - Dict: 缓存数据（包含 generated_code, structure_fingerprint 等）
        - None: 未命中
    """
    cache_file = cache_file or DEFAULT_CACHE_FILE
    key = _build_l1_key(excel_path, sheet_name, subtable_titles)

    with _cache_lock:
        cache = _load_cache(cache_file)
        entry = cache.get(key)

        if entry is None:
            return None

        if _is_expired(entry):
            del cache[key]
            _save_cache(cache_file, cache)
            return None

        _update_last_used(cache, key, cache_file)
        return entry.get("data")


def l1_set(
    excel_path: str,
    sheet_name: str,
    subtable_titles: List[str],
    data: Dict,
    cache_file: Optional[Path] = None
) -> None:
    """
    设置 Level 1 缓存

    参数：
        data: {
            "generated_code": str,
            "structure_fingerprint": Dict,
            "header_map": Dict,
            ...
        }
    """
    cache_file = cache_file or DEFAULT_CACHE_FILE
    key = _build_l1_key(excel_path, sheet_name, subtable_titles)

    with _cache_lock:
        cache = _load_cache(cache_file)

        cache[key] = {
            "data": data,
            "type": "level1",
            "created_at": datetime.now().isoformat(),
            "last_used": datetime.now().isoformat(),
        }

        _evict_if_needed(cache, cache_file)
        _save_cache(cache_file, cache)


# ==================== Level 2 缓存 ====================

def _build_l2_key(sheet_name: str, structure_signature: str) -> str:
    """构建 Level 2 缓存 Key"""
    return f"l2:{sheet_name}:{structure_signature}"


def compute_structure_signature(fingerprint: Dict[str, Any]) -> str:
    """
    从结构指纹计算哈希签名（用于 L2 缓存 Key）

    签名应满足：
    - 相同的表格结构 → 相同的签名
    - 不同的表格结构 → 不同的签名
    - 与绝对位置无关（只关心相对结构）
    """
    if not fingerprint:
        return ""

    # 提取与结构相关的核心特征（排除绝对位置）
    signature_data = {
        "subtables": []
    }

    for sub in fingerprint.get("subtables", []):
        sig_sub = {
            "title_pattern": sub.get("title_pattern", ""),
            "header_rows": sub.get("header_rows", 1),
            "row_count": sub.get("row_count", 0),
            "col_count": sub.get("col_count", 0),
            "header_structure": sub.get("header_structure", {}),
        }
        signature_data["subtables"].append(sig_sub)

    # 添加合并模式
    signature_data["merge_patterns"] = fingerprint.get("merge_patterns", [])

    # 计算哈希
    serialized = json.dumps(signature_data, sort_keys=True, ensure_ascii=False)
    hash_value = hashlib.sha256(serialized.encode()).hexdigest()[:16]

    return hash_value


def l2_get(
    sheet_name: str,
    structure_signature: str,
    cache_file: Optional[Path] = None
) -> Optional[Dict]:
    """
    查询 Level 2 缓存（结构指纹匹配）

    返回：
        - Dict: 缓存数据
        - None: 未命中
    """
    cache_file = cache_file or DEFAULT_CACHE_FILE
    key = _build_l2_key(sheet_name, structure_signature)

    with _cache_lock:
        cache = _load_cache(cache_file)
        entry = cache.get(key)

        if entry is None:
            return None

        if _is_expired(entry):
            del cache[key]
            _save_cache(cache_file, cache)
            return None

        _update_last_used(cache, key, cache_file)
        return entry.get("data")


def l2_set(
    sheet_name: str,
    structure_signature: str,
    data: Dict,
    cache_file: Optional[Path] = None
) -> None:
    """
    设置 Level 2 缓存

    参数：
        data: {
            "generated_code": str,
            "structure_fingerprint": Dict,
            "header_map": Dict,
            "structure_snapshot": Dict,  # 结构快照（用于偏移计算）
        }
    """
    cache_file = cache_file or DEFAULT_CACHE_FILE
    key = _build_l2_key(sheet_name, structure_signature)

    with _cache_lock:
        cache = _load_cache(cache_file)

        cache[key] = {
            "data": data,
            "type": "level2",
            "created_at": datetime.now().isoformat(),
            "last_used": datetime.now().isoformat(),
        }

        _evict_if_needed(cache, cache_file)
        _save_cache(cache_file, cache)


def l2_search_by_structure(
    sheet_name: str,
    current_fingerprint: Dict[str, Any],
    cache_file: Optional[Path] = None
) -> Optional[Tuple[Dict, int]]:
    """
    根据当前结构指纹搜索 Level 2 缓存

    返回：
        - Tuple[Dict, int]: (缓存数据，行偏移量)
        - None: 未命中
    """
    cache_file = cache_file or DEFAULT_CACHE_FILE

    with _cache_lock:
        cache = _load_cache(cache_file)

        # 遍历所有 L2 缓存条目
        for key, entry in cache.items():
            if not key.startswith("l2:") or not sheet_name in key:
                continue

            if _is_expired(entry):
                del cache[key]
                continue

            cached_data = entry.get("data", {})
            cached_fp = cached_data.get("structure_fingerprint", {})

            # 比较结构是否匹配
            if _structure_match(current_fingerprint, cached_fp):
                # 计算偏移量
                offset = _compute_row_offset(current_fingerprint, cached_fp)
                _update_last_used(cache, key, cache_file)
                _save_cache(cache_file, cache)
                return (cached_data, offset)

    return None


def _structure_match(current_fp: Dict, cached_fp: Dict) -> bool:
    """
    判断两个结构指纹是否匹配（同源结构）

    匹配条件：
    - 子表数量相同
    - 每个子表的 title_pattern 相同
    - header_rows、row_count、col_count 相同
    - header_structure 相同
    """
    if not current_fp or not cached_fp:
        return False

    curr_subs = current_fp.get("subtables", [])
    cached_subs = cached_fp.get("subtables", [])

    if len(curr_subs) != len(cached_subs):
        return False

    for curr, cached in zip(curr_subs, cached_subs):
        # 比较标题模式
        if curr.get("title_pattern") != cached.get("title_pattern"):
            return False

        # 比较表头行数
        if curr.get("header_rows") != cached.get("header_rows"):
            return False

        # 比较行列数
        if curr.get("row_count") != cached.get("row_count"):
            return False
        if curr.get("col_count") != cached.get("col_count"):
            return False

        # 比较表头结构
        if curr.get("header_structure") != cached.get("header_structure"):
            return False

    # 比较合并模式
    curr_merges = current_fp.get("merge_patterns", [])
    cached_merges = cached_fp.get("merge_patterns", [])
    if curr_merges != cached_merges:
        return False

    return True


def _compute_row_offset(current_fp: Dict, cached_fp: Dict) -> int:
    """
    计算当前结构与缓存结构的行偏移量

    返回：
        int: 行偏移量（正数=向下偏移，负数=向上偏移）
    """
    curr_subs = current_fp.get("subtables", [])
    cached_subs = cached_fp.get("subtables", [])

    if not curr_subs or not cached_subs:
        return 0

    # 使用第一个子表的起始行计算偏移
    curr_start = curr_subs[0].get("start_row", 0)
    cached_start = cached_subs[0].get("start_row", 0)

    return curr_start - cached_start


def l2_search_by_light_signature(
    sheet_name: str,
    light_fingerprint: Dict[str, Any],
    cache_file: Optional[Path] = None
) -> Optional[Tuple[Dict, int]]:
    """
    根据轻量级结构指纹搜索 Level 2 缓存

    返回：
        - Tuple[Dict, int]: (缓存数据，行偏移量)
        - None: 未命中
    """
    if not light_fingerprint:
        return None

    cache_file = cache_file or DEFAULT_CACHE_FILE

    # 使用轻量级指纹的 signature 直接查询
    signature = light_fingerprint.get("signature")
    if not signature:
        return None

    key = _build_l2_key(sheet_name, signature)

    with _cache_lock:
        cache = _load_cache(cache_file)
        entry = cache.get(key)

        if entry is None:
            # 尝试模糊匹配：遍历所有同 sheet_name 的 L2 缓存
            for k, e in cache.items():
                if not k.startswith("l2:") or sheet_name not in k:
                    continue
                if _is_expired(e):
                    del cache[k]
                    continue

                cached_data = e.get("data", {})
                cached_light_fp = cached_data.get("light_structure_fingerprint", {})

                # 使用轻量级匹配逻辑
                if _light_structure_match(light_fingerprint, cached_light_fp):
                    offset = _compute_light_row_offset(light_fingerprint, cached_light_fp)
                    _update_last_used(cache, k, cache_file)
                    _save_cache(cache_file, cache)
                    return (cached_data, offset)

            return None

        if _is_expired(entry):
            del cache[key]
            _save_cache(cache_file, cache)
            return None

        _update_last_used(cache, key, cache_file)
        cached_data = entry.get("data", {})

        # 计算偏移量
        cached_light_fp = cached_data.get("light_structure_fingerprint", {})
        offset = _compute_light_row_offset(light_fingerprint, cached_light_fp)

        return (cached_data, offset)


def _light_structure_match(light_fp: Dict, cached_light_fp: Dict) -> bool:
    """
    判断两个轻量级结构指纹是否匹配

    匹配条件（宽松匹配）：
    - 子表数量相同
    - 每个子表的 title_pattern 相同
    - header_rows 相同
    - col_count 相同
    - header_signature 相同（关键）
    """
    if not light_fp or not cached_light_fp:
        return False

    curr_subs = light_fp.get("subtables", [])
    cached_subs = cached_light_fp.get("subtables", [])

    if len(curr_subs) != len(cached_subs):
        return False

    for curr, cached in zip(curr_subs, cached_subs):
        # 比较标题模式
        if curr.get("title_pattern") != cached.get("title_pattern"):
            return False

        # 比较表头行数
        if curr.get("header_rows") != cached.get("header_rows"):
            return False

        # 比较列数
        if curr.get("col_count") != cached.get("col_count"):
            return False

        # 比较表头签名（关键：确保表头内容相似）
        if curr.get("header_signature") != cached.get("header_signature"):
            return False

    return True


def _compute_light_row_offset(light_fp: Dict, cached_light_fp: Dict) -> int:
    """
    计算当前轻量级结构与缓存结构的行偏移量

    返回：
        int: 行偏移量（正数=向下偏移，负数=向上偏移）
    """
    curr_subs = light_fp.get("subtables", [])
    cached_subs = cached_light_fp.get("subtables", [])

    if not curr_subs or not cached_subs:
        return 0

    # 使用第一个子表的起始行计算偏移
    curr_start = curr_subs[0].get("start_row", 0)
    cached_start = cached_subs[0].get("start_row", 0)

    return curr_start - cached_start


# ==================== Level 3 缓存 ====================

# Level 3 是 LLM 生成兜底，不需要显式缓存
# 当 LLM 生成成功后，自动写入 L1 和 L2


# ==================== 代码偏移调整 ====================

def apply_code_offset(
    generated_code: str,
    row_offset: int,
    col_offset: int,
    header_map: Dict
) -> Tuple[str, Dict]:
    """
    根据偏移量调整代码中的行列号

    参数：
        generated_code: 原始生成的代码
        row_offset: 行偏移量（正数=向下）
        col_offset: 列偏移量（正数=向右）
        header_map: 表头映射（包含子表位置信息）

    返回：
        (adjusted_code, adjusted_header_map)
    """
    adjusted_code = generated_code

    # 1. 替换元组赋值语句中的 start_row, start_col = X, Y
    def replace_tuple_assign(m):
        indent = m.group(1)
        row_val = int(m.group(6))
        col_val = int(m.group(8))
        return f"{indent}start_row, start_col = {row_val + row_offset}, {col_val + col_offset}"

    adjusted_code = re.sub(
        r'^(\s*)(start_row)(\s*,\s*)(start_col)(\s*=\s*)(\d+)(\s*,\s*)(\d+)',
        replace_tuple_assign,
        adjusted_code,
        flags=re.MULTILINE
    )

    # 2. 替换变量赋值语句中的 start_row = X
    def replace_start_row(m):
        indent = m.group(1)
        value = int(m.group(4))
        return f"{indent}start_row = {value + row_offset}"

    adjusted_code = re.sub(
        r'^(\s*)(start_row)(\s*=\s*)(\d+)',
        replace_start_row,
        adjusted_code,
        flags=re.MULTILINE
    )

    # 3. 替换变量赋值语句中的 start_col = X
    def replace_start_col(m):
        indent = m.group(1)
        value = int(m.group(4))
        return f"{indent}start_col = {value + col_offset}"

    adjusted_code = re.sub(
        r'^(\s*)(start_col)(\s*=\s*)(\d+)',
        replace_start_col,
        adjusted_code,
        flags=re.MULTILINE
    )

    # 4. 替换变量赋值语句中的 end_row = X
    def replace_end_row(m):
        indent = m.group(1)
        value = int(m.group(4))
        return f"{indent}end_row = {value + row_offset}"

    adjusted_code = re.sub(
        r'^(\s*)(end_row)(\s*=\s*)(\d+)',
        replace_end_row,
        adjusted_code,
        flags=re.MULTILINE
    )

    # 5. 替换 range(X, Y) 中的行号
    def replace_range(m):
        start = int(m.group(1)) + row_offset
        end = int(m.group(2)) + row_offset
        return f"range({start}, {end})"

    adjusted_code = re.sub(r'range\((\d+),\s*(\d+)\)', replace_range, adjusted_code)

    # 6. 替换 cell_val(r, c) 中的行号
    def replace_cell_val(m):
        row = int(m.group(1)) + row_offset
        col = int(m.group(2)) + col_offset
        return f"cell_val({row}, {col})"

    adjusted_code = re.sub(r'cell_val\((\d+),\s*(\d+)\)', replace_cell_val, adjusted_code)

    # 调整 header_map 中的行号
    adjusted_header_map = _adjust_header_map(header_map, row_offset, col_offset)

    return adjusted_code, adjusted_header_map


def _adjust_header_map(
    header_map: Dict,
    row_offset: int,
    col_offset: int
) -> Dict:
    """调整 header_map 中的坐标"""
    if not header_map:
        return {}
    adjusted = header_map.copy()

    # 调整子表边界
    for key in ["subtable_start_row", "subtable_end_row"]:
        if key in adjusted:
            adjusted[key] = adjusted[key] + row_offset

    # 调整表头行号
    for key in ["header_row_1", "header_row_2"]:
        if key in adjusted:
            adjusted[key] = adjusted[key] + row_offset

    # 调整列信息中的行号
    if "columns" in adjusted:
        adjusted_columns = []
        for col in adjusted.get("columns", []):
            new_col = col.copy()
            if "row" in new_col:
                new_col["row"] = new_col["row"] + row_offset
            adjusted_columns.append(new_col)
        adjusted["columns"] = adjusted_columns

    return adjusted


# ==================== 缓存管理接口 ====================

def clear_cache(
    level: Optional[str] = None,
    cache_file: Optional[Path] = None
) -> Dict[str, int]:
    """
    清空缓存

    参数：
        level: "level1" | "level2" | "level3" | None（全部）
        cache_file: 缓存文件路径

    返回：
        {"cleared_count": int}
    """
    cache_file = cache_file or DEFAULT_CACHE_FILE
    cleared = 0

    with _cache_lock:
        cache = _load_cache(cache_file)

        if level is None:
            cleared = len(cache)
            cache.clear()
        else:
            keys_to_delete = [
                k for k in cache
                if cache[k].get("type") == level
            ]
            for key in keys_to_delete:
                del cache[key]
                cleared += 1

        _save_cache(cache_file, cache)
        _memory_cache = cache  # 更新内存缓存

    return {"cleared_count": cleared}


def get_cache_stats(cache_file: Optional[Path] = None) -> Dict[str, Any]:
    """获取缓存统计信息"""
    cache_file = cache_file or DEFAULT_CACHE_FILE

    with _cache_lock:
        cache = _load_cache(cache_file)

        stats = {
            "total_entries": len(cache),
            "level1_count": 0,
            "level2_count": 0,
            "level3_count": 0,
        }

        for entry in cache.values():
            t = entry.get("type", "unknown")
            if t == "level1":
                stats["level1_count"] += 1
            elif t == "level2":
                stats["level2_count"] += 1
            elif t == "level3":
                stats["level3_count"] += 1

        return stats


def list_cache_entries(
    level: Optional[str] = None,
    cache_file: Optional[Path] = None
) -> List[Dict]:
    """列出缓存条目"""
    cache_file = cache_file or DEFAULT_CACHE_FILE

    with _cache_lock:
        cache = _load_cache(cache_file)
        entries = []

        for key, entry in cache.items():
            if level is None or entry.get("type") == level:
                entries.append({
                    "key": key,
                    "type": entry.get("type"),
                    "created_at": entry.get("created_at"),
                    "last_used": entry.get("last_used"),
                })

        return entries


def delete_cache_entry(
    key: str,
    cache_file: Optional[Path] = None
) -> bool:
    """删除指定缓存条目"""
    cache_file = cache_file or DEFAULT_CACHE_FILE

    with _cache_lock:
        cache = _load_cache(cache_file)

        if key in cache:
            del cache[key]
            _save_cache(cache_file, cache)
            _memory_cache = cache
            return True

        return False


# ==================== 缓存查询统一入口 ====================

def query_cache(
    excel_path: str,
    sheet_name: str,
    subtable_titles: List[str],
    structure_fingerprint: Optional[Dict] = None,
    cache_file: Optional[Path] = None
) -> Tuple[Optional[Dict], str]:
    """
    统一缓存查询入口（三级缓存逐级查找）

    返回：
        (cached_data, cache_level)
        - cached_data: 缓存数据（None 表示未命中）
        - cache_level: "l1" | "l2" | "l3" (未命中)
    """
    # 1. 尝试 Level 1（完全匹配）
    l1_data = l1_get(excel_path, sheet_name, subtable_titles, cache_file)
    if l1_data:
        return (l1_data, "l1")

    # 2. 尝试 Level 2（结构指纹匹配）
    if structure_fingerprint:
        l2_result = l2_search_by_structure(sheet_name, structure_fingerprint, cache_file)
        if l2_result:
            l2_data, offset = l2_result
            return (l2_data, "l2")

    # 3. Level 3（未命中，需要 LLM 生成）
    return (None, "l3")


def save_to_cache(
    excel_path: str,
    sheet_name: str,
    subtable_titles: List[str],
    data: Dict,
    structure_fingerprint: Optional[Dict] = None,
    light_structure_fingerprint: Optional[Dict] = None,
    cache_file: Optional[Path] = None
) -> None:
    """
    保存数据到缓存（同时写入 L1 和 L2）

    参数：
        data: 要缓存的数据（包含 generated_code, header_map 等）
        structure_fingerprint: 结构指纹（用于 L2，来自 structure_analyzer_node）
        light_structure_fingerprint: 轻量级结构指纹（用于 L2，来自 light_structure_analyzer_node）

    注意：
        - 优先使用 light_structure_fingerprint 保存 L2（用于快速查询）
        - 如果只有 structure_fingerprint，也保存一份（向后兼容）
        - 避免重复保存：如果两种指纹的 signature 相同，只保存一次
    """
    # 保存到 Level 1
    l1_set(excel_path, sheet_name, subtable_titles, data, cache_file)

    # 保存到 Level 2（优先使用轻量级结构指纹）
    if light_structure_fingerprint:
        light_sig = light_structure_fingerprint.get("signature")
        if light_sig:
            l2_set(sheet_name, light_sig, {
                **data,
                "light_structure_fingerprint": light_structure_fingerprint,
                "structure_fingerprint": structure_fingerprint,
            }, cache_file)
    elif structure_fingerprint:
        # 如果没有轻量级指纹，使用完整指纹保存（向后兼容）
        signature = compute_structure_signature(structure_fingerprint)
        if signature:
            l2_set(sheet_name, signature, {
                **data,
                "structure_fingerprint": structure_fingerprint,
            }, cache_file)

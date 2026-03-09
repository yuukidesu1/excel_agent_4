"""
excel_agent/memory.py — 代码记忆管理模块

功能：
  1. 基于场景特征生成唯一 Key
  2. 存储/检索历史成功代码
  3. 支持相似度匹配（模糊复用）
  4. 自动淘汰旧记忆（LRU 策略）
  5. 线程安全的读写操作
  6. 子表级别的记忆复用（核心优化）
"""
import json
import os
import hashlib
import threading
import re
from typing import List, Dict, Any, Optional, Tuple
from datetime import datetime

MEMORY_FILE = "code_library.json"
_memory_lock = threading.Lock()

# 记忆配置
MAX_MEMORIES = 100  # 最大记忆数量，超过时自动淘汰最旧的
MIN_ACCURACY = 0.5  # 相似度匹配阈值（降低以允许部分匹配）
SUBTABLE_MIN_ACCURACY = 0.6  # 子表级别匹配的最低相似度


def _generate_key(sheet_name: str, titles: List[str]) -> str:
    """基于 Sheet 名和需要抽取的子表名生成唯一 Scenario Key"""
    combined = f"{sheet_name}_" + "_".join(sorted(titles))
    return hashlib.md5(combined.encode('utf-8')).hexdigest()


def _normalize_string(s: str) -> str:
    """标准化字符串用于相似度比较：转小写、去除空白和特殊字符"""
    return re.sub(r'[\s\W_]+', '', str(s).lower())


def _compute_similarity(titles1: List[str], titles2: List[str]) -> float:
    """计算两个标题列表的相似度（Jaccard 相似系数）"""
    if not titles1 or not titles2:
        return 0.0

    set1 = {_normalize_string(t) for t in titles1}
    set2 = {_normalize_string(t) for t in titles2}

    if not set1 or not set2:
        return 0.0

    intersection = len(set1 & set2)
    union = len(set1 | set2)
    return intersection / union if union > 0 else 0.0


def _load_memory() -> Dict[str, Any]:
    """安全加载记忆文件，返回结构化数据"""
    if not os.path.exists(MEMORY_FILE):
        return {}

    try:
        with open(MEMORY_FILE, 'r', encoding='utf-8') as f:
            data = json.load(f)
        # 兼容旧格式（纯代码字符串）
        if isinstance(data, dict):
            return data
    except (json.JSONDecodeError, IOError):
        pass
    return {}


def _save_memory(mem: Dict[str, Any]):
    """安全保存记忆到文件"""
    with open(MEMORY_FILE, 'w', encoding='utf-8') as f:
        json.dump(mem, f, ensure_ascii=False, indent=2)


def _enforce_memory_limit(mem: Dict[str, Any], max_count: int = MAX_MEMORIES):
    """当记忆数量超过上限时，淘汰最旧的记忆（基于使用时间）"""
    if len(mem) <= max_count:
        return mem

    # 提取所有记忆的使用时间
    memories_with_time = []
    for key, value in mem.items():
        if isinstance(value, dict):
            last_used = value.get('last_used', '')
        else:
            last_used = ''  # 旧格式没有使用时间
        memories_with_time.append((key, last_used))

    # 按使用时间排序，淘汰最旧的
    memories_with_time.sort(key=lambda x: x[1])
    keys_to_remove = [m[0] for m in memories_with_time[:len(mem) - max_count]]

    for key in keys_to_remove:
        del mem[key]

    return mem


def get_reference_code(
    sheet_name: str,
    titles: List[str],
    use_fuzzy_match: bool = True
) -> Tuple[str, float, List[str]]:
    """
    获取该场景历史成功的代码

    参数:
        sheet_name: Sheet 名称
        titles: 子表标题列表
        use_fuzzy_match: 是否启用模糊匹配（默认 True）

    返回:
        (code, similarity, matched_titles) 元组
        - code: 参考代码字符串，无匹配时返回 ""
        - similarity: 匹配置信度 (0.0-1.0)
        - matched_titles: 匹配的子表名称列表（用于告知 LLM 哪些子表是匹配的）
    """
    with _memory_lock:
        mem = _load_memory()

        # 1. 优先尝试精确匹配
        key = _generate_key(sheet_name, titles)
        if key in mem:
            entry = mem[key]
            if isinstance(entry, dict):
                # 新格式：{"code": "...", "metadata": {...}}
                entry['last_used'] = datetime.now().isoformat()
                _save_memory(mem)
                return entry.get('code', ''), 1.0, titles
            else:
                # 旧格式：直接是代码字符串
                return entry, 1.0, titles

        # 2. 模糊匹配：寻找相似场景
        if use_fuzzy_match:
            best_match = None
            best_similarity = 0.0
            best_match_titles = []
            best_matched_subtables = []

            for stored_key, entry in mem.items():
                if not isinstance(entry, dict):
                    continue

                meta = entry.get('metadata', {})
                stored_titles = meta.get('titles', [])
                stored_sheet = meta.get('sheet_name', '')

                # 只匹配相同 Sheet 的场景（不同 Sheet 结构可能完全不同）
                if stored_sheet != sheet_name:
                    continue

                similarity = _compute_similarity(titles, stored_titles)

                # 新增：子表包含关系的额外加分
                matched_subs = []
                if similarity > 0:
                    target_set = {_normalize_string(t) for t in titles}
                    stored_set = {_normalize_string(t) for t in stored_titles}

                    # 找出匹配的子表（优先精确匹配，其次数字前缀匹配）
                    used_stored = set()  # 记录已匹配的 stored_titles 索引
                    for idx, t in enumerate(titles):
                        norm_t = _normalize_string(t)
                        matched = False

                        # 先尝试精确匹配
                        for sidx, st in enumerate(stored_titles):
                            if sidx in used_stored:
                                continue
                            if _normalize_string(st) == norm_t:
                                matched_subs.append(st)
                                used_stored.add(sidx)
                                matched = True
                                break

                        # 如果没有精确匹配，尝试数字前缀匹配
                        if not matched:
                            for sidx, st in enumerate(stored_titles):
                                if sidx in used_stored:
                                    continue
                                st_no_num = re.sub(r'^\d+', '', _normalize_string(st))
                                t_no_num = re.sub(r'^\d+', '', norm_t)
                                if st_no_num == t_no_num:
                                    matched_subs.append(st)
                                    used_stored.add(sidx)
                                    break

                    # 如果目标子表都被存储的场景包含，提高相似度
                    if target_set.issubset(stored_set):
                        similarity = min(1.0, similarity + 0.2)
                    # 如果存储的场景是目标的子集，也提高相似度
                    elif stored_set.issubset(target_set):
                        similarity = min(1.0, similarity + 0.15)

                if similarity > best_similarity and similarity >= MIN_ACCURACY:
                    best_similarity = similarity
                    best_match = entry.get('code', '')
                    best_match_titles = stored_titles
                    best_matched_subtables = matched_subs

            if best_match:
                # 更新最后使用时间
                for stored_key, entry in mem.items():
                    if isinstance(entry, dict) and entry.get('code') == best_match:
                        if 'metadata' in entry:
                            entry['metadata']['last_used'] = datetime.now().isoformat()
                        else:
                            entry['last_used'] = datetime.now().isoformat()
                        break
                _save_memory(mem)
                return best_match, best_similarity, best_matched_subtables

        return "", 0.0, []


def save_reference_code(
    sheet_name: str,
    titles: List[str],
    code: str,
    quality_score: float = 1.0,
    metadata: Optional[Dict[str, Any]] = None
):
    """
    保存高质量代码到记忆库

    参数:
        sheet_name: Sheet 名称
        titles: 子表标题列表
        code: 成功抽取的代码
        quality_score: 质量评分 (0.0-1.0)
        metadata: 额外元数据
    """
    with _memory_lock:
        mem = _load_memory()

        key = _generate_key(sheet_name, titles)

        # 新格式：结构化存储
        entry = {
            'code': code,
            'metadata': {
                'sheet_name': sheet_name,
                'titles': titles,
                'created_at': datetime.now().isoformat(),
                'last_used': datetime.now().isoformat(),
                'quality_score': quality_score,
            }
        }

        # 合并额外元数据
        if metadata:
            entry['metadata'].update(metadata)

        mem[key] = entry

        # 执行记忆淘汰
        mem = _enforce_memory_limit(mem)

        _save_memory(mem)


def list_memories() -> List[Dict[str, Any]]:
    """列出所有记忆条目（用于调试和管理）"""
    with _memory_lock:
        mem = _load_memory()
        result = []
        for key, entry in mem.items():
            if isinstance(entry, dict):
                result.append({
                    'key': key,
                    'metadata': entry.get('metadata', {}),
                })
            else:
                result.append({
                    'key': key,
                    'metadata': {'format': 'legacy'},
                })
        return result


def clear_memory():
    """清空所有记忆（谨慎使用）"""
    with _memory_lock:
        if os.path.exists(MEMORY_FILE):
            os.remove(MEMORY_FILE)


def delete_memory_entry(sheet_name: str, titles: List[str]):
    """删除指定场景的记忆"""
    with _memory_lock:
        mem = _load_memory()
        key = _generate_key(sheet_name, titles)
        if key in mem:
            del mem[key]
            _save_memory(mem)


def get_reference_code_by_subtable(
    sheet_name: str,
    subtable_title: str,
    use_fuzzy_match: bool = True
) -> Tuple[Optional[str], float, Optional[str]]:
    """
    根据单个子表标题查找最匹配的参考代码

    参数:
        sheet_name: Sheet 名称
        subtable_title: 单个子表标题（如 "4G Configuration"）
        use_fuzzy_match: 是否启用模糊匹配

    返回:
        (code, similarity, source_table_name) 或 (None, 0.0, None)
        - code: 参考代码片段（只包含该子表的提取逻辑）
        - similarity: 匹配置信度
        - source_table_name: 匹配的源代码中的子表名
    """
    with _memory_lock:
        mem = _load_memory()

        best_match = None
        best_similarity = 0.0
        best_source_table = None

        norm_target = _normalize_string(subtable_title)

        for stored_key, entry in mem.items():
            if not isinstance(entry, dict):
                continue

            meta = entry.get('metadata', {})
            stored_sheet = meta.get('sheet_name', '')
            stored_titles = meta.get('titles', [])

            # 只匹配相同 Sheet 的场景
            if stored_sheet != sheet_name:
                continue

            # 检查是否有子表匹配
            for stored_title in stored_titles:
                norm_stored = _normalize_string(stored_title)

                # 精确匹配
                if norm_stored == norm_target:
                    return entry.get('code', ''), 1.0, stored_title

                # 模糊匹配：计算子表名称相似度
                if use_fuzzy_match:
                    # 使用更宽松的相似度计算（包含关系也算）
                    sim = _compute_subtable_similarity(norm_target, norm_stored)
                    if sim > best_similarity and sim >= SUBTABLE_MIN_ACCURACY:
                        best_similarity = sim
                        best_match = entry.get('code', '')
                        best_source_table = stored_title

        if best_match:
            return best_match, best_similarity, best_source_table

        return None, 0.0, None


def _compute_subtable_similarity(s1: str, s2: str) -> float:
    """
    计算两个子表名称的相似度
    使用更宽松的策略：如果一个包含另一个，给予较高相似度
    """
    if not s1 or not s2:
        return 0.0

    # 完全相同
    if s1 == s2:
        return 1.0

    # 包含关系（如 "4G Configuration" 包含 "Configuration"）
    if s1 in s2 or s2 in s1:
        return 0.8

    # 提取共同部分（如 "4G Configuration" vs "5G Configuration"）
    # 移除数字前缀后比较
    s1_no_num = re.sub(r'^\d+', '', s1)
    s2_no_num = re.sub(r'^\d+', '', s2)

    if s1_no_num == s2_no_num:
        return 0.85

    # 使用 Jaccard 相似度作为兜底
    set1 = set(s1)
    set2 = set(s2)
    intersection = len(set1 & set2)
    union = len(set1 | set2)
    return intersection / union if union > 0 else 0.0
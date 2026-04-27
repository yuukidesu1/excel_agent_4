"""
nodes/kv_scene_router.py — LLM 场景路由节点

职责：
    判断 table-configured 的每个子表，底层数据结构是 KV 还是 Table。
    将决策写入 cache.entries[title]["extract_mode"]，供下游路由使用。

注意：
    - extract_type="kv" 或 kv_list 非空时跳过（全局 KV 优先）
    - temperature=0，确保确定性
"""

import json
import re
import os
from pathlib import Path
from typing import Dict, List, Any

from langchain_openai import ChatOpenAI
from langchain_core.messages import SystemMessage, HumanMessage

from excel_agent.state import AgentState


_SYSTEM_PROMPT = """\
你是一个经验丰富的 Excel 数据结构分析专家。请根据传入的子表信息，判断每个子表的底层数据结构。

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
【判断标准】

★★★ KV 模式（extract_mode="kv"）的特征 ★★★
  1. 离散键值对布局：Key 和 Value 分散在多个行上，每行一个 KV 对
  2. Key 列特征：左侧某一列包含配置的"字段名"（如"Site Power Supply"、"Generator Type"）
  3. Value 列特征：右侧某列包含对应的"值"
  4. 行独立性：每一行只包含一个完整的 KV 对，行与行之间数据不连续
  5. 无数据行：没有多行"记录"，只有 KV 对的堆叠
  6. 典型场景：设备配置信息、电源信息、电池信息等"属性列表"

★★★ Table 模式（extract_mode="table"）的特征 ★★★
  1. 连续表格布局：表头在同一行，下方有连续的数据行
  2. 表头特征：所有列名在同一行（通常是子表标题下方第 1-2 行）
  3. 数据行特征：每一行是一条"独立记录"，包含多个字段
  4. 多行数据：通常有多行相似结构的数据（如多个设备、多个站点）
  5. 典型场景：设备列表、站点配置表、多行记录表

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
【判别技巧】

→ 如果配置的 headers 在 Excel 中是"左侧一列字段名 + 右侧对应值"的离散排列 → KV 模式
→ 如果配置的 headers 在 Excel 中是"同一行多个列名 + 下方多行数据" → Table 模式
→ 如果子表区域内大部分行只有 2-3 个非空单元格且分散排列 → KV 模式
→ 如果子表区域内有多行连续的、每行都有多个非空单元格 → Table 模式

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
【输出格式】
严格返回 JSON，不要包含其他说明文字：
{
  "decisions": [
    {
      "title": "子表名称",
      "extract_mode": "kv",  // 或 "table"
      "reason": "简短判断理由，说明是离散 KV 还是连续表格"
    }
  ]
}
"""


def _get_llm() -> ChatOpenAI:
    from dotenv import load_dotenv
    for parent in Path(__file__).resolve().parents:
        env_file = parent / ".env"
        if env_file.exists():
            load_dotenv(env_file, override=True)
            break
    api_key = os.getenv("OPENAI_API_KEY")
    base_url = os.getenv("OPENAI_BASE_URL") or None
    model = os.getenv("LLM_MODEL", "glm-4.7")
    if not api_key:
        raise ValueError("未找到 OPENAI_API_KEY，请检查 .env 文件。")
    return ChatOpenAI(model=model, temperature=0, api_key=api_key, base_url=base_url)


def _build_cell_view(
    title: str,
    start_row: int,
    start_col: int,
    title_row_span: int,
    max_row: int,
    non_empty_cells: List[Dict],
    merged_cells_info: List[Dict],
    search_range: int = 25,
) -> Dict[str, Any]:
    """为单个子表构建压缩单元格视图"""
    end_row = min(start_row + title_row_span + search_range, max_row)

    cells = [
        f"R{c['row']}C{c['col']}:{c['value']}"
        for c in non_empty_cells
        if start_row <= c["row"] <= end_row
    ][:200]

    merges = [
        f"R{m['min_row']}C{m['min_col']}~R{m['max_row']}C{m['max_col']}:{m['value']}"
        for m in merged_cells_info
        if start_row <= m["max_row"] <= end_row
    ][:30]

    return {
        "title": title,
        "title_row": start_row,
        "title_col": start_col,
        "cells": cells,
        "merged_cells": merges,
    }


def _parse_router_output(raw: str) -> List[Dict]:
    """从 LLM 输出中解析 decisions"""
    if not raw:
        return []
    try:
        # 先尝试提取 ```json ... ``` 块
        m = re.search(r"```json\s*([\s\S]*?)```", raw)
        if m:
            parsed = json.loads(m.group(1).strip())
        else:
            m = re.search(r"```\s*([\s\S]*?)```", raw)
            if m:
                parsed = json.loads(m.group(1).strip())
            else:
                parsed = json.loads(raw.strip())
        return parsed.get("decisions", [])
    except json.JSONDecodeError:
        return []


def kv_scene_router_node(state: AgentState) -> dict:
    """LLM 场景路由节点"""
    config = state.get("config", {})

    # 全局 KV 模式优先，跳过路由
    extract_type = config.get("extract_type", "table")
    kv_list = config.get("kv_list", [])
    if (extract_type and extract_type.lower() == "kv") or (kv_list and len(kv_list) > 0):
        return {}

    subtable_titles = config.get("subtable_titles", [])
    subtable_configs = config.get("subtable_configs") or {}
    cache_state = state.get("cache", {})
    entries = cache_state.get("entries", {})

    sheet_structure = state.get("sheet_structure", {})
    non_empty_cells = sheet_structure.get("non_empty_cells", [])
    merged_cells_info = sheet_structure.get("merged_cells_info", [])
    max_row = sheet_structure.get("max_row", 100)

    if not subtable_titles or not entries:
        return {}

    # 构建子表视图
    cell_views = []
    for title in subtable_titles:
        entry = entries.get(title)
        if not entry:
            continue
        start_row = entry.get("start_row", 1)
        start_col = entry.get("start_col", 1)
        # 计算 title_row_span：PSA 写入的 start_row 即标题行
        # 我们需要从 merged_cells_info 中找到标题行的 span
        title_row_span = 1
        for m in merged_cells_info:
            if m["min_row"] == start_row and m["min_col"] == start_col:
                title_row_span = m["row_span"]
                break

        view = _build_cell_view(
            title=title,
            start_row=start_row,
            start_col=start_col,
            title_row_span=title_row_span,
            max_row=max_row,
            non_empty_cells=non_empty_cells,
            merged_cells_info=merged_cells_info,
        )
        cell_views.append(view)

    if not cell_views:
        return {}

    # 提取 subtable_configs 中的 headers
    headers_map = {}
    if isinstance(subtable_configs, dict):
        for title, tc in subtable_configs.items():
            if isinstance(tc, dict):
                headers_map[title] = tc.get("headers", [])
            elif isinstance(tc, list):
                headers_map[title] = [str(item) for item in tc]
    elif isinstance(subtable_configs, list):
        # 全局 headers
        for title in subtable_titles:
            headers_map[title] = [str(item) for item in subtable_configs]

    # 构建 LLM 输入
    ctx = {
        "subtable_titles": subtable_titles,
        "subtable_configs": {
            title: {"headers": headers}
            for title, headers in headers_map.items()
        },
        "subtable_cell_views": cell_views,
    }

    messages = [
        SystemMessage(content=_SYSTEM_PROMPT),
        HumanMessage(content=json.dumps(ctx, ensure_ascii=False, indent=2)),
    ]

    response = _get_llm().invoke(messages)
    decisions = _parse_router_output(response.content) if response else []

    # 将决策写入 cache.entries
    for decision in decisions:
        title = decision.get("title", "")
        mode = decision.get("extract_mode", "table")
        if title in entries:
            entries[title]["extract_mode"] = mode
            # 计算 scope（子表边界）
            entry = entries[title]
            start_row = entry.get("start_row", 1)
            start_col = entry.get("start_col", 1)
            entry["scope"] = {
                "start_row": start_row,
                "end_row": start_row + 30,  # 搜索 30 行
                "start_col": start_col,
                "end_col": start_col + 20,  # 搜索 20 列
            }

    # 如果所有子表都是 KV，切换到 KV 模式
    all_kv = all(
        entries.get(t, {}).get("extract_mode") == "kv"
        for t in subtable_titles
        if t in entries
    )
    if all_kv and entries:
        config["extract_type"] = "kv"
        # 从 KV 子表的 headers 中构建 kv_list（去重）
        all_kv_keys = []
        seen = set()
        for title in subtable_titles:
            for h in headers_map.get(title, []):
                norm_h = h.lower().strip()
                if norm_h not in seen:
                    seen.add(norm_h)
                    all_kv_keys.append(h)
        if all_kv_keys:
            config["kv_list"] = all_kv_keys

    return {"cache": cache_state, "config": config}

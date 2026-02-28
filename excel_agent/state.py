"""
state.py — Agent 全局状态定义

Auto 模式：只需传入 sheet_name + subtable_title，Agent 自动完成：
  1. 探索 Sheet 结构，定位子表
  2. 发现子表全部列（含多级表头）
  3. 识别哪些列需要向前填充（如 SYSTEM MODULE 跨行合并列）
  4. 输出完整 JSON 二维数组
"""

from typing import TypedDict, Optional, List, Dict, Any


class ExtractionConfig(TypedDict):
    """
    用户传入的抽取配置，只需填两个必填项。

    必填：
        sheet_name      : 目标 Sheet 名，如 "CONFIGURATION"
        subtable_title  : 子表标题关键词，如 "3G Configuration"

    可选：
        hints           : 额外定位提示，帮助 LLM 更准确找到子表
                          如 "子表位于 Sheet 上半部分"
    """
    sheet_name:     str
    subtable_title: str
    hints:          Optional[str]


class HeaderMap(TypedDict):
    """
    locate_node (LLM) 的输出：子表位置 + 完整列结构。

    column_map key 格式：
        有父级  →  "ANTENNAS||Antenna Qty."
        无父级  →  "None||CELL"
    值为该列在 Sheet 中的整数列号（从 1 开始，与 openpyxl 一致）。

    all_columns 是有序列表，保持列的原始顺序，供 extract/restore 使用。

    merge_fill_keys 是需要向前填充的列的 key 集合，
    对应 SYSTEM MODULE 这种"一个值覆盖多行"的跨行合并列。
    """
    header_row_count:   int           # 表头行数（1 或 2）
    subtable_start_row: int           # 子表起始行（含表头第一行）
    subtable_end_row:   int           # 子表结束行（含最后数据行）
    subtable_start_col: int           # 子表起始列
    subtable_end_col:   int           # 子表结束列
    column_map:         Dict[str, int]  # "parent||child" → 列号
    all_columns:        List[Dict]    # [{"parent": str|None, "child": str, "col_idx": int}, ...]
    merge_fill_keys:    List[str]     # 需要向前填充的列 key，如 ["None||SYSTEM MODULE"]


class AgentState(TypedDict):

    # ── 输入 ──────────────────────────────────────────────────
    excel_path: str
    config:     ExtractionConfig

    # ── parse_node 输出 ────────────────────────────────────────
    sheet_structure: Optional[Dict[str, Any]]
    # {
    #   "sheet_name": str,
    #   "max_row": int, "max_col": int,
    #   "non_empty_cells": [...],
    #   "potential_headers": [...],    # 粗体/合并单元格（候选表头）
    #   "merged_cells_info": [...],    # 所有合并单元格详情
    #   "potential_subtables": [...]   # 基于空行切分的候选子表块
    # }

    # ── locate_node 输出 ───────────────────────────────────────
    header_map: Optional[HeaderMap]

    # ── extract_node 输出 ──────────────────────────────────────
    raw_data: Optional[List[List[Any]]]
    # 二维列表，行=数据行（不含表头），列按 all_columns 顺序排列
    # None 表示空单元格（normalize 后变为 ""）

    # ── restore_node 输出 ─────────────────────────────────────
    result: Optional[List[List[str]]]
    # 第 0 行为列名，后续为数据行，空值统一为 ""

    # ── 质量控制 ───────────────────────────────────────────────
    quality_score: float   # 0.0 ~ 1.0，阈值 0.75
    retry_count:   int     # 当前重试次数，上限 3
    errors:        List[str]

    # ── 最终交付 ───────────────────────────────────────────────
    final_output: Optional[List[List[str]]]

"""
nodes/kv_feedback.py — KV 多模态反馈处理节点

功能：
    1. 接收用户在前端点选的纠偏数据
    2. 分析原始代码提取坐标与用户指定正确坐标的偏移
    3. 生成修正后的抽取代码

用户反馈数据结构：
{
    "key": "飞机型号",
    "extracted_value": "波音",
    "extracted_coord": [4, 4],  # [row, col]
    "correct_value": "A350-900",
    "correct_coord": [5, 5],     # [row, col]
    "feedback_type": "value_offset"  # "value_offset" | "key_not_found" | "wrong_key"
}
"""

import re
import json
import os
from pathlib import Path
from typing import Dict, Any, Optional, List

from langchain_openai import ChatOpenAI
from langchain_core.messages import SystemMessage, HumanMessage

from excel_agent.state import AgentState


_SYSTEM_PROMPT = """\
你是 Excel KV 抽取代码修复专家。你的任务是根据用户反馈的纠偏数据，修复有问题的 KV 抽取代码。

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
【输入上下文】
  - original_code: 原始生成的 KV 抽取代码
  - kv_list: 待抽取的 Key 列表
  - feedback: 用户反馈的纠偏数据，包含：
      - key: 哪个 Key 抽取错误
      - extracted_value: 你的代码提取的错误值
      - extracted_coord: 你的代码提取的坐标 [row, col]
      - correct_value: 用户指定的正确值
      - correct_coord: 用户指定的正确坐标 [row, col]
      - feedback_type: 反馈类型
  - sheet_structure: Excel 结构信息

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
【修复策略】

1. value_offset (值偏移):
   - 症状：Key 定位正确，但 Value 扫描方向/距离错误
   - 修复：调整 scan_value 函数的扫描逻辑
   - 示例：原代码向右扫描，但正确值在下方 → 改为先向下扫描

2. key_not_found (Key 未找到):
   - 症状：Key 定位函数未找到匹配的单元格
   - 修复：检查 _normalize 函数是否过度归一化，或增加模糊匹配容错

3. wrong_key (Key 定位错误):
   - 症状：定位到了错误的单元格（同名 Key 或多个匹配项）
   - 修复：增加 Key 定位的唯一性校验，或调整搜索范围

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
【代码修复范式】

```python
def extract(ws, merged_map, kv_list):
    def cell_val(r, c):
        v = merged_map.get((r, c), ws.cell(row=r, column=c).value)
        if v is None: return ""
        if isinstance(v, float) and v == int(v): return str(int(v))
        return str(v).replace('\\n', ' ').replace('\\r', '').strip()

    def get_merged_box(r, c):
        info = merged_map.get((r, c))
        if info and isinstance(info, dict) and "top_left" in info:
            tl = info["top_left"]
            return {
                "min_row": tl[0], "max_row": tl[0] + info["r_span"] - 1,
                "min_col": tl[1], "max_col": tl[1] + info["c_span"] - 1
            }
        return {"min_row": r, "max_row": r, "min_col": c, "max_col": c}

    def _normalize(text):
        if text is None: return ""
        return re.sub(r'[\\s\\-\\_\\(\\)\\[\\]\\/\\.,]+', '', str(text).lower())

    def find_key_cell(key_text):
        key_norm = _normalize(key_text)
        for r in range(1, ws.max_row + 1):
            for c in range(1, ws.max_column + 1):
                val = merged_map.get((r, c), ws.cell(row=r, column=c).value)
                if val and key_norm in _normalize(val):
                    box = get_merged_box(r, c)
                    return r, c, box
        return None, None, None

    def scan_value(key_r, key_c, key_box):
        # 修复重点：根据用户反馈调整扫描优先级
        # 如果用户反馈正确值在下方，优先向下扫描
        # 如果用户反馈正确值在右方，优先向右扫描

        # 示例：优先向下扫描（根据反馈调整）
        start_r = key_box["max_row"] + 1
        for r in range(start_r, ws.max_row + 1):
            val = cell_val(r, key_c)
            if val:
                return val, (r, key_c)

        # 其次向右扫描
        start_c = key_box["max_col"] + 1
        for c in range(start_c, ws.max_column + 1):
            val = cell_val(key_r, c)
            if val:
                return val, (key_r, c)

        return "", None

    result = {}
    for key in kv_list:
        key_r, key_c, key_box = find_key_cell(key)
        if key_r is None:
            result[key] = ""
            continue
        val, _ = scan_value(key_r, key_c, key_box)
        result[key] = val

    return result
```

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
【输出格式】
严格返回 JSON：
{
  "refactored_code": "def extract(ws, merged_map, kv_list):\\n    ...",
  "fix_explanation": "根据用户反馈，正确值在 Key 的下方而非右方，因此调整了 scan_value 函数的扫描优先级。"
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
    return ChatOpenAI(model=model, temperature=0.1, api_key=api_key, base_url=base_url)


def _extract_json(raw: str) -> Optional[Dict]:
    """从 LLM 输出中提取 JSON"""
    m = re.search(r"```json\s*([\s\S]*?)```", raw)
    if m:
        try:
            return json.loads(m.group(1).strip())
        except:
            pass
    m = re.search(r"```\s*([\s\S]*?)```", raw)
    if m:
        try:
            return json.loads(m.group(1).strip())
        except:
            pass
    try:
        return json.loads(raw.strip())
    except:
        return None


def kv_feedback_node(state: AgentState, feedback: Dict[str, Any]) -> dict:
    """
    KV 多模态反馈处理节点

    Args:
        state: AgentState
        feedback: 用户反馈数据
            {
                "key": str,
                "extracted_value": str,
                "extracted_coord": List[int],
                "correct_value": str,
                "correct_coord": List[int],
                "feedback_type": str
            }

    Returns:
        {
            "refactored_code": str,
            "fix_explanation": str,
            "success": bool
        }
    """
    config = state.get("config", {})
    kv_list = config.get("kv_list", [])
    sheet_structure = state.get("sheet_structure", {})
    generated_code = state.get("generated_code", [])
    kv_result = state.get("kv_result", {})

    # 获取原始代码
    original_code = generated_code[0] if generated_code else ""

    # 组装 Prompt 上下文
    ctx = {
        "original_code": original_code,
        "kv_list": kv_list,
        "feedback": feedback,
        "sheet_structure": {
            "max_row": sheet_structure.get("max_row", 0),
            "max_col": sheet_structure.get("max_col", 0),
            "sample_cells": sheet_structure.get("non_empty_cells", [])[:50]
        },
        "extracted_kv": kv_result
    }

    messages = [
        SystemMessage(content=_SYSTEM_PROMPT),
        HumanMessage(content=json.dumps(ctx, ensure_ascii=False, indent=2))
    ]

    try:
        response = _get_llm().invoke(messages)
        parsed = _extract_json(response.content)

        if parsed and "refactored_code" in parsed:
            return {
                "refactored_code": parsed["refactored_code"],
                "fix_explanation": parsed.get("fix_explanation", ""),
                "success": True
            }
        else:
            return {
                "refactored_code": original_code,
                "fix_explanation": "LLM 未能生成有效的修复代码",
                "success": False
            }
    except Exception as e:
        return {
            "refactored_code": original_code,
            "fix_explanation": f"修复过程出错：{str(e)}",
            "success": False
        }


def analyze_feedback_offset(feedback: Dict[str, Any]) -> Dict[str, Any]:
    """
    分析用户反馈的偏移类型

    Args:
        feedback: 用户反馈数据

    Returns:
        分析结果：
        {
            "direction": "right" | "down" | "diagonal",
            "distance": int,
            "suggested_fix": str
        }
    """
    extracted = feedback.get("extracted_coord", [])
    correct = feedback.get("correct_coord", [])

    if not extracted or not correct:
        return {"direction": "unknown", "distance": 0, "suggested_fix": "无法分析坐标"}

    row_diff = correct[0] - extracted[0]
    col_diff = correct[1] - extracted[1]

    direction = "unknown"
    suggested_fix = ""

    if row_diff == 0 and col_diff > 0:
        direction = "right"
        suggested_fix = f"正确值在提取位置右侧 {col_diff} 列，建议增加向右扫描的距离"
    elif row_diff == 0 and col_diff < 0:
        direction = "left"
        suggested_fix = f"正确值在提取位置左侧 {abs(col_diff)} 列，建议增加向左扫描逻辑"
    elif row_diff > 0 and col_diff == 0:
        direction = "down"
        suggested_fix = f"正确值在提取位置下方 {row_diff} 行，建议增加向下扫描的距离"
    elif row_diff < 0 and col_diff == 0:
        direction = "up"
        suggested_fix = f"正确值在提取位置上方 {abs(row_diff)} 行，建议增加向上扫描逻辑"
    elif row_diff > 0 and col_diff > 0:
        direction = "diagonal_down_right"
        suggested_fix = f"正确值在提取位置右下方（下{row_diff}行，右{col_diff}列），可能需要调整 Key 定位逻辑"
    elif row_diff > 0 and col_diff < 0:
        direction = "diagonal_down_left"
        suggested_fix = f"正确值在提取位置左下方（下{row_diff}行，左{abs(col_diff)}列），可能需要调整 Key 定位逻辑"
    elif row_diff < 0 and col_diff > 0:
        direction = "diagonal_up_right"
        suggested_fix = f"正确值在提取位置右上方（上{abs(row_diff)}行，右{col_diff}列），可能需要调整 Key 定位逻辑"
    elif row_diff < 0 and col_diff < 0:
        direction = "diagonal_up_left"
        suggested_fix = f"正确值在提取位置左上方（上{abs(row_diff)}行，左{abs(col_diff)}列），可能需要调整 Key 定位逻辑"
    else:
        direction = "same"
        suggested_fix = "提取位置与正确位置相同，可能是值解析错误"

    return {
        "direction": direction,
        "row_diff": row_diff,
        "col_diff": col_diff,
        "distance": abs(row_diff) + abs(col_diff),
        "suggested_fix": suggested_fix
    }

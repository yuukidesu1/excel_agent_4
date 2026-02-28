"""
nodes/locate.py — LLM 节点：自动定位子表 + 发现全部列结构

这是 Agent 唯一调用 LLM 的节点，负责所有需要语义理解的任务：

  【任务一】定位目标子表
      在候选子表块中找到标题匹配 subtable_title 的那个块，
      确定精确的行列边界。

  【任务二】识别表头结构
      判断表头占几行（1行 or 2行）。
      2行时：第一行为父级（可能跨列合并），第二行为子级。

  【任务三】枚举所有列（对标灵犀核心能力）
      对子表内每一列输出完整规格：
        - 单级：parent=null, child="列名"
        - 双级：parent="父级名", child="子级名"
      同时记录每列的实际列号（从 1 开始的整数）。

  【任务四】识别需要向前填充的列（对标灵犀 SYSTEM MODULE 处理）
      若某列存在跨行合并（一个值对应多行，其余行为空），
      标记该列需要 forward-fill。

输入精简原则：
    只传 potential_subtables + potential_headers + merged_cells_info，
    不传全量单元格，避免 token 浪费和 LLM 注意力分散。
"""

import json
import re
import os
from pathlib import Path

from langchain_openai import ChatOpenAI
from langchain_core.messages import SystemMessage, HumanMessage

from excel_agent.state import AgentState, HeaderMap


# ─────────────────────────────────────────────────────────────
# LLM 懒加载（函数内初始化，避免导入时读环境变量报错）
# ─────────────────────────────────────────────────────────────

def _get_llm() -> ChatOpenAI:
    from dotenv import load_dotenv
    # 从本文件往上逐级查找 .env
    for parent in Path(__file__).resolve().parents:
        env_file = parent / ".env"
        if env_file.exists():
            load_dotenv(env_file, override=True)
            break

    api_key  = os.getenv("OPENAI_API_KEY")
    base_url = os.getenv("OPENAI_BASE_URL") or None
    model    = os.getenv("LLM_MODEL", "glm-4")

    if not api_key:
        raise ValueError(
            "未找到 OPENAI_API_KEY。\n"
            "请在项目根目录创建 .env 文件：\n"
            "  OPENAI_API_KEY=你的APIKey\n"
            "  OPENAI_BASE_URL=https://open.bigmodel.cn/api/paas/v4/\n"
            "  LLM_MODEL=glm-4\n"
        )

    return ChatOpenAI(model=model, temperature=0, api_key=api_key, base_url=base_url)


# ─────────────────────────────────────────────────────────────
# System Prompt
# ─────────────────────────────────────────────────────────────

_SYSTEM_PROMPT = """\
你是 Excel 结构分析专家，擅长从复杂 Excel 中自动发现并还原子表的完整列结构。

你会收到一个 Sheet 的结构描述，完成以下四个步骤：

【第一步】定位目标子表
  在 potential_subtables 中找到 title 最匹配 subtable_title 的块。
  忽略大小写差异，允许部分匹配。
  确定子表边界：subtable_start_row、subtable_end_row、subtable_start_col、subtable_end_col。

【第二步】识别表头结构
  分析子表起始区域的 potential_headers 和 merged_cells_info：
  - header_row_count = 1：只有一行表头，直接是列名
  - header_row_count = 2：两行表头，第一行是父级（跨列合并），第二行是子级

【第三步】枚举子表所有列（关键任务）
  对子表列范围内的每一列，输出其完整规格：
  - 单级表头：{ "parent": null, "child": "列名", "col_idx": 列号 }
  - 双级表头：{ "parent": "父级名", "child": "子级名", "col_idx": 列号 }
  col_idx 为该列在 Sheet 中从 1 开始的整数列号。
  column_map 的 key 格式：有父级 → "父级||子级"，无父级 → "None||列名"

【第四步】识别需要向前填充的列
  如果某列存在跨行合并（一个值逻辑上覆盖多行，其余行为空或合并空白），
  将其 key 加入 merge_fill_keys。
  典型例子：SYSTEM MODULE、站点编号等分组列。

严格只返回 JSON，不要任何解释、前缀、markdown 代码块。

返回格式：
{
  "subtable_start_row": <int>,
  "subtable_end_row":   <int>,
  "subtable_start_col": <int>,
  "subtable_end_col":   <int>,
  "header_row_count":   <1 或 2>,
  "all_columns": [
    { "parent": null,       "child": "SYSTEM MODULE", "col_idx": 1 },
    { "parent": null,       "child": "CELL",          "col_idx": 2 },
    { "parent": "RF MODUL", "child": "Type",          "col_idx": 3 },
    { "parent": "ANTENNAS", "child": "NEW/SWAP/EXIST","col_idx": 8 },
    ...
  ],
  "column_map": {
    "None||SYSTEM MODULE":      1,
    "None||CELL":               2,
    "RF MODUL||Type":           3,
    "ANTENNAS||NEW/SWAP/EXIST": 8,
    ...
  },
  "merge_fill_keys": ["None||SYSTEM MODULE"],
  "confidence": <0.0~1.0>,
  "reason":     "<简要说明定位依据，供调试>"
}
"""


# ─────────────────────────────────────────────────────────────
# JSON 解析（容错）
# ─────────────────────────────────────────────────────────────

def _parse_json(raw: str) -> dict:
    raw = raw.strip()
    if "```" in raw:
        m = re.search(r"```(?:json)?\s*([\s\S]*?)```", raw)
        if m:
            raw = m.group(1).strip()
    return json.loads(raw)


# ─────────────────────────────────────────────────────────────
# 节点函数
# ─────────────────────────────────────────────────────────────

def locate_node(state: AgentState) -> dict:
    """
    输入：sheet_structure + config + errors（重试时）
    输出：header_map
    """
    structure = state["sheet_structure"]
    config    = state["config"]
    errors    = state.get("errors", [])

    # 构造给 LLM 的精简上下文
    ctx = {
        "subtable_title":      config["subtable_title"],
        "sheet_size":          f"{structure['max_row']} 行 × {structure['max_col']} 列",
        "potential_subtables": structure["potential_subtables"],
        # 最多传 100 条候选表头（粗体/合并单元格）
        "potential_headers":   structure["potential_headers"][:100],
        # 最多传 80 条合并单元格（多级表头解析必需）
        "merged_cells_info":   structure["merged_cells_info"][:80],
    }

    if config.get("hints"):
        ctx["hints"] = config["hints"]

    # 重试时附上前次错误，让 LLM 针对性修正
    if errors:
        ctx["previous_errors"]  = errors[-3:]
        ctx["retry_instruction"] = "上次定位失败或质量不达标，请仔细阅读错误，重新分析并修正。"

    messages = [
        SystemMessage(content=_SYSTEM_PROMPT),
        HumanMessage(content=json.dumps(ctx, ensure_ascii=False, indent=2)),
    ]

    response = _get_llm().invoke(messages)
    data     = _parse_json(response.content)

    # 过滤 column_map 中值为 null 的项
    column_map = {k: v for k, v in data.get("column_map", {}).items() if v is not None}

    header_map: HeaderMap = {
        "header_row_count":   int(data.get("header_row_count", 1)),
        "subtable_start_row": int(data["subtable_start_row"]),
        "subtable_end_row":   int(data["subtable_end_row"]),
        "subtable_start_col": int(data["subtable_start_col"]),
        "subtable_end_col":   int(data["subtable_end_col"]),
        "column_map":         column_map,
        "all_columns":        data.get("all_columns", []),
        "merge_fill_keys":    data.get("merge_fill_keys", []),
    }

    return {"header_map": header_map}

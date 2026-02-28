"""
nodes/locate.py — LLM 节点：自动定位子表 + 发现全部列结构

核心修复：强化双行表头识别。
当 Excel 有父级（跨列合并）+ 子级两行表头时，LLM 必须：
  - 输出 header_row_count = 2
  - all_columns 中每列使用 parent=父级名, child=子级名
  - 不能把父级名当成 child 输出
"""

import json
import re
import os
from pathlib import Path

from langchain_openai import ChatOpenAI
from langchain_core.messages import SystemMessage, HumanMessage

from excel_agent.state import AgentState


_SYSTEM_PROMPT = """\
你是 Excel 结构分析专家。给你一份 Sheet 的结构描述，完成以下四步：

【第一步】定位目标子表
  在 subtables 中找到 title 最匹配 subtable_title 的块（忽略大小写，允许部分匹配）。
  输出：subtable_start_row, subtable_end_row, subtable_start_col, subtable_end_col

【第二步】识别表头行数（极其重要，请仔细判断）

  判断依据：检查 merged_cells_info 中，是否有位于子表起始行、且 col_span > 1 的合并单元格。

  ★ 如果存在 col_span > 1 的合并单元格（即某个值横跨多列），说明是【双行表头】：
      - 第一行（父级行）：跨列合并的大类名，如 "RF MODULE"、"ANTENNAS"、"RRU Cable"、"TILT"、"Power Splitter"
      - 第二行（子级行）：每列对应的具体列名，如 "TYPE"、"QTY."、"NEW/SWAP/EXIST"、"M"、"E"
      - 此时 header_row_count = 2
      - subtable_start_row 指向父级行（第一表头行）

  ★ 如果没有 col_span > 1 的合并单元格，或所有合并都是 col_span=1（即跨行合并），
    则是【单行表头】，header_row_count = 1。

  ⚠️ 注意：子表的标题行（如"4G Configuration"所在的合并行）不是表头行，
    subtable_start_row 应指向真正的列名所在行，跳过标题行。

【第三步】枚举子表所有列

  双行表头时（header_row_count=2）：
    - parent = 第一行的父级大类名（跨列合并单元格的值）
    - child  = 第二行该列的具体列名
    - 对于没有父级（第一行为空或该列不属于任何合并区域）的列，parent=null
    ❌ 错误：{"parent": null, "child": "ANTENNAS", "col_idx": 9}  ← ANTENNAS是父级，不是child
    ✅ 正确：{"parent": "ANTENNAS", "child": "NEW/SWAP/EXIST", "col_idx": 9}
    ✅ 正确：{"parent": "ANTENNAS", "child": "Antenna Type",   "col_idx": 10}
    ✅ 正确：{"parent": "ANTENNAS", "child": "Antenna Qty.",   "col_idx": 11}

  单行表头时（header_row_count=1）：
    - parent = null，child = 该行的列名

  col_idx 为该列在 Sheet 中从 1 开始的整数列号。
  column_map key 格式：有父级 → "父级||子级"，无父级 → "None||列名"

【第四步】识别需要向前填充的列
  某列存在跨行合并（row_span > 1，一个值覆盖多行，其余行空白），加入 merge_fill_keys。
  典型：SYSTEM MODULE、站点编号等分组列。

严格只返回 JSON，不要任何解释或 markdown 代码块。

返回格式（以双行表头为例）：
{
  "subtable_start_row": 26,
  "subtable_end_row":   39,
  "subtable_start_col": 1,
  "subtable_end_col":   28,
  "header_row_count":   2,
  "all_columns": [
    { "parent": null,         "child": "SYSTEM MODULE",    "col_idx": 2 },
    { "parent": null,         "child": "CELL",             "col_idx": 3 },
    { "parent": "RF MODULE",  "child": "TYPE",             "col_idx": 4 },
    { "parent": "RF MODULE",  "child": "QTY.",             "col_idx": 5 },
    { "parent": null,         "child": "UBBP type",        "col_idx": 6 },
    { "parent": null,         "child": "BAND",             "col_idx": 7 },
    { "parent": null,         "child": "TRX",              "col_idx": 8 },
    { "parent": "ANTENNAS",   "child": "NEW/SWAP/EXIST",   "col_idx": 9 },
    { "parent": "ANTENNAS",   "child": "Antenna Type",     "col_idx": 10 },
    { "parent": "ANTENNAS",   "child": "Antenna Qty.",     "col_idx": 11 },
    { "parent": "RRU Cable",  "child": "POWER LENGTH(m)",  "col_idx": 12 },
    { "parent": "RRU Cable",  "child": "OPT LENGTH(m)",    "col_idx": 13 },
    { "parent": "RRU Cable",  "child": "JUMPER TYPE",      "col_idx": 14 },
    { "parent": "RRU Cable",  "child": "JUMPER LENGTH(m)", "col_idx": 15 },
    { "parent": "RRU Cable",  "child": "Q-TY",             "col_idx": 16 },
    { "parent": null,         "child": "Direction",        "col_idx": 17 },
    { "parent": "TILT",       "child": "M",                "col_idx": 18 },
    { "parent": "TILT",       "child": "E",                "col_idx": 19 }
  ],
  "column_map": {
    "None||SYSTEM MODULE":       2,
    "None||CELL":                3,
    "RF MODULE||TYPE":           4,
    "RF MODULE||QTY.":           5,
    "None||UBBP type":           6,
    "None||BAND":                7,
    "None||TRX":                 8,
    "ANTENNAS||NEW/SWAP/EXIST":  9,
    "ANTENNAS||Antenna Type":    10,
    "ANTENNAS||Antenna Qty.":    11,
    "RRU Cable||POWER LENGTH(m)":12,
    "RRU Cable||OPT LENGTH(m)":  13,
    "RRU Cable||JUMPER TYPE":    14,
    "RRU Cable||JUMPER LENGTH(m)":15,
    "RRU Cable||Q-TY":           16,
    "None||Direction":           17,
    "TILT||M":                   18,
    "TILT||E":                   19
  },
  "merge_fill_keys": ["None||SYSTEM MODULE"],
  "confidence": 0.95,
  "reason": "检测到 RF MODULE、ANTENNAS、RRU Cable、TILT 等跨列合并单元格，确认为双行表头。"
}
"""


def _get_llm() -> ChatOpenAI:
    from dotenv import load_dotenv
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
            "未找到 OPENAI_API_KEY，请在项目根目录创建 .env 文件：\n"
            "  OPENAI_API_KEY=你的APIKey\n"
            "  OPENAI_BASE_URL=https://open.bigmodel.cn/api/paas/v4/\n"
            "  LLM_MODEL=glm-4"
        )
    return ChatOpenAI(model=model, temperature=0, api_key=api_key, base_url=base_url)


def _parse_json(raw: str) -> dict:
    raw = raw.strip()
    m = re.search(r"```(?:json)?\s*([\s\S]*?)```", raw)
    if m:
        raw = m.group(1).strip()
    return json.loads(raw)


def locate_node(state: AgentState) -> dict:
    # st     = state["sheet_structure"]
    st     = state.get("sheet_structure")
    if not st:
        raise ValueError("严重错误：未能获取到 sheet_structure！请检查 parse 节点是否正常执行，或目标 Sheet 页是否存在。")
    errors = state.get("errors", [])

    # ------------------DEBUG
    # print(f"\n[Locate Node] 接收到的 sheet_structure 的键有: {list(st.keys())}\n")

    ctx = {
        "subtable_title":    state["config"]["subtable_title"],
        "sheet_size":        f"{st['max_row']} 行 × {st['max_col']} 列",
        "subtables":         st.get("potential_headers", []),
        "potential_headers": st.get("potential_headers", [])[:120],
        "merged_cells_info": st.get("merged_cells_info", [])[:100],
    }
    if state.get("hints"):
        ctx["hints"] = state["hints"]
    if errors:
        ctx["previous_errors"]   = errors[-3:]
        ctx["retry_instruction"] = "上次失败，请仔细阅读错误，重新分析修正。特别注意双行表头的识别。"

    messages = [
        SystemMessage(content=_SYSTEM_PROMPT),
        HumanMessage(content=json.dumps(ctx, ensure_ascii=False, indent=2)),
    ]

    data = _parse_json(_get_llm().invoke(messages).content)

    header_map = {
        "header_row_count":   int(data.get("header_row_count", 1)),
        "subtable_start_row": int(data["subtable_start_row"]),
        "subtable_end_row":   int(data["subtable_end_row"]),
        "subtable_start_col": int(data["subtable_start_col"]),
        "subtable_end_col":   int(data["subtable_end_col"]),
        "all_columns":        data.get("all_columns", []),
        "column_map":         {k: v for k, v in data.get("column_map", {}).items() if v is not None},
        "merge_fill_keys":    data.get("merge_fill_keys", []),
    }
    return {"header_map": header_map}
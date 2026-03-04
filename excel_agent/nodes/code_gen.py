"""
nodes/code_gen.py — LLM 节点：看 parse 数据，写 openpyxl 抽取代码

LLM 的任务：
    1. 理解 sheet_structure（子表位置、合并单元格、表头结构）
    2. 根据 subtable_titles 定位目标子表
    3. 编写完整的 Python 抽取函数，使用 openpyxl 直接读取数据
    4. 函数必须返回 List[List[str]]（标准二维数组，第0行为列名）

重试时携带上次的错误信息（代码执行异常 or 质量问题），让 LLM 修正。

代码规范（写入 System Prompt）：
    - 必须定义 extract(ws, merged_map) 函数
    - merged_map 是预建的合并填充图 {(row,col): value}，直接用
    - 返回值必须是 List[List[str]]，第0行为列名
    - 不能 import openpyxl（ws 已由 sandbox 传入）
    - 不能读文件、不能访问网络、不能使用 os/sys/subprocess
"""

import json
import re
import os
from pathlib import Path

from langchain_openai import ChatOpenAI
from langchain_core.messages import SystemMessage, HumanMessage

from excel_agent.state import AgentState


_SYSTEM_PROMPT = """\
你是 Excel 数据抽取专家。我会给你一份 Excel Sheet 的完整结构描述，
你需要编写一段 Python 代码，使用已加载好的 openpyxl Worksheet 对象抽取指定子表的数据。

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
你收到的信息：
  - subtable_titles      : 目标子表的标题关键词列表（List[str]）
  - target_columns      : 需要抽取的列配置。格式为字典，键为子表名，值为该表的列配置（None = 抽取全部列）
  - sheet_structure     : Sheet 的完整结构，包含：
      · potential_subtables : 基于空行切分的候选子表块（含起止行）
      · potential_headers   : 粗体/合并单元格（候选表头）
      · merged_cells_info   : 所有合并单元格详情（range, min/max row/col, value, row_span, col_span）
      · non_empty_cells     : 所有非空单元格（row, col, value）
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

你必须编写一个名为 `extract` 的函数，签名如下：

```python
def extract(ws, merged_map: dict) -> list:
    ...
    return result  # Dict[str, List[List[str]]]
```

参数说明：
  ws         : openpyxl Worksheet 对象，已加载好，直接读取即可
  merged_map : 合并单元格填充图，键为 (row, col) 元组，值为该合并区域的实际值
               使用方式：val = merged_map.get((row, col), ws.cell(row, col).value)

返回值规范：
  - 必须返回一个字典（dict）。
  - 字典的键（key）为用户传入的 subtable_titless 中的原名。
  - 字典的值（value）为该字表对应的 List[List[str]]]，所有值转为字符串，None/空值转为""。
  - 每个二维表的第 0 行为列名列表，后续为数据行。

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
重要规则：
  1. 只能使用标准库（re, json, math 等），不能 import openpyxl
  2. 合并单元格必须通过 merged_map 处理，不要直接相信 ws.cell().value（合并区域非左上角格为 None）
  3. 若有双行表头（父级跨列合并 + 子级），需正确拼合列名（如 "ANTENNAS||Antenna Qty."）
     或直接用子级列名（取决于是否有 target_columns 指定了 parent）
  4. 行方向子表、列方向子表、嵌套子表等非常规结构，都要正确处理
  5. 代码必须健壮：行列边界要用变量，不要硬编码"第16行开始"之类的数字
  6. 必须为传入的每一个子表独立提取一份数据，放入返回的字典中。如果某个表在 Excel 中不存在，该键对应的值返回空列表 []。
  7. 只返回代码，不要任何解释。代码包在 ```python ... ``` 中
  8. ⚠️ 绝对禁止在生成的代码中写 `for` 循环去搜索标题、表头或匹配字符串！
     你现在已经看到了 sheet_structure，你必须在你的“大脑”里计算好真实的行号和列号。
     在代码中直接使用硬编码的具体数字。
     ✅ 正确示例：subtable_start_row = 19
     ❌ 错误示例：for header in sheet_structure... if "group3" in header...
  9. 注意：用户提供的 subtable_titles 和 target_columns 可能存在拼写错误或缩写。
     并且可能有多个 subtable_titles, 通常会以 "," 分隔开,例如 subtable_titles="A, B, C, ..."如果有多个 subtable_titles 需要分别独立提取成多个二维数组。
     请发挥你的智能，在 sheet_structure 中找到语义最接近的真实区域和真实列号，
     然后将这些真实的行号、列号、列名写死在你的代码中。
  10. ★ 可用的上下文变量（沙盒中已注入，直接使用）：
       sheet_structure  — Sheet 完整结构（subtables/potential_headers/merged_cells_info 等）
       subtable_titles   — 目标子表标题字符串
       target_columns   — 用户指定的目标列列表（可能为 None）
       hints            — 额外提示（可能为 None）

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
代码模板（参考，根据实际结构调整）：

```python
def extract(ws, merged_map: dict) -> dict:
    def cell_val(r, c):
        v = merged_map.get((r, c), ws.cell(row=r, column=c).value)
        if v is None: return ""
        if isinstance(v, float) and v == int(v): return str(int(v))
        return str(v).strip()

    result = {}
    
    # ==== 提取表1 (硬编码坐标) ====
    # 假设第一个标题为 "Group 1"
    headers_1 = [cell_val(10, c) for c in range(1, 10)]
    rows_1 = [headers_1]
    for r in range(11, 20):
        rows_1.append([cell_val(r, c) for c in range(1, 10)])
    result["Group 1"] = rows_1
    
    # ==== 提取表2 (硬编码坐标) ====
    # ...
    # result["Group 2"] = rows_2

    return result
```
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
        raise ValueError("未找到 OPENAI_API_KEY，请检查 .env 文件。")
    return ChatOpenAI(model=model, temperature=0, api_key=api_key, base_url=base_url)


def _extract_code(raw: str) -> str:
    """从 LLM 输出中提取 ```python ... ``` 之间的代码"""
    m = re.search(r"```python\s*([\s\S]*?)```", raw)
    if m:
        return m.group(1).strip()
    # 没有代码块标记，直接返回（容错）
    return raw.strip()


def code_gen_node(state: AgentState) -> dict:
    """
    输入：sheet_structure + subtable_titles + target_columns
    输出：generated_code（完整的 extract 函数字符串）
    """
    st            = state["sheet_structure"]
    errors        = state.get("errors", [])
    sandbox_error = state.get("sandbox_error")

    # 兼容 config 嵌套结构和平铺结构两种 state 设计
    config         = state.get("config", {})
    subtable_titles = config.get("subtable_titles") or state.get("subtable_titles", [])
    target_columns = config.get("target_columns") or state.get("target_columns")
    hints          = config.get("hints")          or state.get("hints")

    # ── 构造给 LLM 的上下文 ──────────────────────────────────
    ctx: dict = {
        "subtable_titles": subtable_titles,
        "target_columns": target_columns,
        "sheet_structure": {
            "sheet_name":        st["sheet_name"],
            "max_row":           st["max_row"],
            "max_col":           st["max_col"],
            "subtables":         st["potential_subtables"],
            "potential_headers": st["potential_headers"][:150],
            "merged_cells_info": st["merged_cells_info"][:120],
            "sample_cells":      [
                c for c in st["non_empty_cells"]
                if any(
                    sub["start_row"] - 2 <= c["row"] <= sub["end_row"] + 2
                    for sub in st["potential_subtables"]
                    # 遍历判断是否属于目标标题之一
                    if any(t.lower() in sub.get("title", "").lower() or
                           any(t.lower() in tc.lower() for tc in sub.get("title_candidates", []))
                           for t in subtable_titles)
                )
            ][:200],
        },
    }

    if hints:
        ctx["hints"] = hints

    # 重试时附上错误，让 LLM 针对性修正
    if sandbox_error:
        ctx["last_code_error"] = sandbox_error
        ctx["retry_instruction"] = (
            "上次生成的代码执行时出错，请仔细阅读错误信息修正代码。"
            "特别注意：子表标题匹配必须使用模糊匹配（忽略大小写和空格），"
            "不能用 == 精确匹配；合并单元格必须用 merged_map 处理；"
            "行列边界必须从 sheet_structure 动态计算，不能硬编码数字。"
        )
    elif errors:
        ctx["last_quality_errors"] = errors[-3:]
        ctx["retry_instruction"] = (
            "上次代码执行成功但质量不达标，请根据质量问题修正。"
            "常见问题：数据行为空、列数不对、空值率过高。"
        )

    messages = [
        SystemMessage(content=_SYSTEM_PROMPT),
        HumanMessage(content=json.dumps(ctx, ensure_ascii=False, indent=2)),
    ]

    response = _get_llm().invoke(messages)
    code     = _extract_code(response.content)

    return {"generated_code": code, "sandbox_error": None}  # 清空上次的沙盒错误
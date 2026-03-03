"""
nodes/code_gen.py — LLM 节点：看 parse 数据，写 openpyxl 抽取代码

LLM 的任务：
    1. 理解 sheet_structure（子表位置、合并单元格、表头结构）
    2. 根据 subtable_title 定位目标子表
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
  - subtable_title      : 目标子表的标题关键词
  - target_columns      : 需要抽取的列（None = 抽取全部列）
  - sheet_structure     : Sheet 的完整结构，包含：
      · subtables           : 基于空行切分的候选子表块（含起止行）
      · potential_headers   : 粗体/合并单元格（候选表头）
      · merged_cells_info   : 所有合并单元格详情（range, min/max row/col, value, row_span, col_span）
      · non_empty_cells     : 所有非空单元格（row, col, value）
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

你必须编写一个名为 `extract` 的函数，签名如下：

```python
def extract(ws, merged_map: dict) -> list:
    ...
    return result  # List[List[str]]，第0行为列名，后续为数据行
```

参数说明：
  ws         : openpyxl Worksheet 对象，已加载好，直接读取即可
  merged_map : 合并单元格填充图，键为 (row, col) 元组，值为该合并区域的实际值
               使用方式：val = merged_map.get((row, col), ws.cell(row, col).value)

返回值规范：
  - List[List[str]]，所有值转为字符串，None/空值转为 ""
  - 第 0 行为列名列表
  - 后续每行为数据行，列数与表头一致

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
重要规则：
  1. 只能使用标准库（re, json, math 等），不能 import openpyxl
  2. 合并单元格必须通过 merged_map 处理，不要直接相信 ws.cell().value（合并区域非左上角格为 None）
  3. 跨行合并列（如 SYSTEM MODULE）需要向前填充（forward-fill）
  4. 若有双行表头（父级跨列合并 + 子级），需正确拼合列名（如 "ANTENNAS||Antenna Qty."）
     或直接用子级列名（取决于是否有 target_columns 指定了 parent）
  5. 行方向子表、列方向子表、嵌套子表等非常规结构，都要正确处理
  6. 代码必须健壮：行列边界要用变量，不要硬编码"第16行开始"之类的数字
  7. 只返回代码，不要任何解释。代码包在 ```python ... ``` 中

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
代码模板（参考，根据实际结构调整）：

```python
def extract(ws, merged_map: dict) -> list:
    def cell_val(r, c):
        \"\"\"读取单元格，自动处理合并\"\"\"
        v = merged_map.get((r, c), ws.cell(row=r, column=c).value)
        if v is None:
            return ""
        if isinstance(v, float) and v == int(v):
            return str(int(v))
        return str(v).strip()

    # 1. 定位子表（根据 subtable_title 找到起止行列）
    subtable_start_row = ...
    subtable_end_row   = ...
    subtable_start_col = ...
    subtable_end_col   = ...
    header_row_count   = ...   # 1 或 2

    # 2. 读取表头（单/双行）
    header_row1 = [cell_val(subtable_start_row, c)
                   for c in range(subtable_start_col, subtable_end_col + 1)]
    if header_row_count == 2:
        header_row2 = [cell_val(subtable_start_row + 1, c)
                       for c in range(subtable_start_col, subtable_end_col + 1)]
        headers = [f"{p}||{c}" if p else c for p, c in zip(header_row1, header_row2)]
    else:
        headers = header_row1

    # 3. 读取数据行（含 forward-fill）
    data_start = subtable_start_row + header_row_count
    ffill_cols = set()   # 需要向前填充的列索引（相对于 subtable_start_col）
    prev = [""] * len(headers)
    rows = [headers]

    for r in range(data_start, subtable_end_row + 1):
        row = []
        for i, c in enumerate(range(subtable_start_col, subtable_end_col + 1)):
            v = cell_val(r, c)
            if i in ffill_cols and v == "":
                v = prev[i]
            else:
                prev[i] = v
            row.append(v)
        rows.append(row)

    return rows
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
    输入：sheet_structure + subtable_title + target_columns
    输出：generated_code（完整的 extract 函数字符串）
    """
    st             = state["sheet_structure"]
    errors         = state.get("errors", [])
    sandbox_error  = state.get("sandbox_error")
    target_columns = state.get("target_columns")

    # ── 构造给 LLM 的上下文 ──────────────────────────────────
    ctx: dict = {
        "subtable_title": state["config"]["subtable_title"],
        "target_columns": target_columns,   # 告知 LLM 用户只要哪些列
        "sheet_structure": {
            # 精简传输：只传必要字段，避免 token 爆炸
            "sheet_name":        st["sheet_name"],
            "max_row":           st["max_row"],
            "max_col":           st["max_col"],
            "subtables":         st["potential_subtables"],
            "potential_headers": st["potential_headers"][:150],
            "merged_cells_info": st["merged_cells_info"][:120],
            # non_empty_cells 太大，只传子表候选块附近的
            "sample_cells":      [
                c for c in st["non_empty_cells"]
                if any(
                    sub["start_row"] - 2 <= c["row"] <= sub["end_row"] + 2
                    for sub in st["potential_subtables"]
                    if state["config"]["subtable_title"].lower() in sub.get("title", "").lower()
                       or any(state["config"]["subtable_title"].lower() in t.lower()
                              for t in sub.get("title_candidates", []))
                )
            ][:200],
        },
    }

    if state.get("hints"):
        ctx["hints"] = state["hints"]

    # 重试时附上错误，让 LLM 针对性修正
    if sandbox_error:
        ctx["last_code_error"] = sandbox_error
        ctx["retry_instruction"] = (
            "上次生成的代码执行时报错，请仔细阅读错误信息，修正代码。"
            "常见问题：行列边界计算错误、合并单元格未用 merged_map 处理、"
            "表头行数判断错误。"
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
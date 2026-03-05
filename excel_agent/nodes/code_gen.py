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
  - sheet_structure     : Sheet 的完整结构
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

你必须编写一个名为 `extract` 的函数，签名如下：

```python
def extract(ws, merged_map: dict) -> list:
    ...
    return result  # Dict[str, List[List[str]]]
```

返回值规范：
  - 必须返回一个字典（dict）。
  - 字典的键为 subtable_titles 中的原名，值为该子表对应的 List[List[str]]。
  - 第 0 行为列名列表，后续为数据行。

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
重要规则：
  1. 只能使用标准库（re, json, math 等），不能 import openpyxl
  2. 【合并单元格处理】：必须通过 merged_map 读取
  3. 【空值处理策略】：必须保持原始空单元格结构，绝对禁止进行向前的逻辑填充（forward-fill）
  4. 【禁止搜索指令】：绝对禁止写 for 循环去搜索标题。你必须在代码中直接使用硬编码的具体数字！
  5. 【表头处理关键】：如果子表存在双行表头（例如第8行是父类，第9行是子类），必须同时读取两行，并用 "||" 拼接！如果是单行表头，只读一行。
  6. 必须为传入的每一个子表独立提取一份数据。如果表不存在，返回空列表 []。
  7. 只返回代码，不要任何解释。代码包在 ```python ... ``` 中

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
代码模板（参考）：

```python
def extract(ws, merged_map: dict) -> dict:
    def cell_val(r, c):
        v = merged_map.get((r, c), ws.cell(row=r, column=c).value)
        if v is None: return ""
        if isinstance(v, float) and v == int(v): return str(int(v))
        return str(v).replace('\\n', ' ').replace('\\r', '').strip()

    result = {}
    
    # ==== 提取表1 (硬编码坐标) ====
    start_col, end_col = 2, 15
    
    # 【情形A】如果是双行表头（例如第8行是父类，第9行是子类）：
    header_row1 = [cell_val(8, c) for c in range(start_col, end_col + 1)]
    header_row2 = [cell_val(9, c) for c in range(start_col, end_col + 1)]
    headers_1 = [f"{p}||{c}" if p and str(p) != str(c) else c for p, c in zip(header_row1, header_row2)]
    
    # 【情形B】如果是单行表头：
    # headers_1 = [cell_val(8, c) for c in range(start_col, end_col + 1)]

    rows_1 = [headers_1]
    
    # 读取数据行
    for r in range(10, 20):
        rows_1.append([cell_val(r, c) for c in range(start_col, end_col + 1)])
        
    result["Group 1"] = rows_1
    
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
    """高密度压缩测试"""
    st = state["sheet_structure"]
    errors = state.get("errors", [])
    sandbox_error = state.get("sandbox_error")

    config = state.get("config", {})
    subtable_titles = config.get("subtable_titles") or state.get("subtable_titles")

    target_columns = config.get("target_columns") or state.get("target_columns")
    hints = config.get("hints") or state.get("hints")

    # 极限压缩 Token 优化逻辑 _START_
    # 1. 圈定重点观察区域（只关注命中标题的子表附近的行）
    target_rows = set()
    matched_subtables = []

    for sub in st.get("potential_subtables", []):
        is_match = any(
            t.lower() in sub.get("title", "").lower() or
            any(t.lower() in tc.lower() for tc in sub.get("title_candidates", []))
            for t in subtable_titles
        )
        if is_match:
            matched_subtables.append(sub)
            # 扩大视野，包含子表上方 2 行（捕获大标题）和整个表格区域
            target_rows.update(range(max(1, sub["start_row"] - 2), sub["end_row"] + 2))

    # 2. 抛弃冗余的 JSON 键名 (row, col, value...)，改用高密度字符串格式
    # 格式如："R10C2:Sector 1"
    compressed_cells = [
        f"R{c['row']}C{c['col']}:{c['value']}"
        for c in st.get("non_empty_cells", [])
        if c["row"] in target_rows
    ][:200]  # 限定最大数量防止意外超载

    # 合并单元格信息也压缩为简短字符串："R1C1~R1C4:Title"
    compressed_merges = [
        f"R{m['min_row']}C{m['min_col']}~R{m['max_row']}C{m['max_col']}:{m['value']}"
        for m in st.get("merged_cells_info", [])
        if m["min_row"] in target_rows or m["max_row"] in target_rows
    ][:80]

    # 3. 组装极简上下文
    ctx: dict = {
        "subtable_titles": subtable_titles,
        "target_columns": target_columns,
        "sheet_structure": {
            "sheet_name": st["sheet_name"],
            "max_row": st["max_row"],
            "max_col": st["max_col"],
            "matched_subtables": matched_subtables,  # 只传命中的块
            "merged_cells_info": compressed_merges,  # 传入压缩后的字符串列表
            "sample_cells": compressed_cells  # 传入压缩后的字符串列表
        },
    }
    # 极限压缩 Token 优化逻辑 _END_

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

def code_gen_node_(state: AgentState) -> dict:
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
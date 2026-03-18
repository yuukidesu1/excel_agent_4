"""
nodes/code_gen.py — LLM 节点：看 parse 数据，写 openpyxl 抽取代码

LLM 的任务：
    1. 理解 sheet_structure（子表位置、合并单元格、表头结构）
    2. 基于 psa_hints 明确子表的 layout_type（纵表/横表/交叉表）和物理锚点。
    3. 根据 subtable_configs 定位目标行列
    4. 编写完整的 Python 抽取函数，使用 openpyxl 直接读取数据
    5. 函数必须返回 Dict[str, List[List[str]]]（标准二维数组，第0行为列名）
"""

import json
import re
import os
from pathlib import Path

from langchain_openai import ChatOpenAI
from langchain_core.messages import SystemMessage, HumanMessage

from excel_agent.state import AgentState


_SYSTEM_PROMPT = """\
你是顶尖的 Excel 数据抽取算法专家。我会给你一份 Excel Sheet 的完整结构描述以及前置分析器(PSA)给出的精准提示。
你需要编写一段 Python 代码，使用已加载好的 openpyxl Worksheet 对象抽取指定子表的数据。

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
你收到的核心上下文信息：
  - subtable_titles   : 目标子表的标题关键词列表
  - subtable_configs  : 抽取配置。指明了用户想要提取的 col_headers(列方向) 或 row_headers(行方向)
  - psa_hints         : 前置分析器提供的终极提示！包含该子表的 layout_type(布局)、start_row(起始行)和start_col(起始列)
  - sheet_structure   : 包含合并单元格、非空单元格的压缩视图
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

你必须编写一个名为 `extract` 的函数，签名如下：

```python
def extract(ws, merged_map: dict) -> dict:
    ...
    return result  # 必须返回 Dict[str, List[List[str]]]
```
返回值规范：
字典的键为 subtable_titles 中的原名，值为该子表对应的 List[List[str]]。
无论原表是哪种布局，返回的 List[List[str]] 中，第 0 行必须是字段名列表，后续为数据行。

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

 极其重要的布局处理规则 (基于 psa_hints["layout_type"])：
1. 【vertical (仅列/纵表)】：数据向下延伸。
    变体警告：子表的大标题可能在表格的【正上方】，也可能在表格的【最左侧】！
        若标题在上方：表头通常在标题行的下方。
        若标题在左侧（合并单元格）：表头通常与标题位于【同一行】，但在其【右侧】列！
    你需要仔细观察 sample_cells 的坐标关系，灵活定位真实的表头行，然后再向下提取数据。
2. 【horizontal (仅行/横表)】：表头在左侧同列，数据向右延伸！你必须通过行名定位行坐标，然后向右遍历读取！
3. 【cross (交叉表)】：既有行表头又有列表头！你需要分别定位行表头的列范围和列表头的行范围，提取交叉点的值。通常建议将其展平(Melt)为类似于 ["行维度", "列维度", "数值"] 的一维记录返回，或者按照 subtable_configs 要求返回。
4. 【合并单元格处理】：必须通过 merged_map.get((r, c), ws.cell(row=r, column=c).value) 读取！绝对不能直接访问 merged_cells 属性！
5. 【空值策略】：保持原始空单元格为空字符串 ""，绝对禁止使用 forward-fill 导致数据污染。
6. 【锚点利用】：绝对禁止写全表 for 循环去盲搜标题！你必须基于 psa_hints 提供的 start_row 和 start_col 圈定搜索和提取的边界。
7. 代码包在 python ...  中，不要输出额外解释！

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

代码模板(参考)：
def extract(ws, merged_map: dict) -> dict:
    def cell_val(r, c):
        v = merged_map.get((r, c), ws.cell(row=r, column=c).value)
        if v is None: return ""
        if isinstance(v, float) and v == int(v): return str(int(v))
        return str(v).replace('\\n', ' ').replace('\\r', '').strip()

    result = {}

    # ==== 示例 1: 处理 vertical (纵表) ====
    # 假设 psa_hints 给出 start_row=5, start_col=1
    start_row, start_col = 5, 1
    header_row = start_row + 1 # 动态调整
    data_col_start, data_col_end = start_col, start_col + 10

    headers_1 = [cell_val(header_row, c) for c in range(data_col_start, data_col_end + 1)]
    rows_1 = [headers_1]
    for r in range(header_row + 1, header_row + 20): 
        rows_1.append([cell_val(r, c) for c in range(data_col_start, data_col_end + 1)])
    result["Vertical Table"] = rows_1

    # ==== 示例 2: 处理 horizontal (横表) ====
    # 假设 psa_hints 给出 start_row=15, start_col=1
    h_start_row, h_start_col = 15, 1
    header_col = h_start_col 
    data_row_start, data_row_end = h_start_row, h_start_row + 5 
    data_col_start, data_col_end = header_col + 1, header_col + 8 

    headers_2 = [cell_val(r, header_col) for r in range(data_row_start, data_row_end + 1)]
    rows_2 = [headers_2]
    for c in range(data_col_start, data_col_end + 1):
        rows_2.append([cell_val(r, c) for r in range(data_row_start, data_row_end + 1)])
    result["Horizontal Table"] = rows_2

    # ==== 示例 3: 处理 cross (交叉表) ====
    # 假设 psa_hints 给出 start_row=30, start_col=1
    c_start_row, c_start_col = 30, 1
    col_headers_row = c_start_row       # 上方的列表头
    row_headers_col = c_start_col       # 左侧的行表头

    data_row_start, data_row_end = c_start_row + 1, c_start_row + 5
    data_col_start, data_col_end = c_start_col + 1, c_start_col + 5

    # 交叉表通常需要展平 (Unpivot / Melt) 输出
    headers_3 = ["Row_Dimension", "Col_Dimension", "Value"]
    rows_3 = [headers_3]
    for r in range(data_row_start, data_row_end + 1):
        row_title = cell_val(r, row_headers_col)
        for c in range(data_col_start, data_col_end + 1):
            col_title = cell_val(col_headers_row, c)
            val = cell_val(r, c)
            if val: # 可选：只提取有值的交叉点
                rows_3.append([row_title, col_title, val])
    result["Cross Table"] = rows_3

    return result
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
    model    = os.getenv("LLM_MODEL", "glm-4.7")
    if not api_key:
        raise ValueError("未找到 OPENAI_API_KEY，请检查 .env 文件。")
    return ChatOpenAI(model=model, temperature=0, api_key=api_key, base_url=base_url)


def _extract_code(raw: str) -> str:
    """从 LLM 输出中提取 python ...  之间的代码"""
    m = re.search(r"```python\s*([\s\S]*?)```", raw)
    if m:
        return m.group(1).strip()
    return raw.strip()


def code_gen_node(state: AgentState) -> dict:
    """高密度压缩并生成 LLM Prompt 节点"""
    st = state["sheet_structure"]
    errors = state.get("errors", [])
    sandbox_error = state.get("sandbox_error")

    config = state.get("config", {})

    # 1. 优先获取 missed_subtables
    cache_state = state.get("cache", {})
    subtable_titles = cache_state.get("missed_subtables")
    if subtable_titles is None:
        subtable_titles = config.get("subtable_titles", [])

    # 2. 兼容新旧配置
    subtable_configs = config.get("subtable_configs") or config.get("target_columns")
    hints = config.get("hints")

    # 3. ★ 核心提取：提取 PSA 识别出的物理锚点与布局类型
    psa_hints = {}
    for title in subtable_titles:
        entry = cache_state.get("entries", {}).get(title)
        if entry:
            psa_hints[title] = {
                "layout_type": entry.get("layout_type", "vertical"),
                "start_row": entry.get("start_row"),
                "start_col": entry.get("start_col"),
            }

    # 4. 极限压缩 Token 优化逻辑 _START_
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
            target_rows.update(range(max(1, sub["start_row"] - 2), sub["end_row"] + 2))

    compressed_cells = [
        f"R{c['row']}C{c['col']}:{c['value']}"
        for c in st.get("non_empty_cells", [])
        if c["row"] in target_rows
    ][:250]

    compressed_merges = [
        f"R{m['min_row']}C{m['min_col']}~R{m['max_row']}C{m['max_col']}:{m['value']}"
        for m in st.get("merged_cells_info", [])
        if m["min_row"] in target_rows or m["max_row"] in target_rows
    ][:100]

    # 5. 组装最终上下文
    ctx: dict = {
        "subtable_titles": subtable_titles,
        "subtable_configs": subtable_configs,
        "psa_hints": psa_hints,
        "sheet_structure": {
            "sheet_name": st["sheet_name"],
            "max_row": st["max_row"],
            "max_col": st["max_col"],
            "matched_subtables": matched_subtables,
            "merged_cells_info": compressed_merges,
            "sample_cells": compressed_cells
        },
    }
    # 极限压缩 Token 优化逻辑 _END_

    if hints:
        ctx["hints"] = hints

    # 重试时附上错误，让 LLM 针对性修正
    if sandbox_error:
        ctx["last_code_error"] = sandbox_error
        ctx["retry_instruction"] = (
            "🚨 上次生成的代码执行时出错，请仔细阅读上方错误信息修正代码！\n"
            "特别注意：请优先利用 psa_hints 提供的 start_row 和 start_col 进行坐标锚定，切忌写死绝对行号！"
        )
    elif errors:
        ctx["last_quality_errors"] = errors[-3:]
        ctx["retry_instruction"] = (
            "🚨 上次代码执行成功但质量不达标，请根据质量报错修正代码。\n"
            "常见错误：如果是 horizontal 横表，返回的数组中可能行/列发生了颠倒，请参考代码模板中按列遍历的逻辑。"
        )

    messages = [
        SystemMessage(content=_SYSTEM_PROMPT),
        HumanMessage(content=json.dumps(ctx, ensure_ascii=False, indent=2)),
    ]

    response = _get_llm().invoke(messages)
    code     = _extract_code(response.content)

    return {"generated_code": code, "sandbox_error": None}
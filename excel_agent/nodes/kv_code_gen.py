import json
import os
import re
from typing import Dict, Any, Optional
from pathlib import Path

from langchain_openai import ChatOpenAI
from langchain_core.messages import SystemMessage, HumanMessage

from excel_agent.state import AgentState

_SYSTEM_PROMPT = """\
你是顶尖的 Excel 表单数据处理专。当前任务时抽取离散排版的 Key-Value (KV) 键值对表单。

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
【输入上下文信息】
    - subtable_titles   : 目标子表名称列表
    - subtable_configs  : 抽取配置（包含需要的 keys， 通常在 row_headers 中）
    - psa_hints         : 提供子表的 start_row(起步行)、start_col(起步列)
    - sheet_structure   : 包含非空/合并单元格的坐标与值

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
【输出与红线规则】
你必须编写一个名为 `extract_kv(ws, merged_map: dict) -> dict` 的函数。
返回值必须是 Dict[str, List[List[str]]]， 键为子表名，值为仅包含两行的二维数组（第0行为 Key，第1行为 Value）。

核心红线规则：
1. 绝对坐标硬编码（极度重要）：你【绝对不能】在代码中写 for 循环去动态找 Key，你必须直接观察 sheet_structure，找到对应 Value 所在的精确单元格，并将这些 Value 的绝对坐标写死在代码里（如 `val_coords = [(10, 3), (10, 5)]`）。
2. KV 特征：在 KV 表单中，Key 和 Value 可能是左右相邻，也可能是上下相邻。你要观察上下文并分析结构，准确捕获 Value 的坐标。
3. 单元格读取：必须通过 `merged_map.get((r, c), ws.cell(row=r, column=c).value)` 读取。

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
【代码骨架
```Python
def extract(ws, merged_map: dict) -> dict:
    def cell_val(r, c):
        v = merged_map.get((r, c), ws.cell(row=r, column=c).value)
        if v is None: return ""
        if isinstance(v, float) and v == int(v) return str(int(v))
        return str(v).replace('\\n', ' ').replace('\\r', '').strip()
    result = {}
    
    # ==== 示例： 处理 KV 表单 ====
    headers_1 = ["Capacity", "1+0/1+1", "XPIC"]
    # 观察上下文，找到对应 Value 的【绝对坐标】并写死
    # 假设 Capacity 的值在 R10C3，1+0的值在 R10C5，XPIC 的值在 R11C3
    val_coords = [
        (10, 3),
        (10, 5),
        (11, 3)
    ]
    
    row_data = []
    for r, c in val_coords:
        row_data.append(cell_val(r, c))
    
    result["2. MW RTN(IDU) Information"] = [headers_1, row_data]
    
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
    api_key = os.getenv("OPENAI_API_KEY")
    base_url = os.getenv("OPENAI_BASE_URL")
    model = os.getenv("LLM_MODEL", "GLM-4.7")
    if not api_key:
        raise ValueError("OpenAI API key is required")
    return ChatOpenAI(model=model, temperature=0, api_key=api_key, base_url=base_url)

def _extract_code(raw: str) -> str:
    m = re.search(r"```python\s*([\s\S]*?)```", raw)
    if m:
        return m.group(1).strip()
    return raw.strip()


def _compress_structure_for_kv(st: dict, kv_titles: list, psa_hints: dict) -> dict:
    """
    为 KV 表抽取压缩 sheet_structure，只保留相关区域的单元格信息
    """
    # KV 表通常需要标题周围一个小范围的数据
    # 计算每个 KV 表的相关行范围 (标题行 ±20 行)
    target_rows = set()
    target_cols = set()

    for title in kv_titles:
        hint = psa_hints.get(title, {})
        start_row = hint.get("start_row", 1)
        start_col = hint.get("start_col", 1)

        # KV 表通常比较紧凑，搜索标题下方 25 行，左右各 15 列的范围
        for r in range(max(1, start_row - 2), min(start_row + 25, st.get("max_row", 1000) + 1)):
            target_rows.add(r)
        for c in range(max(1, start_col - 2), min(start_col + 15, st.get("max_col", 50) + 1)):
            target_cols.add(c)

    # 压缩单元格表示：R{row}C{col}:{value}
    compressed_cells = [
        f"R{c['row']}C{c['col']}:{c['value']}"
        for c in st.get("non_empty_cells", [])
        if c["row"] in target_rows and c["col"] in target_cols
    ][:300]  # 最多 300 个单元格

    # 压缩合并单元格表示：R{min_row}C{min_col}~R{max_row}C{max_col}:{value}
    compressed_merges = [
        f"R{m['min_row']}C{m['min_col']}~R{m['max_row']}C{m['max_col']}:{m['value']}"
        for m in st.get("merged_cells_info", [])
        if (m["min_row"] in target_rows or m["max_row"] in target_rows) and
           (m["min_col"] in target_cols or m["max_col"] in target_cols)
    ][:50]  # 最多 50 个合并单元格

    return {
        "sheet_name": st.get("sheet_name", ""),
        "max_row": st.get("max_row", 1000),
        "max_col": st.get("max_col", 50),
        "sample_cells": compressed_cells,
        "merged_cells_info": compressed_merges
    }


def kv_code_gen_node(state: AgentState) -> dict:
    missed = state.get("cache", {}).get("missed_subtables", [])
    if not missed:
        return {"generated_code": []}
    # 只筛选 layout_type == "kv" 的表单
    psa_hints = {}
    valid_titles = []
    for title in missed:
        entry = state["cache"]["entries"][title[0]]
        if entry["layout_type"] == "kv":
            psa_hints[title[0]] = {
                "layout_type": "kv",
                "start_row": entry["start_row"],
                "start_col": entry["start_col"]
            }
            valid_titles.append(title[0])

    if not valid_titles:
        return {"generated_code": []}

    # ★ 极限压缩 Token 优化逻辑
    st = state["sheet_structure"]
    compressed_structure = _compress_structure_for_kv(st, valid_titles, psa_hints)

    context = {
        "subtable_titles": valid_titles,
        "subtable_configs": {t: state["config"]["subtable_configs"].get(t) for t in valid_titles},
        "psa_hints": psa_hints,
        "sheet_structure": compressed_structure  # 使用压缩后的结构
    }

    human_msg = f"请根据一下上下文编写 extract_kv 代码： \n```json\n{json.dumps(context, ensure_ascii=False, indent=2)}\n```"
    messages = [SystemMessage(content=_SYSTEM_PROMPT), HumanMessage(content=human_msg)]

    response = _get_llm().invoke(messages)
    kv_code = _extract_code(response.content)

    # —————————————————— DEBUG ——————————————————————————
    # 使用原始字符串确保 '\n' 被正确转义为两个字符而不是换行符
    # kv_code = r"""def extract(ws, merged_map: dict) -> dict:
    #     def cell_val(r, c):
    #         v = merged_map.get((r, c), ws.cell(row=r, column=c).value)
    #         if v is None: return ""
    #         if isinstance(v, float) and v == int(v): return str(int(v))
    #         return str(v).replace('\n', ' ').replace('\r', '').strip()
    #
    #     result = {}
    #
    #     # ==== 处理子表：Spcae Available for New RF Antenna(m) ====
    #     # 根据上下文分析，Keys 为 "Leg 1", "Leg 2", "Leg 3", "Leg 4"
    #     headers_1 = ["Leg 1", "Leg 2", "Leg 3", "Leg 4"]
    #
    #     # 观察 sheet_structure 中的 Row 35:
    #     # "Leg 1" 在 C35 (row 35, col 3), 其值 "32" 在 D35 (row 35, col 4)
    #     # "Leg 2" 在 E35 (row 35, col 5), 其值 "22" 在 F35 (row 35, col 6)
    #     # "Leg 3" 在 G35 (row 35, col 7), 其值 "22" 在 H35 (row 35, col 8)
    #     # "Leg 4" 在 I35 (row 35, col 9), 其值 "NA" 在 J35 (row 35, col 10)
    #     # 硬编码对应的 Value 坐标
    #     val_coords = [
    #         (35, 4), # D35
    #         (35, 6), # F35
    #         (35, 8), # H35
    #         (35, 10) # J35
    #     ]
    #
    #     row_data = []
    #     for r, c in val_coords:
    #         row_data.append(cell_val(r, c))
    #
    #     result["Spcae Available for New RF Antenna(m)"] = [headers_1, row_data]
    #
    #     return result"""

    return {"generated_code": [kv_code], "sandbox_error": None}
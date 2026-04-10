"""
nodes/code_gen.py — LLM 节点：看 parse 数据，写 openpyxl 抽取代码

LLM 的任务：
    1. 理解 sheet_structure（子表位置、合并单元格、表头结构）
    2. 基于 psa_hints 明确子表的 layout_type（纵表/横表/交叉表）和物理锚点。
    3. 根据 subtable_configs 定位目标行列
    4. 编写完整的 Python 抽取函数，使用 openpyxl 直接读取数据
    5. 函数必须返回 Dict[str, List[List[str]]]（标准二维数组，第 0 行为列名）
"""

import json
import re
import os
from pathlib import Path

from langchain_openai import ChatOpenAI
from langchain_core.messages import SystemMessage, HumanMessage

from excel_agent.state import AgentState

_SYSTEM_PROMPT = """\
你是顶尖的 Excel 数据抽取专家。请根据传入的结构视图和前置分析器 (PSA) 提示，编写 Python 代码抽取数据。

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
【输入上下文信息】
  - subtable_titles   : 目标子表名称列表
  - subtable_configs  : 抽取配置（包含需要的 col_headers 或 row_headers）
  - psa_hints         : 提供子表的 layout_type(布局)、start_row(起步行)、start_col(起步列)
  - sheet_structure   : 包含非空/合并单元格的坐标与值

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
【输出与红线规则】
你必须编写一个名为 `extract(ws, merged_map: dict) -> dict` 的函数。
返回值必须是 Dict[str, List[List[str]]]，键为子表名，值为二维数组。

核心红线规则（违反将导致系统崩溃）：
1. 提取策略降维：无论是 vertical(纵表)、horizontal(横表) 还是 cross(交叉表)，你的任务仅仅是把它们当作普通的二维网格提取出来。第 0 行为列名，后续为数据行。不要去执行展平 (Melt) 等复杂操作。
2. 绝对坐标硬编码（极度重要）：你【绝对不能】在代码中调用上下文变量名（会报 NameError）。你必须直接观察上下文，将具体的起始行号、提取列号【写死】在代码里（例如 `cols = [2, 3, 5]`）。
3. 动态终止探针：绝对不能写死结束行（如 `while r < 20`），必须使用 `while r <= ws.max_row`，并通过探测主键列为空来 `break`。
4. 单元格读取：必须通过 `merged_map.get((r, c), ws.cell(row=r, column=c).value)` 读取。
5. 多级表头处理：如果配置中包含 `||` 符号的多级表头（如 `"RF MODULE||TYPE"`），
   你必须使用 `_h(r, c, depth)` 函数从 Excel 中读取垂直堆叠的表头单元格并拼接。

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
【代码骨架模板 (请严格参考此范式)】

```python
def extract(ws, merged_map: dict) -> dict:
    def cell_val(r, c):
        v = merged_map.get((r, c), ws.cell(row=r, column=c).value)
        if v is None: return ""
        if isinstance(v, float) and v == int(v): return str(int(v))
        return str(v).replace('\\n', ' ').replace('\\r', '').strip()

    def _h(r, c, depth=2):
        '''构建多级表头键 (如 '父||子')。
        从第 r 行开始向下读取 depth 行，拼接非空单元格值。
        用于处理 Excel 中垂直堆叠的多级表头结构。
        '''
        parts = []
        for i in range(depth):
            v = merged_map.get((r + i, c), ws.cell(row=r + i, column=c).value)
            if v:
                parts.append(str(v).replace('\\n', ' ').replace('\\r', '').strip())
        return "||".join(parts) if parts else ""

    result = {}

    # ==== 示例 1: 纵表 (或交叉表) 的提取范式 ====
    # 【不要写循环找列！直接从上下文中观察出需要的绝对列号并写死】
    # 【重要：如果配置中包含 "父||子" 形式的多级表头，必须使用 _h() 函数构建！】
    # 观察 Excel 表头位置：第 5 行是表头起始行
    header_row = 5
    # 使用 _h() 读取多级表头 (如 "RF MODULE||TYPE")
    headers_1 = [_h(header_row, c, depth=2) for c in [2, 3, 4, 5, 9, 10]]
    # 如果表头是单级的，也可以直接写死：headers_1 = ["Header A", "Header B"]

    data_cols = [2, 3, 4, 5, 9, 10]  # 观察上下文后写死的绝对列号
    data_start_row = 6              # 观察上下文后写死的数据起始行

    table_1 = [headers_1]
    r = data_start_row
    while r <= ws.max_row:
        # 使用关键列 (如第 2 列) 作为探针，若为空或遇到下一个表头则终止
        if not cell_val(r, 2):
            break
        table_1.append([cell_val(r, c) for c in data_cols])
        r += 1
    result["Table 1 Name"] = table_1

    # ==== 示例 2: 横表 ("仅行"布局) 的提取范式 ====
    # 布局特征：表头在左侧纵向排列（第 1 列），数据向右延伸
    # 提取策略：将横表当作普通二维网格提取，第 0 行为表头，后续每行为数据
    #
    # 示例 Excel 结构：
    #   R10C1="Header A"  R10C2="Val1"  R10C3="Val2" ...
    #   R11C1="Header B"  R11C2="Data1" R11C3="Data2" ...
    #
    # 正确提取结果（返回格式）：
    #   [
    #     ["Header A", "Header B"],      # 第 0 行：表头（从第 1 列读取）
    #     ["Val1", "Data1"],             # 第 1 行：数据列 1
    #     ["Val2", "Data2"],             # 第 2 行：数据列 2
    #     ...
    #   ]
    #
    # 代码范式：
    header_col = 3    # 表头所在的列（绝对坐标）
    data_rows = [10, 11]     # 表头所在的行号列表（绝对坐标）
    data_start_col = 5       # 数据起始列（绝对坐标）

    # 从第 1 列（或指定列）垂直读取表头
    headers_2 = [cell_val(r, header_col) for r in data_rows]

    table_2 = [headers_2]
    c = data_start_col
    while c <= ws.max_column:
        # 探针：检查第一行表头在当前列是否有值
        if not cell_val(data_rows[0], c):
            break
        # 提取当前列的所有行数据，作为结果的一行
        table_2.append([cell_val(r, c) for r in data_rows])
        c += 1
    result["Table 2 Name"] = table_2

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
    # subtable_titles = cache_state.get("missed_subtables")
    subtable_titles = [ms[0] for ms in cache_state.get("missed_subtables") if ms[1] != "kv_table"]
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
                "header_map": entry.get("header_map", {}),  # ← 新增：PSA 识别的表头坐标信息
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
            "上次生成的代码执行时出错，请仔细阅读上方错误信息修正代码！"
        )
    elif errors:
        ctx["last_quality_errors"] = errors[-3:]
        ctx["retry_instruction"] = (
            "上次代码执行成功但质量不达标，请根据质量报错修正代码。\n"
            "常见错误：如果是 horizontal 横表，返回的数组中可能行/列发生了颠倒，请参考代码模板中按列遍历的逻辑。"
        )

    messages = [
        SystemMessage(content=_SYSTEM_PROMPT),
        HumanMessage(content=json.dumps(ctx, ensure_ascii=False, indent=2)),
    ]

    response = _get_llm().invoke(messages)
    code     = _extract_code(response.content)

    # ———————————————————————— DEBUG ——————————————————————————

    # CONFIGURATION.yaml
    # code = "def extract(ws, merged_map: dict) -> dict:\n    def cell_val(r, c):\n        v = merged_map.get((r, c), ws.cell(row=r, column=c).value)\n        if v is None: return \"\"\n        if isinstance(v, float) and v == int(v): return str(int(v))\n        return str(v).replace('\\n', ' ').replace('\\r', '').strip()\n\n    result = {}\n\n    headers_2g = [\"SYSTEM MODULE\", \"CELL\", \"RF MODULE||TYPE\", \"RF MODULE||QTY.\", \"ANTENNAS||NEW/SWAP/EXIST\", \"ANTENNAS||Antenna Type\"]\n    cols_2g = [2, 3, 4, 5, 9, 10]\n    data_start_row_2g = 6\n\n    table_2g = [headers_2g]\n    r = data_start_row_2g\n    while r <= ws.max_row:\n        if not cell_val(r, 2):\n            break\n        table_2g.append([cell_val(r, c) for c in cols_2g])\n        r += 1\n    result[\"2G Configuration\"] = table_2g\n\n    headers_3g = [\"SYSTEM MODULE\", \"CELL\", \"RF MODULE||TYPE\", \"RF MODULE||QTY.\", \"ANTENNAS||NEW/SWAP/EXIST\", \"ANTENNAS||Antenna Type\"]\n    cols_3g = [2, 3, 4, 5, 9, 10]\n    data_start_row_3g = 17\n\n    table_3g = [headers_3g]\n    r = data_start_row_3g\n    while r <= ws.max_row:\n        if not cell_val(r, 2):\n            break\n        table_3g.append([cell_val(r, c) for c in cols_3g])\n        r += 1\n    result[\"3G Configuration\"] = table_3g\n\n    headers_4g = [\"RF MODULE||QTY.\", \"ANTENNAS||NEW/SWAP/EXIST\", \"ANTENNAS||Antenna Type\"]\n    cols_4g = [5, 9, 10]\n    data_start_row_4g = 28\n\n    table_4g = [headers_4g]\n    r = data_start_row_4g\n    while r <= ws.max_row:\n        if not cell_val(r, 3):\n            break\n        table_4g.append([cell_val(r, c) for c in cols_4g])\n        r += 1\n    result[\"4G Configuration\"] = table_4g\n\n    headers_5g = [\"SYSTEM MODULE\", \"CELL\", \"RF MODULE||TYPE\"]\n    cols_5g = [2, 3, 4]\n    data_start_row_5g = 43\n\n    table_5g = [headers_5g]\n    r = data_start_row_5g\n    while r <= ws.max_row:\n        if not cell_val(r, 3):\n            break\n        table_5g.append([cell_val(r, c) for c in cols_5g])\n        r += 1\n    result[\"5G Configuration\"] = table_5g\n\n    return result"
    #
    # TSSR_senario_TEST.yaml
    # code = "def extract(ws, merged_map: dict) -> dict:\n    def cell_val(r, c):\n        v = merged_map.get((r, c), ws.cell(row=r, column=c).value)\n        if v is None: return \"\"\n        if isinstance(v, float) and v == int(v): return str(int(v))\n        return str(v).replace('\\n', ' ').replace('\\r', '').strip()\n\n    def _h(r, c, depth=2):\n        parts = []\n        for i in range(depth):\n            v = merged_map.get((r + i, c), ws.cell(row=r + i, column=c).value)\n            if v:\n                parts.append(str(v).replace('\\n', ' ').replace('\\r', '').strip())\n        return '||'.join(parts) if parts else ''\n\n    result = {}\n\n    # 2G 1800 Mhz Existing Con./Mevcut Kon. - 表头在第 17 行，使用 _h 构建多级表头\n    t1_header_row = 17\n    t1_headers = [_h(t1_header_row, c, depth=2) for c in [3, 4, 6, 9, 10, 12, 17, 20, 37]]\n    t1_cols = [3, 4, 6, 9, 10, 12, 17, 20, 37]\n    t1_data = [t1_headers]\n    r = 19\n    while r <= ws.max_row:\n        if not cell_val(r, 3): break\n        t1_data.append([cell_val(r, c) for c in t1_cols])\n        r += 1\n    result[\"2G 1800 Mhz Existing Con./Mevcut Kon.\"] = t1_data\n\n    # 3G 2100 Mhz Existing Con./Mevcut Kon. - 表头在第 26 行\n    t2_header_row = 26\n    t2_headers = [_h(t2_header_row, c, depth=2) for c in [3, 4, 6, 9, 10, 11, 12, 13, 16, 17, 20, 37]]\n    t2_cols = [3, 4, 6, 9, 10, 11, 12, 13, 16, 17, 20, 37]\n    t2_data = [t2_headers]\n    r = 28\n    while r <= ws.max_row:\n        if not cell_val(r, 3): break\n        t2_data.append([cell_val(r, c) for c in t2_cols])\n        r += 1\n    result[\"3G 2100 Mhz Existing Con./Mevcut Kon.\"] = t2_data\n\n    # L2600 Existing Con./Mevcut Kon. - 表头在第 71 行\n    t3_header_row = 71\n    t3_headers = [_h(t3_header_row, c, depth=2) for c in [3, 4, 6, 9, 10, 12, 17, 20, 22, 37]]\n    t3_cols = [3, 4, 6, 9, 10, 12, 17, 20, 22, 37]\n    t3_data = [t3_headers]\n    r = 73\n    while r <= ws.max_row:\n        if not cell_val(r, 3): break\n        t3_data.append([cell_val(r, c) for c in t3_cols])\n        r += 1\n    result[\"L2600 Existing Con./Mevcut Kon.\"] = t3_data\n\n    # 3G 2100 Mhz Required Con./İstenen Kon. - 表头在第 103 行\n    t4_header_row = 103\n    t4_headers = [_h(t4_header_row, c, depth=2) for c in [3, 4, 6, 9, 10, 12, 17, 20, 37]]\n    t4_cols = [3, 4, 6, 9, 10, 12, 17, 20, 37]\n    t4_data = [t4_headers]\n    r = 105\n    while r <= ws.max_row:\n        if not cell_val(r, 3): break\n        t4_data.append([cell_val(r, c) for c in t4_cols])\n        r += 1\n    result[\"3G 2100 Mhz Required Con./İstenen Kon.\"] = t4_data\n\n    # L2600 Required Con./İstenen Kon. - 表头在第 148 行\n    t5_header_row = 148\n    t5_headers = [_h(t5_header_row, c, depth=2) for c in [3, 4, 6, 9, 10, 11, 12, 13, 16, 17, 20, 37]]\n    t5_cols = [3, 4, 6, 9, 10, 11, 12, 13, 16, 17, 20, 37]\n    t5_data = [t5_headers]\n    r = 150\n    while r <= ws.max_row:\n        if not cell_val(r, 3): break\n        t5_data.append([cell_val(r, c) for c in t5_cols])\n        r += 1\n    result[\"L2600 Required Con./İstenen Kon.\"] = t5_data\n\n    return result"
    #
    # test_horizontal.yaml
    # code = "def extract(ws, merged_map: dict) -> dict:\n    def cell_val(r, c):\n        v = merged_map.get((r, c), ws.cell(row=r, column=c).value)\n        if v is None: return \"\"\n        if isinstance(v, float) and v == int(v): return str(int(v))\n        return str(v).replace('\\n', ' ').replace('\\r', '').strip()\n\n    def _h(r, c, depth=2):\n        '''构建多级表头键 (如 '父||子')。\n        从第 r 行开始向下读取 depth 行，拼接非空单元格值。\n        用于处理 Excel 中垂直堆叠的多级表头结构。\n        '''\n        parts = []\n        for i in range(depth):\n            v = merged_map.get((r + i, c), ws.cell(row=r + i, column=c).value)\n            if v:\n                parts.append(str(v).replace('\\n', ' ').replace('\\r', '').strip())\n        return \"||\".join(parts) if parts else \"\"\n\n    result = {}\n\n    data_rows = [4, 5, 6, 7, 8]\n    data_start_col = 3\n\n    headers = [_h(r, 2, depth=1) for r in data_rows]\n\n    table_data = [headers]\n\n    c = data_start_col\n    while c <= ws.max_column:\n        if not cell_val(4, c):\n            break\n        col_data = [cell_val(r, c) for r in data_rows]\n        table_data.append(col_data)\n        c += 1\n\n    result[\"Server Node Configuration\"] = table_data\n\n    return result"
    #
    # TEST_WL_56A0DS6_mix.yaml
    # 使用原始字符串确保 '\n' 被正确转义为两个字符而不是换行符
#     code = """def extract(ws, merged_map: dict) -> dict:
# def extract(ws, merged_map: dict) -> dict:
#     def cell_val(r, c):
#         v = merged_map.get((r, c), ws.cell(row=r, column=c).value)
#         if v is None: return ""
#         if isinstance(v, float) and v == int(v): return str(int(v))
#         return str(v).replace('\n', ' ').replace('\r', '').strip()
#
#     def _h(r, c, depth=2):
#         '''构建多级表头键 (如 '父||子')。
#         从第 r 行开始向下读取 depth 行，拼接非空单元格值。
#         用于处理 Excel 中垂直堆叠的多级表头结构。
#         '''
#         parts = []
#         for i in range(depth):
#             v = merged_map.get((r + i, c), ws.cell(row=r + i, column=c).value)
#             if v:
#                 parts.append(str(v).replace('\n', ' ').replace('\r', '').strip())
#         return "||".join(parts) if parts else ""
#
#     result = {}
#
#     # ==== 1. DCDU 14B Load Information ====
#     # 布局特征：Horizontal (仅行布局)
#     # 结构分析：
#     # - 第 1 列为行表头列 (Row Headers)，包含 "DCDU 14B Load Information" 和 "Fuse Capacity"
#     # - 第 3 行和第 4 行为数据行所在行
#     # - 数据从第 2 列开始向右延伸 (Load 0 ~ Load 9, Location, Distance...)
#
#     table_name = "1. DCDU 14B Load Information"
#
#     # 绝对坐标配置
#     header_col = 1  # 行表头所在的列
#     data_rows = [3, 4]  # 行表头所在的行号 (DCDU 14B... 在 R3, Fuse Capacity 在 R4)
#     data_start_col = 2  # 数据起始列 (紧接行表头列之后)
#
#     # 第 0 行：构建表头 (从第 1 列垂直读取)
#     headers = [cell_val(r, header_col) for r in data_rows]
#
#     table_data = [headers]
#
#     # 循环提取每一列的数据，直到表头行 (R3) 为空
#     c = data_start_col
#     while c <= ws.max_column:
#         # 探针：检查第一行表头 (data_rows[0], 即 R3) 在当前列是否有值
#         if not cell_val(data_rows[0], c):
#             break
#
#         # 提取当前列的所有行数据，作为结果的一行
#         row_data = [cell_val(r, c) for r in data_rows]
#         table_data.append(row_data)
#         c += 1
#
#     result[table_name] = table_data
#
#     return result
# """

    return {"generated_code": [code], "sandbox_error": None}

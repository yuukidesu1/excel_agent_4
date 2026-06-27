import json
import re
import os

import httpx
from pathlib import Path

from langchain_openai import ChatOpenAI
from langchain_core.messages import SystemMessage, HumanMessage

from excel_agent.state import AgentState

_SYSTEM_PROMPT = """\
你是顶尖的 Excel 数据抽取专家。请根据传入的结构视图和前置分析器 (PSA) 提示，编写 Python 代码抽取指定子表的数据。

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
【输入上下文信息】
  - subtable_titles   : 当前需要抽取的子表名称列表。
  - subtable_configs  : 抽取配置（包含需要提取的 headers 列表）。
  - psa_hints         : 提供子表的 layout_type(布局)、start_row(起步行)、start_col(起步列) 等物理锚点。
  - sheet_structure   : 包含非空单元格、合并单元格坐标与值的压缩视图（供你观察结构）。

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
【输出与红线规则】
你必须编写一个名为 `extract(ws, merged_map: dict) -> dict` 的函数。
返回值必须是 Dict[str, List[List[str]]]，键为子表名，值为标准的二维数组。

核心红线规则（违反将导致系统崩溃）：
1. 提取策略降维：无论是 vertical(纵表)、horizontal(横表) 还是 cross(交叉表)，你的任务仅仅是将它们作为“普通的二维网格”提取出来。第 0 行为表头，后续为数据。不要执行展平 (Melt) 等操作。
2. 坐标数值化（极度重要）：你【绝对不能】在生成的代码中调用外部上下文未定义的变量（会报 NameError）。你必须直接观察传入的 `sheet_structure`，推断出具体的起始行号、列号，并将它们【写死为纯数字】（例如 `data_cols = [2, 3, 5]`）。
3. 防越界终止探针（防止连续表格“一读到底”）：绝对不能写死结束行，必须使用 `while` 循环向下或向右扫描。为了防止多个子表紧密相连导致越界抓取，【必须使用双重防越界探针】：
   - 探针 A（判空）：当探针列/行的值为空时，必须 `break`。
   - 探针 B（语义截断）：当探针列/行的值等于【其他子表的标题】或【特定边界词（如 Total, Note, 备注）】时，必须立即 `break`。
4. 单元格读取安全：必须通过 `merged_map.get((r, c), ws.cell(row=r, column=c).value)` 读取单元格。
5. 多级表头处理：如果目标配置中包含 `||` 符号的多级表头（如 `"RF MODULE||TYPE"`），你必须使用内置的 `_h(r, c, depth)` 函数，从 Excel 中读取垂直堆叠的表头单元格并拼接。
6. 横表转置规则（针对 horizontal 布局）：当提取 horizontal 横向布局的表格时，严禁按行将数据打包。你必须以列为单位进行 while c <= ws.max_column 扫描，将每一列对应的多个属性值打包成一个 List 插入表中，从而在逻辑上完成行列转置。例如表头是 [PropertyA, PropertyB] ，则每一行的数据必须是 [ValueA, ValueB] ，绝不能是 [ValueA, ValueB, ...] 。

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
【代码骨架模板 (请严格遵循此范式)】

```python
def extract(ws, merged_map: dict) -> dict:
    def cell_val(r, c):
        v = merged_map.get((r, c), ws.cell(row=r, column=c).value)
        if v is None:
            return ""
        if isinstance(v, float) and v == int(v):
            v = str(int(v))
        return " ".join(str(v).splitlines()).strip()

    def _h(r, c, depth=2):
        '''构建多级表头键 (如 '父||子')。
        从第 r 行开始向下读取 depth 行，拼接非空单元格值。
        '''
        parts = []
        for i in range(depth):
            v = merged_map.get((r + i, c), ws.cell(row=r + i, column=c).value)
            if v:
                parts.append(str(v).replace('\\n', ' ').replace('\\r', '').strip())
        return "||".join(parts) if parts else ""

    result = {}

    # ==== 示例 1: 纵表 ("vertical"布局) 的提取范式 ====
    header_row = 5
    headers_1 = [_h(header_row, c, depth=2) for c in [2, 3, 4, 5, 9, 10]]

    data_cols = [2, 3, 4, 5, 9, 10]
    data_start_row = 6

    # 重点：定义防越界截断词（从上下文中观察到的下一个表格标题、或其他终止特征）
    stop_words_1 = ["Table 2 Name", "Total", "Notes", "Remark", "Summary"]

    table_1 = [headers_1]
    r = data_start_row
    while r <= ws.max_row:
        probe_val = cell_val(r, 2) # 通常选择主键列（如第2列）作为探针

        # 双重探针：为空，或触碰到截断词，立即终止当前表格的提取
        if not probe_val or any(sw.lower() in probe_val.lower() for sw in stop_words_1):
            break

        table_1.append([cell_val(r, c) for c in data_cols])
        r += 1

    result["Table 1 Name"] = table_1


    # ==== 示例 2: 横表 ("horizontal"布局) 的提取范式 ====
    header_col = 1
    data_rows = [10, 11]
    data_start_col = 2

    headers_2 = [cell_val(r, header_col) for r in data_rows]
    table_2 = [headers_2]

    # 横表同样需要防越界截断词（向右扫描时可能遇到的其他模块标题或无关文字）
    stop_words_2 = ["Next Section", "End"]

    c = data_start_col
    while c <= ws.max_column:
        probe_val = cell_val(data_rows[0], c) # 通常选择第一行作为探针

        # ★ 双重探针限制
        if not probe_val or any(sw.lower() in probe_val.lower() for sw in stop_words_2):
            break

        table_2.append([cell_val(r, c) for r in data_rows])
        c += 1

    result["Table 2 Name"] = table_2

    return result
```
"""

def _get_llm(dynamic_token: str = None) -> ChatOpenAI:
    from dotenv import load_dotenv
    os.environ["http_proxy"] = ""
    os.environ["https_proxy"] = ""
    os.environ["all_proxy"] = ""
    for parent in Path(__file__).resolve().parents:
        env_file = parent / ".env"
        if env_file.exists():
            load_dotenv(env_file, override=True)
            break
    api_key  = os.getenv("OPENAI_API_KEY")
    base_url = os.getenv("OPENAI_BASE_URL") or None
    model    = os.getenv("LLM_MODEL", "glm-4.7")
    header_name = os.getenv("CUSTOM_HEADER_NAME")
    custom_headers = {}

    if header_name and dynamic_token:
        custom_headers[header_name] = dynamic_token

    if not api_key:
        raise ValueError("未找到 OPENAI_API_KEY，请检查 .env 文件。")

    http_client = httpx.Client(verify=False, timeout=600)

    return ChatOpenAI(model=model, temperature=0, api_key=api_key, base_url=base_url, default_headers=custom_headers, http_client=http_client)


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

    cache_state = state.get("cache", {})
    subtable_titles = [ms[0] for ms in cache_state.get("missed_subtables")]
    if subtable_titles is None:
        subtable_titles = config.get("subtable_titles", [])

    subtable_configs = config.get("subtable_configs") or config.get("target_columns")
    hints = config.get("hints")

    psa_hints = {}
    for title in subtable_titles:
        entry = cache_state.get("entries", {}).get(title)
        if entry:
            psa_hints[title] = {
                "layout_type": entry.get("layout_type", "vertical"),
                "start_row": entry.get("start_row"),
                "start_col": entry.get("start_col"),
                "header_map": entry.get("header_map", {}),
            }

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
    if hints:
        ctx["hints"] = hints

    if sandbox_error:
        ctx["last_code_error"] = sandbox_error
        ctx["retry_instruction"] = (
            "上次生成的代码执行时出错，请仔细阅读上方错误信息修正代码！\n"
        )
    elif errors:
        ctx["last_quality_errors"] = errors[-3:]
        ctx["retry_instruction"] = (
            ""
            "-上次代码执行成功但质量不达标，请根据质量报错修正代码。\n"
            "常见错误：如果是 horizontal 横表，返回的数组中可能行/列发生了颠倒，请参考代码模板中按列遍历的逻辑。"
        )

    messages = [
        SystemMessage(content=_SYSTEM_PROMPT),
        HumanMessage(content=json.dumps(ctx, ensure_ascii=False, indent=2)),
    ]

    response = _get_llm().invoke(messages)
    code     = _extract_code(response.content)


    return {"generated_code": [code], "sandbox_error": None}

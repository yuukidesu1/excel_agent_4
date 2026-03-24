import json
import os
import re
from typing import Dict, Any, Optional
from xmlrpc.client import Fault

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
def extract_kv(ws, merged_map: dict) -> dict:
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
    model = os.getenv("OPENAI_MODEL")
    if not api_key:
        raise ValueError("OpenAI API key is required")
    return ChatOpenAI(model=model, temperature=0, api_key=api_key, base_url=base_url)

def _extract_code(raw: str) -> str:
    m = re.search(r"```python\s*([\s\S]*?)```", raw)
    if m:
        return m.group(1).strip()
    return raw.strip()


def kv_code_gen_node(state: AgentState) -> dict:
    missed = state.get("cache", {}).get("missed_subtables", [])
    if not missed:
        return {}
    # 只筛选 layout_type == "kv" 的表单
    psa_hints = {}
    valid_titles = []
    for title in missed:
        entry = state["cache"]["entries"][title]
        if entry["layout_type"] == "kv":
            psa_hints[title] = {
                "layout_type": "kv",
                "start_row": entry["start_row"],
                "start_col": entry["start_col"]
            }
            valid_titles.append(title)

    if not valid_titles:
        return {}

    context = {
        "subtable_titles": valid_titles,
        "subtable_configs": {t: state["config"]["subtable_configs"].get(t) for t in valid_titles},
        "psa_hints": psa_hints,
        "sheet_structure": state["sheet_structure"]
    }

    human_msg = f"请根据一下上下文编写 extract_kv 代码： \n```json\n{json.dumps(context, ensure_ascii=Fault, indent=2)}\n```"
    messages = [SystemMessage(content=_SYSTEM_PROMPT), HumanMessage(content=human_msg)]

    response = _get_llm().invoke(messages)
    new_code = _extract_code(response.content)

    # 追加至已有代码
    existing_code = state.get("generated_code", "")
    final_code = existing_code + "\n\n" + new_code if existing_code else new_code

    return {"generated_code": final_code, "sandbox_error": None}
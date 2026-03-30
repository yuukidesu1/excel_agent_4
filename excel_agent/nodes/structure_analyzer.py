"""
nodes/structure_analyzer.py — 结构分析与代码解耦节点 (SA)

核心逻辑：
遍历所有未命中的子表，将带有硬编码的原代码和该表的物理锚点喂给 LLM。
LLM 通过“数学做减法”，将原代码中的绝对行号/列号，替换为基于 start_row/col 的相对偏移运算。
"""

import re
import json
import os
from pathlib import Path
from typing import Dict, Optional

from langchain_openai import ChatOpenAI
from langchain_core.messages import SystemMessage, HumanMessage

from excel_agent.state import AgentState, CacheState
from excel_agent.nodes.quality import QUALITY_THRESHOLD

_SYSTEM_PROMPT = """\
你是一个资深的 Python 架构师。你的任务是对“一次性”的 Excel 提取脚本进行“参数化重构”。
初级工程师写死了所有的行号和列号。你需要将目标子表的逻辑剥离，并将绝对坐标转换为基于 `start_row` 和 `start_col` 的相对偏移量。

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
【核心重构任务】
1. 剥离专属逻辑：只提取处理当前【目标子表】的代码。
2. 修改函数签名：你必须重构为 `def extract(ws, merged_map: dict, start_row: int, start_col: int) -> list:` 
   （注意：返回值直接是二维数组 list，不要返回 dict）。
3. 坐标相对化 (极度重要！通过做减法替换数字)：
   将所有写死的绝对行号和列号，替换为传入的 `start_row` 和 `start_col` 的加减法运算。
   - 举例 (行)：如果提示中告诉你该表起步 start_row=4，而原代码写了 `data_start_row = 6`，你必须改写为 `data_start_row = start_row + 2`。
   - 举例 (列)：如果起步 start_col=1，而原代码写了 `cols = [2, 3, 4, 5, 9, 10]`，你必须改写为 `cols = [start_col + 1, start_col + 2, start_col + 3, start_col + 4, start_col + 8, start_col + 9]`。
   - 举例 (探针)：原代码中的 `cell_val(r, 2)`，必须改写为 `cell_val(r, start_col + 1)`。

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
【输出格式】
严格返回 JSON：
{{
  "refactored_code": "def extract(ws, merged_map: dict, start_row: int, start_col: int):\\n    def cell_val(r, c):\\n        # ...\\n    # 基于 start_row 的重构逻辑\\n    return table_list"
}}
"""

def _get_llm() -> ChatOpenAI:
    from dotenv import load_dotenv
    for parent in Path(__file__).resolve().parents:
        env_file = parent / ".env"
        if env_file.exists():
            load_dotenv(env_file, override=True)
            break
    api_key = os.getenv("OPENAI_API_KEY")
    base_url = os.getenv("OPENAI_BASE_URL") or None
    # 推荐使用强推理模型做逻辑重构
    model = os.getenv("LLM_MODEL", "glm-4-plus")
    return ChatOpenAI(model=model, temperature=0, api_key=api_key, base_url=base_url)

def _extract_json(raw: str) -> Optional[Dict]:
    m = re.search(r'```json\s*([\s\S]*?)```', raw)
    if m:
        try: return json.loads(m.group(1).strip())
        except: pass
    m = re.search(r'```\s*([\s\S]*?)```', raw)
    if m:
        try: return json.loads(m.group(1).strip())
        except: pass
    try: return json.loads(raw.strip())
    except: return None

def structure_analyzer_node(state: AgentState) -> dict:
    code = state.get("generated_code", "")
    quality_score = state.get("quality_score", 0.0)
    cache_state: CacheState = state.get("cache", {})
    missed_subtables = cache_state.get("missed_subtables", [])
    entries = cache_state.get("entries", {})

    if not code or not missed_subtables:
        cache_state["analyzer_skipped"] = True
        cache_state["analyzer_skip_reason"] = "没有产生新代码或无未命中子表，跳过。"
        return {"cache": cache_state}
    # ———————————————— DEBUG ——————————————————
    if quality_score < QUALITY_THRESHOLD:
    # if quality_score < 2.0:
        cache_state["analyzer_skipped"] = True
        cache_state["analyzer_skip_reason"] = f"质量分 {quality_score} 过低，拒绝入库。"
        return {"cache": cache_state}

    success_count = 0
    for raw_title in missed_subtables:
        title = raw_title[0]
        if title not in entries:
            continue

        entry_meta = entries[title]
        # 提取当前子表真实的物理锚点
        s_row = entry_meta.get("start_row", 1)
        s_col = entry_meta.get("start_col", 1)

        human_msg = (
            f"【目标子表】: {title}\n"
            f"【该表的物理锚点 (用于计算偏移量)】: start_row={s_row}, start_col={s_col}\n\n"
            f"请重构以下代码中处理 '{title}' 的部分：\n"
            f"```python\n{code}\n```"
        )

        try:
            response = _get_llm().invoke([SystemMessage(content=_SYSTEM_PROMPT), HumanMessage(content=human_msg)])
            parsed_dict = _extract_json(response.content)

            if parsed_dict and "refactored_code" in parsed_dict:
                entries[title]["code"] = parsed_dict["refactored_code"]
                success_count += 1
            else:
                print(f"  ⚠ [SA节点] '{title}' 代码重构失败：JSON解析异常。")
        except Exception as e:
            print(f"  ⚠ [SA节点] 处理 '{title}' 时发生异常：{str(e)}")

    cache_state["entries"] = entries
    cache_state["analyzer_skipped"] = False
    cache_state["analyzer_skip_reason"] = f"成功通过 LLM 解耦了 {success_count}/{len(missed_subtables)} 个子表的代码。"
    return {"cache": cache_state}
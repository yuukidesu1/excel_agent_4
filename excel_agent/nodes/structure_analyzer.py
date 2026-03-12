"""
nodes/structure_analyzer.py — 结构分析与代码拆包节点 (SA)

职责：
    在新架构中，物理结构指纹 (Signature) 已由 PSA 节点纯代码生成。
    本节点 (SA) 的核心职责是"代码拆解与入档"：
    当 code_gen 生成了处理多个 missed_subtables 的大段混合代码且跑通后，
    调用 LLM 将这段代码重构、拆分为针对每个子表独立可运行的 extract 函数。

    拆分后的纯净代码将被存入 state["cache"]["entries"][title]["code"] 中，
    供最后一个 cache_save 节点存入数据库。

优化：
    如果 missed_subtables 只有一个表，说明生成的代码已经是该表专属的，
    直接零成本绑定，跳过大模型调用！
"""

import re
import json
import os
from pathlib import Path
from typing import Dict, Any, Optional

from langchain_openai import ChatOpenAI
from langchain_core.messages import SystemMessage, HumanMessage

from excel_agent.state import AgentState, CacheState
from excel_agent.nodes.quality import QUALITY_THRESHOLD


_SYSTEM_PROMPT = """\
你是一个资深的 Python 代码重构专家。

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
【背景与任务】
我有一段已经测试通过的 Excel 解析代码（generated_code）。这段代码目前的逻辑是**混合**的，它在同一个 `extract` 函数中同时提取了以下多个表格的数据：{missed_subtables}。

为了实现高复用性的缓存机制，我需要你将这段大代码**拆分**成针对每个表格的独立、可运行的提取代码块。

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
【重构要求】
1. **彻底解耦**：拆分后的每段代码必须只专注于提取其对应的那一个表格。删除与该表格无关的其他表格的定位和提取逻辑。
2. **函数签名必须保持一致**：每个独立代码块都必须包含且只包含一个入口函数：`def extract(ws, merged_map):`，并返回对应表格提取出的二维数组 (List[List]) 或者是只有该表数据的字典。
3. **保留原汁原味**：不要修改原代码中关于该表格的核心坐标推算、正则提取或数值清洗逻辑，你只是在做"裁剪"和"拆分"。
4. **无需多余解释**：严格以 JSON 格式输出，Key 为子表名称，Value 为拆分后的完整 Python 代码字符串。

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
【输出格式样例】
请严格输出如下 JSON 格式：
```json
{{
  "2G Configuration": "def extract(ws, merged_map):\\n    data = []\\n    # ... 这里是专门提取 2G 的原始代码 ...\\n    return data",
  "4G Configuration": "def extract(ws, merged_map):\\n    data = []\\n    # ... 这里是专门提取 4G 的原始代码 ...\\n    return data"
}}
"""
def _get_llm() -> ChatOpenAI:
    """获取 LLM 实例"""
    from dotenv import load_dotenv

    for parent in Path(__file__).resolve().parents:
        env_file = parent / ".env"
        if env_file.exists():
            load_dotenv(env_file, override=True)
            break

    api_key = os.getenv("OPENAI_API_KEY")
    base_url = os.getenv("OPENAI_BASE_URL") or None
    model = os.getenv("LLM_MODEL", "glm-4")

    if not api_key:
        raise ValueError("未找到 OPENAI_API_KEY，请检查 .env 文件。")

    return ChatOpenAI(model=model, temperature=0, api_key=api_key, base_url=base_url)


def _extract_json(raw: str) -> Optional[Dict]:
    """从 LLM 输出中提取 JSON"""
    m = re.search(r'```json\s*([\s\S]*?)```', raw)

    if m:
        try:
            return json.loads(m.group(1).strip())
        except json.JSONDecodeError:
            pass
    m = re.search(r'```\s*([\s\S]*?)```', raw)
    if m:
        try:
            return json.loads(m.group(1).strip())
        except json.JSONDecodeError:
            pass

    try:
        return json.loads(raw.strip())
    except json.JSONDecodeError:
        return None

def structure_analyzer_node(state: AgentState) -> dict:
    """
    代码拆解与入档节点 (SA)
    """
    code = state.get("generated_code", "")
    quality_score = state.get("quality_score", 0.0)

    cache_state: CacheState = state.get("cache", {})
    missed_subtables = cache_state.get("missed_subtables", [])
    entries = cache_state.get("entries", {})

    # 1. 基础异常与拦截校验
    if not code or not missed_subtables:
        cache_state["analyzer_skipped"] = True
        cache_state["analyzer_skip_reason"] = "没有产生新代码或无未命中子表，直接跳过。"
        return {"cache": cache_state}

    # 尽管 agent.py 的路由已经拦截了低分，这里再加一层保险
    if quality_score < QUALITY_THRESHOLD:
        cache_state["analyzer_skipped"] = True
        cache_state["analyzer_skip_reason"] = f"代码跑通质量分 {quality_score} 偏低，拒绝将其入库污染缓存。"
        return {"cache": cache_state}

    # ── ★ 极速优化路径：只有一个子表未命中时，零成本直接绑定！ ──
    if len(missed_subtables) == 1:
        single_target = missed_subtables[0]
        if single_target in entries:
            entries[single_target]["code"] = code

        cache_state["analyzer_skipped"] = True
        cache_state["analyzer_skip_reason"] = f"仅存在 1 个目标子表 '{single_target}'，新代码即为其专属代码，跳过 LLM 拆分。"
        return {"cache": cache_state}

    # ── 2. LLM 拆包重构路径 (多个子表时) ──
    prompt = _SYSTEM_PROMPT.format(missed_subtables=json.dumps(missed_subtables, ensure_ascii=False))

    messages = [
        SystemMessage(content=prompt),
        HumanMessage(content=f"请拆分以下代码：\n```python\n{code}\n```")
    ]

    try:
        response = _get_llm().invoke(messages)
        parsed_dict = _extract_json(response.content)

        if not parsed_dict:
            cache_state["analyzer_skipped"] = True
            cache_state["analyzer_skip_reason"] = "LLM 代码拆分失败，返回格式无法解析为 JSON。"
            return {"cache": cache_state}

        # ── 3. 将拆分好的代码填写入 Cache 档案中 ──
        success_count = 0
        for table_name, split_code in parsed_dict.items():
            if table_name in entries:
                entries[table_name]["code"] = split_code
                success_count += 1

        # ★ 关键修复：重新赋值 entries 到 cache_state，确保 LangGraph 状态系统能检测到变化
        cache_state["entries"] = entries

        cache_state["analyzer_skipped"] = False
        cache_state["analyzer_skip_reason"] = f"成功通过 LLM 拆分了 {success_count} 个子表的专属代码。"

        return {"cache": cache_state}

    except Exception as e:
        cache_state["analyzer_skipped"] = True
        cache_state["analyzer_skip_reason"] = f"拆分代码过程发生异常：{str(e)}"
        return {"cache": cache_state}

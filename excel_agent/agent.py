"""
agent.py — LangGraph 图组装 + 对外入口（支持三级缓存）

流程（带缓存）：

  第一次运行（无缓存）：
  parse_node
      ↓
  cache_query_node (L1 only) → 未命中
      ↓
  code_gen_node [LLM]
      ↓
  sandbox_node [代码]
      ↓
  restore_node [代码]
      ↓
  quality_node [代码]
      ↓
  structure_analyzer_node [LLM] (质量>=QUALITY_THRESHOLD 时提取结构指纹)
      ↓
  cache_save_node [缓存] (保存 L1+L2)
      ↓
  route ──→ 通过 → END ✅
           └──→ 失败 → retry → code_gen

  后续运行（L1 命中）：
  parse_node
      ↓
  cache_query_node (L1) → 命中
      ↓
  sandbox_node [直接使用缓存代码]
      ↓
  restore_node
      ↓
  quality_node
      ↓
  END ✅

  相似结构文件（L2 命中）：
  由于 L2 需要 structure_fingerprint，而该指纹需要 code_gen 后才能生成，
  因此 L2 缓存暂时不支持跨文件复用。

  【未来优化】：可以在 parse 后立即计算一个轻量级结构指纹用于 L2 查询，
  而不依赖 structure_analyzer 的详细分析。
"""

import sys
import os
from pathlib import Path
from typing import Optional, List, Dict, Any, AsyncGenerator

from langgraph.graph import StateGraph, END

from excel_agent.state import AgentState
from excel_agent.nodes.parse import parse_node
from excel_agent.nodes.pre_structure_analyzer import pre_structure_analyzer_node
from excel_agent.nodes.code_gen import code_gen_node
from excel_agent.nodes.sandbox import sandbox_node
from excel_agent.nodes.restore import restore_node
from excel_agent.nodes.quality import quality_node, route, QUALITY_THRESHOLD
from excel_agent.nodes.structure_analyzer import structure_analyzer_node
from excel_agent.nodes.cache_query import cache_query_node
from excel_agent.nodes.cache_save import cache_save_node

# 开启 LangSmith 追踪
os.environ["LANGCHAIN_TRACING_V2"] = "true"
os.environ["LANGCHAIN_PROJECT"] = "excel_agent"
os.environ["LANGCHAIN_API_KEY"] = "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"


def _retry_node(state: AgentState) -> dict:
    return {"retry_count": state.get("retry_count", 0) + 1}


def _route_after_cache(state: AgentState) -> str:
    """缓存查询后路由：全量命中→跳过LLM直接sandbox，部分/全未命中→code_gen"""
    cache_state = state.get("cache", {})
    if cache_state.get("all_cached", False):
        return "sandbox"
    return "code_gen"


def _route_after_quality(state: AgentState) -> str:
    """
        质量检查后路由：
        - 全量命中且跑完的 → 直接结束（没产生新代码，无需分析保存）
        - 有新产生代码的 + 质量达标 → structure_analyzer -> cache_save
        - 质量不达标 → retry
        """
    cache_state = state.get("cache", {})

    # 如果所有子表都命中了缓存，直接结束，不走后续的保存流
    if cache_state.get("all_cached", False):
        return "end"

    # 以下是有 LLM 新生成代码的情况
    if state["quality_score"] >= QUALITY_THRESHOLD:
        return "analyze"
    if state.get("retry_count", 0) >= 3:
        return "analyze"  # 超过重试上限，死马当活马医，保存备用
    return "retry"



def build_agent():
    g = StateGraph(AgentState)

    # 1. 注册所有节点
    g.add_node("parse", parse_node)
    g.add_node("pre_structure_analyzer", pre_structure_analyzer_node)
    g.add_node("cache_query", cache_query_node)
    g.add_node("code_gen", code_gen_node)
    g.add_node("sandbox", sandbox_node)
    g.add_node("restore", restore_node)
    g.add_node("quality", quality_node)
    g.add_node("structure_analyzer", structure_analyzer_node)
    g.add_node("cache_save", cache_save_node)
    g.add_node("retry", _retry_node)

    # 2. 定义边 (数据流)
    g.set_entry_point("parse")
    g.add_edge("parse", "pre_structure_analyzer")
    g.add_edge("pre_structure_analyzer", "cache_query")

    # 3. 缓存路由：决定是否调用 LLM
    g.add_conditional_edges("cache_query", _route_after_cache, {
        "sandbox": "sandbox",
        "code_gen": "code_gen"
    })

    g.add_edge("code_gen", "sandbox")
    g.add_edge("sandbox", "restore")
    g.add_edge("restore", "quality")

    # 4. 质量控制路由：决定重试还是去切分保存代码
    g.add_conditional_edges("quality", _route_after_quality, {
        "analyze": "structure_analyzer",
        "retry": "retry",
        "end": END
    })

    g.add_edge("structure_analyzer", "cache_save")
    g.add_edge("cache_save", END)

    # 5. 重试逻辑：回到 code_gen 重新生成（针对未命中的子表）
    g.add_edge("retry", "code_gen")

    return g.compile()


def _make_initial(
    excel_path: str,
    sheet_name: str,
    subtable_titles: List[str],
    hints: Optional[str],
    target_columns: Optional[Dict[str, List[Dict[str, Any]]]],
) -> AgentState:
    return {
        "config": {
            "excel_path": excel_path,
            "sheet_name": sheet_name,
            "subtable_titles": subtable_titles,
            "hints": hints,
            "target_columns": target_columns,
        },
        "sheet_structure": None,
        "generated_code": None,
        "raw_result": None,
        "raw_data": None,
        "result": None,
        "final_output": None,
        "quality_score": 0.0,
        "retry_count": 0,
        "errors": [],
        "sandbox_error": None,
        # 缓存相关
        "cache": {
            "entries": {},
            "all_cached": False,
            "partial_cached": False,
            "missed_subtables": [],
            "analyzer_skipped": False,
            "analyzer_skip_reason": None,
        }
    }


def run_extraction(
    excel_path: str,
    sheet_name: str,
    subtable_titles: List[str],
    hints: Optional[str] = None,
    target_columns: Optional[Dict[str, List[Dict[str, Any]]]] = None,
) -> dict:

    agent = build_agent()

    """可视化图"""
    # from IPython.display import Image, display
    #
    # try:
    #     display(Image(agent.get_graph().draw_mermaid_png()))
    # except Exception:
    #     pass
    #
    # import matplotlib.pyplot as plt
    # import matplotlib.image as mpimg
    # import io
    #
    # png_data = agent.get_graph().draw_mermaid_png()
    # img = mpimg.imread(io.BytesIO(png_data))
    # plt.figure(figsize=(15, 10), dpi=300)
    # plt.imshow(img, interpolation='lanczos')  # 使用 lanczos 插值算法平滑边缘
    # plt.axis('off')
    # plt.show()

    final = agent.invoke(
        _make_initial(excel_path, sheet_name, subtable_titles, hints, target_columns)
    )

    return {
        "success": final["quality_score"] >= QUALITY_THRESHOLD,
        "data": final.get("final_output") or final.get("result"),
        "quality_score": final["quality_score"],
        "retry_count": final["retry_count"],
        "errors": final["errors"],
        "generated_code": final.get("generated_code", ""),
        "sandbox_error": final.get("sandbox_error"),
        # 缓存信息
        "all_cached": final.get("cache", {}).get("all_cached", False),
        "partial_cached": final.get("cache", {}).get("partial_cached", False),
    }


async def run_extraction_deep_stream(
    excel_path: str,
    sheet_name: str,
    subtable_titles: List[str],
    hints: Optional[str] = None,
    target_columns: Optional[List[Dict]] = None,
) -> AsyncGenerator[dict, None]:
    """
    异步流式入口（调试用）。
    捕获 Token 吐字、节点完成事件。
    """
    agent = build_agent()
    initial = _make_initial(excel_path, sheet_name, subtable_titles, hints, target_columns)

    async for event in agent.astream_events(initial, version="v2"):
        kind = event["event"]
        name = event["name"]

        if kind == "on_chat_model_stream":
            content = event["data"]["chunk"].content
            if content:
                yield {"type": "token", "content": content}

        elif kind == "on_chain_end" and name in (
            "parse", "cache_query", "code_gen", "sandbox",
            "restore", "quality", "structure_analyzer", "cache_save"
        ):
            yield {
                "type": "node_end",
                "name": name,
                "data": event["data"].get("output", {}),
            }

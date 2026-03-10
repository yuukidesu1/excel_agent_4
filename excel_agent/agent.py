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
  structure_analyzer_node [LLM] (质量>=0.8 时提取结构指纹)
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
from excel_agent.nodes.light_structure_analyzer import light_structure_analyzer_node
from excel_agent.nodes.code_gen import code_gen_node
from excel_agent.nodes.sandbox import sandbox_node
from excel_agent.nodes.restore import restore_node
from excel_agent.nodes.quality import quality_node, route, QUALITY_THRESHOLD
from excel_agent.nodes.structure_analyzer import structure_analyzer_node
from excel_agent.nodes.cache_query import cache_query_node
from excel_agent.nodes.cache_save import cache_save_node

# 开启 LangSmith 追踪
os.environ["LANGCHAIN_TRACING_V2"] = "true"
os.environ["LANGCHAIN_PROJECT"] = "LangGraph_Debug_Test"
os.environ["LANGCHAIN_API_KEY"] = "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"


def _retry_node(state: AgentState) -> dict:
    return {"retry_count": state.get("retry_count", 0) + 1}


def _route_after_cache(state: AgentState) -> str:
    """缓存查询后路由：命中→sandbox，未命中→code_gen"""
    if state.get("cache_hit"):
        return "sandbox"
    return "code_gen"


def _route_after_quality(state: AgentState) -> str:
    """
    质量检查后路由：
    - 通过 + 来自缓存 → 直接结束（无需重复分析/保存）
    - 通过 + 来自 LLM → structure_analyzer → cache_save
    - 失败 → retry
    """
    # 如果数据来自缓存，不需要重复分析和保存
    if state.get("cache_hit"):
        return "end"

    # 来自 LLM 生成
    if state["quality_score"] >= QUALITY_THRESHOLD:
        return "analyze"
    if state.get("retry_count", 0) >= 3:
        return "analyze"  # 超过重试上限，仍然分析结构以便未来复用
    return "retry"


def _route_final(state: AgentState) -> str:
    """最终路由：决定是结束还是重试"""
    if state.get("cache_hit"):
        # 缓存命中的情况，质量达标就直接结束
        if state["quality_score"] >= QUALITY_THRESHOLD:
            return "end"
        # 缓存命中的代码质量不达标，重试也解决不了（还是同样的缓存）
        # 所以直接结束
        return "end"

    # 非缓存命中，按质量路由
    if state["quality_score"] >= QUALITY_THRESHOLD:
        return "end"
    if state.get("retry_count", 0) >= 3:
        return "end"
    return "retry"


def build_agent():
    g = StateGraph(AgentState)

    # 添加节点
    g.add_node("parse", parse_node)
    g.add_node("light_structure_analyzer", light_structure_analyzer_node)
    g.add_node("cache_query", cache_query_node)
    g.add_node("code_gen", code_gen_node)
    g.add_node("sandbox", sandbox_node)
    g.add_node("restore", restore_node)
    g.add_node("quality", quality_node)
    g.add_node("structure_analyzer", structure_analyzer_node)
    g.add_node("cache_save", cache_save_node)
    g.add_node("retry", _retry_node)

    # 设置流程
    g.set_entry_point("parse")
    g.add_edge("parse", "light_structure_analyzer")
    g.add_edge("light_structure_analyzer", "cache_query")

    # 缓存查询后分流
    g.add_conditional_edges("cache_query", _route_after_cache, {
        "sandbox": "sandbox",
        "code_gen": "code_gen"
    })

    # 从 code_gen 或 sandbox 继续
    g.add_edge("code_gen", "sandbox")
    g.add_edge("sandbox", "restore")
    g.add_edge("restore", "quality")

    # 质量检查后决定：通过→分析结构，失败→重试；缓存命中→直接结束
    g.add_conditional_edges("quality", _route_after_quality, {
        "analyze": "structure_analyzer",
        "retry": "retry",
        "end": END
    })

    # 结构分析后保存缓存
    g.add_edge("structure_analyzer", "cache_save")

    # 缓存保存后结束
    g.add_edge("cache_save", END)

    # 重试逻辑
    g.add_edge("retry", "code_gen")  # 重试跳过 parse 和 cache_query，直接重新生成

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
        "cache_hit": None,
        "cache_level": None,
        "structure_fingerprint": None,
        "light_structure_fingerprint": None,
        "analyzer_skipped": None,
        "analyzer_skip_reason": None,
    }


def run_extraction(
    excel_path: str,
    sheet_name: str,
    subtable_titles: List[str],
    hints: Optional[str] = None,
    target_columns: Optional[Dict[str, List[Dict[str, Any]]]] = None,
) -> dict:
    """
    同步调用入口（支持三级缓存）。

    参数：
        excel_path      : Excel 文件路径
        sheet_name      : Sheet 名，如 "CONFIGURATION"
        subtable_titles  : 子表标题关键词，如 "4G Configuration"
        hints           : 可选额外提示
        target_columns  : 列过滤，None=全部。示例：
                          [
                            {"parent": None, "child": "CELL"},
                            {"parent": "ANTENNAS", "child": "Antenna Qty."},
                          ]

    返回：
        {
            "success":        bool,
            "data":           Dict[str, List[List[str]]],
            "quality_score":  float,
            "retry_count":    int,
            "errors":         List[str],
            "generated_code": str,
            "cache_hit":      bool,      # 是否命中缓存
            "cache_level":    str,       # "l1" | "l2" | "l3"
        }
    """
    agent = build_agent()

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
        "cache_hit": final.get("cache_hit"),
        "cache_level": final.get("cache_level"),
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

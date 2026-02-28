"""
agent.py — LangGraph 图组装 + 对外入口

流程：
  parse → locate → extract → restore → quality
            ↑                              |
            └────── retry（最多3次）←──────┘
"""

import sys
import os
from pathlib import Path
from langgraph.graph import StateGraph, END
from typing import Optional, List, Dict, Any, AsyncGenerator

from excel_agent.state            import AgentState
from excel_agent.nodes.parse      import parse_node
from excel_agent.nodes.locate     import locate_node
from excel_agent.nodes.extract    import extract_node
from excel_agent.nodes.restore    import restore_node
from excel_agent.nodes.quality    import quality_node, route, QUALITY_THRESHOLD

# 开启 LangSmith 追踪
# os.environ["LANGCHAIN_TRACING_V2"] = "true"
# os.environ["LANGCHAIN_PROJECT"] = "LangGraph_Debug_Test"
# os.environ["LANGCHAIN_API_KEY"] = "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" # 去 smith.langchain.com 免费申请一个


def _retry_node(state: AgentState) -> dict:
    return {"retry_count": state.get("retry_count", 0) + 1}


def build_agent():
    g = StateGraph(AgentState)
    g.add_node("parse",   parse_node)
    g.add_node("locate",  locate_node)
    g.add_node("extract", extract_node)
    g.add_node("restore", restore_node)
    g.add_node("quality", quality_node)
    g.add_node("retry",   _retry_node)

    g.set_entry_point("parse")
    g.add_edge("parse",   "locate")
    g.add_edge("locate",  "extract")
    g.add_edge("extract", "restore")
    g.add_edge("restore", "quality")
    g.add_edge("retry",   "locate")
    g.add_conditional_edges("quality", route, {"end": END, "retry": "retry"})
    return g.compile()


def _make_initial(excel_path: str, sheet_name: str, subtable_title: str,
                  hints: Optional[str], target_columns: Optional[List[Dict]]) -> AgentState:
    return {
        "config":{
            "excel_path": excel_path,
            "sheet_name": sheet_name,
            "subtable_title": subtable_title,
            "hints": hints,
            "target_columns": target_columns,
        },
        "sheet_structure": None,
        "header_map":      None,
        "raw_data":        None,
        "result":          None,
        "quality_score":   0.0,
        "retry_count":     0,
        "errors":          [],
        "final_output":    None,
    }


def run_extraction(
    excel_path:     str,
    sheet_name:     str,
    subtable_title: str,
    hints:          Optional[str]            = None,
    target_columns: Optional[List[Dict]]     = None,
) -> dict:
    """
    同步调用入口。

    参数：
        excel_path      : Excel 文件路径
        sheet_name      : Sheet 名，如 "CONFIGURATION"
        subtable_title  : 子表标题关键词，如 "4G Configuration"
        hints           : 可选提示
        target_columns  : 可选列过滤，None=全部列，示例：
                          [
                            {"parent": None,       "child": "CELL"},
                            {"parent": "ANTENNAS", "child": "NEW/SWAP/EXIST"},
                            {"parent": "RF MODULE","child": "TYPE"},
                          ]

    返回：
        {
            "success":       bool,
            "data":          List[List[str]],   # 第 0 行为列名
            "quality_score": float,
            "retry_count":   int,
            "errors":        List[str],
        }
    """
    agent = build_agent()

    """对编译好的图进行可视化"""
    import matplotlib.pyplot as plt
    import matplotlib.image as mpimg
    import io
    png_data = agent.get_graph().draw_mermaid_png()
    img = mpimg.imread(io.BytesIO(png_data))

    plt.figure(figsize=(12, 8), dpi=300)
    plt.imshow(img, interpolation="lanczos")
    plt.axis('off')
    plt.show()

    final = agent.invoke(_make_initial(excel_path, sheet_name, subtable_title,
                                       hints, target_columns))
    return {
        "success":       final["quality_score"] >= QUALITY_THRESHOLD,
        "data":          final.get("final_output") or final.get("result"),
        "quality_score": final["quality_score"],
        "retry_count":   final["retry_count"],
        "errors":        final["errors"],
    }


async def run_extraction_deep_stream(
    excel_path:     str,
    sheet_name:     str,
    subtable_title: str,
    hints:          Optional[str]        = None,
    target_columns: Optional[List[Dict]] = None,
) -> AsyncGenerator[dict, None]:
    """
    异步流式调用入口（调试用）。捕获 Token、工具调用、节点完成事件。

    用法：
        async for chunk in run_extraction_deep_stream(...):
            if chunk["type"] == "token":
                print(chunk["content"], end="", flush=True)
            elif chunk["type"] == "node_end":
                print(f"节点 {chunk['name']} 完成")
    """
    agent = build_agent()
    initial = _make_initial(excel_path, sheet_name, subtable_title,
                            hints, target_columns)

    async for event in agent.astream_events(initial, version="v2"):
        kind = event["event"]
        name = event["name"]

        if kind == "on_chat_model_stream":
            content = event["data"]["chunk"].content
            if content:
                yield {"type": "token", "content": content}

        elif kind == "on_tool_start":
            yield {"type": "tool_start", "name": name,
                   "input": event["data"].get("input")}

        elif kind == "on_tool_end":
            yield {"type": "tool_end", "name": name,
                   "output": event["data"].get("output")}

        elif kind == "on_chain_end" and name in ("parse","locate","extract","restore","quality"):
            yield {"type": "node_end", "name": name,
                   "data": event["data"].get("output", {})}
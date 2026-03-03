"""
agent.py — LangGraph 图组装 + 对外入口（Agent B：LLM 写代码版）

流程：
  parse_node
      ↓
  code_gen_node   [LLM]  看 parse 数据，写 openpyxl 抽取函数
      ↓
  sandbox_node    [代码]  在受限命名空间执行 LLM 生成的代码
      ↓
  restore_node    [代码]  格式化 + 列过滤
      ↓
  quality_node    [代码]  质量打分
      ↓
  route ──→ 通过 → END ✅
       └──→ 失败（代码报错 or 质量不达标）→ retry → code_gen（携带错误上下文）

重试策略：
  - sandbox_error 不为空（代码执行失败）→ 告知 LLM 具体报错信息
  - quality_score 不达标              → 告知 LLM 质量问题
  - 最多重试 MAX_RETRY 次后强制输出
"""

import sys
import os
from pathlib import Path
from typing import Optional, List, Dict, Any, AsyncGenerator

from langgraph.graph import StateGraph, END

from excel_agent.state            import AgentState
from excel_agent.nodes.parse      import parse_node
from excel_agent.nodes.code_gen   import code_gen_node
from excel_agent.nodes.sandbox    import sandbox_node
from excel_agent.nodes.restore    import restore_node
from excel_agent.nodes.quality    import quality_node, route, QUALITY_THRESHOLD

# 开启 LangSmith 追踪
os.environ["LANGCHAIN_TRACING_V2"] = "true"
os.environ["LANGCHAIN_PROJECT"] = "LangGraph_Debug_Test"
os.environ["LANGCHAIN_API_KEY"] = "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"

def _retry_node(state: AgentState) -> dict:
    return {"retry_count": state.get("retry_count", 0) + 1}


def build_agent():
    g = StateGraph(AgentState)

    g.add_node("parse",    parse_node)
    g.add_node("code_gen", code_gen_node)
    g.add_node("sandbox",  sandbox_node)
    g.add_node("restore",  restore_node)
    g.add_node("quality",  quality_node)
    g.add_node("retry",    _retry_node)

    g.set_entry_point("parse")
    g.add_edge("parse",    "code_gen")
    g.add_edge("code_gen", "sandbox")
    g.add_edge("sandbox",  "restore")
    g.add_edge("restore",  "quality")
    g.add_edge("retry",    "code_gen")   # 重试跳过 parse，直接重新生成代码

    g.add_conditional_edges("quality", route, {"end": END, "retry": "retry"})

    return g.compile()


def _make_initial(
    excel_path:     str,
    sheet_name:     str,
    subtable_title: str,
    hints:          Optional[str],
    target_columns: Optional[List[Dict]],
) -> AgentState:
    return {
        "config": {
            "excel_path": excel_path,
            "sheet_name": sheet_name,
            "subtable_title": subtable_title,
            "hints": hints,
            "target_columns": target_columns,
        },
        "sheet_structure": None,
        "generated_code":  None,
        "raw_result":      None,
        "raw_data":        None,
        "result":          None,
        "final_output":    None,
        "quality_score":   0.0,
        "retry_count":     0,
        "errors":          [],
        "sandbox_error":   None,
    }


def run_extraction(
    excel_path:     str,
    sheet_name:     str,
    subtable_title: str,
    hints:          Optional[str]        = None,
    target_columns: Optional[List[Dict]] = None,
) -> dict:
    """
    同步调用入口。

    参数：
        excel_path      : Excel 文件路径
        sheet_name      : Sheet 名，如 "CONFIGURATION"
        subtable_title  : 子表标题关键词，如 "4G Configuration"
        hints           : 可选额外提示
        target_columns  : 列过滤，None=全部。示例：
                          [
                            {"parent": None,       "child": "CELL"},
                            {"parent": "ANTENNAS", "child": "Antenna Qty."},
                          ]

    返回：
        {
            "success":        bool,
            "data":           List[List[str]],
            "quality_score":  float,
            "retry_count":    int,
            "errors":         List[str],
            "generated_code": str,   # LLM 生成的代码（供调试）
        }
    """
    agent = build_agent()

    """对编译好的图进行可视化"""
    # import matplotlib.pyplot as plt
    # import matplotlib.image as mpimg
    # import io
    # png_data = agent.get_graph().draw_mermaid_png()
    # img = mpimg.imread(io.BytesIO(png_data))
    #
    # plt.figure(figsize=(12, 8), dpi=300)
    # plt.imshow(img, interpolation="lanczos")
    # plt.axis('off')
    # plt.show()

    final = agent.invoke(
        _make_initial(excel_path, sheet_name, subtable_title, hints, target_columns)
    )
    return {
        "success":        final["quality_score"] >= QUALITY_THRESHOLD,
        "data":           final.get("final_output") or final.get("result"),
        "quality_score":  final["quality_score"],
        "retry_count":    final["retry_count"],
        "errors":         final["errors"],
        "generated_code": final.get("generated_code", ""),
        "sandbox_error":  final.get("sandbox_error"),
    }


async def run_extraction_deep_stream(
    excel_path:     str,
    sheet_name:     str,
    subtable_title: str,
    hints:          Optional[str]        = None,
    target_columns: Optional[List[Dict]] = None,
) -> AsyncGenerator[dict, None]:
    """
    异步流式入口（调试用）。
    捕获 Token 吐字、节点完成事件。
    """
    agent   = build_agent()
    initial = _make_initial(excel_path, sheet_name, subtable_title, hints, target_columns)

    async for event in agent.astream_events(initial, version="v2"):
        kind = event["event"]
        name = event["name"]

        if kind == "on_chat_model_stream":
            content = event["data"]["chunk"].content
            if content:
                yield {"type": "token", "content": content}

        elif kind == "on_chain_end" and name in (
            "parse", "code_gen", "sandbox", "restore", "quality"
        ):
            yield {
                "type": "node_end",
                "name": name,
                "data": event["data"].get("output", {}),
            }
import os
from typing import Optional, List, Dict, Any, AsyncGenerator

from langgraph.graph import StateGraph, END

from excel_agent.state            import AgentState
from excel_agent.nodes.parse      import parse_node
from excel_agent.nodes.code_gen   import code_gen_node
from excel_agent.nodes.sandbox    import sandbox_node
from excel_agent.nodes.restore    import restore_node
from excel_agent.nodes.quality    import quality_node, route, QUALITY_THRESHOLD
from excel_agent.nodes.kv_preprocess import kv_preprocess_node
from excel_agent.nodes.kv_codegen import kv_codegen_node
from excel_agent.nodes.kv_sandbox import kv_sandbox_node
from excel_agent.nodes.kv_quality import kv_quality_node, kv_route

# 开启 LangSmith 追踪
# os.environ["LANGCHAIN_TRACING_V2"] = "true"
# os.environ["LANGCHAIN_PROJECT"] = "LangGraph_Debug_Test"
# os.environ["LANGCHAIN_API_KEY"] = "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"

def _retry_node(state: AgentState) -> dict:
    return {"retry_count": state.get("retry_count", 0) + 1}


def _route_entry(state: AgentState) -> str:
    config = state.get("config", {})
    extract_type = str(config.get("extract_type") or "").lower()
    if extract_type == "kv" or config.get("kv_list"):
        return "kv_preprocess"
    return "parse"


def build_agent():
    g = StateGraph(AgentState)

    g.add_node("entry",    lambda state: {})
    g.add_node("parse",    parse_node)
    g.add_node("code_gen", code_gen_node)
    g.add_node("sandbox",  sandbox_node)
    g.add_node("restore",  restore_node)
    g.add_node("quality",  quality_node)
    g.add_node("retry",    _retry_node)
    g.add_node("kv_preprocess", kv_preprocess_node)
    g.add_node("kv_codegen",    kv_codegen_node)
    g.add_node("kv_sandbox",    kv_sandbox_node)
    g.add_node("kv_quality",    kv_quality_node)
    g.add_node("kv_retry",      _retry_node)

    g.set_entry_point("entry")
    g.add_conditional_edges("entry", _route_entry, {
        "parse": "parse",
        "kv_preprocess": "kv_preprocess",
    })

    g.add_edge("parse",    "code_gen")
    g.add_edge("code_gen", "sandbox")
    g.add_edge("sandbox",  "restore")
    g.add_edge("restore",  "quality")
    g.add_edge("retry",    "code_gen")   # 重试跳过 parse，直接重新生成代码

    g.add_conditional_edges("quality", route, {"end": END, "retry": "retry"})

    g.add_edge("kv_preprocess", "kv_codegen")
    g.add_edge("kv_codegen",    "kv_sandbox")
    g.add_edge("kv_sandbox",    "kv_quality")
    g.add_edge("kv_retry",      "kv_codegen")
    g.add_conditional_edges("kv_quality", kv_route, {"end": END, "retry": "kv_retry"})

    return g.compile()


def _make_initial(
    excel_path:     str,
    sheet_name:     str,
    subtable_titles: List[str],    # 修改 str -> List[str] 以适配多子表抽取
    hints:          Optional[str],
    target_columns: Optional[Dict[str, List[Dict[str, Any]]]],   # 更新为字典
    extract_type:   Optional[str] = None,
    kv_list:        Optional[List[str]] = None,
) -> AgentState:
    if kv_list is None and extract_type == "kv":
        kv_list = subtable_titles or []

    return {
        "config": {
            "excel_path": excel_path,
            "sheet_name": sheet_name,
            "subtable_titles": subtable_titles,
            "hints": hints,
            "target_columns": target_columns,
            "extract_type": extract_type,
            "kv_list": kv_list,
        },
        "sheet_structure": None,
        "generated_code":  [],
        "raw_result":      None,
        "raw_data":        None,
        "result":          None,
        "final_output":    None,
        "quality_score":   0.0,
        "retry_count":     0,
        "errors":          [],
        "sandbox_error":   None,
        "kv_state":        {},
        "cache":           {},
    }


def run_extraction(
    excel_path:     str,
    sheet_name:     str,
    subtable_titles: List[str],    # str -> List[str]
    hints:          Optional[str]        = None,
    target_columns: Optional[Dict[str, List[Dict[str, Any]]]] = None,    # 更新为字典
    extract_type:   Optional[str] = None,
    kv_list:        Optional[List[str]] = None,
) -> dict:
    """
    同步调用入口。

    参数：
        excel_path      : Excel 文件路径
        sheet_name      : Sheet 名，如 "CONFIGURATION"
        subtable_titles  : 子表标题关键词，如 "4G Configuration"
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
        _make_initial(
            excel_path,
            sheet_name,
            subtable_titles,
            hints,
            target_columns,
            extract_type=extract_type,
            kv_list=kv_list,
        )
    )
    kv_state = final.get("kv_state", {}) or {}
    generated_code = final.get("generated_code", "")
    if kv_state.get("generated_code"):
        generated_code = kv_state.get("generated_code", "")

    return {
        "success":        final["quality_score"] >= QUALITY_THRESHOLD,
        "data":           final.get("final_output") or final.get("result"),
        "quality_score":  final["quality_score"],
        "retry_count":    final["retry_count"],
        "errors":         final["errors"],
        "generated_code": generated_code,
        "sandbox_error":  final.get("sandbox_error"),
    }


async def run_extraction_deep_stream(
    excel_path:     str,
    sheet_name:     str,
    subtable_titles: List[str],
    hints:          Optional[str]        = None,
    target_columns: Optional[List[Dict]] = None,
    extract_type:   Optional[str] = None,
    kv_list:        Optional[List[str]] = None,
) -> AsyncGenerator[dict, None]:
    """
    异步流式入口（调试用）。
    捕获 Token 吐字、节点完成事件。
    """
    agent   = build_agent()
    initial = _make_initial(
        excel_path,
        sheet_name,
        subtable_titles,
        hints,
        target_columns,
        extract_type=extract_type,
        kv_list=kv_list,
    )

    async for event in agent.astream_events(initial, version="v2"):
        kind = event["event"]
        name = event["name"]

        if kind == "on_chat_model_stream":
            content = event["data"]["chunk"].content
            if content:
                yield {"type": "token", "content": content}

        elif kind == "on_chain_end" and name in (
            "parse", "code_gen", "sandbox", "restore", "quality",
            "kv_preprocess", "kv_codegen", "kv_sandbox", "kv_quality"
        ):
            yield {
                "type": "node_end",
                "name": name,
                "data": event["data"].get("output", {}),
            }

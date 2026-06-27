from typing import Optional, List, Dict, Any, AsyncGenerator, Union

from langgraph.graph import StateGraph, END

from excel_agent.state import AgentState, SubtableConfig
from excel_agent.nodes.parse import parse_node
from excel_agent.nodes.kv_preprocess import kv_preprocess_node
from excel_agent.nodes.kv_codegen import kv_codegen_node
from excel_agent.nodes.pre_structureanalyzer import pre_structure_analyzer_node
from excel_agent.nodes.code_gen import code_gen_node
from excel_agent.nodes.sandbox import sandbox_node
from excel_agent.nodes.restore import restore_node
from excel_agent.nodes.quality import quality_node, QUALITY_THRESHOLD
from excel_agent.nodes.structure_analyzer import structure_analyzer_node
from excel_agent.nodes.cache_query import cache_query_node
from excel_agent.nodes.cache_save import cache_save_node


# ========================================== #
# 阶段 1：状态与数据初始化
# ========================================== #

def _make_initial(
        excel_path: str,
        sheet_name: str,
        subtable_titles: Optional[List[str]],
        hints: Optional[str],
        subtable_configs: Optional[Dict[str, Union[SubtableConfig, List[Any]]]],
        kv_list: Optional[List[str]] = None,
) -> AgentState:
    return {
        "config": {
            "excel_path": excel_path,
            "sheet_name": sheet_name,
            "subtable_titles": subtable_titles,
            "hints": hints,
            "subtable_configs": subtable_configs,
            "kv_list": kv_list,
        },
        "sheet_structure": None,
        "generated_code": [],
        "raw_result": None,
        "raw_data": None,
        "result": None,
        "final_output": None,
        "quality_score": 0.0,
        "retry_count": 0,
        "errors": [],
        "sandbox_error": None,
        "cache": {
            "entries": {},
            "all_cached": False,
            "partial_cached": False,
            "missed_subtables": [],
            "analyzer_skipped": False,
            "analyzer_skip_reason": None,
        }
    }


# ========================================== #
# 阶段 2：独立节点与条件路由判断函数
# ========================================== #

def _retry_node(state: AgentState) -> dict:
    return {"retry_count": state.get("retry_count", 0) + 1}


def _route_after_cache(state: AgentState) -> str:
    """缓存查询后路由。"""
    cache_state = state.get("cache", {})
    config = state.get("config", {})
    is_kv_model = True if config.get("kv_list") else False

    if is_kv_model:
        kv_state = state.get("kv_state")
        kv_cache = kv_state.get("cache")
        if kv_cache['cache_hit']:
            return "sandbox"
        else:
            return "kv_codegen"

    if cache_state.get("all_cached", False):
        return "sandbox"

    return "code_gen"


def _route_after_sandbox(state: AgentState) -> str:
    is_kv_mode = bool(state.get("config").get("kv_list"))
    if is_kv_mode:
        return "quality"
    return "restore"


def _route_after_quality(state: AgentState) -> str:
    cache_state = state.get("cache", {})
    is_kv_mode = bool(state.get("config").get("kv_list"))

    if cache_state.get("all_cached", False):
        return "end"

    if state["quality_score"] >= QUALITY_THRESHOLD:
        if is_kv_mode:
            return "cache_save"
        return "analyze"

    if state.get("retry_count", 0) >= 3:
        return "analyze"

    return "retry"


# ========================================== #
# 阶段 3：构建有向图 (LangGraph核心)
# ========================================== #

def build_agent():
    g = StateGraph(AgentState)

    g.add_node("parse", parse_node)
    g.add_node("kv_preprocess", kv_preprocess_node)
    g.add_node("kv_codegen", kv_codegen_node)
    g.add_node("pre_structure_analyzer", pre_structure_analyzer_node)
    g.add_node("cache_query", cache_query_node)
    g.add_node("code_gen", code_gen_node)
    g.add_node("sandbox", sandbox_node)
    g.add_node("restore", restore_node)
    g.add_node("quality", quality_node)
    g.add_node("structure_analyzer", structure_analyzer_node)
    g.add_node("cache_save", cache_save_node)
    g.add_node("retry", _retry_node)

    def _route_at_entry(state: AgentState) -> str:
        """入口路由。"""
        config = state.get("config", {})
        if config.get("kv_list"):
            return "kv_preprocess"
        return "parse"

    g.add_conditional_edges("__start__", _route_at_entry, {
        "kv_preprocess": "kv_preprocess",
        "parse": "parse"
    })

    def _route_after_parse(state: AgentState) -> str:
        return "pre_structure_analyzer"

    g.add_conditional_edges("parse", _route_after_parse, {
        "pre_structure_analyzer": "pre_structure_analyzer"
    })

    def _route_after_kv_preprocess(state: AgentState) -> str:
        if state.get("errors"):
            return "end"
        return "cache_query"

    g.add_conditional_edges("kv_preprocess", _route_after_kv_preprocess, {
        "cache_query": "cache_query",
        "end": END
    })

    g.add_edge("kv_codegen", "sandbox")
    g.add_edge("pre_structure_analyzer", "cache_query")

    g.add_conditional_edges("cache_query", _route_after_cache, {
        "sandbox": "sandbox",
        "code_gen": "code_gen",
        "kv_codegen": "kv_codegen"
    })

    g.add_edge("code_gen", "sandbox")

    g.add_conditional_edges("sandbox", _route_after_sandbox, {
        "restore": "restore",
        "quality": "quality"
    })

    g.add_edge("restore", "quality")

    g.add_conditional_edges("quality", _route_after_quality, {
        "analyze": "structure_analyzer",
        "retry": "retry",
        "cache_save": "cache_save",
        "end": END
    })

    g.add_edge("structure_analyzer", "cache_save")
    g.add_edge("cache_save", END)

    g.add_edge("retry", "code_gen")

    return g.compile()


# ========================================== #
# 阶段 4：执行入口 (外部调用)
# ========================================== #

def run_extraction(
        excel_path: str,
        sheet_name: str,
        subtable_titles: Optional[List[str]] = None,
        hints: Optional[str] = None,
        subtable_configs: Optional[Dict[str, Union[SubtableConfig, List[Any]]]] = None,
        kv_list: Optional[List[str]] = None
) -> dict:
    agent = build_agent()

    final = agent.invoke(
        _make_initial(excel_path, sheet_name, subtable_titles, hints, subtable_configs, kv_list)
    )

    return {
        "success": final["quality_score"] >= QUALITY_THRESHOLD,
        "data": final.get("final_output") or final.get("result"),
        "quality_score": final["quality_score"],
        "retry_count": final["retry_count"],
        "errors": final["errors"],
        "generated_code": final.get("generated_code", ""),
        "sandbox_error": final.get("sandbox_error"),
        "all_cached": final.get("cache", {}).get("all_cached", False),
        "partial_cached": final.get("cache", {}).get("partial_cached", False),
    }


async def run_extraction_deep_stream(
        excel_path: str,
        sheet_name: str,
        subtable_titles: Optional[List[str]] = None,
        hints: Optional[str] = None,
        subtable_configs: Optional[Dict[str, Union[SubtableConfig, List[Any]]]] = None,
        kv_list: Optional[List[str]] = None,
) -> AsyncGenerator[dict, None]:
    """
    异步流式入口（调试用）。
    捕获 Token 吐字、节点完成事件。
    """
    agent = build_agent()
    initial = _make_initial(excel_path, sheet_name, subtable_titles, hints, subtable_configs, kv_list)

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

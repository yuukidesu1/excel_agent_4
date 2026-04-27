"""
agent.py — LangGraph 图组装 + 对外入口（支持三级缓存 + KV 模式）

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

  KV 模式流程：
  第一次运行（无缓存）：
  parse_node
      ↓
  cache_query_node (KV 查询) → 未命中
      ↓
  kv_code_gen_node [LLM] → 生成 KV 抽取代码
      ↓
  kv_sandbox_node [KV 模式] → 执行 extract(ws, merged_map, kv_list)
      ↓
  kv_quality_node → 空间置信度打分
      ↓
      ├──────────────┐
      │              │
      v              v
┌────────────┐  ┌──────────────┐
│ 分数 >= 90% │  │ 分数 < 90%   │
└───────┬────┘  └───────┬──────┘
        │               │
        │ Yes           │ No
        │               │
        v               v
┌───────────────┐  ┌────────────────────┐
│ kv_cache_save │  │ retry (max=3)      │
│               │  │ 或人工介入反馈     │
└───────────────┘  └────────────────────┘

  KV 模式后续运行（L1/L2 命中）：
  parse_node
      ↓
  cache_query_node (KV 查询) → 命中
      ↓
  kv_sandbox_node [直接使用缓存代码]
      ↓
  kv_quality_node
      ↓
  END ✅ (跳过 kv_cache_save，避免重复保存)
"""

import sys
import os
from pathlib import Path
from typing import Optional, List, Dict, Any, AsyncGenerator, Union

from langgraph.graph import StateGraph, END

from excel_agent.state import AgentState, SubtableConfig
from excel_agent.nodes.parse import parse_node
from excel_agent.nodes.pre_structure_analyzer import pre_structure_analyzer_node
from excel_agent.nodes.code_gen import code_gen_node
from excel_agent.nodes.sandbox import sandbox_node
from excel_agent.nodes.restore import restore_node
from excel_agent.nodes.quality import quality_node, route, QUALITY_THRESHOLD
from excel_agent.nodes.structure_analyzer import structure_analyzer_node
from excel_agent.nodes.cache_query import cache_query_node
from excel_agent.nodes.cache_save import cache_save_node

# KV 模式节点
from excel_agent.nodes.kv_code_gen import kv_code_gen_node
from excel_agent.nodes.kv_quality import kv_quality_node, route_after_kv_quality, QUALITY_THRESHOLD as KV_QUALITY_THRESHOLD
from excel_agent.nodes.kv_cache_save import kv_cache_save_node
from excel_agent.nodes.kv_sandbox import kv_sandbox_node

# 多场景路由节点
from excel_agent.nodes.kv_scene_router import kv_scene_router_node

# 混合模式结果合并
from excel_agent.nodes.merge_results import merge_results_node

# 开启 LangSmith 追踪
from dotenv import load_dotenv
load_dotenv()
os.environ["LANGCHAIN_TRACING_V2"] = "true"
os.environ["LANGCHAIN_PROJECT"] = "excel_agent"
os.environ["LANGCHAIN_API_KEY"] = os.getenv("LANGCHAIN_API_KEY")


def _retry_node(state: AgentState) -> dict:
    return {"retry_count": state.get("retry_count", 0) + 1}


def _route_after_retry(state: AgentState) -> str:
    """重试后路由：根据抽取类型决定回到哪个代码生成节点"""
    config = state.get("config", {})
    kv_list = config.get("kv_list")
    extract_type = config.get("extract_type", "table")

    is_kv_mode = (kv_list is not None and len(kv_list) > 0) or \
                 (extract_type and extract_type.lower() == "kv")

    if is_kv_mode:
        return "kv_code_gen"

    # 检查是否有 KV 子表需要重试
    cache_state = state.get("cache", {})
    entries = cache_state.get("entries", {})
    has_kv_subtable = any(
        entry.get("extract_mode") == "kv"
        for entry in entries.values()
    )

    # 混合模式或纯 Table 模式：统一回到 code_gen（串行执行 code_gen → kv_code_gen）
    if has_kv_subtable:
        return "code_gen"

    return "code_gen"


def _route_after_cache(state: AgentState) -> str:
    """缓存查询后路由：根据缓存命中情况和场景判定决定"""
    config = state.get("config", {})
    kv_list = config.get("kv_list")
    extract_type = config.get("extract_type", "table")
    cache_state = state.get("cache", {})

    # 1. 全局 KV 模式：有 kv_list 或 extract_type="kv"
    is_global_kv = (kv_list is not None and len(kv_list) > 0) or \
                   (extract_type and extract_type.lower() == "kv")
    if is_global_kv:
        if cache_state.get("hit", False) and not cache_state.get("missed", True):
            return "kv_sandbox"
        return "kv_code_gen"

    # 2. 检查 entries 中的 extract_mode（由 kv_scene_router 写入）
    entries = cache_state.get("entries", {})
    has_kv_subtable = False
    has_table_subtable = False
    has_missed_kv = False
    has_missed_table = False

    for title, entry in entries.items():
        extract_mode = entry.get("extract_mode", "table")
        is_missed = not entry.get("cache_hit", False)
        if extract_mode == "kv":
            has_kv_subtable = True
            if is_missed:
                has_missed_kv = True
        else:
            has_table_subtable = True
            if is_missed:
                has_missed_table = True

    # 3. 纯 Table 模式
    if has_table_subtable and not has_kv_subtable:
        if cache_state.get("all_cached", False):
            return "sandbox"
        return "code_gen"

    # 4. 混合模式（同时存在 KV 和 Table 子表）
    # 改为串行执行：code_gen → kv_code_gen → sandbox
    if has_kv_subtable and has_table_subtable:
        # 有 table 未命中 → 先生成 table 代码
        if has_missed_table:
            return "code_gen"
        # table 全缓存命中，但有 KV 未命中 → 直接生成 KV 代码
        if has_missed_kv:
            return "kv_code_gen"
        # 全部缓存命中
        return "sandbox"

    # 5. 纯 KV 子表（非全局 KV 模式）
    if has_kv_subtable:
        if has_missed_kv:
            return "kv_code_gen"
        return "kv_sandbox"

    # 6. Fallback
    if cache_state.get("all_cached", False):
        return "sandbox"
    return "code_gen"


def _route_after_code_gen(state: AgentState) -> str:
    """code_gen 完成后判断是否还有 KV 代码需要生成（混合模式）"""
    cache_state = state.get("cache", {})
    entries = cache_state.get("entries", {})
    has_missed_kv = any(
        entry.get("extract_mode") == "kv" and not entry.get("cache_hit", False)
        for entry in entries.values()
    )
    if has_missed_kv:
        return "kv_code_gen"
    return "sandbox"


def _route_after_kv_code_gen(state: AgentState) -> str:
    """kv_code_gen 完成后：混合模式 → sandbox，纯 KV → kv_sandbox"""
    cache_state = state.get("cache", {})
    entries = cache_state.get("entries", {})
    has_kv = any(e.get("extract_mode") == "kv" for e in entries.values())
    has_table = any(e.get("extract_mode", "table") != "kv" for e in entries.values())
    if has_kv and has_table:
        return "sandbox"
    return "kv_sandbox"


def _route_after_sandbox(state: AgentState) -> str:
    """sandbox 完成后判断走 restore 还是 merge_results（混合模式）"""
    cache_state = state.get("cache", {})
    entries = cache_state.get("entries", {})
    has_kv = any(e.get("extract_mode") == "kv" for e in entries.values())
    has_table = any(e.get("extract_mode", "table") != "kv" for e in entries.values())
    if has_kv and has_table:
        return "merge_results"
    return "restore"


def _route_after_quality(state: AgentState) -> str:
    """
        质量检查后路由（表格模式）：
        - 全量命中且跑完的 → 直接结束（没产生新代码，无需分析保存）
        - 有新产生代码的 + 质量达标 → structure_analyzer -> cache_save
        - 质量不达标 → retry
        - DEBUG_SKIP_ANALYZER=true → 直接结束（调试用）
        """
    cache_state = state.get("cache", {})

    # 如果所有子表都命中了缓存，直接结束，不走后续的保存流
    if cache_state.get("all_cached", False):
        return "end"

    # 调试模式：跳过 structure_analyzer 和 cache_save，直接结束
    # if DEBUG_SKIP_ANALYZER:
    #     return "end"

    # 以下是有 LLM 新生成代码的情况
    if state["quality_score"] >= QUALITY_THRESHOLD:
        return "analyze"
    if state.get("retry_count", 0) >= 3:
        return "analyze"
    return "retry"


def _route_after_kv_quality(state: AgentState) -> str:
    """KV 质量检查后路由"""
    quality_score = state.get("quality_score", 0.0)
    retry_count = state.get("retry_count", 0)
    cache_state = state.get("cache", {})

    # 如果使用的是缓存代码（缓存命中），无需重复保存，直接结束
    if cache_state.get("hit", False) and not cache_state.get("missed", True):
        return "end"

    # 质量达标，进入缓存保存
    if quality_score >= KV_QUALITY_THRESHOLD:
        return "save"

    # 超过最大重试次数，强制结束（人工介入）
    if retry_count >= 3:
        return "human"

    # 质量不达标，重试
    return "retry"


def build_agent():
    g = StateGraph(AgentState)

    # 1. 注册所有节点（表格模式）
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

    # 注册 KV 模式节点
    g.add_node("kv_code_gen", kv_code_gen_node)
    g.add_node("kv_sandbox", kv_sandbox_node)
    g.add_node("kv_quality", kv_quality_node)
    g.add_node("kv_cache_save", kv_cache_save_node)

    # 注册混合模式结果合并节点
    g.add_node("merge_results", merge_results_node)

    # 注册多场景路由节点
    g.add_node("kv_scene_router", kv_scene_router_node)

    # 2. 定义边 (数据流)
    g.set_entry_point("parse")
    g.add_edge("parse", "pre_structure_analyzer")
    g.add_edge("pre_structure_analyzer", "kv_scene_router")
    g.add_edge("kv_scene_router", "cache_query")

    # 3. 缓存查询后路由：根据抽取类型和缓存命中情况决定
    g.add_conditional_edges("cache_query", _route_after_cache, {
        "sandbox": "sandbox",
        "code_gen": "code_gen",
        "kv_code_gen": "kv_code_gen",
        "kv_sandbox": "kv_sandbox"
    })

    g.add_conditional_edges("code_gen", _route_after_code_gen, {
        "kv_code_gen": "kv_code_gen",
        "sandbox": "sandbox",
    })
    g.add_conditional_edges("kv_code_gen", _route_after_kv_code_gen, {
        "sandbox": "sandbox",
        "kv_sandbox": "kv_sandbox",
    })
    g.add_conditional_edges("sandbox", _route_after_sandbox, {
        "merge_results": "merge_results",
        "restore": "restore",
    })
    g.add_edge("merge_results", "quality")
    g.add_edge("restore", "quality")

    # 4. 质量控制路由：决定重试还是去切分保存代码
    g.add_conditional_edges("quality", _route_after_quality, {
        "analyze": "structure_analyzer",
        "retry": "retry",
        "end": END
    })

    g.add_edge("structure_analyzer", "cache_save")
    g.add_edge("cache_save", END)

    # KV 模式边 (kv_code_gen → kv_sandbox 已在上方的条件路由中处理)
    g.add_edge("kv_sandbox", "kv_quality")

    # KV 质量路由
    g.add_conditional_edges("kv_quality", _route_after_kv_quality, {
        "save": "kv_cache_save",
        "retry": "retry",
        "human": END,  # 人工介入
        "end": END
    })

    g.add_edge("kv_cache_save", END)

    # 重试逻辑：根据抽取类型路由回对应的代码生成节点
    g.add_conditional_edges("retry", _route_after_retry, {
        "code_gen": "code_gen",
        "kv_code_gen": "kv_code_gen"
    })

    return g.compile()


def _make_initial(
    excel_path: str,
    sheet_name: str,
    subtable_titles: List[str],
    hints: Optional[str],
    extract_type: str = "table",
    kv_list: Optional[List[str]] = None,
    subtable_configs: Optional[Dict[str, Union[SubtableConfig, List[Any]]]] = None
) -> AgentState:
    return {
        "config": {
            "excel_path": excel_path,
            "sheet_name": sheet_name,
            "subtable_titles": subtable_titles,
            "hints": hints,
            "subtable_configs": subtable_configs,
            "extract_type": extract_type,
            "kv_list": kv_list or [],
        },
        "sheet_structure": None,
        "generated_code": [],
        "raw_result": None,
        "raw_data": None,
        "kv_result": None,
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
    extract_type: str = "table",
    kv_list: Optional[List[str]] = None,
    subtable_configs: Optional[Dict[str, Union[SubtableConfig, List[Any]]]] = None
) -> dict:
    """
    统一抽取入口（支持表格模式和 KV 模式）

    Args:
        extract_type: "table" | "kv"
        kv_list: KV 模式下的 Key 列表
    """


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
        _make_initial(
            excel_path, sheet_name, subtable_titles, hints,
            extract_type=extract_type,
            kv_list=kv_list,
            subtable_configs=subtable_configs
        )
    )

    # 从 final 状态中读取实际的 extract_type（PSA 可能检测到 KV 布局并修改了它）
    final_extract_type = final.get("config", {}).get("extract_type", extract_type)

    # KV 模式返回
    if final_extract_type == "kv" or kv_list:
        return {
            "success": final["quality_score"] >= KV_QUALITY_THRESHOLD,
            "data": final.get("kv_result"),
            "kv_result": final.get("kv_result"),  # 额外添加，方便 main_yaml.py 读取
            "quality_score": final["quality_score"],
            "retry_count": final["retry_count"],
            "errors": final["errors"],
            "generated_code": final.get("generated_code", ""),
            "sandbox_error": final.get("sandbox_error"),
        }

    # 表格模式返回
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
) -> AsyncGenerator[dict, None]:
    """
    异步流式入口（调试用）。
    捕获 Token 吐字、节点完成事件。
    """
    agent = build_agent()
    initial = _make_initial(excel_path, sheet_name, subtable_titles, hints)

    async for event in agent.astream_events(initial, version="v2"):
        kind = event["event"]
        name = event["name"]

        if kind == "on_chat_model_stream":
            content = event["data"]["chunk"].content
            if content:
                yield {"type": "token", "content": content}

        elif kind == "on_chain_end" and name in (
            "parse", "cache_query", "code_gen", "sandbox",
            "restore", "quality", "structure_analyzer", "cache_save",
            "kv_code_gen", "kv_quality", "kv_cache_save"
        ):
            yield {
                "type": "node_end",
                "name": name,
                "data": event["data"].get("output", {}),
            }

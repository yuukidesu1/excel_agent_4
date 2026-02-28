"""
agent.py — LangGraph 图组装 + 对外调用入口

节点流程：
    START
      ↓
    parse_node    [代码] 读取 Sheet 结构、合并单元格、候选子表块
      ↓
    locate_node   [LLM]  定位子表 → 自动发现所有列 + 识别 forward-fill 列
      ↓
    extract_node  [代码] 按列号精确抽取数据 + 合并单元格填充 + forward-fill
      ↓
    restore_node  [代码] 组装 [表头行] + [数据行] → 二维数组
      ↓
    quality_node  [代码] 质量打分
      ↓
    route ──→ score ≥ 0.75 → END ✅
         └──→ score < 0.75 且 retry < 3 → retry_node → locate_node（带错误上下文）
         └──→ retry ≥ 3 → END（带现有最佳结果）

设计说明：
    - parse_node 只执行一次，IO 不重复
    - 重试从 locate_node 开始，携带 errors 让 LLM 自我修正
    - 纯代码节点（parse/extract/restore）不参与重试循环
"""

from langgraph.graph import StateGraph, END

from excel_agent.state       import AgentState
from excel_agent.nodes.parse   import parse_node
from excel_agent.nodes.locate  import locate_node
from excel_agent.nodes.extract import extract_node
from excel_agent.nodes.restore import restore_node
from excel_agent.nodes.quality import quality_node, route


def _retry_node(state: AgentState) -> dict:
    """重试节点：仅递增计数，errors 已在 quality_node 中追加"""
    return {"retry_count": state.get("retry_count", 0) + 1}


def build_agent():
    """构建并编译 LangGraph Agent"""
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
    g.add_edge("retry",   "locate")   # 重试回到 locate（跳过 parse）

    g.add_conditional_edges("quality", route, {"end": END, "retry": "retry"})

    return g.compile()


def run_extraction(excel_path: str, sheet_name: str, subtable_title: str,
                   hints: str = None) -> dict:
    """
    对外统一调用入口。

    参数：
        excel_path     : Excel 文件路径
        sheet_name     : 目标 Sheet 名，如 "CONFIGURATION"
        subtable_title : 子表标题关键词，如 "3G Configuration"
        hints          : 可选的额外定位提示

    返回：
        {
          "success":       bool,            # quality_score >= 0.75
          "data":          List[List[str]], # 二维数组，第 0 行为列名
          "quality_score": float,
          "retry_count":   int,
          "errors":        List[str],
        }

    ──────────────────────────────────────────────────────────
    使用示例：

        from excel_agent import run_extraction
        import json

        result = run_extraction(
            excel_path     = "./project.xlsx",
            sheet_name     = "CONFIGURATION",
            subtable_title = "3G Configuration",
        )

        if result["success"]:
            print(json.dumps(result["data"], ensure_ascii=False, indent=2))
        else:
            print("抽取质量不达标：", result["errors"])
    """
    agent = build_agent()

    initial: AgentState = {
        "excel_path": excel_path,
        "config": {
            "sheet_name":     sheet_name,
            "subtable_title": subtable_title,
            "hints":          hints,
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

    final = agent.invoke(initial)

    return {
        "success":       final["quality_score"] >= QUALITY_THRESHOLD,
        "data":          final.get("final_output") or final.get("result"),
        "quality_score": final["quality_score"],
        "retry_count":   final["retry_count"],
        "errors":        final["errors"],
    }

# 从 quality.py 导入阈值常量，供 run_extraction 使用
from excel_agent.nodes.quality import QUALITY_THRESHOLD

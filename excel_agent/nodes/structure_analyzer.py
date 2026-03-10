"""
nodes/structure_analyzer.py — 结构指纹提取节点

职责：
    当代码质量足够高时（quality_score >= 0.8），调用 LLM 分析生成的代码，
    提取子表的结构信息（起始行、列定义、表头结构等），生成结构指纹用于 L2 缓存。

结构指纹格式：
    {
        "sheet_name": "CONFIGURATION",
        "subtables": [
            {
                "title": "4G Configuration",
                "title_pattern": "4gconfiguration",       # 归一化标题
                "start_row": 13,                          # 绝对起始行
                "end_row": 25,
                "start_col": 2,
                "end_col": 15,
                "header_structure": [                      # 表头结构
                    {"row": 0, "cols": [                  # 第 0 行（相对行号）
                        {"col": 0, "name": "systemmodule", "span": 1},
                        {"col": 1, "name": "cell", "span": 1},
                        {"col": 2, "name": "rfmodule||type", "span": 4},  # 合并单元格
                    ]},
                    {"row": 1, "cols": [...]}              # 第 1 行（双行表头）
                ],
                "row_count": 10,                           # 数据行数（不含表头）
                "col_count": 14,                           # 总列数
            }
        ],
        "merge_patterns": [                                # 合并模式（相对坐标）
            {"rel_row": 0, "rel_col": 2, "col_span": 4},  # 第 0 行第 2 列开始，跨 4 列
            {"rel_row": 1, "rel_col": 0, "row_span": 5},  # 第 1 行第 0 列开始，跨 5 行
        ]
    }

输出放入 state["structure_fingerprint"]，供缓存系统使用。
"""

import re
import json
import os
from pathlib import Path
from typing import Dict, List, Any, Optional, Tuple

from langchain_openai import ChatOpenAI
from langchain_core.messages import SystemMessage, HumanMessage

from excel_agent.state import AgentState


_SYSTEM_PROMPT = """\
你是 Excel 表格结构分析专家。你的任务是从已生成的 Python 抽取代码中逆向分析出表格的结构指纹。

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
输入信息：
  - generated_code   : LLM 生成的完整抽取代码
  - sheet_structure  : Sheet 的原始结构信息（辅助参考）
  - subtable_titles  : 目标子表标题列表

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
你的任务：
  1. 从代码中提取每个子表的精确位置（起始行、结束行、起始列、结束列）
  2. 分析表头结构（单行/双行表头、列名、合并单元格模式）
  3. 提取数据行的范围（起始行、结束行）
  4. 生成标准化的结构指纹

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
输出格式（严格返回以下 JSON，不要任何解释）：
```json
{
  "subtables": [
    {
      "title": "原始子表名",
      "title_pattern": "归一化标题（小写、无空格、无特殊字符）",
      "start_row": 13,
      "end_row": 25,
      "start_col": 2,
      "end_col": 15,
      "header_rows": 2,
      "header_structure": {
        "row_0": {"cols": [{"col": 0, "name": "systemmodule"}, {"col": 1, "name": "cell"}]},
        "row_1": {"cols": [{"col": 2, "name": "rfmodule||type"}]}
      },
      "data_row_start": 15,
      "data_row_end": 24,
      "row_count": 10,
      "col_count": 14
    }
  ],
  "merge_patterns": [
    {"rel_row": 0, "rel_col": 2, "col_span": 4, "value": "RF MODULE"},
    {"rel_row": 1, "rel_col": 0, "row_span": 2, "value": "CELL"}
  ]
}
```

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
分析要点：
  1. 从代码中的 range(X, Y) 提取行范围
  2. 从 cell_val(r, c) 提取列范围
  3. 从 header_row1/header_row2 判断是否双行表头
  4. 从 merged_map 的使用或注释推断合并模式
  5. 标题归一化：转小写、去空格、去标点 → "4gconfiguration"

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
注意事项：
  - 所有行号、列号必须是代码中实际使用的硬编码值
  - header_structure 中的 col 是相对于 start_col 的偏移
  - merge_patterns 中的 rel_row/rel_col 是相对于子表起始位置的偏移
  - 如果代码中明确有注释说明合并单元格，优先使用注释信息
"""


def _get_llm() -> ChatOpenAI:
    """获取 LLM 实例（复用 code_gen 的配置）"""
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


def _normalize_title(title: str) -> str:
    """归一化标题：小写、去空格、去标点"""
    # 转小写
    s = title.lower()
    # 去除所有空白字符
    s = re.sub(r'\s+', '', s)
    # 去除常见标点
    s = re.sub(r'[^\w\u4e00-\u9fff]', '', s)
    return s


def _extract_json(raw: str) -> Optional[Dict]:
    """从 LLM 输出中提取 JSON"""
    # 尝试匹配 ```json ... ``` 块
    m = re.search(r'```json\s*([\s\S]*?)```', raw)
    if m:
        try:
            return json.loads(m.group(1).strip())
        except json.JSONDecodeError:
            pass

    # 尝试匹配 ``` ... ``` 块
    m = re.search(r'```\s*([\s\S]*?)```', raw)
    if m:
        try:
            return json.loads(m.group(1).strip())
        except json.JSONDecodeError:
            pass

    # 尝试直接解析整个输出
    try:
        return json.loads(raw.strip())
    except json.JSONDecodeError:
        return None


def _analyze_code_structure(code: str) -> Dict[str, Any]:
    """
    从代码中提取结构信息的辅助函数（规则-based）
    用于预分析，帮助 LLM 更准确地理解代码
    """
    info = {
        "range_patterns": [],
        "cell_val_patterns": [],
        "header_patterns": [],
        "merge_comments": [],
    }

    # 提取所有 range(X, Y) 模式
    for m in re.finditer(r'range\((\d+),\s*(\d+)\)', code):
        info["range_patterns"].append({
            "start": int(m.group(1)),
            "end": int(m.group(2)),
            "context": code[max(0, m.start()-50):m.end()+50]
        })

    # 提取所有 cell_val(r, c) 模式
    for m in re.finditer(r'cell_val\((\d+),\s*(\d+)\)', code):
        info["cell_val_patterns"].append({
            "row": int(m.group(1)),
            "col": int(m.group(2)),
            "context": code[max(0, m.start()-30):m.end()+30]
        })

    # 提取表头相关注释
    for m in re.finditer(r'#.*表头|#.*header', code, re.IGNORECASE):
        info["header_patterns"].append(code[m.start():code.find('\n', m.end())])

    # 提取合并单元格相关注释
    for m in re.finditer(r'#.*合并|#.*merge', code, re.IGNORECASE):
        info["merge_comments"].append(code[m.start():code.find('\n', m.end())])

    return info


def structure_analyzer_node(state: AgentState) -> dict:
    """
    结构指纹提取节点

    输入：
        - generated_code    : 已生成的抽取代码
        - sheet_structure   : Sheet 结构信息
        - quality_score     : 质量评分（>= 0.8 才执行）
        - subtable_titles   : 子表标题列表

    输出：
        - structure_fingerprint : 结构指纹字典
        - analyzer_skipped    : 是否跳过分析（质量不足时）
    """
    code = state.get("generated_code", "")
    quality_score = state.get("quality_score", 0.0)

    config = state.get("config", {})
    subtable_titles = config.get("subtable_titles") or state.get("subtable_titles", [])

    # 质量不足，跳过分析（不浪费 Token）
    if quality_score < 0.8:
        return {
            "structure_fingerprint": None,
            "analyzer_skipped": True,
            "skip_reason": f"质量评分 {quality_score} < 0.8，跳过结构分析"
        }

    if not code:
        return {
            "structure_fingerprint": None,
            "analyzer_skipped": True,
            "skip_reason": "generated_code 为空"
        }

    # 预分析代码结构
    code_analysis = _analyze_code_structure(code)

    st = state.get("sheet_structure", {})

    # 组装给 LLM 的上下文
    ctx = {
        "generated_code": code,
        "subtable_titles": subtable_titles,
        "sheet_structure": {
            "sheet_name": st.get("sheet_name", ""),
            "max_row": st.get("max_row", 0),
            "max_col": st.get("max_col", 0),
        },
        "code_analysis": {
            "range_patterns": [
                {"start": p["start"], "end": p["end"], "context": p["context"][:100]}
                for p in code_analysis["range_patterns"][:20]
            ],
            "cell_val_patterns": [
                {"row": p["row"], "col": p["col"], "context": p["context"][:80]}
                for p in code_analysis["cell_val_patterns"][:30]
            ],
            "header_comments": code_analysis["header_patterns"][:5],
            "merge_comments": code_analysis["merge_comments"][:5],
        }
    }

    messages = [
        SystemMessage(content=_SYSTEM_PROMPT),
        HumanMessage(content=json.dumps(ctx, ensure_ascii=False, indent=2)),
    ]

    try:
        response = _get_llm().invoke(messages)
        parsed = _extract_json(response.content)

        if parsed is None:
            return {
                "structure_fingerprint": None,
                "analyzer_skipped": True,
                "skip_reason": "LLM 返回格式无法解析为 JSON",
                "llm_raw_output": response.content[:500] if response.content else ""
            }

        # 添加归一化标题
        if "subtables" in parsed:
            for sub in parsed["subtables"]:
                if "title" in sub and "title_pattern" not in sub:
                    sub["title_pattern"] = _normalize_title(sub["title"])

        # 添加元数据
        fingerprint = {
            **parsed,
            "_meta": {
                "quality_score": quality_score,
                "created_from_code": True,
                "sheet_name": st.get("sheet_name", "")
            }
        }

        return {
            "structure_fingerprint": fingerprint,
            "analyzer_skipped": False
        }

    except Exception as e:
        return {
            "structure_fingerprint": None,
            "analyzer_skipped": True,
            "skip_reason": f"分析过程异常：{str(e)}"
        }


def compute_structure_signature(fingerprint: Dict[str, Any]) -> str:
    """
    从结构指纹计算哈希签名（用于 L2 缓存 Key）

    签名应满足：
    - 相同的表格结构 → 相同的签名
    - 不同的表格结构 → 不同的签名
    - 与绝对位置无关（只关心相对结构）
    """
    import hashlib

    if not fingerprint:
        return ""

    # 提取与结构相关的核心特征（排除绝对位置）
    signature_data = {
        "sheet_name": fingerprint.get("_meta", {}).get("sheet_name", ""),
        "subtables": []
    }

    for sub in fingerprint.get("subtables", []):
        sig_sub = {
            "title_pattern": sub.get("title_pattern", ""),
            "header_rows": sub.get("header_rows", 1),
            "row_count": sub.get("row_count", 0),
            "col_count": sub.get("col_count", 0),
            "header_structure": sub.get("header_structure", {}),
        }
        signature_data["subtables"].append(sig_sub)

    # 添加合并模式
    signature_data["merge_patterns"] = fingerprint.get("merge_patterns", [])

    # 计算哈希
    serialized = json.dumps(signature_data, sort_keys=True, ensure_ascii=False)
    hash_value = hashlib.sha256(serialized.encode()).hexdigest()[:16]

    return hash_value

"""
nodes/kv_code_gen.py — LLM 节点：生成 KV 抽取代码

LLM 的任务：
    1. 理解 sheet_structure（非空单元格、合并单元格信息）
    2. 基于 kv_list 中的 Key 文本，定位 Key 在 Excel 中的位置
    3. 编写 Python 抽取函数，使用 openpyxl 读取数据
    4. 函数必须返回 Dict[str, str]（KV 字典）

核心规则（红线）：
    1. 强制相对寻址：必须先定位 Key 的坐标，再基于该坐标游走寻找 Value
    2. 禁止绝对坐标：不能写死 ws.cell(5, 3).value 这样的硬编码
    3. 支持合并单元格：Key 定位时需考虑合并区域，使用 Bounding Box 计算
    4. 空间扫描优先级：同行向右 > 同列向下 > 其他方向
"""

import json
import re
import os
from pathlib import Path

from langchain_openai import ChatOpenAI
from langchain_core.messages import SystemMessage, HumanMessage

from excel_agent.state import AgentState

_SYSTEM_PROMPT = """
你是顶尖的 Excel KV（键值对）数据抽取专家。请根据传入的结构视图和 Key 列表，编写 Python 代码抽取 KV 数据。

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
【输入上下文信息】
  - kv_list          : 待抽取的 Key 列表（如 ["Site Power Supply", "Rectifier Current Reading(A)"]）
  - sheet_structure  : 包含非空/合并单元格的坐标与值
  - last_code_error  : 上次代码执行错误信息（重试时才有）

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
【输出与红线规则】
你必须编写一个名为 `extract(ws, merged_map, kv_list) -> dict` 的函数。
返回值必须是 Dict[str, str]，键为 Key 原文，值为抽取到的 Value。

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
【核心红线规则（违反将导致系统崩溃）】

1. 禁止导入模块（极度重要！）
   - 沙盒环境已经预置了 re, math, json, openpyxl 等模块
   - 绝对不能写 import xxx 或 from xxx import xxx
   - 直接使用 re.sub() 等函数即可

2. 强制相对寻址（极度重要！）
   - 必须先定位 Key 所在的单元格坐标 (key_row, key_col)
   - 然后基于该坐标游走寻找 Value（如向右扫描、向下扫描）
   - 绝对不能写死绝对坐标（如 ws.cell(5, 3).value 或 start_row = 10）

3. 支持合并单元格
   - Key 可能是一个合并单元格（如占据 A1:A2）
   - 定位 Key 时，需要获取其合并区域的边界（Bounding Box）
   - 扫描 Value 时，从 Key 的边界外延开始扫描

4. 空间扫描优先级
   - 优先向右扫描（同行）：**不限扫描距离**，直到行末
   - 其次向下扫描（同列）：**不限扫描距离**，直到表格底部
   - 如果都找不到，返回空字符串
   - 注意：不能自己添加"5 列"或"10 行"的限制，必须扫描到边界

5. 单元格读取规范
   - 必须通过 `merged_map.get((r, c), ws.cell(row=r, column=c).value)` 读取
   - 处理 None 值、数字转字符串、去除空白

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
【代码骨架模板 (请严格参考此范式)】

```python
def extract(ws, merged_map, kv_list):
    def cell_val(r, c):
        # 安全读取单元格值
        v = merged_map.get((r, c), ws.cell(row=r, column=c).value)
        if v is None:
            return ""
        if isinstance(v, float) and v == int(v):
            return str(int(v))
        return str(v).replace('\n', ' ').replace('\r', '').strip()

    def get_merged_box(r, c):
        # 获取单元格所属的合并区域边界
        # 通过 ws.merged_cells.rangs 获取合并单元格边界
        for rng in ws.merged_cells.ranges:
            if rng.min_row <= r <= rng.max_row and rng.min_col <= c <= rng.max_col:
                return {
                    "min_row": rng.min_row,
                    "max_row": rng.max_row,
                    "min_col": rng.min_col,
                    "max_col": rng.max_col
                }
        
        # 非合并单元格，返回自身
        return {"min_row": r, "max_row": r, "min_col": c, "max_col": c}

    def _normalize(text):
        # 归一化文本用于模糊匹配
        if text is None:
            return ""
        return re.sub(r'[\s\-\_\(\)\[\]\/\.,]+', '', str(text).lower())

    def find_key_cell(key_text):
        # 定位 Key 所在单元格
        # Returns: (key_row, key_col, key_box) 或 (None, None, None)
        key_norm = _normalize(key_text)

        # 全局扫描寻找 Key
        for r in range(1, ws.max_row + 1):
            for c in range(1, ws.max_column + 1):
                # 直接使用 cell_val 函数读取，它会处理 merged_map
                val = cell_val(r, c)
                if val and key_norm in _normalize(val):
                    box = get_merged_box(r, c)
                    return r, c, box

        return None, None, None

    def scan_value(key_r, key_c, key_box):
        # 空间扫描 Value
        # 优先级：1.从 Key 右边界向右扫描  2.从 Key 下边界向下扫描
        # 1. 向右扫描
        start_c = key_box["max_col"] + 1
        for c in range(start_c, ws.max_column + 1):
            val = cell_val(key_r, c)
            if val:
                return val, (key_r, c)

        # 2. 向下扫描
        start_r = key_box["max_row"] + 1
        for r in range(start_r, ws.max_row + 1):
            val = cell_val(r, key_c)
            if val:
                return val, (r, key_c)

        return "", None

    # ==================== 主逻辑 ====================
    result = {}

    for key in kv_list:
        key_r, key_c, key_box = find_key_cell(key)

        if key_r is None:
            # Key 未找到
            result[key] = ""
            continue

        val, val_pos = scan_value(key_r, key_c, key_box)
        result[key] = val

    return result
```

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
【常见错误示例（请避免）】

错误：写死绝对坐标
```python
start_row = 10  # 错误！不能硬编码
for r in range(start_row, ws.max_row + 1):
    ...
```

正确：动态定位
```python
key_r, key_c, key_box = find_key_cell(key)
if key_r is not None:
    val = cell_val(key_r, key_c + 1)  # 基于 Key 坐标游走
```

错误：忽略合并单元格
```python
val = cell_val(key_r, key_c + 1)  # 如果 Key 是合并单元格，可能越界
```

正确：使用 Bounding Box
```python
box = get_merged_box(key_r, key_c)
val = cell_val(key_r, box["max_col"] + 1)  # 从合并区域右边界外延
```
"""


def _get_llm() -> ChatOpenAI:
    from dotenv import load_dotenv
    for parent in Path(__file__).resolve().parents:
        env_file = parent / ".env"
        if env_file.exists():
            load_dotenv(env_file, override=True)
            break
    api_key = os.getenv("OPENAI_API_KEY")
    base_url = os.getenv("OPENAI_BASE_URL") or None
    model = os.getenv("LLM_MODEL", "glm-4.7")
    if not api_key:
        raise ValueError("未找到 OPENAI_API_KEY，请检查 .env 文件。")
    return ChatOpenAI(model=model, temperature=0, api_key=api_key, base_url=base_url)


def _extract_code(raw: str) -> str:
    """从 LLM 输出中提取 python ... 之间的代码"""
    if not raw:
        return ""
    # 先尝试提取 ```python ... ``` 格式
    m = re.search(r"```python\s*([\s\S]*?)```", raw)
    if m:
        return m.group(1).strip()
    # 再尝试提取 ``` ... ``` 格式
    m = re.search(r"```\s*([\s\S]*?)```", raw)
    if m:
        return m.group(1).strip()
    # 如果没有代码块标记，尝试提取 def extract 开头的函数
    m = re.search(r"def extract\([^)]*\)\s*(?::[^]*?)(?=^\s*return|$\n)", raw, re.MULTILINE)
    if m:
        # 提取整个函数
        func_match = re.search(r"(def extract[\s\S]*)", raw)
        if func_match:
            return func_match.group(1).strip()
    return raw.strip()


def kv_code_gen_node(state: AgentState) -> dict:
    """KV 模式代码生成节点"""
    st = state.get("sheet_structure", {})
    errors = state.get("errors", [])
    sandbox_error = state.get("sandbox_error")

    config = state.get("config", {})
    kv_list = config.get("kv_list", [])

    # 极限压缩 Token 优化逻辑
    non_empty_cells = st.get("non_empty_cells", [])
    merged_cells_info = st.get("merged_cells_info", [])

    # 压缩单元格信息（限制数量）
    compressed_cells = [
        f"R{c['row']}C{c['col']}:{c['value']}"
        for c in non_empty_cells
    ][:300]

    compressed_merges = [
        f"R{m['min_row']}C{m['min_col']}~R{m['max_row']}C{m['max_col']}:{m['value']}"
        for m in merged_cells_info
    ][:150]

    # 组装上下文
    ctx = {
        "kv_list": kv_list,
        "sheet_structure": {
            "sheet_name": st.get("sheet_name", ""),
            "max_row": st.get("max_row", 1000),
            "max_col": st.get("max_col", 100),
            "sample_cells": compressed_cells,
            "merged_cells": compressed_merges
        }
    }

    # 重试时附上错误
    if sandbox_error:
        ctx["last_code_error"] = sandbox_error
        ctx["retry_instruction"] = (
            "上次生成的代码执行时出错，请仔细阅读错误信息并修正代码！\n"
            f"错误详情：{sandbox_error}"
        )
    elif errors:
        ctx["last_quality_errors"] = errors[-3:]
        ctx["retry_instruction"] = (
            "上次代码执行成功但质量不达标，请根据质量报错修正代码。\n"
            "常见错误：Key 定位错误、Value 扫描方向错误、未处理合并单元格等。"
        )

    messages = [
        SystemMessage(content=_SYSTEM_PROMPT),
        HumanMessage(content=json.dumps(ctx, ensure_ascii=False, indent=2))
    ]

    # ———————————————————— DEBUG ——————————————————————————————————————————
    response = _get_llm().invoke(messages)

    code = _extract_code(response.content) if response else ""


    # ———————————————————— DEBUG ——————————————————————————————————————————
#     code = """def extract(ws, merged_map, kv_list):
#     def cell_val(r, c):
#         # 安全读取单元格值
#         v = merged_map.get((r, c), ws.cell(row=r, column=c).value)
#         if v is None:
#             return ""
#         if isinstance(v, float) and v == int(v):
#             return str(int(v))
#         return str(v).replace('\\n', ' ').replace('\\r', '').strip()
#
#     def get_merged_box(r, c):
#         # 通过 ws.merged_cells.ranges 获取合并单元格边界
#         for rng in ws.merged_cells.ranges:
#             if rng.min_row <= r <= rng.max_row and rng.min_col <= c <= rng.max_col:
#                 return {
#                     "min_row": rng.min_row,
#                     "max_row": rng.max_row,
#                     "min_col": rng.min_col,
#                     "max_col": rng.max_col
#                 }
#         # 非合并单元格，返回自身
#         return {"min_row": r, "max_row": r, "min_col": c, "max_col": c}
#
#     def _normalize(text):
#         # 归一化文本用于模糊匹配
#         if text is None:
#             return ""
#         return re.sub(r'[\\s\\-\\_\\(\\)\\[\\]\\/\\.,]+', '', str(text).lower())
#
#     def find_key_cell(key_text):
#         # 定位 Key 所在单元格
#         # Returns: (key_row, key_col, key_box) 或 (None, None, None)
#         key_norm = _normalize(key_text)
#
#         # 全局扫描寻找 Key
#         for r in range(1, ws.max_row + 1):
#             for c in range(1, ws.max_column + 1):
#                 # 直接使用 cell_val 函数读取，它会处理 merged_map
#                 val = cell_val(r, c)
#                 if val and key_norm in _normalize(val):
#                     box = get_merged_box(r, c)
#                     return r, c, box
#
#         return None, None, None
#
#     def scan_value(key_r, key_c, key_box):
#         # 空间扫描 Value
#         # 优先级：1.从 Key 右边界向右扫描  2.从 Key 下边界向下扫描
#         # 1. 向右扫描
#         start_c = key_box["max_col"] + 1
#         for c in range(start_c, ws.max_column + 1):
#             val = cell_val(key_r, c)
#             if val:
#                 return val, (key_r, c)
#
#         # 2. 向下扫描
#         start_r = key_box["max_row"] + 1
#         for r in range(start_r, ws.max_row + 1):
#             val = cell_val(r, key_c)
#             if val:
#                 return val, (r, key_c)
#
#         return "", None
#
#     # ==================== 主逻辑 ====================
#     result = {}
#
#     for key in kv_list:
#         key_r, key_c, key_box = find_key_cell(key)
#
#         if key_r is None:
#             # Key 未找到
#             result[key] = ""
#             continue
#
#         val, val_pos = scan_value(key_r, key_c, key_box)
#         result[key] = val
#
#     return result
# """

    return {"generated_code": [code], "sandbox_error": None}
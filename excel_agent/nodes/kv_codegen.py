"""
nodes/kv_codegen.py — KV 独立流水线：代码生成节点

职责：
    1. 接收 kv_preprocess_node 产出的高密度网格视图 (sample_cells)。
    2. 基于 _KV_SYSTEM_PROMPT，让 LLM 根据用户的 kv_list 编写相对坐标偏移提取代码。
    3. 支持自愈循环：接收 sandbox_error 或 quality_errors，让 LLM 进行反思重试。
"""

import os
import re
import json
import ast
import requests
from langchain_core.runnables import RunnableLambda

from langchain_openai import ChatOpenAI
from langchain_core.messages import SystemMessage, HumanMessage, AIMessage

from excel_agent.state import AgentState
_KV_SYSTEM_PROMPT = """\
你是顶尖的 Excel KV（键值对）数据抽取专家。请根据传入的结构视图，编写 Python 代码抽取离散的 KV 数据。

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
【输入上下文信息】
  - kv_configs       : 用户配置的待抽取键值对。格式可能包含层级，如 `["Global Key", "Region A||Key 1"]`
  - sheet_structure  : 包含真实的排版网格以及预先计算好的各区域物理边界 `bboxes`。

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
【输出与红线规则】
你必须编写一个名为 `extract(ws, merged_map: dict, bboxes: dict) -> dict` 的函数。
最终返回值必须是【一维扁平字典 (Flat Dict)】。键的名称必须与 `kv_configs` 中配置的字符串一模一样（包括 `||` 符号）。

核心红线规则：
1. 依赖动态 Bbox 防串台：系统已在 `bboxes` 中传入了各表头的物理边界。你必须使用 `get_bbox("表头名")` 获取边界，并将 min_r, max_r, min_c, max_c 传入 `find_anchors`，实现物理隔离！
2. 动态扫描寻找 Value 防空白隔离：表格中常有为了排版留出的“空白隔离列/行”，绝对禁止硬编码偏移量（如直接取 `max_col + 1`）！寻找 Value 必须严格观察 `sample_cells` 决定方向，并使用骨架中的 `search_right` 或 `search_down` 进行穿透扫描。
   - 【横向平铺表单】：如果子 Key 和区域表头同处一行，Value 必定在右侧，必须使用 `search_right`，禁止向下找！
3. 必须使用候选遍历法防假表头：严禁直接写 `anchors[0]`！必须使用 `for r, c, box in find_anchors(...):` 遍历，并判断 `if val:`（非空）才赋值并 `break`。
4. 智能语义与容错扩展 (Fuzzy Expansion)：全局 Key 以及局部场景的【最末端子 Key】经常存在笔误缩写，你必须发挥常识，传入包含常见拼写错误、同义词的列表（例如扩展写为 `["Azimuth", "Azmuith", "Ant Azimuth"]`）。注意：绝对禁止对【区域表头】（如 `Cell B`）使用扩展列表，表头必须原样单字符串查找！

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
【KV 抽取代码骨架范式 (请严格基于此范式填空与组装)】

```python
# 后续节点已定义好沙箱，这里不要进行任何 import
def extract(ws, merged_map: dict, bboxes: dict) -> dict:
    def cell_val(r, c):
        info = merged_map.get((r, c))
        v = info["value"] if isinstance(info, dict) else ws.cell(row=r, column=c).value
        if v is None: return ""
        if isinstance(v, float) and v == int(v): v = str(int(v))
        return " ".join(str(v).splitlines()).strip()

    def get_merged_box(r, c):
        info = merged_map.get((r, c))
        if isinstance(info, dict):
            return {
                "min_row": info["min_row"],
                "max_row": info["max_row"],
                "min_col": info["min_col"],
                "max_col": info["max_col"]
            }
        return {"min_row": r, "max_row": r, "min_col": c, "max_col": c}

    def _normalize(text):
        if text is None: return ""
        return re.sub(r'[\\s\\-\\_\\(\\)\\[\\]\\/\\.,]+', '', str(text).lower())

    def get_bbox(region_name):
        '''安全获取边界：忽略大小写匹配字典 Key'''
        if not region_name: return {}
        target = str(region_name).strip().lower()
        for k, v in bboxes.items():
            if str(k).strip().lower() == target:
                return v
        return {}

    def find_anchors(keywords, min_r=None, max_r=None, min_c=None, max_c=None):
        '''在指定的二维边界盒内，寻找匹配的关键词（支持同义词容错列表）'''
        if min_r is None: min_r = 1
        if max_r is None: max_r = ws.max_row
        if min_c is None: min_c = 1
        if max_c is None: max_c = ws.max_column

        if isinstance(keywords, str): keywords = [keywords]
        targets = [_normalize(k) for k in keywords]

        matches = []
        for r in range(min_r, max_r + 1):
            for c in range(min_c, max_c + 1):
                cell_norm = _normalize(cell_val(r, c))
                if any(t in cell_norm for t in targets):
                    matches.append((r, c, get_merged_box(r, c)))
        return matches

    def search_right(r, start_col, max_col):
        '''向右射线扫描，自动穿透空白排版列，寻找第一个非空值'''
        for c in range(start_col, max_col + 1):
            val = cell_val(r, c)
            if val: return val
        return ""

    def search_down(c, start_row, max_row):
        '''向下射线扫描，自动穿透空白排版行，寻找第一个非空值'''
        for r in range(start_row, max_row + 1):
            val = cell_val(r, c)
            if val: return val
        return ""

    result = {}

    # =========================================================
    # 下方由 LLM 根据配置 (kv_configs) 和观察到的上下文进行编写
    # =========================================================

    # 1. 处理全局 Key (最末端，使用 List 容错扩展)
    result["Global Key"] = ""
    for r, c, box in find_anchors(["Global Key", "GlobalKey", "Globl Key"]):
        val = search_right(r, box["max_col"] + 1, ws.max_column)
        if val:
            result["Global Key"] = val
            break

    # 2. 处理局部 Key (表头严格匹配，子键 List 容错扩展)
    result["Cell B||Azimuth"] = ""

    cell_b_box = get_bbox("Cell B")  # 区域表头，传入单字符串严格匹配取框
    if cell_b_box:
        # 最末端子 Key (Azimuth)，传入包含错别字的 List 容错扩展
        for key_r, key_c, key_box in find_anchors(["Azimuth", "Azmuith", "Antenna Azimuth", "Ant Azimuth"], 
                                                  min_r=cell_b_box.get("min_r"), 
                                                  max_r=cell_b_box.get("max_r"), 
                                                  min_c=cell_b_box.get("min_c"), 
                                                  max_c=cell_b_box.get("max_c")):

            # 使用射线扫描替代 +1，穿透空白列，且不超越当前大区域的右边界
            val = search_right(key_r, key_box["max_col"] + 1, cell_b_box.get("max_c", ws.max_column))
            if val:
                result["Cell B||Azimuth"] = val
                break

    # 3. 处理横向平铺局部 Key (判定逻辑示例)
    result["Horizontal Table||Leg 1"] = ""
    ht_box = get_bbox("Horizontal Table")
    if ht_box:
        anchor_r = ht_box.get("min_r")

        for key_r, key_c, key_box in find_anchors(["Leg 1", "Leg1", "L1"], 
                                                  min_r=ht_box.get("min_r"), 
                                                  max_r=ht_box.get("max_r"), 
                                                  min_c=ht_box.get("min_c"), 
                                                  max_c=ht_box.get("max_c")):
            if key_r == anchor_r:
                val = search_right(key_r, key_box["max_col"] + 1, ht_box.get("max_c", ws.max_column)) # 同行强制向右
            else:
                # 根据网格自行判断，如若是上下结构可写: val = search_down(key_box["min_col"], key_box["max_row"] + 1, ht_box.get("max_r", ws.max_row))
                val = search_right(key_r, key_box["max_col"] + 1, ht_box.get("max_c", ws.max_column))

            if val:
                result["Horizontal Table||Leg 1"] = val
                break

    return result
    """

# =====================================================================
# 【用户提示词模板】
# 将解析后的物理矩阵直接喂给 LLM，让其观察相对偏移量
# =====================================================================
_USER_PROMPT_TEMPLATE = '''请为以下 KV 提取任务生成完整的 Python 代码：

【KV 配置列表 (kv_configs)】
{kv_configs_str}

【Sheet 结构信息】
- 最大行数：{max_row}
- 最大列数：{max_col}

【预计算区域边界 (Bboxes)】
以下是系统预先为你框定好的各区域绝对安全边界（字典的 Key 已自动转为小写）。
在提取局部区域的子 Key 时，你必须通过 `bboxes.get("表头名")` 获取边界，并将 min_r, max_r, min_c, max_c 传入 `find_anchors` 中，实现物理隔离防串台！
{bboxes_info}

【合并单元格分布】
{merged_cells_info}

【局部排版网格 (Sample Cells)】
请仔细观察下方非空单元格的空间坐标（R行C列）。
你必须通过这些真实坐标，判断 Value 与 Key 的位置关系（同行向右，或在下方），并直接在代码中写死读取坐标。
{sample_cells}

{retry_instruction}
'''


def _init_env():
    import os
    from pathlib import Path
    from dotenv import load_dotenv

    os.environ["http_proxy"] = ""
    os.environ["https_proxy"] = ""
    os.environ["all_proxy"] = ""

    for parent in Path(__file__).resolve().parents:
        env_file = parent / ".env"
        if env_file.exists():
            load_dotenv(env_file, override=True)
            break

def _get_llm(dynamic_token: str = None) -> ChatOpenAI:
    _init_env()

    api_key = os.getenv("OPENAI_API_KEY")
    base_url = os.getenv("OPENAI_BASE_URL") or "http://141.246.3.57:20004/v1"
    model = os.getenv("LLM_MODEL2", "Qwen-V3_6-27B")
    header_name = os.getenv("Authorization")
    extra_body_params = {
        "thinking": False,
        "enable_thinking": False
    }

    custom_headers = {}
    if header_name and dynamic_token:
        custom_headers[header_name] = dynamic_token

    if not api_key:
        raise ValueError("未找到 OPENAI_API_KEY，请检查 .env 文件。")

    return ChatOpenAI(
        model=model,
        temperature=0.0,
        base_url=base_url,
        default_headers=custom_headers,
        timeout=300,
        model_kwargs={"extra_body": extra_body_params}
    )


def _get_anthropic_llm(dynamic_token: str = None):
    _init_env()

    api_key = os.getenv("x-api-key") or os.getenv("X_API_KEY") or os.getenv("ANTHROPIC_API_KEY")
    base_url = os.getenv("ANTHROPIC_BASE_URL") or None
    model = os.getenv("LLM_MODEL1", "MiniMax-M2.7")
    header_name = os.getenv("CUSTOM_HEADER_NAME")

    custom_headers = {}
    if header_name and dynamic_token:
        custom_headers[header_name] = dynamic_token

    if not api_key:
        raise ValueError("未找到 API Key (x-api-key)，请检查 .env 文件。")

    from langchain_anthropic import ChatAnthropic

    return ChatAnthropic(
        model_name=model,
        temperature=0,
        api_key=api_key,
        anthropic_api_url=base_url,
        default_headers=custom_headers,
        streaming=False
    )


def _get_llm_postman(dynamic_token: str = None):
    _init_env()

    api_key = os.getenv("x-api-key") or os.getenv("X_API_KEY") or os.getenv("ANTHROPIC_API_KEY")

    base_url = os.getenv("ANTHROPIC_BASE_URL", "").rstrip("/")
    if not base_url.endswith("/messages"):
        base_url += "/v1/messages"

    model = os.getenv("LLM_MODEL1", "MiniMax-M2.7")
    header_name = os.getenv("CUSTOM_HEADER_NAME")

    if not api_key:
        raise ValueError("未找到 API Key，请检查 .env 文件。")

    def postman_caller(prompt_input) -> AIMessage:
        messages_payload = []
        system_prompt = ""  # 用于存放剥离出来的顶层 System Prompt

        if isinstance(prompt_input, str):
            messages_payload.append({"role": "user", "content": prompt_input})
        elif isinstance(prompt_input, list):
            for msg in prompt_input:
                if isinstance(msg, SystemMessage):
                    system_prompt += msg.content + "\n"
                elif isinstance(msg, HumanMessage):
                    messages_payload.append({"role": "user", "content": msg.content})
                elif isinstance(msg, AIMessage):
                    messages_payload.append({"role": "assistant", "content": msg.content})
                elif isinstance(msg, dict) and "role" in msg:
                    if msg["role"] == "system":
                        system_prompt += msg["content"] + "\n"
                    else:
                        messages_payload.append(msg)

        headers = {
            "x-api-key": f"{api_key}",
            "Content-Type": "application/json"
        }
        if header_name and dynamic_token:
            headers[header_name] = dynamic_token

        payload = {
            "model": model,
            "temperature": 0.0,
            "stream": False,
            "messages": messages_payload
        }

        if system_prompt.strip():
            payload["system"] = system_prompt.strip()

        try:
            response = requests.post(base_url, headers=headers, json=payload, timeout=600)
            response.raise_for_status()

            response_json = response.json()

            content_list = response_json.get("content", [])
            if content_list and isinstance(content_list, list):
                content = content_list[0].get("text", "")
            else:
                content = ""

            return AIMessage(
                content=content,
                response_metadata={
                    "postman_raw_response": response_json
                }
            )

        except Exception as e:
            print(f"❌ Postman 直连请求失败: {str(e)}")
            if 'response' in locals() and hasattr(response, 'text'):
                print(f"❌ 原始报错报文: {response.text}")
            raise e

    return RunnableLambda(postman_caller)

def _get_llm_postman_openai(dynamic_token: str = None):
    _init_env()

    api_key = os.getenv("OPENAI_API_KEY")

    base_url = os.getenv("OPENAI_BASE_URL", "").rstrip("/")
    base_url += "/chat/completions"

    model = os.getenv("LLM_MODEL114514", "Qwen-V3_6-27B")
    header_name = os.getenv("CUSTOM_HEADER_NAME", "Authorization")

    if not api_key:
        raise ValueError("can not FIND API key, please check the .env file")

    def postman_caller(prompt_input) -> AIMessage:
        messages_payload = []

        if isinstance(prompt_input, str):
            messages_payload.append({"role": "user", "content": prompt_input})
        elif isinstance(prompt_input, list):
            for msg in prompt_input:
                if isinstance(msg, SystemMessage):
                    messages_payload.append({"role": "system", "content": msg.content})
                elif isinstance(msg, HumanMessage):
                    messages_payload.append({"role": "user", "content": msg.content})
                elif isinstance(msg, AIMessage):
                    messages_payload.append({"role": "assistant", "content": msg.content})
                elif isinstance(msg, dict) and "role" in msg:
                    messages_payload.append(msg)

        headers = {
            "Authorization": f"{api_key}",
            "Content-Type": "application/json; charset=utf-8"
        }
        if header_name and dynamic_token:
            headers[header_name] = dynamic_token

        payload = {
            "model": model,
            "temperature": 0.0,
            "stream": False,
            "messages": messages_payload
        }
        payload_bytes = json.dumps(payload, ensure_ascii=False).encode("utf8")

        try:
            response = requests.post(base_url,
                                     headers=headers,
                                     data=payload_bytes,
                                     timeout=600)
            response.raise_for_status()

            response_json = response.json()

            choices = response_json.get("choices", [])
            if choices and isinstance(choices, list):
                content = choices[0].get("message", {}).get("content", "")
            else:
                content = ""

            return AIMessage(
                content=content,
                response_metadata = {
                    "postman_raw_response": response_json
                }
            )
        except Exception as e:
            print(f"Postman 请求失败: {str(e)}")
            if 'response' in locals() and hasattr(response, 'text'):
                print(f"原始报错问: {response.text}")

            raise e
    return RunnableLambda(postman_caller)

def _extract_code(raw: str) -> str:
    """从 LLM 输出中安全提取最后一个 Python 代码块，并清理所有 import 语句"""

    matches = re.findall(r"```python\s*([\s\S]*?)```", raw, re.IGNORECASE)

    if matches:
        code = matches[-1].strip()
    else:
        fallback_matches = re.findall(r"```\s*([\s\S]*?)```", raw)
        code = fallback_matches[-1].strip() if fallback_matches else raw.strip()

    try:
        tree = ast.parse(code)
        tree.body = [node for node in tree.body if not isinstance(node, (ast.Import, ast.ImportFrom))]
        code = ast.unparse(tree)

    except Exception:
        code = re.sub(r"^(?:from\s+[\w.]+\s+)?import\s+.*$", "", code, flags=re.MULTILINE).strip()
        code = re.sub(r'\n\s*\n', '\n', code)

    return code


def kv_codegen_node(state: AgentState) -> dict:
    """KV 专属代码生成节点"""

    config = state.get("config", {})
    kv_list = config.get("kv_list", [])

    kv_state = state.get("kv_state", {})
    st = kv_state.get("sheet_structure", {})

    errors = state.get("errors", [])
    sandbox_error = state.get("sandbox_error")

    max_row = st.get("max_row", 1000)
    max_col = st.get("max_col", 100)
    merged_cells_info = st.get("merged_cells_info", [])
    sample_cells = st.get("sample_cells", [])

    bboxes = st.get("bboxes", {})
    bboxes_info = json.dumps(bboxes, ensure_ascii=False, indent=2) if bboxes else "{}"

    retry_instruction = ""
    if sandbox_error:
        retry_instruction = (
            "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n"
            "【警告：代码执行崩溃】\n"
            f"上次你生成的代码在沙盒中执行时报错了，报错堆栈如下：\n{sandbox_error}\n"
            "请仔细分析上述报错，修复代码中引发异常的漏洞！\n"
            "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n"
        )
    elif errors:
        retry_instruction = (
            "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n"
            "【警告：质量校验不达标】\n"
            f"上次你生成的代码执行成功，但业务线校验报错：\n{errors[-1]}\n"
            "请根据上述反馈，重新审视你的偏移量计算或视窗下边界探针！\n"
            "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n"
        )

    user_prompt = _USER_PROMPT_TEMPLATE.format(
        kv_configs_str=json.dumps(kv_list, ensure_ascii=False, indent=2),
        max_row=max_row,
        max_col=max_col,
        bboxes_info=bboxes_info,
        merged_cells_info="\n".join(merged_cells_info) if merged_cells_info else "无合并单元格",
        sample_cells="\n".join(sample_cells) if sample_cells else "无非空数据",
        retry_instruction=retry_instruction
    )

    messages = [
        SystemMessage(content=_KV_SYSTEM_PROMPT),
        HumanMessage(content=user_prompt),
    ]

    try:
        response = _get_llm_postman_openai().invoke(messages)
        code = _extract_code(response.content)
        print("———————————————————————— 大模型生成的代码 ————————————————————————————————")
        print(code)

    except Exception as e:
        return {"errors": errors + [f"生成 KV 代码请求失败: {str(e)}"]}

    if not code:
        return {"errors": errors + ["大模型未能生成有效的 Python 代码"]}

    return {
        "kv_state":
            {
                "generated_code": [code],
                "sandbox_error": None
            }
    }

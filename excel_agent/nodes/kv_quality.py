"""
nodes/kv_quality.py — KV 模式质量打分 + 熔断

质量校验维度（空间置信度算法）：
    1. 视线法则（50% 权重）：Key 和 Value 的矩形框在 X 轴或 Y 轴必须有投影交集
       - 兼容合并单元格：将 Key 和 Value 视为 Bounding Box 而非点
       - 对角线偏移直接熔断（置信度 0）

    2. 距离衰减（30% 权重）：
       - 相邻（无间隔）：1.0
       - 间隔 1-2 个单元格：0.8
       - 间隔 3-5 个单元格：0.5
       - 间隔 >5 个单元格：0.1

    3. Key 发现率（20% 权重）：found_keys / len(kv_list)

熔断阈值：
    - 综合置信度 < 0.90 → 触发验证失败，拒绝保存代码
    - 任一 KV 对视线法则不通过 → 直接返回 0 分
"""

import re
from typing import Dict, Any, Optional, Tuple, List

from excel_agent.state import AgentState

# ==================== 配置常量 ====================

QUALITY_THRESHOLD = 0.90  # 熔断阈值
MAX_RETRY = 3             # 最大重试次数


# ==================== Bounding Box 工具函数 ====================

def _get_bounding_box(r: int, c: int, merged_map: Dict) -> Dict[str, int]:
    """
    获取单元格所属的合并区域边界

    Args:
        r, c: 单元格坐标
        merged_map: 合并单元格信息字典

    Returns:
        {"min_row", "max_row", "min_col", "max_col"}
    """
    info = merged_map.get((r, c))
    if info and isinstance(info, dict) and "top_left" in info:
        tl = info["top_left"]
        return {
            "min_row": tl[0],
            "max_row": tl[0] + info["r_span"] - 1,
            "min_col": tl[1],
            "max_col": tl[1] + info["c_span"] - 1
        }
    # 非合并单元格，返回自身
    return {"min_row": r, "max_row": r, "min_col": c, "max_col": c}


def _check_line_of_sight(key_box: Dict[str, int], val_box: Dict[str, int]) -> float:
    """
    检查 Key 和 Value 是否有视线交集（投影交集）

    Bounding Box 投影交集条件：
    - X 轴交集：key_box 的列范围 与 val_box 的列范围有重叠
    - 或 Y 轴交集：key_box 的行范围 与 val_box 的行范围有重叠

    Returns:
        1.0: 有交集（同行或同列）
        0.0: 无交集（对角线偏移，熔断）
    """
    # X 轴投影（列范围）
    key_col_range = (key_box["min_col"], key_box["max_col"])
    val_col_range = (val_box["min_col"], val_box["max_col"])
    x_overlap = not (key_col_range[1] < val_col_range[0] or val_col_range[1] < key_col_range[0])

    # Y 轴投影（行范围）
    key_row_range = (key_box["min_row"], key_box["max_row"])
    val_row_range = (val_box["min_row"], val_box["max_row"])
    y_overlap = not (key_row_range[1] < val_row_range[0] or val_row_range[1] < key_row_range[0])

    return 1.0 if (x_overlap or y_overlap) else 0.0


def _compute_distance_score(key_box: Dict[str, int], val_box: Dict[str, int]) -> float:
    """
    计算距离衰减分数

    规则：
    - 相邻（无间隔单元格）：1.0
    - 间隔 1-2 个单元格：0.8
    - 间隔 3-5 个单元格：0.5
    - 间隔 >5 个单元格：0.1
    """
    # 计算最小间隔
    gap = 999  # 默认最大值

    # 检查同行情况
    if key_box["max_row"] == val_box["min_row"] or key_box["min_row"] == val_box["max_row"]:
        # 同行，计算列间隔
        if val_box["min_col"] > key_box["max_col"]:
            gap = val_box["min_col"] - key_box["max_col"] - 1
        elif key_box["min_col"] > val_box["max_col"]:
            gap = key_box["min_col"] - val_box["max_col"] - 1
        else:
            gap = 0  # 有重叠

    # 检查同列情况
    elif key_box["max_col"] == val_box["min_col"] or key_box["min_col"] == val_box["max_col"]:
        # 同列，计算行间隔
        if val_box["min_row"] > key_box["max_row"]:
            gap = val_box["min_row"] - key_box["max_row"] - 1
        elif key_box["min_row"] > val_box["max_row"]:
            gap = key_box["min_row"] - val_box["max_row"] - 1
        else:
            gap = 0  # 有重叠

    if gap <= 0:
        return 1.0
    elif gap <= 2:
        return 0.8
    elif gap <= 5:
        return 0.5
    else:
        return 0.1


def _find_cell_by_text(text: str, ws, merged_map: Dict) -> Optional[Tuple[int, int]]:
    """
    在工作表中查找包含指定文本的单元格

    Returns:
        (row, col) 或 None
    """
    if not text:
        return None

    text_norm = re.sub(r'[\s\-_\(\)\[\]\/\.,]+', '', str(text).lower())

    for r in range(1, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            val = merged_map.get((r, c), ws.cell(row=r, column=c).value)
            if val and text_norm in re.sub(r'[\s\-_\(\)\[\]\/\.,]+', '', str(val).lower()):
                return (r, c)

    return None


def compute_spatial_confidence(
    kv_result: Dict[str, str],
    kv_list: List[str],
    ws,
    merged_map: Dict
) -> Tuple[float, List[str]]:
    """
    计算 KV 抽取结果的空间置信度

    Args:
        kv_result: 抽取结果 {key: value}
        kv_list: 原始 Key 列表
        ws: openpyxl worksheet
        merged_map: 合并单元格填充图

    Returns:
        (confidence_score, details)
        - confidence_score: 0.0 ~ 1.0
        - details: 详细分析信息列表
    """
    if not kv_result:
        return 0.0, ["KV 抽取结果为空"]

    if not kv_list:
        return 0.0, ["KV 列表为空"]

    details = []
    individual_scores = []
    found_keys = 0

    for key, value in kv_result.items():
        if not value:
            details.append(f"Key '{key}' 未找到 Value")
            continue

        found_keys += 1

        # 1. 定位 Key 和 Value 的坐标
        key_pos = _find_cell_by_text(key, ws, merged_map)
        val_pos = _find_cell_by_text(value, ws, merged_map)

        if not key_pos or not val_pos:
            details.append(f"Key '{key}' 或 Value 无法定位到单元格")
            individual_scores.append(0.0)
            continue

        # 2. 计算 Bounding Box
        key_box = _get_bounding_box(key_pos[0], key_pos[1], merged_map)
        val_box = _get_bounding_box(val_pos[0], val_pos[1], merged_map)

        # 3. 视线法则检查
        los_score = _check_line_of_sight(key_box, val_box)
        if los_score == 0:
            details.append(
                f"Key '{key}' 视线法则熔断：Key[{key_box}] 与 Value[{val_box}] 无投影交集（对角线偏移）"
            )
            individual_scores.append(0.0)  # 熔断
            continue

        # 4. 距离衰减计算
        dist_score = _compute_distance_score(key_box, val_box)

        # 5. 单项得分
        item_score = los_score * 0.5 + dist_score * 0.3
        individual_scores.append(item_score)

        details.append(
            f"Key '{key}': 视线={los_score:.1f}, 距离={dist_score:.1f}, 得分={item_score:.2f}"
        )

    # Key 发现率
    discovery_rate = found_keys / len(kv_list) if kv_list else 0
    discovery_score = discovery_rate * 0.2

    # 综合得分
    if individual_scores:
        avg_score = sum(individual_scores) / len(individual_scores)
    else:
        avg_score = 0

    final_score = avg_score + discovery_score
    details.append(f"Key 发现率：{discovery_rate:.0%} (+{discovery_score:.2f})")
    details.append(f"最终置信度：{final_score:.2f}")

    return min(1.0, final_score), details


# ==================== 质量节点主逻辑 ====================

def kv_quality_node(state: AgentState) -> dict:
    """KV 模式质量打分节点"""
    kv_result = state.get("kv_result", {})
    config = state.get("config", {})
    kv_list = config.get("kv_list", [])

    errors = list(state.get("errors", []))

    # 空结果检查
    if not kv_result:
        errors.append("[KV 质量] 抽取结果为空字典")
        return {
            "quality_score": 0.0,
            "errors": errors
        }

    # 空 Key 列表检查
    if not kv_list:
        errors.append("[KV 质量] KV 列表为空")
        return {
            "quality_score": 0.0,
            "errors": errors
        }

    # 计算空间置信度
    # 注意：这里需要访问 ws 和 merged_map，但从 state 中无法直接获取
    # 简化版本：只检查 Key 发现率和基本格式
    found_keys = sum(1 for k, v in kv_result.items() if v)
    discovery_rate = found_keys / len(kv_list)

    # 简化版置信度计算（完整版本需要在 sandbox 中传递 ws 和 merged_map）
    confidence = discovery_rate

    details = [
        f"Key 发现率：{discovery_rate:.0%} ({found_keys}/{len(kv_list)})"
    ]

    # 质量打分
    if discovery_rate < 0.5:
        confidence = 0.0
        errors.append(f"[KV 质量] Key 发现率过低：{discovery_rate:.0%}")

    if discovery_rate < 1.0:
        missing_keys = [k for k in kv_list if not kv_result.get(k)]
        errors.append(f"[KV 质量] 以下 Key 未找到值：{missing_keys}")

    return {
        "quality_score": round(confidence, 3),
        "errors": errors,
        "kv_quality_details": details
    }


def route_after_kv_quality(state: AgentState) -> str:
    """KV 质量检查后路由"""
    quality_score = state.get("quality_score", 0.0)
    retry_count = state.get("retry_count", 0)

    # 质量达标，进入缓存保存
    if quality_score >= QUALITY_THRESHOLD:
        return "save"

    # 超过最大重试次数，强制结束（人工介入）
    if retry_count >= MAX_RETRY:
        return "human"

    # 质量不达标，重试
    return "retry"

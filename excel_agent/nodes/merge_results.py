"""
nodes/merge_results.py — 混合模式结果合并节点

职责：
    合并 table 模式的 raw_result（{title: 2D 数组}）和 KV 模式的 kv_result
    （{title: {key: value}} 或 {key: value}）。

    输出格式：
    {
        "TableSubTable": [[header1, header2], [row1], [row2]],  # table 结果不变
        "KvSubTable": {"key1": "value1", "key2": "value2"},     # KV 结果
    }

注意：
    - 混合模式下，table 子表结果经过 restore 列过滤，KV 子表结果不需要列过滤
    - 纯 KV 模式不走此节点（走 kv_quality 路径）
    - 纯 Table 模式不走此节点（走 restore → quality 路径）
"""

from typing import Any

from excel_agent.state import AgentState


def merge_results_node(state: AgentState) -> dict:
    """合并 table raw_result 和 kv_result"""
    raw_result = state.get("raw_result") or {}
    kv_result = state.get("kv_result") or {}

    merged: dict[str, Any] = {}
    merged.update(raw_result)

    # KV 结果可能是 {title: {key: value}}（多子表嵌套）或 {key: value}（全局 KV）
    if kv_result:
        first_val = next(iter(kv_result.values()), None)
        if isinstance(first_val, dict):
            # 多子表嵌套格式：{title: {key: value}}，直接合并
            merged.update(kv_result)
        elif isinstance(kv_result, dict):
            # 全局 KV 格式：{key: value}
            # 这种情况在混合模式下不应该出现（混合模式的 KV 结果总是嵌套的）
            # 但为安全起见，用一个默认 key 存储
            if not raw_result:
                merged.update(kv_result)

    return {"result": merged}

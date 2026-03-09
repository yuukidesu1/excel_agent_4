"""
测试跨文件缓存复用

目的：
  验证 Level 2 缓存能否在**不同 Excel 文件**之间复用
  场景：两个 Excel 文件结构相同，但子表起始行不同
"""

import os
import time
from excel_agent.agent import run_extraction
from excel_agent.cache import clear_cache, get_cache_stats

# 测试配置 - 使用两个结构相同但起始行不同的 Excel 文件
EXCEL_PATH_1 = "./ws_test_file_use.xlsx"  # 原始文件
EXCEL_PATH_2 = "./ws_test_file_use.xlsx"  # 如果有另一个测试文件，请修改这里

SHEET_NAME = "CONFIGURATION"
SUBTABLE_TITLES = ["2G Configuration", "4G Configuration", "5G Configuration"]


def format_time(seconds: float) -> str:
    if seconds < 0.1:
        return f"{seconds * 1000:.1f}ms"
    elif seconds < 1:
        return f"{seconds * 1000:.0f}ms"
    else:
        return f"{seconds:.2f}s"


def test_cross_file_cache():
    """测试跨文件缓存复用"""

    print("=" * 70)
    print("Excel Agent 跨文件缓存复用测试")
    print("=" * 70)

    # 清空缓存
    print("\n[准备] 清空缓存...")
    clear_cache()

    # ───────────────────────────────────────────────────────────────
    # 文件 1：第一次运行（LLM 生成）
    # ───────────────────────────────────────────────────────────────
    print("\n" + "-" * 70)
    print(f"【文件 1】{EXCEL_PATH_1}")
    print("第一次运行（无缓存，LLM 生成代码）")
    print("-" * 70)

    start = time.time()
    result1 = run_extraction(
        excel_path=EXCEL_PATH_1,
        sheet_name=SHEET_NAME,
        subtable_titles=SUBTABLE_TITLES,
    )
    elapsed1 = time.time() - start

    print(f"\n耗时：{format_time(elapsed1)}")
    print(f"缓存命中：{result1.get('cache_hit', False)}")
    print(f"成功：{result1.get('success', False)}")
    print(f"质量评分：{result1.get('quality_score', 0)}")

    if result1.get('success'):
        data = result1.get('data', {})
        total_rows = sum(len(table) - 1 for table in data.values() if table)
        print(f"提取数据：{total_rows} 行")

    # ───────────────────────────────────────────────────────────────
    # 文件 1：第二次运行（Level 1 命中）
    # ───────────────────────────────────────────────────────────────
    print("\n" + "-" * 70)
    print(f"【文件 1】{EXCEL_PATH_1}")
    print("第二次运行（Level 1 缓存命中）")
    print("-" * 70)

    start = time.time()
    result2 = run_extraction(
        excel_path=EXCEL_PATH_1,
        sheet_name=SHEET_NAME,
        subtable_titles=SUBTABLE_TITLES,
    )
    elapsed2 = time.time() - start

    print(f"\n耗时：{format_time(elapsed2)}")
    print(f"缓存命中：{result2.get('cache_hit', False)}")
    print(f"缓存层级：{result2.get('cache_level', 'N/A')}")

    if elapsed2 > 0:
        print(f"加速比：{elapsed1 / elapsed2:.1f}x")

    # ───────────────────────────────────────────────────────────────
    # 文件 2：第一次运行（期望 Level 2 命中）
    # ───────────────────────────────────────────────────────────────
    print("\n" + "-" * 70)
    print(f"【文件 2】{EXCEL_PATH_2}")
    print("第一次运行（期望 Level 2 缓存命中 - 结构相同，起始行偏移）")
    print("-" * 70)

    start = time.time()
    result3 = run_extraction(
        excel_path=EXCEL_PATH_2,
        sheet_name=SHEET_NAME,
        subtable_titles=SUBTABLE_TITLES,
    )
    elapsed3 = time.time() - start

    print(f"\n耗时：{format_time(elapsed3)}")
    print(f"缓存命中：{result3.get('cache_hit', False)}")
    print(f"缓存层级：{result3.get('cache_level', 'N/A')}")
    if result3.get('cache_hit') and result3.get('cache_level') == 'level2':
        print(f"行偏移量：{result3.get('row_offset', 0)}")
    print(f"成功：{result3.get('success', False)}")

    if elapsed1 > 0 and elapsed3 > 0:
        print(f"相比首次运行加速比：{elapsed1 / elapsed3:.1f}x")

    # ───────────────────────────────────────────────────────────────
    # 缓存统计
    # ───────────────────────────────────────────────────────────────
    print("\n" + "-" * 70)
    print("【缓存统计】")
    print("-" * 70)

    stats = get_cache_stats()
    print(f"总缓存条目：{stats.get('total_entries', 0)}")
    print(f"Level 1 缓存：{stats.get('level1_count', 0)}")
    print(f"Level 2 缓存：{stats.get('level2_count', 0)}")

    # ───────────────────────────────────────────────────────────────
    # 总结
    # ───────────────────────────────────────────────────────────────
    print("\n" + "=" * 70)
    print("【测试总结】")
    print("=" * 70)

    print(f"""
文件 1 - 首次运行（无缓存）: {format_time(elapsed1)}
文件 1 - 二次运行（Level 1）: {format_time(elapsed2)}, {elapsed1/elapsed2:.1f}x 加速
文件 2 - 首次运行（Level 2）: {format_time(elapsed3)}, {elapsed1/elapsed3:.1f}x 加速

预期结果:
  - 文件 1 第二次运行应命中 Level 1（完全相同文件）
  - 文件 2 第一次运行应命中 Level 2（结构相同，起始行偏移）
""")

    # 验证
    if result2.get('cache_hit') and result2.get('cache_level') == 'level1':
        print("[✓] Level 1 缓存工作正常（同文件复用）")
    else:
        print("[✗] Level 1 缓存未命中")

    if result3.get('cache_hit') and result3.get('cache_level') == 'level2':
        print("[✓] Level 2 缓存工作正常（跨文件复用）")
    else:
        print("[✗] Level 2 缓存未命中（可能原因：文件完全相同导致命中 Level 1，或结构指纹不匹配）")


if __name__ == "__main__":
    test_cross_file_cache()

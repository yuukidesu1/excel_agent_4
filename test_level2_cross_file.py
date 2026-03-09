"""
测试跨文件缓存复用（Level 2 缓存）

场景：
  - Excel1: ws_test_file_use.xlsx（包含 2G|3G|4G Configuration）
  - Excel2: ws_test_file_use2.xlsx（包含 2G|3G|4G Configuration，起始行可能不同）

预期：
  - Excel1 第一次：LLM 生成（慢）
  - Excel2 第一次：Level 2 命中（快，自动计算偏移量）
"""

import os
import time
from excel_agent.agent import run_extraction
from excel_agent.cache import clear_cache, get_cache_stats

# 测试配置
EXCEL_PATH_1 = "./ws_test_file_use.xlsx"
EXCEL_PATH_2 = "./ws_test_file_use2.xlsx"

SHEET_NAME = "CONFIGURATION"
# 使用相同的子表标题
SUBTABLE_TITLES = ["2G Configuration", "3G Configuration", "4G Configuration"]


def format_time(seconds: float) -> str:
    if seconds < 0.1:
        return f"{seconds * 1000:.1f}ms"
    elif seconds < 1:
        return f"{seconds * 1000:.0f}ms"
    else:
        return f"{seconds:.2f}s"


def test_cross_file_level2():
    """测试跨文件 Level 2 缓存复用"""

    print("=" * 70)
    print("跨文件 Level 2 缓存复用测试")
    print("=" * 70)

    # 检查文件是否存在
    if not os.path.exists(EXCEL_PATH_2):
        print(f"\n警告：{EXCEL_PATH_2} 不存在，无法进行跨文件测试")
        print("请确保有两个结构相同但起始行不同的 Excel 文件")
        return

    # 清空缓存
    print("\n[准备] 清空缓存...")
    clear_cache()

    # ───────────────────────────────────────────────────────────────
    # Excel1: 第一次运行（LLM 生成）
    # ───────────────────────────────────────────────────────────────
    print("\n" + "-" * 70)
    print(f"【Excel1】{EXCEL_PATH_1}")
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

    data1 = result1.get('data', {})
    if data1:
        for title, rows in data1.items():
            print(f"  {title}: {len(rows) - 1} 行数据")

    # ───────────────────────────────────────────────────────────────
    # Excel2: 第一次运行（期望 Level 2 命中）
    # ───────────────────────────────────────────────────────────────
    print("\n" + "-" * 70)
    print(f"【Excel2】{EXCEL_PATH_2}")
    print("第一次运行（期望 Level 2 缓存命中 - 结构相同）")
    print("-" * 70)

    start = time.time()
    result2 = run_extraction(
        excel_path=EXCEL_PATH_2,
        sheet_name=SHEET_NAME,
        subtable_titles=SUBTABLE_TITLES,
    )
    elapsed2 = time.time() - start

    print(f"\n耗时：{format_time(elapsed2)}")
    print(f"缓存命中：{result2.get('cache_hit', False)}")
    print(f"缓存层级：{result2.get('cache_level', 'N/A')}")
    if result2.get('cache_hit'):
        print(f"行偏移量：{result2.get('row_offset', 0)}")
    print(f"成功：{result2.get('success', False)}")

    data2 = result2.get('data', {})
    if data2:
        for title, rows in data2.items():
            print(f"  {title}: {len(rows) - 1} 行数据")

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
Excel1 - 首次运行（无缓存）: {format_time(elapsed1)}
Excel2 - 首次运行（Level 2）: {format_time(elapsed2)}

加速比：{elapsed1 / elapsed2:.1f}x
""")

    # 验证
    if result2.get('cache_hit') and result2.get('cache_level') == 'level2':
        print("[✓] Level 2 缓存工作正常（跨文件复用成功）")
    elif result2.get('cache_hit') and result2.get('cache_level') == 'level1':
        print("[✓] Level 1 缓存命中（两个文件哈希相同）")
    else:
        print("[✗] 缓存未命中，可能原因:")
        print("  1. 两个文件结构指纹不同（列名/合并模式有差异）")
        print("  2. 子表名称不完全匹配")
        print("  3. 缓存代码不包含请求的所有子表")


if __name__ == "__main__":
    test_cross_file_level2()

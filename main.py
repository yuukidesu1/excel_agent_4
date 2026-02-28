"""
main.py — 运行入口

用法：
    python main.py

调用方式极简：只需传 excel_path + sheet_name + subtable_title，
Agent 自动完成定位、列发现、合并处理、forward-fill，输出完整二维数组。
"""

import sys
import json
from pathlib import Path

# 将项目根目录加入 Python 路径（兼容 PyCharm 直接运行）
# _ROOT = Path(__file__).resolve().parent
# if str(_ROOT) not in sys.path:
#     sys.path.insert(0, str(_ROOT))

from dotenv import load_dotenv
# load_dotenv(_ROOT / ".env", override=True)
load_dotenv()

from excel_agent import run_extraction


def main():
    result = run_extraction(
        excel_path     = "ws_test_file_use.xlsx",       # ← 替换为你的 Excel 路径
        sheet_name     = "CONFIGURATION",       # ← 目标 Sheet 名
        subtable_title = "4G Configuration",    # ← 子表标题关键词
        # hints        = "子表位于 Sheet 上半部分",  # 可选
    )

    print("=" * 60)
    print(f"抽取{'成功 ✓' if result['success'] else '失败 ✗（质量不达标）'}")
    print(f"质量分：{result['quality_score']:.2f}  |  重试次数：{result['retry_count']}")

    if result["errors"]:
        print("\n⚠ 警告/错误：")
        for e in result["errors"]:
            print(f"  {e}")

    if result["data"]:
        rows = result["data"]
        print(f"\n📋 结果：{len(rows)-1} 行数据 × {len(rows[0])} 列")
        print(json.dumps(rows, ensure_ascii=False, indent=2))
    else:
        print("\n❌ 无数据输出")


if __name__ == "__main__":
    main()

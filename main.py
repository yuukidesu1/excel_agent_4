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
from excel_agent import run_extraction_stream


# def main():
#     result = run_extraction(
#         excel_path     = "ws_test_file_use.xlsx",       # ← 替换为你的 Excel 路径
#         sheet_name     = "CONFIGURATION",       # ← 目标 Sheet 名
#         subtable_title = "4G Configuration",    # ← 子表标题关键词
#         # hints        = "子表位于 Sheet 上半部分",  # 可选
#     )
#
#     print("=" * 60)
#     print(f"抽取{'成功 ✓' if result['success'] else '失败 ✗（质量不达标）'}")
#     print(f"质量分：{result['quality_score']:.2f}  |  重试次数：{result['retry_count']}")
#
#     if result["errors"]:
#         print("\n⚠ 警告/错误：")
#         for e in result["errors"]:
#             print(f"  {e}")
#
#     if result["data"]:
#         rows = result["data"]
#         print(f"\n📋 结果：{len(rows)-1} 行数据 × {len(rows[0])} 列")
#         print(json.dumps(rows, ensure_ascii=False, indent=2))
#     else:
#         print("\n❌ 无数据输出")


def main():
    print("🚀 开始执行流式抽取任务...")

    # 1. 准备变量，用于在流式过程中捕获最终状态
    final_score = 0.0
    retry_count = 0
    errors = []
    final_data = None

    # 2. 用 for 循环迭代生成器，实现流式打印
    for chunk in run_extraction_stream(
            excel_path="ws_test_file_use.xlsx",
            sheet_name="CONFIGURATION",
            subtable_title="4G Configuration",
            # hints        = "子表位于 Sheet 上半部分",
    ):
        # 实时打印节点进度
        print(f"⏳ {chunk['message']}")

        # 捕获状态更新（LangGraph 的 updates 模式输出的是增量）
        data_update = chunk.get("data", {})

        # 不断刷新我们的最终结果变量
        if "quality_score" in data_update:
            final_score = data_update["quality_score"]
        if "retry_count" in data_update:
            retry_count = data_update["retry_count"]
        if "errors" in data_update:
            errors = data_update["errors"]

        # 捕获最终产出的数据
        if data_update.get("final_output"):
            final_data = data_update["final_output"]
        elif data_update.get("result"):
            final_data = data_update["result"]

    # 3. 流式执行完毕后，根据捕获到的最终状态进行总结判断
    success = final_score >= 0.75

    print("\n" + "=" * 60)
    print(f"抽取{'成功 ✓' if success else '失败 ✗（质量不达标）'}")
    print(f"质量分：{final_score:.2f}  |  重试次数：{retry_count}")

    if errors:
        print("\n⚠ 警告/错误：")
        for e in errors:
            print(f"  {e}")

    if final_data:
        rows = final_data
        print(f"\n📋 结果：{len(rows) - 1} 行数据 × {len(rows[0])} 列")
        print(json.dumps(rows, ensure_ascii=False, indent=2))
    else:
        print("\n❌ 无数据输出")

if __name__ == "__main__":
    main()

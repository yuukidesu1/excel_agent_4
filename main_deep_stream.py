import json
import asyncio
from dotenv import load_dotenv

load_dotenv()

# 确保你在 __init__.py 导出了 run_extraction_deep_stream
from excel_agent import run_extraction_deep_stream

QUALITY_THRESHOLD = 0.75


async def main():
    print("🚀 开始执行底层事件流抽取任务...")

    final_score = 0.0
    retry_count = 0
    errors = []
    final_data = None

    # 使用 async for 迭代异步流
    async for chunk in run_extraction_deep_stream(
            excel_path="ws_test_file_use.xlsx",
            sheet_name="CONFIGURATION",
            subtable_title="4G Configuration",
    ):
        event_type = chunk["type"]

        # 打印大模型思考过程（打字机效果）
        if event_type == "token":
            print(chunk["content"], end="", flush=True)

        # 打印工具调用
        elif event_type == "tool_start":
            print(f"\n\n🔧 [调用工具] {chunk['name']}")
            print(f"📦 参数: {json.dumps(chunk['input'], ensure_ascii=False)}")

        elif event_type == "tool_end":
            print(f"✅ [工具返回] {chunk['name']} 执行完毕\n")

        # 打印图节点进度，并收集最终状态
        elif event_type == "node_end":
            node_name = chunk["name"]
            print(f"\n🟢 节点 [{node_name}] 执行完毕")

            data_update = chunk.get("data", {})
            if "quality_score" in data_update:
                final_score = data_update["quality_score"]
            if "retry_count" in data_update:
                retry_count = data_update["retry_count"]
            if "errors" in data_update:
                errors = data_update["errors"]

            if data_update.get("final_output"):
                final_data = data_update["final_output"]
            elif data_update.get("result"):
                final_data = data_update["result"]

    # --- 最终结果输出 ---
    success = final_score >= QUALITY_THRESHOLD
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
    # 使用 asyncio.run 运行异步主函数
    asyncio.run(main())
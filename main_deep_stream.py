"""
main_deep_stream.py — 流式调用示例（调试用，支持列过滤）
"""

import sys, json, asyncio
from pathlib import Path

_ROOT = Path(__file__).resolve().parent
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

from dotenv import load_dotenv
load_dotenv(_ROOT / ".env", override=True)

from excel_agent import run_extraction_deep_stream

QUALITY_THRESHOLD = 0.75


async def main():
    print("🚀 开始执行底层事件流抽取任务...")

    final_score = 0.0
    retry_count = 0
    errors      = []
    final_data  = None

    async for chunk in run_extraction_deep_stream(
        excel_path     = "./ws_test_file_use.xlsx",
        sheet_name     = "CONFIGURATION",
        subtable_title = "4G Configuration",

        # ── 列过滤（取消注释即启用）──────────────────────────
        target_columns = [
            {"parent": None,        "child": "SYSTEM MODULE"},
            {"parent": None,        "child": "CELL"},
            {"parent": "RF MODULE", "child": "TYPE"},
            {"parent": "RF MODULE", "child": "QTY."},
            {"parent": "ANTENNAS",  "child": "NEW/SWAP/EXIST"},
            {"parent": "ANTENNAS",  "child": "Antenna Type"},
            {"parent": "ANTENNAS",  "child": "Antenna Qty."},
            {"parent": "RRU Cable", "child": "POWER LENGTH(m)"},
            {"parent": "RRU Cable", "child": "OPT LENGTH(m)"},
            {"parent": "TILT",      "child": "M"},
            {"parent": "TILT",      "child": "E"},
        ],
    ):
        t = chunk["type"]

        if t == "token":
            print(chunk["content"], end="", flush=True)

        elif t == "tool_start":
            print(f"\n\n🔧 [工具] {chunk['name']}")
            print(f"📦 参数: {json.dumps(chunk['input'], ensure_ascii=False)}")

        elif t == "tool_end":
            print(f"✅ [工具完成] {chunk['name']}\n")

        elif t == "node_end":
            node = chunk["name"]
            data = chunk.get("data", {})
            print(f"\n🟢 节点 [{node}] 完成")
            if node == "locate" and data.get("header_map"):
                hm = data["header_map"]
                print(f"   表头行数: {hm['header_row_count']}，"
                      f"子表范围: R{hm['subtable_start_row']}-R{hm['subtable_end_row']} "
                      f"C{hm['subtable_start_col']}-C{hm['subtable_end_col']}")
                print(f"   发现列数: {len(hm.get('all_columns', []))}")
            if "quality_score" in data: final_score = data["quality_score"]
            if "retry_count"   in data: retry_count = data["retry_count"]
            if "errors"        in data: errors       = data["errors"]
            if data.get("final_output"): final_data = data["final_output"]
            elif data.get("result"):     final_data = data["result"]

    success = final_score >= QUALITY_THRESHOLD
    print("\n" + "=" * 60)
    print(f"抽取{'成功 ✓' if success else '失败 ✗'}")
    print(f"质量分：{final_score:.2f}  |  重试次数：{retry_count}")

    if errors:
        print("\n⚠ 警告/错误：")
        for e in errors: print(f"  {e}")

    if final_data:
        rows = final_data
        print(f"\n📋 结果：{len(rows)-1} 行数据 × {len(rows[0])} 列")
        print(json.dumps(rows, ensure_ascii=False, indent=2))
    else:
        print("\n❌ 无数据输出")


if __name__ == "__main__":
    asyncio.run(main())
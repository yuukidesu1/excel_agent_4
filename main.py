"""
main.py — 同步调用

演示两种用法：
  1. 全列模式（target_columns=None，默认）
  2. 列过滤模式（只保留指定列）
"""

import sys, json
from pathlib import Path

_ROOT = Path(__file__).resolve().parent
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

from dotenv import load_dotenv
load_dotenv(_ROOT / ".env", override=True)

from excel_agent import run_extraction


# ── 示例 1：全列模式（保留子表所有列）──────────────────────────
result = run_extraction(
    excel_path     = "./ws_test_file_use.xlsx",
        sheet_name     = "CONFIGURATION",
        subtable_titles = "3G Configuration",
        # target_columns = [
        #     {"parent": None,        "child": "SYSTEM MODULE"},
        #     {"parent": None,        "child": "CELL"},
        #     {"parent": "RF MODULE", "child": "TYPE"},
        #     {"parent": "RF MODULE", "child": "QTY."},
        #     {"parent": "ANTENNAS",  "child": "NEW/SWAP/EXIST"},
        #     {"parent": "ANTENNAS",  "child": "Antenna Type"},
        #     {"parent": "ANTENNAS",  "child": "Antenna Qty."},
        #     {"parent": "RRU Cable", "child": "POWER LENGTH(m)"},
        #     {"parent": "RRU Cable", "child": "OPT LENGTH(m)"},
        #     {"parent": "TILT",      "child": "M"},
        #     {"parent": "TILT",      "child": "E"},
        # ],
)

success = result["success"]
print(f"\n{'成功 ✓' if success else '失败 ✗'}  "
      f"质量分：{result['quality_score']:.2f}  重试：{result['retry_count']} 次")

if result.get("sandbox_error"):
    print(f"\n⛔ 沙盒错误：\n{result['sandbox_error']}")

if result["errors"]:
    for e in result["errors"]:
        print(f"  ⚠ {e}")

print(f"\n── LLM 生成的代码 ──────────────────────────────────")
print(result.get("generated_code", "（无）"))

if result["data"]:
    rows = result["data"]
    print(f"\n── 结果：{len(rows)-1} 行数据 × {len(rows[0])} 列 ──────────────")
    print(json.dumps(rows, ensure_ascii=False, indent=2))
else:
    print("\n❌ 无数据输出")
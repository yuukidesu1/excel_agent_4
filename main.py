import json
from pathlib import Path

from dotenv import load_dotenv
from excel_agent import run_extraction

load_dotenv(Path(__file__).resolve().parent / ".env", override=True)

result = run_extraction(
    excel_path="./ws_test_file_use.xlsx",
    sheet_name="CONFIGURATION",
    subtable_titles=["3G Configuration"],
)

success = result["success"]
print(f"\n{'成功 ✓' if success else '失败 ✗'}  "
      f"质量分：{result['quality_score']:.2f}  重试：{result['retry_count']} 次")

if result.get("sandbox_error"):
    print(f"\n⛔ 沙盒错误：\n{result['sandbox_error']}")

if result["errors"]:
    for e in result["errors"]:
        print(f"  ⚠ {e}")

print("\n── LLM 生成的代码 ──────────────────────────────────")
print(result.get("generated_code", "（无）"))

if result["data"]:
    rows = result["data"]
    print(f"\n── 结果：{len(rows)-1} 行数据 × {len(rows[0])} 列 ──────────────")
    print(json.dumps(rows, ensure_ascii=False, indent=2))
else:
    print("\n❌ 无数据输出")

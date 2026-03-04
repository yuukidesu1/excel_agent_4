import sys
import json
import argparse
import yaml
from pathlib import Path
from dotenv import load_dotenv
from excel_agent import run_extraction
from utils.helpers import timer

load_dotenv()

@timer
def main(config_file: str):
    config_path = Path(config_file)
    if not config_path.exists():
        raise FileNotFoundError(f"cannot find the config files: {config_file}")

    with open(config_path, "r", encoding="utf8") as f:
        config = yaml.safe_load(f)

    excel_path = config.get("excel_path")
    sheet_name = config.get("sheet_name")
    subtable_title = config.get("subtable_title")
    target_columns = config.get("target_columns", None)

    if not all([excel_path, sheet_name, subtable_title]):
        raise ValueError("配置文件中必须包含 excel_path, sheet_name 和 subtable_title")

    print("="*100)
    print("loading config...")
    print(config_path.name)
    print("="*100)
    print("loading excel...")
    print(f"📊 目标 Excel: {excel_path} | Sheet: {sheet_name} | 子表: {subtable_title}")
    print("="*100)

    kwargs = {
        "excel_path": excel_path,
        "sheet_name": sheet_name,
        "subtable_title": subtable_title,
    }
    if target_columns is not None:
        kwargs["target_columns"] = target_columns

    print("开始执行抽取任务...")
    result = run_extraction(**kwargs)

    # ── 打印结果 ──────────────────────────────────────────────
    success = result.get("success", False)
    print(f"\n{'成功 ✓' if success else '失败 ✗'}  "
          f"质量分：{result.get('quality_score', 0):.2f}  重试：{result.get('retry_count', 0)} 次")

    if result.get("sandbox_error"):
        print(f"\n⛔ 沙盒错误：\n{result['sandbox_error']}")

    if result.get("errors"):
        for e in result["errors"]:
            print(f"  ⚠ {e}")

    print(f"\n── LLM 生成的代码 ──────────────────────────────────")
    print(result.get("generated_code", "（无）"))

    if result.get("data"):
        rows = result["data"]
        print(f"\n── 结果：{len(rows) - 1} 行数据 × {len(rows[0])} 列 ──────────────")
        print(json.dumps(rows, ensure_ascii=False, indent=2))
    else:
        print("\n❌ 无数据输出")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Excel 抽取代理")
    parser.add_argument(
        "-c", "--config",
        type=str,
        default="./configs/glm5_WL_56A0DS6_multi_subtitle_test.yaml",
        help="YAML 配置文件路径 (例如./configs/task_3g.yaml)"
    )
    args = parser.parse_args()

    main(args.config)
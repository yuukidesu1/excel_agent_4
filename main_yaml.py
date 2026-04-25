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
    extract_type = config.get("extract_type", "table")

    # KV 模式使用 kv_list，表格模式使用 subtable_titles
    subtable_titles = config.get("subtable_titles") or config.get("kv_list")
    subtable_configs = config.get("subtable_configs") or config.get("target_columns")

    if not all([excel_path, sheet_name, subtable_titles]) and extract_type != "kv" and extract_type != "KV":
        raise ValueError("配置文件中必须包含 excel_path, sheet_name 和 subtable_titles 或 kv_list")

    print("="*100)
    print("loading config...")
    print(config_path.name)
    print("="*100)
    print("loading excel...")
    print(f"📊 目标 Excel: {excel_path} | Sheet: {sheet_name} | 子表: {subtable_titles}")
    print("="*100)

    kwargs = {
        "excel_path": excel_path,
        "sheet_name": sheet_name,
        "subtable_titles": subtable_titles,
    }
    if subtable_configs is not None:
        kwargs["subtable_configs"] = subtable_configs

    # KV 模式需要传递额外参数
    if extract_type and extract_type.lower() == "kv":
        kwargs["extract_type"] = "kv"
        kwargs["kv_list"] = subtable_titles

    # 用户明确意图优先：如果配置了 kv_list，直接传递（无论 extract_type 是什么）
    kv_list_from_config = config.get("kv_list")
    if kv_list_from_config:
        kwargs["kv_list"] = kv_list_from_config
        # 如果没有显式指定 extract_type，也设置为 kv
        if "extract_type" not in kwargs:
            kwargs["extract_type"] = "kv"

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

    # 判断是否是 KV 模式（配置文件指定或 PSA 检测到 KV 布局）
    is_kv_mode = (extract_type and extract_type.lower() == "kv") or (result.get("kv_result") is not None)

    if result.get("data") or result.get("kv_result"):
        # KV 模式优先使用 kv_result
        extracted_data = result.get("kv_result") if is_kv_mode else result.get("data")

        # 1. 如果返回的是字典（多子表或 KV 模式）
        if isinstance(extracted_data, dict):
            if is_kv_mode:
                # KV 模式：{key: value} 格式
                print("\n── KV 抽取结果 ──────────────────────────────")
                print(f" 📊 共抽取 {len(extracted_data)} 个键值对")
                for key, value in extracted_data.items():
                    print(f"   • {key}: {value[:50] if value and len(value) > 50 else value}")
                print("────────────────────────────────────────────────")
            else:
                # 表格模式：多子表
                print("\n── 多子表抽取结果汇总 ──────────────────────────────")
                for table_name, rows in extracted_data.items():
                    if rows and isinstance(rows, list):
                        data_rows = len(rows) - 1 if len(rows) > 0 else 0
                        cols = len(rows[0]) if data_rows >= 0 and isinstance(rows[0], list) else 0
                        print(f" 📊 [{table_name}]: {data_rows} 行数据 × {cols} 列")
                    else:
                        print(f" ⚠ [{table_name}]: 提取结果为空或格式不符")
                print("──────────────────────────────────────────────────")

        # 2. 如果返回的是二维列表（单个表格，向下兼容）
        elif isinstance(extracted_data, list):
            rows = extracted_data
            if rows:
                data_rows = len(rows) - 1 if len(rows) > 0 else 0
                cols = len(rows[0]) if isinstance(rows[0], list) else 0
                print(f"\n── 单表结果：{data_rows} 行数据 × {cols} 列 ──────────────")
            else:
                print("\n── 结果：空列表 ──────────────")

        # 无论哪种格式，都打印完整的 JSON 数据供核对
        print(json.dumps(extracted_data, ensure_ascii=False, indent=2))
    else:
        print("\n❌ 无数据输出")



if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Excel 抽取代理")
    parser.add_argument(
        "-c", "--config",
        type=str,
        # default="./configs/generalization_test/Power_56A0NNC_kv_table.yaml",
        # default="./configs/TSSR_senario_TEST.yaml",
        # default="./configs/generalization_test/CONFIGURATION.yaml",
        # default="./configs/generalization_test/MW_56A0DS6.yaml",
        # default="./configs/generalization_test/WL_56A0DS6.yaml",
        # default="./configs/test/test_horizontal.yaml",
        # default="./configs/generalization_test/56A0DS6_PowerLoadInformation.yaml",
        # default="./configs/generalization_test/Power_56A0NNC.yaml",
        # default="./configs/test/test_cross.yaml",
        # default="./configs/test/TEST_WL_56A0DS6_mix.yaml",
        # default="./configs/kv_test/kv_list.yaml",
        default="./configs/generalization_test/CONFIGURATION.yaml",
        # default="./configs/kv_test/kv_table_test.yaml",
        help="YAML 配置文件路径 (例如./configs/task_3g.yaml)"
    )
    args = parser.parse_args()

    main(args.config)
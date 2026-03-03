import sys
import pandas as pd
import json


def read_json_from_stdin():
    """
    从标准输入读取 JSON 字符串
    CMD+D (Mac/Linux) 或 Ctrl+Z + Enter (Windows) 结束输入
    """
    print("请粘贴 JSON 数据，输入完成后按 Ctrl+D 结束：")
    input_data = sys.stdin.read()
    return json.loads(input_data)


def json_to_excel(data, output_file):
    """
    将二维列表 JSON 转换为 Excel
    data[0] 为表头
    data[1:] 为数据行
    """
    df = pd.DataFrame(data[1:], columns=data[0])
    df.to_excel(output_file, index=False)
    print(f"成功！Excel 文件已保存为: {output_file}")


if __name__ == "__main__":
    json_data = read_json_from_stdin()
    json_to_excel(json_data, "result.xlsx")
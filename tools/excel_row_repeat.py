from itertools import repeat

import pandas as pd
from pathlib import Path

def run():
    base_path = Path(__file__).resolve().parents[1]
    data_dir = base_path / "data" / "excel_row_repeat"
    input_dir = data_dir / "input"
    output_dir = data_dir / "output"

    input_dir.mkdir(parents=True, exist_ok=True)
    output_dir.mkdir(parents=True, exist_ok=True)

    print(f"\n📂 输入文件目录: {input_dir}")
    print(f"📂 输出文件目录: {output_dir}")
    # input file
    input_file = "复制N行内容-input.xlsx"
    input_path = input_dir / input_file
    df = pd.read_excel(input_path, dtype=str)

    # 每行重复 10 次（含原行），如果只想添加9行，则改为 repeat(9)
    repeat_num = int(input("请输入重复次数（如 17）: ").strip())
    repeated_df = df.loc[df.index.repeat(repeat_num)].reset_index(drop=True)

    # 保存为 Excel，强制所有单元格为“文本”格式
    output_file = "复制N行内容-output.xlsx"
    output_path = output_dir / output_file
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        repeated_df.to_excel(writer, index=False, sheet_name="Sheet1")

        # 设置单元格格式为文本
        ws = writer.sheets["Sheet1"]
        for col in ws.columns:
            for cell in col:
                cell.number_format = "@"

    print("✅ 已完成：每行已复制" + str(repeat_num) + "次，所有长数字已保持原样（非科学计数法）")

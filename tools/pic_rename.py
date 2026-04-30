import os
import shutil
from pathlib import Path

def run():
    base_path = Path(__file__).resolve().parents[1]
    data_dir = base_path / "data" / "pic_rename"
    input_dir = data_dir / "input"
    output_dir = data_dir / "output"

    input_dir.mkdir(parents=True, exist_ok=True)
    output_dir.mkdir(parents=True, exist_ok=True)

    print(f"\n📂 输入文件目录: {input_dir}")
    print(f"📂 输出文件目录: {output_dir}")

    # ======== 配置 ========
    asin_file = data_dir / "asins.txt"       # ASIN 列表文件
    # 读取 ASIN 列表
    with open(asin_file, "r", encoding="utf-8") as f:
        asins = [line.strip() for line in f if line.strip()]

    if not asins:
        raise ValueError("❌ asins.txt 文件为空或未找到有效 ASIN。")

    # 获取模板图片列表（按名称排序）
    templates = sorted([f for f in os.listdir(input_dir) if f.lower().endswith(".jpg")])

    if not templates:
        raise ValueError("❌ 未在模板文件夹中找到任何 .jpg 文件。")

    # 创建输出文件夹
    os.makedirs(output_dir, exist_ok=True)

    # 处理每个 ASIN
    for asin in asins:
        print(f"\n🔹 生成图片组：{asin}")
        for tpl in templates:
            # 拆分模板名（如 MAIN.jpg -> MAIN）
            name_part = os.path.splitext(tpl)[0]
            new_filename = f"{asin}.{name_part}.jpg"

            src_path = os.path.join(input_dir, tpl)
            dst_path = os.path.join(output_dir, new_filename)

            shutil.copy2(src_path, dst_path)
            print(f"✅ {tpl} → {new_filename}")

    print("\n🎯 全部生成完成！结果已保存至：", output_dir)
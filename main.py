import importlib
import sys

# 工具清单（按编号注册）
TOOLS = {
    "1": ("周报数据清洗（领星 产品表现-MSKU）→ 生成汇总文件", "amazon_toolkit.tools.weekly_data_clean"),
    "2": ("周报数据写入（需先运行第1步）", "amazon_toolkit.tools.weekly_report_append"),
    "3": ("Excel 行复制 N 次", "amazon_toolkit.tools.excel_row_repeat"),
    "4": ("ASIN图片重命名", "amazon_toolkit.tools.asin_pic_rename"),
    "5": ("广告批量更新Bid", "amazon_toolkit.tools.ad_bulk_update"),
}

def main():
    while True:
        print("\n=== Amazon Automation Toolkit ===")
        print("请选择要运行的功能：")
        for key, (desc, _) in TOOLS.items():
            print(f"  {key}. {desc}")
        print("  0. 退出")

        choice = input("\n输入编号并回车：").strip()

        if choice == "0":
            print("👋 再见！")
            sys.exit(0)

        if choice not in TOOLS:
            print("❌ 无效选项，请重新选择。")
            continue

        _, module_name = TOOLS[choice]
        try:
            module = importlib.import_module(module_name)
            if hasattr(module, "run"):
                module.run()
            else:
                print(f"❌ 工具 {module_name} 缺少 run() 函数。")
        except Exception as e:
            print(f"⚠️ 执行失败：{e}")


if __name__ == "__main__":
    main()
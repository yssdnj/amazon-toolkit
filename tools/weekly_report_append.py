from pathlib import Path
import datetime
import pandas as pd
import xlwings as xw


def run():
    base_path = Path(__file__).resolve().parents[1]
    data_dir = base_path / "data" / "weekly_report_append"
    input_dir = data_dir / "input"
    output_dir = data_dir / "output"

    input_dir.mkdir(parents=True, exist_ok=True)
    output_dir.mkdir(parents=True, exist_ok=True)

    print(f"\n📂 输入文件目录: {input_dir}")
    print(f"📂 输出文件目录: {output_dir}")

    week_number = input("请输入输入周数（如 26W11）: ").strip()

    input_file = '周销售数据统计US-'+ week_number + '.xlsx'
    output_file = '周销售数据统计US-'+ week_number + '.xlsx'

    input_path = input_dir / input_file
    output_path = output_dir / output_file

    if not input_path.exists():
        print(f"❌ 文件不存在: {input_path}")
        return

    output_file = 'output_aggregated-' + week_number + '.xlsx'
    # ==============================================================================
    # 20260503修改：修复 source_file 路径 Bug
    # 旧代码：source_file = base_path / "data" / "weekly_report" / "output" / output_file
    #         dev 分支已将目录重命名为 weekly_data_clean，旧路径会导致文件找不到直接报错
    # 新代码：路径与 weekly_data_clean.py 的输出目录保持一致
    # ==============================================================================
    source_file = base_path / "data" / "weekly_data_clean" / "output" / output_file

    add_row(input_path, output_path)
    fill_data(source_file, output_path)

def add_row(file_path, save_path=None):

    sheets = ["SL","Toy", "ToyDH", "DL", "DSL", "MFL", "SFM"]

    # ==============================================================================
    # 20260503修改：App 生命周期改用 try/finally 保护
    # 旧代码：app.quit() 直接写在函数末尾，若中途异常会导致 Excel 进程残留后台
    # 新代码：无论是否报错，finally 块都保证 wb.close() + app.quit() 被执行
    # ==============================================================================
    app = xw.App(visible=False)
    app.display_alerts = False
    app.screen_updating = False

    try:
        wb = app.books.open(file_path)

        today = datetime.date.today() - datetime.timedelta(days=7)

        monday = today - datetime.timedelta(days=today.weekday())
        sunday = monday + datetime.timedelta(days=6)

        year, week, _ = monday.isocalendar()

        year_week = f"{year}W{week:02d}"
        week_range = f"{monday.strftime('%m%d')}~{sunday.strftime('%m%d')}"

        for sheet_name in sheets:

            if sheet_name not in [s.name for s in wb.sheets]:
                print(f"sheet不存在: {sheet_name}")
                continue

            sht = wb.sheets[sheet_name]

            # 找最后一行
            last_row = sht.range("A" + str(sht.cells.last_cell.row)).end("up").row

            insert_row = last_row
            template_row = last_row - 1

            # 插入新行（倒数两行之间）
            sht.range(f"{insert_row}:{insert_row}").api.Insert()

            # 复制模板行
            sht.range(f"{template_row}:{template_row}").api.Copy()
            # 粘贴格式
            sht.range(f"{insert_row}:{insert_row}").api.PasteSpecial(-4122)
            # 粘贴公式
            sht.range(f"{insert_row}:{insert_row}").api.PasteSpecial(-4123)

            # # 删除常量单元格（只保留公式）
            # used_cols = sht.used_range.last_cell.column
            # for col in range(3, used_cols + 1):
            #     cell = sht.range(insert_row, col)
            #     if not cell.formula:
            #         cell.value = None

            # 填充前两个单元格
            sht.range(insert_row, 1).value = week_range
            sht.range(insert_row, 2).value = year_week

            print(f"完成 sheet: {sheet_name}")

        if save_path:
            wb.save(save_path)
        else:
            wb.save()

    finally:
        # 旧代码：wb.close() / app.quit() 直接在函数末尾调用，异常时不执行
        wb.close()
        app.quit()

    return "全部sheet处理完成"

def fill_data(file1, file2):

    sheet_map = ["SL", "Toy", "ToyDH", "DL", "DSL", "MFL","SFM"]

    summary_cols = [
        "销量", "订单量", "销售额", "促销销量", "促销订单量", "促销销售额",
        "促销折扣", "退款量", "退款金额", "展示", "点击",
        "广告订单量", "广告花费", "广告销售额", "CPC"
    ]

    # 读取附件1两个表
    df_summary = pd.read_excel(file1, sheet_name="标签汇总sht")
    df_product = pd.read_excel(file1, sheet_name="标签品名汇总sht")

    # 关键修复
    df_summary["listing标签"] = df_summary["listing标签"].astype(str).str.strip()
    df_product["listing标签"] = df_product["listing标签"].astype(str).str.strip()

    # ==============================================================================
    # 20260503修改：App 生命周期改用 try/finally 保护
    # 旧代码：wb.close() / app.quit() 直接写在函数末尾，异常时不会执行
    # 新代码：finally 块保证 Excel 进程一定被释放
    # ==============================================================================
    app = xw.App(visible=False)
    try:
        wb = app.books.open(file2)

        for tag in sheet_map:

            # ==============================================================================
            # 20260503修改：裸 except 改为 except Exception as e，避免掩盖真实错误
            # 旧代码：except:（捕获所有异常包括 KeyboardInterrupt，且不打印原因）
            # 新代码：except Exception as e，并把错误信息打印出来
            # ==============================================================================
            try:
                sht = wb.sheets[tag]
            except Exception as e:
                print(f"⚠ sheet不存在或打开失败: {tag}，原因: {e}")
                continue

            # 找写入行
            last_row = sht.used_range.last_cell.row
            write_row = last_row - 1

            headers = sht.range("A3").expand("right").value

            # ==============================================================================
            # 20260503修改：逐单元格写入改为先组装整行数据再一次性写入
            # 旧代码：for col in summary_cols: sht.cells(write_row, col_index).value = ...
            #         每次写一个单元格触发一次 COM 调用，列多时性能差
            # 新代码：先把所有要写的值组装到 row_values 列表，最后一次性 range.value 写入
            # ==============================================================================

            # 先把整行当前值读出来作为基础（保留公式列不被覆盖）
            total_cols = len(headers)
            row_values = list(sht.range(write_row, 1).resize(1, total_cols).value or [None] * total_cols)

            # 第一部分：标签汇总数据
            df_tag = df_summary[df_summary["listing标签"] == tag]

            if not df_tag.empty:
                row_data = df_tag.iloc[0]
                for col in summary_cols:
                    if col not in headers:
                        continue
                    # 旧代码：sht.cells(write_row, col_index).value = row_data[col]
                    row_values[headers.index(col)] = row_data[col]

            # 第二部分：品名数据
            df_tag_product = df_product[df_product["listing标签"] == tag]

            if not df_tag_product.empty:
                for _, row in df_tag_product.iterrows():
                    product_name = str(row["品名"]).strip()
                    sales = row["销量"]
                    refund_qty = row["退款量"]

                    # 写入销量列（例如 3.6FT）
                    if product_name in headers:
                        # 旧代码：sht.cells(write_row, col_index).value = sales
                        row_values[headers.index(product_name)] = sales

                    # 写入退款列（例如 退款3.6FT）
                    refund_col = f"退款{product_name}"
                    if refund_col in headers:
                        # 旧代码：sht.cells(write_row, col_index).value = refund_qty
                        row_values[headers.index(refund_col)] = refund_qty

            # 一次性写入整行（原来逐格写入改为单次 COM 调用）
            sht.range(write_row, 1).resize(1, total_cols).value = row_values

            print(f"✓ 完成 sheet: {tag}")

        wb.save()

    finally:
        # 旧代码：wb.close() / app.quit() 直接写在函数末尾，异常时不执行
        wb.close()
        app.quit()

    print("✓ 全部完成")
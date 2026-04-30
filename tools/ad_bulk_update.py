def run():
    import pandas as pd
    from datetime import datetime
    from openpyxl import load_workbook
    from openpyxl.styles import PatternFill
    import shutil
    import os

    # ========================
    # 基于当前文件定位项目根目录
    # ========================

    # 当前文件所在目录：tools/
    current_dir = os.path.dirname(os.path.abspath(__file__))

    # 项目根目录：amazon_toolkit/
    project_root = os.path.dirname(current_dir)

    input_path = os.path.join(
        project_root,
        "data/ad_bulk_update/input/BulkSheetExport-ASIN-高Acos出单和高点击不出单.xlsx"
    )

    output_path = os.path.join(
        project_root,
        "data/ad_bulk_update/output/BulkSheetExport-ASIN-高Acos出单和高点击不出单-processed.xlsx"
    )

    # ========================
    # 复制文件
    # ========================
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    shutil.copy(input_path, output_path)

    # ========================
    # 工具：列字母 → index
    # ========================
    def col_idx(col):
        num = 0
        for c in col:
            num = num * 26 + (ord(c.upper()) - ord('A') + 1)
        return num - 1

    # ========================
    # 读取数据
    # ========================
    df1 = pd.read_excel(output_path, sheet_name="Sponsored Products Campaigns")

    try:
        df2 = pd.read_excel(output_path, sheet_name="targeting_labels_SL_ASIN")
        sheet2_name = "targeting_labels_SL_ASIN"
        targeting_col_1 = col_idx("AK")
        filter_type = "Product Targeting"
    except:
        df2 = pd.read_excel(output_path, sheet_name="targeting_labels_SL_KW")
        sheet2_name = "targeting_labels_SL_KW"
        targeting_col_1 = col_idx("AC")
        filter_type = "Keyword"

    # 表1列
    col_B = col_idx("B")  # Entity（Product Targeting / Keyword）
    col_C = col_idx("C")  # Operation
    col_L = col_idx("L")  # Campaign Name (Informational only)
    col_M = col_idx("M")  # Ad Group Name (Informational only)
    col_AB = col_idx("AB")  # Bid

    # 表2列
    col_A2 = col_idx("A")  # Campaign Name
    col_B2 = col_idx("B")  # Ad Group Name
    col_C2 = col_idx("C")  # Targeting
    col_Q2 = col_idx("Q")  # adj_pct
    col_T2 = col_idx("T")  # 原竞价（Old Bid）
    col_U2 = col_idx("U")  # 新竞价（New Bid）
    col_V2 = col_idx("V")  # 操作日期（Operation Date）

    # ========================
    # 构建索引
    # ========================
    index_map = {}

    for i in range(len(df1)):
        if str(df1.iat[i, col_B]).strip() == filter_type:
            key = (
                str(df1.iat[i, col_L]).strip() + "|" +
                str(df1.iat[i, col_M]).strip() + "|" +
                str(df1.iat[i, targeting_col_1]).strip()
            )
            index_map[key] = i

    # ========================
    # 更新逻辑
    # ========================
    updated_rows = []

    for i in range(len(df2)):
        key = (
            str(df2.iat[i, col_A2]).strip() + "|" +
            str(df2.iat[i, col_B2]).strip() + "|" +
            str(df2.iat[i, col_C2]).strip()
        )

        if key in index_map:
            r = index_map[key]

            old_bid = df1.iat[r, col_AB]
            adj = df2.iat[i, col_Q2]

            if pd.isna(old_bid) or pd.isna(adj):
                continue

            # 防重复执行
            if str(df1.iat[r, col_C]).strip() == "Update":
                continue

            new_bid = round(old_bid * (1 + adj), 4)

            # 表2写入
            df2.iat[i, col_T2] = old_bid
            df2.iat[i, col_U2] = new_bid
            df2.iat[i, col_V2] = datetime.now().strftime("%Y-%m-%d")

            # 表1更新
            df1.iat[r, col_AB] = new_bid
            df1.iat[r, col_C] = "Update"

            updated_rows.append(r)

    # ========================
    # 写回
    # ========================
    with pd.ExcelWriter(output_path, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        df1.to_excel(writer, sheet_name="Sponsored Products Campaigns", index=False)
        df2.to_excel(writer, sheet_name=sheet2_name, index=False)

    # ========================
    # 高亮 + 删除Sheet
    # ========================
    wb = load_workbook(output_path)
    ws1 = wb["Sponsored Products Campaigns"]

    yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

    bid_col_excel = 28  # AB列

    for r in updated_rows:
        ws1.cell(row=r + 2, column=bid_col_excel).fill = yellow_fill

    for name in ["targeting_labels_SL_ASIN", "RAS Search Term Report"]:
        if name in wb.sheetnames:
            del wb[name]

    wb.save(output_path)

    print("✅ 完成：路径自适应 + 输出成功")
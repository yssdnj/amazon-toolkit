"""
ad_bulk_update.py
广告竞价批量更新工具

输入目录：amazon_toolkit/data/ad_bulk_update/input/
  - targeting_labels_{产品}_{ASIN|KW}_{日期}.xlsx  竞价更新指导文件（独立文件）
  - BulkSheetExport*.xlsx                          广告批量文件

输出目录：amazon_toolkit/data/ad_bulk_update/output/
  - BulkSheetExport*_updated.xlsx  更新竞价后的 Bulk 文件（已删除 RAS Search Term Report）

竞价指导文件原地回写 T列原竞价 / U列新竞价 / V列操作日期。
"""

import re
import shutil
import warnings
from datetime import datetime
from pathlib import Path

import pandas as pd

# 屏蔽 openpyxl 读取亚马逊导出文件时的"无默认样式"警告（不影响功能）
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')
from openpyxl import load_workbook
from openpyxl.styles import PatternFill


# ──────────────────────────────────────────────
# 工具函数
# ──────────────────────────────────────────────

def find_file(directory: Path, pattern: str) -> Path | None:
    """在目录中按正则匹配文件名，返回第一个匹配项，无则返回 None"""
    for f in sorted(directory.iterdir()):
        if re.search(pattern, f.name, re.IGNORECASE):
            return f
    return None


def detect_label_files(input_dir: Path) -> list:
    """
    扫描 input 目录，找出所有 targeting_labels 文件。
    返回 [(路径, 'ASIN'|'KW'), ...]
    """
    results = []
    for f in sorted(input_dir.iterdir()):
        m = re.search(r'targeting_labels_.+_(ASIN|KW)_[\d-]+\.xlsx', f.name, re.IGNORECASE)  # 支持 2026-04-20 或 20260420
        if m:
            label_type = m.group(1).upper()
            results.append((f, label_type))
    return results


# ──────────────────────────────────────────────
# 核心逻辑
# ──────────────────────────────────────────────

def load_label_df(label_path: Path):
    """
    读取竞价更新指导文件，筛选需处理的行。
    筛选条件：O列label含"高ACoS出单"或"高点击不出单"，且G列orders < 10
    返回 (筛选后df, 完整df)
    """
    df = pd.read_excel(label_path, dtype=str)
    df.columns = df.columns.str.strip()

    # G列 orders 转数值
    df['orders'] = pd.to_numeric(df['orders'], errors='coerce').fillna(0)

    # O列 label 筛选（含emoji前缀，用关键词匹配）
    label_col = df['label'].fillna('')
    mask_label = label_col.str.contains('高ACoS出单|高点击不出单', na=False)

    # G列 orders < 10
    mask_order = df['orders'] < 10

    filtered = df[mask_label & mask_order].copy()
    return filtered, df


def build_bulk_index(df_bulk: pd.DataFrame, label_type: str) -> dict:
    """
    从 Bulk SP Campaigns 表建立三元组查找索引。
    key = (Campaign Name, Ad Group Name, Targeting文本)
    value = df_bulk 的行索引（0-based）

    ASIN → Entity='Product Targeting'，Targeting列='Product Targeting Expression'
    KW   → Entity='Keyword'，          Targeting列='Keyword Text'
    """
    entity_filter  = 'Product Targeting' if label_type == 'ASIN' else 'Keyword'
    targeting_col  = 'Product Targeting Expression' if label_type == 'ASIN' else 'Keyword Text'

    index_map = {}
    for i, row in df_bulk.iterrows():
        if str(row.get('Entity', '')).strip() != entity_filter:
            continue
        key = (
            str(row.get('Campaign Name (Informational only)', '')).strip(),
            str(row.get('Ad Group Name (Informational only)', '')).strip(),
            str(row.get(targeting_col, '')).strip(),
        )
        if key not in index_map:
            index_map[key] = i
    return index_map


def process_updates(df_label_filtered, df_label_full, df_bulk, index_map, label_type, log):
    """
    遍历筛选出的竞价指导行，在 Bulk 中找匹配并更新竞价。
    - 更新 df_bulk 的 Bid 列
    - 回写 df_label_full 的 原竞价/新竞价/操作日期 列
    返回 (df_label_full, df_bulk, 被更新的bulk行索引列表)
    """
    updated_bulk_rows = []
    now_str = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    for label_idx, label_row in df_label_filtered.iterrows():
        camp  = str(label_row.get('Campaign Name', '')).strip()
        adgrp = str(label_row.get('Ad Group Name', '')).strip()
        tgt   = str(label_row.get('Targeting', '')).strip()
        key   = (camp, adgrp, tgt)

        # ── 找匹配行 ──
        if key not in index_map:
            log.append(f'[未匹配] {label_type} | {camp} | {adgrp} | {tgt}')
            continue

        bulk_idx = index_map[key]

        # ── 读原竞价 ──
        old_bid = pd.to_numeric(df_bulk.at[bulk_idx, 'Bid'], errors='coerce')
        if pd.isna(old_bid):
            log.append(f'[跳过-Bid为空] {label_type} | {camp} | {adgrp} | {tgt}')
            continue

        # ── 读 adj_pct ──
        adj_pct = pd.to_numeric(label_row.get('adj_pct'), errors='coerce')
        if pd.isna(adj_pct):
            log.append(f'[跳过-adj_pct为空] {label_type} | {camp} | {adgrp} | {tgt}')
            continue

        new_bid = round(float(old_bid) * (1 + float(adj_pct)), 4)

        # ── 更新 Bulk（df 以 dtype=str 读入，写回须转 str）──
        df_bulk.at[bulk_idx, 'Bid'] = str(new_bid)
        # 20260503需求补充：修改了 Bid 的行，C列 Operation 同步标记为 Update
        df_bulk.at[bulk_idx, 'Operation'] = 'Update'
        updated_bulk_rows.append(bulk_idx)

        # ── 回写竞价指导（同样 dtype=str，用 str 写入；openpyxl 写文件时再用 float）──
        df_label_full.at[label_idx, '原竞价']  = str(float(old_bid))
        df_label_full.at[label_idx, '新竞价']  = str(new_bid)
        df_label_full.at[label_idx, '操作日期'] = now_str

        log.append(
            f'[已更新] {label_type} | {camp} | {adgrp} | {tgt} '
            f'| {old_bid} → {new_bid} (adj={adj_pct:+.0%})'
        )

    return df_label_full, df_bulk, updated_bulk_rows


def save_bulk_updated(bulk_path: Path, output_path: Path, df_bulk: pd.DataFrame, updated_rows: list):
    """
    将更新后的 Bid 写回 Bulk 文件：
    - 复制原文件保留所有 sheet 和格式
    - 仅修改被更新的行的 Bid 单元格（不整列覆盖）
    - 高亮修改行（黄色）
    - 删除 RAS Search Term Report 表
    """
    shutil.copy(bulk_path, output_path)

    wb = load_workbook(output_path)
    ws = wb['Sponsored Products Campaigns']

    # 找 Bid 和 Operation 列（1-based）
    header_row = [cell.value for cell in ws[1]]
    try:
        bid_col_excel = header_row.index('Bid') + 1
    except ValueError:
        bid_col_excel = 28  # 默认 AB 列
    try:
        op_col_excel = header_row.index('Operation') + 1
    except ValueError:
        op_col_excel = 3   # 默认 C 列

    yellow_fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')

    # 只写被修改行的 Bid 和 Operation 单元格，其余行保持原文件不动
    for df_idx in updated_rows:
        excel_row = df_idx + 2  # +1 表头行, +1 转 1-based
        # Bid
        bid_val = df_bulk.at[df_idx, 'Bid']
        cell = ws.cell(row=excel_row, column=bid_col_excel)
        try:
            cell.value = float(bid_val)
        except (ValueError, TypeError):
            cell.value = bid_val
        cell.fill = yellow_fill
        # Operation = Update
        ws.cell(row=excel_row, column=op_col_excel).value = 'Update'

    # 删除 RAS Search Term Report
    if 'RAS Search Term Report' in wb.sheetnames:
        del wb['RAS Search Term Report']
        print('   🗑️  已删除 RAS Search Term Report 表')

    wb.save(output_path)


def save_label_updated(label_path: Path, df_label_full: pd.DataFrame):
    """
    将 原竞价/新竞价/操作日期 回写到竞价指导文件（原地保存）。
    按列名定位，列不存在时自动追加到末尾。
    """
    wb = load_workbook(label_path)
    ws = wb.active

    header = [cell.value for cell in ws[1]]

    def get_or_append_col(name):
        """找列，找不到就在表头末尾追加，返回 1-based 列号"""
        if name in header:
            return header.index(name) + 1
        new_col = len(header) + 1
        ws.cell(row=1, column=new_col, value=name)
        header.append(name)
        return new_col

    col_T = get_or_append_col('原竞价')
    col_U = get_or_append_col('新竞价')
    col_V = get_or_append_col('操作日期')

    for df_idx, row in df_label_full.iterrows():
        orig = row.get('原竞价')
        if pd.isna(orig) or orig == '' or orig == 'nan':
            continue
        excel_row = df_idx + 2  # +1 表头, +1 1-based
        # df 以 str 存储，写入 Excel 时转回 float 保留数值格式
        ws.cell(row=excel_row, column=col_T, value=float(orig))
        ws.cell(row=excel_row, column=col_U, value=float(row['新竞价']))
        ws.cell(row=excel_row, column=col_V, value=str(row['操作日期']))

    wb.save(label_path)


# ──────────────────────────────────────────────
# 入口
# ──────────────────────────────────────────────

def run():
    base_path  = Path(__file__).resolve().parents[1]
    input_dir  = base_path / 'data' / 'ad_bulk_update' / 'input'
    output_dir = base_path / 'data' / 'ad_bulk_update' / 'output'
    output_dir.mkdir(parents=True, exist_ok=True)

    log = []

    # ── 1. 找 BulkSheet 文件 ──
    bulk_path = find_file(input_dir, r'BulkSheetExport.*\.xlsx')
    if bulk_path is None:
        print('❌ 未找到 BulkSheetExport*.xlsx，请检查 input 目录')
        return
    print(f'📄 Bulk 文件: {bulk_path.name}')

    # ── 2. 找所有 targeting_labels 文件 ──
    label_files = detect_label_files(input_dir)
    if not label_files:
        print('❌ 未找到 targeting_labels_*_{ASIN|KW}_*.xlsx，请检查 input 目录')
        return
    if len(label_files) > 1:
        names = [f.name for f, _ in label_files]
        # 同时存在 ASIN 和 KW 两个文件是正常情况，仅当产品名不同时才警告
        products = set()
        for f, _ in label_files:
            m = re.search(r'targeting_labels_(.+?)_(ASIN|KW)_', f.name, re.IGNORECASE)
            if m:
                products.add(m.group(1))
        if len(products) > 1:
            msg = f'[警告] 发现多个不同产品的 targeting_labels 文件，将全部处理: {names}'
            log.append(msg)
            print(f'⚠️  {msg}')
        else:
            print(f'📋 同时处理 ASIN + KW 两份文件: {[f.name for f, _ in label_files]}')

    # ── 3. 读取 Bulk SP Campaigns 表（所有 label 文件共享同一份）──
    df_bulk = pd.read_excel(
        bulk_path,
        sheet_name='Sponsored Products Campaigns',
        dtype=str
    )
    df_bulk.columns = df_bulk.columns.str.strip()

    all_updated_rows = []

    # ── 4. 逐个 targeting_labels 文件处理 ──
    for label_path, label_type in label_files:
        print(f'\n🔍 处理: {label_path.name}（{label_type}）')

        df_label_filtered, df_label_full = load_label_df(label_path)

        n = len(df_label_filtered)
        print(f'   筛选结果: {n} 行（高ACoS出单/高点击不出单 且 orders<10）')

        if df_label_filtered.empty:
            log.append(f'[{label_type}] {label_path.name} 筛选结果为空，跳过')
            continue

        # 建立 Bulk 查找索引
        index_map = build_bulk_index(df_bulk, label_type)

        # 执行更新
        df_label_full, df_bulk, updated_rows = process_updates(
            df_label_filtered, df_label_full, df_bulk, index_map, label_type, log
        )
        all_updated_rows.extend(updated_rows)

        # 回写竞价指导文件（原地）
        save_label_updated(label_path, df_label_full)
        print(f'   ✅ 竞价指导文件已回写: {label_path.name}')

    # ── 5. 保存更新后的 Bulk 文件（加产品名和 _updated 后缀）──
    # 从 targeting_labels 文件名提取产品名，如 targeting_labels_SL_ASIN_... → _SL
    product_name = ''
    if label_files:
        m = re.search(r'targeting_labels_(.+?)_(ASIN|KW)_', label_files[0][0].name, re.IGNORECASE)
        if m:
            product_name = f'_{m.group(1)}'
    output_path = output_dir / f'{bulk_path.stem}{product_name}_updated.xlsx'
    save_bulk_updated(bulk_path, output_path, df_bulk, all_updated_rows)
    print(f'\n✅ Bulk 文件已保存: {output_path.name}')
    print(f'   共更新 {len(all_updated_rows)} 条竞价，已黄色高亮 Bid 单元格')

    # ── 6. 打印完整日志 ──
    print('\n─── 操作日志 ───')
    for line in log:
        print(f'  {line}')
    print('────────────────')
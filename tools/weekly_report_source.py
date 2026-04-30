import pandas as pd
from pathlib import Path

def process_xlsx(input_file, output_file):
    # 读取数据
    df = pd.read_excel(input_file, dtype=str).fillna('')
    
    # 列名清理，确保列名无空格
    df.columns = df.columns.str.strip()
    
    # 必须确保这些列存在，否则报错
    required_cols = ['listing标签','品名','订单量','销量','销售额','促销销量','促销订单量',
                     '促销销售额','退款量','退款金额','展示','点击','广告订单量','广告花费','广告销售额']
    for col in required_cols:
        if col not in df.columns:
            raise ValueError(f"缺少必要列: {col}")

    # 转换数值列为数值类型（先去逗号再转换）
    def to_num(series, is_int=False):
        s = series.str.replace(',','').replace('', '0')
        if is_int:
            return pd.to_numeric(s, errors='coerce').fillna(0).astype(int)
        else:
            return pd.to_numeric(s, errors='coerce').fillna(0).astype(float)
    
    int_cols = ['订单量','销量','促销销量','促销订单量','退款量','展示','点击','广告订单量']
    float_cols = ['销售额','促销销售额','退款金额','广告花费','广告销售额']
    
    for col in int_cols:
        df[col] = to_num(df[col].astype(str), True)
    for col in float_cols:
        df[col] = to_num(df[col].astype(str), False)
    
    # 品名去空格
    df['品名'] = df['品名'].str.strip()
    df['listing标签'] = df['listing标签'].str.strip()
    
    # 退款金额和广告花费转正数（绝对值）
    df['退款金额'] = df['退款金额'].abs()
    df['广告花费'] = df['广告花费'].abs()
    
    # --- 一级分组：按 listing标签 ---
    grouped = df.groupby('listing标签').agg({
        '订单量':'sum',
        '销量':'sum',
        '销售额':'sum',
        '促销销量':'sum',
        '促销订单量':'sum',
        '促销销售额':'sum',
        '退款量':'sum',
        '退款金额':'sum',
        '展示':'sum',
        '点击':'sum',
        '广告订单量':'sum',
        '广告花费':'sum',
        '广告销售额':'sum'
    }).reset_index()
    
    # 计算派生列
    grouped['退款率'] = grouped.apply(lambda r: r['退款量']/r['销量'] if r['销量']>0 else 0, axis=1)
    grouped['CPC'] = grouped.apply(lambda r: r['广告花费']/r['点击'] if r['点击']>0 else 0, axis=1)
    grouped['促销折扣'] = grouped.apply(lambda r: r['促销销售额']/r['促销销量'] if r['促销销量']>0 else 0, axis=1)
    
    # --- 二级分组规则 ---
    subgroup_rules = {
        'SL': ['3.6FT', '5.0FT', '5.5FT'],
        'Toy': ['Short w/oBall 1Pack', 'Long w/oBall 1Pack', 'w/Ball 1Pack', 'Long w/oBall 2Pack'],
        'ToyDH': ['1Pack', '2Pack'],
        'DL': ['w/oHandle AL', 'w/Handle AL','w/oHandle ZN','w/Handle ZN'],
        'DSL': ['3FT', '6FT'],
        'MFL': ['AL','ZN'],
        'SFM': ['SFM 1', 'SFM 2', 'SFM 3']
    }
    
    subgroup_results = []
    for listing_tag, names in subgroup_rules.items():
        df_filtered = df[df['listing标签'] == listing_tag]
        if df_filtered.empty:
            continue
        for name in names:
            # 模糊匹配品名包含规则中的name，忽略大小写
            mask = df_filtered['品名'].str.contains(name, case=False, na=False)
            df_sub = df_filtered[mask]
            if df_sub.empty:
                continue
            
            sub_group = df_sub.agg({
                '订单量':'sum',
                '销量':'sum',
                '销售额':'sum',
                '促销销量':'sum',
                '促销订单量':'sum',
                '促销销售额':'sum',
                '退款量':'sum',
                '退款金额':'sum',
                '展示':'sum',
                '点击':'sum',
                '广告订单量':'sum',
                '广告花费':'sum',
                '广告销售额':'sum'
            }).to_dict()
            
            # 计算派生列
            sub_group['退款率'] = sub_group['退款量']/sub_group['销量'] if sub_group['销量'] > 0 else 0
            sub_group['CPC'] = sub_group['广告花费']/sub_group['点击'] if sub_group['点击'] > 0 else 0
            sub_group['促销折扣'] = sub_group['促销销售额']/sub_group['促销销量'] if sub_group['促销销量'] > 0 else 0
            
            # 绝对值
            sub_group['退款金额'] = abs(sub_group['退款金额'])
            sub_group['广告花费'] = abs(sub_group['广告花费'])
            
            # 填充标签
            sub_group['listing标签'] = listing_tag
            sub_group['品名'] = name
            
            subgroup_results.append(sub_group)
    
    df_subgroups_all = pd.DataFrame(subgroup_results)
    
    # --- 数据格式化 ---
    def format_data(df_out):
        # 整数列
        for col in ['订单量','销量','促销销量','促销订单量','退款量','展示','点击','广告订单量']:
            if col in df_out.columns:
                df_out[col] = df_out[col].fillna(0).astype(int)
        # 保留两位小数列
        for col in ['销售额','促销销售额','促销折扣','退款金额','广告花费','广告销售额','CPC']:
            if col in df_out.columns:
                df_out[col] = df_out[col].fillna(0).round(2)
        # 百分比列，保留两位小数
        if '退款率' in df_out.columns:
            df_out['退款率'] = df_out['退款率'].fillna(0).round(4)  # 先四位小数
            # df_out['退款率'] = df_out['退款率'] * 100  # 转换成百分比
            # df_out['退款率'] = df_out['退款率'].round(2)
        return df_out
    
    grouped = format_data(grouped)
    df_subgroups_all = format_data(df_subgroups_all)
    
    
    # --- 导出 ---
    # 指定导出列顺序
    export_columns = [
        'listing标签', '品名', '销量', '订单量', '销售额', '促销销量', '促销订单量', '促销销售额', '促销折扣',
        '退款量', '退款率', '退款金额', '展示', '点击', '广告订单量', '广告花费', '广告销售额', 'CPC'
    ]

    # 只保留并按顺序导出指定列（如果有些列不存在则自动跳过）
    grouped_export = grouped[[col for col in export_columns if col in grouped.columns]]
    df_subgroups_all_export = df_subgroups_all[[col for col in export_columns if col in df_subgroups_all.columns]]

    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        grouped_export.to_excel(writer, sheet_name='标签汇总sht', index=False)
        if not df_subgroups_all_export.empty:
            df_subgroups_all_export.to_excel(writer, sheet_name='标签品名汇总sht', index=False)
        
        # 设置Excel格式（整数、浮点、百分比）
        workbook  = writer.book
        int_fmt = workbook.add_format({'num_format': '0'})
        float_fmt = workbook.add_format({'num_format': '0.00'})
        pct_fmt = workbook.add_format({'num_format': '0.00%'})
        
        # 格式化第一个sheet
        ws1 = writer.sheets['标签汇总sht']
        fmt_map = {
            '订单量': int_fmt, '销量': int_fmt,  '销售额': float_fmt,
            '促销销量': int_fmt, '促销订单量': int_fmt, '促销销售额': float_fmt, '促销折扣': float_fmt, 
            '退款量': int_fmt, '退款率': pct_fmt,'退款金额': float_fmt,
            '展示': int_fmt, '点击': int_fmt, '广告订单量': int_fmt, 
            '广告花费': float_fmt, '广告销售额': float_fmt, 'CPC': float_fmt
        }
        for col_num, col_name in enumerate(grouped_export.columns):
            if col_name in fmt_map:
                ws1.set_column(col_num, col_num, 15, fmt_map[col_name])
            else:
                ws1.set_column(col_num, col_num, 20)
        
        # 格式化第二个sheet
        if not df_subgroups_all_export.empty:
            ws2 = writer.sheets['标签品名汇总sht']
            for col_num, col_name in enumerate(df_subgroups_all_export.columns):
                if col_name in fmt_map:
                    ws2.set_column(col_num, col_num, 15, fmt_map[col_name])
                else:
                    ws2.set_column(col_num, col_num, 20)

    print(f"处理完成，结果已保存至 {output_file}")

def run():
    base_path = Path(__file__).resolve().parents[1]
    data_dir = base_path / "data" / "weekly_report"
    input_dir = data_dir / "input"
    output_dir = data_dir / "output"

    input_dir.mkdir(parents=True, exist_ok=True)
    output_dir.mkdir(parents=True, exist_ok=True)

    print(f"\n📂 输入文件目录: {input_dir}")
    print(f"📂 输出文件目录: {output_dir}")

    # input_file = input("请输入输入文件名（如 产品表现MSKU-W39.xlsx）: ").strip()
    # output_file = input("请输入输出文件名（如 output_aggregated-W39.xlsx）: ").strip()
    week_number = input("请输入输入周数（如 26W11）: ").strip()
    input_file = '产品表现MSKU-' + week_number + '.xlsx'
    output_file = 'output_aggregated-' + week_number + '.xlsx'

    input_path = input_dir / input_file
    output_path = output_dir / output_file

    if not input_path.exists():
        print(f"❌ 文件不存在: {input_path}")
        return

    process_xlsx(input_path, output_path)
    print(f"✅ 周报已生成：{output_path}")






import streamlit as st
import pandas as pd
import io
import re

st.set_page_config(layout="wide", page_title="Coupang 利润核算 (手动启动版)")
st.title("🔘 步骤五：多店铺利润核算 (手动启动模式)")
st.markdown("### 操作流程：上传文件 -> 确认就绪 -> **点击按钮** -> 生成报表")

# ==========================================
# 1. 上传区域
# ==========================================
with st.sidebar:
    st.header("1. 文件上传区")
    file_master = st.file_uploader("基础信息表 (Master - 1个)", type=['csv', 'xlsx'])
    files_sales = st.file_uploader("销售表 (Sales - 支持多个)", type=['csv', 'xlsx'], accept_multiple_files=True)
    files_ads = st.file_uploader("广告表 (Ads - 支持多个)", type=['csv', 'xlsx'], accept_multiple_files=True)

    st.markdown("---")
    # 侧边栏状态提示
    if file_master and files_sales and files_ads:
        st.success("✅ 所有文件已上传，请去右侧点击按钮开始。")
    else:
        st.info("⏳ 等待文件上传完整...")

# ==========================================
# 2. 工具函数
# ==========================================
def clean_id(series):
    return series.astype(str).str.replace(r'\.0$', '', regex=True).str.replace('"', '').str.replace('\n', '').str.strip()
def clean_num(series):
    return pd.to_numeric(series, errors='coerce').fillna(0)
def extract_product_code(text):
    if pd.isna(text): return None
    match = re.search(r'([Cc]\d+)', str(text))
    if match: return match.group(1).upper()
    return None

def read_and_combine(file_list, file_type_name=""):
    if not file_list: return pd.DataFrame()
    all_dfs = []
    for file in file_list:
        try:
            file.seek(0)
            if file.name.endswith('.csv'):
                try: df = pd.read_csv(file)
                except: file.seek(0); df = pd.read_csv(file, encoding='gbk')
            else: df = pd.read_excel(file)
            all_dfs.append(df)
        except Exception as e: st.error(f"❌ {file.name} 读取失败: {e}")
    
    if all_dfs:
        combined = pd.concat(all_dfs, ignore_index=True)
        rows_before = len(combined)
        combined.drop_duplicates(inplace=True)
        rows_after = len(combined)
        removed = rows_before - rows_after
        if removed > 0: st.warning(f"⚠️ 【{file_type_name}】剔除了 {removed} 条重复数据")
        return combined
    return pd.DataFrame()

# ==========================================
# 3. 主界面逻辑 (带按钮控制)
# ==========================================

# 只有当三个文件都存在时，才显示“开始按钮”
if file_master and files_sales and files_ads:
    
    st.divider()
    col1, col2 = st.columns([3, 1])
    with col1:
        st.subheader("📂 文件状态确认")
        st.write(f"• 基础表：1 个")
        st.write(f"• 销售表：{len(files_sales)} 个 (待合并)")
        st.write(f"• 广告表：{len(files_ads)} 个 (待合并)")
    
    with col2:
        st.write("##") # 占位符，让按钮下沉对齐
        # type='primary' 让按钮变成红色醒目款
        start_btn = st.button("🚀 点击开始计算", type="primary", use_container_width=True)

    if start_btn:
        st.divider()
        with st.status("🔄 正在全速计算中...", expanded=True):
            try:
                # A. Master
                st.write("正在处理基础表...")
                if file_master.name.endswith('.csv'): df_master = pd.read_csv(file_master)
                else: df_master = pd.read_excel(file_master)
                df_master['__ORDER__'] = range(len(df_master))

                col_code = df_master.columns[0]; col_sku = df_master.columns[3]; col_profit = df_master.columns[10]
                df_master['关联ID'] = clean_id(df_master[col_sku])
                df_master['单件毛利'] = clean_num(df_master[col_profit])
                df_master['产品编号_清洗'] = clean_id(df_master[col_code]).str.upper()

                # B. Sales
                st.write("正在合并并清洗销售数据...")
                df_sales_all = read_and_combine(files_sales, "销售表")
                col_sales_id = df_sales_all.columns[0]; col_sales_qty = df_sales_all.columns[8]
                df_sales_all['关联ID'] = clean_id(df_sales_all[col_sales_id])
                df_sales_all['销量'] = clean_num(df_sales_all[col_sales_qty])
                sales_agg = df_sales_all.groupby('关联ID')['销量'].sum().reset_index()
                sales_agg.rename(columns={'销量': 'O列_合并销量'}, inplace=True)

                # C. Ads
                st.write("正在匹配广告花费...")
                df_ads_all = read_and_combine(files_ads, "广告表")
                col_campaign = df_ads_all.columns[5]; col_ad_spend = df_ads_all.columns[15]
                df_ads_all['提取编号'] = df_ads_all[col_campaign].apply(extract_product_code)
                df_ads_all['含税广告费'] = clean_num(df_ads_all[col_ad_spend]) * 1.1
                ads_agg = df_ads_all.groupby('提取编号')['含税广告费'].sum().reset_index()
                ads_agg.rename(columns={'提取编号': '产品编号_清洗', '含税广告费': 'R列_产品总广告费'}, inplace=True)

                # D. Merge
                st.write("正在生成最终报表...")
                df_final = pd.merge(df_master, sales_agg, on='关联ID', how='left')
                df_final['O列_合并销量'] = df_final['O列_合并销量'].fillna(0).astype(int)
                df_final['P列_SKU总毛利'] = df_final['O列_合并销量'] * df_final['单件毛利']
                df_final['Q列_产品总利润'] = df_final.groupby('产品编号_清洗')['P列_SKU总毛利'].transform('sum')
                df_final = pd.merge(df_final, ads_agg, on='产品编号_清洗', how='left')
                df_final['R列_产品总广告费'] = df_final['R列_产品总广告费'].fillna(0)
                df_final['S列_最终净利润'] = df_final['Q列_产品总利润'] - df_final['R列_产品总广告费']

                df_final.sort_values(by=['__ORDER__'], inplace=True)
                df_final.drop(columns=['__ORDER__', '关联ID', '单件毛利', '产品编号_清洗', '提取编号'], inplace=True, errors='ignore')

                # E. Excel Output
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    wb = writer.book
                    
                    # Sheet 1
                    df_final.to_excel(writer, index=False, sheet_name='1_超级数据源')
                    ws1 = writer.sheets['1_超级数据源']
                    (mr, mc) = df_final.shape
                    cols_settings = [{'header': c} for c in df_final.columns]
                    ws1.add_table(0, 0, mr, mc-1, {'columns': cols_settings, 'name': 'Data', 'style': 'TableStyleMedium9'})
                    ws1.set_column(0, mc-1, 15)

                    # Sheet 2
                    df_final.to_excel(writer, index=False, sheet_name='2_老板视图')
                    ws2 = writer.sheets['2_老板视图']
                    
                    merge_fmt = wb.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'fg_color': '#FFFFFF'})
                    green_fmt = wb.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#C6EFCE', 'font_color': '#006100'})
                    red_fmt = wb.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#FFC7CE', 'font_color': '#9C0006'})
                    
                    ws2.set_column('A:A', 15)
                    cols = df_final.columns.tolist()
                    idx_A = 0
                    idx_Q = cols.index('Q列_产品总利润') if 'Q列_产品总利润' in cols else -1
                    idx_R = cols.index('R列_产品总广告费') if 'R列_产品总广告费' in cols else -1
                    idx_S = cols.index('S列_最终净利润') if 'S列_最终净利润' in cols else -1

                    start_row = 1
                    codes = df_final.iloc[:, 0].astype(str).tolist()
                    q_vals = df_final['Q列_产品总利润'].tolist()
                    r_vals = df_final['R列_产品总广告费'].tolist()
                    s_vals = df_final['S列_最终净利润'].tolist()

                    for i in range(1, len(codes) + 1):
                        if i == len(codes) or codes[i] != codes[i-1]:
                            profit = s_vals[start_row-1]
                            s_fmt = green_fmt if profit >= 0 else red_fmt
                            cnt = i - start_row
                            if cnt > 1:
                                ws2.merge_range(start_row, idx_A, i, idx_A, codes[start_row-1], merge_fmt)
                                ws2.merge_range(start_row, idx_Q, i, idx_Q, q_vals[start_row-1], merge_fmt)
                                ws2.merge_range(start_row, idx_R, i, idx_R, r_vals[start_row-1], merge_fmt)
                                ws2.merge_range(start_row, idx_S, i, idx_S, profit, s_fmt)
                            else:
                                ws2.write(start_row, idx_A, codes[start_row-1], merge_fmt)
                                ws2.write(start_row, idx_Q, q_vals[start_row-1], merge_fmt)
                                ws2.write(start_row, idx_R, r_vals[start_row-1], merge_fmt)
                                ws2.write(start_row, idx_S, profit, s_fmt)
                            start_row = i + 1
                    
                st.success("✅ 计算完成！")
                st.download_button("📥 下载结果报表", output.getvalue(), "Coupang_Final_Result.xlsx", "application/vnd.ms-excel", type='primary')

            except Exception as e:
                st.error(f"发生错误: {e}")

else:
    # 初始状态提示
    st.info("👈 请在左侧上传文件：1个基础表 + 多个销售/广告表。上传完成后，此处会出现开始按钮。") 我找到了之前的代码，在这个基础上改
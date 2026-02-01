import streamlit as st
import pandas as pd
import io
import re

# ==========================================
# 1. 页面配置
# ==========================================
st.set_page_config(layout="wide", page_title="Coupang 利润核算 (网页看板版)")
st.title("📊 最终定稿：利润核算 (网页预览 + 下载)")
st.markdown("""
### 🎯 功能升级：
* **网页看板**：计算完成后，直接在网页下方预览“明细表”和“业务报表”。
* **双重匹配**：G列广告组为主，F列活动名为辅。
* **数据清洗**：自动去除千分位逗号，修复数值异常。
""")

# --- 列号配置 ---
IDX_M_CODE   = 0    # Master表 A列: 内部编码
IDX_M_SKU    = 3    # Master表 D列: SKU ID
IDX_M_PROFIT = 10   # Master表 K列: 单品毛利

IDX_S_ID     = 0    # Sales表 A列: 选项ID
IDX_S_QTY    = 8    # Sales表 I列: 购买数量

# 广告表配置
IDX_A_CAMPAIGN = 5  # Ads表 F列: 广告活动名 (兜底)
IDX_A_GROUP    = 6  # Ads表 G列: 广告组 (首选)
IDX_A_SPEND    = 15 # Ads表 P列: 广告费
# -----------------

# ==========================================
# 2. 上传区域
# ==========================================
with st.sidebar:
    st.header("📂 文件上传")
    st.info("基础表 1 个，销售/广告表支持多个")
    file_master = st.file_uploader("1. 基础信息表 (Master)", type=['csv', 'xlsx', 'xlsm'])
    files_sales = st.file_uploader("2. 销售表 (Sales - 多选)", type=['csv', 'xlsx', 'xlsm'], accept_multiple_files=True)
    files_ads = st.file_uploader("3. 广告表 (Ads - 多选)", type=['csv', 'xlsx', 'xlsm'], accept_multiple_files=True)

# ==========================================
# 3. 清洗工具
# ==========================================
def clean_for_match(series):
    return series.astype(str).str.replace(r'\.0$', '', regex=True).str.replace('"', '').str.strip().str.upper()

def clean_num(series):
    return pd.to_numeric(series.astype(str).str.replace(',', ''), errors='coerce').fillna(0)

def extract_code_from_text(text):
    if pd.isna(text): return None
    match = re.search(r'([Cc]\d+)', str(text))
    if match: return match.group(1).upper()
    return None

def read_file_strict(file):
    try:
        file.seek(0)
        if file.name.endswith('.csv'):
            return pd.read_csv(file, dtype=str)
        else:
            return pd.read_excel(file, dtype=str, engine='openpyxl')
    except:
        file.seek(0)
        return pd.read_csv(file, dtype=str, encoding='gbk')

def get_col_width(series):
    return series.astype(str).map(len).max()

# ==========================================
# 4. 主逻辑
# ==========================================
if file_master and files_sales and files_ads:
    st.divider()
    if st.button("🚀 开始计算 & 生成看板", type="primary", use_container_width=True):
        try:
            with st.status("🔄 正在清洗与计算...", expanded=True):
                
                # --------------------------------------------
                # Step 1: 基础表
                # --------------------------------------------
                st.write("1. 读取基础表...")
                df_master = read_file_strict(file_master)
                col_code_name = df_master.columns[IDX_M_CODE]

                df_master['_MATCH_SKU'] = clean_for_match(df_master.iloc[:, IDX_M_SKU])
                df_master['_MATCH_CODE'] = clean_for_match(df_master.iloc[:, IDX_M_CODE])
                df_master['_VAL_PROFIT'] = clean_num(df_master.iloc[:, IDX_M_PROFIT])

                # --------------------------------------------
                # Step 2: 销售表
                # --------------------------------------------
                st.write(f"2. 合并 {len(files_sales)} 个销售表...")
                sales_list = [read_file_strict(f) for f in files_sales]
                df_sales_all = pd.concat(sales_list, ignore_index=True)
                
                df_sales_all['_MATCH_SKU'] = clean_for_match(df_sales_all.iloc[:, IDX_S_ID])
                df_sales_all['销量'] = clean_num(df_sales_all.iloc[:, IDX_S_QTY])
                
                sales_agg = df_sales_all.groupby('_MATCH_SKU')['销量'].sum().reset_index()
                sales_agg.rename(columns={'销量': 'O列_合并销量'}, inplace=True)

                # --------------------------------------------
                # Step 3: 广告表 (双重匹配)
                # --------------------------------------------
                st.write(f"3. 合并 {len(files_ads)} 个广告表...")
                ads_list = [read_file_strict(f) for f in files_ads]
                df_ads_all = pd.concat(ads_list, ignore_index=True)

                # 清洗 & 双重提取
                df_ads_all['含税广告费'] = clean_num(df_ads_all.iloc[:, IDX_A_SPEND]) * 1.1
                df_ads_all['Code_Group'] = df_ads_all.iloc[:, IDX_A_GROUP].apply(extract_code_from_text)
                df_ads_all['Code_Campaign'] = df_ads_all.iloc[:, IDX_A_CAMPAIGN].apply(extract_code_from_text)
                df_ads_all['_MATCH_CODE'] = df_ads_all['Code_Group'].fillna(df_ads_all['Code_Campaign'])

                # 匹配自检
                st.info("🔍 **广告匹配自检 (前5条):**")
                debug_df = df_ads_all[['_MATCH_CODE', '含税广告费']].dropna().head()
                if not debug_df.empty: st.dataframe(debug_df)
                
                # 聚合
                valid_ads = df_ads_all.dropna(subset=['_MATCH_CODE'])
                ads_agg = valid_ads.groupby('_MATCH_CODE')['含税广告费'].sum().reset_index()
                ads_agg.rename(columns={'含税广告费': 'R列_产品总广告费'}, inplace=True)

                total = df_ads_all['含税广告费'].sum()
                matched = ads_agg['R列_产品总广告费'].sum()
                st.success(f"💰 广告匹配统计：总额 {total:,.0f} | 成功 {matched:,.0f} ({(matched/total if total>0 else 0):.1%})")

                # --------------------------------------------
                # Step 4: 关联计算
                # --------------------------------------------
                df_final = pd.merge(df_master, sales_agg, on='_MATCH_SKU', how='left', sort=False)
                df_final['O列_合并销量'] = df_final['O列_合并销量'].fillna(0).astype(int)
                
                df_final['P列_SKU总毛利'] = df_final['O列_合并销量'] * df_final['_VAL_PROFIT']
                df_final['Q列_产品总利润'] = df_final.groupby('_MATCH_CODE', sort=False)['P列_SKU总毛利'].transform('sum')
                
                df_final = pd.merge(df_final, ads_agg, on='_MATCH_CODE', how='left', sort=False)
                df_final['R列_产品总广告费'] = df_final['R列_产品总广告费'].fillna(0)
                
                df_final['S列_最终净利润'] = df_final['Q列_产品总利润'] - df_final['R列_产品总广告费']

                # --------------------------------------------
                # Step 5: 输出准备
                # --------------------------------------------
                df_sheet2 = df_final[[col_code_name, 'Q列_产品总利润', 'R列_产品总广告费', 'S列_最终净利润']].copy()
                df_sheet2 = df_sheet2.drop_duplicates(subset=[col_code_name], keep='first')
                
                # 防爆修复：转str后再查 startswith
                cols_to_drop = [c for c in df_final.columns if str(c).startswith('_') or str(c).startswith('Code_')]
                df_final.drop(columns=cols_to_drop, inplace=True)

                # ==========================================
                # 新增功能：网页看板展示
                # ==========================================
                st.divider()
                st.subheader("📈 报表预览看板")
                
                tab1, tab2 = st.tabs(["📝 1. 利润分析 (明细)", "📊 2. 业务报表 (汇总)"])
                
                with tab1:
                    st.caption("展示所有 SKU 的详细利润情况")
                    # 高亮 S列 (净利润) 小于 0 的行
                    st.dataframe(
                        df_final.style.format(precision=0).background_gradient(subset=['S列_最终净利润'], cmap='RdYlGn', vmin=-10000, vmax=10000),
                        use_container_width=True
                    )
                
                with tab2:
                    st.caption("展示按产品归集的最终经营结果")
                    # 高亮净利润
                    st.dataframe(
                        df_sheet2.style.format(precision=0).background_gradient(subset=['S列_最终净利润'], cmap='RdYlGn', vmin=-10000, vmax=10000),
                        use_container_width=True
                    )

                # ==========================================
                # Excel 下载逻辑
                # ==========================================
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    
                    # Sheet 1
                    df_final.to_excel(writer, index=False, sheet_name='利润分析')
                    wb = writer.book
                    ws = writer.sheets['利润分析']
                    
                    base_font = {'font_name': 'Microsoft YaHei', 'bold': True, 'border': 1, 'align': 'center', 'valign': 'vcenter'}
                    fmt_row_grey = wb.add_format(dict(base_font, bg_color='#BFBFBF'))
                    fmt_row_white = wb.add_format(dict(base_font, bg_color='#FFFFFF'))
                    fmt_s_profit = wb.add_format(dict(base_font, bg_color='#C6EFCE'))
                    fmt_s_loss = wb.add_format(dict(base_font, bg_color='#FFC7CE'))

                    for i, col in enumerate(df_final.columns):
                        max_len = get_col_width(df_final[col])
                        header_len = len(str(col)) * 1.5
                        final_width = max(max_len, header_len) + 2
                        ws.set_column(i, i, min(max(final_width, 10), 50))
                    ws.freeze_panes(1, 0)

                    col_code_idx = IDX_M_CODE 
                    cols_list = df_final.columns.tolist()
                    col_profit_idx = cols_list.index('S列_最终净利润') if 'S列_最终净利润' in cols_list else -1
                    raw_codes = df_final.iloc[:, col_code_idx].astype(str).tolist()
                    clean_codes = [str(x).replace('.0','').replace('"','').strip().upper() for x in raw_codes]
                    
                    is_grey = False
                    for i in range(len(raw_codes)):
                        excel_row = i + 1
                        if i > 0 and clean_codes[i] != clean_codes[i-1]: is_grey = not is_grey
                        ws.set_row(excel_row, None, fmt_row_grey if is_grey else fmt_row_white)
                        if col_profit_idx != -1:
                            val = df_final.iloc[i, col_profit_idx]
                            try: num_val = float(val)
                            except: num_val = 0
                            if num_val > 0: ws.write(excel_row, col_profit_idx, val, fmt_s_profit)
                            elif num_val < 0: ws.write(excel_row, col_profit_idx, val, fmt_s_loss)
                            else: ws.write(excel_row, col_profit_idx, val, fmt_row_grey if is_grey else fmt_row_white)

                    # Sheet 2
                    df_sheet2.to_excel(writer, index=False, sheet_name='业务报表')
                    ws2 = writer.sheets['业务报表']
                    fmt_header2 = wb.add_format({'font_name': 'Microsoft YaHei', 'bold': True, 'font_size': 12, 'bg_color': '#4472C4', 'font_color': 'white', 'border': 1, 'align': 'center'})
                    fmt_money2 = wb.add_format({'font_name': 'Microsoft YaHei', 'font_size': 11, 'border': 1, 'align': 'center', 'valign': 'vcenter', 'num_format': '#,##0'})
                    
                    for col_num, value in enumerate(df_sheet2.columns.values): ws2.write(0, col_num, value, fmt_header2)
                    ws2.set_column(0, 0, 25)
                    ws2.set_column(1, 3, 18, fmt_money2)
                    ws2.freeze_panes(1, 0)
                    (max_r2, max_c2) = df_sheet2.shape
                    ws2.conditional_format(1, 3, max_r2, 3, {'type': 'data_bar', 'bar_color': '#63C384', 'bar_negative_color': '#FF0000', 'bar_axis_position': 'middle'})

            st.success("✅ 计算完成！请在上方标签页预览数据，或点击下方按钮下载。")
            st.download_button("📥 下载完整报表", output.getvalue(), "Coupang_Dashboard_Report.xlsx")

        except Exception as e:
            st.error(f"❌ 错误: {e}")
else:
    st.info("👈 请上传所有必需文件")

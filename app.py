import streamlit as st
import pandas as pd
import io
import re

# ==========================================
# 1. 页面配置 (宽屏)
# ==========================================
st.set_page_config(layout="wide", page_title="Coupang 经营看板 Pro")
st.title("📊 Coupang 经营分析看板 (最终版)")

# --- 列号配置 ---
IDX_M_CODE   = 0    # Master A列
IDX_M_SKU    = 3    # Master D列
IDX_M_PROFIT = 10   # Master K列

IDX_S_ID     = 0    # Sales A列
IDX_S_QTY    = 8    # Sales I列

IDX_A_CAMPAIGN = 5  # Ads F列
IDX_A_GROUP    = 6  # Ads G列
IDX_A_SPEND    = 15 # Ads P列
# -----------------

# ==========================================
# 2. 侧边栏上传
# ==========================================
with st.sidebar:
    st.header("📂 数据源上传")
    file_master = st.file_uploader("1. 基础信息表 (Master)", type=['csv', 'xlsx', 'xlsm'])
    files_sales = st.file_uploader("2. 销售表 (Sales)", type=['csv', 'xlsx', 'xlsm'], accept_multiple_files=True)
    files_ads = st.file_uploader("3. 广告表 (Ads)", type=['csv', 'xlsx', 'xlsm'], accept_multiple_files=True)

# ==========================================
# 3. 清洗工具函数
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

# ==========================================
# 4. 主逻辑
# ==========================================
if file_master and files_sales and files_ads:
    st.divider()
    
    if st.button("🚀 生成看板 & 准备下载", type="primary", use_container_width=True):
        try:
            with st.spinner("正在处理多店数据..."):
                
                # --- Step 1: 基础表 ---
                df_master = read_file_strict(file_master)
                col_code_name = df_master.columns[IDX_M_CODE]

                df_master['_MATCH_SKU'] = clean_for_match(df_master.iloc[:, IDX_M_SKU])
                df_master['_MATCH_CODE'] = clean_for_match(df_master.iloc[:, IDX_M_CODE])
                df_master['_VAL_PROFIT'] = clean_num(df_master.iloc[:, IDX_M_PROFIT])

                # --- Step 2: 销售表 ---
                sales_list = [read_file_strict(f) for f in files_sales]
                df_sales_all = pd.concat(sales_list, ignore_index=True)
                
                df_sales_all['_MATCH_SKU'] = clean_for_match(df_sales_all.iloc[:, IDX_S_ID])
                df_sales_all['销量'] = clean_num(df_sales_all.iloc[:, IDX_S_QTY])
                
                sales_agg = df_sales_all.groupby('_MATCH_SKU')['销量'].sum().reset_index()
                sales_agg.rename(columns={'销量': 'O列_合并销量'}, inplace=True)

                # --- Step 3: 广告表 ---
                ads_list = [read_file_strict(f) for f in files_ads]
                df_ads_all = pd.concat(ads_list, ignore_index=True)

                df_ads_all['含税广告费'] = clean_num(df_ads_all.iloc[:, IDX_A_SPEND]) * 1.1
                df_ads_all['Code_Group'] = df_ads_all.iloc[:, IDX_A_GROUP].apply(extract_code_from_text)
                df_ads_all['Code_Campaign'] = df_ads_all.iloc[:, IDX_A_CAMPAIGN].apply(extract_code_from_text)
                df_ads_all['_MATCH_CODE'] = df_ads_all['Code_Group'].fillna(df_ads_all['Code_Campaign'])

                valid_ads = df_ads_all.dropna(subset=['_MATCH_CODE'])
                ads_agg = valid_ads.groupby('_MATCH_CODE')['含税广告费'].sum().reset_index()
                ads_agg.rename(columns={'含税广告费': 'R列_产品总广告费'}, inplace=True)

                # --- Step 4: 关联 & 计算 ---
                df_final = pd.merge(df_master, sales_agg, on='_MATCH_SKU', how='left', sort=False)
                df_final['O列_合并销量'] = df_final['O列_合并销量'].fillna(0).astype(int)
                
                # 1. 算 SKU 毛利
                df_final['P列_SKU总毛利'] = df_final['O列_合并销量'] * df_final['_VAL_PROFIT']
                
                # 2. 算 产品总利润 (汇总)
                df_final['Q列_产品总利润'] = df_final.groupby('_MATCH_CODE', sort=False)['P列_SKU总毛利'].transform('sum')
                
                # 3. 算 产品总销量 (新增功能!)
                df_final['产品总销量'] = df_final.groupby('_MATCH_CODE', sort=False)['O列_合并销量'].transform('sum')
                
                # 4. 关联广告
                df_final = pd.merge(df_final, ads_agg, on='_MATCH_CODE', how='left', sort=False)
                df_final['R列_产品总广告费'] = df_final['R列_产品总广告费'].fillna(0)
                
                # 5. 算净利
                df_final['S列_最终净利润'] = df_final['Q列_产品总利润'] - df_final['R列_产品总广告费']

                # --- Step 5: Sheet2 逻辑 (加入销量) ---
                # 选取列：加入 '产品总销量'
                df_sheet2 = df_final[[col_code_name, '产品总销量', 'Q列_产品总利润', 'R列_产品总广告费', 'S列_最终净利润']].copy()
                df_sheet2 = df_sheet2.drop_duplicates(subset=[col_code_name], keep='first')
                
                # 计算比值
                df_sheet2['广告/毛利比'] = df_sheet2.apply(
                    lambda x: x['R列_产品总广告费'] / x['Q列_产品总利润'] if x['Q列_产品总利润'] != 0 else 0, 
                    axis=1
                )

                # --- Step 6: 清理 ---
                cols_to_drop = [c for c in df_final.columns if str(c).startswith('_') or str(c).startswith('Code_')]
                df_final.drop(columns=cols_to_drop, inplace=True)

                # ==========================================
                # 🔥 看板展示
                # ==========================================
                
                # KPI
                total_qty = df_sheet2['产品总销量'].sum()
                total_profit = df_sheet2['Q列_产品总利润'].sum()
                total_ads = df_sheet2['R列_产品总广告费'].sum()
                net_profit = df_sheet2['S列_最终净利润'].sum()
                
                st.subheader("📈 经营概览")
                k1, k2, k3, k4 = st.columns(4)
                k1.metric("💰 最终净利润", f"{net_profit:,.0f}")
                k2.metric("📦 总销售数量", f"{total_qty:,.0f}") # 这里改成显示总销量
                k3.metric("📢 总广告费", f"{total_ads:,.0f}")
                k4.metric("📉 整体广告比", f"{(total_ads/total_profit if total_profit else 0):.1%}")

                st.divider()

                # Tab 分页
                tab1, tab2 = st.tabs(["📝 1. 利润明细 (查账)", "📊 2. 业务报表 (汇报)"])
                
                # 尝试应用颜色样式
                def try_style(df, cols, is_pct=False):
                    try:
                        styler = df.style.format(precision=0)
                        if is_pct:
                            styler = styler.format({'广告/毛利比': '{:.1%}'})
                        return styler.background_gradient(subset=cols, cmap='RdYlGn', vmin=-10000, vmax=10000)
                    except:
                        return df

                with tab1:
                    st.caption("🔍 明细数据")
                    st.dataframe(
                        try_style(df_final, ['S列_最终净利润']),
                        use_container_width=True,
                        height=800
                    )
                
                with tab2:
                    st.caption("🏆 汇总数据 (含新增列：产品总销量)")
                    st.dataframe(
                        try_style(df_sheet2, ['S列_最终净利润'], is_pct=True),
                        use_container_width=True,
                        height=800
                    )

                # ==========================================
                # 📥 总下载逻辑
                # ==========================================
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    # Sheet1
                    df_final.to_excel(writer, index=False, sheet_name='利润分析')
                    
                    # Sheet2
                    df_sheet2.to_excel(writer, index=False, sheet_name='业务报表')
                    
                    # Excel 美化
                    wb = writer.book
                    ws2 = writer.sheets['业务报表']
                    fmt_pct = wb.add_format({'num_format': '0.0%'})
                    ws2.set_column(5, 5, 15, fmt_pct) # 第6列(F列)是广告比

                st.divider()
                st.success("✅ 报表已生成！")
                
                st.download_button(
                    label="📥 一键下载完整报表 (含销量统计)",
                    data=output.getvalue(),
                    file_name="Coupang_Final_Report_v2.xlsx",
                    mime="application/vnd.ms-excel",
                    type="primary",
                    use_container_width=True
                )

        except Exception as e:
            st.error(f"❌ 运行出错: {e}")
else:
    st.info("👈 请上传文件")

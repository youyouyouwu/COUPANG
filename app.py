import streamlit as st
import pandas as pd
import io
import re

# ==========================================
# 1. 页面配置 (开启宽屏模式)
# ==========================================
st.set_page_config(layout="wide", page_title="Coupang 经营看板 Pro")

st.title("📊 Coupang 经营分析看板")
st.markdown("### 🚀 核心功能：多店铺数据合并 + 智能广告匹配 + 财务看板")

# --- 列号配置 ---
IDX_M_CODE   = 0    # Master A列
IDX_M_SKU    = 3    # Master D列
IDX_M_PROFIT = 10   # Master K列

IDX_S_ID     = 0    # Sales A列
IDX_S_QTY    = 8    # Sales I列

IDX_A_CAMPAIGN = 5  # Ads F列 (兜底)
IDX_A_GROUP    = 6  # Ads G列 (首选)
IDX_A_SPEND    = 15 # Ads P列
# -----------------

# ==========================================
# 2. 侧边栏上传
# ==========================================
with st.sidebar:
    st.header("📂 数据源上传")
    st.info("💡 提示：支持 .xlsx, .xlsm, .csv")
    
    file_master = st.file_uploader("1. 基础信息表 (Master - 单文件)", type=['csv', 'xlsx', 'xlsm'])
    files_sales = st.file_uploader("2. 销售表 (Sales - 多文件)", type=['csv', 'xlsx', 'xlsm'], accept_multiple_files=True)
    files_ads = st.file_uploader("3. 广告表 (Ads - 多文件)", type=['csv', 'xlsx', 'xlsm'], accept_multiple_files=True)

# ==========================================
# 3. 工具函数
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
    
    col_btn, _ = st.columns([1, 3])
    with col_btn:
        start_calc = st.button("🚀 生成看板 & 报表", type="primary", use_container_width=True)

    if start_calc:
        try:
            with st.spinner("正在清洗数据、匹配广告、核算利润..."):
                
                # --- Step 1: 读取基础表 ---
                df_master = read_file_strict(file_master)
                col_code_name = df_master.columns[IDX_M_CODE]

                df_master['_MATCH_SKU'] = clean_for_match(df_master.iloc[:, IDX_M_SKU])
                df_master['_MATCH_CODE'] = clean_for_match(df_master.iloc[:, IDX_M_CODE])
                df_master['_VAL_PROFIT'] = clean_num(df_master.iloc[:, IDX_M_PROFIT])

                # --- Step 2: 合并销售表 ---
                sales_list = [read_file_strict(f) for f in files_sales]
                df_sales_all = pd.concat(sales_list, ignore_index=True)
                
                df_sales_all['_MATCH_SKU'] = clean_for_match(df_sales_all.iloc[:, IDX_S_ID])
                df_sales_all['销量'] = clean_num(df_sales_all.iloc[:, IDX_S_QTY])
                
                sales_agg = df_sales_all.groupby('_MATCH_SKU')['销量'].sum().reset_index()
                sales_agg.rename(columns={'销量': 'O列_合并销量'}, inplace=True)

                # --- Step 3: 合并广告表 (双重提取) ---
                ads_list = [read_file_strict(f) for f in files_ads]
                df_ads_all = pd.concat(ads_list, ignore_index=True)

                df_ads_all['含税广告费'] = clean_num(df_ads_all.iloc[:, IDX_A_SPEND]) * 1.1
                df_ads_all['Code_Group'] = df_ads_all.iloc[:, IDX_A_GROUP].apply(extract_code_from_text)
                df_ads_all['Code_Campaign'] = df_ads_all.iloc[:, IDX_A_CAMPAIGN].apply(extract_code_from_text)
                df_ads_all['_MATCH_CODE'] = df_ads_all['Code_Group'].fillna(df_ads_all['Code_Campaign'])

                valid_ads = df_ads_all.dropna(subset=['_MATCH_CODE'])
                ads_agg = valid_ads.groupby('_MATCH_CODE')['含税广告费'].sum().reset_index()
                ads_agg.rename(columns={'含税广告费': 'R列_产品总广告费'}, inplace=True)

                # --- Step 4: 关联计算 ---
                df_final = pd.merge(df_master, sales_agg, on='_MATCH_SKU', how='left', sort=False)
                df_final['O列_合并销量'] = df_final['O列_合并销量'].fillna(0).astype(int)
                df_final['P列_SKU总毛利'] = df_final['O列_合并销量'] * df_final['_VAL_PROFIT']
                df_final['Q列_产品总利润'] = df_final.groupby('_MATCH_CODE', sort=False)['P列_SKU总毛利'].transform('sum')
                
                df_final = pd.merge(df_final, ads_agg, on='_MATCH_CODE', how='left', sort=False)
                df_final['R列_产品总广告费'] = df_final['R列_产品总广告费'].fillna(0)
                df_final['S列_最终净利润'] = df_final['Q列_产品总利润'] - df_final['R列_产品总广告费']

                # --- Step 5: 生成 Sheet2 数据 ---
                # 在这里计算比值
                df_sheet2 = df_final[[col_code_name, 'Q列_产品总利润', 'R列_产品总广告费', 'S列_最终净利润']].copy()
                df_sheet2 = df_sheet2.drop_duplicates(subset=[col_code_name], keep='first')
                
                # 新增计算：广告占比 = 广告费 / 总利润
                # 注意处理分母为0的情况
                df_sheet2['广告/毛利比'] = df_sheet2.apply(
                    lambda x: x['R列_产品总广告费'] / x['Q列_产品总利润'] if x['Q列_产品总利润'] != 0 else 0, 
                    axis=1
                )
                
                # --- Step 6: 清理辅助列 (str修复版) ---
                cols_to_drop = [c for c in df_final.columns if str(c).startswith('_') or str(c).startswith('Code_')]
                df_final.drop(columns=cols_to_drop, inplace=True)

                # ==========================================
                # 🔥 看板展示区 (Dashboard)
                # ==========================================
                
                # 1. 顶部 KPI 指标卡
                total_profit = df_sheet2['Q列_产品总利润'].sum()
                total_ads = df_sheet2['R列_产品总广告费'].sum()
                net_profit = df_sheet2['S列_最终净利润'].sum()
                
                st.subheader("📈 经营概览")
                kpi1, kpi2, kpi3, kpi4 = st.columns(4)
                kpi1.metric("💰 最终净利润", f"{net_profit:,.0f}", delta_color="normal")
                kpi2.metric("📦 产品总毛利", f"{total_profit:,.0f}")
                kpi3.metric("📢 总广告费", f"{total_ads:,.0f}", delta_color="inverse")
                
                if total_profit > 0:
                    overall_ads_ratio = (total_ads / total_profit)
                    kpi4.metric("📉 整体广告/毛利比", f"{overall_ads_ratio:.1%}")
                else:
                    kpi4.metric("📉 整体广告/毛利比", "N/A")

                st.divider()

                # 2. 标签页展示表格
                tab1, tab2 = st.tabs(["📝 Sheet1: 利润明细表 (查账用)", "📊 Sheet2: 业务报表 (含占比)"])
                
                with tab1:
                    st.caption("展示所有 SKU 的详细利润情况。")
                    st.dataframe(
                        df_final.style.format(precision=0)
                        .background_gradient(subset=['S列_最终净利润'], cmap='RdYlGn', vmin=-10000, vmax=10000),
                        use_container_width=True,
                        height=500
                    )
                
                with tab2:
                    st.caption("展示按产品归集的结果。新增【广告/毛利比】列。")
                    # 设置格式：金额列0位小数，比值列百分比
                    format_dict = {
                        'Q列_产品总利润': '{:,.0f}',
                        'R列_产品总广告费': '{:,.0f}', 
                        'S列_最终净利润': '{:,.0f}',
                        '广告/毛利比': '{:.1%}'
                    }
                    st.dataframe(
                        df_sheet2.style.format(format_dict)
                        .background_gradient(subset=['S列_最终净利润'], cmap='RdYlGn', vmin=-10000, vmax=10000)
                        # 广告比大于100% (即1.0) 标红，说明亏本推广
                        .text_gradient(subset=['广告/毛利比'], cmap='coolwarm', vmin=0, vmax=1.5),
                        use_container_width=True,
                        height=500
                    )

                # ==========================================
                # 📥 下载逻辑
                # ==========================================
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    # Sheet 1
                    df_final.to_excel(writer, index=False, sheet_name='利润分析')
                    
                    # Sheet 2
                    df_sheet2.to_excel(writer, index=False, sheet_name='业务报表')
                    
                    # Excel 格式化
                    wb = writer.book
                    ws2 = writer.sheets['业务报表']
                    
                    fmt_header = wb.add_format({'bold': True, 'bg_color': '#4472C4', 'font_color': 'white', 'border': 1, 'align': 'center'})
                    fmt_pct = wb.add_format({'num_format': '0.0%', 'align': 'center'})
                    fmt_money = wb.add_format({'num_format': '#,##0', 'align': 'center'})
                    
                    # 写表头
                    for col_num, value in enumerate(df_sheet2.columns.values):
                        ws2.write(0, col_num, value, fmt_header)
                    
                    # 设置列宽和格式
                    ws2.set_column(0, 0, 20) # A列 产品
                    ws2.set_column(1, 3, 15, fmt_money) # B,C,D列 金额
                    ws2.set_column(4, 4, 15, fmt_pct)   # E列 占比

                st.divider()
                st.download_button(
                    label="📥 点击下载 Excel 完整报表",
                    data=output.getvalue(),
                    file_name="Coupang_Pro_Report.xlsx",
                    mime="application/vnd.ms-excel",
                    type="primary",
                    use_container_width=True
                )

        except Exception as e:
            st.error(f"❌ 运行出错: {e}")
else:
    st.info("👈 请在左侧上传文件以开始...")

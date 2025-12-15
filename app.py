import streamlit as st
import pandas as pd
import io
import re

st.set_page_config(layout="wide", page_title="Coupang 利润核算 (列号锁定版)")
st.title("🔘 步骤五：多店铺利润核算 (列号锁定版)")
st.markdown("### 操作流程：上传文件 -> 确认就绪 -> **点击按钮** -> 生成报表")
st.caption("💡 已启用列号锁定：自动忽略表头语言（中/韩），仅依据列的位置读取数据。")

# ==========================================
# 0. 【核心配置区】在这里统一管理列号
# 说明：A列=0, B列=1, C列=2, ... L列=11, P列=15
# ==========================================

# 1. 基础信息表 (Master)
IDX_M_CODE   = 0    # 产品编号 (通常在 A列)
IDX_M_SKU    = 3    # SKU/关联ID (通常在 D列)
IDX_M_PROFIT = 10   # 单件毛利 (通常在 K列)

# 2. 销售表 (Sales)
IDX_S_ID     = 0    # 注册商品ID (通常在 A列)
IDX_S_QTY    = 8    # 销量 (通常在 I列)

# 3. 广告表 (Ads)
IDX_A_NAME   = 5    # 广告活动名称(用于提取Cxx编号) (通常在 F列)
IDX_A_SPEND  = 15   # 广告花费 (通常在 P列)

# ==========================================
# 1. 上传区域
# ==========================================
with st.sidebar:
    st.header("1. 文件上传区")
    file_master = st.file_uploader("基础信息表 (Master - 1个)", type=['csv', 'xlsx'])
    files_sales = st.file_uploader("销售表 (Sales - 支持多个)", type=['csv', 'xlsx'], accept_multiple_files=True)
    files_ads = st.file_uploader("广告表 (Ads - 支持多个)", type=['csv', 'xlsx'], accept_multiple_files=True)

    st.markdown("---")
    if file_master and files_sales and files_ads:
        st.success("✅ 所有文件已上传，请去右侧点击按钮开始。")
    else:
        st.info("⏳ 等待文件上传完整...")

# ==========================================
# 2. 工具函数
# ==========================================
def clean_id(series):
    """清洗ID：转字符串，去小数，去引号，去空格"""
    return series.astype(str).str.replace(r'\.0$', '', regex=True).str.replace('"', '').str.replace('\n', '').str.strip()

def clean_num(series):
    """清洗数值：转数字，无法转换的变0"""
    return pd.to_numeric(series, errors='coerce').fillna(0)

def extract_product_code(text):
    """从广告名称中提取 C01 这种编号"""
    if pd.isna(text): return None
    match = re.search(r'([Cc]\d+)', str(text))
    if match: return match.group(1).upper()
    return None

def read_and_combine(file_list, file_type_name=""):
    """读取并合并多个文件"""
    if not file_list: return pd.DataFrame()
    all_dfs = []
    for file in file_list:
        try:
            file.seek(0)
            if file.name.endswith('.csv'):
                try: df = pd.read_csv(file)
                except: file.seek(0); df = pd.read_csv(file, encoding='gbk')
            else:
                df = pd.read_excel(file)
            all_dfs.append(df)
        except Exception as e: st.error(f"❌ {file.name} 读取失败: {e}")
    
    if all_dfs:
        combined = pd.concat(all_dfs, ignore_index=True)
        # 简单去重
        rows_before = len(combined)
        combined.drop_duplicates(inplace=True)
        rows_after = len(combined)
        removed = rows_before - rows_after
        if removed > 0: st.warning(f"⚠️ 【{file_type_name}】剔除了 {removed} 条完全重复的数据")
        return combined
    return pd.DataFrame()

# ==========================================
# 3. 主界面逻辑
# ==========================================

if file_master and files_sales and files_ads:
    
    st.divider()
    col1, col2 = st.columns([3, 1])
    with col1:
        st.subheader("📂 文件状态确认")
        st.write(f"• 基础表：1 个")
        st.write(f"• 销售表：{len(files_sales)} 个 (待合并)")
        st.write(f"• 广告表：{len(files_ads)} 个 (待合并)")
    
    with col2:
        st.write("##")
        start_btn = st.button("🚀 点击开始计算", type="primary", use_container_width=True)

    if start_btn:
        st.divider()
        with st.status("🔄 正在全速计算中...", expanded=True):
            try:
                # -------------------------------------------------------
                # A. Master (基础表处理)
                # -------------------------------------------------------
                st.write("1. 正在读取基础表并锁定列位置...")
                if file_master.name.endswith('.csv'): df_master = pd.read_csv(file_master)
                else: df_master = pd.read_excel(file_master)
                
                # 【优化点】使用配置区的常量读取列
                df_master['__ORDER__'] = range(len(df_master)) # 保留原始顺序
                
                # 锁定关键列的数据
                raw_col_code = df_master.iloc[:, IDX_M_CODE]
                raw_col_sku  = df_master.iloc[:, IDX_M_SKU]
                raw_col_profit = df_master.iloc[:, IDX_M_PROFIT]

                df_master['关联ID'] = clean_id(raw_col_sku)
                df_master['单件毛利'] = clean_num(raw_col_profit)
                df_master['产品编号_清洗'] = clean_id(raw_col_code).str.upper()

                # -------------------------------------------------------
                # B. Sales (销售表处理)
                # -------------------------------------------------------
                st.write("2. 正在合并销售数据...")
                df_sales_all = read_and_combine(files_sales, "销售表")
                
                # 锁定关键列
                raw_sale_id = df_sales_all.iloc[:, IDX_S_ID]
                raw_sale_qty = df_sales_all.iloc[:, IDX_S_QTY]

                df_sales_all['关联ID'] = clean_id(raw_sale_id)
                df_sales_all['销量'] = clean_num(raw_sale_qty)
                
                sales_agg = df_sales_all.groupby('关联ID')['销量'].sum().reset_index()
                sales_agg.rename(columns={'销量': 'O列_合并销量'}, inplace=True)

                # -------------------------------------------------------
                # C. Ads (广告表处理)
                # -------------------------------------------------------
                st.write("3. 正在匹配广告花费...")
                df_ads_all = read_and_combine(files_ads, "广告表")
                
                # 锁定关键列
                raw_ad_name = df_ads_all.iloc[:, IDX_A_NAME]
                raw_ad_spend = df_ads_all.iloc[:, IDX_A_SPEND]

                df_ads_all['提取编号'] = raw_ad_name.apply(extract_product_code)
                df_ads_all['含税广告费'] = clean_num(raw_ad_spend) * 1.1 # 加上10%税点
                
                ads_agg = df_ads_all.groupby('提取编号')['含税广告费'].sum().reset_index()
                ads_agg.rename(columns={'提取编号': '产品编号_清洗', '含税广告费': 'R列_产品总广告费'}, inplace=True)

                # -------------------------------------------------------
                # D. Merge (合并计算)
                # -------------------------------------------------------
                st.write("4. 正在生成最终报表...")
                df_final = pd.merge(df_master, sales_agg, on='关联ID', how='left')
                df_final['O列_合并销量'] = df_final['O列_合并销量'].fillna(0).astype(int)
                df_final['P列_SKU总毛利'] = df_final['O列_合并销量'] * df_final['单件毛利']
                
                # 计算产品维度的总利润
                df_final['Q列_产品总利润'] = df_final.groupby('产品编号_清洗')['P列_SKU总毛利'].transform('sum')
                
                # 匹配广告费
                df_final = pd.merge(df_final, ads_agg, on='产品编号_清洗', how='left')
                df_final['R列_产品总广告费'] = df_final['R列_产品总广告费'].fillna(0)
                
                # 最终净利
                df_final['S列_最终净利润'] = df_final['Q列_产品总利润'] - df_final['R列_产品总广告费']

                # 恢复排序并清理中间列
                df_final.sort_values(by=['__ORDER__'], inplace=True)
                cols_to_drop = ['__ORDER__', '关联ID', '单件毛利', '产品编号_清洗', '提取编号']
                df_final.drop(columns=[c for c in cols_to_drop if c in df_final.columns], inplace=True, errors='ignore')

                # -------------------------------------------------------
                # E. Excel Output (保留你的样式代码)
                # -------------------------------------------------------
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    wb = writer.book
                    
                    # Sheet 1: 明细
                    df_final.to_excel(writer, index=False, sheet_name='1_超级数据源')
                    ws1 = writer.sheets['1_超级数据源']
                    (mr, mc) = df_final.shape
                    # 注意：这里 header 需要转 string 防止报错
                    cols_settings = [{'header': str(c)} for c in df_final.columns]
                    ws1.add_table(0, 0, mr, mc-1, {'columns': cols_settings, 'name': 'Data', 'style': 'TableStyleMedium9'})
                    ws1.set_column(0, mc-1, 15)

                    # Sheet 2: 老板视图
                    df_final.to_excel(writer, index=False, sheet_name='2_老板视图')
                    ws2 = writer.sheets['2_老板视图']
                    
                    merge_fmt = wb.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'fg_color': '#FFFFFF'})
                    green_fmt = wb.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#C6EFCE', 'font_color': '#006100'})
                    red_fmt = wb.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#FFC7CE', 'font_color': '#9C0006'})
                    
                    ws2.set_column('A:A', 15)
                    cols_list = df_final.columns.tolist()
                    
                    # 动态寻找列索引 (防止列位置变动导致写入错位)
                    idx_A = 0
                    idx_Q = cols_list.index('Q列_产品总利润') if 'Q列_产品总利润' in cols_list else -1
                    idx_R = cols_list.index('R列_产品总广告费') if 'R列_产品总广告费' in cols_list else -1
                    idx_S = cols_list.index('S列_最终净利润') if 'S列_最终净利润' in cols_list else -1

                    # 你的原始合并逻辑
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
                                if idx_Q >= 0: ws2.merge_range(start_row, idx_Q, i, idx_Q, q_vals[start_row-1], merge_fmt)
                                if idx_R >= 0: ws2.merge_range(start_row, idx_R, i, idx_R, r_vals[start_row-1], merge_fmt)
                                if idx_S >= 0: ws2.merge_range(start_row, idx_S, i, idx_S, profit, s_fmt)
                            else:
                                ws2.write(start_row, idx_A, codes[start_row-1], merge_fmt)
                                if idx_Q >= 0: ws2.write(start_row, idx_Q, q_vals[start_row-1], merge_fmt)
                                if idx_R >= 0: ws2.write(start_row, idx_R, r_vals[start_row-1], merge_fmt)
                                if idx_S >= 0: ws2.write(start_row, idx_S, profit, s_fmt)
                            start_row = i + 1
                    
                st.success("✅ 计算完成！")
                st.download_button("📥 下载结果报表", output.getvalue(), "Coupang_Final_Result.xlsx", "application/vnd.ms-excel", type='primary')

            except Exception as e:
                st.error(f"发生错误: {e}")
                st.error("💡 建议检查：上传的表格列顺序是否发生了变化？请核对代码最上方的配置区列号。")

else:
    st.info("👈 请在左侧上传文件：1个基础表 + 多个销售/广告表。上传完成后，此处会出现开始按钮。")
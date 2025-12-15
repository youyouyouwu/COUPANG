import streamlit as st
import pandas as pd
import io
import re

# ==========================================
# 1. 页面配置
# ==========================================
st.set_page_config(layout="wide", page_title="Coupang 利润核算 (修复版)")
st.title("🔧 步骤五：全链路利润核算 (强力修复版)")
st.markdown("### 修复内容：解决 CSV 换行符/引号问题，增加未匹配数据检视功能")

# ==========================================
# 2. 上传区域
# ==========================================
with st.sidebar:
    st.header("请上传文件")
    file_master = st.file_uploader("1. 基础信息表 (Master)", type=['csv', 'xlsx'])
    file_sales = st.file_uploader("3. 销售表 (Sales)", type=['csv', 'xlsx'])
    file_ads = st.file_uploader("4. 广告表 (Ads)", type=['csv', 'xlsx'])

# ==========================================
# 3. 工具函数
# ==========================================
def clean_id(series):
    """基础ID清洗：转文本 -> 去.0 -> 去空格 -> 去引号 -> 去换行"""
    return series.astype(str).str.replace(r'\.0$', '', regex=True).str.replace('"', '').str.replace('\n', '').str.strip()

def clean_num(series):
    """数字清洗"""
    return pd.to_numeric(series, errors='coerce').fillna(0)

def extract_product_code(text):
    """
    从广告活动名中提取 C+数字 的编号
    优化：忽略大小写，支持提取 C001, c001
    """
    if pd.isna(text):
        return None
    # 正则：寻找 C 或 c 开头，后面紧跟数字的组合
    match = re.search(r'([Cc]\d+)', str(text))
    if match:
        # 统一转为大写 (C001) 以便匹配
        return match.group(1).upper()
    return None

# ==========================================
# 4. 执行逻辑
# ==========================================
if file_master and file_sales and file_ads:
    st.info("🔄 正在清洗数据并计算，请稍候...")
    
    try:
        # ------------------------------------------------
        # A. 基础表 (Master) 读取与强力清洗
        # ------------------------------------------------
        if file_master.name.endswith('.csv'):
            df_master = pd.read_csv(file_master)
        else:
            df_master = pd.read_excel(file_master)

        # 锁定列 (按索引，防止列名乱码)
        # A列(0): 产品编号, D列(3): SKU ID, K列(10): 毛利润
        col_code = df_master.columns[0]
        col_sku = df_master.columns[3]
        col_profit = df_master.columns[10]

        # 【核心修复】针对 CSV 里的 "C0001\n" 进行清洗
        df_master['关联ID'] = clean_id(df_master[col_sku])
        df_master['单件毛利'] = clean_num(df_master[col_profit])
        df_master['产品编号_清洗'] = clean_id(df_master[col_code]).str.upper() # 统一大写

        # ------------------------------------------------
        # B. 销售表 (Sales)
        # ------------------------------------------------
        if file_sales.name.endswith('.csv'):
            df_sales = pd.read_csv(file_sales)
        else:
            df_sales = pd.read_excel(file_sales)

        col_sales_id = df_sales.columns[0]
        col_sales_qty = df_sales.columns[8]

        df_sales['关联ID'] = clean_id(df_sales[col_sales_id])
        df_sales['销量'] = clean_num(df_sales[col_sales_qty])
        
        # 汇总销量
        sales_agg = df_sales.groupby('关联ID')['销量'].sum().reset_index()
        sales_agg.rename(columns={'销量': 'O列_合并销量'}, inplace=True)

        # ------------------------------------------------
        # C. 广告表 (Ads)
        # ------------------------------------------------
        if file_ads.name.endswith('.csv'):
            df_ads = pd.read_csv(file_ads)
        else:
            df_ads = pd.read_excel(file_ads)

        col_campaign = df_ads.columns[5] # F列
        col_ad_spend = df_ads.columns[15] # P列

        # 提取并统一大写
        df_ads['提取编号'] = df_ads[col_campaign].apply(extract_product_code)
        df_ads['含税广告费'] = clean_num(df_ads[col_ad_spend]) * 1.1 # 补税点
        
        # 汇总广告费
        ads_valid = df_ads.dropna(subset=['提取编号'])
        ads_agg = ads_valid.groupby('提取编号')['含税广告费'].sum().reset_index()
        ads_agg.rename(columns={'提取编号': '产品编号_清洗', '含税广告费': 'R列_产品总广告费'}, inplace=True)

        # ------------------------------------------------
        # 🔍 调试面板 (帮你找错误)
        # ------------------------------------------------
        with st.expander("🕵️‍♂️ 数据匹配侦探 (点我查看匹配情况)"):
            st.write("基础表里的产品编号示例:", df_master['产品编号_清洗'].unique()[:5])
            st.write("广告表提取出的编号示例:", ads_agg['产品编号_清洗'].unique()[:5])
            
            # 检查有多少广告费没匹配上
            master_codes = set(df_master['产品编号_清洗'])
            unmatched_ads = ads_agg[~ads_agg['产品编号_清洗'].isin(master_codes)]
            if not unmatched_ads.empty:
                st.warning(f"⚠️ 警告：有 {len(unmatched_ads)} 个广告活动没找到对应的产品！(可能是编号写错了)")
                st.dataframe(unmatched_ads)
            else:
                st.success("✅ 完美！所有提取到的广告费都匹配到了产品。")

        # ------------------------------------------------
        # D. 全链路合并
        # ------------------------------------------------
        # 1. Master + Sales
        df_final = pd.merge(df_master, sales_agg, on='关联ID', how='left')
        df_final['O列_合并销量'] = df_final['O列_合并销量'].fillna(0).astype(int)
        
        # 2. 算 SKU 毛利 (P列)
        df_final['P列_SKU总毛利'] = df_final['O列_合并销量'] * df_final['单件毛利']
        
        # 3. 算 产品总毛利 (Q列)
        df_final['Q列_产品总利润'] = df_final.groupby('产品编号_清洗')['P列_SKU总毛利'].transform('sum')
        
        # 4. 匹配 广告费 (R列)
        df_final = pd.merge(df_final, ads_agg, on='产品编号_清洗', how='left')
        df_final['R列_产品总广告费'] = df_final['R列_产品总广告费'].fillna(0)
        
        # 5. 算 最终盈亏 (S列)
        df_final['S列_最终净利润'] = df_final['Q列_产品总利润'] - df_final['R列_产品总广告费']

        # ------------------------------------------------
        # E. 排序与输出
        # ------------------------------------------------
        df_final.sort_values(by=['产品编号_清洗'], inplace=True)
        
        # 清理列
        cols_to_drop = ['关联ID', '单件毛利', '产品编号_清洗', '提取编号']
        df_final.drop(columns=[c for c in cols_to_drop if c in df_final.columns], inplace=True)

        st.success("✅ 计算成功！")

        # Excel 导出
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            sheet_name = '盈亏分析'
            df_final.to_excel(writer, index=False, sheet_name=sheet_name)
            
            workbook = writer.book
            worksheet = writer.sheets[sheet_name]
            
            # 样式
            merge_fmt = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'fg_color': '#FFFFFF'})
            green_fmt = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#C6EFCE', 'font_color': '#006100'})
            red_fmt = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#FFC7CE', 'font_color': '#9C0006'})
            
            # 列宽
            worksheet.set_column('A:A', 15)
            
            # 合并逻辑
            cols = df_final.columns.tolist()
            idx_A = 0
            # 动态寻找列索引 (防止列变动)
            idx_Q = cols.index('Q列_产品总利润') if 'Q列_产品总利润' in cols else len(cols)-3
            idx_R = cols.index('R列_产品总广告费') if 'R列_产品总广告费' in cols else len(cols)-2
            idx_S = cols.index('S列_最终净利润') if 'S列_最终净利润' in cols else len(cols)-1
            
            start_row = 1
            # 取出数据用于循环对比
            codes = df_final.iloc[:, 0].astype(str).tolist() # 确保是字符串对比
            q_vals = df_final['Q列_产品总利润'].tolist()
            r_vals = df_final['R列_产品总广告费'].tolist()
            s_vals = df_final['S列_最终净利润'].tolist()
            
            for i in range(1, len(codes) + 1):
                if i == len(codes) or codes[i] != codes[i-1]:
                    profit = s_vals[start_row-1]
                    s_fmt = green_fmt if profit >= 0 else red_fmt
                    
                    if i - start_row > 0:
                        worksheet.merge_range(start_row, idx_A, i, idx_A, codes[start_row-1], merge_fmt)
                        worksheet.merge_range(start_row, idx_Q, i, idx_Q, q_vals[start_row-1], merge_fmt)
                        worksheet.merge_range(start_row, idx_R, i, idx_R, r_vals[start_row-1], merge_fmt)
                        worksheet.merge_range(start_row, idx_S, i, idx_S, profit, s_fmt)
                    else:
                        worksheet.write(start_row, idx_A, codes[start_row-1], merge_fmt)
                        worksheet.write(start_row, idx_Q, q_vals[start_row-1], merge_fmt)
                        worksheet.write(start_row, idx_R, r_vals[start_row-1], merge_fmt)
                        worksheet.write(start_row, idx_S, profit, s_fmt)
                    start_row = i + 1

        st.download_button(
            label="📥 下载最终报表 (修复版)",
            data=output.getvalue(),
            file_name="Coupang_Profit_Fixed.xlsx",
            mime="application/vnd.ms-excel"
        )

    except Exception as e:
        st.error(f"❌ 程序发生错误: {e}")
        st.warning("提示：如果报错 'KeyError'，通常是因为上传的文件列名不对，请确保上传的是原始文件。")

else:
    st.info("👈 请上传 3 个文件 (Master, Sales, Ads)")
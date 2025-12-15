import streamlit as st
import pandas as pd
import io
import re

st.set_page_config(layout="wide", page_title="Coupang 利润核算 (列号锁定版)")
st.title("🔘 步骤五：多店铺利润核算 (列号锁定版)")
st.markdown("### 操作流程：上传文件 -> 确认就绪 -> **点击按钮** -> 生成报表")
st.caption("✅ 已启用列号锁定功能：自动无视表头语言（中/韩/英），只识别列的位置。")

# ==========================================
# 0. 【核心配置区】在这里修改列号
# 如果Coupang表格格式变了，只需要改这里的数字 (A列=0, B列=1, C列=2...)
# ==========================================
# 1. 基础信息表 (Master)
IDX_M_CODE   = 0    # 产品编号所在列 (A列)
IDX_M_SKU    = 3    # SKU/关联ID所在列 (D列)
IDX_M_PROFIT = 10   # 单件毛利所在列 (K列)

# 2. 销售表 (Sales)
IDX_S_ID     = 0    # 关联ID/注册商品ID所在列 (A列)
IDX_S_QTY    = 8    # 销量所在列 (I列)

# 3. 广告表 (Ads)
IDX_A_NAME   = 5    # 广告活动名称(用于提取编号)所在列 (F列)
IDX_A_SPEND  = 15   # 广告花费所在列 (P列)

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
        # 去重逻辑
        rows_before = len(combined)
        combined.drop_duplicates(inplace=True)
        rows_after = len(combined)
        removed = rows_before - rows_after
        if removed > 0: st.warning(f"⚠️ 【{file_type_name}】剔除了 {removed} 条完全重复的数据")
        return combined
    return pd.DataFrame()

# ==========================================
# 3. 主界面逻辑 (核心修改部分)
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
        with st.status("🔄 正在全速计算中 (列号锁定模式)...", expanded=True):
            try:
                # -------------------------------------------------------
                # A. 处理 Master (基础表)
                # -------------------------------------------------------
                st.write("1. 正在读取基础表并锁定关键列...")
                if file_master.name.endswith('.csv'): df_master = pd.read_csv(file_master)
                else: df_master = pd.read_excel(file_master)
                
                # 【关键逻辑】不再读取列名，而是直接通过列号(Index)重命名
                # 获取原列名
                cols = df_master.columns
                # 建立映射关系： 原名 -> 标准名
                rename_dict = {
                    cols[IDX_M_CODE]: 'RAW_CODE',     # 对应第0列
                    cols[IDX_M_SKU]: 'RAW_SKU',       # 对应第3列
                    cols[IDX_M_PROFIT]: 'RAW_PROFIT'  # 对应第10列
                }
                df_master.rename(columns=rename_dict, inplace=True)
                
                # 保留原始顺序用于排序
                df_master['__ORDER__'] = range(len(df_master))

                # 使用标准名进行清洗
                df_master['关联ID'] = clean_id(df_master['RAW_SKU'])
                df_master['单件毛利'] = clean_num(df_master['RAW_PROFIT'])
                df_master['产品编号_清洗'] = clean_id(df_master['RAW_CODE']).str.upper()

                # -------------------------------------------------------
                # B. 处理 Sales (销售表)
                # -------------------------------------------------------
                st.write("2. 正在合并销售数据...")
                df_sales_all = read_and_combine(files_sales, "销售表")
                
                if not df_sales_all.empty:
                    # 【关键逻辑】列号锁定
                    cols_s = df_sales_all.columns
                    rename_dict_s = {
                        cols_s[IDX_S_ID]: 'RAW_SALE_ID',  # 对应第0列
                        cols_s[IDX_S_QTY]: 'RAW_SALE_QTY' # 对应第8列
                    }
                    df_sales_all.rename(columns=rename_dict_s, inplace=True)

                    df_sales_all['关联ID'] = clean_id(df_sales_all['RAW_SALE_ID'])
                    df_sales_all['销量'] = clean_num(df_sales_all['RAW_SALE_QTY'])
                    
                    sales_agg = df_sales_all.groupby('关联ID')['销量'].sum().reset_index()
                    sales_agg.rename(columns={'销量': 'O列_合并销量'}, inplace=True)
                else:
                    sales_agg = pd.DataFrame(columns=['关联ID', 'O列_合并销量'])

                # -------------------------------------------------------
                # C. 处理 Ads (广告表)
                # -------------------------------------------------------
                st.write("3. 正在匹配广告花费...")
                df_ads_all = read_and_combine(files_ads, "广告表")

                if not df_ads_all.empty:
                    # 【关键逻辑】列号锁定
                    cols_a = df_ads_all.columns
                    rename_dict_a = {
                        cols_a[IDX_A_NAME]: 'RAW_AD_NAME',   # 对应第5列
                        cols_a[IDX_A_SPEND]: 'RAW_AD_SPEND'  # 对应第15列
                    }
                    df_ads_all.rename(columns=rename_dict_a, inplace=True)

                    df_ads_all['提取编号'] = df_ads_all['RAW_AD_NAME'].apply(extract_product_code)
                    # 加上10%税点
                    df_ads_all['含税广告费'] = clean_num(df_ads_all['RAW_AD_SPEND']) * 1.1
                    
                    ads_agg = df_ads_all.groupby('提取编号')['含税广告费'].sum().reset_index()
                    ads_agg.rename(columns={'提取编号': '产品编号_清洗', '含税广告费': 'R列_产品总广告费'}, inplace=True)
                else:
                    ads_agg = pd.DataFrame(columns=['产品编号_清洗', 'R列_产品总广告费'])

                # -------------------------------------------------------
                # D. Merge (合并计算)
                # -------------------------------------------------------
                st.write("4. 正在生成最终报表...")
                
                # 合并销量
                df_final = pd.merge(df_master, sales_agg, on='关联ID', how='left')
                df_final['O列_合并销量'] = df_final['O列_合并销量'].fillna(0).astype(int)
                
                # 计算 SKU 总毛利
                df_final['P列_SKU总毛利'] = df_final['O列_合并销量'] * df_final['单件毛利']
                
                # 计算 产品总利润 (同产品ID的所有SKU利润之和)
                df_final['Q列_产品总利润'] = df_final.groupby('产品编号_清洗')['P列_SKU总毛利'].transform('sum')
                
                # 合并广告费 (按产品编号)
                df_final = pd.merge(df_final, ads_agg, on='产品编号_清洗', how='left')
                df_final['R列_产品总广告费'] = df_final['R列_产品总广告费'].fillna(0)
                
                # 计算最终净利润
                df_final['S列_最终净利润'] = df_final['Q列_产品总利润'] - df_final['R列_产品总广告费']

                # 整理格式
                df_final.sort_values(by=['__ORDER__'], inplace=True)
                # 移除临时列，但可以保留原始列以供参考，这里选择移除中间计算用的列
                cols_to_drop = ['__ORDER__', '关联ID', '单件毛利', '产品编号_清洗', 'RAW_CODE', 'RAW_SKU', 'RAW_PROFIT']
                df_final.drop(columns=[c for c in cols_to_drop if c in df_final.columns], inplace=True, errors='ignore')

                # -------------------------------------------------------
                # E. Excel Output (样式美化)
                # -------------------------------------------------------
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    wb = writer.book
                    
                    # Sheet 1: 原始明细
                    df_final.to_excel(writer, index=False, sheet_name='1_超级数据源')
                    ws1 = writer.sheets['1_超级数据源']
                    (mr, mc) = df_final.shape
                    cols_settings = [{'header': str(c)} for c in df_final.columns]
                    ws1.add_table(0, 0, mr, mc-1, {'columns': cols_settings, 'name': 'Data', 'style': 'TableStyleMedium9'})
                    ws1.set_column(0, mc-1, 15)

                    # Sheet 2: 老板视图 (带颜色和合并单元格)
                    df_final.to_excel(writer, index=False, sheet_name='2_老板视图')
                    ws2 = writer.sheets['2_老板视图']
                    
                    merge_fmt = wb.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'fg_color': '#FFFFFF'})
                    green_fmt = wb.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#C6EFCE', 'font_color': '#006100'})
                    red_fmt = wb.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#FFC7CE', 'font_color': '#9C0006'})
                    
                    ws2.set_column('A:A', 15)
                    cols_list = df_final.columns.tolist()
                    
                    # 动态寻找列的位置
                    idx_A = 0
                    idx_Q = cols_list.index('Q列_产品总利润') if 'Q列_产品总利润' in cols_list else -1
                    idx_R = cols_list.index('R列_产品总广告费') if 'R列_产品总广告费' in cols_list else -1
                    idx_S = cols_list.index('S列_最终净利润') if 'S列_最终净利润' in cols_list else -1

                    start_row = 1
                    # 假设第一列是产品ID用于合并
                    codes = df_final.iloc[:, 0].astype(str).tolist()
                    q_vals = df_final['Q列_产品总利润'].tolist()
                    r_vals = df_final['R列_产品总广告费'].tolist()
                    s_vals = df_final['S列_最终净利润'].tolist()

                    for i in range(1, len(codes) + 1):
                        if i == len(codes) or codes[i] != codes[i-1]:
                            profit = s_vals[start_row-1]
                            s_fmt = green_fmt if profit >= 0 else red_fmt
                            cnt = i - start_row
                            
                            # 执行合并写入
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
                    
                st.success("✅ 计算完成！表头语言差异已自动解决。")
                st.download_button("📥 下载结果报表", output.getvalue(), "Coupang_Final_Result.xlsx", "application/vnd.ms-excel", type='primary')

            except Exception as e:
                st.error(f"发生错误: {e}")
                st.error("💡 排查建议：请检查上传的表格列顺序是否发生了变化？请核对代码顶部的配置区列号。")

else:
    st.info("👈 请在左侧上传文件。上传完成后，此处会出现开始按钮。")
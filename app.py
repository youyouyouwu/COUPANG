import streamlit as st
import pandas as pd
import io
import re

# ==========================================
# 1. 页面配置
# ==========================================
st.set_page_config(layout="wide", page_title="Coupang 利润核算 (精修版)")
st.title("🎨 步骤五：利润核算 (最终精修版)")
st.markdown("""
### 🛡️ 优化细节：
1.  **0值处理**：净利润为 0 时不显示颜色，保持表格整洁。
2.  **自动列宽**：列宽自动根据内容调整，避免过宽或过窄，一眼看全数据。
3.  **样式保留**：首行冻结 + 微软雅黑加粗 + 深灰斑马纹。
""", unsafe_allow_html=True)

# --- 列号配置 ---
IDX_M_CODE   = 0    # A列
IDX_M_SKU    = 3    # D列
IDX_M_PROFIT = 10   # K列
IDX_S_ID     = 0    # A列
IDX_S_QTY    = 8    # I列
IDX_A_NAME   = 5    # F列
IDX_A_SPEND  = 15   # P列
# -----------------

# ==========================================
# 2. 上传区域
# ==========================================
with st.sidebar:
    st.header("📂 文件上传")
    file_master = st.file_uploader("1. 基础信息表 (Master)", type=['csv', 'xlsx'])
    file_sales = st.file_uploader("3. 销售表 (Sales)", type=['csv', 'xlsx'])
    file_ads = st.file_uploader("4. 广告表 (Ads)", type=['csv', 'xlsx'])

# ==========================================
# 3. 清洗工具
# ==========================================
def clean_for_match(series):
    return series.astype(str).str.replace(r'\.0$', '', regex=True).str.replace('"', '').str.strip().str.upper()

def clean_num(series):
    return pd.to_numeric(series, errors='coerce').fillna(0)

def extract_code_from_ad(text):
    if pd.isna(text): return None
    match = re.search(r'([Cc]\d+)', str(text))
    if match: return match.group(1).upper()
    return None

def read_file_strict(file):
    try:
        if file.name.endswith('.csv'):
            return pd.read_csv(file, dtype=str)
        else:
            return pd.read_excel(file, dtype=str)
    except:
        file.seek(0)
        return pd.read_csv(file, dtype=str, encoding='gbk')

# 计算字符宽度的辅助函数 (粗略估算)
def get_col_width(series):
    # 计算每行字符长度，中文按2个字符算可能更准，这里简单用len
    max_len = series.astype(str).map(len).max()
    return max_len

# ==========================================
# 4. 主逻辑
# ==========================================
if file_master and file_sales and file_ads:
    st.divider()
    if st.button("🚀 开始计算 (自动调整列宽)", type="primary", use_container_width=True):
        try:
            with st.status("🔄 正在计算... (正在适配列宽...)", expanded=True):
                # --------------------------------------------
                # Step A-D: 读取与计算 (逻辑不变)
                # --------------------------------------------
                df_master = read_file_strict(file_master)
                df_master['_MATCH_SKU'] = clean_for_match(df_master.iloc[:, IDX_M_SKU])
                df_master['_MATCH_CODE'] = clean_for_match(df_master.iloc[:, IDX_M_CODE])
                df_master['_VAL_PROFIT'] = clean_num(df_master.iloc[:, IDX_M_PROFIT])

                df_sales = read_file_strict(file_sales)
                df_sales['_MATCH_SKU'] = clean_for_match(df_sales.iloc[:, IDX_S_ID])
                df_sales['销量'] = clean_num(df_sales.iloc[:, IDX_S_QTY])
                sales_agg = df_sales.groupby('_MATCH_SKU')['销量'].sum().reset_index()
                sales_agg.rename(columns={'销量': 'O列_合并销量'}, inplace=True)

                df_ads = read_file_strict(file_ads)
                df_ads['提取编号'] = df_ads.iloc[:, IDX_A_NAME].apply(extract_code_from_ad)
                df_ads['含税广告费'] = clean_num(df_ads.iloc[:, IDX_A_SPEND]) * 1.1
                valid_ads = df_ads.dropna(subset=['提取编号'])
                ads_agg = valid_ads.groupby('提取编号')['含税广告费'].sum().reset_index()
                ads_agg.rename(columns={'提取编号': '_MATCH_CODE', '含税广告费': 'R列_产品总广告费'}, inplace=True)

                # 合并
                df_final = pd.merge(df_master, sales_agg, on='_MATCH_SKU', how='left', sort=False)
                df_final['O列_合并销量'] = df_final['O列_合并销量'].fillna(0).astype(int)
                df_final['P列_SKU总毛利'] = df_final['O列_合并销量'] * df_final['_VAL_PROFIT']
                df_final['Q列_产品总利润'] = df_final.groupby('_MATCH_CODE', sort=False)['P列_SKU总毛利'].transform('sum')
                df_final = pd.merge(df_final, ads_agg, on='_MATCH_CODE', how='left', sort=False)
                df_final['R列_产品总广告费'] = df_final['R列_产品总广告费'].fillna(0)
                df_final['S列_最终净利润'] = df_final['Q列_产品总利润'] - df_final['R列_产品总广告费']

                # 清理
                cols_to_drop = [c for c in df_final.columns if c.startswith('_')]
                df_final.drop(columns=cols_to_drop, inplace=True)

                # --------------------------------------------
                # Step E: 输出 Excel (样式精修)
                # --------------------------------------------
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df_final.to_excel(writer, index=False, sheet_name='利润分析')
                    wb = writer.book
                    ws = writer.sheets['利润分析']
                    
                    # 样式对象
                    base_font = {'font_name': 'Microsoft YaHei', 'bold': True, 'border': 1, 'align': 'center', 'valign': 'vcenter'}
                    fmt_row_grey = wb.add_format(dict(base_font, bg_color='#BFBFBF'))
                    fmt_row_white = wb.add_format(dict(base_font, bg_color='#FFFFFF'))
                    
                    # 盈亏样式 (仅背景色不同)
                    fmt_s_profit = wb.add_format(dict(base_font, bg_color='#C6EFCE')) # 绿
                    fmt_s_loss = wb.add_format(dict(base_font, bg_color='#FFC7CE'))   # 红

                    # === 自动列宽调整 ===
                    # 遍历每一列，计算最大内容长度，并设置宽度
                    for i, col in enumerate(df_final.columns):
                        # 获取该列最长内容的长度
                        max_len = get_col_width(df_final[col])
                        # 表头长度也要考虑
                        header_len = len(str(col)) * 1.5 # 中文表头稍微加权
                        
                        # 最终宽度：取内容和表头的最大值，稍微加点余量
                        final_width = max(max_len, header_len) + 2
                        
                        # 限制一下最大宽度，防止描述列太宽撑爆屏幕
                        if final_width > 50: final_width = 50
                        if final_width < 10: final_width = 10 # 最小宽度
                        
                        ws.set_column(i, i, final_width)

                    # === 冻结首行 ===
                    ws.freeze_panes(1, 0)

                    # === 智能着色 ===
                    col_code_idx = IDX_M_CODE
                    cols_list = df_final.columns.tolist()
                    col_profit_idx = cols_list.index('S列_最终净利润') if 'S列_最终净利润' in cols_list else -1

                    raw_codes = df_final.iloc[:, col_code_idx].astype(str).tolist()
                    clean_codes = [str(x).replace('.0','').replace('"','').strip().upper() for x in raw_codes]
                    
                    is_grey = False
                    for i in range(len(raw_codes)):
                        excel_row = i + 1
                        # 切换斑马纹
                        if i > 0 and clean_codes[i] != clean_codes[i-1]:
                            is_grey = not is_grey
                        
                        # 应用行样式
                        ws.set_row(excel_row, None, fmt_row_grey if is_grey else fmt_row_white)
                        
                        # 单独处理 S列
                        if col_profit_idx != -1:
                            val = df_final.iloc[i, col_profit_idx]
                            try:
                                num_val = float(val)
                            except:
                                num_val = 0
                            
                            # 逻辑修改：只有不等于0才上色
                            if num_val > 0:
                                ws.write(excel_row, col_profit_idx, val, fmt_s_profit)
                            elif num_val < 0:
                                ws.write(excel_row, col_profit_idx, val, fmt_s_loss)
                            else:
                                # 等于0，保持该行的原样 (什么都不做，或者显式写回去以防万一)
                                # 为了稳妥，用当前行的默认格式把值写回去
                                ws.write(excel_row, col_profit_idx, val, fmt_row_grey if is_grey else fmt_row_white)

            st.success("✅ 报表生成！列宽已自动适配，0值显示已优化。")
            st.download_button("📥 下载精修版报表", output.getvalue(), "Coupang_Perfect_Report.xlsx")

        except Exception as e:
            st.error(f"❌ 错误: {e}")
else:
    st.info("👈 请在左侧上传文件")
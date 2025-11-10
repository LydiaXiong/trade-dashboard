# trade_dashboard.py
# 贸易业务分析仪表盘（Streamlit 单文件）
# 功能：
# - 从 Excel 上传 / 读取数据（支持多 sheet）
# - 使用 pandas DataFrame 进行读取、筛选、分组、合并、分析
# - 计算关键指标：总销售额、前5大客户、按产品/客户/月份汇总等
# - 使用 matplotlib / seaborn 绘制折线图、柱状图、散点图、热力图
# - 导出筛选后的数据为 CSV
# 运行前：pip install streamlit pandas matplotlib seaborn openpyxl

import io
from datetime import datetime
import pandas as pd
import numpy as np
import streamlit as st
import matplotlib.pyplot as plt
import seaborn as sns

st.set_page_config(layout="wide", page_title="贸易业务分析仪表盘")

st.title("贸易业务分析仪表盘")

# ---------- 辅助函数 ----------
@st.cache_data
def read_excel_file(uploaded_file):
    # 返回 dict: sheet_name -> DataFrame
    try:
        xls = pd.read_excel(uploaded_file, sheet_name=None, engine='openpyxl')
        return xls
    except Exception as e:
        st.error(f"读取 Excel 出错: {e}")
        return {}

def auto_detect_columns(df):
    # 常见列名的映射
    cols = {c.lower(): c for c in df.columns}
    mapping = {}
    def find(name_options):
        for opt in name_options:
            if opt in cols:
                return cols[opt]
        return None
    mapping['order_date'] = find(['orderdate','order_date','date','交易日期','下单日期'])
    mapping['customer'] = find(['customer','客户','client','buyer'])
    mapping['product'] = find(['product','sku','item','货品','商品'])
    mapping['quantity'] = find(['quantity','qty','数量'])
    mapping['unit_price'] = find(['unitprice','unit_price','price','单价','价格'])
    mapping['sales'] = find(['sales','amount','金额','total','总额'])
    return mapping

@st.cache_data
def compute_basic_metrics(df, date_col, cust_col, qty_col, price_col, sales_col):
    d = df.copy()
    # Parse date
    if date_col:
        d[date_col] = pd.to_datetime(d[date_col], errors='coerce')
    # Ensure numeric
    if qty_col:
        d[qty_col] = pd.to_numeric(d[qty_col], errors='coerce').fillna(0)
    if price_col:
        d[price_col] = pd.to_numeric(d[price_col], errors='coerce').fillna(0)
    if sales_col:
        d[sales_col] = pd.to_numeric(d[sales_col], errors='coerce')
    else:
        # 计算 Sales = qty * unit_price
        if qty_col and price_col:
            d['Sales_Calc'] = d[qty_col] * d[price_col]
            sales_col = 'Sales_Calc'
    total_sales = d[sales_col].sum() if sales_col in d.columns else 0
    order_count = len(d)
    avg_order = total_sales / order_count if order_count else 0
    top_customers = None
    if cust_col and sales_col in d.columns:
        top_customers = d.groupby(cust_col)[sales_col].sum().sort_values(ascending=False)
    return d, total_sales, order_count, avg_order, top_customers, sales_col

# ---------- 上传与读取 ----------
st.sidebar.header("上传或读取数据")
uploaded_file = st.sidebar.file_uploader("上传 Excel 文件 (.xlsx/.xls)", type=['xlsx','xls'])

sheet_dfs = {}
if uploaded_file:
    sheet_dfs = read_excel_file(uploaded_file)
    sheet_names = list(sheet_dfs.keys())
    sheet_choice = st.sidebar.selectbox("选择工作表", sheet_names)
    df = sheet_dfs[sheet_choice]
else:
    st.sidebar.info("请在左侧上传 Excel 文件以开始分析。你也可以将示例数据粘贴到下面区域进行测试。")
    sample_text = st.sidebar.text_area("粘贴 CSV 示例（可选）", height=120)
    if sample_text:
        try:
            df = pd.read_csv(io.StringIO(sample_text))
            st.sidebar.success('已读取示例 CSV 数据')
        except Exception as e:
            st.sidebar.error('读取示例 CSV 失败: ' + str(e))
            df = None
    else:
        df = None

if df is None:
    st.warning("尚未加载数据 — 请上传 Excel 或粘贴 CSV 示例。")
    st.stop()

st.subheader("原始数据预览（前 50 行）")
st.dataframe(df.head(50))

# ---------- 列映射 & 清洗 ----------
st.sidebar.header("列映射与清洗")
col_map = auto_detect_columns(df)
st.sidebar.write("自动检测（若不正确，请手动选择）")

date_col = st.sidebar.selectbox("订单日期列（用于时间分析）", options=[None] + list(df.columns), index=0 if col_map['order_date'] is None else list(df.columns).index(col_map['order_date'])+1)
cust_col = st.sidebar.selectbox("客户列", options=[None] + list(df.columns), index=0 if col_map['customer'] is None else list(df.columns).index(col_map['customer'])+1)
product_col = st.sidebar.selectbox("产品列", options=[None] + list(df.columns), index=0 if col_map['product'] is None else list(df.columns).index(col_map['product'])+1)
qty_col = st.sidebar.selectbox("数量列", options=[None] + list(df.columns), index=0 if col_map['quantity'] is None else list(df.columns).index(col_map['quantity'])+1)
price_col = st.sidebar.selectbox("单价列", options=[None] + list(df.columns), index=0 if col_map['unit_price'] is None else list(df.columns).index(col_map['unit_price'])+1)
sales_col = st.sidebar.selectbox("销售额列（若无可留空，系统会尝试计算）", options=[None] + list(df.columns), index=0 if col_map['sales'] is None else list(df.columns).index(col_map['sales'])+1)

# ---------- 基本指标计算 ----------
st.sidebar.header("分析设置")
resample_period = st.sidebar.selectbox("时间重采样周期（趋势）", options=['M','W','D','Q'], index=0, help='M=按月, W=按周, D=按日, Q=按季度')
show_top_n = st.sidebar.number_input("展示前 N 名客户/产品", min_value=1, max_value=50, value=5)

with st.spinner('计算指标中...'):
    clean_df, total_sales, order_count, avg_order, top_customers, sales_col_used = compute_basic_metrics(df, date_col, cust_col, qty_col, price_col, sales_col)

# ---------- KPI 卡片 ----------
st.markdown("### 关键指标")
col1, col2, col3, col4 = st.columns(4)
col1.metric("总销售额", f"{total_sales:,.2f}")
col2.metric("订单数量", f"{order_count:,}")
col3.metric("平均每单金额", f"{avg_order:,.2f}")
if date_col:
    earliest = clean_df[date_col].min()
    latest = clean_df[date_col].max()
    col4.metric("时间范围", f"{earliest.date() if pd.notna(earliest) else '-'} — {latest.date() if pd.notna(latest) else '-'}")

# ---------- 过滤器 ----------
st.sidebar.header("数据过滤")
filters = {}
# 时间过滤
if date_col:
    min_date = pd.to_datetime(clean_df[date_col], errors='coerce').min()
    max_date = pd.to_datetime(clean_df[date_col], errors='coerce').max()
    if pd.notna(min_date) and pd.notna(max_date):
        date_range = st.sidebar.date_input("订单日期范围", value=(min_date.date(), max_date.date()))
        if len(date_range) == 2:
            start_dt = pd.to_datetime(date_range[0])
            end_dt = pd.to_datetime(date_range[1])
            filters['date'] = (start_dt, end_dt)
# 客户过滤
if cust_col:
    unique_customers = clean_df[cust_col].dropna().unique().tolist()
    sel_customers = st.sidebar.multiselect("选择客户（不选表示全部）", options=unique_customers, default=None)
    if sel_customers:
        filters['customers'] = sel_customers
# 产品过滤
if product_col:
    unique_products = clean_df[product_col].dropna().unique().tolist()
    sel_products = st.sidebar.multiselect("选择产品（不选表示全部）", options=unique_products, default=None)
    if sel_products:
        filters['products'] = sel_products

# 应用过滤
df_filtered = clean_df.copy()
if 'date' in filters and date_col:
    sdt, edt = filters['date']
    df_filtered = df_filtered[(df_filtered[date_col] >= sdt) & (df_filtered[date_col] <= edt + pd.Timedelta(days=1))]
if 'customers' in filters and cust_col:
    df_filtered = df_filtered[df_filtered[cust_col].isin(filters['customers'])]
if 'products' in filters and product_col:
    df_filtered = df_filtered[df_filtered[product_col].isin(filters['products'])]

st.subheader(f"筛选后数据预览（{len(df_filtered)} 行）")
st.dataframe(df_filtered.head(100))

# ---------- 列表与排名 ----------
st.markdown("### 排名与汇总")
rank_col1, rank_col2 = st.columns(2)
with rank_col1:
    if cust_col and sales_col_used in df_filtered.columns:
        top_cust = df_filtered.groupby(cust_col)[sales_col_used].sum().sort_values(ascending=False).head(show_top_n)
        st.write(f"前 {show_top_n} 大客户（按销售额）")
        st.table(top_cust.reset_index().rename(columns={cust_col: '客户', sales_col_used: '销售额'}))
with rank_col2:
    if product_col and sales_col_used in df_filtered.columns:
        top_prod = df_filtered.groupby(product_col)[sales_col_used].sum().sort_values(ascending=False).head(show_top_n)
        st.write(f"前 {show_top_n} 热销产品（按销售额）")
        st.table(top_prod.reset_index().rename(columns={product_col: '产品', sales_col_used: '销售额'}))

# ---------- 时间序列趋势 ----------
st.markdown("### 时间序列趋势")
if date_col and sales_col_used in df_filtered.columns:
    ts = df_filtered.set_index(pd.to_datetime(df_filtered[date_col], errors='coerce'))
    ts_resampled = ts[sales_col_used].resample(resample_period).sum()
    fig, ax = plt.subplots(figsize=(10,4))
    ax.plot(ts_resampled.index, ts_resampled.values)
    ax.set_title('销售额时间趋势')
    ax.set_xlabel('时间')
    ax.set_ylabel('销售额')
    plt.tight_layout()
    st.pyplot(fig)
else:
    st.info('缺少时间列或销售额列，无法绘制时间序列趋势。')

# ---------- 柱状图：按月/客户/产品 ----------
st.markdown("### 柱状图：分组汇总")
plot_choice = st.selectbox('选择要绘制柱状图的维度', options=['按客户','按产品','按月份'])
if plot_choice == '按客户' and cust_col and sales_col_used in df_filtered.columns:
    series = df_filtered.groupby(cust_col)[sales_col_used].sum().sort_values(ascending=False).head(20)
    fig, ax = plt.subplots(figsize=(10,5))
    ax.bar(series.index.astype(str), series.values)
    ax.set_xticklabels(series.index.astype(str), rotation=45, ha='right')
    ax.set_title('客户销售额排名（Top 20）')
    st.pyplot(fig)
elif plot_choice == '按产品' and product_col and sales_col_used in df_filtered.columns:
    series = df_filtered.groupby(product_col)[sales_col_used].sum().sort_values(ascending=False).head(20)
    fig, ax = plt.subplots(figsize=(10,5))
    ax.bar(series.index.astype(str), series.values)
    ax.set_xticklabels(series.index.astype(str), rotation=45, ha='right')
    ax.set_title('产品销售额排名（Top 20）')
    st.pyplot(fig)
elif plot_choice == '按月份' and date_col and sales_col_used in df_filtered.columns:
    s = pd.to_datetime(df_filtered[date_col]).dt.to_period('M').astype(str)
    series = df_filtered.groupby(s)[sales_col_used].sum()
    fig, ax = plt.subplots(figsize=(10,5))
    ax.bar(series.index, series.values)
    ax.set_xticklabels(series.index, rotation=45, ha='right')
    ax.set_title('每月销售额')
    st.pyplot(fig)
else:
    st.info('请选择合适的维度并确保数据列存在。')

# ---------- 散点图：数量 vs 单价（可视化异常/聚类） ----------
st.markdown('### 散点图：数量 vs 单价')
if qty_col and price_col and qty_col in df_filtered.columns and price_col in df_filtered.columns:
    fig, ax = plt.subplots(figsize=(8,5))
    ax.scatter(df_filtered[qty_col], df_filtered[price_col], alpha=0.6)
    ax.set_xlabel('数量')
    ax.set_ylabel('单价')
    ax.set_title('数量 vs 单价 散点图')
    st.pyplot(fig)
else:
    st.info('缺少数量列或单价列，无法绘制散点图。')

# ---------- 热力图：相关性矩阵 ----------
st.markdown('### 热力图：数值字段相关性')
numeric_df = df_filtered.select_dtypes(include=[np.number])
if not numeric_df.empty:
    corr = numeric_df.corr()
    fig, ax = plt.subplots(figsize=(8,6))
    sns.heatmap(corr, annot=True, fmt='.2f', ax=ax)
    ax.set_title('数值字段相关性热力图')
    st.pyplot(fig)
else:
    st.info('筛选后无数值字段用于计算相关性。')

# ---------- 数据导出 ----------
st.markdown('### 导出与下载')
if not df_filtered.empty:
    csv = df_filtered.to_csv(index=False).encode('utf-8')
    st.download_button('下载筛选后数据 CSV', data=csv, file_name='filtered_data.csv', mime='text/csv')

# ---------- 额外功能建议（留白供扩展） ----------
st.markdown('---')
st.write('提示：你可以在此基础上扩展：添加更多 KPI（毛利、毛利率）、对账功能、退货处理、客户分层（RFM）、地图可视化、导入多币种并自动换算等。')

# 结束
st.write('分析完成 — 如需我把这个部署为网页（Heroku/Streamlit Cloud/Docker）、或改造成 Flask+React 的完整系统，我可以继续提供代码。')

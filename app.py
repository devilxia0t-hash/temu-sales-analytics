from io import BytesIO
import pandas as pd
import plotly.express as px
import streamlit as st
from analytics import (ALIASES, REQUIRED, demo_data, read_report, infer_mapping,
                       normalize, daily_summary, ranking, signals, export_excel)

st.set_page_config(page_title='Temu 采购订单分析', page_icon='📊', layout='wide')
st.title('Temu 全托管采购订单分析')
st.warning('本系统分析平台采购订单。采购件数不等于消费者实际销量，采购金额不等于真实销售额、回款或利润。')
with st.sidebar:
    st.header('数据来源')
    source = st.radio('选择数据', ['模拟数据', '上传报表'])
    st.caption('模拟数据仅用于体验，不代表真实经营情况。上传文件仅在当前服务内存中处理。')
    uploaded = st.file_uploader('Excel / CSV 报表', type=['xlsx', 'xls', 'csv']) if source == '上传报表' else None
    st.download_button('下载模拟 CSV', demo_data().to_csv(index=False).encode('utf-8-sig'), '模拟采购订单.csv', 'text/csv')

if source == '上传报表' and uploaded is None:
    st.info('请上传订单报表，或在左侧选择模拟数据。')
    st.stop()
try:
    if source == '模拟数据':
        raw = demo_data()
    else:
        data = uploaded.getvalue()
        sheet = 0
        if uploaded.name.lower().endswith(('.xlsx', '.xls')):
            sheets = pd.ExcelFile(BytesIO(data)).sheet_names
            sheet = st.sidebar.selectbox('工作表', sheets)
        raw = read_report(data, uploaded.name, sheet)
except Exception as exc:
    st.error(f'文件读取失败，请检查文件格式、编码及是否加密：{exc}')
    st.stop()
if raw.empty:
    st.info('报表没有数据行。')
    st.stop()
with st.expander('原始数据预览与字段匹配', expanded=True):
    st.dataframe(raw.head(20), hide_index=True)
    st.caption('必填：日期、SKU、数量、供货单价、状态。供货价必须匹配每件单价，不能使用行金额。可选维度缺失时归入“未提供”，不会从 SKU 猜测。')
    detected = infer_mapping(raw.columns)
    options = [None] + list(raw.columns)
    mapping = {}
    columns = st.columns(3)
    for i, field in enumerate(ALIASES):
        mapping[field] = columns[i % 3].selectbox(field + (' *' if field in REQUIRED else '（可选）'), options,
              index=options.index(detected[field]), format_func=lambda x: '未匹配' if x is None else str(x), key=f'map_{field}')
try:
    valid, rejected = normalize(raw, mapping)
except (ValueError, TypeError) as exc:
    st.error(str(exc))
    st.stop()
st.caption(f'读取 {len(raw)} 行；有效 {len(valid)} 行；无效 {len(rejected)} 行。源数据行以首行为表头计数。')
if not rejected.empty:
    with st.expander('查看无效数据（不纳入分析）'):
        st.dataframe(rejected, hide_index=True)
if valid.empty:
    st.error('没有可分析的有效数据，请检查字段匹配和无效原因。')
    st.stop()
with st.sidebar:
    st.header('统计范围')
    statuses = sorted(valid['状态'].dropna().unique())
    defaults = [s for s in statuses if not any(word in s.casefold() for word in ['取消', '关闭', '作废', 'cancel', 'closed'])]
    selected = st.multiselect('纳入的采购订单状态', statuses, default=defaults)
    st.caption('默认排除含“取消、关闭、作废、cancel、closed”的状态，请根据实际业务复核。')
    bounds = (valid['日期'].min().date(), valid['日期'].max().date())
    dates = st.date_input('观察日期范围', value=bounds, min_value=bounds[0], max_value=bounds[1])
    currency = st.text_input('金额币种 / 单位（不做汇率换算）', '人民币元')
if len(dates) != 2:
    st.info('请选择起止日期。')
    st.stop()
start, end = map(pd.Timestamp, dates)
frame = valid.loc[valid['状态'].isin(selected) & valid['日期'].between(start, end)].copy()
if frame.empty:
    st.info('当前日期和状态筛选下没有有效采购记录。')
    st.stop()
has_id = mapping.get('订单号') is not None
if has_id and frame['订单号'].isna().any():
    st.error('已匹配订单号，但筛选结果存在空订单号。请修正源数据，或取消订单号匹配，按明细行数统计。')
    st.stop()
if not has_id:
    st.info('未提供订单号，仅统计采购明细行数，不能将其称为采购订单数。')
st.caption('同一订单多行会保留；有订单号时订单数按唯一订单号统计，每日按当日去重。重复导出的行不会自动删除，请避免重复导入。')
daily = daily_summary(frame, start, end, has_id)
changes, anomalies = signals(daily)
metrics = st.columns(3)
metrics[0].metric('采购订单数' if has_id else '采购明细行数', f"{frame['订单号'].nunique() if has_id else len(frame):,}")
metrics[1].metric('采购件数', f"{frame['数量'].sum():,.0f}")
metrics[2].metric(f'采购金额（{currency}）', f"{frame['采购金额'].sum():,.2f}")
st.caption(f'采购金额 = 每行数量 × 供货单价（每行四舍五入至2位）再汇总；统计截至 {end:%Y-%m-%d}，最后一天可能尚未结束。缺失日期补0仅代表报表无记录，不能证明实际无采购。')
tabs = st.tabs(['订单趋势', '商品排行', '增长与异常', '明细与导出'])
with tabs[0]:
    window = st.radio('趋势窗口（以所选结束日为基准）', [7, 30], format_func=lambda x: f'近{x}天', horizontal=True)
    measure = st.selectbox('趋势指标', daily.columns)
    chart = daily.loc[daily.index >= end - pd.Timedelta(days=window-1)].reset_index()
    st.plotly_chart(px.line(chart, x='日期', y=measure, markers=True), use_container_width=True)
    st.dataframe(daily.reset_index(), hide_index=True)
with tabs[1]:
    dimension = st.radio('排行维度', ['SPU', 'SKU', '颜色', '尺码'], horizontal=True)
    ranks = ranking(frame, dimension)
    st.plotly_chart(px.bar(ranks.head(20), x=dimension, y='采购件数', title='采购件数前20名'), use_container_width=True)
    st.dataframe(ranks, hide_index=True)
with tabs[2]:
    st.caption('比较观察范围末尾7/30天与紧邻的上一等长周期；增长率为采购件数增长率。观察天数不足则不计算。')
    st.dataframe(changes.style.format({'增长率': '{:.1%}'}, na_rep='不适用'), hide_index=True)
    st.caption('异常规则：当日采购件数偏离此前7个自然日均值超过3倍总体标准差，且绝对偏差至少5件。仅为排查线索；缺报、节假日和未结束日期均可能产生提示。')
    if anomalies.empty:
        st.info('当前范围未检测到符合规则的异常。')
    else:
        st.dataframe(anomalies.reset_index(), hide_index=True)
with tabs[3]:
    st.dataframe(frame, hide_index=True)
    metadata = [('数据性质', '模拟数据' if source == '模拟数据' else '用户上传采购报表'),
                ('业务口径', '平台采购订单；非消费者实际销量、销售额、回款或利润'),
                ('订单计数', '按唯一订单号统计；每日按当日去重' if has_id else '未提供订单号，仅统计明细行'),
                ('日期范围', f'{start:%Y-%m-%d} 至 {end:%Y-%m-%d}'), ('纳入状态', '、'.join(selected)),
                ('币种/单位', currency), ('缺失日期', '观察范围内补0，不证明实际无采购'),
                ('金额公式', '数量×供货单价；每行保留2位后求和'), ('重复数据', '不自动删除重复行'),
                ('无效数据', f'{len(rejected)}行已排除；导出无效数据表包含整个上传文件的无效行'),
                ('增长异常规则', '7/30天等长环比；偏离此前7日均值超过3倍标准差且至少5件')]
    st.download_button('导出 Excel 分析结果', export_excel(frame, daily, changes, anomalies, rejected, metadata),
                       'Temu采购订单分析.xlsx', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

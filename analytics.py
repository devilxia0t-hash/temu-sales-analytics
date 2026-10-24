"""采购订单分析：所有金额均为供货采购口径。"""
from io import BytesIO
import numpy as np
import pandas as pd

ALIASES = {
    '日期': ['订单日期', '下单时间', '创建时间', '采购日期', '订单创建时间', 'date'],
    'SKU': ['SKU', 'sku', 'SKU编码', '货品SKU', '商品SKU'],
    '数量': ['数量', '采购数量', '件数', '商品数量', 'quantity'],
    '供货价': ['供货价', '供货单价', '采购单价', '单价', 'price'],
    '状态': ['订单状态', '采购状态', '状态', 'status'],
    '订单号': ['订单号', '采购订单号', '订单编号', 'order_id'],
    'SPU': ['SPU', 'spu', 'SPU编码', '商品编号'],
    '颜色': ['颜色', '色号', 'color'],
    '尺码': ['尺码', '尺寸', 'size'],
}
REQUIRED = ['日期', 'SKU', '数量', '供货价', '状态']

def infer_mapping(columns):
    lookup = {str(c).strip().casefold(): c for c in columns}
    return {field: next((lookup[a.casefold()] for a in aliases if a.casefold() in lookup), None)
            for field, aliases in ALIASES.items()}

def read_report(data, filename, sheet=0):
    if filename.lower().endswith('.csv'):
        for encoding in ('utf-8-sig', 'gb18030'):
            try:
                return pd.read_csv(BytesIO(data), encoding=encoding, dtype=str)
            except UnicodeDecodeError:
                continue
        raise ValueError('CSV 编码无法识别，请另存为 UTF-8 CSV。')
    return pd.read_excel(BytesIO(data), sheet_name=sheet, dtype=str)

def normalize(raw, mapping):
    missing = [k for k in REQUIRED if not mapping.get(k)]
    if missing:
        raise ValueError('请匹配必填字段：' + '、'.join(missing))
    selected = [v for v in mapping.values() if v is not None]
    if len(selected) != len(set(selected)):
        raise ValueError('一个原始列不能同时匹配多个字段。')
    frame = pd.DataFrame({k: raw[v] for k, v in mapping.items() if v is not None})
    frame.insert(0, '源数据行', range(2, len(frame) + 2))
    for key in ['SKU', '状态', '订单号', 'SPU', '颜色', '尺码']:
        if key not in frame:
            frame[key] = pd.NA
        frame[key] = frame[key].astype('string').str.strip().replace('', pd.NA)
    dates = pd.to_datetime(frame['日期'], errors='coerce', format='mixed')
    # Excel serial dates may arrive as strings when the source cell has no date format.
    serial = pd.to_numeric(frame['日期'], errors='coerce')
    serial_mask = serial.between(20000, 80000)
    dates.loc[serial_mask] = pd.to_datetime(serial[serial_mask], unit='D', origin='1899-12-30')
    frame['日期'] = dates.dt.normalize()
    for key in ['数量', '供货价']:
        frame[key] = pd.to_numeric(frame[key].astype('string').str.replace(r'[,，¥￥\s]', '', regex=True), errors='coerce')
    reasons = pd.Series('', index=frame.index)
    checks = {
        '日期无效': frame['日期'].isna(),
        'SKU为空': frame['SKU'].isna(),
        '状态为空': frame['状态'].isna(),
        '数量必须为正整数': frame['数量'].isna() | ~np.isfinite(frame['数量'].fillna(0)) | (frame['数量'] <= 0) | (frame['数量'] % 1 != 0),
        '供货价必须为非负有限数': frame['供货价'].isna() | ~np.isfinite(frame['供货价'].fillna(0)) | (frame['供货价'] < 0),
    }
    for reason, mask in checks.items():
        reasons.loc[mask] += reason + '；'
    rejected = raw.loc[reasons.ne('')].copy()
    rejected.insert(0, '源数据行', frame.loc[reasons.ne(''), '源数据行'])
    rejected['无效原因'] = reasons[reasons.ne('')]
    valid = frame.loc[reasons.eq('')].copy()
    valid['采购金额'] = (valid['数量'] * valid['供货价']).round(2)
    return valid, rejected

def daily_summary(frame, start, end, has_order_id):
    index = pd.date_range(start, end, freq='D', name='日期')
    grouped = frame.groupby('日期')
    count = grouped['订单号'].nunique() if has_order_id else grouped.size()
    result = pd.DataFrame({'采购订单数' if has_order_id else '采购明细行数': count,
                           '采购件数': grouped['数量'].sum(), '采购金额': grouped['采购金额'].sum()})
    return result.reindex(index, fill_value=0).fillna(0)

def ranking(frame, field):
    return (frame.assign(**{field: frame[field].fillna('未提供')})
            .groupby(field, dropna=False).agg(采购件数=('数量', 'sum'), 采购金额=('采购金额', 'sum'), 采购明细行数=('SKU', 'size'))
            .sort_values('采购件数', ascending=False).reset_index())

def signals(daily):
    """Compare complete equal-length windows; anomalies use preceding 7 days only."""
    values = daily['采购件数']
    previous = values.shift(1).rolling(7, min_periods=7)
    mean, std = previous.mean(), previous.std(ddof=0)
    mask = ((values - mean).abs() > 3 * std) & ((values - mean).abs() >= 5) & mean.notna()
    anomalies = daily.loc[mask].copy()
    anomalies['前7日均值'] = mean[mask]
    anomalies['方向'] = np.where(values[mask] > mean[mask], '异常增长', '异常下降')
    changes = []
    for days in (7, 30):
        if len(values) < days * 2:
            changes.append({'周期': f'{days}天', '本期采购件数': None, '上期采购件数': None, '增长率': None, '说明': f'至少需要{days * 2}天观察范围'})
            continue
        current, prior = values.iloc[-days:].sum(), values.iloc[-2*days:-days].sum()
        changes.append({'周期': f'{days}天', '本期采购件数': current, '上期采购件数': prior,
                        '增长率': (current / prior - 1) if prior else None,
                        '说明': '上期为0，增长率不适用' if prior == 0 else ('增长' if current > prior else '下降' if current < prior else '持平')})
    return pd.DataFrame(changes), anomalies

def export_excel(frame, daily, changes, anomalies, rejected, metadata):
    output = BytesIO()
    tables = {'分析口径': pd.DataFrame(metadata, columns=['项目', '说明']), '每日采购汇总': daily.reset_index(),
              '增长分析': changes, '异常日期': anomalies.reset_index(), '有效采购明细': frame,
              '无效数据': rejected}
    tables.update({f'{field}排行': ranking(frame, field) for field in ['SPU', 'SKU', '颜色', '尺码']})
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for name, table in tables.items():
            table.to_excel(writer, sheet_name=name, index=False)
            ws = writer.sheets[name]
            ws.freeze_panes = 'A2'
            ws.auto_filter.ref = ws.dimensions
            # Treat uploaded text as text, including spreadsheet formula prefixes.
            for row in ws.iter_rows():
                for cell in row:
                    if isinstance(cell.value, str):
                        cell.data_type = 's'
    return output.getvalue()

def demo_data():
    rng = np.random.default_rng(42)
    rows = []
    for day in pd.date_range(pd.Timestamp.today().normalize() - pd.Timedelta(days=89), periods=90):
        for i in range(int(rng.integers(5, 15))):
            sku = int(rng.integers(1, 13))
            rows.append({'订单日期': day, '采购订单号': f'PO{day:%Y%m%d}{i:03}',
                         'SKU': f'SKU{sku:03}', 'SPU': f'SPU{(sku-1)//4+1:02}',
                         '颜色': ['黑色', '白色', '蓝色'][sku % 3], '尺码': ['S', 'M', 'L', 'XL'][sku % 4],
                         '采购数量': int(rng.integers(1, 20)), '供货单价': round(10 + sku * 2.5, 2),
                         '订单状态': rng.choice(['已完成', '待发货', '已取消'], p=[.7, .25, .05])})
    return pd.DataFrame(rows)

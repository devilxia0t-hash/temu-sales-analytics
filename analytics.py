"""采购订单分析：所有金额均为供货采购口径。"""
from io import BytesIO
import numpy as np
import pandas as pd
import re
import unicodedata
import traceback
from excel_reader import sheets, xml_grid, merged_ranges
from openpyxl.utils.cell import range_boundaries

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
    '条码': ['条码', '商品条码', 'barcode', 'EAN', 'UPC'],
}
REQUIRED = ['日期', 'SKU', '数量', '供货价', '状态']

def clean_text(value):
    return re.sub(r'[\s\u200b\ufeff]+', '', unicodedata.normalize('NFKC', str(value))).casefold()


def infer_mapping(columns):
    lookup = {}
    for c in columns:
        lookup.setdefault(clean_text(c), c)
        lookup.setdefault(clean_text(str(c).split(' / ')[-1]), c)
    return {field: next((lookup[clean_text(a)] for a in aliases if clean_text(a) in lookup), None)
            for field, aliases in ALIASES.items()}


def list_sheets(data, filename):
    if filename.lower().endswith('.xlsx'):
        return [name for name, _ in sheets(data)[0]]
    with pd.ExcelFile(BytesIO(data)) as book:
        return book.sheet_names


def load_grid(data, filename, sheet=0):
    diagnostics = []
    epoch = '1899-12-30'
    if filename.lower().endswith('.csv'):
        for encoding in ('utf-8-sig', 'gb18030'):
            try:
                grid = pd.read_csv(BytesIO(data), encoding=encoding, dtype=str, header=None,
                                   keep_default_na=False, skip_blank_lines=False)
                break
            except UnicodeDecodeError:
                continue
        else:
            raise ValueError('CSV 编码无法识别，请另存为 UTF-8 CSV。')
    else:
        if filename.lower().endswith('.xlsx'):
            epoch = sheets(data)[1]
        try:
            grid = pd.read_excel(BytesIO(data), sheet_name=sheet, dtype=str, header=None, keep_default_na=False)
        except Exception as exc:
            if not filename.lower().endswith('.xlsx'):
                raise ValueError(f'Excel读取失败：{type(exc).__name__}: {exc}；请用Excel另存为.xlsx重试。') from exc
            trace = traceback.format_exc()
            grid, issues, epoch = xml_grid(data, sheet)
            diagnostics = [f'标准Excel解析失败：{type(exc).__name__}: {exc}', trace,
                           '已切换OOXML原始值读取；不依赖样式，数值不强制转整数。日期列中的Excel序号按工作簿日期系统转换。'] + issues
    grid.attrs.update(diagnostics=diagnostics, excel_epoch=epoch)
    if filename.lower().endswith('.xlsx'):
        grid.attrs['merged_ranges'] = merged_ranges(data, sheet)
    return grid


def detect_header(grid):
    best, score = 0, -1
    for index, row in grid.head(50).iterrows():
        found = infer_mapping(row.dropna().tolist())
        current = sum(found.get(k) is not None for k in REQUIRED)
        if current > score:
            best, score = index, current
    return best


def report_from_grid(grid, header=0, header_rows=1):
    if grid.empty:
        return grid.copy()
    if not 0 <= header < len(grid):
        raise ValueError('表头行超出文件范围。')
    if not 1 <= header_rows <= 5 or header + header_rows > len(grid):
        raise ValueError('表头层数必须为1至5，且不能超过表格末行。')
    block = grid.iloc[header:header+header_rows].map(lambda v: pd.NA if isinstance(v, str) and not v.strip() else v)
    if header_rows > 1:
        block = block.ffill(axis=1)
    values = []
    for col in block.columns:
        parts = list(dict.fromkeys(str(v).strip() for v in block[col] if pd.notna(v)))
        values.append(' / '.join(parts))
    names, seen = [], {}
    for i, value in enumerate(values):
        name = str(value).strip() if pd.notna(value) and str(value).strip() else f'未命名列{i+1}'
        seen[name] = seen.get(name, 0) + 1
        label = name if seen[name] == 1 else f'{name} [{seen[name]}]'
        while label in names:
            seen[name] += 1
            label = f'{name} [{seen[name]}]'
        names.append(label)
    frame = grid.iloc[header+header_rows:].copy()
    frame.columns = names
    frame = frame.map(lambda v: pd.NA if isinstance(v, str) and not v.strip() else v).dropna(how='all')
    frame.attrs.update(grid.attrs)
    frame.attrs['source_rows'] = [int(i)+1 for i in frame.index]
    return frame.reset_index(drop=True)


def read_report(data, filename, sheet=0, header=0):
    return report_from_grid(load_grid(data, filename, sheet), header)


def fill_merged_text(raw, mapping):
    """Only fill explicitly merged, vertical text fields; never infer numeric data."""
    result = raw.copy()
    source = dict(zip(raw.attrs.get('source_rows', []), raw.index))
    fields = ['订单号', 'SKU', '条码', 'SPU', '颜色', '尺码', '状态']
    for ref in raw.attrs.get('merged_ranges', []):
        left, top, right, bottom = range_boundaries(ref)
        if left != right or top not in source or left > len(raw.columns):
            continue
        column = raw.columns[left-1]
        if not any(mapping.get(field) == column for field in fields):
            continue
        value = raw.at[source[top], column]
        if pd.isna(value):
            continue
        for row in range(top+1, bottom+1):
            if row in source and pd.isna(result.at[source[row], column]):
                result.at[source[row], column] = value
    return result


def normalize(raw, mapping):
    missing = [k for k in REQUIRED if not mapping.get(k)]
    if missing:
        raise ValueError('请匹配必填字段：' + '、'.join(missing))
    selected = [v for v in mapping.values() if v is not None]
    if len(selected) != len(set(selected)):
        raise ValueError('一个原始列不能同时匹配多个字段。')
    frame = pd.DataFrame({k: raw[v] for k, v in mapping.items() if v is not None})
    frame.insert(0, '源数据行', raw.attrs.get('source_rows', range(2, len(frame) + 2)))
    for key in ['SKU', '状态', '订单号', 'SPU', '颜色', '尺码', '条码']:
        if key not in frame:
            frame[key] = pd.NA
        frame[key] = frame[key].astype('string').str.strip().replace('', pd.NA)
    date_text = frame['日期'].astype('string').str.replace('年', '-', regex=False).str.replace('月', '-', regex=False).str.replace('日', '', regex=False)
    def parse_local_date(value):
        try:
            parsed = pd.Timestamp(value)
            return parsed.tz_localize(None) if parsed.tzinfo else parsed
        except (ValueError, TypeError, OverflowError):
            return pd.NaT
    dates = pd.to_datetime(date_text.map(parse_local_date), errors='coerce').copy()
    # Excel serial dates may arrive as strings when the source cell has no date format.
    serial = pd.to_numeric(frame['日期'], errors='coerce')
    serial_mask = serial.between(20000, 80000)
    dates.loc[serial_mask] = pd.to_datetime(serial[serial_mask], unit='D', origin=raw.attrs.get('excel_epoch', '1899-12-30'))
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
    def excel_safe(value):
        if isinstance(value, str):
            return re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f]', lambda m: f'\\x{ord(m.group()):02x}', value)
        return value
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for name, table in tables.items():
            table = table.map(excel_safe).rename(columns=excel_safe)
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

"""OOXML fallback: retain malformed numeric values as text, never discard rows."""
from io import BytesIO
import posixpath
from zipfile import ZipFile
from xml.etree import ElementTree as ET
import pandas as pd
from openpyxl.utils.cell import coordinate_to_tuple

NS = {'m': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'


def package(data):
    z = ZipFile(BytesIO(data))
    if sum(i.file_size for i in z.infolist()) > 256 * 1024 * 1024:
        z.close()
        raise ValueError('Excel 解压内容超过256MB安全限制，请拆分文件。')
    return z


def sheets(data):
    with package(data) as z:
        book = ET.fromstring(z.read('xl/workbook.xml'))
        relationships = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
        targets = {r.attrib['Id']: r.attrib['Target'] for r in relationships}
        entries = []
        for s in book.findall('m:sheets/m:sheet', NS):
            path = targets[s.attrib[f'{{{REL}}}id']]
            path = path.lstrip('/') if path.startswith('/') else posixpath.normpath(posixpath.join('xl', path))
            entries.append((s.attrib['name'], path))
        props = book.find('m:workbookPr', NS)
        epoch = '1904-01-01' if props is not None and props.get('date1904') in ('1', 'true') else '1899-12-30'
        return entries, epoch


def merged_ranges(data, sheet):
    entries, _ = sheets(data)
    _, path = entries[sheet] if isinstance(sheet, int) else next(e for e in entries if e[0] == sheet)
    with package(data) as z, z.open(path) as stream:
        ranges = []
        for _, element in ET.iterparse(stream, events=('end',)):
            if element.tag == f'{{{NS["m"]}}}mergeCell':
                ranges.append(element.attrib['ref'])
            element.clear()
        return ranges


def xml_grid(data, sheet):
    entries, epoch = sheets(data)
    name, path = entries[sheet] if isinstance(sheet, int) else next(e for e in entries if e[0] == sheet)
    with package(data) as z:
        strings = []
        if 'xl/sharedStrings.xml' in z.namelist():
            strings = [''.join(t.text or '' for t in si.iter(f'{{{NS["m"]}}}t'))
                       for si in ET.fromstring(z.read('xl/sharedStrings.xml'))]
        root = ET.fromstring(z.read(path))
        values = {}
        issues = []
        for row in root.findall('m:sheetData/m:row', NS):
            for c in row.findall('m:c', NS):
                ref = c.get('r')
                r, col = coordinate_to_tuple(ref)
                if r > 200000 or col > 1000:
                    raise ValueError(f'工作表 {name} 单元格 {ref} 超过读取限制（20万行/1000列），请拆分文件。')
                kind = c.get('t', 'n')
                v = c.find('m:v', NS)
                text = v.text if v is not None else None
                if kind == 's' and text is not None:
                    try:
                        index = int(text)
                        if index < 0:
                            raise IndexError(index)
                        text = strings[index]
                    except (ValueError, IndexError) as exc:
                        raise ValueError(f'工作表 {name} 单元格 {ref} 的共享字符串索引无效：{text!r}') from exc
                elif kind == 'inlineStr':
                    text = ''.join(t.text or '' for t in c.iter(f'{{{NS["m"]}}}t'))
                elif kind == 'n' and text is not None:
                    try:
                        float(text)
                    except ValueError:
                        issues.append(f'{name}!{ref} 标记为数值但内容为 {text!r}，已保留原文')
                if c.find('m:f', NS) is not None and text is None:
                    issues.append(f'{name}!{ref} 公式无缓存结果，读取为空；请在Excel中重算后保存')
                values[r-1, col-1] = text
        if not values:
            return pd.DataFrame(), [], epoch
        height, width = max(r for r, _ in values)+1, max(c for _, c in values)+1
        if height * width > 5_000_000:
            raise ValueError(f'工作表 {name} 有效区域超过500万单元格，请移除远端多余单元格或拆分文件。')
        grid = pd.DataFrame(index=range(height), columns=range(width), dtype=object)
        for (r, c), value in values.items():
            grid.iat[r, c] = value
        # Do not fill merged data cells: that could duplicate quantities/amounts.
        return grid, issues, epoch

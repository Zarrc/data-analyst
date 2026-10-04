"""零依赖 xlsx 读取器（纯 Python 标准库 zipfile + xml.etree）。

为什么不用 openpyxl/pandas.read_excel：
  本机未装 openpyxl / xlsxwriter / python-calamine 等任何 xlsx 引擎，
  pandas 3.0.6 自带无 xlsx 解析能力。xlsx 本质是 zip 里的 XML，
  用标准库即可解析，无需安装任何依赖（符合护栏 3「不擅自装依赖」）。

用法：
    from xlsx_read import read_xlsx
    shared, rows = read_xlsx("file.xlsx")
    # rows: [(rownum, {col_letter: value_or_None}), ...]  含表头行
"""
import zipfile
import re
from xml.etree import ElementTree as ET

NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'


def read_xlsx(path):
    """返回 (shared_strings, rows)。

    rows 为 [(row_number:int, {col_letter: value|None}), ...]，保持 Excel 行序。
    单元格值已解析：共享字符串(shared)、内联字符串(inlineStr)、数值/日期原始文本。
    注意：数值型日期会以 Excel 序列号(list)形式出现 —— 本脚本将 self-closing
    单元格判定为「空」，但真正的日期在本数据集中以字符串形式存储（Excel 导出），
    故无需序列号转换。此函数保持“原样返回”，转换在分析层处理。
    """
    z = zipfile.ZipFile(path)
    names = z.namelist()

    shared = []
    if 'xl/sharedStrings.xml' in names:
        root = ET.fromstring(z.read('xl/sharedStrings.xml'))
        for si in root.findall(NS + 'si'):
            txt = ''.join(t.text or '' for t in si.iter(NS + 't'))
            shared.append(txt)

    sheet_name = sorted(n for n in names if n.startswith('xl/worksheets/sheet'))[0]
    root = ET.fromstring(z.read(sheet_name))

    rows = []
    for row in root.iter(NS + 'row'):
        rnum = int(row.get('r')) if row.get('r') else None
        cells = {}
        for c in row.findall(NS + 'c'):
            ref = c.get('r')
            t = c.get('t')
            v = c.find(NS + 'v')
            isv = c.find(NS + 'is')
            val = None
            if t == 's' and v is not None:
                val = shared[int(v.text)]
            elif t == 'inlineStr' and isv is not None:
                val = ''.join(x.text or '' for x in isv.iter(NS + 't'))
            elif t == 'str' and v is not None:
                val = v.text
            elif v is not None:
                val = v.text
            col = re.match(r'[A-Z]+', ref).group(0)
            cells[col] = val
        rows.append((rnum, cells))
    return shared, rows


def rows_to_dicts(rows):
    """把 [(rownum, cells)] 转成 (header_list, [dict, ...])，跳过表头前的空行。"""
    # 表头 = 第一个非空行
    header = None
    header_rownum = None
    data = []
    for rnum, cells in rows:
        if header is None:
            if any((v is not None and str(v).strip() != '') for v in cells.values()):
                header = cells
                header_rownum = rnum
            continue
        if not any((v is not None and str(v).strip() != '') for v in cells.values()):
            continue  # 跳过整行空行
        data.append(cells)
    return header, header_rownum, data


if __name__ == '__main__':
    import sys
    shared, rows = read_xlsx(sys.argv[1])
    print('shared strings:', len(shared), 'rows:', len(rows))

"""在竞品明细的价格列（H）后插入「FBA配送费」列，后面各列右移一格；说明页在「价格」行后加说明。
纯 XML，嵌入图片不受影响。FBA 费用读自 /workspace/sorftime_cache.json。
用法: python3 add_fba_fee_col.py 输入.xlsx 输出.xlsx
"""
import zipfile, re, sys, os, json, html
from openpyxl.utils import column_index_from_string as ci, get_column_letter as gl

CACHE = '/workspace/sorftime_cache.json'
src, dst = sys.argv[1], sys.argv[2]
cache = {}
if os.path.exists(CACHE):
    cache = json.load(open(CACHE, encoding='utf-8'))

zi = zipfile.ZipFile(src)
s = zi.read('xl/worksheets/sheet1.xml').decode()

def header(col):
    m = re.search(rf'<c r="{col}1"[^>]*>.*?<t[^>]*>(.*?)</t>', s, re.S)
    return html.unescape(m.group(1)) if m else None

if header('H') != '价格':
    sys.exit(f'skip: H={header("H")}')
if header('I') == 'FBA配送费':
    sys.exit('already has FBA配送费')

# ASIN (D) -> row
asin_row = {}
for m in re.finditer(r'<c r="D(\d+)"[^>]*>(.*?)</c>', s, re.S):
    r = m.group(1)
    if r == '1':
        continue
    t = re.search(r'<t[^>]*>(.*?)</t>', m.group(2))
    if t:
        asin_row[int(r)] = html.unescape(t.group(1))

# 1) 把 I 及之后的列右移一格（单元格引用、超链接）
def sh(col):
    i = ci(col)
    return gl(i + 1) if i >= 9 else col

s = re.sub(r'(<c r=")([A-Z]+)(\d+")', lambda m: m.group(1) + sh(m.group(2)) + m.group(3), s)
s = re.sub(r'(<hyperlink ref=")([A-Z]+)(\d+")', lambda m: m.group(1) + sh(m.group(2)) + m.group(3), s)

# 2) 列宽：min/max>=9 的 +1，并在 H 后插入宽 9 的新 I 列
def col(m):
    prefix, a, b = m.group(1), int(m.group(2)), int(m.group(3))
    if a >= 9:
        a, b = a + 1, b + 1
    return f'{prefix}min="{a}" max="{b}"/>'

s = re.sub(r'(<col [^>]*?)min="(\d+)" max="(\d+)"/>', col, s)
# 插在 min=8 之后
s = re.sub(
    r'(<col [^>]*min="8" max="8"/>)',
    r'\1<col width="9" customWidth="1" min="9" max="9"/>',
    s,
    count=1,
)

# 3) dimension / autoFilter：末列 AD -> AE（或当前末列右移）
def bump_end(m):
    return f'{m.group(1)}{gl(ci(m.group(2))+1)}{m.group(3)}'

s = re.sub(r'(ref="A1:)([A-Z]+)(\d+")', bump_end, s)

# 4) 在每行 H 单元格后插入 I 列
def fee_of(asin):
    d = cache.get(asin) or {}
    fee = d.get('fba_fee') if isinstance(d, dict) else None
    if fee is None or fee == '' or (isinstance(fee, (int, float)) and fee < 0):
        return None
    try:
        return float(fee)
    except (TypeError, ValueError):
        return None

def insert_i(m):
    row = m.group(0)
    r = re.search(r'<row r="(\d+)"', row).group(1)
    if re.search(rf'<c r="I{r}"', row):
        return row
    if r == '1':
        cell = '<c r="I1" s="1" t="inlineStr"><is><t>FBA配送费</t></is></c>'
    else:
        fee = fee_of(asin_row.get(int(r)))
        if fee is None:
            cell = f'<c r="I{r}" s="6" t="n"></c>'
        else:
            fee2 = round(float(fee), 2)
            if fee2 == int(fee2):
                v = str(int(fee2))
            else:
                v = f'{fee2:.2f}'.rstrip('0').rstrip('.')
            cell = f'<c r="I{r}" s="6" t="n"><v>{v}</v></c>'
    # 插在 H 单元格之后；若无 H 则插在行内第一个单元格位置之后或行开头
    if re.search(rf'<c r="H{r}"[^>]*?(?:/>|>.*?</c>)', row, re.S):
        row = re.sub(rf'(<c r="H{r}"[^>]*?(?:/>|>.*?</c>))', r'\1' + cell, row, count=1, flags=re.S)
    else:
        row = re.sub(r'(<row r="\d+"[^>]*>)', r'\1' + cell, row, count=1)
    return row

s = re.sub(r'<row r="\d+"[^>]*>.*?</row>', insert_i, s, flags=re.S)

wb = zi.read('xl/workbook.xml').decode()
# FilterDatabase 末列 +1
wb = re.sub(
    r"('竞品明细'!\$A\$1:\$)([A-Z]+)(\$\d+)",
    lambda m: m.group(1) + gl(ci(m.group(2)) + 1) + m.group(3),
    wb,
)

# 说明页：在「价格」行后插入 FBA 说明，后续行号 +1
s3 = zi.read('xl/worksheets/sheet3.xml').decode()
NOTE_A = 'FBA配送费'
NOTE_B = (
    '来自 Sorftime 商品详情接口的 fba_fee（美元，按包装尺寸重量计算的亚马逊 FBA 配送费，查询日期 2026-11-16），'
    '无数据留空；自发货（FBM）商品同样给出若按 FBA 发货的费用（如果数据如此）。'
)

def esc(t):
    return html.escape(t, quote=False)

rows = re.findall(r'<row [^>]*>.*?</row>', s3, re.S)
new = []
inserted = False
for r in rows:
    a = re.search(r'<c r="A\d+"[^>]*>.*?<t[^>]*>(.*?)</t>', r, re.S)
    title = html.unescape(a.group(1)).strip() if a else ''
    new.append(r)
    if title == '价格' and not inserted:
        # 临时行号，后面统一重排
        new.append(
            '<row r="0">'
            f'<c r="A0" s="16" t="inlineStr"><is><t>{esc(NOTE_A)}</t></is></c>'
            f'<c r="B0" s="5" t="inlineStr"><is><t>{esc(NOTE_B)}</t></is></c>'
            '</row>'
        )
        inserted = True
if not inserted:
    sys.exit('说明页未找到「价格」行')

renum = []
for i, r in enumerate(new, 1):
    old = re.search(r'<row r="(\d+)"', r).group(1)
    r = re.sub(r'<row r="\d+"', f'<row r="{i}"', r, count=1)
    r = re.sub(rf'(<c r="[A-Z]+){old}"', rf'\g<1>{i}"', r)
    renum.append(r)
body = re.search(r'<sheetData>.*?</sheetData>', s3, re.S).group(0)
s3 = s3.replace(body, '<sheetData>' + ''.join(renum) + '</sheetData>')
s3 = re.sub(
    r'<dimension ref="A1:([A-Z]+)\d+"/>',
    lambda m: f'<dimension ref="A1:{m.group(1)}{len(renum)}"/>',
    s3,
)

zo = zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED)
for it in zi.infolist():
    d = zi.read(it.filename)
    d = {
        'xl/worksheets/sheet1.xml': s.encode(),
        'xl/workbook.xml': wb.encode(),
        'xl/worksheets/sheet3.xml': s3.encode(),
    }.get(it.filename, d)
    zo.writestr(it, d)
zo.close()

n_fee = n_empty = 0
for r, asin in asin_row.items():
    if fee_of(asin) is None:
        n_empty += 1
    else:
        n_fee += 1
print(dst, 'rows', len(asin_row), 'fba', n_fee, 'empty', n_empty, 'note_inserted', inserted)

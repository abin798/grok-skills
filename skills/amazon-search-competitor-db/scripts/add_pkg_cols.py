"""在竞品明细的 FBA配送费列（I）后插入「包装尺寸(cm)」「包装重量(g)」两列，后面各列右移两格。
纯 XML，嵌入图片不受影响。数据读自 /workspace/sorftime_cache.json。
用法: python3 add_pkg_cols.py 输入.xlsx 输出.xlsx
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

if header('I') != 'FBA配送费':
    sys.exit(f'skip: I={header("I")}')
if header('J') in ('包装尺寸(cm)', '包装尺寸'):
    sys.exit('already has 包装尺寸')

# ASIN (D) -> row
asin_row = {}
for m in re.finditer(r'<c r="D(\d+)"[^>]*>(.*?)</c>', s, re.S):
    r = m.group(1)
    if r == '1':
        continue
    t = re.search(r'<t[^>]*>(.*?)</t>', m.group(2))
    if t:
        asin_row[int(r)] = html.unescape(t.group(1))

# 1) J 及之后的列右移两格（单元格引用、超链接）
def sh(col):
    i = ci(col)
    return gl(i + 2) if i >= 10 else col

s = re.sub(r'(<c r=")([A-Z]+)(\d+")', lambda m: m.group(1) + sh(m.group(2)) + m.group(3), s)
s = re.sub(r'(<hyperlink ref=")([A-Z]+)(\d+")', lambda m: m.group(1) + sh(m.group(2)) + m.group(3), s)

# 2) 列宽：min/max>=10 的 +2，并在 I 后插入 J、K
def col(m):
    prefix, a, b = m.group(1), int(m.group(2)), int(m.group(3))
    if a >= 10:
        a, b = a + 2, b + 2
    return f'{prefix}min="{a}" max="{b}"/>'

s = re.sub(r'(<col [^>]*?)min="(\d+)" max="(\d+)"/>', col, s)
s = re.sub(
    r'(<col [^>]*min="9" max="9"/>)',
    r'\1<col width="18" customWidth="1" min="10" max="10"/>'
    r'<col width="10" customWidth="1" min="11" max="11"/>',
    s,
    count=1,
)

# 3) dimension / autoFilter 末列 +2
def bump_end(m):
    return f'{m.group(1)}{gl(ci(m.group(2))+2)}{m.group(3)}'

s = re.sub(r'(ref="A1:)([A-Z]+)(\d+")', bump_end, s)

def size_of(asin):
    d = cache.get(asin) or {}
    raw = d.get('package_size_cm') if isinstance(d, dict) else None
    if not raw:
        return None
    if isinstance(raw, list):
        parts = []
        for item in raw:
            if item is None or str(item).strip() == '':
                continue
            parts.extend([p.strip() for p in str(item).split(',') if p.strip()])
        if not parts:
            return None
        return ' × '.join(parts)
    text = str(raw).strip()
    return text or None

def weight_of(asin):
    d = cache.get(asin) or {}
    w = d.get('weight_g') if isinstance(d, dict) else None
    if w is None or w == '' or (isinstance(w, (int, float)) and w < 0):
        return None
    try:
        return int(float(w))
    except (TypeError, ValueError):
        return None

def insert_jk(m):
    row = m.group(0)
    r = re.search(r'<row r="(\d+)"', row).group(1)
    if re.search(rf'<c r="J{r}"', row):
        return row
    if r == '1':
        cells = (
            f'<c r="J1" s="1" t="inlineStr"><is><t>包装尺寸(cm)</t></is></c>'
            f'<c r="K1" s="1" t="inlineStr"><is><t>包装重量(g)</t></is></c>'
        )
    else:
        asin = asin_row.get(int(r))
        sz = size_of(asin)
        wt = weight_of(asin)
        if sz is None:
            j = f'<c r="J{r}" s="5" t="inlineStr"><is><t></t></is></c>'
        else:
            j = f'<c r="J{r}" s="5" t="inlineStr"><is><t>{html.escape(sz, quote=False)}</t></is></c>'
        if wt is None:
            k = f'<c r="K{r}" s="8" t="n"></c>'
        else:
            k = f'<c r="K{r}" s="8" t="n"><v>{wt}</v></c>'
        cells = j + k
    if re.search(rf'<c r="I{r}"[^>]*?(?:/>|>.*?</c>)', row, re.S):
        row = re.sub(rf'(<c r="I{r}"[^>]*?(?:/>|>.*?</c>))', r'\1' + cells, row, count=1, flags=re.S)
    else:
        row = re.sub(r'(<row r="\d+"[^>]*>)', r'\1' + cells, row, count=1)
    return row

s = re.sub(r'<row r="\d+"[^>]*>.*?</row>', insert_jk, s, flags=re.S)

wb = zi.read('xl/workbook.xml').decode()
wb = re.sub(
    r"('竞品明细'!\$A\$1:\$)([A-Z]+)(\$\d+)",
    lambda m: m.group(1) + gl(ci(m.group(2)) + 2) + m.group(3),
    wb,
)

s3 = zi.read('xl/worksheets/sheet3.xml').decode()
NOTES = [
    ('包装尺寸(cm)',
     '来自 Sorftime 商品详情接口的 package_size_cm，单位厘米，格式为长 × 宽 × 高；无数据留空。查询日期 2026-11-16。'),
    ('包装重量(g)',
     '来自 Sorftime 商品详情接口的 weight_g，单位克；无数据或无效值（如 -1）留空。查询日期 2026-11-16。'),
]

def esc(t):
    return html.escape(t, quote=False)

rows = re.findall(r'<row [^>]*>.*?</row>', s3, re.S)
new = []
inserted = False
for r in rows:
    a = re.search(r'<c r="A\d+"[^>]*>.*?<t[^>]*>(.*?)</t>', r, re.S)
    title = html.unescape(a.group(1)).strip() if a else ''
    new.append(r)
    if title == 'FBA配送费' and not inserted:
        for name, note in NOTES:
            new.append(
                '<row r="0">'
                f'<c r="A0" s="16" t="inlineStr"><is><t>{esc(name)}</t></is></c>'
                f'<c r="B0" s="5" t="inlineStr"><is><t>{esc(note)}</t></is></c>'
                '</row>'
            )
        inserted = True
if not inserted:
    sys.exit('说明页未找到「FBA配送费」行')

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

n_sz = n_wt = n_empty_sz = n_empty_wt = 0
for r, asin in asin_row.items():
    if size_of(asin) is None:
        n_empty_sz += 1
    else:
        n_sz += 1
    if weight_of(asin) is None:
        n_empty_wt += 1
    else:
        n_wt += 1
print(dst, 'rows', len(asin_row), 'size', n_sz, 'empty_size', n_empty_sz,
      'weight', n_wt, 'empty_weight', n_empty_wt, 'note_inserted', inserted)

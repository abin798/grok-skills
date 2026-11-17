"""按表头名称重排竞品明细列（纯 XML，保留嵌入图片/链接/列宽/样式）。
用法: python3 reorder_cols.py 输入.xlsx 输出.xlsx
"""
import zipfile, re, html, sys
from openpyxl.utils import column_index_from_string as ci, get_column_letter as gl
ORDER = ['自然排名','页码','页内位置','ASIN','主图','标题','品牌','价格','近一月销量','变体数量',
 '上架时间','FBA配送费','包装尺寸(cm)','包装重量(g)','评分','评论数','大类BSR排名','大类','小类BSR排名','小类名称',
 '父ASIN','库存状态','卖家','配送方式','主体所在国家/地区','公司地址','卖家精灵数据更新时间','卖家ID','店铺名称','公司主体',
 '卖家ASIN数量','卖家最佳自然排名','卖家主页']
src, dst = sys.argv[1], sys.argv[2]
zi = zipfile.ZipFile(src)
s = zi.read('xl/worksheets/sheet1.xml').decode()
hdr = {}
for m in re.finditer(r'<c r="([A-Z]+)1"[^>]*>.*?<t[^>]*>(.*?)</t>', s, re.S):
    hdr[html.unescape(m.group(2))] = ci(m.group(1))
missing = set(ORDER) ^ set(hdr)
if missing: sys.exit(f'header mismatch: {missing}')
newpos = {hdr[name]: i+1 for i, name in enumerate(ORDER)}
mp = {gl(o): gl(n) for o, n in newpos.items()}
s = re.sub(r'(<c r=")([A-Z]+)(\d+")', lambda m: m.group(1)+mp.get(m.group(2), m.group(2))+m.group(3), s)
s = re.sub(r'(<hyperlink ref=")([A-Z]+)(\d+")', lambda m: m.group(1)+mp.get(m.group(2), m.group(2))+m.group(3), s)
w = {int(m.group(2)): m.group(1) for m in re.finditer(r'<col width="([\d.]+)" customWidth="1" min="(\d+)" max="\2"/>', s)}
neww = {newpos[o]: w[o] for o in newpos if o in w}
s = re.sub(r'<col width="([\d.]+)" customWidth="1" min="(\d+)" max="\2"/>',
           lambda m: f'<col width="{neww.get(int(m.group(2)), m.group(1))}" customWidth="1" min="{m.group(2)}" max="{m.group(2)}"/>', s)
cellpat = re.compile(r'<c r="([A-Z]+)\d+"[^>]*?(?:/>|>.*?</c>)', re.S)
def fixrow(m):
    body = m.group(2)
    cells = sorted(((ci(c.group(1)), c.group(0)) for c in cellpat.finditer(body)), key=lambda x: x[0])
    assert cellpat.sub('', body).strip() == ''
    return m.group(1)+''.join(c for _, c in cells)+m.group(3)
s = re.sub(r'(<row [^>]*>)(.*?)(</row>)', fixrow, s, flags=re.S)
zo = zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED)
for it in zi.infolist():
    d = zi.read(it.filename)
    if it.filename == 'xl/worksheets/sheet1.xml': d = s.encode()
    zo.writestr(it, d)
zo.close()
print('ok', dst)

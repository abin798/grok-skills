"""删除竞品明细的 I(原价/参考价)、J(折扣) 两列，后面各列左移；删说明页「折扣」行。纯 XML，嵌入图片不受影响。
用法: python3 drop_cols_IJ.py 输入.xlsx 输出.xlsx"""
import zipfile,re,sys,html
from openpyxl.utils import column_index_from_string as ci, get_column_letter as gl
src,dst=sys.argv[1],sys.argv[2]
zi=zipfile.ZipFile(src); s=zi.read('xl/worksheets/sheet1.xml').decode()
h=lambda c:html.unescape(re.search(rf'<c r="{c}1"[^>]*>.*?<t[^>]*>(.*?)</t>',s,re.S).group(1))
if h('I')!='原价/参考价' or h('J')!='折扣': sys.exit(f'skip: I={h("I")} J={h("J")}')
s=re.sub(r'<c r="[IJ]\d+"[^>]*?(?:/>|>.*?</c>)','',s)
s=re.sub(r'<hyperlink ref="[IJ]\d+"[^>]*/>','',s)
def sh(col): i=ci(col); return gl(i-2) if i>10 else col
s=re.sub(r'(<c r=")([A-Z]+)(\d+")',lambda m:m.group(1)+sh(m.group(2))+m.group(3),s)
s=re.sub(r'(<hyperlink ref=")([A-Z]+)(\d+")',lambda m:m.group(1)+sh(m.group(2))+m.group(3),s)
def col(m):
    a,b=int(m.group(2)),int(m.group(3))
    if a in(9,10): return ''
    if a>10: a,b=a-2,b-2
    return f'{m.group(1)}min="{a}" max="{b}"/>'
s=re.sub(r'(<col [^>]*?)min="(\d+)" max="(\d+)"/>',col,s)
s=re.sub(r'(ref="A1:)AF(\d+")',r'\1AD\2',s)
assert ':AF' not in s and 'r="AE' not in s
wb=zi.read('xl/workbook.xml').decode().replace('$AF$','$AD$')
s3=zi.read('xl/worksheets/sheet3.xml').decode()
rows=re.findall(r'<row [^>]*>.*?</row>',s3,re.S); keep=[]
for r in rows:
    a=re.search(r'<c r="A\d+"[^>]*>.*?<t[^>]*>(.*?)</t>',r,re.S)
    if a and html.unescape(a.group(1)).strip()=='折扣': continue
    keep.append(r)
new=[]
for i,r in enumerate(keep,1):
    old=re.search(r'<row r="(\d+)"',r).group(1)
    r=re.sub(r'<row r="\d+"',f'<row r="{i}"',r,count=1)
    r=re.sub(rf'(<c r="[A-Z]+){old}"',rf'\g<1>{i}"',r); new.append(r)
body=re.search(r'<sheetData>.*?</sheetData>',s3,re.S).group(0)
s3=s3.replace(body,'<sheetData>'+''.join(new)+'</sheetData>')
s3=re.sub(r'<dimension ref="A1:([A-Z]+)\d+"/>',lambda m:f'<dimension ref="A1:{m.group(1)}{len(keep)}"/>',s3)
zo=zipfile.ZipFile(dst,'w',zipfile.ZIP_DEFLATED)
for it in zi.infolist():
    d=zi.read(it.filename)
    d={'xl/worksheets/sheet1.xml':s.encode(),'xl/workbook.xml':wb.encode(),'xl/worksheets/sheet3.xml':s3.encode()}.get(it.filename,d)
    zo.writestr(it,d)
zo.close()

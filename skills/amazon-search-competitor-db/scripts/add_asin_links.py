"""给竞品明细 D 列 ASIN 加商品页超链接（蓝色下划线），只改 zip 内 XML，不经 openpyxl，嵌入图片不受影响。
用法: python3 add_asin_links.py 输入.xlsx 输出.xlsx"""
import zipfile,re,sys,html
src,dst=sys.argv[1],sys.argv[2]
zi=zipfile.ZipFile(src)
wb=zi.read('xl/workbook.xml').decode()
# 竞品明细 is sheet1 by convention
S='xl/worksheets/sheet1.xml'; R='xl/worksheets/_rels/sheet1.xml.rels'
s=zi.read(S).decode(); rels=zi.read(R).decode(); st=zi.read('xl/styles.xml').decode()
if re.search(r'<hyperlink ref="D\d+"',s): sys.exit('already has D links')
# style: ensure hyperlink font
fonts=re.search(r'<fonts count="(\d+)"[^>]*>(.*?)</fonts>',st,re.S)
fl=re.findall(r'<font>.*?</font>|<font/>',fonts.group(2),re.S)
fid=next((i for i,x in enumerate(fl) if '0563C1' in x and '<u' in x),None)
if fid is None:
    fid=len(fl); newf=fonts.group(0).replace('</fonts>','<font><color rgb="FF0563C1"/><u val="single"/></font></fonts>')
    newf=re.sub(r'<fonts count="\d+"',f'<fonts count="{fid+1}"',newf,1); st=st.replace(fonts.group(0),newf)
cx=re.search(r'<cellXfs count="(\d+)"[^>]*>(.*?)</cellXfs>',st,re.S)
xl=re.findall(r'<xf [^>]*?/>|<xf [^>]*?>.*?</xf>',cx.group(2),re.S)
cache={}
def newstyle(old):
    if old in cache: return cache[old]
    x=xl[old]; x=re.sub(r'fontId="\d+"',f'fontId="{fid}"',x)
    if 'applyFont' not in x: x=x.replace('<xf ','<xf applyFont="1" ',1)
    xl.append(x); cache[old]=len(xl)-1; return cache[old]
ids=[int(i) for i in re.findall(r'Id="rId(\d+)"',rels)]; nxt=max(ids or [0])+1
links=[];newrels=[]
def rep(m):
    global nxt
    row,sid,asin=m.group(1),int(m.group(2)),m.group(3)
    if row=='1' or not re.fullmatch(r'[A-Z0-9]{10}',asin): return m.group(0)
    ns=newstyle(sid); rid=f'rId{nxt}'; nxt+=1
    links.append(f'<hyperlink ref="D{row}" r:id="{rid}"/>')
    newrels.append(f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://www.amazon.com/dp/{asin}" TargetMode="External"/>')
    return m.group(0).replace(f's="{sid}"',f's="{ns}"',1)
s=re.sub(r'<c r="D(\d+)" s="(\d+)" t="inlineStr"><is><t>([^<]+)</t></is></c>',rep,s)
s=s.replace('<hyperlinks>','<hyperlinks>'+''.join(links),1)
rels=rels.replace('</Relationships>',''.join(newrels)+'</Relationships>')
newcx=f'<cellXfs count="{len(xl)}">'+''.join(xl)+'</cellXfs>'
st=st.replace(cx.group(0),newcx)
zo=zipfile.ZipFile(dst,'w',zipfile.ZIP_DEFLATED)
for it in zi.infolist():
    d=zi.read(it.filename)
    if it.filename==S: d=s.encode()
    elif it.filename==R: d=rels.encode()
    elif it.filename=='xl/styles.xml': d=st.encode()
    zo.writestr(it,d)
zo.close(); print(dst,'D links',len(links))

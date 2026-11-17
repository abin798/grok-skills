"""竞品明细格式紧凑化：整数加千分位、卖家精灵更新时间只留日期、行高列宽字号收紧。纯 XML。
用法: python3 compact_format.py 输入.xlsx 输出.xlsx"""
import zipfile,re,sys,math
src,dst=sys.argv[1],sys.argv[2]
zi=zipfile.ZipFile(src)
st=zi.read('xl/styles.xml').decode(); s=zi.read('xl/worksheets/sheet1.xml').decode()
# 1 千分位：cellXfs 中 numFmtId="1"(0) 改为 3(#,##0)
cx=re.search(r'<cellXfs[^>]*>.*?</cellXfs>',st,re.S).group(0)
cx2=cx.replace('numFmtId="1" ','numFmtId="3" ')
# 3 日期时间格式 168/164 -> 166 (yyyy-mm-dd)
cx2=cx2.replace('numFmtId="168" ','numFmtId="166" ').replace('numFmtId="164" ','numFmtId="166" ')
st=st.replace(cx,cx2)
# 字号 10
def fz(m):
    f=m.group(0)
    f=re.sub(r'<sz val="[\d.]+"/>','',f)
    return f.replace('</font>','<sz val="10"/></font>')
st=re.sub(r'<font>.*?</font>|<font/>',lambda m:fz(m) if m.group(0)!='<font/>' else '<font><sz val="10"/></font>',re.search(r'<fonts.*?</fonts>',st,re.S).group(0)) if False else st
fonts=re.search(r'<fonts[^>]*>.*?</fonts>',st,re.S).group(0)
st=st.replace(fonts,re.sub(r'<font>.*?</font>',fz,fonts,flags=re.S))
# V 列值去掉时间部分
hdr=dict((html_c,t) for html_c,t in re.findall(r'<c r="([A-Z]+)1"[^>]*>.*?<t[^>]*>(.*?)</t>',s,re.S))
import html as H
vcol=[c for c,t in hdr.items() if H.unescape(t)=='卖家精灵数据更新时间']
for vc in vcol:
    s=re.sub(rf'(<c r="{vc}\d+"[^>]*><v>)([\d.]+)(</v>)',lambda m:m.group(1)+str(math.floor(float(m.group(2))))+m.group(3),s)
# 4 行高、列宽
s=re.sub(r'(<row r="1" ht=")[\d.]+"',r'\g<1>40"',s)
s=re.sub(r'(<row r="(?!1")\d+" ht=")[\d.]+"',r'\g<1>86"',s)
W=[6,5,6,13.5,15.7,40,12,8,8,12,5,7,8,14,7,18,6,11,9,14,7,12,15,14,22,28,9,6,7,24]
def col(m):
    i=int(m.group(2)); return m.group(1)+f'width="{W[i-1]}"'+m.group(3) if i<=len(W) else m.group(0)
s=re.sub(r'(<col )width="[\d.]+"( customWidth="1" min="(\d+)")',lambda m:m.group(1)+f'width="{W[int(m.group(3))-1]}"'+m.group(2),s)
zo=zipfile.ZipFile(dst,'w',zipfile.ZIP_DEFLATED)
for it in zi.infolist():
    d=zi.read(it.filename)
    d={'xl/styles.xml':st.encode(),'xl/worksheets/sheet1.xml':s.encode()}.get(it.filename,d)
    zo.writestr(it,d)
zo.close()

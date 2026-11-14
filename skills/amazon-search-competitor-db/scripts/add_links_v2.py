import zipfile,re,sys
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from xml.sax.saxutils import escape
SRC,STAGE,DST=sys.argv[1:4]
ws=load_workbook(STAGE,read_only=True)['竞品明细']
rows=list(ws.iter_rows(min_row=1,values_only=True)); hdr=list(rows[0])
iA=hdr.index('ASIN'); iP=hdr.index('主图'); iS=hdr.index('卖家主页')
LP=get_column_letter(iP+1); LS=get_column_letter(iS+1)
assert LP=='E', LP
links=[]  # (ref,url)
for n,r in enumerate(rows[1:],2):
    a=r[iA]; assert a and re.fullmatch(r'[A-Z0-9]{10}',a),(n,a)
    links.append((f'{LP}{n}',f'https://www.amazon.com/dp/{a}'))
for n,r in enumerate(rows[1:],2):
    u=r[iS]
    if u:
        assert u.startswith('https://www.amazon.com/sp?seller='),(n,u)
        links.append((f'{LS}{n}',u))
z=zipfile.ZipFile(SRC); files={k:z.read(k) for k in z.namelist()}; order=z.namelist()
s=files['xl/worksheets/sheet1.xml'].decode()
assert '<hyperlinks>' not in s
if 'xmlns:r=' not in s[:500]:
    s=s.replace('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"','<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"',1)
assert 'xmlns:r=' in s[:500]
hl='<hyperlinks>'+''.join(f'<hyperlink ref="{ref}" r:id="rId{i+1}"/>' for i,(ref,u) in enumerate(links))+'</hyperlinks>'
pos=[s.find(t) for t in ('<printOptions','<pageMargins','<pageSetup','<headerFooter','<drawing','<legacyDrawing','<tableParts','<extLst','</worksheet>') if s.find(t)>=0]
p=min(pos); s=s[:p]+hl+s[p:]
files['xl/worksheets/sheet1.xml']=s.encode()
rel='xl/worksheets/_rels/sheet1.xml.rels'
assert rel not in files
files[rel]=('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'+''.join(f'<Relationship Id="rId{i+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="{escape(u)}" TargetMode="External"/>' for i,(ref,u) in enumerate(links))+'</Relationships>').encode()
with zipfile.ZipFile(DST,'w',zipfile.ZIP_DEFLATED) as out:
    for k in order+[k for k in files if k not in order]: out.writestr(k,files[k])
print('links',len(links))

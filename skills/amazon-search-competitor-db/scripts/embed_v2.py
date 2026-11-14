import zipfile, re, os
SRC='stage_entity.xlsx'
DST='stage_entity_embed.xlsx'
COL='E'
zin=zipfile.ZipFile(SRC)
files={n:zin.read(n) for n in zin.namelist()}
order=zin.namelist()
rows=list(range(2,145))
# sheet1
s=files['xl/worksheets/sheet1.xml'].decode()
vm={}
def rep(m):
    r=int(m.group(1)); st=m.group(2)
    k=rows.index(r)+1; vm[r]=k
    return f'<c r="{COL}{r}"{st} t="e" vm="{k}"><v>#VALUE!</v></c>'
s,n=re.subn(r'<c r="'+COL+r'(\d+)"( s="\d+")? t="inlineStr"><is><t>[^<]*</t></is></c>',lambda m: rep(m) if int(m.group(1))>=2 else m.group(0),s)
assert len(vm)==143, len(vm)
# strip any hyperlinks/drawing (stage should have none)
assert '<hyperlink' not in s
s=re.sub(r'<drawing [^>]*/>','',s)
rp='xl/worksheets/_rels/sheet1.xml.rels'
if rp in files:
    rels=re.sub(r'<Relationship [^>]*drawing[^>]*/>','',files[rp].decode())
    if '<Relationship ' in rels: files[rp]=rels.encode()
    else: del files[rp]
files['xl/worksheets/sheet1.xml']=s.encode()
drop=set()
# drop drawing + old media
for k in list(files):
    if k.startswith('xl/drawings/') or k.startswith('xl/media/'): del files[k]
ct=files['[Content_Types].xml'].decode()
ct=ct.replace('<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>','')
add=[('/xl/metadata.xml','application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml'),
('/xl/richData/rdRichValueTypes.xml','application/vnd.ms-excel.rdrichvaluetypes+xml'),
('/xl/richData/rdrichvalue.xml','application/vnd.ms-excel.rdrichvalue+xml'),
('/xl/richData/rdrichvaluestructure.xml','application/vnd.ms-excel.rdrichvaluestructure+xml'),
('/xl/richData/richValueRel.xml','application/vnd.ms-excel.richvaluerel+xml')]
ct=ct.replace('</Types>',''.join(f'<Override PartName="{p}" ContentType="{c}"/>' for p,c in add)+'</Types>')
files['[Content_Types].xml']=ct.encode()
wr=files['xl/_rels/workbook.xml.rels'].decode()
ids=[int(x) for x in re.findall(r'Id="rId(\d+)"',wr)]; b=max(ids)
nr=[('http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata','metadata.xml'),
('http://schemas.microsoft.com/office/2022/10/relationships/richValueRel','richData/richValueRel.xml'),
('http://schemas.microsoft.com/office/2017/06/relationships/rdRichValue','richData/rdrichvalue.xml'),
('http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueStructure','richData/rdrichvaluestructure.xml'),
('http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueTypes','richData/rdRichValueTypes.xml')]
wr=wr.replace('</Relationships>',''.join(f'<Relationship Id="rId{b+i+1}" Type="{t}" Target="{g}"/>' for i,(t,g) in enumerate(nr))+'</Relationships>')
files['xl/_rels/workbook.xml.rels']=wr.encode()
H='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
N=143
for i,r in enumerate(rows):
    files[f'xl/media/image{i+1}.jpeg']=open(f'napkin_imgs/{r}.jpg','rb').read()
files['xl/richData/richValueRel.xml']=(H+'<richValueRels xmlns="http://schemas.microsoft.com/office/spreadsheetml/2022/richvaluerel" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'+''.join(f'<rel r:id="rId{i+1}"/>' for i in range(N))+'</richValueRels>').encode()
files['xl/richData/_rels/richValueRel.xml.rels']=(H+'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'+''.join(f'<Relationship Id="rId{i+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image{i+1}.jpeg"/>' for i in range(N))+'</Relationships>').encode()
files['xl/richData/rdrichvalue.xml']=(H+f'<rvData xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata" count="{N}">'+''.join(f'<rv s="0"><v>{i}</v><v>5</v></rv>' for i in range(N))+'</rvData>').encode()
ref='ref/xl/'
for f in ['richData/rdrichvaluestructure.xml','richData/rdRichValueTypes.xml']:
    files['xl/'+f]=open(ref+f,'rb').read()
files['xl/metadata.xml']=(H+'<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:xlrd="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata"><metadataTypes count="1"><metadataType name="XLRICHVALUE" minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" coerce="1"/></metadataTypes>'
 +f'<futureMetadata name="XLRICHVALUE" count="{N}">'+''.join(f'<bk><extLst><ext uri="{{3e2802c4-a4d2-4d8b-9148-e3be6c30e623}}"><xlrd:rvb i="{i}"/></ext></extLst></bk>' for i in range(N))+'</futureMetadata>'
 +f'<valueMetadata count="{N}">'+''.join(f'<bk><rc t="1" v="{i}"/></bk>' for i in range(N))+'</valueMetadata></metadata>').encode()
neworder=[n for n in order if n in files]+[n for n in files if n not in order]
with zipfile.ZipFile(DST,'w',zipfile.ZIP_DEFLATED) as z:
    for n in neworder: z.writestr(n,files[n])
print('ok',len(drop))

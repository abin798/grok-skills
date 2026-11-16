"""Finish a newly authored XLSX with native Excel in-cell images and hyperlinks.

Uses ZIP/XML only; never opens and re-saves rich images through openpyxl.
"""
from __future__ import annotations
import argparse
import copy
import json
import os
from pathlib import Path, PurePosixPath
import posixpath
import re
import tempfile
import xml.etree.ElementTree as ET
import zipfile

MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PKG = 'http://schemas.openxmlformats.org/package/2006/relationships'
CT = 'http://schemas.openxmlformats.org/package/2006/content-types'
RD = 'http://schemas.microsoft.com/office/spreadsheetml/2017/richdata'
RD2 = 'http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2'
RREL = 'http://schemas.microsoft.com/office/spreadsheetml/2022/richvaluerel'
Q = lambda ns, tag: f'{{{ns}}}{tag}'

def xml_bytes(root, original=None):
    ns=root.tag.split('}')[0][1:] if root.tag.startswith('{') else ''
    if ns: ET.register_namespace('',ns)
    ET.register_namespace('r',REL)
    ET.register_namespace('xlrd',RD)
    result=ET.tostring(root,encoding='utf-8',xml_declaration=True)
    # Preserve namespace declarations referenced only by mc:Ignorable values.
    if original:
        first=re.search(rb'<(?!\?)[^>]+>',original)
        if first:
            for decl in re.findall(rb'xmlns(?::[\w.-]+)?=["\'][^"\']+["\']',first.group()):
                key=decl.split(b'=')[0]
                start=re.search(rb'<(?!\?)[^>]+>',result)
                if start and key+b'=' not in start.group():
                    p=start.end()-1
                    result=result[:p]+b' '+decl+result[p:]
    return result

def resolve_part(owner, target):
    return posixpath.normpath(posixpath.join(posixpath.dirname(owner),target.lstrip('/'))) if not target.startswith('/') else target.lstrip('/')

def sheet_part(files, name):
    wb=ET.fromstring(files['xl/workbook.xml'])
    rels=ET.fromstring(files['xl/_rels/workbook.xml.rels'])
    lookup={r.get('Id'):resolve_part('xl/workbook.xml',r.get('Target','')) for r in rels}
    for s in wb.find(Q(MAIN,'sheets')):
        if s.get('name')==name: return lookup[s.get(Q(REL,'id'))]
    raise ValueError(f'Missing sheet: {name}')

def rel_path(part):
    p=PurePosixPath(part)
    return str(p.parent/'_rels'/(p.name+'.rels'))

def next_id(root):
    ids={r.get('Id') for r in root}
    n=1
    while f'rId{n}' in ids: n+=1
    return f'rId{n}'

def add_relationship(root, typ, target, external=False):
    rid=next_id(root)
    attr={'Id':rid,'Type':typ,'Target':target}
    if external: attr['TargetMode']='External'
    ET.SubElement(root,Q(PKG,'Relationship'),attr)
    return rid

def finish_workbook(src, dst, data, schema, allow_missing_images=False):
    with zipfile.ZipFile(src) as z:
        files={name:z.read(name) for name in z.namelist()}
    if any(n.startswith('xl/richData/') for n in files) or 'xl/metadata.xml' in files:
        raise ValueError('Input already contains rich-data/metadata. Rebuild from JSON to avoid overwriting existing images or metadata.')
    part=sheet_part(files,'竞品明细')
    original=files[part]
    sheet=ET.fromstring(original)
    sheet_data=sheet.find(Q(MAIN,'sheetData'))
    rows_by_num={int(r.get('r')):r for r in sheet_data}
    rpath=rel_path(part)
    rels=ET.fromstring(files[rpath]) if rpath in files else ET.Element(Q(PKG,'Relationships'))
    hlinks=sheet.find(Q(MAIN,'hyperlinks'))
    if hlinks is None:
        hlinks=ET.Element(Q(MAIN,'hyperlinks'))
        # ECMA worksheet order: hyperlinks follow dataValidation and precede printOptions.
        after={'printOptions','pageMargins','pageSetup','headerFooter','rowBreaks','colBreaks','customProperties','cellWatches','ignoredErrors','smartTags','drawing','legacyDrawing','legacyDrawingHF','picture','oleObjects','controls','webPublishItems','tableParts','extLst'}
        idx=next((i for i,e in enumerate(sheet) if e.tag.split('}')[-1] in after),len(sheet))
        sheet.insert(idx,hlinks)
    existing_links={h.get('ref'):h for h in hlinks}
    headers=schema['headers']
    image_col=column_letter(headers.index('主图')+1)
    seller_col=column_letter(headers.index('卖家主页')+1)
    images=[]
    for i,record in enumerate(data.get('records',[]),2):
        row=rows_by_num.get(i)
        if row is None: raise ValueError(f'Missing workbook row {i}')
        ref=f'{image_col}{i}'
        cell=next((c for c in row if c.get('r')==ref),None)
        if cell is None:
            cell=ET.SubElement(row,Q(MAIN,'c'),{'r':ref})
            ordered=sorted(list(row),key=lambda c:column_number(re.match(r'[A-Z]+',c.get('r')).group()))
            row[:]=ordered
        image=record.get('_image')
        if image:
            raw=Path(image).read_bytes()
            if raw.startswith(b'\x89PNG\r\n\x1a\n'): ext='png'; mime='image/png'
            elif raw.startswith(b'\xff\xd8\xff'): ext='jpeg'; mime='image/jpeg'
            else: raise ValueError(f'Only PNG/JPEG images supported: {image}')
            images.append((ref,raw,ext,mime))
            for c in list(cell): cell.remove(c)
            cell.set('t','e'); cell.set('vm',str(len(images)))
            ET.SubElement(cell,Q(MAIN,'v')).text='#VALUE!'
        elif not allow_missing_images:
            raise ValueError(f'{record["ASIN"]}: _image is required')
        else:
            for c in list(cell): cell.remove(c)
            cell.set('t','inlineStr')
            ET.SubElement(ET.SubElement(cell,Q(MAIN,'is')),Q(MAIN,'t')).text='图片缺失'
        for href,target in [(ref,f'https://www.amazon.com/dp/{record["ASIN"]}'),(f'{seller_col}{i}',record.get('卖家主页'))]:
            if not target: continue
            if not str(target).startswith(('https://','http://')): raise ValueError(f'Invalid hyperlink: {target}')
            rid=add_relationship(rels,REL+'/hyperlink',str(target),True)
            h=existing_links.get(href)
            if h is None: h=ET.SubElement(hlinks,Q(MAIN,'hyperlink'),{'ref':href})
            h.set(Q(REL,'id'),rid)
    if not len(hlinks): sheet.remove(hlinks)
    # Replace the worksheet autoFilter and pane explicitly: both work with zero records.
    auto=sheet.find(Q(MAIN,'autoFilter'))
    if auto is None:
        auto=ET.Element(Q(MAIN,'autoFilter'))
        data_idx=list(sheet).index(sheet_data)
        sheet.insert(data_idx+1,auto)
    auto.set('ref',f'A1:AI{max(1,len(data.get("records",[]))+1)}')
    views=sheet.find(Q(MAIN,'sheetViews'))
    if views is None:
        views=ET.Element(Q(MAIN,'sheetViews')); sheet.insert(0,views)
    view=views.find(Q(MAIN,'sheetView'))
    if view is None: view=ET.SubElement(views,Q(MAIN,'sheetView'),{'workbookViewId':'0'})
    for e in list(view):
        if e.tag in {Q(MAIN,'pane'),Q(MAIN,'selection')}: view.remove(e)
    view.insert(0,ET.Element(Q(MAIN,'pane'),{'xSplit':'5','ySplit':'1','topLeftCell':'F2','activePane':'bottomRight','state':'frozen'}))
    ET.SubElement(view,Q(MAIN,'selection'),{'pane':'bottomRight','activeCell':'F2','sqref':'F2'})
    files[part]=xml_bytes(sheet,original)
    if len(rels): files[rpath]=xml_bytes(rels)
    if images:
        n=len(images)
        ct=ET.fromstring(files['[Content_Types].xml'])
        def content_override(partname,ctype):
            existing=next((e for e in ct if e.get('PartName')==partname),None)
            if existing is None: ET.SubElement(ct,Q(CT,'Override'),{'PartName':partname,'ContentType':ctype})
        def content_default(ext,mime):
            existing=next((e for e in ct if e.get('Extension')==ext),None)
            if existing is None: ET.SubElement(ct,Q(CT,'Default'),{'Extension':ext,'ContentType':mime})
        extra=[('/xl/metadata.xml','application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml'),('/xl/richData/rdRichValueTypes.xml','application/vnd.ms-excel.rdrichvaluetypes+xml'),('/xl/richData/rdrichvalue.xml','application/vnd.ms-excel.rdrichvalue+xml'),('/xl/richData/rdrichvaluestructure.xml','application/vnd.ms-excel.rdrichvaluestructure+xml'),('/xl/richData/richValueRel.xml','application/vnd.ms-excel.richvaluerel+xml')]
        for p,t in extra: content_override(p,t)
        wr=ET.fromstring(files['xl/_rels/workbook.xml.rels'])
        for typ,target in [(REL+'/sheetMetadata','metadata.xml'),('http://schemas.microsoft.com/office/2022/10/relationships/richValueRel','richData/richValueRel.xml'),('http://schemas.microsoft.com/office/2017/06/relationships/rdRichValue','richData/rdrichvalue.xml'),('http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueStructure','richData/rdrichvaluestructure.xml'),('http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueTypes','richData/rdRichValueTypes.xml')]: add_relationship(wr,typ,target)
        files['xl/_rels/workbook.xml.rels']=xml_bytes(wr)
        rr=ET.Element(Q(RREL,'richValueRels'))
        rr_rels=ET.Element(Q(PKG,'Relationships'))
        rv=ET.Element(Q(RD,'rvData'),{'count':str(n)})
        for index,(ref,raw,ext,mime) in enumerate(images):
            media=f'xl/media/competitor_{index+1}.{ext}'
            if media in files: raise ValueError(f'Unexpected image filename collision: {media}')
            files[media]=raw; content_default(ext,mime)
            rid=add_relationship(rr_rels,REL+'/image',f'../media/competitor_{index+1}.{ext}')
            ET.SubElement(rr,Q(RREL,'rel'),{Q(REL,'id'):rid})
            record=ET.SubElement(rv,Q(RD,'rv'),{'s':'0'})
            ET.SubElement(record,Q(RD,'v')).text=str(index)
            ET.SubElement(record,Q(RD,'v')).text='5'
        files['xl/richData/richValueRel.xml']=xml_bytes(rr)
        files['xl/richData/_rels/richValueRel.xml.rels']=xml_bytes(rr_rels)
        files['xl/richData/rdrichvalue.xml']=xml_bytes(rv)
        structure=ET.Element(Q(RD,'rvStructures'),{'count':'1'})
        st=ET.SubElement(structure,Q(RD,'s'),{'t':'_localImage'})
        ET.SubElement(st,Q(RD,'k'),{'n':'_rvRel:LocalImageIdentifier','t':'i'})
        ET.SubElement(st,Q(RD,'k'),{'n':'CalcOrigin','t':'i'})
        files['xl/richData/rdrichvaluestructure.xml']=xml_bytes(structure)
        types=ET.Element(Q(RD2,'rvTypesInfo'))
        flags=ET.SubElement(ET.SubElement(types,Q(RD2,'global')),Q(RD2,'keyFlags'))
        for key in ['_Self','_DisplayString','_Flags','_Format','_SubLabel','_Attribution','_Icon','_Display','_CanonicalPropertyNames','_ClassificationId']:
            k=ET.SubElement(flags,Q(RD2,'key'),{'name':key})
            if key=='_Self': ET.SubElement(k,Q(RD2,'flag'),{'name':'ExcludeFromFile','value':'1'})
            ET.SubElement(k,Q(RD2,'flag'),{'name':'ExcludeFromCalcComparison','value':'1'})
        files['xl/richData/rdRichValueTypes.xml']=xml_bytes(types)
        metadata=ET.Element(Q(MAIN,'metadata'))
        mt=ET.SubElement(metadata,Q(MAIN,'metadataTypes'),{'count':'1'})
        ma={'name':'XLRICHVALUE','minSupportedVersion':'120000',**{k:'1' for k in ['copy','pasteAll','pasteValues','merge','splitFirst','rowColShift','clearFormats','clearComments','assign','coerce']}}
        ET.SubElement(mt,Q(MAIN,'metadataType'),ma)
        future=ET.SubElement(metadata,Q(MAIN,'futureMetadata'),{'name':'XLRICHVALUE','count':str(n)})
        value=ET.SubElement(metadata,Q(MAIN,'valueMetadata'),{'count':str(n)})
        for i in range(n):
            ext=ET.SubElement(ET.SubElement(ET.SubElement(future,Q(MAIN,'bk')),Q(MAIN,'extLst')),Q(MAIN,'ext'),{'uri':'{3e2802c4-a4d2-4d8b-9148-e3be6c30e623}'})
            ET.SubElement(ext,Q(RD,'rvb'),{'i':str(i)})
            ET.SubElement(ET.SubElement(value,Q(MAIN,'bk')),Q(MAIN,'rc'),{'t':'1','v':str(i)})
        files['xl/metadata.xml']=xml_bytes(metadata)
        files['[Content_Types].xml']=xml_bytes(ct)
    dst=Path(dst); dst.parent.mkdir(parents=True,exist_ok=True)
    fd,temp=tempfile.mkstemp(suffix='.xlsx',prefix='.competitor-',dir=dst.parent)
    os.close(fd)
    try:
        with zipfile.ZipFile(temp,'w',zipfile.ZIP_DEFLATED) as z:
            for name,raw in files.items(): z.writestr(name,raw)
        os.replace(temp,dst)
    finally:
        if os.path.exists(temp): os.unlink(temp)
    return {'records':len(data.get('records',[])),'embedded_images':len(images),'hyperlinks':len(hlinks)}

def column_letter(number):
    out=''
    while number: number,rem=divmod(number-1,26); out=chr(65+rem)+out
    return out

def column_number(text):
    n=0
    for c in text: n=n*26+ord(c)-64
    return n

#!/usr/bin/env python3
"""Add invisible linked rectangle hot spots over existing native E-cell images.

The image data and richData remain unchanged. Hot spots are DrawingML shapes,
not replacement floating pictures. Native WPS click/sort/filter testing remains
necessary after generation.
"""
from __future__ import annotations
import argparse
import hashlib
import json
import os
from pathlib import Path
import posixpath
import re
import tempfile
import xml.etree.ElementTree as ET
import zipfile
from patch_link_styles import S,R,P,Q,xml_dump,resolve,find_sheet,text_value

XDR='http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'
A='http://schemas.openxmlformats.org/drawingml/2006/main'
CT='http://schemas.openxmlformats.org/package/2006/content-types'
PREFIX='CodexASINHotspot_'

def rels_path(part):
    return posixpath.dirname(part)+'/_rels/'+posixpath.basename(part)+'.rels'

def next_rid(root):
    ids={x.get('Id') for x in root}
    n=1
    while f'rId{n}' in ids: n+=1
    return f'rId{n}'

def drawing_dump(root,original=None):
    ET.register_namespace('xdr',XDR)
    ET.register_namespace('a',A)
    ET.register_namespace('r',R)
    result=ET.tostring(root,encoding='utf-8',xml_declaration=True)
    if original:
        first=re.search(rb'<(?!\?)[^>]+>',original)
        if first:
            for d in re.findall(rb'xmlns(?::[\w.-]+)?=["\'][^"\']+["\']',first.group()):
                start=re.search(rb'<(?!\?)[^>]+>',result)
                if start and d.split(b'=')[0]+b'=' not in start.group():
                    i=start.end()-1
                    result=result[:i]+b' '+d+result[i:]
    return result

def add_image_link_overlays(source,destination,sheet_name='竞品明细'):
    source,destination=Path(source),Path(destination)
    if source.resolve()==destination.resolve(): raise ValueError('Source and destination must differ')
    with zipfile.ZipFile(source) as archive:
        infos=archive.infolist()
        files={info.filename:archive.read(info.filename) for info in infos}
    protected={n:hashlib.sha256(raw).hexdigest() for n,raw in files.items() if n.startswith(('xl/media/','xl/richData/')) or n=='xl/metadata.xml'}
    sheet_part=find_sheet(files,sheet_name)
    sheet_original=files[sheet_part]
    sheet=ET.fromstring(sheet_original)
    cells={c.get('r'):c for row in sheet.find(Q(S,'sheetData')) for c in row}
    strings=[]
    if 'xl/sharedStrings.xml' in files:
        strings=[''.join(t.text or '' for t in x.iter(Q(S,'t'))) for x in ET.fromstring(files['xl/sharedStrings.xml'])]
    records=[]
    for ref,cell in cells.items():
        if not re.fullmatch(r'E[1-9]\d*',ref) or not cell.get('vm'): continue
        row=int(ref[1:])
        asin=text_value(cells[f'D{row}'],strings)
        if not re.fullmatch(r'[A-Z0-9]{10}',str(asin)): raise ValueError(f'Bad ASIN at D{row}: {asin}')
        records.append((row,asin))
    records.sort()
    if not records: raise ValueError('No native image cells found')
    worksheet_rels_name=rels_path(sheet_part)
    worksheet_rels=ET.fromstring(files[worksheet_rels_name]) if worksheet_rels_name in files else ET.Element(Q(P,'Relationships'))
    drawing_tag=sheet.find(Q(S,'drawing'))
    if drawing_tag is not None:
        rid=drawing_tag.get(Q(R,'id'))
        relationship=next((x for x in worksheet_rels if x.get('Id')==rid),None)
        if relationship is None: raise ValueError('Existing worksheet drawing relationship is missing')
        drawing_name=resolve(sheet_part,relationship.get('Target'))
        drawing_original=files[drawing_name]
        drawing=ET.fromstring(drawing_original)
    else:
        n=1
        while f'xl/drawings/drawing{n}.xml' in files: n+=1
        drawing_name=f'xl/drawings/drawing{n}.xml'
        drawing_original=None
        drawing=ET.Element(Q(XDR,'wsDr'))
        rid=next_rid(worksheet_rels)
        target=posixpath.relpath(drawing_name,posixpath.dirname(sheet_part))
        ET.SubElement(worksheet_rels,Q(P,'Relationship'),{'Id':rid,'Type':R+'/drawing','Target':target})
        drawing_tag=ET.Element(Q(S,'drawing'),{Q(R,'id'):rid})
        after={'legacyDrawing','legacyDrawingHF','picture','oleObjects','controls','webPublishItems','tableParts','extLst'}
        index=next((i for i,child in enumerate(sheet) if child.tag.split('}')[-1] in after),len(sheet))
        sheet.insert(index,drawing_tag)
    drawing_rels_name=rels_path(drawing_name)
    drawing_rels=ET.fromstring(files[drawing_rels_name]) if drawing_rels_name in files else ET.Element(Q(P,'Relationships'))
    if any((node.get('name') or '').startswith(PREFIX) for node in drawing.iter(Q(XDR,'cNvPr'))):
        raise ValueError('Overlay hot spots already exist; rebuild from the workbook without overlays')
    max_id=max([int(x.get('id','0')) for x in drawing.iter(Q(XDR,'cNvPr'))]+[0])
    original_anchors=len(drawing)
    # Shape xfrm repeats the full-cell anchor in EMUs for consumers that use it.
    # The authoritative sizing behavior is editAs=twoCell + next-cell boundaries.
    format_pr=sheet.find(Q(S,'sheetFormatPr'))
    default_height=float(format_pr.get('defaultRowHeight','15')) if format_pr is not None else 15.0
    row_heights={int(row.get('r')):float(row.get('ht',default_height)) for row in sheet.find(Q(S,'sheetData'))}
    col_widths={i:8.43 for i in range(5)}
    columns=sheet.find(Q(S,'cols'))
    if columns is not None:
        for column in columns:
            for i in range(max(0,int(column.get('min'))-1),min(5,int(column.get('max')))):
                col_widths[i]=float(column.get('width','8.43'))
    def pixels(width): return int(width*12+0.5) if width<1 else int(width*7+0.5)+5
    left=sum(pixels(col_widths[c]) for c in range(4))*9525
    width=pixels(col_widths[4])*9525
    expected={}
    for number,(row,asin) in enumerate(records,1):
        url='https://www.amazon.com/dp/'+asin
        anchor=ET.SubElement(drawing,Q(XDR,'twoCellAnchor'),{'editAs':'twoCell'})
        for tag,col_index,row_index in [('from',4,row-1),('to',5,row)]:
            marker=ET.SubElement(anchor,Q(XDR,tag))
            for name,value in [('col',col_index),('colOff',0),('row',row_index),('rowOff',0)]: ET.SubElement(marker,Q(XDR,name)).text=str(value)
        shape=ET.SubElement(anchor,Q(XDR,'sp'),{'macro':'','textlink':''})
        nonvisual=ET.SubElement(shape,Q(XDR,'nvSpPr'))
        nv=ET.SubElement(nonvisual,Q(XDR,'cNvPr'),{'id':str(max_id+number),'name':PREFIX+f'E{row}_{asin}','descr':f'透明商品链接热区：{asin}；底图仍为单元格内图片'})
        link_id=next_rid(drawing_rels)
        ET.SubElement(drawing_rels,Q(P,'Relationship'),{'Id':link_id,'Type':R+'/hyperlink','Target':url,'TargetMode':'External'})
        ET.SubElement(nv,Q(A,'hlinkClick'),{Q(R,'id'):link_id,'tooltip':f'打开亚马逊商品 {asin}'})
        ET.SubElement(nonvisual,Q(XDR,'cNvSpPr'))
        props=ET.SubElement(shape,Q(XDR,'spPr'))
        xfrm=ET.SubElement(props,Q(A,'xfrm'))
        top=sum(row_heights.get(r,default_height) for r in range(1,row))*12700
        height=row_heights.get(row,default_height)*12700
        ET.SubElement(xfrm,Q(A,'off'),{'x':str(left),'y':str(round(top))})
        ET.SubElement(xfrm,Q(A,'ext'),{'cx':str(width),'cy':str(round(height))})
        ET.SubElement(ET.SubElement(props,Q(A,'prstGeom'),{'prst':'rect'}),Q(A,'avLst'))
        fill=ET.SubElement(props,Q(A,'solidFill'))
        ET.SubElement(ET.SubElement(fill,Q(A,'srgbClr'),{'val':'FFFFFF'}),Q(A,'alpha'),{'val':'0'})
        ET.SubElement(ET.SubElement(props,Q(A,'ln')),Q(A,'noFill'))
        ET.SubElement(anchor,Q(XDR,'clientData'))
        expected[f'E{row}']=url
    content_types=ET.fromstring(files['[Content_Types].xml'])
    if not any(x.get('PartName')=='/'+drawing_name for x in content_types):
        ET.SubElement(content_types,Q(CT,'Override'),{'PartName':'/'+drawing_name,'ContentType':'application/vnd.openxmlformats-officedocument.drawing+xml'})
    files[drawing_name]=drawing_dump(drawing,drawing_original)
    files[drawing_rels_name]=xml_dump(drawing_rels,files.get(drawing_rels_name))
    files[worksheet_rels_name]=xml_dump(worksheet_rels,files.get(worksheet_rels_name))
    files[sheet_part]=xml_dump(sheet,sheet_original)
    files['[Content_Types].xml']=xml_dump(content_types,files['[Content_Types].xml'])
    for name,digest in protected.items():
        if hashlib.sha256(files[name]).hexdigest()!=digest: raise ValueError(f'Native image part changed: {name}')
    # Independent structural check: each transparent shape target matches its D ASIN.
    saved=ET.fromstring(files[drawing_name])
    relmap={x.get('Id'):x.get('Target') for x in ET.fromstring(files[drawing_rels_name])}
    checked={}
    for anchor in saved:
        nv=anchor.find('./'+Q(XDR,'sp')+'/'+Q(XDR,'nvSpPr')+'/'+Q(XDR,'cNvPr'))
        if nv is None or not (nv.get('name') or '').startswith(PREFIX): continue
        frm=anchor.find(Q(XDR,'from'))
        row=int(frm.findtext(Q(XDR,'row')))+1
        hyperlink=nv.find(Q(A,'hlinkClick'))
        checked[f'E{row}']=relmap[hyperlink.get(Q(R,'id'))]
        if anchor.get('editAs')!='twoCell': raise ValueError('Wrong move/size behavior')
    if checked!=expected: raise ValueError('Hot spot links do not match row ASINs')
    destination.parent.mkdir(parents=True,exist_ok=True)
    fd,temp=tempfile.mkstemp(suffix='.xlsx',prefix='.image-hotspots-',dir=destination.parent)
    os.close(fd)
    try:
        with zipfile.ZipFile(temp,'w',zipfile.ZIP_DEFLATED) as archive:
            for info in infos: archive.writestr(info,files.pop(info.filename))
            for name,raw in files.items(): archive.writestr(name,raw)
        with zipfile.ZipFile(temp) as archive:
            if archive.testzip(): raise ValueError('ZIP CRC failure')
            for name in archive.namelist():
                if name.endswith(('.xml','.rels')): ET.fromstring(archive.read(name))
        os.replace(temp,destination)
    finally:
        if os.path.exists(temp): os.unlink(temp)
    return {'overlays':len(records),'existing_drawing_anchors_preserved':original_anchors,'native_image_and_richdata_parts_unchanged':len(protected),'drawing_part':drawing_name,'move_and_size_with_cells':True,'native_wps_click_tested':False,'native_sort_filter_tested':False,'note':'Transparent linked shapes sit above native in-cell images; sorting/filtering must be tested in the target WPS version.'}

if __name__=='__main__':
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('source',type=Path)
    parser.add_argument('destination',type=Path)
    args=parser.parse_args()
    print(json.dumps(add_image_link_overlays(args.source,args.destination),ensure_ascii=False))

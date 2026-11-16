#!/usr/bin/env python3
"""Add standard Hyperlink cell styling/tooltips while preserving native image parts.

This matches XlsxWriter embed_image(url=...)'s documented cell-hyperlink
mechanism. It does not invent a nonexistent _localImage URL field and cannot
prove WPS click behavior without a native-application test.
"""
from __future__ import annotations
import argparse
import copy
import hashlib
import json
import os
from pathlib import Path
import posixpath
import re
import tempfile
import xml.etree.ElementTree as ET
import zipfile

S='http://schemas.openxmlformats.org/spreadsheetml/2006/main'
R='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
P='http://schemas.openxmlformats.org/package/2006/relationships'
Q=lambda ns,name:f'{{{ns}}}{name}'

def xml_dump(root, original=None):
    ET.register_namespace('',root.tag.split('}')[0][1:])
    ET.register_namespace('r',R)
    result=ET.tostring(root,encoding='utf-8',xml_declaration=True)
    if original:
        first=re.search(rb'<(?!\?)[^>]+>',original)
        if first:
            for declaration in re.findall(rb'xmlns(?::[\w.-]+)?=["\'][^"\']+["\']',first.group()):
                start=re.search(rb'<(?!\?)[^>]+>',result)
                if start and declaration.split(b'=')[0]+b'=' not in start.group():
                    pos=start.end()-1
                    result=result[:pos]+b' '+declaration+result[pos:]
    return result

def resolve(owner,target):
    if target.startswith('/'): return target.lstrip('/')
    return posixpath.normpath(posixpath.join(posixpath.dirname(owner),target))

def find_sheet(files,name):
    wb=ET.fromstring(files['xl/workbook.xml'])
    relations=ET.fromstring(files['xl/_rels/workbook.xml.rels'])
    lookup={r.get('Id'):resolve('xl/workbook.xml',r.get('Target')) for r in relations}
    for sheet in wb.find(Q(S,'sheets')):
        if sheet.get('name')==name: return lookup[sheet.get(Q(R,'id'))]
    raise ValueError(f'No worksheet {name!r}')

def text_value(cell,strings):
    if cell.get('t')=='s': return strings[int(cell.findtext(Q(S,'v')))]
    if cell.get('t')=='inlineStr': return ''.join(t.text or '' for t in cell.iter(Q(S,'t')))
    return cell.findtext(Q(S,'v'))

def patch_link_styles(src,dst,add_asin_links=True):
    src,dst=Path(src),Path(dst)
    if src.resolve()==dst.resolve(): raise ValueError('Write to a reviewable copy; source and destination must differ')
    with zipfile.ZipFile(src) as archive:
        infos=archive.infolist()
        files={i.filename:archive.read(i.filename) for i in infos}
    protected={n:hashlib.sha256(b).hexdigest() for n,b in files.items() if n.startswith(('xl/richData/','xl/media/')) or n=='xl/metadata.xml'}
    part=find_sheet(files,'竞品明细')
    original_sheet=files[part]
    sheet=ET.fromstring(original_sheet)
    cells={c.get('r'):c for row in sheet.find(Q(S,'sheetData')) for c in row}
    strings=[]
    if 'xl/sharedStrings.xml' in files:
        strings=[''.join(t.text or '' for t in si.iter(Q(S,'t'))) for si in ET.fromstring(files['xl/sharedStrings.xml'])]
    asins={ref[1:]:text_value(cell,strings) for ref,cell in cells.items() if re.fullmatch(r'D[2-9]\d*|D1\d+',ref)}
    asins={row:asin for row,asin in asins.items() if re.fullmatch(r'[A-Z0-9]{10}',str(asin))}
    rel_path=posixpath.dirname(part)+'/_rels/'+posixpath.basename(part)+'.rels'
    relations=ET.fromstring(files[rel_path]) if rel_path in files else ET.Element(Q(P,'Relationships'))
    links=sheet.find(Q(S,'hyperlinks'))
    if links is None: raise ValueError('Source has no worksheet hyperlinks; expected existing E-image links')
    link_by_ref={x.get('ref'):x for x in links}
    relationship_by_id={x.get('Id'):x for x in relations}
    for row,asin in asins.items():
        expected='https://www.amazon.com/dp/'+asin
        ref='E'+row
        if ref not in link_by_ref: raise ValueError(f'Missing original image hyperlink {ref}')
        actual=relationship_by_id[link_by_ref[ref].get(Q(R,'id'))].get('Target')
        if actual!=expected: raise ValueError(f'Image hyperlink {ref} does not match D-column ASIN')
        link_by_ref[ref].set('tooltip',f'打开亚马逊商品 {asin}')
        if add_asin_links:
            ref='D'+row
            hyperlink=link_by_ref.get(ref)
            if hyperlink is None:
                n=1
                while f'rId{n}' in relationship_by_id: n+=1
                rid=f'rId{n}'
                relation=ET.SubElement(relations,Q(P,'Relationship'),{'Id':rid,'Type':R+'/hyperlink','Target':expected,'TargetMode':'External'})
                relationship_by_id[rid]=relation
                hyperlink=ET.SubElement(links,Q(S,'hyperlink'),{'ref':ref,Q(R,'id'):rid})
                link_by_ref[ref]=hyperlink
            else:
                relation=relationship_by_id[hyperlink.get(Q(R,'id'))]
                if relation.get('Target')!=expected: raise ValueError(f'Existing D-column hyperlink mismatches ASIN at {ref}')
            hyperlink.set('tooltip',f'打开亚马逊商品 {asin}')
    styles=ET.fromstring(files['xl/styles.xml'])
    fonts=styles.find(Q(S,'fonts'))
    xfs=styles.find(Q(S,'cellXfs'))
    style_xfs=styles.find(Q(S,'cellStyleXfs'))
    cell_styles=styles.find(Q(S,'cellStyles'))
    if any(x is None for x in [fonts,xfs,style_xfs,cell_styles]): raise ValueError('Incomplete style table')
    named=next((x for x in cell_styles if x.get('builtinId')=='8' or x.get('name')=='Hyperlink'),None)
    hyperlink_font_ids={}
    def hyperlink_font(old_id):
        if old_id in hyperlink_font_ids: return hyperlink_font_ids[old_id]
        font=copy.deepcopy(fonts[old_id])
        for tag in ['u','color']:
            for element in list(font):
                if element.tag==Q(S,tag): font.remove(element)
        ET.SubElement(font,Q(S,'u'))
        ET.SubElement(font,Q(S,'color'),{'theme':'10'})
        # Reuse a prior identical font instead of growing duplicates.
        for index,existing in enumerate(fonts):
            if ET.tostring(existing)==ET.tostring(font):
                hyperlink_font_ids[old_id]=index
                return index
        index=len(fonts); fonts.append(font); fonts.set('count',str(len(fonts)))
        hyperlink_font_ids[old_id]=index
        return index
    if named is None:
        style_xf_index=len(style_xfs)
        standard_font=hyperlink_font(int(xfs[0].get('fontId','0')))
        sx=ET.SubElement(style_xfs,Q(S,'xf'),{'numFmtId':'0','fontId':str(standard_font),'fillId':'0','borderId':'0','applyNumberFormat':'0','applyFill':'0','applyBorder':'0','applyAlignment':'0','applyProtection':'0'})
        ET.SubElement(sx,Q(S,'alignment'),{'vertical':'top'})
        ET.SubElement(sx,Q(S,'protection'),{'locked':'0'})
        style_xfs.set('count',str(len(style_xfs)))
        ET.SubElement(cell_styles,Q(S,'cellStyle'),{'name':'Hyperlink','xfId':str(style_xf_index),'builtinId':'8'})
        cell_styles.set('count',str(len(cell_styles)))
    else: style_xf_index=int(named.get('xfId'))
    style_cache={}
    touched=[]
    for ref,hyperlink in link_by_ref.items():
        cell=cells.get(ref)
        if cell is None: continue
        old=int(cell.get('s','0'))
        if old not in style_cache:
            xf=copy.deepcopy(xfs[old])
            xf.set('fontId',str(hyperlink_font(int(xf.get('fontId','0')))))
            xf.set('xfId',str(style_xf_index))
            xf.set('applyFont','1')
            style_cache[old]=len(xfs)
            xfs.append(xf)
        cell.set('s',str(style_cache[old]))
        touched.append(ref)
    xfs.set('count',str(len(xfs)))
    files['xl/styles.xml']=xml_dump(styles,files['xl/styles.xml'])
    files[part]=xml_dump(sheet,original_sheet)
    files[rel_path]=xml_dump(relations)
    for name,digest in protected.items():
        if hashlib.sha256(files[name]).hexdigest()!=digest: raise ValueError(f'Image/rich-data modified unexpectedly: {name}')
    dst.parent.mkdir(parents=True,exist_ok=True)
    fd,temp=tempfile.mkstemp(suffix='.xlsx',prefix='.linked-images-',dir=dst.parent)
    os.close(fd)
    try:
        with zipfile.ZipFile(temp,'w',zipfile.ZIP_DEFLATED) as archive:
            for info in infos: archive.writestr(info,files.pop(info.filename))
            for name,raw in files.items(): archive.writestr(name,raw)
        with zipfile.ZipFile(temp) as archive:
            if archive.testzip(): raise ValueError('ZIP CRC failure')
            for name in archive.namelist():
                if name.endswith(('.xml','.rels')): ET.fromstring(archive.read(name))
        os.replace(temp,dst)
    finally:
        if os.path.exists(temp): os.unlink(temp)
    return {'asins':len(asins),'image_cell_links':sum(ref.startswith('E') for ref in link_by_ref),'asin_text_links':sum(ref.startswith('D') for ref in link_by_ref),'hyperlinks':len(link_by_ref),'styled_hyperlinks':len(touched),'protected_image_and_richdata_parts_unchanged':len(protected),'native_wps_click_tested':False,'mechanism':'standard worksheet cell hyperlink + built-in Hyperlink style (XlsxWriter embed_image(url=...) mechanism)'}

if __name__=='__main__':
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('source',type=Path)
    parser.add_argument('destination',type=Path)
    parser.add_argument('--no-asin-text-links',action='store_true')
    args=parser.parse_args()
    print(json.dumps(patch_link_styles(args.source,args.destination,not args.no_asin_text_links),ensure_ascii=False))

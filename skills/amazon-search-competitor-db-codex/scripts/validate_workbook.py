#!/usr/bin/env python3
"""Validate the fixed schema, typed dates, native images, links and WPS click layer."""
from __future__ import annotations
import argparse
import copy
from datetime import date,datetime,timedelta,timezone
import json
import math
from pathlib import Path
import re
import xml.etree.ElementTree as ET
import zipfile
from in_cell_images import MAIN,REL,RD,Q,sheet_part,rel_path,resolve_part

XDR='http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'
DRAW='http://schemas.openxmlformats.org/drawingml/2006/main'
BEIJING=timezone(timedelta(hours=8))
EPOCH=date(1899,12,30)
HOTSPOT_PREFIX='CodexASINHotspot_'
COMPANY_FIELDS={'公司主体','公司地址','主体所在国家/地区'}

def removed_legacy_field(value):
    if not isinstance(value,str): return False
    return ('商标' in value or 'trademark' in value.casefold() or
            re.fullmatch(r'\$?(?:AG|AH|AI)\$?\d*',value.strip(),re.I) is not None)

def legacy_note(value):
    text=value if isinstance(value,str) else json.dumps(value,ensure_ascii=False)
    return re.search(r'商标|trademark|\bUSPTO\b|\bTSDR\b',text,re.I) is not None

def remove_legacy_export_data(data):
    """Discard retired result fields/caches; ordinary brand/company facts stay."""
    for key in list(data):
        if removed_legacy_field(key): data.pop(key,None)
    for row in [*data.get('records',[]),*data.get('variants',[])]:
        if not isinstance(row,dict): continue
        for key in list(row):
            if removed_legacy_field(key): row.pop(key,None)
        for field in ['_sources','_field_provenance']:
            mapping=row.get(field)
            if isinstance(mapping,dict): row[field]={key:value for key,value in mapping.items() if not removed_legacy_field(key) and not (isinstance(value,dict) and value.get('source_type')=='third_party_trademark_candidates')}
        for field in ['_highlight_fields','_highlight_fields_resolved']:
            value=row.get(field)
            if isinstance(value,dict): row[field]={key:color for key,color in value.items() if not removed_legacy_field(key)}
            elif isinstance(value,list): row[field]=[key for key in value if not removed_legacy_field(key)]
            elif removed_legacy_field(value): row.pop(field,None)
    if isinstance(data.get('notes'),list): data['notes']=[note for note in data['notes'] if not legacy_note(note)]
    return data

def excel_date_serial(value):
    """Calendar date: YYYY-MM-DD, aware ISO, or Unix milliseconds -> integer."""
    if value is None or value=='': return None
    if isinstance(value,bool): raise ValueError('A Boolean is not a calendar date')
    if isinstance(value,(int,float)):
        day=datetime.fromtimestamp(value/1000,tz=timezone.utc).astimezone(BEIJING).date()
    elif re.fullmatch(r'\d{4}-\d{2}-\d{2}',str(value)):
        day=date.fromisoformat(value)
    else:
        dt=datetime.fromisoformat(str(value).replace('Z','+00:00'))
        if dt.tzinfo is None: raise ValueError('Listing date ISO datetime must include a timezone')
        day=dt.astimezone(BEIJING).date()
    serial=(day-EPOCH).days
    if serial<=0: raise ValueError(f'Date outside supported Excel range: {value}')
    return serial

def excel_datetime_serial(value):
    """Unix ms / aware ISO / legacy Beijing YYYY-MM-DD HH:MM -> wall-clock serial."""
    if value is None or value=='': return None
    if isinstance(value,bool): raise ValueError('A Boolean is not an update timestamp')
    if isinstance(value,(int,float)) or re.fullmatch(r'-?\d{12,16}',str(value)):
        milliseconds=float(value)
        if not math.isfinite(milliseconds): raise ValueError('Non-finite update timestamp')
        dt=datetime.fromtimestamp(milliseconds/1000,tz=timezone.utc)
    elif re.fullmatch(r'\d{4}-\d{2}-\d{2} \d{2}:\d{2}(?::\d{2}(?:\.\d{1,6})?)?',str(value)):
        dt=datetime.fromisoformat(str(value)).replace(tzinfo=BEIJING)
    else:
        dt=datetime.fromisoformat(str(value).replace('Z','+00:00'))
        if dt.tzinfo is None: raise ValueError('Update ISO datetime needs a timezone; legacy space-separated timestamps are Beijing time')
    local=dt.astimezone(BEIJING)
    fraction=(local.hour*3600+local.minute*60+local.second+local.microsecond/1e6)/86400
    return (local.date()-EPOCH).days+fraction

def expected_date_values(record):
    updated=record.get('_sellersprite_updated_at_ms')
    if updated is None or updated=='': updated=record.get('卖家精灵数据更新时间')
    return {'上架时间':excel_date_serial(record.get('上架时间')),'卖家精灵数据更新时间':excel_datetime_serial(updated)}

def highlighted_fields(record,headers):
    result=set()
    for field,provenance in (record.get('_field_provenance') or {}).items():
        if not isinstance(provenance,dict): continue
        third=provenance.get('source_type')=='third_party_database'
        if field in headers and third and provenance.get('verified_on_amazon') is not True and field in COMPANY_FIELDS: result.add(field)
    old=record.get('_highlight_fields') or []
    if isinstance(old,dict): old=[key for key,value in old.items() if value is not False and value is not None]
    if isinstance(old,str): old=[old]
    for field in old:
        if field in headers: result.add(field); continue
        match=re.fullmatch(r'([A-Z]+)\d*',str(field))
        if match:
            number=0
            for letter in match.group(1): number=number*26+ord(letter)-64
            if 1<=number<=len(headers): result.add(headers[number-1])
    if record.get('_sales_estimated'): result.discard('近一月销量')
    return result

def cell_text(cell,shared):
    if cell is None: return None
    if cell.get('t')=='inlineStr': return ''.join(t.text or '' for t in cell.iter(Q(MAIN,'t')))
    value=cell.findtext(Q(MAIN,'v'))
    return shared[int(value)] if cell.get('t')=='s' else value

def column_letter(number):
    result=''
    while number: number,rem=divmod(number-1,26); result=chr(65+rem)+result
    return result

def validate_workbook(path,data=None,schema=None,require_click_layer=True,require_asin_text_links=True):
    if data is not None: data=remove_legacy_export_data(copy.deepcopy(data))
    if schema is None: schema=json.loads((Path(__file__).resolve().parent.parent/'references/schema.json').read_text(encoding='utf-8'))
    if len(schema['headers'])!=32: raise ValueError('Schema must contain exactly 32 columns')
    with zipfile.ZipFile(path) as archive:
        if archive.testzip(): raise ValueError('ZIP CRC check failed')
        files={name:archive.read(name) for name in archive.namelist()}
    for name,raw in files.items():
        if name.endswith(('.xml','.rels')): ET.fromstring(raw)
    for name in ['竞品明细','变体明细','说明']: sheet_part(files,name)
    shared=[]
    if 'xl/sharedStrings.xml' in files: shared=[''.join(t.text or '' for t in si.iter(Q(MAIN,'t'))) for si in ET.fromstring(files['xl/sharedStrings.xml'])]
    part=sheet_part(files,'竞品明细')
    sheet=ET.fromstring(files[part])
    cells={c.get('r'):c for row in sheet.find(Q(MAIN,'sheetData')) for c in row}
    headers=[cell_text(cells.get(column_letter(i)+'1'),shared) for i in range(1,33)]
    if headers!=schema['headers']: raise ValueError('32-column header/order mismatch')
    for ref in cells:
        letters=re.match(r'[A-Z]+',ref).group(); number=0
        for letter in letters: number=number*26+ord(letter)-64
        if number>32: raise ValueError('Main worksheet contains retired columns after AF')
    pane=sheet.find('./'+Q(MAIN,'sheetViews')+'/'+Q(MAIN,'sheetView')+'/'+Q(MAIN,'pane'))
    if pane is None or pane.get('topLeftCell')!='F2' or pane.get('xSplit')!='5' or pane.get('ySplit')!='1': raise ValueError('Freeze pane must be F2')
    if sheet.find(Q(MAIN,'autoFilter')) is None: raise ValueError('Worksheet autoFilter missing')
    rels=ET.fromstring(files[rel_path(part)]) if rel_path(part) in files else []
    lookup={r.get('Id'):r for r in rels}
    hyperlinks=list(sheet.find(Q(MAIN,'hyperlinks'))) if sheet.find(Q(MAIN,'hyperlinks')) is not None else []
    if len({h.get('ref') for h in hyperlinks})!=len(hyperlinks): raise ValueError('Duplicate cell hyperlinks')
    linkmap={}
    for h in hyperlinks:
        relation=lookup[h.get(Q(REL,'id'))]
        if relation.get('Type')!=REL+'/hyperlink' or relation.get('TargetMode')!='External': raise ValueError('Invalid hyperlink relationship')
        linkmap[h.get('ref')]=relation.get('Target')
    asins={int(ref[1:]):cell_text(cell,shared) for ref,cell in cells.items() if re.fullmatch(r'D\d+',ref) and int(ref[1:])>=2 and cell_text(cell,shared)}
    if sheet.find(Q(MAIN,'autoFilter')).get('ref')!=f'A1:AF{max(1,len(asins)+1)}': raise ValueError('Main autofilter must end at AF and the actual data row')
    if any(not re.fullmatch(r'[A-Z0-9]{10}',str(a)) for a in asins.values()): raise ValueError('Invalid main ASIN value')
    image_cells={ref:c for ref,c in cells.items() if c.get('vm')}
    for ref,c in cells.items():
        if c.get('t')=='e' and not c.get('vm'): raise ValueError(f'Unexpected Excel error at {ref}')
    styles=ET.fromstring(files['xl/styles.xml']); xfs=styles.find(Q(MAIN,'cellXfs')); fills=styles.find(Q(MAIN,'fills'))
    formats={int(x.get('numFmtId')):x.get('formatCode') for x in styles.find(Q(MAIN,'numFmts'))} if styles.find(Q(MAIN,'numFmts')) is not None else {}
    def format_code(cell): return formats.get(int(xfs[int(cell.get('s','0'))].get('numFmtId','0')),'').replace('"','').replace('\\','')
    def fill_color(cell):
        fill=fills[int(xfs[int(cell.get('s','0'))].get('fillId','0'))]
        color=fill.find('.//'+Q(MAIN,'fgColor'))
        return color.get('rgb','')[-6:].upper() if color is not None else ''
    def check_date_cell(ref,expected,date_only=False,cell_map=None):
        cell=(cells if cell_map is None else cell_map).get(ref); actual=cell.findtext(Q(MAIN,'v')) if cell is not None else None
        if expected is None:
            if actual is not None and actual!='': raise ValueError(f'Unexpected substituted date/time at {ref}')
            return
        if cell is None or cell.get('t')!='n': raise ValueError(f'{ref} must be numeric Excel date/time (t=n)')
        actual=float(actual)
        if abs(actual-expected)>1e-8 or (date_only and not actual.is_integer()): raise ValueError(f'Date/time mismatch at {ref}')
        expected_format='yyyy年mm月dd日' if date_only else 'yyyy-mm-dd hh:mm'
        if format_code(cell)!=expected_format: raise ValueError(f'Wrong date/time format at {ref}: {format_code(cell)}')
    # Even without source JSON, reject ISO-string dates that WPS displays as text.
    for row in asins:
        for column,date_only in [('L',True),('X',False)]:
            ref=f'{column}{row}'; cell=cells.get(ref)
            raw=cell.findtext(Q(MAIN,'v')) if cell is not None else None
            if raw is not None and raw!='':
                if cell.get('t')!='n': raise ValueError(f'{ref} must be a numeric Excel date/time, not a text/ISO date')
                check_date_cell(ref,float(raw),date_only)
    verified_image_bytes=0
    if image_cells:
        rv=ET.fromstring(files['xl/richData/rdrichvalue.xml']); md=ET.fromstring(files['xl/metadata.xml']); rr=ET.fromstring(files['xl/richData/richValueRel.xml'])
        image_relations={r.get('Id'):r for r in ET.fromstring(files['xl/richData/_rels/richValueRel.xml.rels'])}
        value=md.find(Q(MAIN,'valueMetadata')); future=next((f for f in md.findall(Q(MAIN,'futureMetadata')) if f.get('name')=='XLRICHVALUE'),None)
        structures=ET.fromstring(files['xl/richData/rdrichvaluestructure.xml'])
        for root in [rv,value,future]:
            if root is None or int(root.get('count'))!=len(root): raise ValueError('Rich-data/metadata declared count mismatch')
        for ref,cell in image_cells.items():
            if not re.fullmatch(r'E\d+',ref): raise ValueError('Unexpected native image column')
            if cell.get('t')!='e' or cell.findtext(Q(MAIN,'v'))!='#VALUE!': raise ValueError('Wrong native image cell representation')
            vm=int(cell.get('vm'))-1
            if not 0<=vm<len(value): raise ValueError('Native image metadata index outside range')
            fi=int(value[vm][0].get('v'))
            if not 0<=fi<len(future): raise ValueError('Native image future-metadata index outside range')
            rvb=future[fi].find('.//'+Q(RD,'rvb')); rvi=int(rvb.get('i'))
            if not 0<=rvi<len(rv): raise ValueError('Native rich-value index outside range')
            record=rv[rvi]; structure=structures[int(record.get('s'))]
            if structure.get('t')!='_localImage': raise ValueError('Wrong rich-value structure')
            keys=[k.get('n') for k in structure]; image_index=int(record[keys.index('_rvRel:LocalImageIdentifier')].text)
            if not 0<=image_index<len(rr): raise ValueError('Native image relationship index outside range')
            relation=image_relations[rr[image_index].get(Q(REL,'id'))]
            if relation.get('Type')!=REL+'/image': raise ValueError('Invalid image relationship')
            raw=files[resolve_part('xl/richData/richValueRel.xml',relation.get('Target'))]
            if not (raw.startswith(b'\x89PNG\r\n\x1a\n') or raw.startswith(b'\xff\xd8\xff')): raise ValueError('Invalid embedded image')
            if data is not None:
                source=data['records'][int(ref[1:])-2].get('_image')
                if not source or raw!=Path(source).read_bytes(): raise ValueError(f'Image/source byte mismatch at {ref}')
                verified_image_bytes+=1
    for name,raw in files.items():
        if not name.endswith('.rels'): continue
        if name=='_rels/.rels': owner=''
        else:
            parent,base=name.rsplit('/_rels/',1); owner=parent+'/'+base[:-5]
        root=ET.fromstring(raw)
        if len({r.get('Id') for r in root})!=len(root): raise ValueError(f'Duplicate relationship IDs: {name}')
        for relation in root:
            if relation.get('TargetMode')!='External' and resolve_part(owner,relation.get('Target','')) not in files: raise ValueError(f'Broken relationship: {name}')
    overlays={}
    drawing_tag=sheet.find(Q(MAIN,'drawing'))
    if drawing_tag is not None:
        drawing_name=resolve_part(part,lookup[drawing_tag.get(Q(REL,'id'))].get('Target')); drawing=ET.fromstring(files[drawing_name])
        draw_rels={r.get('Id'):r for r in ET.fromstring(files[rel_path(drawing_name)])}
        ids=[n.get('id') for n in drawing.iter(Q(XDR,'cNvPr'))]
        if len(set(ids))!=len(ids): raise ValueError('Duplicate drawing object IDs')
        for anchor in drawing:
            shape=anchor.find(Q(XDR,'sp'))
            if shape is None: continue
            nv=shape.find('./'+Q(XDR,'nvSpPr')+'/'+Q(XDR,'cNvPr'))
            if nv is None or not (nv.get('name') or '').startswith(HOTSPOT_PREFIX): continue
            if anchor.tag!=Q(XDR,'twoCellAnchor') or anchor.get('editAs')!='twoCell': raise ValueError('Hot spot must move/size with cells')
            frm,to=anchor.find(Q(XDR,'from')),anchor.find(Q(XDR,'to')); row=int(frm.findtext(Q(XDR,'row')))+1; ref=f'E{row}'
            if ref in overlays: raise ValueError('Duplicate image click overlay')
            expected_from={'col':'4','colOff':'0','row':str(row-1),'rowOff':'0'}; expected_to={'col':'5','colOff':'0','row':str(row),'rowOff':'0'}
            if any(frm.findtext(Q(XDR,k))!=v for k,v in expected_from.items()) or any(to.findtext(Q(XDR,k))!=v for k,v in expected_to.items()): raise ValueError('Overlay does not cover its E cell')
            props=shape.find(Q(XDR,'spPr')); alpha=props.find('./'+Q(DRAW,'solidFill')+'/'+Q(DRAW,'srgbClr')+'/'+Q(DRAW,'alpha'))
            if alpha is None or alpha.get('val')!='0' or props.find('./'+Q(DRAW,'ln')+'/'+Q(DRAW,'noFill')) is None: raise ValueError('Overlay must be transparent with no outline')
            if props.find(Q(DRAW,'prstGeom')).get('prst')!='rect' or any((t.text or '') for t in shape.iter(Q(DRAW,'t'))): raise ValueError('Unexpected overlay geometry/text')
            click=nv.find(Q(DRAW,'hlinkClick')); relation=draw_rels[click.get(Q(REL,'id'))]; url=f'https://www.amazon.com/dp/{asins[row]}'
            if relation.get('Type')!=REL+'/hyperlink' or relation.get('TargetMode')!='External' or relation.get('Target')!=url: raise ValueError('Overlay/ASIN mismatch')
            overlays[ref]=url
    if (require_click_layer and image_cells) or overlays:
        if set(overlays)!=set(image_cells): raise ValueError('Click overlay rows do not match native images')
    records=data.get('records',[]) if data is not None else None
    expected_count=len(records) if records is not None else len(asins)
    if len(asins)!=expected_count: raise ValueError('Product row count mismatch')
    for row,asin in asins.items():
        expected=f'https://www.amazon.com/dp/{asin}'
        if linkmap.get(f'E{row}')!=expected: raise ValueError(f'Missing E{row} image-cell hyperlink')
        if require_asin_text_links and linkmap.get(f'D{row}')!=expected: raise ValueError(f'Missing D{row} ASIN text hyperlink')
    if records is not None:
        if len(image_cells)!=sum(bool(r.get('_image')) for r in records): raise ValueError('Native image count mismatch')
        for row,record in enumerate(records,2):
            if asins.get(row)!=record['ASIN']: raise ValueError(f'ASIN mismatch at D{row}')
            if cells[f'D{row}'].get('t') not in {'s','inlineStr','str'}: raise ValueError('ASIN must remain text')
            dates=expected_date_values(record); check_date_cell(f'L{row}',dates['上架时间'],True); check_date_cell(f'X{row}',dates['卖家精灵数据更新时间'])
            for field in ['标题','公司主体','公司地址']:
                original=record.get(field)
                if original is not None and isinstance(original,str):
                    actual=cell_text(cells.get(f'{column_letter(headers.index(field)+1)}{row}'),shared)
                    expected="'"+original if re.match(r'^[=+@]',original) else original
                    if actual!=expected: raise ValueError(f'Text changed/truncated at {field}, row {row}')
            for field in highlighted_fields(record,headers):
                if fill_color(cells[f'{column_letter(headers.index(field)+1)}{row}'])!='FFF2CC': raise ValueError(f'Missing third-party yellow marker: {field}, row {row}')
            if record.get('_sales_estimated') and fill_color(cells[f'K{row}'])!='FCE4D6': raise ValueError('Estimated sales orange marker missing')
            if record.get('卖家主页') and linkmap.get(f'AF{row}')!=record['卖家主页']: raise ValueError('Seller hyperlink mismatch')
        expected_links=expected_count*(2 if require_asin_text_links else 1)+sum(bool(r.get('卖家主页')) for r in records)
        if len(hyperlinks)!=expected_links: raise ValueError('Cell hyperlink count mismatch')
        variant_sheet=ET.fromstring(files[sheet_part(files,'变体明细')])
        variant_cells={c.get('r'):c for c in variant_sheet.iter(Q(MAIN,'c'))}
        variants=data.get('variants',[])
        variant_headers=list(dict.fromkeys(['父ASIN','ASIN']+[key for row in variants for key in row if not key.startswith('_')]))
        date_fields=data.get('variant_date_fields',['上架时间'])
        datetime_fields=data.get('variant_datetime_fields',['卖家精灵数据更新时间'])
        for row,record in enumerate(variants,2):
            for index,field in enumerate(variant_headers,1):
                ref=f'{column_letter(index)}{row}'
                if field in date_fields: check_date_cell(ref,excel_date_serial(record.get(field)),True,variant_cells)
                elif field in datetime_fields: check_date_cell(ref,excel_datetime_serial(record.get(field)),False,variant_cells)
                elif isinstance(record.get(field),str):
                    expected=record[field]
                    if re.match(r'^[=+@]',expected): expected="'"+expected
                    if cell_text(variant_cells.get(ref),shared)!=expected: raise ValueError(f'Variant text changed/truncated at {ref}')
    return {'passed':True,'columns':32,'image_cells':len(image_cells),'click_overlays':len(overlays),'hyperlinks':len(hyperlinks),'source_image_bytes_verified':verified_image_bytes,'xml_parts':sum(n.endswith(('.xml','.rels')) for n in files),'native_application_visual_check':False,'native_sort_filter_check':False}

def main():
    parser=argparse.ArgumentParser(description=__doc__); parser.add_argument('workbook',type=Path); parser.add_argument('--input',type=Path)
    parser.add_argument('--no-wps-click-layer',action='store_true'); parser.add_argument('--no-asin-text-links',action='store_true'); args=parser.parse_args()
    data=json.loads(args.input.read_text(encoding='utf-8')) if args.input else None
    if data is not None:
        for record in data.get('records',[]):
            if record.get('_image'):
                image=Path(record['_image'])
                if not image.is_absolute(): record['_image']=str(args.input.resolve().parent/image)
    print(json.dumps(validate_workbook(args.workbook,data,require_click_layer=not args.no_wps_click_layer,require_asin_text_links=not args.no_asin_text_links),ensure_ascii=False))

if __name__=='__main__': main()

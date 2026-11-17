#!/usr/bin/env python3
"""Offline installation regression test; generated fixtures contain no live data."""
from pathlib import Path
from datetime import datetime
import argparse
import json
import struct
import subprocess
import sys
import tempfile
import xml.etree.ElementTree as ET
import zipfile
import zlib

SKILL=Path(__file__).resolve().parents[1]
S='http://schemas.openxmlformats.org/spreadsheetml/2006/main'
XDR='http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'
Q=lambda name:f'{{{S}}}{name}'

def png(rgb):
    def chunk(kind,data):
        return struct.pack('>I',len(data))+kind+data+struct.pack('>I',zlib.crc32(kind+data)&0xffffffff)
    pixels=b''.join(b'\0'+bytes(rgb)*8 for _ in range(8))
    return b'\x89PNG\r\n\x1a\n'+chunk(b'IHDR',struct.pack('>IIBBBBB',8,8,8,2,0,0,0))+chunk(b'IDAT',zlib.compress(pixels))+chunk(b'IEND',b'')

def build(directory,count):
    records=[]
    for i in range(1,count+1):
        image=directory/f'image-{i}.png'
        image.write_bytes(png((i%256,(i*37)%256,(i*71)%256)))
        records.append({'自然排名':i,'页码':(i-1)//50+1,'页内位置':(i-1)%50+1,
            'ASIN':f'B0TEST{i:04d}','标题':f'合成测试商品 {i}','品牌':'合成测试品牌',
            '价格':19.99 if i==count else 10.25,'近一月销量':2000,'变体数量':3,'上架时间':'2024-01-02',
            'FBA配送费':4.5 if i%2 else None,'包装尺寸(cm)':[30,12.5,8] if i%2 else None,'包装重量(g)':250 if i%2 else -1,
            '评论数':1200,'大类BSR排名':3000,'小类BSR排名':40,'卖家ASIN数量':8,'卖家最佳自然排名':i,
            '库存状态':'无购买框' if i==count else '有货',
            '卖家ID':f'TESTSELLER{i}','卖家主页':f'https://www.amazon.com/sp?seller=TESTSELLER{i}',
            '公司主体':'合成公司','公司地址':'合成地址，仅用于离线测试',
            '商标持有人':'STALE_TRADEMARK_HOLDER','商标持有人国家/地区':'STALE_TRADEMARK_COUNTRY',
            '商标信息':'STALE_TRADEMARK_INFO','_trademark_details':'STALE_TRADEMARK_DETAILS',
            '_sources':{'商标持有人':'STALE_TRADEMARK_SOURCE','AG':'STALE_TRADEMARK_ALIAS','公司主体':'保留的合成公司来源'},
            '_highlight_fields':['AG','AH2','AI99','商标信息','公司地址'],
            '_highlight_fields_resolved':['商标持有人','公司主体'],
            '_sellersprite_updated_at_ms':1704153600000,'_image':image.name,
            '_sales_display':'2K+ bought in past month','_sales_lower_bound':True,'_sales_estimated':True,
            '_field_provenance':{
              '公司主体':{'source_type':'third_party_database','source':'合成来源','verified_on_amazon':False},
              '公司地址':{'source_type':'third_party_database','source':'合成来源','verified_on_amazon':False},
              '商标持有人':{'source_type':'third_party_trademark_candidates','source':'STALE_TRADEMARK_PROVENANCE'},
              '商标持有人国家/地区':{'source_type':'third_party_trademark_candidates','source':'STALE_TRADEMARK_PROVENANCE'}}})
    payload={'keyword':'offline synthetic sample','marketplace':'US','captured_at':'2024-02-01T00:00:00+00:00',
             'records':records,'variants':[{'父ASIN':'B0TEST0001','ASIN':'B0TEST0001','说明':'保留的合成变体说明','商标信息':'STALE_TRADEMARK_VARIANT','_sources':{'商标信息':'STALE_TRADEMARK_VARIANT_SOURCE'}}] if records else [],
             'notes':['合成数据，禁止作为真实竞品调研交付。','商标旧缓存：STALE_TRADEMARK_NOTE'],
             'trademark_research':{'mode':'full'},'_trademark_scope_note':'STALE_TRADEMARK_SCOPE','points_used':None}
    target=directory/'input.json'
    target.write_text(json.dumps(payload,ensure_ascii=False),encoding='utf-8')
    return target

def run(input_path,output,*options,success=True):
    p=subprocess.run([sys.executable,str(SKILL/'scripts/render_workbook.py'),str(input_path),'--output',str(output),*options],text=True,capture_output=True)
    if success and p.returncode: raise AssertionError(p.stdout+p.stderr)
    if not success and not p.returncode: raise AssertionError('Expected rejection, but workbook was created')
    return p

def shared_strings(archive):
    if 'xl/sharedStrings.xml' not in archive.namelist(): return []
    root=ET.fromstring(archive.read('xl/sharedStrings.xml'))
    return [''.join(node.text or '' for node in item.iter(Q('t'))) for item in root]

def cell_text(cell,shared):
    if cell is None: return None
    if cell.get('t')=='inlineStr': return ''.join(node.text or '' for node in cell.iter(Q('t')))
    value=cell.findtext(Q('v'))
    return shared[int(value)] if cell.get('t')=='s' else value

def verify(output,count,overlays):
    schema=json.loads((SKILL/'references/schema.json').read_text(encoding='utf-8'))
    assert schema['headers']==['自然排名','页码','页内位置','ASIN','主图','标题','品牌','价格','近一月销量','变体数量','上架时间','FBA配送费','包装尺寸(cm)','包装重量(g)','评分','评论数','大类BSR排名','大类','小类BSR排名','小类名称','父ASIN','库存状态','卖家','配送方式','主体所在国家/地区','公司地址','卖家精灵数据更新时间','卖家ID','店铺名称','公司主体','卖家ASIN数量','卖家最佳自然排名','卖家主页']
    assert '原价/参考价' not in schema['headers'] and '折扣' not in schema['headers']
    with zipfile.ZipFile(output) as z:
        assert z.testzip() is None
        sheet=ET.fromstring(z.read('xl/worksheets/sheet1.xml'))
        styles=ET.fromstring(z.read('xl/styles.xml'))
        shared=shared_strings(z)
        cells={c.get('r'):c for c in sheet.iter(Q('c'))}
        header_row=next(row for row in sheet.find(Q('sheetData')) if row.get('r')=='1')
        assert len(list(header_row))==33
        assert abs(float(header_row.get('ht'))-40)<0.05
        assert [cell_text(cells[f'{chr(65+i) if i<26 else "A"+chr(65+i-26)}1'],shared) for i in range(33)]==schema['headers']
        assert sheet.find(Q('autoFilter')).get('ref')==f'A1:AG{max(1,count+1)}'
        xml=b'\n'.join(z.read(name) for name in z.namelist() if name.endswith('.xml'))
        assert b'STALE_TRADEMARK' not in xml
        assert '商标'.encode() not in xml
        assert '原价/参考价'.encode() not in xml
        assert '折扣'.encode() not in xml
        assert '合成公司'.encode() in xml if count else True
        assert '保留的合成变体说明'.encode() in xml if count else True
        images=[c for c in cells.values() if c.get('vm')]
        assert len(images)==count
        links=sheet.find(Q('hyperlinks'))
        assert (len(links) if links is not None else 0)==count*3
        xfs=styles.find(Q('cellXfs'))
        fills=styles.find(Q('fills'))
        formats={int(item.get('numFmtId')):item.get('formatCode') for item in styles.find(Q('numFmts'))} if styles.find(Q('numFmts')) is not None else {}
        builtin={3:'#,##0'}
        def format_code(cell):
            fmt=int(xfs[int(cell.get('s','0'))].get('numFmtId','0'))
            return (formats.get(fmt) or builtin.get(fmt) or '').split(';')[0]
        def fill_color(cell):
            fill=fills[int(xfs[int(cell.get('s','0'))].get('fillId','0'))]
            color=fill.find('.//'+Q('fgColor'))
            return color.get('rgb','')[-6:].upper() if color is not None else ''
        widths={}
        for col in sheet.find(Q('cols')):
            for index in range(int(col.get('min')),int(col.get('max'))+1): widths[index]=float(col.get('width'))
        assert abs(widths[11]-12)<0.05 and abs(widths[27]-12)<0.05
        listed=(datetime(2024,1,2)-datetime(1899,12,30)).days
        for i in range(2,count+2):
            assert cells[f'K{i}'].get('t') in {None,'n'}
            assert float(cells[f'K{i}'].findtext(Q('v')))==listed
            assert format_code(cells[f'K{i}'])=='yyyy-mm-dd'
            assert cells[f'AA{i}'].get('t') in {None,'n'}
            assert float(cells[f'AA{i}'].findtext(Q('v')))==listed
            assert format_code(cells[f'AA{i}'])=='yyyy-mm-dd'
            assert cells[f'K{i}'].find(Q('f')) is None and cells[f'J{i}'].find(Q('f')) is None
            assert format_code(cells[f'I{i}'])=='#,##0'
            assert fill_color(cells[f'I{i}'])=='FCE4D6'
            record_number=i-1
            if record_number%2:
                assert format_code(cells[f'L{i}'])=='$#,##0.00'
                assert cell_text(cells[f'M{i}'],shared)=='30 × 12.5 × 8'
                assert format_code(cells[f'N{i}'])=='#,##0'
                assert float(cells[f'N{i}'].findtext(Q('v')))==250
            else:
                assert cells.get(f'L{i}') is None or not cells[f'L{i}'].findtext(Q('v'))
                assert cell_text(cells.get(f'M{i}'),shared) in (None,'')
                assert cells.get(f'N{i}') is None or not cells[f'N{i}'].findtext(Q('v'))
            assert fill_color(cells[f'N{i}'])!='FCE4D6' if cells.get(f'N{i}') is not None else True
            if record_number==count:
                assert cell_text(cells.get(f'H{i}'),shared) in (None,'')
            else:
                assert format_code(cells[f'H{i}'])=='$#,##0.00'
        for ref,cell in cells.items():
            if fill_color(cell)=='FCE4D6':
                assert ref.startswith('I') and ref[1:].isdigit()
        drawing_files=[n for n in z.namelist() if n.startswith('xl/drawings/') and n.endswith('.xml')]
        shapes=sum(len(list(ET.fromstring(z.read(n)).iter(f'{{{XDR}}}sp'))) for n in drawing_files)
        pictures=sum(len(list(ET.fromstring(z.read(n)).iter(f'{{{XDR}}}pic'))) for n in drawing_files)
        assert shapes==overlays and pictures==0

def verify_company_source_retained():
    sys.path.insert(0,str(SKILL/'scripts'))
    from validate_workbook import remove_legacy_export_data,highlighted_fields,normalize_fba_fee,normalize_package_size,normalize_package_weight
    assert normalize_fba_fee(None) is None and normalize_fba_fee('') is None and normalize_fba_fee(-1) is None
    assert normalize_fba_fee('3.25')==3.25
    assert normalize_package_size(None) is None and normalize_package_size([30,None,8]) is None
    assert normalize_package_size([30,'12.5',8])=='30 × 12.5 × 8'
    assert normalize_package_size('10,5,3')=='10 × 5 × 3'
    assert normalize_package_weight(None) is None and normalize_package_weight(-1) is None and normalize_package_weight(250.9)==250
    source='合成 Trademark Company 资料页，仅作测试'
    provenance={'source_type':'third_party_database','source':source,'verified_on_amazon':False}
    data={'records':[{'品牌':'Synthetic Trademark Brand','公司主体':'合成 Trademark Company',
        '_sources':{'公司主体':source,'商标持有人':'STALE_TRADEMARK'},
        '_field_provenance':{'公司主体':provenance,'商标持有人':{'source_type':'third_party_trademark_candidates'}}}]}
    remove_legacy_export_data(data)
    row=data['records'][0]
    assert row['_sources']=={'公司主体':source} and row['_field_provenance']=={'公司主体':provenance}
    assert row['品牌']=='Synthetic Trademark Brand' and row['公司主体']=='合成 Trademark Company'
    assert highlighted_fields(row,['品牌','公司主体'])=={'公司主体'}

def main():
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--work-dir',type=Path)
    args=parser.parse_args()
    verify_company_source_retained()
    if args.work_dir: args.work_dir.mkdir(parents=True,exist_ok=True)
    with tempfile.TemporaryDirectory(prefix='codex-skill-test-',dir=args.work_dir) as temp:
        root=Path(temp)
        results=[]
        for count in (1,150,0):
            directory=root/f'rows-{count}'; directory.mkdir()
            input_path=build(directory,count); output=directory/'output.xlsx'
            run(input_path,output); verify(output,count,count)
            results.append({'rows':count,'passed':True})
        directory=root/'no-click-layer'; directory.mkdir()
        input_path=build(directory,1); output=directory/'output.xlsx'
        run(input_path,output,'--no-wps-click-layer'); verify(output,1,0)
        directory=root/'missing-image'; directory.mkdir()
        input_path=build(directory,1)
        for image in directory.glob('*.png'): image.unlink()
        output=directory/'output.xlsx'
        run(input_path,output,success=False); assert not output.exists()
        run(input_path,output,'--allow-missing-images')
        results.extend([{'no_click_layer':True,'passed':True},{'missing_image_default_rejected':True,'explicit_partial_delivery':True}])
        print(json.dumps(results,ensure_ascii=False))

if __name__=='__main__': main()

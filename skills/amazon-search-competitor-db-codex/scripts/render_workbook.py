#!/usr/bin/env python3
"""Create the fixed 33-column competitor workbook from collected/enriched JSON."""
from __future__ import annotations
import argparse
import json
import os
from pathlib import Path
import re
import shutil
import subprocess
import sys
import tempfile
from datetime import datetime
from zoneinfo import ZoneInfo
from in_cell_images import finish_workbook
from validate_workbook import validate_workbook,expected_date_values,excel_date_serial,excel_beijing_date_serial,highlighted_fields,remove_legacy_export_data,normalize_fba_fee,normalize_package_size,normalize_package_weight
from patch_link_styles import patch_link_styles
from add_link_overlays import add_image_link_overlays

HERE=Path(__file__).resolve().parent

def runtime_paths():
    base=Path.home()/'.cache/codex-runtimes/codex-primary-runtime/dependencies'
    node=Path(os.environ.get('CODEX_NODE_BIN',str(base/'node/bin/node')))
    modules=Path(os.environ.get('CODEX_NODE_MODULES',str(base/'node/node_modules')))
    if not node.is_file() or not (modules/'@oai/artifact-tool').exists():
        raise ValueError('Bundled artifact runtime not found. Use load_workspace_dependencies and set CODEX_NODE_BIN and CODEX_NODE_MODULES to its returned paths.')
    return node,modules

def prepare_data(input_path,allow_missing_images=False):
    data=json.loads(input_path.read_text(encoding='utf-8'))
    if not isinstance(data,dict) or not isinstance(data.get('records',[]),list): raise ValueError('Expected JSON object with records list')
    if not isinstance(data.get('keyword'),str) or not data['keyword'].strip(): raise ValueError('keyword is required')
    if data.get('marketplace','US')!='US': raise ValueError('This skill currently supports amazon.com US only')
    for field in ['variants','notes']:
        if field in data and not isinstance(data[field],list): raise ValueError(f'{field} must be a list')
    remove_legacy_export_data(data)
    data.setdefault('notes',[])
    if data.get('sample'):
        data['notes'].insert(0,'演示数据，仅用于验证格式和功能，不代表真实商品、卖家或销量。')
    if data.get('captured_at'):
        captured=datetime.fromisoformat(str(data['captured_at']).replace('Z','+00:00'))
        if captured.tzinfo is None: raise ValueError('captured_at must include a timezone offset')
    else:
        data['captured_at']=datetime.now(ZoneInfo('Asia/Shanghai')).isoformat()
        data['notes'].append('未提供抓取时间；此处记录的是工作簿生成时间，不作为实时抓取证明。')
    seen=set()
    missing=[]
    schema_headers=json.loads((HERE.parent/'references/schema.json').read_text(encoding='utf-8'))['headers']
    for i,record in enumerate(data.get('records',[]),2):
        if not isinstance(record,dict): raise ValueError(f'Row {i}: expected object')
        asin=str(record.get('ASIN',''))
        if not re.fullmatch(r'[A-Z0-9]{10}',asin): raise ValueError(f'Row {i}: invalid ASIN {asin!r}')
        if asin in seen: raise ValueError(f'Duplicate ASIN: {asin}')
        seen.add(asin)
        provenance=record.get('_field_provenance') or {}
        if not isinstance(provenance,dict): raise ValueError(f'{asin}: _field_provenance must be an object')
        sources=record.setdefault('_sources',{})
        if not isinstance(sources,dict): raise ValueError(f'{asin}: _sources must be an object')
        for field,meta in provenance.items():
            if not isinstance(meta,dict): raise ValueError(f'{asin}: provenance for {field} must be an object')
            source=meta.get('source')
            if source:
                if not isinstance(source,str): source=json.dumps(source,ensure_ascii=False)
                old=sources.get(field)
                if old and old!=source:
                    old=old if isinstance(old,str) else json.dumps(old,ensure_ascii=False)
                    if source not in old: sources[field]=old+'；'+source
                else: sources[field]=source
        # These values are calculated only from source fields, never from captured_at/current time.
        record['FBA配送费']=normalize_fba_fee(record.get('FBA配送费'))
        record['包装尺寸(cm)']=normalize_package_size(record.get('包装尺寸(cm)'))
        record['包装重量(g)']=normalize_package_weight(record.get('包装重量(g)'))
        if record.get('库存状态')=='无购买框': record['价格']=None
        record['_excel_date_values']=expected_date_values(record)
        record['_highlight_fields_resolved']=sorted(highlighted_fields(record,schema_headers))
        image=record.get('_image')
        if image:
            p=Path(image).expanduser()
            if not p.is_absolute(): p=input_path.parent/p
            if not p.is_file():
                if allow_missing_images:
                    record['_image']=None
                    missing.append(asin)
                else: raise ValueError(f'{asin}: missing image file: {p}')
            else: record['_image']=str(p.resolve())
        elif allow_missing_images: missing.append(asin)
        else: raise ValueError(f'{asin}: _image required (use --allow-missing-images only for explicit partial delivery)')
        if record.get('_sales_estimated') and record.get('近一月销量') is not None:
            data['notes'].append(f'{asin}：近一月销量为第三方估算，已用浅橙底标识。')
    if missing: data['notes'].append('主图缺失，部分交付：'+', '.join(missing))
    date_fields=data.get('variant_date_fields',['上架时间'])
    datetime_fields=data.get('variant_datetime_fields',['卖家精灵数据更新时间'])
    if not isinstance(date_fields,list) or not isinstance(datetime_fields,list): raise ValueError('variant_date_fields and variant_datetime_fields must be lists')
    for variant in data.get('variants',[]):
        if not isinstance(variant,dict): raise ValueError('Each variant must be an object')
        serials={field:excel_date_serial(variant.get(field)) for field in date_fields if field in variant}
        serials.update({field:excel_beijing_date_serial(variant.get(field)) for field in datetime_fields if field in variant})
        variant['_excel_date_values']=serials
    return data

def main():
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('input',type=Path)
    parser.add_argument('--output',type=Path)
    parser.add_argument('--preview-dir',type=Path)
    parser.add_argument('--allow-missing-images',action='store_true')
    parser.add_argument('--no-wps-click-layer',action='store_true',help='Keep native images and normal hyperlinks but omit transparent WPS click hot spots')
    args=parser.parse_args()
    input_path=args.input.resolve()
    data=prepare_data(input_path,args.allow_missing_images)
    schema_path=HERE.parent/'references/schema.json'
    schema=json.loads(schema_path.read_text(encoding='utf-8'))
    if len(schema['headers'])!=33: raise ValueError('Schema must contain exactly 33 columns')
    safe_keyword=re.sub(r'[/\\\x00-\x1f]','_',data['keyword'])
    stamp=datetime.now(ZoneInfo('Asia/Shanghai')).strftime('%Y%m%d')
    output=(args.output or Path.cwd()/'outputs'/f'当季热销品竞品数据库--{safe_keyword}{stamp}.xlsx').resolve()
    node,modules=runtime_paths()
    with tempfile.TemporaryDirectory(prefix='codex-competitor-') as td:
        temp=Path(td)
        os.symlink(modules,temp/'node_modules',target_is_directory=True)
        builder=temp/'build_workbook.mjs'
        shutil.copy2(HERE/'build_workbook.mjs',builder)
        prepared=temp/'input.json'; prepared.write_text(json.dumps(data,ensure_ascii=False),encoding='utf-8')
        base=temp/'base.xlsx'
        command=[str(node),str(builder),str(prepared),str(schema_path),str(base)]
        if args.preview_dir: command.append(str(args.preview_dir.resolve()))
        result=subprocess.run(command,text=True,capture_output=True)
        if result.returncode:
            raise RuntimeError('Artifact workbook creation failed:\n'+result.stderr[-6000:]+result.stdout[-2000:])
        output.parent.mkdir(parents=True,exist_ok=True)
        fd,pending_name=tempfile.mkstemp(prefix='.competitor-validated-',suffix='.xlsx',dir=output.parent)
        os.close(fd)
        pending=Path(pending_name)
        try:
            native=temp/'native.xlsx'
            report=finish_workbook(base,native,data,schema,args.allow_missing_images)
            working=native
            link_report=None
            if data.get('records'):
                working=temp/'linked.xlsx'
                link_report=patch_link_styles(native,working,add_asin_links=True)
            overlay_report=None
            if report['embedded_images'] and not args.no_wps_click_layer:
                overlay_report=add_image_link_overlays(working,pending)
            else: shutil.copyfile(working,pending)
            validation=validate_workbook(pending,data,schema,require_click_layer=not args.no_wps_click_layer)
            os.replace(pending,output)
        finally:
            pending.unlink(missing_ok=True)
    print(json.dumps({'output':str(output),**report,'link_style':link_report,'click_layer':overlay_report,'validation':validation},ensure_ascii=False))

if __name__=='__main__':
    try: main()
    except (ValueError,RuntimeError,OSError,KeyError,json.JSONDecodeError) as exc:
        print(f'Error: {exc}',file=sys.stderr)
        sys.exit(1)

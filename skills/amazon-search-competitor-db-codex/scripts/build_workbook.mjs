import fs from 'node:fs/promises';
import path from 'node:path';
import { Workbook, SpreadsheetFile } from '@oai/artifact-tool';

// Called by render_workbook.py from an isolated runtime directory.
const [inputPath, schemaPath, outputPath, previewDir = ''] = process.argv.slice(2);
if (!inputPath || !schemaPath || !outputPath) throw new Error('Usage: build_workbook.mjs input.json schema.json base.xlsx [preview-dir]');
const data = JSON.parse(await fs.readFile(inputPath, 'utf8'));
const schema = JSON.parse(await fs.readFile(schemaPath, 'utf8'));
const headers = schema.headers;
const rows = data.records ?? [];
const wb = Workbook.create();
const detail = wb.worksheets.add('竞品明细');
const variants = wb.worksheets.add('变体明细');
const notes = wb.worksheets.add('说明');
const col = n => { let out=''; for (;n;n=Math.floor((n-1)/26)) out=String.fromCharCode(65+(n-1)%26)+out; return out; };
const literal = v => typeof v === 'string' && /^[=+@]/.test(v) ? `'${v}` : v;
const integerColumns = new Set(schema.integer_columns);
const numericColumns = new Set([...schema.integer_columns, ...schema.price_columns, '评分', '折扣']);
const dateFormat='yyyy年mm月dd日';
const datetimeFormat='yyyy-mm-dd hh:mm';
const visualLength = value => Array.from(String(value??'')).reduce((n,ch)=>n+(ch.codePointAt(0)>0x2ff?2:1),0);
const lineHeight = (value,width) => String(value??'').split('\n').reduce((sum,line)=>sum+Math.max(1,Math.ceil(visualLength(line)/width)),0)*13+10;
const values = rows.map((record,i) => headers.map(header => {
  if (header === '主图') return record._image ? record.ASIN : null;
  if (header === '上架时间' || header === '卖家精灵数据更新时间') return record._excel_date_values?.[header] ?? null;
  const value=record[header] ?? null;
  if (numericColumns.has(header) && value!==null && (typeof value!=='number' || !Number.isFinite(value))) throw new Error(`${record.ASIN}: ${header} must be a number or null`);
  if (integerColumns.has(header) && value!==null && (!Number.isInteger(value)||value<0)) throw new Error(`${record.ASIN}: ${header} must be a nonnegative integer`);
  return literal(value);
}));
detail.getRange('A1:AF1').values=[headers];
if (values.length) detail.getRange(`A2:AF${values.length+1}`).values=values;
const used=detail.getRange(`A1:AF${Math.max(1,rows.length+1)}`);
used.format.font={name:'Arial',size:10,color:'#202938'};
used.format.verticalAlignment='center';
used.format.columnWidth=16;
used.format.rowHeight=76;
used.format.wrapText=true;
detail.showGridLines=false;
detail.tabColor='#203A58';
const headerRange=detail.getRange('A1:AF1');
headerRange.format={fill:'#203A58',font:{name:'Arial',size:10,bold:true,color:'#FFFFFF'},rowHeight:38,wrapText:true,horizontalAlignment:'center',verticalAlignment:'center'};
detail.getRange(`A2:C${Math.max(2,rows.length+1)}`).format.columnWidth=9;
detail.getRange(`D1:D${Math.max(1,rows.length+1)}`).format.columnWidth=15;
detail.getRange(`E1:E${Math.max(1,rows.length+1)}`).format.columnWidth=18;
detail.getRange(`L1:L${Math.max(1,rows.length+1)}`).format.columnWidth=19;
detail.getRange(`X1:X${Math.max(1,rows.length+1)}`).format.columnWidth=24;
detail.getRange(`F1:F${Math.max(1,rows.length+1)}`).format.columnWidth=54;
detail.getRange(`F2:G${Math.max(2,rows.length+1)}`).format.wrapText=true;
detail.getRange(`AA1:AB${Math.max(1,rows.length+1)}`).format.columnWidth=36;
detail.getRange(`AA2:AB${Math.max(2,rows.length+1)}`).format.wrapText=true;
detail.getRange(`AF1:AF${Math.max(1,rows.length+1)}`).format.columnWidth=30;
detail.getRange(`AF2:AF${Math.max(2,rows.length+1)}`).format.wrapText=true;
for (let i=0;i<headers.length;i++) {
  const h=headers[i], end=Math.max(2,rows.length+1), r=detail.getRange(`${col(i+1)}2:${col(i+1)}${end}`);
  if(integerColumns.has(h)) r.setNumberFormat('0');
  if(schema.price_columns.includes(h)) r.setNumberFormat('$#,##0.00');
  if(h==='评分') r.setNumberFormat('0.0');
  if(h==='折扣') r.setNumberFormat('0%');
  if(h==='上架时间') r.setNumberFormat(dateFormat);
  if(h==='卖家精灵数据更新时间') r.setNumberFormat(datetimeFormat);
  if(['ASIN','父ASIN','卖家ID'].includes(h)) r.setNumberFormat('@');
  if(numericColumns.has(h)) r.format.horizontalAlignment='right';
}
for(let i=0;i<rows.length;i++) {
  const row=i+2, r=rows[i];
  if(typeof r['价格']==='number'&&typeof r['原价/参考价']==='number'&&r['原价/参考价']>r['价格']) detail.getRange(`J${row}`).formulas=[[`=1-H${row}/I${row}`]];
  else detail.getRange(`J${row}`).values=[[null]];
  for(const field of r._highlight_fields_resolved??[]) {
    const index=headers.indexOf(field);
    if(index>=0) detail.getRange(`${col(index+1)}${row}`).format.fill='#FFF2CC';
  }
  if(r._sales_estimated) detail.getRange(`K${row}`).format.fill='#FCE4D6';
  const longest=Math.max(76,...[['标题',54],['品牌',16],['公司主体',36],['公司地址',36]].map(([field,width])=>lineHeight(r[field],width)));
  detail.getRange(`A${row}:AF${row}`).format.rowHeight=Math.min(409.5,longest);
}
detail.freezePanes.freezeRows(1);
detail.freezePanes.freezeColumns(5);
// Native worksheet autoFilter is added after export by the OOXML finishing step.
const variantRows=data.variants ?? [];
const variantHeaders=[...new Set(['父ASIN','ASIN',...variantRows.flatMap(r=>Object.keys(r).filter(k=>!k.startsWith('_')))])];
variants.getRangeByIndexes(0,0,1,variantHeaders.length).values=[variantHeaders];
if(variantRows.length) variants.getRangeByIndexes(1,0,variantRows.length,variantHeaders.length).values=variantRows.map(r=>variantHeaders.map(k=>Object.hasOwn(r._excel_date_values??{},k)?r._excel_date_values[k]:literal(r[k]??null)));
const vr=variants.getRangeByIndexes(0,0,Math.max(1,variantRows.length+1),variantHeaders.length);
vr.format={font:{name:'Arial',size:10},verticalAlignment:'center',columnWidth:24,rowHeight:24,wrapText:true};
variants.getRangeByIndexes(0,0,1,variantHeaders.length).format={fill:'#203A58',font:{name:'Arial',size:10,bold:true,color:'#FFFFFF'},horizontalAlignment:'center',rowHeight:30};
variants.showGridLines=false;
variants.freezePanes.freezeRows(1);
for(let i=0;i<variantRows.length;i++) {
  const record=variantRows[i],row=i+2;
  for(let j=0;j<variantHeaders.length;j++) {
    const header=variantHeaders[j];
    if(Object.hasOwn(record._excel_date_values??{},header)) variants.getRange(`${col(j+1)}${row}`).setNumberFormat((data.variant_datetime_fields??['卖家精灵数据更新时间']).includes(header)?datetimeFormat:dateFormat);
    if(['ASIN','父ASIN'].includes(header)) variants.getRange(`${col(j+1)}${row}`).setNumberFormat('@');
  }
  variants.getRangeByIndexes(row-1,0,1,variantHeaders.length).format.rowHeight=Math.min(409.5,Math.max(24,...variantHeaders.map(h=>lineHeight(record[h],24))));
}
const captured=data.captured_at ? new Date(data.captured_at) : new Date();
const beijing=Number.isNaN(captured.getTime()) ? String(data.captured_at) : new Intl.DateTimeFormat('sv-SE',{timeZone:'Asia/Shanghai',year:'numeric',month:'2-digit',day:'2-digit',hour:'2-digit',minute:'2-digit',second:'2-digit',hour12:false}).format(captured);
const notesRows=[
 ['项目','内容',''],['搜索词',data.keyword??'',''],['站点',data.marketplace??'US',''],['抓取时间',`${beijing}（北京时间）`,''],
 ['配送邮编',data.delivery_postcode??'未核实',''],['商品数量',rows.length,''],
 ['广告剔除规则','只统计已确认的自然搜索卡片；剔除 Sponsored、广告组件；按 ASIN 去重。实际覆盖缺口见下方说明。',''],
 ['近一月销量口径','页面 bought in past month 的数字可能是下限，原文保存在来源说明；第三方估算填浅橙底。',''],
 ['上架时间口径','历史最早记录只作近似，不等同于真实上架时间；近似项逐条注明。',''],
 ['卖家精灵点数',data.points_used??'未核实',''],
 ['来源颜色','浅黄表示第三方公司资料；浅橙表示估算销量。来源和核实状态逐字段列示。',''],
 ['更新时间口径','只取真实源更新时间，按北京时间显示至分钟；数值保留原时间精度，不用本次查询时间代填。',''],
];
const appendLong=(asin,label,value)=>{
  const text=typeof value==='string'?value:JSON.stringify(value);
  const chars=Array.from(text??'');
  for(let offset=0;offset<Math.max(1,chars.length);offset+=1000) notesRows.push([asin,offset?`${label}（续${Math.floor(offset/1000)+1}）`:label,chars.slice(offset,offset+1000).join('')]);
};
for(const note of data.notes??[]) appendLong('说明','',String(note));
notesRows.push(['字段来源','',''],['ASIN','字段','来源及完整说明']);
const sourceLabels={amazon_direct:'Amazon直接披露',third_party_database:'第三方数据库'};
for(const record of rows) {
  for(const [field,source] of Object.entries(record._sources??{})) appendLong(record.ASIN,field,source);
  for(const [field,meta] of Object.entries(record._field_provenance??{})) {
    const label=sourceLabels[meta.source_type]??String(meta.source_type??'来源类型未标记');
    appendLong(record.ASIN,`${field}来源标记`,`${label}${meta.verified_on_amazon===true?'；已在Amazon核实':meta.verified_on_amazon===false?'；未在Amazon核实':''}${meta.source?'；'+(typeof meta.source==='string'?meta.source:JSON.stringify(meta.source)):''}`);
  }
  if(record._sales_display) appendLong(record.ASIN,'近一月销量页面原文',`${record._sales_display}${record._sales_lower_bound?'；所填数字为页面披露下限，不是精确销量。':''}`);
  if(record._listing_date_approx) appendLong(record.ASIN,'上架时间','近似：使用历史最早记录，不是确证上架日期。');
  if(record._sellersprite_updated_at_ms!==null&&record._sellersprite_updated_at_ms!==undefined) appendLong(record.ASIN,'更新时间原始毫秒',String(record._sellersprite_updated_at_ms));
  else if(record['卖家精灵数据更新时间']) appendLong(record.ASIN,'更新时间原始值',String(record['卖家精灵数据更新时间']));
  for(const field of ['标题','公司主体','公司地址']) if(typeof record[field]==='string'&&Array.from(record[field]).length>160) appendLong(record.ASIN,`${field}完整内容`,record[field]);
}
for(const record of variantRows) {
  for(const [field,value] of Object.entries(record)) if(!field.startsWith('_')&&typeof value==='string'&&Array.from(value).length>120) appendLong(record.ASIN??record['父ASIN']??'变体',`变体${field}完整内容`,value);
  for(const [field,source] of Object.entries(record._sources??{})) appendLong(record.ASIN??'变体',`变体${field}来源`,source);
}
notes.getRange(`A1:C${notesRows.length}`).values=notesRows.map(r=>r.map(literal));
notes.getRange(`A1:C${notesRows.length}`).format={font:{name:'Arial',size:10},verticalAlignment:'center',rowHeight:32,wrapText:true};
notes.getRange(`A1:A${notesRows.length}`).format.columnWidth=20;
notes.getRange(`B1:B${notesRows.length}`).format.columnWidth=44;
notes.getRange(`C1:C${notesRows.length}`).format.columnWidth=90;
notes.getRange('A1:C1').format={fill:'#203A58',font:{name:'Arial',size:10,bold:true,color:'#FFFFFF'}};
for(let i=1;i<notesRows.length;i++) notes.getRange(`A${i+1}:C${i+1}`).format.rowHeight=Math.min(409.5,Math.max(32,lineHeight(notesRows[i][0],20),lineHeight(notesRows[i][1],44),lineHeight(notesRows[i][2],90)));
notes.showGridLines=false;
wb.recalculate();
const inspected=await wb.inspect({kind:'table',range:`竞品明细!A1:M${Math.min(4,rows.length+1)}`,include:'values,formulas',tableMaxRows:4,tableMaxCols:13,maxChars:2500});
const errors=await wb.inspect({kind:'match',searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!',options:{useRegex:true,maxResults:20},maxChars:1000});
await fs.mkdir(path.dirname(outputPath),{recursive:true});
await (await SpreadsheetFile.exportXlsx(wb)).save(outputPath);
if(previewDir) {
  await fs.mkdir(previewDir,{recursive:true});
  // The PNG renderer does not recognize Chinese date literals. This preview-only
  // format change happens after export; the saved XLSX keeps yyyy年mm月dd日.
  if(rows.length) detail.getRange(`L2:L${rows.length+1}`).setNumberFormat('yyyy-mm-dd');
  if(variantRows.length) for(const field of data.variant_date_fields??['上架时间']) {
    const index=variantHeaders.indexOf(field);
    if(index>=0) variants.getRange(`${col(index+1)}2:${col(index+1)}${variantRows.length+1}`).setNumberFormat('yyyy-mm-dd');
  }
  // Preview-only drawings let the renderer show thumbnails; base/final XLSX contain no such drawings.
  for(let i=0;i<Math.min(rows.length,4);i++) if(rows[i]._image) {
    const bytes=await fs.readFile(rows[i]._image);
    const mime=bytes[0]===0x89?'image/png':'image/jpeg';
    detail.getRange(`E${i+2}`).values=[[null]];
    detail.images.add({dataUrl:`data:${mime};base64,${bytes.toString('base64')}`,anchor:{from:{row:i+1,col:4,rowOffsetPx:3,colOffsetPx:3},extent:{widthPx:104,heightPx:92}}});
  }
  for(const [sheetName,range,name] of [['竞品明细',`A1:M${Math.min(5,rows.length+1)}`,'竞品明细'],['竞品明细',`N1:AF${Math.min(5,rows.length+1)}`,'竞品明细右侧'],['变体明细',`A1:${col(Math.min(variantHeaders.length,6))}${Math.min(5,variantRows.length+1)}`,'变体明细'],['说明',`A1:C${Math.min(13,notesRows.length)}`,'说明']]) {
    const blob=await wb.render({sheetName,range,scale:1,format:'png'});
    await fs.writeFile(path.join(previewDir,`${name}.png`),new Uint8Array(await blob.arrayBuffer()));
  }
  await fs.writeFile(path.join(previewDir,'artifact-inspection.json'),JSON.stringify({table:inspected.ndjson,errors:errors.ndjson},null,2));
}
console.log(JSON.stringify({base:outputPath,records:rows.length,columns:headers.length}));

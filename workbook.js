import {CATEGORIES,fields,contentHash,version,validateExpense,decorate,uid,validDate,validateSettings,isInternal} from './model.js';
// Master input sheet. Finance may edit these columns directly in Excel.
const HEADERS=['Expense ID','Category','Cost type','Date','Vendor / team','Description','Quantity','Duration','Unit','Unit price','Base amount','Tax treatment','Tax %','Tax adjustment','Billable amount','Payable amount','Paid amount','Reference','Notes','Supporting files (app managed)','Document status (app managed)'];
const KEYS=['id','category','costType','date','vendor','description','qty','duration','unit','price','baseAmount','taxType','taxPct','taxAmount','billable','payable','paid','reference','notes'];
const NUMERIC=['qty','duration','price','baseAmount','taxAmount','billable','payable','paid'];
const col=k=>KEYS.indexOf(k)+1;
// Category detail sheets follow the layout of the 2026 recap sent to the promoter.
const REPORT=['Cost type','Date','Vendor / Team','Description','Qty','Metrics','Unit','Base price','Base amount','Tax type','Tax %','Tax / gross-up','Total incl. tax','Billable cost','Source / Reference','Notes'];
const REPORT_WIDTHS=[20,12,26,48,9,10,10,16,18,16,9,16,18,18,44,32];
const HEAD=4;
const AMOUNT='#,##0.00;(#,##0.00);"–"';
const PCT='0.0#%';
const quote=name=>`'${name.replace(/'/g,"''")}'`;
function cellValue(c){
 const v=c.value;if(v===null||v===undefined)return '';if(v instanceof Date)return v.toISOString().slice(0,10);
 if(typeof v==='object'){if('formula'in v||'sharedFormula'in v)throw Error('Use values, not formulas, in the master input sheets.');if(v.richText)return v.richText.map(x=>x.text).join('');if(v.text!==undefined)return v.text;throw Error('Unsupported input cell')}
 return v;
}
function paint(row){row.font={bold:true,color:{argb:'FFFFFFFF'}};row.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FF263D30'}};row.height=30}
function style(s,headers){s.addRow(headers);paint(s.getRow(1));s.columns.forEach(c=>c.width=22);s.views=[{state:'frozen',ySplit:1}];s.autoFilter={from:{row:1,column:1},to:{row:1,column:headers.length}}}
function formula(cell,text,result){cell.value={formula:text,result}}
// Print each report sheet landscape, one page wide, so a PDF or printout keeps all columns together.
function printable(s){s.pageSetup={orientation:'landscape',paperSize:9,fitToPage:true,fitToWidth:1,fitToHeight:0}}
function docNames(e,data){return e.documents.map(id=>data.documents.find(d=>d.id===id)?.name).filter(Boolean)}
function metaRows(wb,data){
 const s=wb.addWorksheet('_12W App Data',{state:'veryHidden'});
 s.addRow(['12WIRED_FINANCE_MASTER',1]);s.addRow(['project',JSON.stringify(data.project)]);
 for(const d of data.documents)s.addRow(['document',JSON.stringify(d)]);
 for(const e of data.expenses)s.addRow(['expense',JSON.stringify({id:e.id,documents:e.documents,documentChecked:e.documentChecked,source:e.source,sourceKey:e.sourceKey,contentHash:e.contentHash,updated:e.updated})]);
 for(const h of data.history.slice(0,100))s.addRow(['history',JSON.stringify(h)]);
}
// One category detail sheet: title, comparison line, then the cost rows and a total row.
function categorySheet(s,cat,es,data){
 const p=data.project;REPORT_WIDTHS.forEach((w,i)=>s.getColumn(i+1).width=w);
 s.getCell('A1').value=cat.toUpperCase();s.getCell('A1').font={bold:true,size:14,color:{argb:'FF263D30'}};
 s.getRow(HEAD).values=REPORT;paint(s.getRow(HEAD));s.views=[{state:'frozen',ySplit:HEAD}];s.autoFilter={from:{row:HEAD,column:1},to:{row:HEAD,column:REPORT.length}};
 let r=HEAD;const sum={base:0,tax:0,total:0,bill:0};
 for(const e of es){
  r++;const row=s.getRow(r);const blank=isInternal(e)&&e.billable===null;
  const src=[e.reference,...docNames(e,data)].filter(Boolean).join(' | ');
  row.values=[e.costType||'',e.date||'',e.vendor,e.description,e.qty,e.duration,e.unit,e.price,e.baseAmount,e.taxType,e.taxPct===null||e.taxPct===undefined?null:e.taxPct/100,e.taxAmount,null,e.billable,src,e.notes||(blank?'To be filled by the project head':'')];
  if(e.baseAmount!==null||e.taxAmount!==null)formula(row.getCell(13),`I${r}+L${r}`,(e.baseAmount||0)+(e.taxAmount||0));
  for(const c of [5,6,8,9,12,13,14])row.getCell(c).numFmt=AMOUNT;row.getCell(11).numFmt=PCT;
  sum.base+=e.baseAmount||0;sum.tax+=e.taxAmount||0;sum.total+=(e.baseAmount||0)+(e.taxAmount||0);sum.bill+=e.billable||0;
 }
 const tot=r+1,trow=s.getRow(tot);trow.getCell(4).value='CATEGORY TOTAL';trow.font={bold:true};
 for(const [c,v] of [[9,sum.base],[12,sum.tax],[13,sum.total],[14,sum.bill]]){const L=s.getColumn(c).letter;if(es.length)formula(trow.getCell(c),`SUM(${L}${HEAD+1}:${L}${r})`,v);else trow.getCell(c).value=0;trow.getCell(c).numFmt=AMOUNT}
 const prev=p.previous[cat];
 // Totals line above the matching table columns: tax, total incl. tax, billable, then the comparison.
 for(const [c,label] of [['L','Tax / gross-up'],['M','Total incl. tax'],['N','Billable cost'],['O',p.previousLabel],['P','Variance']]){s.getCell(c+'2').value=label;s.getCell(c+'2').font={bold:true}}
 formula(s.getCell('L3'),`L${tot}`,sum.tax);formula(s.getCell('M3'),`M${tot}`,sum.total);formula(s.getCell('N3'),`N${tot}`,sum.bill);s.getCell('O3').value=prev;formula(s.getCell('P3'),'IF(O3="","n.a.",N3-O3)',prev===null?'n.a.':sum.bill-prev);
 for(const c of ['L3','M3','N3','O3','P3']){s.getCell(c).numFmt=AMOUNT;s.getCell(c).alignment={horizontal:'right'}}
 return tot;
}
export async function buildWorkbook(data,{final=false,master=true}={}){
 decorate(data);for(const e of data.expenses){e.contentHash=await contentHash(e);e.version=await version(e)}
 const wb=new globalThis.ExcelJS.Workbook();wb.creator='12WIRED Finance';wb.calcProperties.fullCalcOnLoad=true;
 const p=data.project,es=data.expenses;
 const summary=wb.addWorksheet('Summary');
 const categories=Object.fromEntries(CATEGORIES.map(c=>[c,wb.addWorksheet(c)]));
 const register=wb.addWorksheet('Source Register');
 // The promoter report holds only Summary, category sheets and Source Register. Vendor payables,
 // payments, settings and app data stay in the master.
 const inputs=master?wb.addWorksheet('Expenses'):null,settings=master?wb.addWorksheet('Project Settings'):null,guide=master?wb.addWorksheet('How to use'):null;
 if(inputs){
  style(inputs,HEADERS);inputs.getColumn(col('id')).width=38;inputs.getColumn(col('description')).width=48;inputs.getColumn(HEADERS.length-1).width=48;inputs.getColumn(HEADERS.length).width=45;
  for(const e of es)inputs.addRow([...KEYS.map(k=>k==='taxPct'?(e.taxPct===null||e.taxPct===undefined?null:e.taxPct/100):e[k]??null),docNames(e,data).join('; '),e.issues.join('; ')||(isInternal(e)?'12WIRED row filled':'Supported and checked')]);
  for(const k of NUMERIC)inputs.getColumn(col(k)).numFmt=AMOUNT;inputs.getColumn(col('taxPct')).numFmt=PCT;
 }
 const totals=Object.fromEntries(CATEGORIES.map(c=>[c,categorySheet(categories[c],c,es.filter(e=>e.category===c),data)]));
 // The budget is internal (owner decision, 26 Oct 2026): it appears in the master and the app, never in promoter reports.
 const budgetCols=master;
 summary.columns=budgetCols?[{width:30},{width:24},{width:24},{width:24},{width:18},{width:24},{width:24},{width:42}]:[{width:30},{width:24},{width:24},{width:24},{width:18},{width:42},{width:24}];
 summary.getCell('A1').value=p.name+' — COST RECAP & QUOTATION';summary.getCell('A1').font={bold:true,size:16,color:{argb:'FF263D30'}};
 summary.getCell('A2').value=(master?'WORKING MASTER':final?'FINAL':'DRAFT')+' · Artist hospitality · IDR';
 summary.getRow(4).values=['Category',p.previousLabel,'Current billable','Variance','Variance %',...(budgetCols?['Approved budget','Budget variance']:[]),'Explanation'];paint(summary.getRow(4));summary.getRow(4).height=undefined;
 const catL=inputs?.getColumn(col('category')).letter,billL=inputs?.getColumn(col('billable')).letter;
 let subtotal=0;
 for(let i=0;i<CATEGORIES.length;i++){
  const c=CATEGORIES[i],row=5+i,bill=es.filter(e=>e.category===c).reduce((a,e)=>a+(e.billable||0),0),prev=p.previous[c],budget=p.budgets[c];subtotal+=bill;
  summary.getCell(`A${row}`).value=c;summary.getCell(`B${row}`).value=prev;
  formula(summary.getCell(`C${row}`),master?`SUMIF(Expenses!${catL}2:${catL}10001,A${row},Expenses!${billL}2:${billL}10001)`:`${quote(c)}!N${totals[c]}`,bill);
  formula(summary.getCell(`D${row}`),`IF(B${row}="","n.a.",C${row}-B${row})`,prev===null?'n.a.':bill-prev);
  formula(summary.getCell(`E${row}`),`IF(OR(B${row}="",B${row}=0),"n.a.",D${row}/B${row})`,!prev?'n.a.':(bill-prev)/prev);
  if(budgetCols){summary.getCell(`F${row}`).value=budget;formula(summary.getCell(`G${row}`),`IF(F${row}="","n.a.",C${row}-F${row})`,budget===null?'n.a.':bill-budget);summary.getCell(`H${row}`).value=p.explanations[c]}else summary.getCell(`F${row}`).value=p.explanations[c];
 }
 summary.getCell('A11').value='COST SUBTOTAL';formula(summary.getCell('C11'),'SUM(C5:C10)',subtotal);
 const fee=Math.round(subtotal*p.feePct)/100,vat=Math.round((subtotal+fee)*p.vatPct)/100,total=subtotal+fee+vat;
 for(const [row,label,text,value] of [[14,'Billable cost subtotal','C11',subtotal],[15,'Agency service fee','ROUND(C14*B15,2)',fee],[16,'VAT on costs + agency fee','ROUND(SUM(C14:C15)*B16,2)',vat],[17,'Grand total','SUM(C14:C16)',total],[18,'Less: promoter advances','-B18',-p.advance],[19,'TOTAL DUE','SUM(C17:C18)',total-p.advance]]){summary.getCell(`A${row}`).value=label;formula(summary.getCell(`C${row}`),text,value)}
 summary.getCell('B15').value=p.feePct/100;summary.getCell('B16').value=p.vatPct/100;summary.getCell('B18').value=p.advance;summary.getCell('B15').numFmt=summary.getCell('B16').numFmt='#,##0.0%';summary.getCell('B18').numFmt=AMOUNT;
 for(const c of budgetCols?['B','C','D','F','G']:['B','C','D'])for(let row=5;row<=11;row++)summary.getCell(`${c}${row}`).numFmt=AMOUNT;
 for(let row=5;row<=10;row++)summary.getCell(`E${row}`).numFmt='#,##0.0%;(#,##0.0%)';for(let row=14;row<=19;row++)summary.getCell(`C${row}`).numFmt=AMOUNT;
 if(master||!final){
  summary.getCell('F14').value='Expense records';summary.getCell('G14').value=es.length;summary.getCell('G14').numFmt='#,##0';
  summary.getCell('F15').value='Rows needing attention';summary.getCell('G15').value=es.filter(e=>e.issues.length).length;summary.getCell('G15').numFmt='#,##0';
  summary.getCell('F16').value='Settings confirmed';summary.getCell('G16').value=p.settingsConfirmed?'Yes':'No';
 }
 if(master){summary.getCell('A22').value='Edit Expenses and Project Settings, then save Excel and click Refresh from Dropbox in the app.';summary.getCell('A23').value='Category detail sheets and document checks are refreshed when the app saves the master.'}
 style(register,['No.','Document','Type','Expense references','Dropbox path']);register.getColumn(2).width=48;register.getColumn(4).width=55;register.getColumn(5).width=70;
 const seen=new Set();for(const d of data.documents){if(seen.has(d.sha256)||!es.some(e=>e.documents.includes(d.id)))continue;seen.add(d.sha256);register.addRow([seen.size,d.name,d.kind,es.filter(e=>e.documents.includes(d.id)).map(e=>e.reference||e.description).join('; '),'/Apps/12WIRED Finance Pilot'+d.path])}
 register.getColumn(1).numFmt='#,##0';
 for(const s of [summary,register,...Object.values(categories)])printable(s);
 if(!master)return new Uint8Array(await wb.xlsx.writeBuffer());
 style(settings,['Setting','Value','Instructions']);settings.getColumn(1).width=44;settings.getColumn(3).width=70;
 for(const [key,label] of [['name','Project name'],['eventEnd','Festival end date (optional, YYYY-MM-DD)'],['feePct','Agency service fee (%)'],['vatPct','VAT (%)'],['advance','Promoter advances (IDR)'],['previousLabel','Previous-project comparison label']])settings.addRow([key,p[key],label]);
 for(const c of CATEGORIES){settings.addRow(['budget:'+c,p.budgets[c],'Approved budget; blank means not supplied']);settings.addRow(['previous:'+c,p.previous[c],'Previous project total; blank means not supplied']);settings.addRow(['explanation:'+c,p.explanations[c],'Variance explanation'])}
 settings.getColumn(2).numFmt=AMOUNT;settings.getColumn(2).width=36;
 guide.getColumn(1).width=120;
 for(const line of ['12WIRED FINANCE PILOT','The master workbook is the project record. Supporting files are stored separately in this app’s Dropbox folder.','Edit the Expenses sheet for daily entries. Keep its headers, existing Expense IDs, and project/category names intact.','For new Excel rows, leave Expense ID blank; the app assigns one on the next save. Enter dates as YYYY-MM-DD.','Use numbers, not formulas, in Expenses and Project Settings. Category names must match the six report categories.','Cost type says how the cost was paid: Accounts payable, Company card (add the cardholder, e.g. Company card NY), Cash advance, LO payment recap, Direct receipt, Inventory, or 12WIRED.','Tax %: enter as a percentage, e.g. 2.5%. Tax adjustment is the amount in rupiah. For WHT gross-up, tax = base ÷ (1 − rate) − base, rounded to whole rupiah.','12WIRED rows are 12WIRED’s own manpower. Finance may leave their amounts blank; the project head fills them. They need no invoice, but the final report waits until each has a billable amount.','Edit only Expenses and Project Settings. Summary and category sheets are generated reports.','After editing Excel, save it and wait for Dropbox sync, then use Refresh from Dropbox in the app.','Changes to expense details clear their document check. Recheck the invoice or quotation in the app.','Changes to Project Settings clear finance’s settings confirmation. Confirm the settings again in the app.','If another copy changes during a save, the app stops. Refresh and review before retrying; it never silently overwrites.','Files uploaded from your phone are saved in Supporting Documents and referenced in this workbook.','Final reports require a cost type and a checked invoice or vendor quotation for every vendor cost, a filled amount on every 12WIRED row, and confirmed billing settings.','The promoter report contains Summary, category sheets and Source Register only. Budgets, payables, payments and settings stay in this master.','Submission timing is supervised manually by the project head.','Pilot limit: 2,000 expenses per project, 15 MB per attachment, and 10 MB for an edited master workbook.','Anyone with the address can open the app page, but each person sees only the files in the Dropbox account they sign in with. It requires internet access; it does not queue offline changes.'])guide.addRow([line]);
 metaRows(wb,data);return new Uint8Array(await wb.xlsx.writeBuffer());
}
export async function readWorkbook(bytes){
 if(bytes.byteLength>10*1024*1024)throw Error('The master workbook exceeds the 10 MB pilot limit');
 const wb=new globalThis.ExcelJS.Workbook();await wb.xlsx.load(bytes);const meta=wb.getWorksheet('_12W App Data');
 if(!meta||meta.getCell('A1').value!=='12WIRED_FINANCE_MASTER'||meta.getCell('B1').value!==1)throw Error('This file is not a supported 12WIRED master workbook');
 const allowed=['Summary',...CATEGORIES,'Source Register','Expenses','Project Settings','How to use','_12W App Data'];if(wb.worksheets.some(s=>!allowed.includes(s.name)))throw Error('This master contains extra sheets. Keep additional analysis in a separate workbook before using app saves.');
 const data={project:null,expenses:[],documents:[],history:[]},old=new Map();
 meta.eachRow((r,n)=>{if(n===1)return;const kind=r.getCell(1).value;const value=JSON.parse(String(r.getCell(2).value));if(kind==='project')data.project=value;if(kind==='document')data.documents.push(value);if(kind==='expense')old.set(value.id,value);if(kind==='history')data.history.push(value)});
 if(!data.project?.id)throw Error('Master project settings are missing');
 // Masters saved by the short-lived no-budget version (26 Oct 2026) have no budgets: start them blank.
 if(!data.project.budgets)data.project.budgets=Object.fromEntries(CATEGORIES.map(c=>[c,null]));
 const settings=wb.getWorksheet('Project Settings');if(!settings)throw Error('Project Settings sheet is missing');const settingsInput=structuredClone(data.project);
 settings.eachRow((r,n)=>{if(n===1)return;const key=String(cellValue(r.getCell(1))),value=cellValue(r.getCell(2));if(['name','eventEnd','feePct','vatPct','advance','previousLabel'].includes(key))settingsInput[key]=value;else if(key.startsWith('budget:'))settingsInput.budgets[key.slice(7)]=value;else if(key.startsWith('previous:'))settingsInput.previous[key.slice(9)]=value;else if(key.startsWith('explanation:'))settingsInput.explanations[key.slice(12)]=value});
 const updated=validateSettings(settingsInput,data.project);if(JSON.stringify(updated)!==JSON.stringify(data.project)){updated.settingsConfirmed=false;data.settingsEditedInExcel=true}data.project=updated;
 const s=wb.getWorksheet('Expenses');if(!s)throw Error('Expenses sheet is missing');if(s.rowCount>2001)throw Error('Pilot limit: 2,000 expenses per project');
 for(let i=0;i<HEADERS.length;i++)if(s.getRow(1).getCell(i+1).value!==HEADERS[i])throw Error('Keep the master Expenses column headers unchanged');
 const ids=new Set();
 for(let row=2;row<=s.rowCount;row++){
  const r=s.getRow(row);if(!KEYS.some((_,i)=>cellValue(r.getCell(i+1))!==''))continue;
  // Excel holds Tax % as a true percentage (0.025); the app holds percentage points (2.5).
  const raw=Object.fromEntries(KEYS.map((key,i)=>{let v=cellValue(r.getCell(i+1));if(key==='taxPct'&&typeof v==='number')v=Math.round(v*1e8)/1e6;return [key,v]}));
  let prior=old.get(raw.id);if(raw.id&&!prior)throw Error(`Unknown Expense ID on row ${row}. Leave IDs blank for new entries.`);
  const e=validateExpense(raw,data.project.id,prior||null);if(!prior)e.id='excel-'+(await contentHash(e)).slice(0,20)+'-'+row;
  if(ids.has(e.id))throw Error('Duplicate Expense ID in the master');ids.add(e.id);
  e.contentHash=await contentHash(e);e.documentChecked=!!prior?.documentChecked&&prior.contentHash===e.contentHash;e.version=await version(e);if(prior&&prior.contentHash!==e.contentHash)e.editedInExcel=true;
  data.expenses.push(e);
 }
 return decorate(data);
}
export async function readUpload(name,bytes){
 if(bytes.byteLength>15*1024*1024)throw Error('Maximum spreadsheet size is 15 MB');if(name.toLowerCase().endsWith('.csv'))return {CSV:parseCSV(new TextDecoder().decode(bytes))};if(!name.toLowerCase().endsWith('.xlsx'))throw Error('Choose XLSX or CSV');
 const wb=new globalThis.ExcelJS.Workbook();await wb.xlsx.load(bytes);const out={};
 for(const s of wb.worksheets){if(s.rowCount>10000||s.columnCount>250)throw Error('Import limit: 10,000 rows and 250 columns per sheet');out[s.name]=[];for(let i=1;i<=s.rowCount;i++){const row=[];for(let j=1;j<=s.columnCount;j++){let v=s.getRow(i).getCell(j).value;if(v instanceof Date)v=v.toISOString().slice(0,10);else if(v&&typeof v==='object'){if('formula'in v||'sharedFormula'in v){if(v.result===undefined)throw Error('Recalculate and save the source workbook in Excel before importing');v=v.result}else if(v.richText)v=v.richText.map(x=>x.text).join('');else v=v.text??''}row.push(v??'')}out[s.name].push(row)}}
 return out;
}
export function parseCSV(text){
 text=text.replace(/^﻿/,'');const first=text.split(/\r?\n/)[0];const sep=[';',',','\t'].sort((a,b)=>first.split(b).length-first.split(a).length)[0];const rows=[];let row=[],value='',quoted=false;
 for(let i=0;i<text.length;i++){const c=text[i];if(c==='"'){if(quoted&&text[i+1]==='"'){value+='"';i++}else quoted=!quoted}else if(c===sep&&!quoted){row.push(value);value=''}else if((c==='\n'||c==='\r')&&!quoted){if(c==='\r'&&text[i+1]==='\n')i++;row.push(value);rows.push(row);row=[];value=''}else value+=c;if(rows.length>10000||row.length>250)throw Error('Import exceeds the row or column limit')}
 if(quoted)throw Error('Unclosed quote in CSV');if(value||row.length){row.push(value);rows.push(row)}return rows;
}

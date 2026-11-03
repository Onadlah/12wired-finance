import * as dbx from './dropbox.js';
import {uid,now,clone,projectDefaults,validateExpense,validateSettings,decorate,finalReady,hash,version,contentHash} from './model.js';
import {buildWorkbook,readWorkbook,readUpload} from './workbook.js';
// A project is one event's FINANCE folder: /PROJECTS/<year>/<event>/<nn FINANCE>/Master Finance.xlsx.
// The app refers to it by the master file's Dropbox ID ("key"); the workbook keeps its own project ID.
let projects=[],cache=new Map(),imports=new Map(),created=new Map(),accessCache=new Map(),writing=false;
const MAX=2000,ACCESS_TTL=5*60*1000;
const bytes64=s=>Uint8Array.from(atob(s),c=>c.charCodeAt(0));
const masterPath=p=>p.path+'/'+dbx.MASTER;
function find(key){const p=projects.find(p=>p.id===key);if(!p)throw Error('Project not found. Refresh the project list.');return p}
async function access(p,fresh=false){const c=accessCache.get(p.id);if(!fresh&&c&&Date.now()-c.at<ACCESS_TTL)return c.value;const value=await dbx.checkAccess(p.path);accessCache.set(p.id,{at:Date.now(),value});return value}
async function load(key){const p=find(key),{bytes,meta}=await dbx.download(masterPath(p));const data=await readWorkbook(bytes);Object.assign(data,{revision:meta.rev,path:p.path,key,event:p.event,year:p.year,access:await access(p)});const who=await dbx.whoAmI();data.account={name:who.name,email:who.email};cache.set(key,data);return clone(data)}
async function unchanged(key,expected){const data=await load(key);if(!expected||data.revision!==expected)throw Error('This master changed in Dropbox. Your changes were not saved. Close this form, refresh from Dropbox and review the latest values before retrying.');return data}
async function save(data,message){if(data.expenses.length>MAX)throw Error('Pilot limit: 2,000 expenses per project');data.history.unshift({at:now(),message});data.history=data.history.slice(0,100);const bytes=await buildWorkbook(data);const meta=await dbx.upload(masterPath(data),bytes,data.revision);data.revision=meta.rev;cache.set(data.key,clone(decorate(data)));return data}
function expense(data,id,v){const e=data.expenses.find(e=>e.id===id);if(!e||e.version!==v)throw Error('This expense changed. Refresh and reopen it before saving.');return e}
const ownDocument=(data,d)=>{try{return dbx.documentPath(data.path,d.path)}catch{throw Error('Supporting document '+d.name+' is outside this event’s FINANCE folder. Reattach it.')}};
export async function api(route,r){
 if(route==='projects'){const listed=await dbx.listProjects();for(const id of created.keys())if(listed.some(p=>p.id===id))created.delete(id);projects=[...listed,...created.values()];return projects.map(p=>({id:p.id,name:p.name,demo:false}))}
 if(route==='events')return dbx.listEvents();
 if(route.startsWith('state?'))return load(new URLSearchParams(route.split('?')[1]).get('project'));
 if(writing)throw Error('A save is already in progress. Please wait.');writing=true;
 try{return await action(route,r)}finally{writing=false}
}
async function action(route,r){
 if(route==='project/new'){
  const event=String(r.event||'');if(!dbx.isEventPath(event))throw Error('Choose an event folder');
  const fin=await dbx.findFinanceFolder(event);
  if(!fin)throw Error('This event folder has no FINANCE folder. Ask the CEO to add one (for example 09 FINANCE) and restrict it to the finance team, then try again.');
  if(fin.noAccess)throw Error('You are not on the access list of this event’s FINANCE folder.');
  const check=await dbx.checkAccess(fin.path);
  if(check.restricted===false)throw Error('This FINANCE folder is open to everyone who can open PROJECTS. Ask the CEO to restrict it to the finance team, then try again.');
  if(check.restricted!==true)throw Error('Could not confirm that this FINANCE folder is restricted. '+(check.error||''));
  const [, ,year,eventName]=event.split('/'),name=String(r.name||'').trim()||eventName;
  const p=projectDefaults(name);for(const sub of Object.values(dbx.SUBFOLDERS))await dbx.ensureFolder(fin.path+'/'+sub);
  const data={project:p,expenses:[],documents:[],history:[{at:now(),message:'Created project in Dropbox'}],path:fin.path};
  let meta;try{meta=await dbx.upload(fin.path+'/'+dbx.MASTER,await buildWorkbook(data))}catch(e){if(e instanceof dbx.DropboxError&&e.summary.includes('conflict'))throw Error('This event already has Master Finance.xlsx in its FINANCE folder. Refresh from Dropbox to open it.');throw e}
  const entry={id:meta.id,path:fin.path,year,event:eventName,name:eventName.includes(year)?eventName:`${eventName} (${year})`};created.set(entry.id,entry);projects.push(entry);accessCache.set(entry.id,{at:Date.now(),value:check});return {id:entry.id,name:entry.name};
 }
 if(route==='import/read'){const sheets=await readUpload(r.name,bytes64(r.content)),id=uid();imports.set(id,{id,project:r.project,name:r.name,sheets});return {id,sheets}}
 if(route==='import/preview'){
  const obj=imports.get(r.id);if(!obj||obj.project!==r.project)throw Error('Read the spreadsheet again');const data=await load(r.project),rows=obj.sheets[r.sheet];if(!rows||!Number.isInteger(r.header)||r.header<1||r.header>rows.length)throw Error('Choose a valid header row');const mapping=r.mapping;for(const col of Object.values(mapping))if(col!==''&&(!/^\d+$/.test(String(col))||Number(col)>249))throw Error('Invalid column mapping');const out=[],seen=new Set();
  for(let i=r.header;i<rows.length;i++){
   const raw=Object.fromEntries(Object.entries(mapping).map(([k,col])=>[k,col===''?'':rows[i][Number(col)]??'']));const desc=String(raw.description||'').trim();if(!desc||/^(sub\s*total|grand\s*total|total|balance|diff|control)\b/i.test(desc))continue;raw.category=raw.category||r.defaultCategory;raw.taxAmount=raw.taxAmount===''||raw.taxAmount==null?0:raw.taxAmount;raw.paid=raw.paid===''||raw.paid==null?0:raw.paid;raw.date=parseDate(raw.date);raw.taxType=raw.taxType||'Imported — review treatment';raw.unit=raw.unit||'package';const sourceKey=obj.name+'|'+r.sheet+'|'+(i+1);let e,errors=[];try{e=validateExpense(raw,data.project.id)}catch(err){e=raw;errors=[err.message]}
   const matches=data.expenses.filter(x=>x.sourceKey===sourceKey||(raw.reference&&x.reference===String(raw.reference).trim())||(x.vendor.toLowerCase()===String(raw.vendor||'').toLowerCase()&&x.description.toLowerCase()===desc.toLowerCase()&&x.date===raw.date)).map(x=>({id:x.id,description:x.description,billable:x.billable,version:x.version}));const sig=JSON.stringify([raw.reference,raw.vendor,desc,raw.date,raw.billable]);const duplicate=seen.has(sig);seen.add(sig);out.push({row:i+1,expense:e,errors,matches,duplicate,action:errors.length||matches.length||duplicate?'skip':'new',sourceKey});if(out.length>MAX)throw Error('Import at most 2,000 expenses per batch')
  }obj.rows=out;obj.sheet=r.sheet;obj.revision=data.revision;return {rows:out};
 }
 const data=await unchanged(r.project,route==='import/commit'?imports.get(r.id)?.revision:r.revision);
 if(route==='project/save'){data.project=validateSettings(r.values,data.project);await save(data,'Updated project settings');return data.project}
 if(route==='expense/save'){const raw=r.expense,old=raw.id?expense(data,raw.id,raw.version):null;const e=validateExpense(raw,data.project.id,old);if(old&&await contentHash(e)!==await contentHash(old))e.documentChecked=false;e.version=await version(e);data.expenses=old?data.expenses.map(x=>x.id===old.id?e:x):[...data.expenses,e];await save(data,(old?'Updated':'Added')+' expense: '+e.description);return e}
 if(route==='expense/delete'){expense(data,r.id,r.version);data.expenses=data.expenses.filter(e=>e.id!==r.id);await save(data,'Removed expense from master; supporting files retained');return {ok:true}}
 if(route==='document/add'){
  const e=expense(data,r.expense,r.version);if(!['Invoice','Quotation','Payment evidence','Other'].includes(r.kind))throw Error('Choose a document type');const bytes=bytes64(r.content);if(!bytes.length||bytes.length>15*1024*1024)throw Error('Choose a file up to 15 MB');const sha256=await hash(bytes);let d=data.documents.find(x=>x.sha256===sha256&&x.kind===r.kind);
  if(!d){const id=(await hash(sha256+r.kind)).slice(0,32),folder=dbx.folderFor(r.kind),rel=folder+'/'+dbx.documentFileName({vendor:e.vendor,description:e.description,kind:r.kind,id,original:r.name}),path=dbx.documentPath(data.path,rel);try{await dbx.upload(path,bytes)}catch(err){if(!(err instanceof dbx.DropboxError&&err.summary.includes('conflict')))throw err;const existing=await dbx.download(path);if(await hash(existing.bytes)!==sha256)throw Error('A different document already exists at this path')};d={id,project:data.project.id,name:String(r.name).slice(0,200),kind:r.kind,size:bytes.length,sha256,path:rel,created:now()};data.documents.push(d)}
  if(!e.documents.includes(d.id))e.documents.push(d.id);e.documentChecked=false;e.version=await version(e);try{await save(data,'Attached '+d.kind.toLowerCase()+': '+d.name)}catch(err){throw Error('The file is in the FINANCE folder ('+d.path.split('/')[0]+'), but its expense link could not be saved. Refresh and attach it again. '+err.message)}return d;
 }
 if(route==='import/commit'){const obj=imports.get(r.id);if(!obj||obj.project!==r.project||obj.committed)throw Error('This import was already used. Read it again.');const choices=new Map(r.choices.map(x=>[x.row,x.action])),targets=new Set();let count=0,updated=0;for(const entry of obj.rows){const choice=choices.get(entry.row)||'skip';if(choice==='skip')continue;if(entry.errors.length)throw Error('Correct invalid rows before importing');let old=null;if(choice!=='new'){if(targets.has(choice))throw Error('Two rows cannot update the same expense');targets.add(choice);const match=entry.matches.find(m=>m.id===choice);if(!match)throw Error('Unknown import update target');old=expense(data,choice,match.version)}const e=validateExpense(entry.expense,data.project.id,old);e.source=obj.name+' / '+obj.sheet;e.sourceKey=entry.sourceKey;e.documentChecked=false;e.version=await version(e);data.expenses=old?data.expenses.map(x=>x.id===old.id?e:x):[...data.expenses,e];count++;if(old)updated++}if(!count)throw Error('Choose at least one row to import');await save(data,'Imported '+(count-updated)+' new and updated '+updated+' expenses');obj.committed=true;return {count,updated}}
 throw Error('Unknown operation');
}
function parseDate(value){const s=String(value||'').trim();const m=s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2}|\d{4})$/);return m?`${m[3].length===2?'20'+m[3]:m[3]}-${m[2].padStart(2,'0')}-${m[1].padStart(2,'0')}`:s}
const stamp=(d=new Date())=>`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')} ${[d.getHours(),d.getMinutes(),d.getSeconds()].map(n=>String(n).padStart(2,'0')).join('.')}`;
export async function exportReportFile(key,revision,final){if(writing)throw Error('Wait for the current save to finish');writing=true;try{const data=await unchanged(key,revision);if(final&&!finalReady(data))throw Error('Final report blocked: every vendor cost needs a cost type and a checked invoice or quotation, every 12WIRED row needs its amount from the project head, and billing settings must be confirmed.');if(final){const evidence=data.documents.filter(d=>['Invoice','Quotation'].includes(d.kind)&&data.expenses.some(e=>e.documents.includes(d.id)));for(const d of evidence){const path=ownDocument(data,d);try{const meta=await dbx.rpc('files/get_metadata',{path});if(meta['.tag']!=='file')throw Error('Not a file')}catch{throw Error('Supporting document unavailable in Dropbox: '+d.name+'. Reattach it before final export.')}}}const bytes=await buildWorkbook(data,{master:false,final});const filename=`${dbx.fileName(data.project.name)} - ${final?'Final':'Draft'} report - ${stamp()}.xlsx`;await dbx.upload(data.path+'/'+dbx.SUBFOLDERS.reports+'/'+filename,bytes);return {bytes,filename}}finally{writing=false}}
export async function documentFile(key,id){const data=cache.get(key);const d=data?.documents.find(x=>x.id===id);if(!d)throw Error('Document not found. Refresh first.');return {bytes:(await dbx.download(ownDocument(data,d))).bytes,filename:d.name}}
export async function masterFile(key){const p=find(key);return {bytes:(await dbx.download(masterPath(p))).bytes,filename:dbx.MASTER}}
export async function folderAccess(key){return access(find(key),true)}

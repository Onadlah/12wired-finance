export const CATEGORIES=['Manpower','Transportation','Catering & Meals','Riders & Dressing Rooms','Operations','Miscellaneous'];
// Cost type says how a cost was paid or sourced, as in the promoter recap. "12WIRED" marks
// 12WIRED's own manpower lines: no vendor invoice; the project head fills the amounts.
export const INTERNAL='12WIRED';
export const COST_TYPES=['Accounts payable','Company card','Cash advance','LO payment recap','Direct receipt','Inventory',INTERNAL];
export const TAX_TYPES=['No tax','WHT gross-up','VAT added','Tax included in receipt'];
export const uid=()=>crypto.randomUUID().replaceAll('-','');
export const now=()=>new Date().toISOString();
export const fields=['category','costType','vendor','description','date','qty','duration','unit','price','baseAmount','taxType','taxPct','taxAmount','billable','payable','paid','reference','notes'];
export const clone=v=>structuredClone(v);
export const isInternal=e=>String(e?.costType??'').trim().toUpperCase()===INTERNAL;
export function projectDefaults(name){return {id:uid(),name:name.trim().slice(0,120),demo:false,eventEnd:'',feePct:0,vatPct:0,advance:0,settingsConfirmed:false,previous:Object.fromEntries(CATEGORIES.map(c=>[c,null])),previousLabel:'Previous project',explanations:Object.fromEntries(CATEGORIES.map(c=>[c,''])),created:now()}}
export function money(value,nullable=false){
 if(value===null||value===undefined||value===''){if(nullable)return null;throw Error('A required amount is missing. Enter 0 only for a genuine zero.')}
 if(typeof value==='boolean'||typeof value==='object')throw Error('Invalid amount');
 let s=String(value).trim().replace(/IDR|Rp|\s/g,'');const neg=s.startsWith('(')&&s.endsWith(')');s=s.replace(/[()]/g,'');
 if(s==='-')s='0';
 if(s.includes(',')&&s.includes('.'))s=s.lastIndexOf(',')>s.lastIndexOf('.')?s.replaceAll('.','').replace(',','.'):s.replaceAll(',','');
 else if(s.includes(','))s=/^-?\d{1,3}(,\d{3})+$/.test(s)?s.replaceAll(',',''):s.replace(',','.');
 else if(typeof value==='string'&&/^-?\d{1,3}(\.\d{3})+$/.test(s))s=s.replaceAll('.','');
 if(!s)throw Error('A required amount is missing');let n=Number(s)*(neg?-1:1);if(!Number.isFinite(n)||Math.abs(n)>1e15)throw Error('Invalid or unsupported amount');return Math.round((n+Number.EPSILON)*100)/100;
}
// Tax % is held as percentage points (2.5 means 2.5%). A trailing "%" is accepted.
export function percent(value){if(value===null||value===undefined||String(value).trim()==='')return null;const n=money(String(value).trim().replace(/%$/,''));if(n<0||n>=100)throw Error('Tax % must be at least 0 and below 100');return n}
// Tax amount from a rate. WHT gross-up: the vendor receives the base, so the billed total is
// base ÷ (1 − rate), rounded to whole rupiah. Other treatments: base × rate. "No tax" and
// "Tax included in receipt" give zero.
export function taxFromRate(base,pct,type=''){const t=String(type).toLowerCase();if(base===null||base===undefined||pct===null||pct===undefined)throw Error('Enter the base amount and tax % first');if(t.includes('no tax')||t.includes('included'))return 0;const r=pct/100;return t.includes('gross')?Math.round(base/(1-r))-base:Math.round(base*r)}
export function validDate(s){return /^\d{4}-\d{2}-\d{2}$/.test(s)&&!isNaN(Date.parse(s))&&new Date(s+'T12:00:00Z').toISOString().slice(0,10)===s}
export function validateExpense(raw,pid,old=null){
 const e={};for(const k of ['category','costType','vendor','description','date','unit','taxType','reference','notes'])e[k]=String(raw[k]??'').trim().slice(0,2000);
 const internal=isInternal(e);if(internal){e.costType=INTERNAL;if(!e.vendor)e.vendor=INTERNAL}
 if(!e.vendor||!e.description)throw Error('Vendor and description are required');if(!CATEGORIES.includes(e.category))throw Error('Choose a report category');if(!(internal&&!e.date)&&!validDate(e.date))throw Error('A valid expense date is required');
 // 12WIRED rows may be saved blank: the project head fills them before the final report.
 for(const k of ['baseAmount','taxAmount','billable'])e[k]=money(raw[k],internal);for(const k of ['payable','paid'])e[k]=internal?(money(raw[k],true)??0):money(raw[k]);
 for(const k of ['qty','duration','price']){e[k]=money(raw[k],true);if(e[k]<0)throw Error('Quantity, duration and unit price cannot be negative')}
 e.taxPct=percent(raw.taxPct);
 return {...e,id:old?.id||uid(),project:pid,documents:[...(old?.documents||[])],documentChecked:!!raw.documentChecked,source:old?.source||'Direct entry',sourceKey:old?.sourceKey||'',updated:now()};
}
export async function hash(value){const bytes=typeof value==='string'?new TextEncoder().encode(value):value;return [...new Uint8Array(await crypto.subtle.digest('SHA-256',bytes))].map(n=>n.toString(16).padStart(2,'0')).join('')}
export const contentHash=e=>hash(JSON.stringify(fields.map(k=>e[k]??null)));
export async function version(e){return hash(JSON.stringify([await contentHash(e),e.documents,e.documentChecked]))}
export const AWAITING_HEAD='12WIRED amount to be filled by the project head';
export function readiness(e,docs){
 if(isInternal(e))return e.billable===null||e.billable===undefined?[AWAITING_HEAD]:[];
 const issues=[];if(!String(e.costType??'').trim())issues.push('Cost type missing');
 const found=e.documents.map(id=>docs.find(d=>d.id===id)).filter(d=>d&&['Invoice','Quotation'].includes(d.kind));
 if(!found.length)issues.push('Invoice or quotation missing');else if(!e.documentChecked)issues.push('Confirm document matches the cost');
 return issues;
}
export function decorate(data){data.categories=CATEGORIES;for(const e of data.expenses)e.issues=readiness(e,data.documents);return data}
export function finalReady(data){return data.expenses.length>0&&data.project.settingsConfirmed&&data.expenses.every(e=>!readiness(e,data.documents).length)}
export function validateSettings(raw,old){const p={...old,name:String(raw.name||'').trim().slice(0,120),eventEnd:raw.eventEnd||'',settingsConfirmed:!!raw.settingsConfirmed,previousLabel:String(raw.previousLabel||'Previous project').slice(0,120)};
 if(!p.name)throw Error('Project name is required');if(p.eventEnd&&!validDate(p.eventEnd))throw Error('Invalid festival date');for(const k of ['feePct','vatPct','advance'])p[k]=money(raw[k]);if(p.feePct<0||p.feePct>100||p.vatPct<0||p.vatPct>100||p.advance<0)throw Error('Rates must be 0–100 and advances nonnegative');
 // Budgets were removed from the app on the owner's decision (26 Oct 2026); older masters' budget values are dropped.
 delete p.budgets;p.previous=Object.fromEntries(CATEGORIES.map(c=>[c,money(raw.previous?.[c],true)]));p.explanations=Object.fromEntries(CATEGORIES.map(c=>[c,String(raw.explanations?.[c]||'').slice(0,1000)]));return p;
}

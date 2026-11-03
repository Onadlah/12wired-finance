// Dropbox access for 12WIRED Finance.
// The Dropbox app has Full Dropbox permission because the finance folders sit in the team space.
// This code limits itself to the FINANCE folder of an event: /PROJECTS/<year>/<event>/<nn FINANCE>,
// and inside it only to Master Finance.xlsx, Invoices and Receipts, Quotations and Reports.
// It never creates, renames, moves or deletes event or FINANCE folders.
export const APP_KEY='dct3g1w0w58p4r7'; // public client ID of the new Full Dropbox app (not a secret)
export const REDIRECT='https://onadlah.github.io/12wired-finance/';
export const SCOPES='account_info.read files.metadata.read files.content.read files.content.write sharing.read';
export const PROJECTS_ROOT='/PROJECTS';
export const MASTER='Master Finance.xlsx';
export const SUBFOLDERS={invoices:'Invoices and Receipts',quotations:'Quotations',reports:'Reports'};
const SUB=Object.values(SUBFOLDERS);
const NO_PROJECTS='This Dropbox account has no PROJECTS folder. Sign in with your 12WIRED team Dropbox account.';
const TOKEN_KEY='12w-dropbox-session',STATE_KEY='12w-dropbox-pkce';
const b64url=b=>btoa(String.fromCharCode(...new Uint8Array(b))).replaceAll('+','-').replaceAll('/','_').replaceAll('=','');
const random=()=>b64url(crypto.getRandomValues(new Uint8Array(32)));
let account=null;
export function connected(){try{return !!JSON.parse(sessionStorage.getItem(TOKEN_KEY)||'null')?.access_token}catch{return false}}
export function forgetConnection(){sessionStorage.removeItem(TOKEN_KEY);sessionStorage.removeItem(STATE_KEY);account=null}
export async function connectDropbox(){if(APP_KEY.startsWith('PENDING'))throw Error('This version is not linked to a Dropbox app yet.');if(location.origin!==new URL(REDIRECT).origin)throw Error('Dropbox sign-in works only from the app’s web address, https://onadlah.github.io/12wired-finance/. This copy is a layout preview.');const verifier=random(),state=random();sessionStorage.setItem(STATE_KEY,JSON.stringify({verifier,state,created:Date.now()}));const challenge=b64url(await crypto.subtle.digest('SHA-256',new TextEncoder().encode(verifier)));const q=new URLSearchParams({client_id:APP_KEY,response_type:'code',redirect_uri:REDIRECT,code_challenge:challenge,code_challenge_method:'S256',state,scope:SCOPES,token_access_type:'offline'});location.assign('https://www.dropbox.com/oauth2/authorize?'+q);}
async function exchange(body){const res=await fetch('https://api.dropboxapi.com/oauth2/token',{method:'POST',body:new URLSearchParams(body),cache:'no-store',credentials:'omit',referrerPolicy:'no-referrer'});const data=await res.json();if(!res.ok||!data.access_token)throw Error('Dropbox authorization failed. Please connect again.');return data}
export async function finishConnection(){const q=new URLSearchParams(location.search);if(!q.has('code')&&!q.has('error'))return false;const stored=JSON.parse(sessionStorage.getItem(STATE_KEY)||'null');sessionStorage.removeItem(STATE_KEY);history.replaceState(null,'',location.pathname);if(q.has('error'))throw Error('Dropbox connection was cancelled. You can try again.');if(!stored||stored.state!==q.get('state')||Date.now()-stored.created>600000)throw Error('This connection request expired or could not be verified. Please connect again.');const data=await exchange({grant_type:'authorization_code',code:q.get('code'),client_id:APP_KEY,redirect_uri:REDIRECT,code_verifier:stored.verifier});sessionStorage.setItem(TOKEN_KEY,JSON.stringify({...data,expires:Date.now()+data.expires_in*1000}));return true;}
let refreshing=null;
async function token(){let t=JSON.parse(sessionStorage.getItem(TOKEN_KEY)||'null');if(!t?.access_token)throw Error('Connect Dropbox first');if(Date.now()>t.expires-60000){if(!t.refresh_token){forgetConnection();throw Error('Your Dropbox session expired. Connect again.')}if(!refreshing)refreshing=exchange({grant_type:'refresh_token',refresh_token:t.refresh_token,client_id:APP_KEY}).then(fresh=>{t={...t,...fresh,expires:Date.now()+fresh.expires_in*1000};sessionStorage.setItem(TOKEN_KEY,JSON.stringify(t));return t.access_token}).catch(e=>{forgetConnection();throw e}).finally(()=>refreshing=null);return refreshing}return t.access_token}
export class DropboxError extends Error{constructor(message,status,summary){super(message);this.status=status;this.summary=summary}}
async function checked(res){
 if(res.ok)return res;let data;try{data=await res.json()}catch{data={}}const summary=data.error_summary||'';let message;
 if(res.status===401&&summary.startsWith('missing_scope'))message='The Dropbox app is missing a permission ('+(data.error?.required_scope||'unknown')+'). Ask the app owner to tick it in the Dropbox app settings, then disconnect and connect again.';
 else if(res.status===401){forgetConnection();message='Your Dropbox connection expired. Reconnect to continue.'}
 else if(res.status===429)message='Dropbox is busy. Wait a moment, then try again.';
 else if(summary.includes('conflict'))message='Dropbox has a newer version. Your save was stopped. Refresh from Dropbox, review the changes, and try again.';
 else if(res.status===403||summary.includes('no_permission')||summary.includes('restricted_content'))message='Dropbox denied access to this folder. Check that you are on the folder’s access list.';
 else message='Dropbox request failed. Check your connection and retry.';
 throw new DropboxError(message,res.status,summary)}
async function send(url,headers,body){return checked(await fetch(url,{method:'POST',headers:{Authorization:'Bearer '+await token(),...headers},body,credentials:'omit',cache:'no-store',referrerPolicy:'no-referrer'}))}
// Who is signed in, and the team-space root. Team folders such as PROJECTS sit in the team root,
// so file calls carry Dropbox-API-Path-Root when the team root differs from the member's home.
export async function whoAmI(){if(account)return account;const a=await (await send('https://api.dropboxapi.com/2/users/get_current_account',{'Content-Type':'application/json'},'null')).json();const r=a.root_info||{};account={name:a.name?.display_name||'',email:a.email||'',team:r['.tag']==='team',root:r.root_namespace_id&&r.root_namespace_id!==r.home_namespace_id?String(r.root_namespace_id):null};return account}
async function pathRoot(){const a=await whoAmI();return a.root?{'Dropbox-API-Path-Root':JSON.stringify({'.tag':'root',root:a.root})}:{}}
export async function rpc(route,args){const headers={'Content-Type':'application/json',...(route.startsWith('files/')?await pathRoot():{})};const res=await send('https://api.dropboxapi.com/2/'+route,headers,args===undefined?'null':JSON.stringify(args));return res.status===204?null:res.json()}

// ---- Path rules -----------------------------------------------------------------------------
const FINANCE_NAME=/^(\d{1,2} )?FINANCE$/i;
const FILE=/^[A-Za-z0-9_. ()-]+$/;
const okName=n=>!!n&&n!=='.'&&n!=='..'&&!/[\\\u0000-\u001f]/.test(n);
const parts=path=>{if(typeof path!=='string'||path.includes('..'))return null;const s=path.split('/');return s[0]===''&&/^PROJECTS$/i.test(s[1]??'')&&/^\d{4}$/.test(s[2]??'')&&okName(s[3])?s:null};
export function isEventPath(path){const s=parts(path);return !!s&&s.length===4}
export function isFinanceFolder(path){const s=parts(path);return !!s&&s.length===5&&FINANCE_NAME.test(s[4])}
// Every Dropbox path the app reads or writes passes here.
export function safePath(path){
 const s=parts(path);if(!s||s.length<5||!FINANCE_NAME.test(s[4]))throw Error('Unexpected Dropbox path');
 const rest=s.slice(5);
 if(!rest.length||(rest.length===1&&rest[0]===MASTER))return path;
 if(SUB.includes(rest[0])&&(rest.length===1||(rest.length===2&&FILE.test(rest[1])&&!rest[1].startsWith('.'))))return path;
 throw Error('Unexpected Dropbox path');
}
// Supporting documents are stored relative to the FINANCE folder ("Invoices and Receipts/<file>"),
// so renaming or moving the event folder does not break the links.
export function documentPath(financeFolder,relative){if(!isFinanceFolder(financeFolder))throw Error('Unexpected Dropbox path');const full=financeFolder+'/'+relative;safePath(full);const s=full.split('/');if(s.length!==7||s[5]===SUBFOLDERS.reports)throw Error('Unexpected document location');return full}
export const folderFor=kind=>kind==='Quotation'?SUBFOLDERS.quotations:SUBFOLDERS.invoices;

// ---- File transfer -----------------------------------------------------------------------------
const headerJSON=value=>JSON.stringify(value).replace(/[\u007f-￿]/g,c=>'\\u'+c.charCodeAt(0).toString(16).padStart(4,'0'));
export async function download(path){safePath(path);const res=await send('https://content.dropboxapi.com/2/files/download',{'Dropbox-API-Arg':headerJSON({path}),...await pathRoot()},undefined);const meta=JSON.parse(res.headers.get('Dropbox-API-Result')||'null');if(!meta?.rev)throw Error('Dropbox did not provide a file revision');return {bytes:new Uint8Array(await res.arrayBuffer()),meta}}
export async function upload(path,bytes,rev=null){safePath(path);const mode=rev?{'.tag':'update',update:rev}:'add';const res=await send('https://content.dropboxapi.com/2/files/upload',{'Content-Type':'application/octet-stream','Dropbox-API-Arg':headerJSON({path,mode,autorename:false,strict_conflict:true,mute:true}),...await pathRoot()},bytes);return res.json()}
async function listAll(path){let d=await rpc('files/list_folder',{path,recursive:false,limit:2000});const out=[...d.entries];while(d.has_more){d=await rpc('files/list_folder/continue',{cursor:d.cursor});out.push(...d.entries);if(out.length>10000)break}return out}
const notFound=e=>e instanceof DropboxError&&e.summary.includes('not_found');

// ---- Projects, events and folder access ------------------------------------------------------------
const label=(event,year)=>event.includes(year)?event:`${event} (${year})`;
// Projects are event FINANCE folders that hold a Master Finance.xlsx the signed-in person can open.
// Dropbox search can take a minute to list a brand-new master; the store keeps new ones meanwhile.
export async function listProjects(){
 let data;try{data=await rpc('files/search_v2',{query:MASTER,options:{path:PROJECTS_ROOT,filename_only:true,file_status:'active',max_results:100}})}catch(e){if(notFound(e))throw Error(NO_PROJECTS);throw e}
 const found=new Map();
 const take=matches=>{for(const m of matches||[]){const md=m.metadata?.metadata;if(!md||md['.tag']!=='file'||String(md.name).toLowerCase()!==MASTER.toLowerCase())continue;const dir=md.path_display.slice(0,md.path_display.length-md.name.length-1);if(!isFinanceFolder(dir))continue;const [, ,year,event]=dir.split('/');found.set(md.id,{id:md.id,path:dir,year,event,name:label(event,year)})}};
 take(data.matches);let pages=0;while(data.has_more&&pages++<5){data=await rpc('files/search/continue_v2',{cursor:data.cursor});take(data.matches)}
 return [...found.values()].sort((a,b)=>b.year.localeCompare(a.year)||a.event.localeCompare(b.event));
}
// Event folders of every year in PROJECTS, newest first, for "New project".
export async function listEvents(){
 let top;try{top=await listAll(PROJECTS_ROOT)}catch(e){if(notFound(e))throw Error(NO_PROJECTS);throw e}
 const years=top.filter(e=>e['.tag']==='folder'&&/^\d{4}$/.test(e.name)).map(e=>e.name).sort().reverse();const groups=[];
 for(const year of years){const events=(await listAll(PROJECTS_ROOT+'/'+year)).filter(e=>e['.tag']==='folder'&&okName(e.name)).map(e=>({name:e.name,path:`${PROJECTS_ROOT}/${year}/${e.name}`})).sort((a,b)=>a.name.localeCompare(b.name));if(events.length)groups.push({year,events})}
 return groups;
}
export async function findFinanceFolder(eventPath){if(!isEventPath(eventPath))throw Error('Unexpected Dropbox path');const found=(await listAll(eventPath)).filter(e=>e['.tag']==='folder'&&FINANCE_NAME.test(e.name));if(!found.length)return null;if(found.length>1)throw Error('This event has more than one FINANCE folder. Keep one, then try again.');const e=found[0];return {path:eventPath+'/'+e.name,name:e.name,noAccess:!!e.sharing_info?.no_access}}
// Restricted means a shared folder that does not inherit access from the folder above it ("no_inherit"):
// only the people and groups on its own list can open it. The restriction may sit on the FINANCE folder
// itself or on a folder above it, such as the event folder. Reaching the team folder (PROJECTS) without
// finding one means the FINANCE folder is open to everyone who can open PROJECTS.
export async function checkAccess(path){
 if(!isFinanceFolder(path))throw Error('Unexpected Dropbox path');
 try{
  const meta=await rpc('files/get_metadata',{path});let id=meta.sharing_info?.shared_folder_id||meta.sharing_info?.parent_shared_folder_id;
  for(let level=0;id&&level<8;level++){
   const shared=await rpc('sharing/get_folder_metadata',{shared_folder_id:id});
   if(shared.access_inheritance?.['.tag']==='no_inherit'){let members=[],groups=[];
    try{const list=await rpc('sharing/list_folder_members',{shared_folder_id:id,limit:100});members=(list.users||[]).map(u=>({name:u.user?.display_name||u.user?.email||'Team member',email:u.user?.email||'',access:u.access_type?.['.tag']||''}));groups=(list.groups||[]).map(g=>({name:g.group?.group_name||'Group',count:g.group?.member_count??null,access:g.access_type?.['.tag']||''}))}catch{}
    return {restricted:true,restrictedAt:shared.name||'',members,groups}}
   if(shared.is_team_folder)break;
   id=shared.parent_shared_folder_id;
  }
  return {restricted:false,members:[],groups:[]};
 }catch(e){return {restricted:null,error:e.message,members:[],groups:[]}}
}
// Creates Invoices and Receipts, Quotations or Reports inside a FINANCE folder when missing. Nothing else.
export async function ensureFolder(path){safePath(path);const s=path.split('/');if(s.length!==6)throw Error('Unexpected Dropbox path');try{return await rpc('files/create_folder_v2',{path,autorename:false})}catch(e){if(!(e instanceof DropboxError&&e.summary.startsWith('path/conflict/folder')))throw e}}

// ---- File names ------------------------------------------------------------------------------
export function readable(s,max=60){return String(s??'').normalize('NFKD').replace(/[̀-ͯ]/g,'').replace(/&/g,' and ').replace(/[^A-Za-z0-9_. ()-]+/g,' ').replace(/\.{2,}/g,'.').replace(/\s+/g,' ').trim().slice(0,max).trim()}
export function fileName(name){return readable(name,150).replace(/^\.+/,'')||'attachment'}
// "Vendor - Description - Invoice - 1a2b3c4d.jpg": readable in Dropbox; the 8-character code keeps names unique.
export function documentFileName({vendor,description,kind,id,original}){const ext=(String(original||'').match(/\.([A-Za-z0-9]{1,5})$/)||[])[1]?.toLowerCase();const base=[readable(vendor,40),readable(description,60),readable(kind,20)].filter(Boolean).join(' - ')||'Document';return (`${base} - ${String(id).slice(0,8)}`+(ext?'.'+ext:'')).replace(/^\.+/,'')}

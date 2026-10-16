// Lift Board <-> Training sheet. Deploy: Web app, Execute as Me, Access Anyone.
// v3. Writes ONLY to the Lift tab. Reads the Jiujitsu tab for hours; never writes it.
const SHEET_ID='1tUl1mCnkoc8Ri3Alen3ODYBU82iIX6NGOw5x_huwV6M', LIFT='Lift', JJ='Jiujitsu', HOURS_SINCE='2026-10-12';
const COLS=['Date','Session','Exercise','Sets','Reps','Weight','Note','Status'];
function ss_(){return SpreadsheetApp.openById(SHEET_ID);}
function out_(o){return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);}
function fmt_(v){return v instanceof Date?Utilities.formatDate(v,'America/New_York','yyyy-MM-dd'):String(v==null?'':v);}
function day_(v,tz){return v instanceof Date?Utilities.formatDate(v,tz||'America/New_York','yyyy-MM-dd'):String(v==null?'':v).trim().slice(0,10);}
function dur_(note){let h=0;String(note).toLowerCase().split(' ').join(',').split(',').forEach(t=>{t=t.trim();if(t.length>1&&t.charAt(t.length-1)==='h'){const n=parseFloat(t);if(n>0)h+=n;}});return h;}
function col_(h,name,dflt){for(let j=0;j<h.length;j++)if(String(h[j]).trim().toLowerCase()===name)return j;return dflt;}
function hours_(dbg){const ss=ss_(),sh=ss.getSheetByName(JJ),tz=ss.getSpreadsheetTimeZone();
  if(!sh){if(dbg)dbg.error='no tab '+JJ;return 0;}
  const v=sh.getDataRange().getValues(),h=v[0]||[],cd=col_(h,'date',0),cn=col_(h,'note',4),cs=col_(h,'status',5);let tot=0;
  if(dbg){dbg.sheet=ss.getName();dbg.id=ss.getId();dbg.tab=sh.getName();dbg.tz=tz;dbg.rows=v.length;dbg.cols={date:cd,note:cn,status:cs};dbg.seen=[];}
  for(let i=1;i<v.length;i++){const d=day_(v[i][cd],tz),st=String(v[i][cs]).toLowerCase(),add=dur_(v[i][cn]);
    const ok=d.length===10&&d.charAt(4)==='-'&&d>=HOURS_SINCE&&st.indexOf('logged')>=0;
    if(dbg)dbg.seen.push([i+1,typeof v[i][cd]==='object'?'Date':typeof v[i][cd],d,String(v[i][cn]),st,add,ok]);
    if(ok)tot+=add;}
  return tot;}
function list_(dbg){const v=ss_().getSheetByName(LIFT).getDataRange().getValues(),rows=[];
  for(let i=1;i<v.length;i++){if(!String(v[i][2]).trim())continue;const r={row:i+1};
    COLS.forEach((n,j)=>r[n.toLowerCase()]=j===0?day_(v[i][j]):fmt_(v[i][j]));rows.push(r);}
  return {updated:Utilities.formatDate(new Date(),'America/New_York',"yyyy-MM-dd'T'HH:mm"),rows,hours:hours_(dbg),v:3};}
function log_(p){if(!p.exercise||!String(p.weight||'').trim())return {ok:false,error:'need exercise and weight'};
  const s=ss_().getSheetByName(LIFT),d=p.date?String(p.date).slice(0,10):Utilities.formatDate(new Date(),'America/New_York','yyyy-MM-dd');
  s.appendRow([d,p.session||'',p.exercise,p.sets||'',p.reps||'',p.weight,p.note||'','logged']);return {ok:true,row:s.getLastRow()};}
function undo_(p){const s=ss_().getSheetByName(LIFT),row=parseInt(p.row,10);
  if(!(row>1&&row<=s.getLastRow()))return {ok:false,error:'bad row'};
  if(p.exercise&&String(s.getRange(row,3).getValue())!==p.exercise)return {ok:false,error:'row moved; reload'};
  s.getRange(row,8).setValue('undone');return {ok:true,row};}
function handle_(p){const l=LockService.getScriptLock();l.waitLock(10000);
  try{return out_(p.action==='log'?log_(p):p.action==='undo'?undo_(p):p.debug?(function(){const d={};const r=list_(d);r.debug=d;return r;})():list_());}finally{l.releaseLock();}}
function doGet(e){return handle_(e.parameter||{});}
function doPost(e){let p={};try{p=JSON.parse(e.postData.contents);}catch(x){p=e.parameter||{};}return handle_(p);}

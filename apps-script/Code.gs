// Lift Board <-> Training sheet. Deploy: Web app, Execute as Me, Access Anyone.
// Writes ONLY to the Lift tab. Reads the Jiujitsu tab for hours; never writes it.
const SHEET_ID='1tUl1mCnkoc8Ri3Alen3ODYBU82iIX6NGOw5x_huwV6M', LIFT='Lift', JJ='Jiujitsu', HOURS_SINCE='2026-10-12';
const COLS=['Date','Session','Exercise','Sets','Reps','Weight','Note','Status'];
function ss_(){return SpreadsheetApp.openById(SHEET_ID);}
function out_(o){return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);}
function fmt_(v){return v instanceof Date?Utilities.formatDate(v,Session.getScriptTimeZone(),'yyyy-MM-dd'):String(v==null?'':v);}
function hours_(){const v=ss_().getSheetByName(JJ).getDataRange().getValues();let h=0;
  for(let i=1;i<v.length;i++){const d=fmt_(v[i][0]);if(!/^\d{4}-\d\d-\d\d$/.test(d)||d<HOURS_SINCE)continue;
    if(!/logged/i.test(String(v[i][5])))continue;const m=String(v[i][4]).match(/(\d+(?:\.\d+)?)\s*h/i);if(m)h+=parseFloat(m[1]);}
  return h;}
function list_(){const v=ss_().getSheetByName(LIFT).getDataRange().getValues(),rows=[];
  for(let i=1;i<v.length;i++){if(!String(v[i][2]).trim())continue;const r={row:i+1};
    COLS.forEach((n,j)=>r[n.toLowerCase()]=fmt_(v[i][j]));rows.push(r);}
  return {updated:Utilities.formatDate(new Date(),Session.getScriptTimeZone(),"yyyy-MM-dd'T'HH:mm"),rows,hours:hours_()};}
function log_(p){if(!p.exercise||!String(p.weight||'').trim())return {ok:false,error:'need exercise and weight'};
  const s=ss_().getSheetByName(LIFT),d=p.date||Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyy-MM-dd');
  s.appendRow([d,p.session||'',p.exercise,p.sets||'',p.reps||'',p.weight,p.note||'','logged']);return {ok:true,row:s.getLastRow()};}
function undo_(p){const s=ss_().getSheetByName(LIFT),row=parseInt(p.row,10);
  if(!(row>1&&row<=s.getLastRow()))return {ok:false,error:'bad row'};
  if(p.exercise&&String(s.getRange(row,3).getValue())!==p.exercise)return {ok:false,error:'row moved; reload'};
  s.getRange(row,8).setValue('undone');return {ok:true,row};}
function handle_(p){const l=LockService.getScriptLock();l.waitLock(10000);
  try{return out_(p.action==='log'?log_(p):p.action==='undo'?undo_(p):list_());}finally{l.releaseLock();}}
function doGet(e){return handle_(e.parameter||{});}
function doPost(e){let p={};try{p=JSON.parse(e.postData.contents);}catch(x){p=e.parameter||{};}return handle_(p);}

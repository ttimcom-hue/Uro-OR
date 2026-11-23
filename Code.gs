// 泌尿科刀房調度板：網頁 + 同步，同一個 Apps Script 專案
const TOKEN = '';  // 想加密碼再填，例如 'uro2026'；留空＝不用密碼
function sh(){const ss=SpreadsheetApp.getActiveSpreadsheet();return ss.getSheetByName('data')||ss.insertSheet('data')}
function row(s,k){const n=s.getLastRow();if(!n)return 0;const v=s.getRange(1,1,n,1).getValues();for(let i=0;i<n;i++)if(v[i][0]===k)return i+1;return 0}
function apiGet(k){const s=sh(),r=row(s,k);return r?String(s.getRange(r,2).getValue()||''):''}
function apiSet(k,v){const L=LockService.getScriptLock();L.waitLock(15000);try{const s=sh();let r=row(s,k);if(!r){r=s.getLastRow()+1;s.getRange(r,1).setValue(k)}s.getRange(r,2).setValue(v);s.getRange(r,3).setValue(new Date())}finally{L.releaseLock()}return true}
function out(o){return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON)}
// ===== 癌症討論會追蹤表連動 =====
const CANCER_ID='1OaCFv8mcIf8lKyyX__Tj_at4Y8KsX5q50r7F2F6riiI';
function cancerStatus(){const ss=SpreadsheetApp.openById(CANCER_ID);
  const newN=+ss.getSheetByName('刀表候選').getRange('B4').getValue()||0;
  const tr=ss.getSheetByName('個案追蹤');const n=tr.getLastRow();const v=n>1?tr.getRange(2,1,n-1,8).getValues():[];
  const cnt={};let warn=0,late=0;v.forEach(r=>{if(!r[1]&&!r[2])return;const st=String(r[0]||'');cnt[st]=(cnt[st]||0)+1;const h=String(r[7]||'');if(/需排入/.test(h))warn++;if(/已過/.test(h))late++});
  const mtg=ss.getSheetByName('週會名單');const d=mtg.getRange('B1').getDisplayValue();const m=mtg.getLastRow();const listed=m>5?mtg.getRange(6,1,m-5,1).getValues().filter(r=>r[0]).length:0;
  return{newN,warn,late,cnt,mtg:d,listed,url:ss.getUrl()}}
function cancerImport(text){const ss=SpreadsheetApp.openById(CANCER_ID);const sh=ss.getSheetByName('刀表匯入');const last=sh.getLastRow();if(last>1)sh.getRange(2,1,last-1,2).clearContent();
  const lines=String(text||'').replace(/\r/g,'').split('\n').slice(0,5990);if(lines.length)sh.getRange(2,1,lines.length,1).setValues(lines.map(x=>[x]));SpreadsheetApp.flush();return cancerStatus()}
function apiCancer(action,text){return JSON.stringify(action==='import'?cancerImport(text):cancerStatus())}
function doGet(e){const p=(e&&e.parameter)||{};
  if(p.action==='cancerStatus'){if(TOKEN&&p.token!==TOKEN)return out({ok:false,err:'token'});try{return out({ok:true,value:cancerStatus()})}catch(x){return out({ok:false,err:String(x)})}}
  if(p.key){if(TOKEN&&p.token!==TOKEN)return out({ok:false,err:'token'});const v=apiGet(p.key);return out({ok:true,value:v?JSON.parse(v):null})}
  return HtmlService.createHtmlOutputFromFile('index').setTitle('泌尿科刀房調度板').addMetaTag('viewport','width=device-width, initial-scale=1').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)}
function doPost(e){const b=JSON.parse(e.postData.contents);if(TOKEN&&b.token!==TOKEN)return out({ok:false,err:'token'});if(b.action==='cancerImport'){try{return out({ok:true,value:cancerImport(b.text)})}catch(x){return out({ok:false,err:String(x)})}}apiSet(b.key,JSON.stringify(b.value));return out({ok:true})}

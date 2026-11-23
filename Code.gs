// 泌尿科刀房調度板：網頁 + 同步，同一個 Apps Script 專案
const TOKEN = '';  // 想加密碼再填，例如 'uro2026'；留空＝不用密碼
function sh(){const ss=SpreadsheetApp.getActiveSpreadsheet();return ss.getSheetByName('data')||ss.insertSheet('data')}
function row(s,k){const n=s.getLastRow();if(!n)return 0;const v=s.getRange(1,1,n,1).getValues();for(let i=0;i<n;i++)if(v[i][0]===k)return i+1;return 0}
function apiGet(k){if(k==='meeting/drive')return JSON.stringify(meetingRead());const s=sh(),r=row(s,k);return r?String(s.getRange(r,2).getValue()||''):''}
function apiSet(k,v){const L=LockService.getScriptLock();L.waitLock(15000);try{const s=sh();let r=row(s,k);if(!r){r=s.getLastRow()+1;s.getRange(r,1).setValue(k)}s.getRange(r,2).setValue(v);s.getRange(r,3).setValue(new Date())}finally{L.releaseLock()}return true}
function out(o){return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON)}
// ===== 癌症討論會追蹤表連動 =====
const CANCER_ID='1OaCFv8mcIf8lKyyX__Tj_at4Y8KsX5q50r7F2F6riiI';
function cancerStatus(){const ss=SpreadsheetApp.openById(CANCER_ID);
  const newN=+ss.getSheetByName('刀表候選').getRange('B4').getValue()||0;
  const tr=ss.getSheetByName('個案追蹤');const n=tr.getLastRow();const v=n>1?tr.getRange(2,1,n-1,8).getValues():[];
  const cnt={};let warn=0,late=0;v.forEach(r=>{if(!r[1]&&!r[2])return;const st=String(r[0]||'');cnt[st]=(cnt[st]||0)+1;const h=String(r[7]||'');if(/需排入/.test(h))warn++;if(/已過/.test(h))late++});
  const mtg=ss.getSheetByName('週會名單');const d=mtg.getRange('B1').getDisplayValue();const m=mtg.getLastRow();const listed=m>5?mtg.getRange(6,1,m-5,1).getValues().filter(r=>r[0]&&!/^（/.test(String(r[0]))).length:0;
  // gid：入口網頁「癌症討論會」分頁用來直接開到各分頁（只回傳人數、日期與分頁編號，不含病人資料）
  const gid={};[['next','下次討論會'],['todo','注意事項'],['dv','達文西術前'],['trk','個案追蹤'],['stat','統計']].forEach(([k,n])=>{const s=ss.getSheetByName(n);if(s)gid[k]=s.getSheetId()});
  return{newN,warn,late,cnt,mtg:d,listed,url:ss.getUrl(),gid}}
function cancerImport(text){const ss=SpreadsheetApp.openById(CANCER_ID);const sh=ss.getSheetByName('刀表匯入');const last=sh.getLastRow();if(last>1)sh.getRange(2,1,last-1,2).clearContent();
  const lines=String(text||'').replace(/\r/g,'').split('\n').slice(0,5990);if(lines.length)sh.getRange(2,1,lines.length,1).setValues(lines.map(x=>[x]));SpreadsheetApp.flush();try{cancerSync()}catch(x){console.log(x)}return cancerStatus()}
function apiCancer(action,text){return JSON.stringify(action==='import'?cancerImport(text):cancerStatus())}
// ===== 晨會（Google Drive「W2 科部晨會」資料夾）=====
// 每週一個子資料夾（名稱 MMDD），主簡報 YYYYMMDD.pptx 的 OUTLINE 頁＝當週議程。
// 只存議程標題與檔名（不讀各頁內文，避免存到病人姓名；不存 Drive 連結，入口不提供進入 Drive 的連結），存在「meeting」工作表，用同步鍵 meeting/drive 讀出給科務入口「晨會」分頁用。
// 需要 drive.readonly 權限（Updater.gs 的 GH_SCOPES）。每小時（ghAutoUpdate）與入口按「從 Drive 更新」時執行。
const MTG_FOLDER='1NsEvZfchGMt49LbPkJztHz2t1Ewhg6mL';
const MTG_SKIP=/履歷|甄審|申請|成績|病歷/;  // 這些檔名不列出
// 結果存在「meeting」工作表，一週一列（整包放一格會超過試算表 5 萬字上限）
function mtgSheet_(){const ss=SpreadsheetApp.getActiveSpreadsheet();return ss.getSheetByName('meeting')||ss.insertSheet('meeting')}
function meetingRead(){const sh=mtgSheet_(),n=sh.getLastRow();let ts=0,left=0;const weeks=[];
  if(n)sh.getRange(1,1,n,2).getDisplayValues().forEach(r=>{if(r[0]==='ts')ts=+r[1]||0;else if(r[0]==='left')left=+r[1]||0;else if(r[1]){try{weeks.push(JSON.parse(r[1]))}catch(e){}}});
  weeks.sort((a,b)=>a.d<b.d?1:-1);return{ts,left,weeks}}
function meetingWrite_(o){const sh=mtgSheet_();sh.clearContents();
  const rows=[['ts',String(o.ts)],['left',String(o.left||0)]].concat(o.weeks.map(w=>[w.d,JSON.stringify(w)]));
  const rg=sh.getRange(1,1,rows.length,2);rg.setNumberFormat('@');rg.setValues(rows);SpreadsheetApp.flush()}
// 一次最多跑 4.5 分鐘（Apps Script 上限 6 分鐘）；沒讀完的 1 分鐘後自動接著讀，讀過且沒改的簡報不會重讀
function meetingScan(){const t0=Date.now();
  const old=meetingRead(),cache={},byFolder={};
  old.weeks.forEach(w=>{if(w.mid)cache[w.mid+'|'+w.mt]=w;if(w.fk)byFolder[w.fk]=w});
  const folders=[];
  // 週資料夾＝MMDD（月 01–12）；年份資料夾（例如「2026」）往下找，並記下年份
  const walk=(f,depth,yr)=>{const it=f.getFolders();while(it.hasNext()){const c=it.next(),n=c.getName();
    if(/^(0[1-9]|1[0-2])(0[1-9]|[12]\d|3[01])$/.test(n))folders.push({c,yr});
    else if(depth<3)walk(c,depth+1,(n.match(/^(20\d\d)\b/)||[])[1]||yr)}};
  walk(DriveApp.getFolderById(MTG_FOLDER),0,'');
  const weeks=[];let left=0,errs=0;
  folders.forEach(({c,yr})=>{
    if(Date.now()-t0>270000){const o=byFolder[(yr||'')+'/'+c.getName()];if(o)weeks.push(o);left++;return}
    try{const w=mtgWeek_(c,cache,yr);if(w.err)errs++;weeks.push(w)}catch(e){errs++;console.log(c.getName(),e)}});
  weeks.sort((a,b)=>a.d<b.d?1:-1);
  const res={ts:Date.now(),left,weeks};meetingWrite_(res);
  ScriptApp.getProjectTriggers().filter(t=>t.getHandlerFunction()==='meetingScan').forEach(t=>ScriptApp.deleteTrigger(t));
  if(left)ScriptApp.newTrigger('meetingScan').timeBased().after(60*1000).create();
  console.log('晨會：找到 '+folders.length+' 個週資料夾，讀到議程 '+weeks.filter(w=>w.outline&&w.outline.length).length+' 週'+(errs?'，'+errs+' 週讀取失敗':'')+(left?'，還有 '+left+' 週下次繼續':''));
  return res}
function mtgWeek_(f,cache,yr){const mmdd=f.getName();const files=[];let main=null;
  const re=new RegExp('^(20\\d\\d)'+mmdd+'(\\.pptx?)?$','i');
  const it=f.getFiles();while(it.hasNext()){const x=it.next(),n=x.getName();
    if(re.test(n)&&(!main||/\.pptx$/i.test(n)))main=x;
    if(!MTG_SKIP.test(n))files.push({n,m:String(x.getMimeType()).split(/[./]/).pop()})}
  let y;if(main)y=+main.getName().match(re)[1];else if(yr)y=+yr;else{const c=f.getDateCreated();y=c.getFullYear();
    const d=new Date(y,+mmdd.slice(0,2)-1,+mmdd.slice(2));if(d<new Date(c.getTime()-60*864e5))y++}
  const w={d:y+'-'+mmdd.slice(0,2)+'-'+mmdd.slice(2),fk:(yr||'')+'/'+mmdd,files,outline:[]};
  if(main){w.mid=main.getId();w.mt=main.getLastUpdated().getTime();w.mn=main.getName();
    const c=cache[w.mid+'|'+w.mt];
    if(c&&c.outline&&c.outline.length){w.outline=c.outline;w.how=c.how}
    else{try{const r=mtgOutline_(mtgBlob_(main));w.outline=r.items;w.how=r.how}catch(e){w.err=String(e).slice(0,120)}}}
  return w}
function mtgBlob_(file){
  if(file.getMimeType()==='application/vnd.google-apps.presentation'){
    return UrlFetchApp.fetch('https://docs.google.com/presentation/d/'+file.getId()+'/export/pptx',{headers:{Authorization:'Bearer '+ScriptApp.getOAuthToken()}}).getBlob()}
  if(file.getSize()>40*1024*1024)throw new Error('檔案太大，略過');
  return file.getBlob()}
// 從 pptx 取議程：有「OUTLINE／大綱／議程」那頁就用那頁的各行；沒有就用每頁第一行（標題）
function mtgOutline_(blob){
  const dec=s=>s.replace(/&lt;/g,'<').replace(/&gt;/g,'>').replace(/&quot;/g,'"').replace(/&apos;/g,"'").replace(/&#(\d+);/g,(m,n)=>String.fromCharCode(+n)).replace(/&amp;/g,'&');
  const num=b=>+(b.getName().match(/(\d+)\.xml$/)||[0,0])[1];
  const slides=Utilities.unzip(blob.setContentType('application/zip')).filter(b=>/^ppt\/slides\/slide\d+\.xml$/.test(b.getName())).sort((a,b)=>num(a)-num(b));
  const paras=b=>{const s=b.getDataAsString('UTF-8');const out=[];const re=/<a:p(?:\s[^>]*)?>([\s\S]*?)<\/a:p>/g;let m;
    while((m=re.exec(s))){const t=dec((m[1].match(/<a:t(?:\s[^>]*)?>[^<]*<\/a:t>/g)||[]).map(x=>x.replace(/<[^>]+>/g,'')).join('')).replace(/\s+/g,' ').trim();
      if(t&&!/^\d+$/.test(t)&&!/^20\d\d[.\/-]\d/.test(t))out.push(t)}return out};
  const all=slides.map(paras);
  const oi=all.findIndex(p=>p.some(t=>/^(outline|agenda|大綱|議程)$/i.test(t)));
  if(oi>=0)return{how:'outline',items:all[oi].filter(t=>!/^(outline|agenda|大綱|議程)$/i.test(t)).slice(0,20)};
  const seen=new Set();return{how:'titles',items:all.slice(1).map(p=>p[0]).filter(t=>t&&!seen.has(t)&&seen.add(t)).slice(0,15)}}
function doGet(e){const p=(e&&e.parameter)||{};
  if(p.action==='cancerScan'){if(TOKEN&&p.token!==TOKEN)return out({ok:false,err:'token'});try{return out({ok:true,value:cancerDriveScan()})}catch(x){return out({ok:false,err:String(x)})}}
  if(p.action==='meetingScan'){if(TOKEN&&p.token!==TOKEN)return out({ok:false,err:'token'});try{return out({ok:true,value:meetingScan()})}catch(x){return out({ok:false,err:String(x)})}}
  if(p.action==='cancerStatus'){if(TOKEN&&p.token!==TOKEN)return out({ok:false,err:'token'});try{return out({ok:true,value:cancerStatus()})}catch(x){return out({ok:false,err:String(x)})}}
  if(p.key){if(TOKEN&&p.token!==TOKEN)return out({ok:false,err:'token'});const v=apiGet(p.key);return out({ok:true,value:v?JSON.parse(v):null})}
  return HtmlService.createHtmlOutputFromFile('index').setTitle('泌尿科刀房調度板').addMetaTag('viewport','width=device-width, initial-scale=1').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)}
function doPost(e){const b=JSON.parse(e.postData.contents);if(TOKEN&&b.token!==TOKEN)return out({ok:false,err:'token'});if(b.action==='cancerImport'){try{return out({ok:true,value:cancerImport(b.text)})}catch(x){return out({ok:false,err:String(x)})}}apiSet(b.key,JSON.stringify(b.value));return out({ok:true})}

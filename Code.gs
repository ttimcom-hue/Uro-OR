// 泌尿科刀房調度板：網頁 + 同步，同一個 Apps Script 專案
const TOKEN = '';  // 想加密碼再填，例如 'uro2026'；留空＝不用密碼
function sh(){const ss=SpreadsheetApp.getActiveSpreadsheet();return ss.getSheetByName('data')||ss.insertSheet('data')}
function row(s,k){const n=s.getLastRow();if(!n)return 0;const v=s.getRange(1,1,n,1).getValues();for(let i=0;i<n;i++)if(v[i][0]===k)return i+1;return 0}
function apiGet(k){if(k==='meeting/drive')return JSON.stringify(meetingRead());const s=sh(),r=row(s,k);return r?String(s.getRange(r,2).getValue()||''):''}
function apiSet(k,v){const L=LockService.getScriptLock();L.waitLock(15000);try{const s=sh();let r=row(s,k);if(!r){r=s.getLastRow()+1;s.getRange(r,1).setValue(k)}s.getRange(r,2).setValue(v);s.getRange(r,3).setValue(new Date())}finally{L.releaseLock()}return true}
function apiGetMany(keys){const s=sh(),n=s.getLastRow(),m={};if(n)s.getRange(1,1,n,2).getValues().forEach(r=>{if(r[0])m[r[0]]=r[1]});
  const v={};keys.forEach(k=>{try{const x=k==='meeting/drive'?JSON.stringify(meetingRead()):m[k];v[k]=x?JSON.parse(String(x)):null}catch(e){v[k]=null}});return v}
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
    const c=cache[w.mid+'|'+w.mt];let blob=null;const B=()=>blob||(blob=mtgBlob_(main));
    if(c&&c.outline&&c.outline.length){w.outline=c.outline;w.how=c.how}
    else{try{const r=mtgOutline_(B());w.outline=r.items;w.how=r.how}catch(e){w.err=String(e).slice(0,120)}}
    // 截圖辨識：只做最近 3 週～未來 5 週的晨會，簡報沒改過就沿用上次結果
    if(c&&c.img)w.img=c.img;
    else if(mtgRecent_(w.d)){try{w.img=mtgImgEvents_(B(),w.d)}catch(e){w.imgErr=String(e).slice(0,120)}}}
  return w}
const mtgRecent_=d=>{const t=Date.now(),x=new Date(d+'T00:00:00+08:00').getTime();return x>=t-21*864e5&&x<=t+35*864e5};

/* ---------- 晨會簡報裡的會議截圖 → 文字辨識 ----------
   有些重要會議只貼截圖。只看「標題／內文提到會議」的投影片、完全沒有文字的投影片，以及會議投影片的下一張（多半是接續的截圖），
   每週最多 8 張 png／jpg（15 KB～8 MB）。用 Drive 把圖片轉成 Google 文件（OCR）→ 讀文字 → 立刻刪除暫存檔（drive.file 權限）。
   **只存辨識出的會議名稱、日期、時間、地點**，不存整段文字；含長串數字、病歷、床號、歲等字眼的行一律丟掉。 */
const MTG_IMG_KW=/會議|月會|年會|研討|課程|講座|論壇|工作坊|公告|活動|報名|workshop|symposium|congress|meeting|conference|course|seminar|forum/i;
function mtgImgEvents_(blob,wd){
  const files=Utilities.unzip(blob.setContentType('application/zip')),by={};files.forEach(f=>by[f.getName()]=f);
  const num=b=>+(b.getName().match(/(\d+)\.xml$/)||[0,0])[1];
  const slides=files.filter(b=>/^ppt\/slides\/slide\d+\.xml$/.test(b.getName())).sort((a,b)=>num(a)-num(b));
  const out=[],seen={};let n=0,prevKw=false;
  slides.forEach((sl,i)=>{const x=sl.getDataAsString('UTF-8');
    const txt=(x.match(/<a:t(?:\s[^>]*)?>[^<]*<\/a:t>/g)||[]).map(t=>t.replace(/<[^>]+>/g,'')).join(' ').replace(/\s+/g,' ').trim();
    const kw=MTG_IMG_KW.test(txt),take=i>0&&(kw||!txt||(prevKw&&txt.length<30));prevKw=kw;if(!take||n>=8)return;
    const rel=by['ppt/slides/_rels/'+sl.getName().split('/').pop()+'.rels'];if(!rel)return;
    const R={};(rel.getDataAsString('UTF-8').match(/<Relationship\b[^>]*>/g)||[]).forEach(r=>{const id=(r.match(/Id="([^"]+)"/)||[])[1],tg=(r.match(/Target="([^"]+)"/)||[])[1];if(id&&tg)R[id]=tg});
    (x.match(/r:embed="[^"]+"/g)||[]).forEach(m=>{if(n>=8)return;const tg=R[m.slice(9,-1)];if(!tg||!/\.(png|jpe?g)$/i.test(tg))return;
      const nm='ppt/'+tg.replace(/^\.\.\//,''),img=by[nm];if(!img||seen[nm])return;seen[nm]=1;const sz=img.getBytes().length;if(sz<15000||sz>8*1024*1024)return;
      n++;try{mtgImgParse_(mtgOcr_(img,nm),wd).forEach(e=>{if(!out.some(o=>o.d===e.d&&o.title===e.title))out.push(e)})}catch(e){console.log('OCR',nm,e)}})});
  return out.slice(0,10)}
function mtgOcr_(img,nm){
  const h={Authorization:'Bearer '+ScriptApp.getOAuthToken()},b='mi'+Date.now(),ct=/\.png$/i.test(nm)?'image/png':'image/jpeg';
  const meta=JSON.stringify({name:'暫存-晨會截圖辨識（會自動刪除）',mimeType:'application/vnd.google-apps.document'});
  const body=Utilities.newBlob(`--${b}\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n${meta}\r\n--${b}\r\nContent-Type: ${ct}\r\n\r\n`).getBytes()
    .concat(img.getBytes()).concat(Utilities.newBlob(`\r\n--${b}--`).getBytes());
  const r=UrlFetchApp.fetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&ocrLanguage=zh-TW&fields=id',{method:'post',contentType:'multipart/related; boundary='+b,payload:body,headers:h,muteHttpExceptions:true});
  if(r.getResponseCode()>=300)throw new Error('圖片辨識失敗 '+r.getResponseCode());
  const id=JSON.parse(r.getContentText()).id;
  try{return DocumentApp.openById(id).getBody().getText()}
  finally{UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/'+id,{method:'delete',headers:h,muteHttpExceptions:true})}}
/* 純文字解析（可用 node 的 vm 測試）：一張截圖 → 0～1 個會議 {d,end,t,title,place}。日期要在該次晨會前 1 天之後（過去的日期不算） */
function mtgImgParse_(text,wd){
  const L=String(text||'').replace(/\r/g,'').split('\n').map(s=>s.replace(/[\s\u3000]+/g,' ').trim()).filter(Boolean);
  const bad=s=>/\d{6,}|病歷|床號|\d\s*歲|y\/o|病人|患者|身分證/i.test(s);
  const iso=(y,m,d)=>y+'-'+String(m).padStart(2,'0')+'-'+String(d).padStart(2,'0');
  const dd=(a,b)=>Math.round((new Date(b+'T00:00:00Z')-new Date(a+'T00:00:00Z'))/864e5);
  const DRE=/(?:(20\d\d|1[01]\d)\s*[年.\/-]\s*)?(\d{1,2})\s*[月.\/-]\s*(\d{1,2})(?!\d)\s*日?(?:\s*[(（][一二三四五六日][)）])?(?:\s*[~～\-–至到]\s*(?:(\d{1,2})\s*[月.\/-]\s*)?(\d{1,2})(?![\d:：])\s*日?)?/g;
  let hit=null;
  const order=L.map((s,i)=>i).sort((a,b)=>(/日期|時間|date/i.test(L[b])?1:0)-(/日期|時間|date/i.test(L[a])?1:0)||a-b);
  for(const i of order){if(bad(L[i]))continue;DRE.lastIndex=0;let m;
    while((m=DRE.exec(L[i]))){const mo=+m[2],da=+m[3];if(mo<1||mo>12||da<1||da>31)continue;if(!m[1]&&/[:：]/.test(L[i].slice(m.index+m[0].length,m.index+m[0].length+1)))continue;
      let y=m[1]?+m[1]:+wd.slice(0,4);if(y<200)y+=1911;let d=iso(y,mo,da);if(!m[1]&&dd(wd,d)<-60)d=iso(y+1,mo,da);
      if(dd(wd,d)<-1||dd(wd,d)>400)continue;let end='';if(m[5]){const em=m[4]?+m[4]:mo,ed=+m[5];if(em>=1&&em<=12&&ed>=1&&ed<=31){end=iso(+d.slice(0,4)+(em<mo?1:0),em,ed);if(end<=d)end=''}}
      hit={i,d,end,rest:L[i].slice(m.index+m[0].length)};break}
    if(hit)break}
  if(!hit)return[];
  let t='';const tm=(hit.rest+' '+(L[hit.i+1]||'')).match(/(上午|下午|晚上|AM|PM)?\s*(\d{1,2})\s*[:：]\s*(\d{2})/i);
  if(tm){let h=+tm[2];if(/下午|晚上|PM/i.test(tm[1]||'')&&h<12)h+=12;if(h<24)t=String(h).padStart(2,'0')+':'+tm[3]}
  const KW=/會|課程|研討|講座|論壇|工作坊|workshop|symposium|congress|meeting|conference|course|seminar|forum|masterclass/i,STRONG=/月會|年會|研討會|討論會|課程|講座|論壇|工作坊|workshop|symposium|congress|meeting|conference|course|seminar|forum|masterclass/i,SKIP=/^(地點|地址|會場|場地|日期|時間|報名|主辦|協辦|聯絡|費用|電話|網址|備註|講者|主持|座長)/;
  const ok=s=>!SKIP.test(s)&&!bad(s)&&s.length>=3&&s.length<=40&&!(DRE.lastIndex=0,DRE.test(s));
  const title=L.find(s=>STRONG.test(s)&&ok(s))||L.find(s=>KW.test(s)&&ok(s));DRE.lastIndex=0;
  if(!title)return[];
  const pl=L.map(s=>s.match(/^(?:地點|地址|會場|場地|venue|location)\s*[:：]?\s*(.+)/i)).find(Boolean);
  return[{d:hit.d,end:hit.end,t,title:title.replace(/^[【\[「]|[】\]」]$/g,'').slice(0,40),place:pl&&!bad(pl[1])?pl[1].slice(0,40):''}]}
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
  // ?keys=a,b,c：一次讀多個鍵（入口網頁載入用，只讀一次 data 工作表，比逐鍵請求快很多）
  if(p.keys){if(TOKEN&&p.token!==TOKEN)return out({ok:false,err:'token'});return out({ok:true,value:apiGetMany(String(p.keys).split(',').slice(0,40))})}
  if(p.action==='handoff'){if(TOKEN&&p.token!==TOKEN)return out({ok:false,err:'token'});try{return out({ok:true,value:handoffFor(p.dates)})}catch(x){return out({ok:false,err:String(x)})}}
  if(p.key){if(TOKEN&&p.token!==TOKEN)return out({ok:false,err:'token'});const v=apiGet(p.key);return out({ok:true,value:v?JSON.parse(v):null})}
  return HtmlService.createHtmlOutputFromFile('index').setTitle('泌尿科刀房調度板').addMetaTag('viewport','width=device-width, initial-scale=1').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)}
function doPost(e){const b=JSON.parse(e.postData.contents);if(TOKEN&&b.token!==TOKEN)return out({ok:false,err:'token'});if(b.action==='cancerImport'){try{return out({ok:true,value:cancerImport(b.text)})}catch(x){return out({ok:false,err:String(x)})}}apiSet(b.key,JSON.stringify(b.value));return out({ok:true})}
// ===== 黃部長交班單 → 調度板術前資料（即時讀取，只回傳數值，不回傳姓名、病歷號、床號，也不存進同步資料） =====
const HO_DOC_ID='1MUwbZVat96vpyil2ahjP6lIv-NzW_MHbOC99EXTx7Yk';
function hoDocId(){try{const v=JSON.parse(apiGet('common/handoff')||'null');const m=v&&String(v.u||'').match(/document\/d\/([\w-]{20,})/);if(m)return m[1]}catch(e){}return HO_DOC_ID}
// 讀交班單所有分頁裡 7 欄的表格 → [[床號姓名病歷號, PH, PI, Drain, Treatment, Exam/lab, Plan], …]
function hoRows(){const doc=DocumentApp.openById(hoDocId());const bodies=[];
  const walk=tabs=>(tabs||[]).forEach(t=>{try{bodies.push(t.asDocumentTab().getBody())}catch(e){}walk(t.getChildTabs&&t.getChildTabs())});
  if(doc.getTabs)walk(doc.getTabs());if(!bodies.length)bodies.push(doc.getBody());const rows=[];
  bodies.forEach(b=>b.getTables().forEach(tb=>{for(let i=0;i<tb.getNumRows();i++){const r=tb.getRow(i);if(r.getNumCells()<7)continue;const c=[];for(let j=0;j<7;j++)c.push(r.getCell(j).getText());rows.push(c)}}));return rows}
// 純文字解析（node 可測）：rows 同上；dates＝['YYYY-MM-DD',…]
function hoParse(rows,dates){const want={};(dates||[]).forEach(d=>{want[+d.slice(5,7)+'/'+(+d.slice(8))]=d});const out=[];
  const L=s=>String(s||'').replace(/\r/g,'').split('\n').map(x=>x.trim()).filter(Boolean);
  const last=(re,txt)=>{let m,r=null;const g=new RegExp(re.source,re.flags.includes('g')?re.flags:re.flags+'g');while((m=g.exec(txt)))r=m;return r};
  const dated=(lines,re)=>{let r=null,cd='';lines.forEach(l=>{const d=l.match(/^(\d{1,2}\/\d{1,2})\b/);if(d)cd=d[1];const m=l.match(re);if(m)r={m,d:cd}});return r};  // 沒寫日期的續行沿用上一行日期
  rows.forEach(c=>{if(/past history/i.test(c[1]||''))return;
    const tr=L(c[4]),ops=[];tr.forEach(l=>{const m=l.match(/^(\d{1,2})\/(\d{1,2})\s*(.*)$/);const sl=l.match(/\bTF\s*:?\s*(\d+)/i),tm=l.match(/\b(\d{1,2}:\d{2})\b/);
      if(m){ops.push({md:+m[1]+'/'+(+m[2]),op:m[3].replace(/\bTF\s*:?\s*\d+/i,'').replace(/\b\d{1,2}:\d{2}\b/,'').trim(),slot:sl?+sl[1]:tm?(tm[1].replace(/^0/,'')==='8:30'?0:null):null,time:tm?tm[1]:''})}
      else if(ops.length){const o=ops[ops.length-1];if(sl&&o.slot==null)o.slot=+sl[1];else if(tm&&!o.time){o.time=tm[1];if(o.slot==null&&tm[1].replace(/^0/,'')==='8:30')o.slot=0}else if(!sl&&!tm)o.note=(o.note?o.note+'；':'')+l}});
    const hit=ops.filter(o=>want[o.md]);if(!hit.length)return;
    const pi=L(c[2]),ex=L(c[5]),all=[...pi,...ex],txt=all.join('\n'),ph=String(c[1]||''),plan=String(c[6]||''),f={};
    const v=(re,k,fmt)=>{const r=dated(all,re);if(r)f[k]=fmt(r.m)+(r.d?' ('+r.d+')':'')};
    v(/PSA\s*:?\s*(\d+(?:\.\d+)?)/i,'psa',m=>m[1]);
    v(/TPV\s*[=:]?\s*(\d+(?:\.\d+)?)/i,'tpv',m=>m[1]+'cc');
    v(/(?:PC|GS)\s*(?:GS)?\s*(\d\s*\+\s*\d)/i,'gs',m=>m[1].replace(/\s/g,''));
    v(/\bPR-?\s*(\d)\b|PI-?RADS\s*(\d)/i,'pirads',m=>m[1]||m[2]);
    v(/\bcT\d[a-c]?(?:N\d)?(?:M\d)?/,'ct',m=>m[0]);
    const lab=dated(ex.length?ex:all,/\bW\s*(\d+(?:\.\d+)?)\s+H\s*(\d+(?:\.\d+)?)\s+C\s*(\d+(?:\.\d+)?)/);
    if(lab){f.wbc=lab.m[1];f.hb=lab.m[2];f.cr=lab.m[3];if(lab.d)f.lab=lab.d}else v(/\bC\s?(\d\.\d+)\b/,'cr',m=>m[1]);
    v(/\bUA\s+(R\S*\s+W\S*\s+B\S*)/,'ua',m=>m[1]);
    v(/\bU\/C\s*:?\s*(.+)/i,'uc',m=>m[1]);
    v(/(?:心超|echo\s*cardio|\bEF\b)\s*:?\s*(.+)/i,'echo',m=>m[1].replace(/^EF\s*/i,'EF '));
    v(/肺功\s*:?\s*(.+)/,'pft',m=>m[1]);
    const iief=last(/IIEF\s*:?\s*(\d+)/i,plan+'\n'+txt);if(iief)f.iief=iief[1];
    const AC=/aspirin|plavix|clopidogrel|warfarin|coumadin|eliquis|apixaban|xarelto|rivaroxaban|lixiana|edoxaban|pradaxa|dabigatran|brilinta|ticagrelor|NOAC|efient|pletaal|cilostazol/gi;
    const ac=[...new Set(((ph+'\n'+plan+'\n'+txt).match(AC)||[]).map(x=>x[0].toUpperCase()+x.slice(1).toLowerCase()))];if(ac.length)f.ac=ac.join('、');
    const hold=last(/hold\s+[A-Za-z]+/i,plan);if(hold)f.hold=hold[0].replace(/^hold/i,'Hold');
    if(String(c[3]||'').trim())f.drain=L(c[3]).join('；');
    hit.forEach(o=>out.push({date:want[o.md],slot:o.slot,time:o.time||'',op:o.op,note:o.note||'',f}))});
  return out}
function handoffFor(dates){dates=String(dates||'').split(',').filter(d=>/^\d{4}-\d{2}-\d{2}$/.test(d)).slice(0,7);
  const C=CacheService.getScriptCache(),k='ho:'+dates.join(',');const h=C.get(k);if(h)return JSON.parse(h);
  const r={at:Date.now(),list:hoParse(hoRows(),dates)};try{C.put(k,JSON.stringify(r),300)}catch(e){}return r}
function apiHandoff(dates){return JSON.stringify(handoffFor(dates))}
// 在編輯器手動執行一次：授權讀 Google 文件並確認讀得到（只記錄筆數）
function hoTest(){const rows=hoRows();const d=new Date(),iso=x=>Utilities.formatDate(x,Session.getScriptTimeZone(),'yyyy-MM-dd');const ds=[0,1,2,3].map(i=>iso(new Date(d.getTime()+i*864e5)));Logger.log('交班單表格列數 '+rows.length+'；今天起 4 天對到 '+hoParse(rows,ds).length+' 筆手術')}

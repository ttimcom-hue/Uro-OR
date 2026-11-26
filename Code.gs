// 泌尿科刀房調度板：網頁 + 同步，同一個 Apps Script 專案
const TOKEN = '';  // 想加密碼再填，例如 'uro2026'；留空＝不用密碼
function sh(){const ss=SpreadsheetApp.getActiveSpreadsheet();return ss.getSheetByName('data')||ss.insertSheet('data')}
function row(s,k){const n=s.getLastRow();if(!n)return 0;const v=s.getRange(1,1,n,1).getValues();for(let i=0;i<n;i++)if(v[i][0]===k)return i+1;return 0}
function dataGet_(k){if(k==='meeting/drive')return JSON.stringify(meetingRead());const s=sh(),r=row(s,k);return r?String(s.getRange(r,2).getValue()||''):''}
function dataSet_(k,v){const L=LockService.getScriptLock();L.waitLock(15000);try{const s=sh();let r=row(s,k);if(!r){r=s.getLastRow()+1;s.getRange(r,1).setValue(k)}s.getRange(r,2).setValue(v);s.getRange(r,3).setValue(new Date())}finally{L.releaseLock()}return true}
function dataGetMany_(keys){const s=sh(),n=s.getLastRow(),m={};if(n)s.getRange(1,1,n,2).getValues().forEach(r=>{if(r[0])m[r[0]]=r[1]});
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
function cancerImport_(text){const ss=SpreadsheetApp.openById(CANCER_ID);const sh=ss.getSheetByName('刀表匯入');const last=sh.getLastRow();if(last>1)sh.getRange(2,1,last-1,2).clearContent();
  const lines=String(text||'').replace(/\r/g,'').split('\n').slice(0,5990);if(lines.length)sh.getRange(2,1,lines.length,1).setValues(lines.map(x=>[x]));SpreadsheetApp.flush();try{cancerSync()}catch(x){console.log(x)}return cancerStatus()}
function apiCancer_(action,text){return JSON.stringify(action==='import'?cancerImport_(text):cancerStatus())}
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
// ===== 登入與權限（Google 帳號）=====
// 1) 病人相關資料（交班單術前數值 handoff、Drive 討論會掃描 cancerScan）只給「核可帳號」：
//    ＝黃部長交班單（hoDocId()）的擁有者、編輯者、檢視者／加註者，加上指令碼屬性 AUTH_EXTRA（逗號分隔 email）。每 10 分鐘重抓一次。
// 2) 刀表（days/）、急診刀（er/list）不含可辨識病人資料，調度板所有功能（含排刀模式）都不用登入。
//    （保留 EDITORS／action=editors 刀表編輯者名單的伺服器功能，目前網頁沒有使用。）
//    管理員＝交班單的擁有者與編輯者，加上指令碼屬性 AUTH_ADMIN（逗號分隔 email）。
// GOOGLE_CLIENT_ID：Google Cloud「OAuth 用戶端 ID（網頁應用程式）」，已授權的 JavaScript 來源要加 GitHub Pages 網址。
//   留空＝不啟用登入（資料維持原本開放）。設定步驟見 開發說明.md「登入與權限」。
const GOOGLE_CLIENT_ID='271644198276-u9n58etp9ie1mu0o6f0i59c3tu763p6f.apps.googleusercontent.com';
const PRIVATE_KEY=/^(user\/|sess)/;            // 只能經由登入 API 存取
const SESSION_DAYS=30;
const PAGES_URL='https://ttimcom-hue.github.io/Uro-OR/';
function authOn(){return !!GOOGLE_CLIENT_ID}
function sessHash_(t){return Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,String(t)))}
function sessEmail_(t){if(!t)return '';const h=sessHash_(t),C=CacheService.getScriptCache();let v=C.get('s:'+h);
  if(!v){v=PropertiesService.getScriptProperties().getProperty('sess:'+h);if(v)try{C.put('s:'+h,v,600)}catch(e){}}
  if(!v)return '';const s=JSON.parse(v);return s.exp>Date.now()?s.e:''}
const em_=e=>String(e||'').trim().toLowerCase();
const emails_=s=>String(s||'').split(/[,;\s]+/).map(em_).filter(e=>/@/.test(e));
// 交班單共用名單：{all:核可帳號, admins:管理員}，快取 10 分鐘
function shareList_(){const C=CacheService.getScriptCache(),c=C.get('allow2');if(c)return JSON.parse(c);
  const all={},adm={},P=PropertiesService.getScriptProperties();
  try{const f=DriveApp.getFileById(hoDocId());try{const o=em_(f.getOwner().getEmail());all[o]=adm[o]=1}catch(e){}
    f.getEditors().forEach(u=>{const e=em_(u.getEmail());if(e)all[e]=adm[e]=1});f.getViewers().forEach(u=>{const e=em_(u.getEmail());if(e)all[e]=1})}catch(e){console.log('shareList',e)}
  emails_(P.getProperty('AUTH_EXTRA')).forEach(e=>all[e]=1);emails_(P.getProperty('AUTH_ADMIN')).forEach(e=>all[e]=adm[e]=1);
  const r={all:Object.keys(all),admins:Object.keys(adm)};try{C.put('allow2',JSON.stringify(r),600)}catch(e){}return r}
function allowList_(){return shareList_().all}
function approved_(email){return !!email&&allowList_().indexOf(em_(email))>=0}
function admin_(email){return !!email&&shareList_().admins.indexOf(em_(email))>=0}
function editorsGet_(){try{return JSON.parse(PropertiesService.getScriptProperties().getProperty('EDITORS')||'[]').map(em_).filter(Boolean)}catch(e){return[]}}
function editor_(email){return !!email&&(admin_(email)||editorsGet_().indexOf(em_(email))>=0)}
// 管理員維護刀表編輯者名單
function authEditors_(sess,list){const e=sessEmail_(sess);if(!admin_(e))throw new Error('admin');
  if(Array.isArray(list)){const L=[...new Set(list.map(em_).filter(x=>/^[^@\s]+@[^@\s]+$/.test(x)))].slice(0,100);PropertiesService.getScriptProperties().setProperty('EDITORS',JSON.stringify(L))}
  return{editors:editorsGet_(),admins:shareList_().admins}}
// 這個請求能不能讀寫病人資料：沒啟用登入＝可以；有啟用＝要有效登入且是核可帳號
function canPHI_(sess){return !authOn()||approved_(sessEmail_(sess))}
function userGet_(email){try{return JSON.parse(dataGet_('user/'+email)||'{}')||{}}catch(e){return{}}}
function userPut_(email,u){dataSet_('user/'+email,JSON.stringify(u))}
function meFor_(email){const u=userGet_(email);return{email,name:u.name||'',pic:u.pic||'',ok:approved_(email),edit:editor_(email),admin:admin_(email),prefs:u.prefs||{}}}
function sessClean_(P){const all=P.getProperties(),now=Date.now();Object.keys(all).forEach(k=>{if(/^sess:/.test(k)){try{if(JSON.parse(all[k]).exp<now)P.deleteProperty(k)}catch(e){P.deleteProperty(k)}}})}
function sessNew_(email){const tok=Utilities.getUuid().replace(/-/g,'')+Utilities.getUuid().replace(/-/g,''),s={e:email,exp:Date.now()+SESSION_DAYS*864e5};
  const P=PropertiesService.getScriptProperties();sessClean_(P);P.setProperty('sess:'+sessHash_(tok),JSON.stringify(s));return{token:tok,exp:s.exp,me:meFor_(email)}}
// 用 Google 登入回傳的 ID token 換一個 30 天的登入憑證
function authLogin_(idToken){if(!authOn())throw new Error('off');
  const r=UrlFetchApp.fetch('https://oauth2.googleapis.com/tokeninfo?id_token='+encodeURIComponent(String(idToken||'')),{muteHttpExceptions:true});
  if(r.getResponseCode()!==200)throw new Error('token');const t=JSON.parse(r.getContentText());
  if(t.aud!==GOOGLE_CLIENT_ID||String(t.email_verified)!=='true'||+t.exp*1000<Date.now()||!/accounts\.google\.com$/.test(String(t.iss)))throw new Error('token');
  const email=String(t.email).toLowerCase(),u=userGet_(email);if(!u.name)u.name=t.name||'';u.gname=t.name||'';u.pic=t.picture||'';u.last=Date.now();userPut_(email,u);
  return sessNew_(email)}
function authLogout_(sess){if(!sess)return;const h=sessHash_(sess);PropertiesService.getScriptProperties().deleteProperty('sess:'+h);CacheService.getScriptCache().remove('s:'+h)}
// 登入碼：已登入的瀏覽器產生 8 碼（10 分鐘、一次有效），在主畫面 App 等無法跳出 Google 登入的地方輸入
function authCode_(sess){const e=sessEmail_(sess);if(!e)throw new Error('auth');const A='ABCDEFGHJKLMNPQRSTUVWXYZ23456789';let c='';const b=Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,Utilities.getUuid());for(let i=0;i<8;i++)c+=A[(b[i]+256)%A.length];
  CacheService.getScriptCache().put('code:'+c,e,600);return c}
function authClaim_(code){code=String(code||'').toUpperCase().replace(/[^A-Z0-9]/g,'');const C=CacheService.getScriptCache(),e=code.length===8&&C.get('code:'+code);if(!e)throw new Error('code');C.remove('code:'+code);return sessNew_(e)}
function authPrefs_(sess,b){const e=sessEmail_(sess);if(!e)throw new Error('auth');const u=userGet_(e);
  if(b.prefs&&typeof b.prefs==='object'){const s=JSON.stringify(b.prefs);if(s.length>4000)throw new Error('too big');u.prefs=b.prefs}
  if(typeof b.name==='string')u.name=b.name.slice(0,40);u.ts=Date.now();userPut_(e,u);return meFor_(e)}
// 讀取：私密鍵一律不給，其他都開放
function readKey_(k,sess){if(PRIVATE_KEY.test(k))return{deny:1};const v=dataGet_(k);return{v:v?JSON.parse(v):null}}
// 【注意】只要有 HtmlService 網頁（這個上傳頁），任何人都能用 google.script.run 呼叫「名稱不以 _ 結尾」的伺服器函式。
//   所以讀資料、寫刀表、讀交班單的函式一律以 _ 結尾（dataGet_、dataSet_、hoRows_、handoffFor_、cancerImport_…），
//   新增的伺服器函式也要照做；只有網頁真的要叫的（apiUpload）、觸發器與選單用的才不加 _。
// ===== 院內刀表上傳頁（?page=upload）：醫院電腦擋 github.io 時用。病歷號、姓名在瀏覽器就先去掉（病歷號換成 3 字元病人代碼 99xxxx），
//       只存去識別化的文字到 inbox/or；調度板打開時自動套用。原始全文只在勾選時送到癌症討論會試算表（跟調度板匯入相同）。
function apiUpload(txt,raw,by){txt=String(txt||'');if(!txt.trim()||txt.length>800000)throw new Error('內容是空的或太大');
  const now=Date.now(),who=String(by||'').slice(0,20);
  dataSet_('inbox/or',JSON.stringify({txt,ts:now,by:who,done:0}));
  dataSet_('meta/import',JSON.stringify({ts:now,src:'院內上傳',by:who}));
  let can=null;if(raw){try{can=cancerImport_(String(raw))}catch(e){can={err:String(e&&e.message||e)}}}
  return{last:Utilities.formatDate(new Date(now),'Asia/Taipei','M/d HH:mm'),can}}
function uploadPage_(){let mi=null;try{mi=JSON.parse(dataGet_('meta/import')||'null')}catch(e){}
  const esc=s=>String(s).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
  const last=mi&&mi.ts?esc(Utilities.formatDate(new Date(mi.ts),'Asia/Taipei','M/d HH:mm')+(mi.src?'（'+mi.src+(mi.by?'・'+mi.by:'')+'）':'')):'尚無紀錄';
  return HtmlService.createHtmlOutput(UPLOAD_HTML.replace('{{LAST}}',last)).setTitle('刀表上傳（院內）').addMetaTag('viewport','width=device-width, initial-scale=1').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)}
const UPLOAD_HTML=`<!DOCTYPE html><html><head><meta charset="utf-8"><style>
body{font:16px/1.6 "Microsoft JhengHei","Noto Sans TC",sans-serif;color:#0f262b;background:#f3f6f6;margin:0}
.w{max-width:720px;margin:0 auto;padding:18px}h1{font-size:22px;margin:0 0 4px}.n{color:#587076;font-size:14px}
textarea{width:100%;box-sizing:border-box;height:150px;font:14px monospace;border:2px dashed #0b5563;border-radius:12px;padding:10px;background:#fff}
button{font:700 16px inherit;padding:10px 16px;border-radius:10px;border:0;background:#0b5563;color:#fff;cursor:pointer}
#m{margin-top:10px;padding:10px 12px;border-radius:10px;display:none}.ok{background:#e3f5ec;color:#14532d}.bad{background:#fdecec;color:#b42318}
.last{background:#fff;border-radius:12px;padding:10px 12px;margin:10px 0}</style></head><body><div class="w">
<h1>📥 刀表上傳（院內）</h1><div class="n">HIS 手術工作表 Ctrl+A、Ctrl+C → 點下面的框 → Ctrl+V，就會自動上傳。病歷號、姓名會先在這台電腦去掉才送出。</div>
<div class="last">刀表最後更新：<b id="last">{{LAST}}</b></div>
<textarea id="t" placeholder="在這裡按 Ctrl+V 貼上，會自動上傳"></textarea>
<p><label><input type="checkbox" id="can" checked> 同時送到「黃部長癌症討論會追蹤表」</label>　<label>上傳者 <input id="by" style="width:90px" placeholder="可不填"></label>　<button id="go">上傳</button></p>
<div id="m"></div><p class="n">上傳後，任何人打開刀房調度板就會自動套用（保留已排好的順序，只更新術式；新刀排在最後；刀表上刪除的刀會移除）。</p></div>
<script>
(function(){var TM=/^(\\d{1,2}:\\d{2}|TF\\s*:?\\s*\\d+|[小大]夜:?)$/i,PID=/^\\d{6,10}$/,BED=/^[0-9]{3,5}[A-Za-z]{0,2}$/,TAIL=/常規|急|插刀|科$/;
function fnv(s){var h=2166136261;for(var i=0;i<s.length;i++){h^=s.charCodeAt(i);h=Math.imul(h,16777619)>>>0}return h}
function code(p){var v=/^99\\d{4}$/.test(p)?+p.slice(2)%4096:fnv(p)%4096;var s=String(v);while(s.length<4)s='0'+s;return '99'+s}
function clean(txt){var L=String(txt).replace(/\\r/g,'').split('\\n'),i,j,k,q,n=0,ds={};
 var tab=0;for(i=0;i<L.length;i++)if(L[i].split('\\t').length>=5)tab++;
 if(tab){var out=[];for(i=0;i<L.length;i++){var p=L[i].split('\\t');if(p.length<5){out.push(L[i]);continue}var di=-1;
   for(k=0;k<p.length-1;k++)if(/^\\d{7}$/.test(p[k].replace(/^\\s+|\\s+$/g,''))&&TM.test(p[k+1].replace(/^\\s+|\\s+$/g,''))){di=k;break}
   if(di<0){out.push(L[i]);continue}n++;ds[p[di].replace(/\\s/g,'')]=1;
   for(k=0;k<di;k++){var c=p[k].replace(/^\\s+|\\s+$/g,'');if(!c)continue;if(PID.test(c))p[k]=code(c);else if(!BED.test(c))p[k]='-'}
   out.push(p.join('\\t'))}return{t:out.join('\\n'),n:n,ds:ds}}
 var a=[];for(i=0;i<L.length;i++){var x=L[i].replace(/^\\s+|\\s+$/g,'');if(x)a.push(x)}
 for(j=0;j<a.length-1;j++){if(!(/^\\d{7}$/.test(a[j])&&TM.test(a[j+1])))continue;n++;ds[a[j]]=1;var m=-1;
  for(k=Math.max(0,j-3);k<j;k++)if(PID.test(a[k])){m=k;break}
  if(m>=0){a[m]=code(a[m]);for(q=m+1;q<j;q++)if(!BED.test(a[q]))a[q]='-'}
  else for(q=Math.max(0,j-2);q<j;q++)if(!BED.test(a[q])&&!TAIL.test(a[q])&&!/^\\d/.test(a[q]))a[q]='-'}
 return{t:a.join('\\n'),n:n,ds:ds}}
function md(r){var m=String(r).match(/^(\\d{3})(\\d{2})(\\d{2})$/);return m?(+m[2])+'/'+(+m[3]):r}
function msg(t,ok){var e=document.getElementById('m');e.style.display='block';e.className=ok?'ok':'bad';e.innerHTML=t}
var busy=false;
function up(){if(busy)return;var ta=document.getElementById('t'),raw=ta.value;if(!raw.replace(/\\s/g,'')){msg('框是空的：請先在 HIS 複製，再到框裡按 Ctrl+V');return}
 var c=clean(raw);if(!c.n){msg('看不出是 HIS 手術工作表，請確認複製的內容（要包含手術日與時間）');return}
 var dl=[];for(var d in ds_(c))dl.push(md(d));busy=true;msg('上傳中…（'+c.n+' 筆，'+dl.join('、')+'）',true);
 google.script.run.withSuccessHandler(function(r){busy=false;ta.value='';document.getElementById('last').textContent=r.last+'（院內上傳）';
   var cn=r.can?(r.can.err?'<br>癌症討論會表：送出失敗（'+r.can.err+'）':'<br>癌症討論會表：已送出'):'';
   msg('✓ 已上傳 '+c.n+' 筆（'+dl.join('、')+'）。打開刀房調度板就會自動套用。'+cn,true)})
  .withFailureHandler(function(e){busy=false;msg('上傳失敗：'+(e&&e.message||e)+'。請再試一次。')})
  .apiUpload(c.t,document.getElementById('can').checked?raw:'',document.getElementById('by').value)}
function ds_(c){return c.ds}
document.getElementById('go').onclick=up;
document.getElementById('t').addEventListener('paste',function(){setTimeout(up,60)});
try{var b=localStorage.getItem('upBy');if(b)document.getElementById('by').value=b;document.getElementById('by').onchange=function(){localStorage.setItem('upBy',this.value)}}catch(e){}
document.getElementById('t').focus();
window.__clean=clean;})();
</script></body></html>`;
// ===== Google 網址版調度板（?page=board）：醫院電腦擋 github.io 時用。內容就是 index.html，
//       但和 GitHub Pages 版一樣只走公開的 HTTP 資料管道（不用 google.script.run，不多開伺服器函式）；
//       auth.js 由這裡內嵌（Google 登入按鈕在這裡不能用，要登入請用科務入口「👤 我的帳號」產生的登入碼）。調度板本身不用登入。
// 調度板內容直接抓 GitHub main 的 index.html（快取 10 分鐘，分段存），不依賴專案裡的 index 檔（曾發生專案裡的 index 停在舊版）；抓不到才用專案裡的
function indexHtml_(){const C=CacheService.getScriptCache();try{const n=+C.get('idx:n')||0;if(n){const ks=[];for(let i=0;i<n;i++)ks.push('idx:'+i);const g=C.getAll(ks);let s='';for(const k of ks){if(g[k]==null){s='';break}s+=g[k]}if(s)return s}}catch(e){}
  try{const r=UrlFetchApp.fetch(GH_RAW+'index.html?t='+Date.now(),{muteHttpExceptions:true});if(r.getResponseCode()===200){const s=r.getContentText('UTF-8');if(s.indexOf('泌尿科刀房調度板')>0){try{const SZ=25000,o={};let k=0;for(let i=0;i<s.length;i+=SZ)o['idx:'+(k++)]=s.slice(i,i+SZ);o['idx:n']=String(k);C.putAll(o,600)}catch(e){}return s}}}catch(e){}
  return HtmlService.createHtmlOutputFromFile('index').getContent()}
function boardPage_(){let html=indexHtml_(),auth='';
  try{const C=CacheService.getScriptCache();auth=C.get('authjs')||'';if(!auth){const r=UrlFetchApp.fetch(GH_RAW+'auth.js',{muteHttpExceptions:true});if(r.getResponseCode()===200){auth=r.getContentText();if(auth.length<95000)C.put('authjs',auth,3600)}}}catch(e){}
  html=html.replace(/<script src="auth\.js[^"]*"><\/script>/,auth?'<script>'+auth.replace(/<\/script/gi,'<\\/script')+'</script>':'');
  html=html.replace('<head>','<head><script>window.OR_HOST="gas";window.OR_EXEC='+JSON.stringify(ScriptApp.getService().getUrl())+';</script>');
  return HtmlService.createHtmlOutput(html).setTitle('泌尿科刀房調度板').addMetaTag('viewport','width=device-width, initial-scale=1').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)}
function menuPage_(){const u=ScriptApp.getService().getUrl(),esc=s=>String(s).replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
  const b=(href,t,d)=>'<a target="_top" href="'+esc(href)+'"><b>'+t+'</b><span>'+d+'</span></a>';
  return HtmlService.createHtmlOutput('<!DOCTYPE html><html><head><meta charset="utf-8"><style>body{font:16px/1.6 "Microsoft JhengHei","Noto Sans TC",sans-serif;background:#f3f6f6;color:#0f262b;margin:0}.w{max-width:560px;margin:0 auto;padding:22px 18px}h1{font-size:24px;margin:0 0 14px}a{display:block;background:#fff;border:1px solid #d5e3e5;border-radius:14px;padding:14px 16px;margin-bottom:10px;text-decoration:none;color:inherit}a b{display:block;font-size:19px;color:#0b5563}a span{color:#587076;font-size:14.5px}</style></head><body><div class="w"><h1>泌尿科刀房調度板</h1>'
    +b(u+'?page=board','🗓️ 刀房調度板（Google 版）','醫院電腦打不開 github.io 時用。看刀表、排刀、急診刀都可以。')
    +b(u+'?page=upload','📥 上傳刀表','HIS 手術工作表複製後貼上，自動上傳。')
    +b(PAGES_URL,'🌐 刀房調度板（一般版）','手機、個人電腦用這個：'+PAGES_URL)
    +'</div></body></html>').setTitle('泌尿科刀房調度板').addMetaTag('viewport','width=device-width, initial-scale=1').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)}
function doGet(e){const p=(e&&e.parameter)||{};
  if(p.page==='upload')return uploadPage_();
  if(p.page==='board')return boardPage_();
  const tk=()=>TOKEN&&p.token!==TOKEN,run=f=>{try{return out({ok:true,value:f()})}catch(x){return out({ok:false,err:String(x&&x.message||x)})}},deny=()=>out({ok:false,err:'auth'});
  if(p.action||p.key||p.keys){if(tk())return out({ok:false,err:'token'})}
  if(p.action==='authcfg')return out({ok:true,value:{on:authOn(),client:GOOGLE_CLIENT_ID}});
  if(p.action==='editors')return run(()=>authEditors_(p.sess));
  if(p.action==='me')return run(()=>{const em=sessEmail_(p.sess);return em?meFor_(em):null});
  if(p.action==='cancerScan'){if(!canPHI_(p.sess))return deny();return run(cancerDriveScan)}
  if(p.action==='meetingScan')return run(meetingScan);
  if(p.action==='cancerStatus')return run(cancerStatus);
  if(p.action==='handoff'){if(!canPHI_(p.sess))return deny();return run(()=>handoffFor_(p.dates))}
  // ?keys=a,b,c：一次讀多個鍵（入口網頁載入用，只讀一次 data 工作表）；沒權限的鍵回 null 並列在 denied
  if(p.keys){const ks=String(p.keys).split(',').slice(0,40),denied=ks.filter(k=>PRIVATE_KEY.test(k));
    const v=dataGetMany_(ks.filter(k=>denied.indexOf(k)<0));denied.forEach(k=>v[k]=null);return out({ok:true,value:v,denied})}
  if(p.key){const r=readKey_(p.key,p.sess);return r.deny?deny():out({ok:true,value:r.v})}
  // 不再提供 Apps Script 版網頁：HtmlService 網頁可讓任何人呼叫所有伺服器函式（包含讀交班單原文），一律改用 GitHub Pages
  return menuPage_()}
function doPost(e){const b=JSON.parse(e.postData.contents);if(TOKEN&&b.token!==TOKEN)return out({ok:false,err:'token'});
  const run=f=>{try{return out({ok:true,value:f()})}catch(x){return out({ok:false,err:String(x&&x.message||x)})}};
  if(b.action==='login')return run(()=>authLogin_(b.id_token));
  if(b.action==='claim')return run(()=>authClaim_(b.code));
  if(b.action==='logout')return run(()=>{authLogout_(b.sess);return true});
  if(b.action==='code')return run(()=>authCode_(b.sess));
  if(b.action==='prefs')return run(()=>authPrefs_(b.sess,b));
  if(b.action==='editors')return run(()=>authEditors_(b.sess,b.list));
  if(b.action==='cancerImport'){return run(()=>cancerImport_(b.text))}
  const k=String(b.key||'');if(!k||PRIVATE_KEY.test(k))return out({ok:false,err:'key'});
  dataSet_(k,JSON.stringify(b.value));return out({ok:true})}
// ===== 黃部長交班單 → 調度板術前資料（即時讀取，只回傳數值，不回傳姓名、病歷號、床號，也不存進同步資料） =====
const HO_DOC_ID='1MUwbZVat96vpyil2ahjP6lIv-NzW_MHbOC99EXTx7Yk';
function hoDocId(){try{const v=JSON.parse(dataGet_('common/handoff')||'null');const m=v&&String(v.u||'').match(/document\/d\/([\w-]{20,})/);if(m)return m[1]}catch(e){}return HO_DOC_ID}
// 讀交班單所有分頁裡 7 欄的表格 → [[床號姓名病歷號, PH, PI, Drain, Treatment, Exam/lab, Plan], …]
function hoRows_(){const doc=DocumentApp.openById(hoDocId());const bodies=[];
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
function handoffFor_(dates){dates=String(dates||'').split(',').filter(d=>/^\d{4}-\d{2}-\d{2}$/.test(d)).slice(0,7);
  const C=CacheService.getScriptCache(),k='ho:'+dates.join(',');const h=C.get(k);if(h)return JSON.parse(h);
  const r={at:Date.now(),list:hoParse(hoRows_(),dates)};try{C.put(k,JSON.stringify(r),300)}catch(e){}return r}
function apiHandoff_(dates){return JSON.stringify(handoffFor_(dates))}
// 在編輯器手動執行一次：授權讀 Google 文件並確認讀得到（只記錄筆數）
function hoTest(){const rows=hoRows_();const d=new Date(),iso=x=>Utilities.formatDate(x,Session.getScriptTimeZone(),'yyyy-MM-dd');const ds=[0,1,2,3].map(i=>iso(new Date(d.getTime()+i*864e5)));Logger.log('交班單表格列數 '+rows.length+'；今天起 4 天對到 '+hoParse(rows,ds).length+' 筆手術')}

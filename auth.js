/* 泌尿科網站共用：Google 帳號登入、核可狀態、個人偏好（刀房調度板、科務入口、晨會都載入這個檔）
   - 伺服器（Code.gs）決定是否啟用登入：?action=authcfg → {on, client}。沒啟用時一切照舊。
   - 登入：Google Identity Services 按鈕 → ID token → POST action=login → 30 天登入憑證（存本機 or:auth）。
   - 之後每個請求帶 sess（AUTH.q() 加在網址、AUTH.body() 加在 POST）。
   - 核可帳號＝黃部長交班單的共用對象，才讀得到刀表、急診刀等病人資料。
   - 偏好 prefs（主題、我是誰、入口預設分頁、調度板預設模式…）存在伺服器 user/<email>，本機也留一份。
   使用：AUTH.init(syncUrl, token).then(...)；監聽 window 'auth:change'。 */
(function(){
const LSK='or:auth',CFGK='or:authcfg';
const lsGet=k=>{try{return JSON.parse(localStorage.getItem(k))}catch(e){return null}},lsSet=(k,v)=>{try{v==null?localStorage.removeItem(k):localStorage.setItem(k,JSON.stringify(v))}catch(e){}};
const esc=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
let url='',tok='',cfg=lsGet(CFGK),S=lsGet(LSK);if(S&&S.exp&&S.exp<Date.now())S=null;
const qs=()=>(url.includes('?')?'&':'?');
const emit=()=>{applyTheme();window.dispatchEvent(new CustomEvent('auth:change'))};
async function api(action,extra){const r=await fetch(url,{method:'POST',headers:{'Content-Type':'text/plain;charset=utf-8'},body:JSON.stringify(Object.assign({token:tok,action},extra||{}))});const j=await r.json();if(!j.ok)throw new Error(j.err||'fail');return j.value}
function applyTheme(){const t=(AUTH.prefs().theme)||'';document.documentElement.style.colorScheme=t==='light'?'light':t==='dark'?'dark':''}
/* 登出或失去權限時，清掉本機暫存的病人資料（刀表、急診刀、入口快取） */
function wipePHI(){try{Object.keys(localStorage).forEach(k=>{if(/^or:(days\/|er\/list|hubcache)/.test(k))localStorage.removeItem(k)})}catch(e){}}
let gisP=null;
const loadGIS=()=>gisP||(gisP=new Promise((ok,no)=>{const s=document.createElement('script');s.src='https://accounts.google.com/gsi/client';s.async=true;s.onload=ok;s.onerror=()=>{gisP=null;no(new Error('gis'))};document.head.appendChild(s)}));
const inApp=()=>/Line\/|FBAN|FBAV|Instagram|MicroMessenger/i.test(navigator.userAgent);
const extLink=()=>location.origin+location.pathname+'?openExternalBrowser=1'+location.hash;
function done(v){S={token:v.token,exp:v.exp,me:v.me};lsSet(LSK,S);if(!v.me.ok)wipePHI();emit()}
function modal(){const old=document.querySelector('.authm');if(old)old.remove();
 const m=document.createElement('div');m.className='authm';m.style.cssText='position:fixed;inset:0;background:rgba(0,0,0,.45);display:flex;align-items:center;justify-content:center;z-index:10001;padding:12px';
 m.innerHTML=`<div style="background:var(--card,#fff);color:var(--ink,#0f262b);border-radius:16px;max-width:400px;width:100%;padding:18px 20px;font-size:15px;line-height:1.6;box-shadow:0 10px 40px rgba(0,0,0,.3)">
  <b style="font-size:18px">🔑 用 Google 帳號登入</b>
  <p style="margin:6px 0 12px;color:var(--mut,#587076);font-size:13.5px">刀表、急診刀、交班單等含病人資料的內容，只開放<b>黃部長交班單有共用的 Google 帳號</b>。登入後也會記住你的偏好設定。</p>
  ${inApp()?`<p><b>你現在在 LINE 裡面</b>，Google 不允許在這裡登入，請先改用手機瀏覽器開啟。</p><a href="${esc(extLink())}" style="display:inline-block;background:#0b5563;color:#fff;border-radius:10px;padding:8px 14px;text-decoration:none">用瀏覽器開啟</a>`:`<div id="gbtn" style="min-height:44px">載入 Google 登入中…</div>`}
  <details style="margin-top:12px"><summary style="cursor:pointer;font-size:13.5px">在主畫面 App 裡跳不出 Google 登入？用登入碼</summary>
   <p style="font-size:13px;color:var(--mut,#587076);margin:6px 0">先在瀏覽器登入，到科務入口「👤 我的帳號」按「產生登入碼」，再把 8 碼輸入這裡（10 分鐘內有效）。</p>
   <div style="display:flex;gap:6px"><input id="acode" placeholder="例：AB12CD34" maxlength="9" autocomplete="one-time-code" style="flex:1;font:inherit;padding:6px 8px;border:1px solid var(--line,#ccc);border-radius:8px;text-transform:uppercase;background:var(--card,#fff);color:inherit"><button data-aclaim style="font:inherit;border:1px solid #0b5563;background:#0b5563;color:#fff;border-radius:8px;padding:6px 12px">登入</button></div></details>
  <p id="amsg" style="color:#c0392b;font-size:13.5px;margin:8px 0 0"></p>
  <button data-aclose style="font:inherit;margin-top:10px;border:1px solid var(--line,#ccc);background:transparent;color:inherit;border-radius:10px;padding:6px 14px">關閉</button></div>`;
 document.body.appendChild(m);const msg=t=>{const e=m.querySelector('#amsg');if(e)e.textContent=t};
 const fail=e=>msg(/token|aud/.test(e.message)?'Google 登入驗證失敗，請再試一次':/off/.test(e.message)?'登入功能尚未啟用':'登入失敗：'+e.message);
 m.addEventListener('click',async e=>{if(e.target===m||e.target.closest('[data-aclose]'))m.remove();
  if(e.target.closest('[data-aclaim]')){msg('');try{done(await api('claim',{code:m.querySelector('#acode').value}));m.remove()}catch(x){msg(/code/.test(x.message)?'登入碼錯誤或已過期':'登入失敗：'+x.message)}}});
 if(!inApp())loadGIS().then(()=>{const g=window.google&&google.accounts&&google.accounts.id;if(!g)throw new Error('gis');
   g.initialize({client_id:cfg.client,callback:async r=>{msg('驗證中…');try{done(await api('login',{id_token:r.credential}));m.remove()}catch(x){fail(x)}},ux_mode:'popup',itp_support:true});
   const b=m.querySelector('#gbtn');b.textContent='';g.renderButton(b,{theme:'filled_blue',size:'large',text:'signin_with',shape:'pill',locale:'zh-TW',width:280})})
  .catch(()=>{const b=m.querySelector('#gbtn');if(b)b.textContent='Google 登入載入失敗，請確認網路後再試'})}
const AUTH=window.AUTH={
 ready:null,
 init(u,t){url=u||'';tok=t||'';applyTheme();if(!url&&!cfg)cfg={on:false};
  this.ready=(async()=>{if(!url)return;const was=!!(cfg&&cfg.on);try{const r=await fetch(url+qs()+'action=authcfg&token='+encodeURIComponent(tok));const j=await r.json();if(j.ok&&j.value){cfg=j.value;lsSet(CFGK,cfg)}else if(!cfg)cfg={on:false}}catch(e){if(!cfg)cfg={on:false}}
   if(was!==!!cfg.on)emit();if(cfg.on&&S)this.refresh()})();
  return this.ready},
 on:()=>!!(cfg&&cfg.on),
 known:()=>!!cfg,
 signedIn:()=>!!(S&&S.token),
 me:()=>(S&&S.me)||null,
 /* 需要登入／核可才能看病人資料：啟用登入且（沒登入或不是核可帳號） */
 need:()=>!!(cfg&&cfg.on)&&!(S&&S.me&&S.me.ok),
 q:()=>S&&S.token?'&sess='+encodeURIComponent(S.token):'',
 body:o=>(S&&S.token?Object.assign(o,{sess:S.token}):o),
 prefs:()=>(S&&S.me&&S.me.prefs)||lsGet('or:prefs')||{},
 name:()=>{const m=S&&S.me;return m?(m.name||m.email.split('@')[0]):''},
 login(){if(!cfg||!cfg.on){alert('登入功能尚未啟用');return}modal()},
 async logout(){if(!confirm('登出這個裝置？本機暫存的刀表與急診刀資料也會一起清除。'))return;const t=S&&S.token;S=null;lsSet(LSK,null);wipePHI();if(t)api('logout',{sess:t}).catch(()=>{});emit()},
 async refresh(){if(!S)return;try{const r=await fetch(url+qs()+'action=me&token='+encodeURIComponent(tok)+this.q());const j=await r.json();if(!j.ok)return;
   if(!j.value){S=null;lsSet(LSK,null);wipePHI();emit();return}const was=S.me&&S.me.ok;S.me=j.value;lsSet(LSK,S);if(was&&!S.me.ok)wipePHI();emit()}catch(e){}},
 async savePrefs(p,name){lsSet('or:prefs',Object.assign({},this.prefs(),p));if(!S){applyTheme();return}const v=await api('prefs',{sess:S.token,prefs:Object.assign({},this.prefs(),p),name});S.me=v;lsSet(LSK,S);emit();return v},
 async code(){if(!S)throw new Error('auth');return api('code',{sess:S.token})},
 /* 伺服器回 err:'auth'（登入過期或權限被取消）時呼叫：清掉登入並請使用者重新登入 */
 expired(){if(S){S=null;lsSet(LSK,null)}wipePHI();emit();if(cfg&&cfg.on)modal()},
 /* 標題列的小按鈕：已登入顯示頭像＋名字（連到我的帳號），沒登入顯示「登入」 */
 badge(hubHref){if(!cfg||!cfg.on)return '';const m=this.me();const st='display:inline-flex;align-items:center;gap:6px;color:#fff;background:rgba(255,255,255,.14);border:1px solid rgba(255,255,255,.3);border-radius:99px;padding:2px 10px 2px 3px;font-size:13px;text-decoration:none;white-space:nowrap';
  if(!m)return `<a href="#" data-alogin style="${st};padding:3px 12px">🔑 登入</a>`;
  const pic=m.pic?`<img src="${esc(m.pic)}" alt="" width="22" height="22" referrerpolicy="no-referrer" style="border-radius:50%">`:'<span style="width:22px;text-align:center">👤</span>';
  return `<a href="${esc(hubHref||'hub.html#me')}" style="${st}" title="${esc(m.email)}">${pic}${esc(this.name())}${m.ok?'':' ⚠'}</a>`},
 /* 鎖住的內容：給頁面放在需要核可的區塊 */
 lockCard(what){const m=this.me();return `<section style="background:var(--card,#fff);border:1px solid var(--line,#d8e2e3);border-radius:14px;padding:18px;max-width:560px;margin:10px auto;text-align:center">
  <div style="font-size:36px">🔒</div><h2 style="margin:4px 0 8px;justify-content:center">${esc(what||'這裡有病人資料')}</h2>
  ${m?`<p>你用 <b>${esc(m.email)}</b> 登入，但這個帳號<b>還沒有權限</b>。</p><p style="color:var(--mut,#587076);font-size:13.5px">權限名單＝黃部長交班單（Google 文件）有共用的帳號。請總醫師把交班單共用給這個帳號（檢視者即可），約 10 分鐘內生效；或改用已共用的帳號登入。</p><button data-alogout style="font:inherit;border:1px solid var(--line,#ccc);background:transparent;color:inherit;border-radius:10px;padding:8px 14px;margin-top:6px">換帳號</button> <button data-arefresh style="font:inherit;border:1px solid var(--line,#ccc);background:transparent;color:inherit;border-radius:10px;padding:8px 14px;margin-top:6px">重新檢查</button>`
   :`<p>請用<b>黃部長交班單有共用的 Google 帳號</b>登入後查看。</p><button data-alogin style="font:inherit;border:0;background:#0b5563;color:#fff;border-radius:10px;padding:10px 18px;margin-top:6px">🔑 用 Google 帳號登入</button>`}</section>`}
};
document.addEventListener('click',e=>{const t=e.target.closest&&e.target.closest('[data-alogin],[data-alogout],[data-arefresh]');if(!t)return;e.preventDefault();e.stopPropagation();
 if(t.hasAttribute('data-alogin'))AUTH.login();else if(t.hasAttribute('data-alogout')){const t0=S&&S.token;S=null;lsSet(LSK,null);wipePHI();if(t0)api('logout',{sess:t0}).catch(()=>{});emit();AUTH.login()}else AUTH.refresh()},true);
applyTheme();
})();

// 從 GitHub 自動更新本 Apps Script 專案（不用再手動貼程式）
// 第一次：照「開發說明.md › 自動更新」設定好後，在編輯器執行 setupAutoUpdate() 一次並授權。
// 之後：每小時檢查 GitHub main 分支，有變更就自動
//   1) 更新專案裡的 Code、CancerSheet、Updater、index 四個檔（其他檔案與 appsscript.json 不動）
//   2) 建立新版本並更新現有的網頁應用程式部署（/exec 網址不變）
//   3) CancerSheet 有變就在 1 分鐘後自動執行 setupCancerSheet()
// 也可以在試算表選單「🔄 程式更新」手動立即更新。
// 需要：script.google.com/home/usersettings 開啟「Google Apps Script API」，
//       appsscript.json 列出 oauthScopes（見開發說明）。

const GH_RAW = 'https://raw.githubusercontent.com/ttimcom-hue/Uro-OR/main/';
// [repo 路徑, 專案內檔名, 類型, 用來找舊檔的特徵（行首的函式宣告；用 ^ 避免比對到本檔這幾行字）]
const GH_FILES = [
  ['Code.gs', 'Code', 'SERVER_JS', /^function doGet\b/m],
  ['cancer/CancerSheet.gs', 'CancerSheet', 'SERVER_JS', /^function setupCancerSheet\b/m],
  ['Updater.gs', 'Updater', 'SERVER_JS', /^function updateFromGitHub\b/m],
  ['index.html', 'index', 'HTML', null],
];
const GH_API = 'https://script.googleapis.com/v1/projects/';
// appsscript.json 需要的權限。只在手動執行（updateNow／setupAutoUpdate）時補上，
// 因為加了權限之後要有人當場重新授權，否則網頁應用程式與排程會停擺。
const GH_SCOPES = [
  'https://www.googleapis.com/auth/spreadsheets',
  'https://www.googleapis.com/auth/script.external_request',
  'https://www.googleapis.com/auth/script.projects',
  'https://www.googleapis.com/auth/script.deployments',
  'https://www.googleapis.com/auth/script.scriptapp',
  'https://www.googleapis.com/auth/script.container.ui',
  'https://www.googleapis.com/auth/documents',
  'https://www.googleapis.com/auth/drive.readonly',
  'https://www.googleapis.com/auth/drive.file',   // 癌症討論會：PDF 暫時轉成文件讀文字，讀完刪除
];

function onOpen() {
  SpreadsheetApp.getUi().createMenu('🔄 程式更新')
    .addItem('立即從 GitHub 更新', 'updateNow')
    .addItem('重新套用癌症討論會表格式', 'ghRunCancerSetup')
    .addItem('開啟每小時自動更新', 'setupAutoUpdate')
    .addToUi();
}

// 第一次設定：建立每小時自動更新，並馬上更新一次
function setupAutoUpdate() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'ghAutoUpdate')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('ghAutoUpdate').timeBased().everyHours(1).create();
  return updateNow();
}

function updateNow() {
  const msg = updateFromGitHub(true);
  ghToast_(msg);
  return msg;
}

function ghAutoUpdate() {
  try { updateFromGitHub(false); } catch (e) { console.log(e); }
  // 每小時順便做癌症討論會表的自動加入與交班單同步（編輯時沒觸發到的也會補上）
  try { if (typeof cancerDriveScan === 'function') cancerDriveScan(); } catch (e) { console.log(e); }
  try { if (typeof cancerSync === 'function') cancerSync(); } catch (e) { console.log(e); }
  // 晨會：掃 Drive「W2 科部晨會」資料夾，更新 meeting/drive（入口「晨會」分頁）
  try { if (typeof meetingScan === 'function') meetingScan(); } catch (e) { console.log(e); }
}

function ghRunCancerSetup() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'ghRunCancerSetup')
    .forEach(t => ScriptApp.deleteTrigger(t));
  setupCancerSheet();
  ghLog_('已套用癌症討論會表格式');
}

function updateFromGitHub(withScopes) {
  const id = ScriptApp.getScriptId();
  const cur = ghApi_('get', id + '/content').files;

  // 1) 抓 GitHub 上的新版，任何一個抓不到就整個放棄（避免把專案改壞）
  const fresh = GH_FILES.map(([path, name, type, sig]) => {
    const r = UrlFetchApp.fetch(GH_RAW + path, { muteHttpExceptions: true });
    if (r.getResponseCode() !== 200) throw new Error('抓不到 GitHub 的 ' + path + '（' + r.getResponseCode() + '），請確認已合併到 main');
    const src = r.getContentText();
    if (sig && !sig.test(src)) throw new Error(path + ' 內容不對，已停止更新');
    return { path, name, type, sig, src };
  });

  // 2) 對應到專案裡現有的檔案：先比檔名，再比特徵字串；找不到就新增
  const files = cur.map(f => ({ name: f.name, type: f.type, source: f.source }));
  const used = new Set();
  let changed = [], cancerChanged = false;
  fresh.forEach(n => {
    const same = f => ghName_(f.name) === ghName_(n.name);
    // 順序：同名同類型 → 內容有特徵字串 → 同名但類型不同（多半是建錯成 HTML，直接接管）→ 新增
    let i = files.findIndex((f, k) => !used.has(k) && f.type === n.type && same(f));
    if (i < 0 && n.sig) i = files.findIndex((f, k) => !used.has(k) && f.type === n.type && n.sig.test(f.source));
    if (i < 0) i = files.findIndex((f, k) => !used.has(k) && same(f));
    if (i < 0) {
      if (files.some(same)) throw new Error('專案裡已有同名檔「' + n.name + '」無法對應，現有檔案：' + ghList_(files));
      files.push({ name: n.name, type: n.type, source: '' }); i = files.length - 1;
    }
    used.add(i);
    if (files[i].type !== n.type) { files[i].type = n.type; files[i].source = ''; }
    if (files[i].source.trim() !== n.src.trim()) {
      files[i].source = n.src;
      changed.push(n.path);
      if (n.name === 'CancerSheet') cancerChanged = true;
    }
  });
  let needAuth = false;
  if (withScopes) {
    const m = files.find(f => f.type === 'JSON' && f.name === 'appsscript');
    if (m) {
      const j = JSON.parse(m.source);
      const have = j.oauthScopes || [];
      const add = GH_SCOPES.filter(x => have.indexOf(x) < 0);
      if (add.length) {
        j.oauthScopes = have.concat(add);
        m.source = JSON.stringify(j, null, 2);
        changed.push('appsscript.json（新增權限）');
        needAuth = true;
      }
    }
  }
  if (!changed.length) { ghLog_('已是最新'); return '已是最新，不用更新'; }

  // 3) 寫回專案、建新版本、更新網頁應用程式部署
  try { ghApi_('put', id + '/content', { files }); }
  catch (e) { throw new Error(e.message + '\n現有檔案：' + ghList_(cur)); }
  const ver = ghApi_('post', id + '/versions', { description: 'GitHub 自動更新 ' + ghNow_() }).versionNumber;
  const deps = (ghApi_('get', id + '/deployments').deployments || [])
    .filter(d => d.deploymentConfig && d.deploymentConfig.versionNumber &&
      (d.entryPoints || []).some(e => e.entryPointType === 'WEB_APP'));
  deps.forEach(d => ghApi_('put', id + '/deployments/' + d.deploymentId, {
    deploymentConfig: { scriptId: id, versionNumber: ver, manifestFileName: 'appsscript',
      description: d.deploymentConfig.description || ('GitHub 自動更新 v' + ver) },
  }));

  // 4) 新程式要在下一次執行才生效，所以 1 分鐘後再跑 setupCancerSheet
  if (cancerChanged) ScriptApp.newTrigger('ghRunCancerSetup').timeBased().after(60 * 1000).create();

  const msg = '已更新：' + changed.join('、') + '（版本 ' + ver + '，部署 ' + deps.length + ' 個）' +
    (cancerChanged ? '；1 分鐘後自動套用癌症討論會表格式' : '') +
    (needAuth ? '；⚠ 權限有變，請馬上在編輯器執行 cancerSync 並允許授權' : '');
  ghLog_(msg);
  return msg;
}

function ghApi_(method, path, body) {
  const opt = { method, muteHttpExceptions: true, contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() } };
  if (body) opt.payload = JSON.stringify(body);
  const r = UrlFetchApp.fetch(GH_API + path, opt);
  const t = r.getContentText();
  if (r.getResponseCode() >= 300) {
    const hint = /has not been used|disabled|usersettings/i.test(t)
      ? '請到 script.google.com/home/usersettings 開啟「Google Apps Script API」' : '';
    throw new Error('Apps Script API 錯誤 ' + r.getResponseCode() + ' ' + hint + '\n' + t.slice(0, 500));
  }
  return t ? JSON.parse(t) : {};
}
function ghName_(s) { return String(s || '').trim().replace(/\.(gs|html)$/i, '').toLowerCase(); }
function ghList_(files) { return files.map(f => f.name + (f.type === 'HTML' ? '.html' : f.type === 'JSON' ? '.json' : '.gs')).join('、'); }
function ghNow_() { return Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy/MM/dd HH:mm'); }
function ghLog_(m) {
  console.log(m);
  PropertiesService.getScriptProperties().setProperty('ghLastUpdate', ghNow_() + ' ' + m);
}
function ghToast_(m) {
  try { SpreadsheetApp.getActiveSpreadsheet().toast(m, '程式更新', 10); } catch (e) { /* 從編輯器執行時沒有試算表畫面 */ }
}

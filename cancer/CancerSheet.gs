// 癌症討論會追蹤表：重建公式分頁＋自動加入新個案＋同步交班單（Google Apps Script）
// 放在跟 Code.gs 同一個 Apps Script 專案（Updater.gs 會自動從 GitHub 更新並執行 setupCancerSheet）。
// setupCancerSheet 只會改「公式／格式」，不會動到使用者輸入的資料：
//   - 個案追蹤：只重寫 G、H 兩欄公式與格式，A–F、I–Q 的資料不動
//   - 其他公式分頁整頁重寫，但保留 下次討論會!B1（手動改的日期）、刀表候選!B2:B3（主刀、關鍵字）、
//     說明!B1:B3（交班單網址、分頁名稱、自動加入開關）
//   - 刀表匯入：只寫 A1 標題，A2 以下的原始刀表不動
// cancerSync()：刀表新個案自動加入個案追蹤＋把注意事項寫進交班單（編輯追蹤表時、每小時、匯入刀表時執行）
// CANCER_ID 定義在 Code.gs。交班單網址只存在試算表「說明」B1，不放在 repo（repo 是公開的）。

const CS = {
  NEXT: '下次討論會', TODO: '注意事項', TRK: '個案追蹤', DONE: '已討論個案', WEEK: '週會名單',
  PEND: '待討論清單', CAND: '刀表候選', RAW: '刀表匯入', PARSE: '刀表解析', HELP: '說明',
};
const CS_RENAMED = { '會後待辦': CS.TODO };   // 舊分頁名 → 新分頁名
// 前四頁是平常會用的，其餘是後台（隱藏，需要時從左下「所有工作表」打開）
const CS_ORDER = [CS.NEXT, CS.TODO, CS.TRK, CS.HELP, CS.DONE, CS.WEEK, CS.PEND, CS.CAND, CS.RAW, CS.PARSE];
const CS_HIDDEN = [CS.DONE, CS.WEEK, CS.PEND, CS.CAND, CS.RAW, CS.PARSE];
const CS_STATUS = ['待討論', '已排入', '已討論-待手術', '已開刀', '待確認', '暫緩/改期', '不需手術'];
const CS_STATUS_COLOR = {
  '待討論': '#F4CCCC', '已排入': '#CFE2F3', '已討論-待手術': '#FFF2CC', '已開刀': '#D9EAD3',
  '待確認': '#FCE5CD', '暫緩/改期': '#E0E0E0', '不需手術': '#D9D2E9',
};
// 「注意事項」要列出的狀態（已討論但還沒開刀）
const CS_OPEN = ['已討論-待手術', '暫緩/改期', '待確認'];
const CS_TRK_HEAD = ['狀態', '姓名', '病歷號', '手術日期', '預定手術', '討論會日期', '最晚討論日', '提醒',
  '主治VS', '診斷 / 影像分期', '討論結論', '術後結果 / 病理', '報告醫師', '來源 / 備註',
  '會後待辦', '負責人', '待辦期限'];
const CS_KEYWORDS = '(?i)davinci|prostatectomy|nephr\\w*ectomy|nephrouret|\\bRNU\\b|cystectomy|RARC|adrenal|thrombect|retroperitoneal';
const CS_PARSE_ROWS = 400;   // 一次刀表最多解析幾筆手術
const CS_RAW_ROWS = 6000;    // 刀表匯入最多幾行
const HDR_BG = '#EFEFEF', RED = '#C00000';
const CS_TAB_COLOR = {
  [CS.NEXT]: '#3D85C6', [CS.TODO]: '#E06666', [CS.TRK]: '#6AA84F', [CS.HELP]: '#CCCCCC', [CS.DONE]: '#CCCCCC',
  [CS.WEEK]: '#CCCCCC', [CS.PEND]: '#CCCCCC', [CS.CAND]: '#CCCCCC', [CS.RAW]: '#FFD966', [CS.PARSE]: '#CCCCCC',
};
const CS_DOC_TAB = '癌症討論會';   // 交班單裡要寫入的分頁名稱（預設值，可在 說明!B2 改）
// 個案追蹤：預設收合的欄 [起始欄, 欄數]（按欄上方的 ＋ 展開）。G 最晚討論日、L–N 術後病理／報告醫師／備註
const CS_TRK_FOLD = [[7, 1], [12, 3]];
const CS_FINISHED = ['已開刀', '不需手術'];

function q_(name) { return "'" + name + "'"; }

// ---------- 公式 ----------
function csTrackerG_(r) { return `=IF(D${r}="","",D${r}-(MOD(WEEKDAY(D${r})-6,7)+1))`; }
function csTrackerH_(r) {
  return `=IF(D${r}="","",IF(A${r}="待討論",IF(G${r}<TODAY(),"⚠ 已過最晚討論日",IF(G${r}-TODAY()<=14,"⚠ 需排入討論會","")),` +
    `IF(AND(OR(A${r}="已討論-待手術",A${r}="已排入"),D${r}<TODAY()),"手術日已過，請更新","")))`;
}
// 下次討論會：上半「已排入這次」，下半「還沒討論、可挑」（依手術日）。欄：提示,手術日,姓名,病歷號,主治,預定手術,診斷
function csNextFormula_() {
  const T = q_(CS.TRK), E = 'HSTACK("","","","","","")';
  return `=LET(t,${T}!A2:Q,d,$B$1,st,INDEX(t,,1),nm,INDEX(t,,2),md,INDEX(t,,6),ld,INDEX(t,,7),` +
    `c,HSTACK(INDEX(t,,4),nm,INDEX(t,,3),INDEX(t,,9),INDEX(t,,5),INDEX(t,,10)),` +
    `tag,ARRAYFORMULA(IF(ld="","未定手術日",IF(ld<d,"⚠ 已過最晚討論日",IF(ld=d,"⚠ 這次一定要排",IF(ld<=d+7,"下次可排",""))))),` +
    `na,COUNTIFS(md,d,nm,"<>"),nb,COUNTIFS(st,"待討論",nm,"<>"),` +
    `a,IF(na=0,HSTACK("（還沒有排入）",${E}),FILTER(HSTACK(st,c),md=d,nm<>"")),` +
    `b,IF(nb=0,HSTACK("（沒有待討論個案）",${E}),SORT(FILTER(HSTACK(tag,c),st="待討論",nm<>""),2,TRUE)),` +
    `VSTACK(HSTACK("▶ 已排入 "&TEXT(d,"m/d")&" 討論會："&na&" 人",${E}),a,HSTACK("",${E}),` +
    `HSTACK("▶ 還沒討論、可以挑的個案："&nb&" 人（⚠ 優先）",${E}),b))`;
}
// 注意事項：已討論、還沒開刀。欄：狀態,手術日,剩幾天,姓名,病歷號,預定手術,討論結論,注意事項/會後待辦,負責/期限
function csTodoFormula_() {
  const T = q_(CS.TRK);
  const cond = CS_OPEN.map(s => `(st="${s}")`).join('+');
  return `=IFERROR(LET(t,${T}!A2:Q,st,INDEX(t,,1),f,SORT(FILTER(t,${cond}),4,TRUE),d,INDEX(f,,4),` +
    `HSTACK(INDEX(f,,1),d,ARRAYFORMULA(IF(d="","",d-TODAY())),INDEX(f,,2),INDEX(f,,3),INDEX(f,,5),INDEX(f,,11),INDEX(f,,15),` +
    `ARRAYFORMULA(TRIM(INDEX(f,,16)&IF(INDEX(f,,17)="",""," "&TEXT(INDEX(f,,17),"m/d")))))),"（目前沒有）")`;
}
function csParseRow2_() {
  const R = q_(CS.RAW), C = q_(CS.CAND), T = q_(CS.TRK);
  return [
    `=IFERROR(FILTER(ROW(${R}!A2:A${CS_RAW_ROWS}),REGEXMATCH(${R}!A2:A${CS_RAW_ROWS}&"","^\\d{8}$")),"")`,
    `=IF(A2="","",IF(REGEXMATCH(INDEX(${R}!A:A,A2+3)&"","^1\\d{6}$"),A2+3,A2+4))`,
    `=IF(A2="","",IF(INDEX(${R}!A:A,B2+3)&""="",B2+5,B2+4))`,
    `=IF(A2="","",INDEX(${R}!A:A,A2)&"")`,
    `=IF(A2="","",INDEX(${R}!A:A,A2+1)&"")`,
    `=IF(A2="","",LET(n,INDEX(${R}!A:A,B2),DATE(INT(n/10000)+1911,MOD(INT(n/100),100),MOD(n,100))))`,
    `=IF(A2="","",LET(t,INDEX(${R}!A:A,B2+1),IF(ISNUMBER(t),TEXT(t,"hh:mm"),t&"")))`,
    `=IF(A2="","",INDEX(${R}!A:A,B2+2)&"")`,
    `=IF(A2="","",IF(C2=B2+4,INDEX(${R}!A:A,B2+3)&"",""))`,
    `=IF(A2="","",INDEX(${R}!A:A,C2+1)&"")`,
    `=IF(A2="","",INDEX(${R}!A:A,C2+3)&"")`,
    `=IF(A2="",FALSE,AND(REGEXMATCH(H2,${C}!$B$2),REGEXMATCH(J2,${C}!$B$3)))`,
    `=IF(A2="","",IF(COUNTIF(${T}!C:C,D2)>0,"已在追蹤表","新個案"))`,
    `=IF(A2="","",J2&" "&G2)`,
  ];
}

// ---------- 主程式 ----------
function setupCancerSheet() {
  const ss = SpreadsheetApp.openById(CANCER_ID);
  ss.setSpreadsheetLocale('zh_TW');
  ss.setSpreadsheetTimeZone('Asia/Taipei');
  Object.entries(CS_RENAMED).forEach(([o, n]) => {
    const sh = ss.getSheetByName(o);
    if (sh && !ss.getSheetByName(n)) sh.setName(n);
  });
  const nx = ss.getSheetByName(CS.NEXT);
  const keep = {
    // 下次討論會!B1：使用者手動改成某個日期（不是公式）且還沒過，就保留
    nextDate: nx && !nx.getRange('B1').getFormula() && nx.getRange('B1').getValue() instanceof Date &&
      nx.getRange('B1').getValue() >= csToday_() ? nx.getRange('B1').getValue() : '',
    surgeon: csGet_(ss, CS.CAND, 'B2'),
    keywords: csGet_(ss, CS.CAND, 'B3'),
    settings: csSettings_(ss),
  };
  CS_ORDER.forEach((name, i) => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name, i);
    ss.setActiveSheet(sh); ss.moveActiveSheet(i + 1);
    sh.setTabColor(CS_TAB_COLOR[name]);
  });
  csNext_(ss.getSheetByName(CS.NEXT), keep.nextDate);
  csTracker_(ss.getSheetByName(CS.TRK));
  csTodo_(ss.getSheetByName(CS.TODO));
  csDone_(ss.getSheetByName(CS.DONE));
  csWeek_(ss.getSheetByName(CS.WEEK));
  csPend_(ss.getSheetByName(CS.PEND));
  csCand_(ss.getSheetByName(CS.CAND), keep.surgeon, keep.keywords);
  csRaw_(ss.getSheetByName(CS.RAW));
  csParse_(ss.getSheetByName(CS.PARSE));
  csHelp_(ss.getSheetByName(CS.HELP), keep.settings);
  ss.setActiveSheet(ss.getSheetByName(CS.NEXT));
  CS_ORDER.forEach(name => {
    const sh = ss.getSheetByName(name);
    if (CS_HIDDEN.includes(name)) sh.hideSheet(); else if (sh.isSheetHidden()) sh.showSheet();
  });
  SpreadsheetApp.flush();
  csEnsureTrigger_();
}

function csToday_() { const d = new Date(); d.setHours(0, 0, 0, 0); return d; }

// ---------- 下次討論會 ----------
function csNext_(sh, nextDate) {
  csReset_(sh);
  sh.getRange('A1').setValue('下次討論會').setFontWeight('bold').setFontSize(13);
  const b1 = sh.getRange('B1');
  if (nextDate) b1.setValue(nextDate); else b1.setFormula('=TODAY()+MOD(5-WEEKDAY(TODAY()),7)');
  b1.setNumberFormat('yyyy/mm/dd（ddd）').setFontWeight('bold').setFontSize(13).setBackground('#FFF2CC');
  sh.getRange('C1').setValue('自動＝下一個週四；要看別週，直接在黃色格打日期（刪掉就恢復自動）').setFontColor('#888888');
  const head = ['提示 / 狀態', '手術日', '姓名', '病歷號', '主治', '預定手術', '診斷 / 影像分期'];
  sh.getRange(3, 1, 1, head.length).setValues([head]);
  csHeader_(sh.getRange(3, 1, 1, head.length));
  sh.getRange('A4').setFormula(csNextFormula_());
  sh.getRange('B4:B').setNumberFormat('m/d（ddd）');
  sh.getRange('A4:G').setVerticalAlignment('top');
  sh.getRange('F4:G').setWrap(true);
  const rng = sh.getRange('A4:G300');
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=LEFT($A4,1)="▶"')
      .setBackground('#CFE2F3').setBold(true).setRanges([rng]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=LEFT($A4,1)="⚠"')
      .setFontColor(RED).setBold(true).setRanges([sh.getRange('A4:A300')]).build(),
  ]);
  sh.setFrozenRows(3);
  csWidths_(sh, [150, 95, 80, 90, 70, 300, 300]);
}

function csGet_(ss, name, a1) {
  const sh = ss.getSheetByName(name);
  return sh ? sh.getRange(a1).getValue() : '';
}
function csReset_(sh) {
  sh.clear(); sh.clearConditionalFormatRules();
  sh.getDataRange().clearDataValidations();
}
function csHeader_(rng) { rng.setFontWeight('bold').setBackground(HDR_BG); }
function csWidths_(sh, widths) { widths.forEach((w, i) => sh.setColumnWidth(i + 1, w)); }
function csStatusRules_(sh, a1) {
  const rng = sh.getRange(a1);
  return Object.entries(CS_STATUS_COLOR).map(([k, c]) =>
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(k).setBackground(c).setRanges([rng]).build());
}
// 已結案（已開刀／不需手術）整列字變淡灰，眼睛只需要看還在進行的個案
function csFinishedRule_(sh, a1, statusCol, firstRow) {
  const f = '=OR(' + CS_FINISHED.map(s => `$${statusCol}${firstRow}="${s}"`).join(',') + ')';
  return SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(f)
    .setFontColor('#A0A0A0').setRanges([sh.getRange(a1)]).build();
}
// 欄群組：先清掉舊群組（重跑 setup 不會越疊越深），再建立並收合
function csFold_(sh, groups) {
  const max = sh.getMaxColumns();
  for (let c = 1; c <= max; c++) {
    const d = sh.getColumnGroupDepth(c);
    if (d > 0) sh.getRange(1, c, 1, 1).shiftColumnGroupDepth(-d);
  }
  sh.setColumnGroupControlPosition(SpreadsheetApp.GroupControlTogglePosition.BEFORE);
  groups.forEach(([c, n]) => {
    sh.getRange(1, c, 1, n).shiftColumnGroupDepth(1);
    sh.getColumnGroup(c, 1).collapse();
  });
}

function csTracker_(sh) {
  const last = Math.max(sh.getLastRow(), 1);
  sh.getRange(1, 1, 1, CS_TRK_HEAD.length).setValues([CS_TRK_HEAD]);
  // 標題列分三段上色：個案資料（灰）／討論會（藍）／會後待辦（紅）
  sh.getRange(1, 1, 1, 17).setFontWeight('bold').setWrap(true).setVerticalAlignment('middle');
  sh.getRange(1, 1, 1, 8).setBackground(HDR_BG);
  sh.getRange(1, 9, 1, 6).setBackground('#CFE2F3');
  sh.getRange(1, 15, 1, 3).setBackground('#F4CCCC');
  // G、H 公式：寫到最後一列 + 50 列，之後新增的列也有公式
  const n = last - 1 + 50, rows = [];
  for (let r = 2; r < 2 + n; r++) rows.push([csTrackerG_(r), csTrackerH_(r)]);
  sh.getRange(2, 7, n, 2).setFormulas(rows);
  sh.getRange('D2:D').setNumberFormat('yyyy/mm/dd');
  sh.getRange('F2:G').setNumberFormat('yyyy/mm/dd');
  sh.getRange('Q2:Q').setNumberFormat('yyyy/mm/dd');
  sh.getRange('H2:H').setFontWeight('bold').setFontColor(RED);
  sh.getRange('A2:Q').setVerticalAlignment('top');
  sh.getRange('J2:L').setWrap(true); sh.getRange('N2:O').setWrap(true);
  sh.getRange('A2:A500').setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(CS_STATUS, true).setAllowInvalid(true).build());
  sh.setConditionalFormatRules([csFinishedRule_(sh, 'B2:Q500', 'A', 2), ...csStatusRules_(sh, 'A2:A500')]);
  sh.setFrozenRows(1); sh.setFrozenColumns(2);
  csWidths_(sh, [110, 90, 90, 90, 230, 100, 100, 130, 70, 260, 360, 220, 70, 240, 340, 90, 90]);
  csFold_(sh, CS_TRK_FOLD);
  // 標題列加篩選（可只看某狀態、某VS）；已有篩選就保留使用者目前的條件
  if (!sh.getFilter()) sh.getRange(1, 1, sh.getMaxRows(), CS_TRK_HEAD.length).createFilter();
}

function csTodo_(sh) {
  csReset_(sh);
  sh.getRange('A1').setValue('已討論、還沒開刀：注意事項').setFontWeight('bold').setFontSize(13);
  sh.getRange('A2').setValue('自動產生（會同步到交班單）。內容請在「個案追蹤」K 討論結論、O–Q 會後待辦修改；狀態改「已開刀」就會消失。')
    .setFontColor('#888888');
  const head = ['狀態', '手術日', '剩幾天', '姓名', '病歷號', '預定手術', '討論結論', '注意事項 / 會後待辦', '負責 / 期限'];
  sh.getRange(3, 1, 1, head.length).setValues([head]);
  csHeader_(sh.getRange(3, 1, 1, head.length));
  sh.getRange('H3').setBackground('#F4CCCC');
  sh.getRange('A4').setFormula(csTodoFormula_());
  sh.getRange('B4:B').setNumberFormat('m/d（ddd）');
  sh.getRange('C4:C').setNumberFormat('0').setHorizontalAlignment('center');
  sh.getRange('A4:I').setVerticalAlignment('top');
  sh.getRange('F4:I').setWrap(true);
  sh.getRange('H4:H').setFontWeight('bold');
  const rules = csStatusRules_(sh, 'A4:A200');
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND($C4<>"",$C4>=0,$C4<=7)')
    .setBackground('#F4CCCC').setBold(true).setRanges([sh.getRange('C4:C200')]).build());
  sh.setConditionalFormatRules(rules);
  sh.setFrozenRows(3);
  csWidths_(sh, [100, 90, 55, 75, 85, 200, 280, 300, 90]);
}

function csDone_(sh) {
  csReset_(sh);
  const T = q_(CS.TRK);
  sh.getRange('A1:B5').setValues([
    ['已討論個案追蹤（自動從「個案追蹤」產生：討論日已過的個案，由新到舊）', ''],
    ['已討論', `=COUNTIFS(${T}!F2:F,"<="&TODAY())`],
    ['已開刀', `=COUNTIFS(${T}!F2:F,"<="&TODAY(),${T}!A2:A,"已開刀")`],
    ['已討論-待手術', `=COUNTIFS(${T}!F2:F,"<="&TODAY(),${T}!A2:A,"已討論-待手術")`],
    ['待確認', `=COUNTIFS(${T}!F2:F,"<="&TODAY(),${T}!A2:A,"待確認")`],
  ]);
  sh.getRange('A1:A5').setFontWeight('bold');
  const head = ['討論會日期', '姓名', '病歷號', '主治VS', '狀態', '手術日期', '預定手術', '診斷 / 影像分期', '討論結論', '術後結果 / 病理', '報告醫師'];
  sh.getRange(7, 1, 1, head.length).setValues([head]);
  csHeader_(sh.getRange(7, 1, 1, head.length));
  sh.getRange('A8').setFormula(`=IFERROR(SORT(FILTER(CHOOSECOLS(${T}!A2:N,6,2,3,9,1,4,5,10,11,12,13),${T}!F2:F<>"",${T}!F2:F<=TODAY()),1,FALSE),"（尚無）")`);
  sh.getRange('A8:A').setNumberFormat('yyyy/mm/dd');
  sh.getRange('F8:F').setNumberFormat('yyyy/mm/dd');
  sh.getRange('G8:J').setWrap(true).setVerticalAlignment('top');
  sh.setConditionalFormatRules([csFinishedRule_(sh, 'A8:D500', 'E', 8), csFinishedRule_(sh, 'F8:K500', 'E', 8),
    ...csStatusRules_(sh, 'E8:E500')]);
  sh.setFrozenRows(7); sh.setFrozenColumns(2);
  csWidths_(sh, [130, 100, 100, 100, 100, 100, 200, 240, 340, 220, 100]);
}

function csWeek_(sh) {
  csReset_(sh);
  const T = q_(CS.TRK);
  // 後台頁（調度板讀 B1 與第 6 列起的人數），日期跟著「下次討論會」B1
  sh.getRange('A1:A2').setValues([['討論會日期（跟著「下次討論會」B1）'], ['下一次週四']]);
  sh.getRange('B1').setFormula(`=${q_(CS.NEXT)}!B1`);
  sh.getRange('B2').setFormula('=TODAY()+MOD(5-WEEKDAY(TODAY()),7)');
  sh.getRange('B1:B2').setNumberFormat('yyyy/mm/dd');
  sh.getRange('B1').setBackground('#FFF2CC');
  sh.getRange('A4').setValue('該週已排入名單');
  sh.getRange('A1:A4').setFontWeight('bold');
  const head = ['姓名', '病歷號', '主治VS', '手術日期', '預定手術', '診斷 / 影像分期', '報告醫師'];
  sh.getRange(5, 1, 1, head.length).setValues([head]);
  csHeader_(sh.getRange(5, 1, 1, head.length));
  // Code.gs 的 cancerStatus() 從第 6 列開始數人數，不要移動這張表
  sh.getRange('A6').setFormula(`=IFERROR(FILTER(CHOOSECOLS(${T}!A2:N,2,3,9,4,5,10,13),${T}!F2:F=B1),"（此日無排入個案）")`);
  sh.getRange('D6:D').setNumberFormat('yyyy/mm/dd');
  sh.getRange('F6:F').setWrap(true).setVerticalAlignment('top');
  sh.setFrozenRows(5);
  csWidths_(sh, [150, 110, 110, 110, 240, 300, 90]);
}

function csPend_(sh) {
  csReset_(sh);
  const T = q_(CS.TRK);
  sh.getRange('A1').setValue('待討論個案（依手術日排序；在「個案追蹤」把狀態改為已排入並填討論會日期，就會從這裡消失）').setFontWeight('bold');
  const head = ['姓名', '病歷號', '主治VS', '手術日期', '預定手術', '最晚討論日', '提醒'];
  sh.getRange(2, 1, 1, head.length).setValues([head]);
  csHeader_(sh.getRange(2, 1, 1, head.length));
  sh.getRange('A3').setFormula(`=IFERROR(SORT(FILTER(CHOOSECOLS(${T}!A2:N,2,3,9,4,5,7,8),${T}!A2:A="待討論"),4,TRUE),"（目前沒有待討論個案）")`);
  sh.getRange('D3:D').setNumberFormat('yyyy/mm/dd');
  sh.getRange('F3:F').setNumberFormat('yyyy/mm/dd');
  sh.getRange('G3:G').setFontWeight('bold').setFontColor(RED);
  sh.setFrozenRows(2);
  csWidths_(sh, [110, 110, 110, 110, 260, 130, 130]);
}

function csCand_(sh, surgeon, keywords) {
  csReset_(sh);
  const P = q_(CS.PARSE);
  sh.getRange('A1:A5').setValues([
    ['刀表候選：主刀的重大癌症手術（貼新刀表到「刀表匯入」後自動更新）'], ['主刀'], ['重大手術關鍵字'], ['尚未進追蹤表'],
    ['「新個案」會自動加入「個案追蹤」（狀態＝待討論）；說明!B3 取消勾選就改回手動複製 B–E'],
  ]);
  sh.getRange('B2').setValue(surgeon || '黃冠華');
  sh.getRange('B3').setNumberFormat('@').setValue(keywords || CS_KEYWORDS);
  // Code.gs 的 cancerStatus() 讀 B4，不要移動
  sh.getRange('B4').setFormula('=COUNTIF(A8:A,"新個案")').setFontWeight('bold').setFontColor(RED);
  sh.getRange('A1:A4').setFontWeight('bold');
  sh.getRange('B2:B3').setBackground('#FFF2CC');
  const head = ['比對', '姓名', '病歷號', '手術日期', '預定手術', '癌', '來源'];
  sh.getRange(7, 1, 1, head.length).setValues([head]);
  csHeader_(sh.getRange(7, 1, 1, head.length));
  const e = CS_PARSE_ROWS + 2;
  const cols = ['M', 'E', 'D', 'F', 'N', 'I', 'K'].map(c => `${P}!${c}2:${c}${e}`).join(',');
  sh.getRange('A8').setFormula(`=IFERROR(SORT(UNIQUE(FILTER(HSTACK(${cols}),${P}!L2:L${e}=TRUE)),4,TRUE),"（刀表中沒有符合的手術）")`);
  sh.getRange('D8:D').setNumberFormat('yyyy/mm/dd');
  sh.setFrozenRows(7);
  csWidths_(sh, [140, 100, 100, 100, 400, 60, 60]);
}

function csRaw_(sh) {
  if (sh.getMaxRows() < CS_RAW_ROWS) sh.insertRowsAfter(sh.getMaxRows(), CS_RAW_ROWS - sh.getMaxRows());
  sh.getRange('A1').setValue('刀表原始資料：把HIS「手術工作表」文字整份貼在A2（先清掉舊的），「刀表候選」會自動更新')
    .setFontWeight('bold').setBackground('#FFF2CC');
  sh.setFrozenRows(1);
  sh.setColumnWidth(1, 360);
}

function csParse_(sh) {
  csReset_(sh);
  const head = ['起始列', '日期列', '位置列', '病歷號', '姓名', '手術日', '時', '主刀', '癌', '手術名稱', '來源', '符合條件', '比對', '預定手術'];
  if (sh.getMaxRows() < CS_PARSE_ROWS + 10) sh.insertRowsAfter(sh.getMaxRows(), CS_PARSE_ROWS + 10 - sh.getMaxRows());
  sh.getRange(1, 1, 1, head.length).setValues([head]);
  csHeader_(sh.getRange(1, 1, 1, head.length));
  const f = csParseRow2_();
  sh.getRange(2, 1, 1, f.length).setFormulas([f]);
  sh.getRange(2, 2, 1, f.length - 1).copyTo(sh.getRange(3, 2, CS_PARSE_ROWS, f.length - 1));
  sh.getRange('F2:F').setNumberFormat('yyyy/mm/dd');
  sh.setFrozenRows(1);
}

function csSettings_(ss) {
  const sh = ss.getSheetByName(CS.HELP);
  const v = sh ? sh.getRange('B1:B3').getValues().map(r => r[0]) : [];
  return {
    docUrl: /docs\.google\.com\/document/.test(String(v[0] || '')) ? String(v[0]) : '',
    docTab: v[0] !== undefined && /docs\.google/.test(String(v[0] || '')) && v[1] ? String(v[1]) : CS_DOC_TAB,
    autoAdd: v[2] === false ? false : true,
  };
}

function csHelp_(sh, st) {
  csReset_(sh);
  st = st || { docUrl: '', docTab: CS_DOC_TAB, autoAdd: true };
  sh.getRange('A1:B3').setValues([
    ['交班單 Google 文件網址', st.docUrl],
    ['交班單裡的分頁名稱', st.docTab],
    ['刀表新個案自動加入追蹤', st.autoAdd],
  ]);
  sh.getRange('B3').insertCheckboxes().setValue(st.autoAdd);
  sh.getRange('A1:A3').setFontWeight('bold');
  sh.getRange('B1:B3').setBackground('#FFF2CC');
  sh.getRange('C1').setValue('「注意事項」會自動寫進這份文件的這個分頁（整個分頁會被覆蓋，請用專用分頁）').setFontColor('#888888');
  const rows = [
    ['項目', '內容'],
    ['平常只看三頁', '① 下次討論會：上面是已排入這次的人，下面是還沒討論、可以挑的個案（⚠ 是最晚這次就要討論）。\n② 注意事項：已討論、還沒開刀的人，討論結論與會後待辦，會自動同步到交班單。\n③ 個案追蹤：唯一要手動填的地方。其他分頁都是自動產生，已隱藏（左下「所有工作表」可打開）。'],
    ['一個個案的流程', '刀表匯入後自動加入（狀態＝待討論）→ 排進討論會：填 F 討論會日期，狀態改「已排入」→ 討論完：填 K 討論結論、O 會後待辦（P 負責人、Q 期限），狀態改「已討論-待手術」→ 開完刀：狀態改「已開刀」、填 L 病理。不用討論的改「不需手術」。'],
    ['最晚討論日', '手術日前一個週四（自動計算）。'],
    ['匯入新刀表', '用調度板匯入刀表時會自動送到這裡；也可以把 HIS 手術工作表文字貼在隱藏的「刀表匯入」A2。符合主刀與關鍵字（隱藏的「刀表候選」B2、B3）的重大手術會自動加入個案追蹤。'],
    ['同步到交班單', '編輯「個案追蹤」後幾秒內、以及每小時，會把「注意事項」寫進上面設定的交班單分頁。交班單那一頁不要手動改，會被覆蓋。'],
    ['公式從哪裡來', '所有公式與格式由 GitHub repo ttimcom-hue/Uro-OR 的 cancer/CancerSheet.gs 產生，會自動更新。'],
  ];
  sh.getRange(5, 1, rows.length, 2).setValues(rows);
  csHeader_(sh.getRange(5, 1, 1, 2));
  sh.getRange(6, 1, rows.length - 1, 1).setFontWeight('bold');
  sh.getRange(6, 2, rows.length - 1, 1).setWrap(true).setVerticalAlignment('top');
  sh.getRange(6, 1, rows.length - 1, 1).setVerticalAlignment('top');
  csWidths_(sh, [180, 640]);
}

// ---------- 自動化：新個案自動加入、同步交班單 ----------
function csEnsureTrigger_() {
  const has = ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'cancerOnEdit');
  if (!has) ScriptApp.newTrigger('cancerOnEdit').forSpreadsheet(CANCER_ID).onEdit().create();
}

// 試算表被編輯時（installable trigger）：改了個案追蹤或貼了刀表才做事
function cancerOnEdit(e) {
  const name = e && e.range ? e.range.getSheet().getName() : '';
  if (name !== CS.TRK && name !== CS.RAW) return;
  cancerSync();
}

function cancerSync() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return '忙碌中，略過';
  try {
    const ss = SpreadsheetApp.openById(CANCER_ID);
    const st = csSettings_(ss);
    const added = st.autoAdd ? csAutoAdd_(ss) : 0;
    let doc;
    try { doc = csSyncDoc_(ss, st); } catch (e) { doc = '同步交班單失敗：' + (e && e.message || e); }
    const msg = (added ? '自動加入 ' + added + ' 位新個案；' : '') + doc;
    console.log(msg);
    return msg;
  } finally { lock.releaseLock(); }
}

function csAutoAdd_(ss) {
  SpreadsheetApp.flush();
  const cand = ss.getSheetByName(CS.CAND), trk = ss.getSheetByName(CS.TRK);
  const n = cand.getLastRow();
  if (n < 8) return 0;
  const have = new Set(trk.getRange('C2:C').getValues().map(r => String(r[0]).trim().replace(/^'/, '')).filter(String));
  const add = new Map();   // 病歷號 → 一列；同一人多筆手術合併
  cand.getRange(8, 1, n - 7, 5).getValues().forEach(([flag, name, mrn, date, proc]) => {
    mrn = String(mrn).trim();
    if (flag !== '新個案' || !mrn || have.has(mrn)) return;
    if (add.has(mrn)) { const r = add.get(mrn); if (proc && r[4].indexOf(proc) < 0) r[4] += ' + ' + proc; return; }
    add.set(mrn, ['待討論', name, "'" + mrn, date, String(proc || '')]);
  });
  if (!add.size) return 0;
  // 找個案追蹤最後一個有姓名或病歷號的列（G、H 欄預先填了公式，不能用 getLastRow）
  const bc = trk.getRange('B2:C').getValues();
  let last = 1;
  bc.forEach((r, i) => { if (String(r[0]).trim() || String(r[1]).trim()) last = i + 2; });
  const rows = [...add.values()], r0 = last + 1;
  trk.getRange(r0, 1, rows.length, 5).setValues(rows);
  const note = '刀表自動加入 ' + Utilities.formatDate(new Date(), 'Asia/Taipei', 'M/d');
  trk.getRange(r0, 14, rows.length, 1).setValues(rows.map(() => [note]));
  for (let r = r0; r < r0 + rows.length; r++) {
    if (!trk.getRange(r, 7).getFormula()) trk.getRange(r, 7, 1, 2).setFormulas([[csTrackerG_(r), csTrackerH_(r)]]);
  }
  return rows.length;
}

function csDateStr_(d) {
  if (!(d instanceof Date)) return String(d || '');
  return Utilities.formatDate(d, 'Asia/Taipei', 'M/d') + '（' + '日一二三四五六'[d.getDay()] + '）';
}

function csFindTab_(tabs, title) {
  for (const t of tabs) {
    if (t.getTitle().trim() === title) return t;
    const c = csFindTab_(t.getChildTabs(), title);
    if (c) return c;
  }
  return null;
}

function csSyncDoc_(ss, st) {
  if (!st.docUrl) return '未設定交班單網址（說明!B1），略過同步';
  const trk = ss.getSheetByName(CS.TRK);
  const n = trk.getLastRow();
  const v = n > 1 ? trk.getRange(2, 1, n - 1, 17).getValues() : [];
  const list = v.filter(r => CS_OPEN.includes(String(r[0])) && (String(r[1]).trim() || String(r[2]).trim()))
    .sort((a, b) => (a[3] instanceof Date ? a[3].getTime() : 9e15) - (b[3] instanceof Date ? b[3].getTime() : 9e15));
  const head = ['手術日', '病人', '預定手術', '討論結論', '注意事項 / 會後待辦', '負責 / 期限'];
  const rows = list.map(r => [
    [csDateStr_(r[3]), r[0] === '已討論-待手術' ? '' : '【' + r[0] + '】'].filter(String).join('\n'),
    String(r[1]).trim() + '\n' + String(r[2]).trim(),
    String(r[4]).trim(),
    String(r[10]).trim(),
    String(r[14]).trim(),
    [String(r[15]).trim(), r[16] instanceof Date ? csDateStr_(r[16]) : String(r[16]).trim()].filter(String).join('\n'),
  ]);
  const sig = JSON.stringify([st.docUrl, st.docTab, rows]);
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('csDocSig') === sig) return '交班單已是最新';

  const id = (st.docUrl.match(/\/document\/d\/([\w-]+)/) || [])[1];
  if (!id) return '交班單網址格式不對（說明!B1）';
  const tab = csFindTab_(DocumentApp.openById(id).getTabs(), st.docTab);
  if (!tab) throw new Error('交班單找不到「' + st.docTab + '」分頁：請在交班單新增一個這個名稱的分頁');
  const body = tab.asDocumentTab().getBody();
  body.clear();
  body.appendParagraph('癌症討論會：已討論、還沒開刀的注意事項').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph('自動同步自「癌症討論會追蹤表」，更新於 ' +
    Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy/MM/dd HH:mm') + '。請到追蹤表修改，這一頁的修改會被覆蓋。')
    .editAsText().setFontSize(9).setForegroundColor('#888888');
  if (rows.length) {
    const tb = body.appendTable([head, ...rows]);
    tb.editAsText().setFontSize(10);
    const h = tb.getRow(0);
    for (let c = 0; c < head.length; c++) h.getCell(c).setBackgroundColor('#EFEFEF');
    h.editAsText().setBold(true);
    [70, 75, 110, 150, 150, 65].forEach((w, c) => tb.setColumnWidth(c, w));
  } else {
    body.appendParagraph('（目前沒有）');
  }
  try { const p0 = body.getChild(0); if (p0.getType() === DocumentApp.ElementType.PARAGRAPH && !p0.asParagraph().getText()) p0.removeFromParent(); } catch (e) { /* 只剩一段時不能刪 */ }
  props.setProperty('csDocSig', sig);
  return '已同步 ' + rows.length + ' 位到交班單';
}

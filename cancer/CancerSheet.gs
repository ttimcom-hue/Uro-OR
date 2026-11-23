// 癌症討論會追蹤表：重建公式分頁（Google Apps Script）
// 放在跟 Code.gs 同一個 Apps Script 專案。改完後在編輯器執行 setupCancerSheet() 一次。
// 只會改「公式／格式」，不會動到使用者輸入的資料：
//   - 個案追蹤：只重寫 G、H 兩欄公式與格式，A–F、I–Q 的資料不動
//   - 會後待辦、已討論個案、週會名單、待討論清單、刀表候選、刀表解析：整頁重寫，
//     但保留 週會名單!B1（討論會日期）、刀表候選!B2:B3（主刀、關鍵字）
//   - 刀表匯入：只寫 A1 標題，A2 以下的原始刀表不動
// CANCER_ID 定義在 Code.gs。

const CS = {
  TRK: '個案追蹤', TODO: '會後待辦', DONE: '已討論個案', WEEK: '週會名單',
  PEND: '待討論清單', CAND: '刀表候選', RAW: '刀表匯入', PARSE: '刀表解析', HELP: '說明',
};
const CS_ORDER = [CS.TRK, CS.TODO, CS.DONE, CS.WEEK, CS.PEND, CS.CAND, CS.RAW, CS.PARSE, CS.HELP];
const CS_STATUS = ['待討論', '已排入', '已討論-待手術', '已開刀', '待確認', '暫緩/改期', '不需手術'];
const CS_STATUS_COLOR = {
  '待討論': '#F4CCCC', '已排入': '#CFE2F3', '已討論-待手術': '#FFF2CC', '已開刀': '#D9EAD3',
  '待確認': '#FCE5CD', '暫緩/改期': '#E0E0E0', '不需手術': '#D9D2E9',
};
// 「會後待辦」要列出的狀態（已討論但還沒開刀）
const CS_OPEN = ['已討論-待手術', '暫緩/改期', '待確認'];
const CS_TRK_HEAD = ['狀態', '姓名', '病歷號', '手術日期', '預定手術', '討論會日期', '最晚討論日', '提醒',
  '主治VS', '診斷 / 影像分期', '討論結論', '術後結果 / 病理', '報告醫師', '來源 / 備註',
  '會後待辦', '負責人', '待辦期限'];
const CS_KEYWORDS = '(?i)davinci|prostatectomy|nephr\\w*ectomy|nephrouret|\\bRNU\\b|cystectomy|RARC|adrenal|thrombect|retroperitoneal';
const CS_PARSE_ROWS = 400;   // 一次刀表最多解析幾筆手術
const CS_RAW_ROWS = 6000;    // 刀表匯入最多幾行
const HDR_BG = '#EFEFEF', RED = '#C00000';

function q_(name) { return "'" + name + "'"; }

// ---------- 公式 ----------
function csTrackerG_(r) { return `=IF(D${r}="","",D${r}-(MOD(WEEKDAY(D${r})-6,7)+1))`; }
function csTrackerH_(r) {
  return `=IF(D${r}="","",IF(A${r}="待討論",IF(G${r}<TODAY(),"⚠ 已過最晚討論日",IF(G${r}-TODAY()<=14,"⚠ 需排入討論會","")),` +
    `IF(AND(OR(A${r}="已討論-待手術",A${r}="已排入"),D${r}<TODAY()),"手術日已過，請更新","")))`;
}
function csTodoFormula_() {
  const T = q_(CS.TRK);
  const cond = CS_OPEN.map(s => `(st="${s}")`).join('+');
  // 欄位：姓名,病歷號,VS,狀態,手術日,距手術(天),討論會日期,會後待辦,負責人,期限,討論結論
  return `=IFERROR(LET(t,${T}!A2:Q,st,INDEX(t,,1),f,FILTER(t,${cond}),s,SORT(f,4,TRUE),d,CHOOSECOLS(s,4),` +
    `HSTACK(CHOOSECOLS(s,2,3,9,1,4),ARRAYFORMULA(IF(d="","",d-TODAY())),CHOOSECOLS(s,6,15,16,17,11))),"（目前沒有）")`;
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
  const keep = {
    weekDate: csGet_(ss, CS.WEEK, 'B1'),
    surgeon: csGet_(ss, CS.CAND, 'B2'),
    keywords: csGet_(ss, CS.CAND, 'B3'),
  };
  CS_ORDER.forEach((name, i) => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name, i);
    ss.setActiveSheet(sh); ss.moveActiveSheet(i + 1);
  });
  csTracker_(ss.getSheetByName(CS.TRK));
  csTodo_(ss.getSheetByName(CS.TODO));
  csDone_(ss.getSheetByName(CS.DONE));
  csWeek_(ss.getSheetByName(CS.WEEK), keep.weekDate);
  csPend_(ss.getSheetByName(CS.PEND));
  csCand_(ss.getSheetByName(CS.CAND), keep.surgeon, keep.keywords);
  csRaw_(ss.getSheetByName(CS.RAW));
  csParse_(ss.getSheetByName(CS.PARSE));
  csHelp_(ss.getSheetByName(CS.HELP));
  SpreadsheetApp.flush();
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

function csTracker_(sh) {
  const last = Math.max(sh.getLastRow(), 1);
  sh.getRange(1, 1, 1, CS_TRK_HEAD.length).setValues([CS_TRK_HEAD]);
  csHeader_(sh.getRange(1, 1, 1, 14));
  sh.getRange(1, 15, 1, 3).setFontWeight('bold').setBackground('#F4CCCC');
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
  sh.setConditionalFormatRules(csStatusRules_(sh, 'A2:A500'));
  sh.setFrozenRows(1); sh.setFrozenColumns(2);
  csWidths_(sh, [110, 90, 90, 90, 230, 100, 100, 130, 70, 260, 360, 220, 70, 240, 340, 90, 90]);
}

function csTodo_(sh) {
  csReset_(sh);
  sh.setTabColor('#E06666');
  sh.getRange('A1').setValue('癌症討論會：已討論、尚未手術病人的會後待辦').setFontWeight('bold').setFontSize(12);
  sh.getRange('A2').setValue('自動產生，請勿在此頁編輯。待辦、負責人、期限請填在「個案追蹤」O–Q欄；狀態改為「已開刀」就會從這裡消失。')
    .setFontStyle('italic').setFontColor('#666666');
  const head = ['姓名', '病歷號', '主治VS', '狀態', '手術日期', '距手術(天)', '討論會日期', '會後待辦', '負責人', '期限', '討論結論'];
  sh.getRange(3, 1, 1, head.length).setValues([head]);
  csHeader_(sh.getRange(3, 1, 1, head.length));
  sh.getRange('H3:J3').setBackground('#F4CCCC');
  sh.getRange('A4').setFormula(csTodoFormula_());
  sh.getRange('E4:E').setNumberFormat('yyyy/mm/dd');
  sh.getRange('G4:G').setNumberFormat('yyyy/mm/dd');
  sh.getRange('J4:J').setNumberFormat('yyyy/mm/dd');
  sh.getRange('F4:F').setNumberFormat('0').setHorizontalAlignment('center');
  sh.getRange('A4:K').setVerticalAlignment('top');
  sh.getRange('H4:H').setWrap(true).setFontWeight('bold');
  sh.getRange('K4:K').setWrap(true);
  const rules = csStatusRules_(sh, 'D4:D200');
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND($F4<>"",$F4>=0,$F4<=7)')
    .setBackground('#F4CCCC').setBold(true).setRanges([sh.getRange('F4:F200')]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND($J4<>"",$J4<TODAY())')
    .setBackground('#F4CCCC').setBold(true).setRanges([sh.getRange('J4:J200')]).build());
  sh.setConditionalFormatRules(rules);
  sh.setFrozenRows(3); sh.setFrozenColumns(1);
  csWidths_(sh, [80, 80, 80, 105, 90, 75, 90, 340, 85, 85, 340]);
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
  sh.setConditionalFormatRules(csStatusRules_(sh, 'E8:E500'));
  sh.setFrozenRows(7); sh.setFrozenColumns(2);
  csWidths_(sh, [130, 100, 100, 100, 100, 100, 200, 240, 340, 220, 100]);
}

function csWeek_(sh, weekDate) {
  csReset_(sh);
  const T = q_(CS.TRK);
  sh.getRange('A1:A2').setValues([['討論會日期（改這格）'], ['下一次週四']]);
  if (weekDate) sh.getRange('B1').setValue(weekDate);
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
    ['做法：把「新個案」那列的B–E複製，貼到「個案追蹤」最後一列的B–E，狀態選「待討論」'],
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

function csHelp_(sh) {
  csReset_(sh);
  const rows = [
    ['項目', '內容'],
    ['選案原則', '黃冠華部長的主要癌症/腎上腺手術：Davinci RARP(含SP)、部分/根除性腎切、RNU、RARC、adrenalectomy、IVC thrombectomy；另加需影像討論的切片個案。Cystoscopy、URS、TURBT、HoLEP、結石、hernia不列入。'],
    ['狀態', '待討論 → 已排入(填討論會日期) → 已討論-待手術(填結論與會後待辦) → 已開刀(填術後結果/病理)。其他：待確認、暫緩/改期、不需手術。'],
    ['會後待辦（交班用）', '每位病人一列，自動列出已討論-待手術、暫緩/改期、待確認的病人，依手術日排序；改為已開刀就消失。待辦、負責人、期限填在「個案追蹤」O–Q欄。這頁的A3:K30是連結到交班文件的範圍。'],
    ['最晚討論日', '手術日前一個週四(自動計算)。週四上午開的刀可當天早上07:30討論，公式仍保守算前一週。'],
    ['提醒欄', '待討論且最晚討論日在14天內 → ⚠ 需排入討論會；已過 → ⚠ 已過最晚討論日；已排入/待手術但手術日已過 → 請更新狀態。'],
    ['新增個案', '在「個案追蹤」最後一列下方輸入；G、H欄已預先填好公式。'],
    ['匯入新刀表', '1) 到「刀表匯入」清掉A2以下舊資料 2) 把HIS手術工作表文字整份貼在A2(一個欄位一列、空白列保留)，或用調度板的匯入功能 3) 看「刀表候選」：標「新個案」的是還沒進追蹤表的重大手術，把B–E複製貼到「個案追蹤」。主刀與關鍵字可在「刀表候選」B2、B3修改。「刀表解析」是後台計算，不用動(一次最多400筆手術)。'],
    ['已討論個案', '自動列出討論日已過的個案(含結論、手術、術後病理)。內容請在「個案追蹤」修改。'],
    ['同步到交班文件', '在本試算表選取要同步的範圍(「會後待辦」A3:K30)→複製→到交班Google文件貼上→選「連結至試算表」。之後表格右上角出現「更新」就按一下。'],
    ['公式從哪裡來', '所有公式與格式由 GitHub repo ttimcom-hue/Uro-OR 的 cancer/CancerSheet.gs 產生；改結構請改那份程式後執行 setupCancerSheet()。'],
  ];
  sh.getRange(1, 1, rows.length, 2).setValues(rows);
  csHeader_(sh.getRange(1, 1, 1, 2));
  sh.getRange(2, 1, rows.length - 1, 1).setFontWeight('bold');
  sh.getRange(2, 2, rows.length - 1, 1).setWrap(true).setVerticalAlignment('top');
  csWidths_(sh, [180, 560]);
}

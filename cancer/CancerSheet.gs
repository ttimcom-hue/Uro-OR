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
  NEXT: '下次討論會', TODO: '注意事項', TRK: '個案追蹤', DV: '達文西術前', STAT: '統計', DONE: '已討論個案', WEEK: '週會名單',
  PEND: '待討論清單', CAND: '刀表候選', RAW: '刀表匯入', PARSE: '刀表解析', HELP: '說明',
};
const CS_RENAMED = { '會後待辦': CS.TODO };   // 舊分頁名 → 新分頁名
// 前四頁是平常會用的，其餘是後台（隱藏，需要時從左下「所有工作表」打開）
const CS_ORDER = [CS.NEXT, CS.TODO, CS.DV, CS.TRK, CS.STAT, CS.HELP, CS.DONE, CS.WEEK, CS.PEND, CS.CAND, CS.RAW, CS.PARSE];
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
  [CS.NEXT]: '#3D85C6', [CS.TODO]: '#E06666', [CS.TRK]: '#6AA84F', [CS.DV]: '#F6B26B', [CS.STAT]: '#8E7CC3', [CS.HELP]: '#CCCCCC', [CS.DONE]: '#CCCCCC',
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
    statFrom: csGet_(ss, CS.STAT, 'B2'),
    statTo: csGet_(ss, CS.STAT, 'D2'),
  };
  CS_ORDER.forEach((name, i) => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name, i);
    ss.setActiveSheet(sh); ss.moveActiveSheet(i + 1);
    sh.setTabColor(CS_TAB_COLOR[name]);
  });
  csNext_(ss.getSheetByName(CS.NEXT), keep.nextDate);
  csStat_(ss.getSheetByName(CS.STAT), keep.statFrom, keep.statTo);
  csDv_(ss.getSheetByName(CS.DV));
  csSyncDv_(ss);
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

// ---------- 達文西術前 ----------
// 一人一列（以病歷號對應）。A–H 由程式從個案追蹤更新；I–R 人工填寫（第一次列入時從診斷欄帶入讀得到的值）；L、S 是公式
const CS_DV_HEAD = ['手術日', '時段', '姓名', '病歷號', '術式', '主治', '狀態', '討論會',
  'PSA（最近／日期）', '切片 GS／陽性針數', '切片日期', '切片→手術(週)', 'MRI（日期／PI-RADS／cT／EPE・SV・LN）',
  '攝護腺體積(cc)', 'ADT（藥物／開始日）', '過去手術史（TURP、疝氣補片、腹部手術）', 'IIEF-5', '術前注意／備註', 'RARP 完成度'];
const CS_DV_N = CS_DV_HEAD.length;   // 19 欄（A–S）
const CS_DV_ROBOT = /davinci|robot|RARP|RAPN|RALN|RANU|RaNxUx|RARC|RALA|SP-RA/i;
const CS_DV_RARP = '(?i)RARP|prostatectomy';
function csDvFormulas_(r) {
  return [`=IF(OR(A${r}="",K${r}=""),"",ROUND((A${r}-K${r})/7,1))`,
    `=IF(REGEXMATCH(E${r}&"","${CS_DV_RARP}"),COUNTA(I${r}:K${r},M${r}:Q${r})&"/8","")`];
}
function csDv_(sh) {
  sh.getRange(1, 1, 1, CS_DV_N).setValues([CS_DV_HEAD]).setFontWeight('bold').setWrap(true).setVerticalAlignment('middle');
  sh.getRange(1, 1, 1, 8).setBackground(HDR_BG);
  sh.getRange(1, 9, 1, 10).setBackground('#FCE5CD');
  sh.getRange(1, 19).setBackground(HDR_BG);
  sh.getRange('A2:A').setNumberFormat('m/d（ddd）');
  sh.getRange('B2:B').setNumberFormat('@');   // 時段 08:30 存成文字，不要變成時間
  sh.getRange('K2:K').setNumberFormat('yyyy/mm/dd');
  sh.getRange('A2:S').setVerticalAlignment('top');
  sh.getRange('E2:E').setWrap(true); sh.getRange('I2:R').setWrap(true);
  sh.getRange('A2:H500').setBackground('#F3F3F3');
  sh.getRange('L2:L500').setBackground('#F3F3F3'); sh.getRange('S2:S500').setBackground('#F3F3F3');
  const rule = (f, rng, fmt) => fmt(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(f)).setRanges([sh.getRange(rng)]).build();
  sh.setConditionalFormatRules([
    rule('=OR($G2="已開刀",$G2="不需手術")', 'A2:S500', b => b.setFontColor('#A0A0A0')),
    rule('=AND(ISNUMBER($L2),$L2<6)', 'L2:L500', b => b.setFontColor(RED).setBold(true)),
    rule('=LEFT($H2,1)="⚠"', 'H2:H500', b => b.setFontColor(RED).setBold(true)),
    rule('=AND($S2<>"",$S2<>"8/8")', 'S2:S500', b => b.setFontColor(RED)),
  ]);
  sh.setFrozenRows(1); sh.setFrozenColumns(4);
  csWidths_(sh, [85, 55, 70, 80, 170, 60, 95, 80, 120, 120, 90, 70, 260, 75, 150, 200, 60, 220, 70]);
}
function csDvPrefill_(j) {
  j = String(j || '');
  const m = re => (j.match(re) || [])[0] || '';
  const psa = j.match(/PSA[\s:：]*([\d.]+)(\s*→\s*[\d.]+)?/i);
  const mri = [m(/PI-?RADS\s*\d/i), m(/y?c?T[1-4x][a-c]?\s*N[0-9x]\b/i)].filter(String).join('，');
  const hx = [...new Set((j.match(/TURP|TUEP|TUIP|HIFU|hernia|TURBT/gi) || []).map(x => x.toUpperCase()))].join('、');
  return {
    psa: psa ? psa[1] + (psa[2] || '').replace(/\s/g, '') : '',
    gs: m(/GS\s*\d\s*\+\s*\d/i).replace(/\s/g, ' '),
    mri, vol: (j.match(/(\d+(?:\.\d+)?)\s*(?:cc|ml)\b/i) || [])[1] || '',
    adt: /ADT/i.test(j) ? '有（' + m(/neo-?ADT|s\/p ADT|ADT/i) + '，藥物／日期待補）' : '',
    hx: hx ? 'Hx ' + hx : '',
  };
}
// 從個案追蹤更新「達文西術前」；回傳新增人數。只在自動欄位有變時才寫回
function csSyncDv_(ss) {
  const sh = ss.getSheetByName(CS.DV);
  if (!sh) return 0;
  const trk = ss.getSheetByName(CS.TRK), tn = trk.getLastRow();
  const tv = tn > 1 ? trk.getRange(2, 1, tn - 1, 17).getValues() : [];
  const today = csToday_();
  const n0 = sh.getLastRow();
  const old = n0 > 1 ? sh.getRange(2, 1, n0 - 1, CS_DV_N).getValues().filter(r => String(r[3]).trim()) : [];
  const byMrn = new Map(old.map(r => [String(r[3]).trim().replace(/^'/, ''), r]));
  let added = 0;
  tv.forEach(t => {
    const mrn = String(t[2]).trim().replace(/^'/, ''), op = String(t[4] || ''), st = String(t[0] || '');
    if (!mrn || !CS_DV_ROBOT.test(op)) return;
    const d = t[3] instanceof Date ? t[3] : '';
    const mt = t[5] instanceof Date ? t[5] : '';
    const slot = (op.match(/\b(0?8:?30|12:?00|TF\s*\d?)\b/i) || [])[1] || '';
    const disc = mt ? (mt <= today ? Utilities.formatDate(mt, 'Asia/Taipei', 'M/d') + ' 已討論' : Utilities.formatDate(mt, 'Asia/Taipei', 'M/d') + ' 已排入') : '⚠ 未上討論會';
    const auto = [d, slot.replace(/\s/g, '').toUpperCase(), t[1], "'" + mrn, op, t[8], st, disc];
    let r = byMrn.get(mrn);
    if (!r) {
      const open = !['已開刀', '不需手術'].includes(st) && (!d || d >= today);
      if (!open) return;
      const p = csDvPrefill_(t[9]);
      r = new Array(CS_DV_N).fill('');
      r[8] = p.psa; r[9] = p.gs; r[12] = p.mri; r[13] = p.vol; r[14] = p.adt; r[15] = p.hx;
      r[17] = '自動帶入（請核對）';
      byMrn.set(mrn, r); added++;
    }
    auto.forEach((v, i) => { r[i] = v; });
  });
  const rows = [...byMrn.values()];
  const done = r => ['已開刀', '不需手術'].includes(String(r[6]));
  const tm = r => r[0] instanceof Date ? r[0].getTime() : 9e15;
  rows.sort((a, b) => done(a) - done(b) || (done(a) ? tm(b) - tm(a) : tm(a) - tm(b)));
  const key = r => JSON.stringify(r.slice(0, 18).map((v, i) => i === 11 ? '' : v instanceof Date ? v.getTime() : String(v).replace(/^'/, '')));
  if (!added && rows.length === old.length && rows.every((r, i) => key(r) === key(old[i]))) return 0;
  if (n0 > 1) sh.getRange(2, 1, n0 - 1, CS_DV_N).clearContent();
  if (!rows.length) return 0;
  sh.getRange(2, 1, rows.length, CS_DV_N).setValues(rows.map(r => r.map((v, i) => i === 11 || i === 18 ? '' : v)));
  const f = rows.map((_, i) => csDvFormulas_(i + 2));
  sh.getRange(2, 12, rows.length, 1).setFormulas(f.map(x => [x[0]]));
  sh.getRange(2, 19, rows.length, 1).setFormulas(f.map(x => [x[1]]));
  return added;
}

// ---------- 統計 ----------
// 癌別：先看「預定手術」(E)，判斷不出來再看「診斷」(J)；依序比對，第一個符合的就是
const CS_CATS = [
  ['腎上腺', '(?i)adrenal|RALA|aldoster|腎上腺'],
  ['上泛尿路', '(?i)UTUC|RNU|RANU|NxUx|NUx|nephroureter|ureter.{0,20}(UC|tumor)|renal pelvi'],
  ['膀胱', '(?i)UBUC|RARC|cystectomy|bladder|TURBT'],
  ['攝護腺', '(?i)RARP|prostat|\\bPC\\b|PSA|fusion'],
  ['腎臟', '(?i)RCC|renal|kidney|nephrect|\\bNx\\b|RAPN|RALN|\\bPN\\b|AML|thrombect'],
  ['其他', '(?i)retroperiton'],
];
const CS_CAT_LIST = ['攝護腺', '腎臟', '上泛尿路', '膀胱', '腎上腺', '其他'];
// 統計頁右側隱藏的明細（AB:AH）：討論日,月份,癌別,狀態,主治,報告醫師,病歷號
function csStatBaseFormula_() {
  const T = q_(CS.TRK);
  const ifs = v => 'ARRAYFORMULA(IFS(' + CS_CATS.map(([n, r]) => `REGEXMATCH(${v}&"","${r}"),"${n}"`).join(',') + ',TRUE,""))';
  return `=IFERROR(LET(t,${T}!A2:Q,f,INDEX(t,,6),s,IF($B$2="",DATE(2000,1,1),$B$2),e,IF($D$2="",TODAY(),$D$2),` +
    `b,FILTER(t,f<>"",f>=s,f<=e,INDEX(t,,2)<>""),pe,INDEX(b,,5),pj,INDEX(b,,10),ce,${ifs('pe')},cj,${ifs('pj')},` +
    `HSTACK(INDEX(b,,6),ARRAYFORMULA(TEXT(INDEX(b,,6),"yyyy-mm")),ARRAYFORMULA(IF(ce<>"",ce,IF(cj<>"",cj,"其他"))),` +
    `INDEX(b,,1),INDEX(b,,9),INDEX(b,,13),INDEX(b,,3))),"")`;
}
function csStatMonthFormula_() {
  const c = CS_CAT_LIST.join(',');
  return `=IFERROR(LET(m,SORT(UNIQUE(FILTER(AC4:AC,AC4:AC<>""))),c,SPLIT("${c}",","),n,COLUMNS(c),` +
    `body,MAKEARRAY(ROWS(m),n+1,LAMBDA(i,j,IF(j>n,COUNTIF(AC4:AC,INDEX(m,i)),COUNTIFS(AC4:AC,INDEX(m,i),AD4:AD,INDEX(c,1,j))))),` +
    `tot,MAKEARRAY(1,n+1,LAMBDA(i,j,IF(j>n,COUNTIF(AC4:AC,"?*"),COUNTIF(AD4:AD,INDEX(c,1,j))))),` +
    `VSTACK(HSTACK("月份",c,"合計"),HSTACK(m,body),HSTACK("合計",tot))),"（無資料）")`;
}
function csStatGroupFormula_(col, label) {
  return `=IFERROR(QUERY(AB4:AH,"select Col${col}, count(Col1) where Col1 is not null and Col${col} <> '' ` +
    `group by Col${col} order by count(Col1) desc label Col${col} '${label}', count(Col1) '人數'",0),"（無資料）")`;
}
function csStatMissedFormula_() {
  const T = q_(CS.TRK);
  return `=IFERROR(LET(t,${T}!A2:Q,s,IF($B$2="",DATE(2000,1,1),$B$2),e,IF($D$2="",TODAY(),$D$2),d,INDEX(t,,4),f,INDEX(t,,6),` +
    `SORT(FILTER(HSTACK(d,INDEX(t,,2),INDEX(t,,3),INDEX(t,,5),INDEX(t,,9)),INDEX(t,,1)="已開刀",d>=s,d<=e,(f="")+(f>d)),1,FALSE)),"（沒有）")`;
}

function csStat_(sh, from, to) {
  csReset_(sh);
  const T = q_(CS.TRK);
  sh.getRange('A1').setValue('癌症討論會統計').setFontWeight('bold').setFontSize(13);
  sh.getRange('A2:D2').setValues([['統計期間（討論會日期）', from || '', '～', to || '']]);
  sh.getRange('B2').setNumberFormat('yyyy/mm/dd').setBackground('#FFF2CC');
  sh.getRange('D2').setNumberFormat('yyyy/mm/dd').setBackground('#FFF2CC');
  sh.getRange('E2').setValue('黃色格空白＝不限起日／到今天為止。統計來源是「個案追蹤」，討論會日期 ≤ 今天才算已討論。').setFontColor('#888888');
  const range = `s,IF($B$2="",DATE(2000,1,1),$B$2),e,IF($D$2="",TODAY(),$D$2)`;
  const st = x => `=COUNTIF($AE$4:$AE,"${x}")`;
  const rows = [
    ['總覽', '人數'],
    ['已討論人數', '=COUNT($AB$4:$AB)'],
    ['討論會場數', '=IFERROR(ROWS(UNIQUE(FILTER($AB$4:$AB,ISNUMBER($AB$4:$AB)))),0)'],
    ['平均每場', '=IFERROR(ROUND(B6/B7,1),0)'],
    ['討論後：已開刀', st('已開刀')],
    ['　　　　待手術', st('已討論-待手術')],
    ['　　　　暫緩/改期', st('暫緩/改期')],
    ['　　　　待確認', st('待確認')],
    ['　　　　不需手術', st('不需手術')],
    ['期間內已開刀（依手術日）', `=LET(${range},COUNTIFS(${T}!A2:A,"已開刀",${T}!D2:D,">="&s,${T}!D2:D,"<="&e))`],
    ['　其中手術前沒上討論會', '=COUNT($V$4:$V)'],
    ['　討論率', '=IFERROR(TEXT(1-B16/B15,"0%"),"-")'],
  ];
  sh.getRange(5, 1, rows.length, 2).setValues(rows.map(r => [r[0], '']));
  sh.getRange(6, 2, rows.length - 1, 1).setFormulas(rows.slice(1).map(r => [r[1]]));
  sh.getRange('B5').setValue('人數');
  csHeader_(sh.getRange('A5:B5'));
  sh.getRange('A6:A16').setFontWeight('bold');

  sh.getRange('D4').setValue('每月討論人數（依癌別）').setFontWeight('bold');
  sh.getRange('D5').setFormula(csStatMonthFormula_());
  sh.getRange('M4').setValue('依主治').setFontWeight('bold');
  sh.getRange('M5').setFormula(csStatGroupFormula_(5, '主治'));
  sh.getRange('P4').setValue('依報告醫師').setFontWeight('bold');
  sh.getRange('P5').setFormula(csStatGroupFormula_(6, '報告醫師'));
  sh.getRange('S4').setValue('每場討論會').setFontWeight('bold');
  sh.getRange('S5').setFormula(`=IFERROR(QUERY(AB4:AH,"select Col1, count(Col1) where Col1 is not null group by Col1 order by Col1 desc label Col1 '討論會日期', count(Col1) '人數'",0),"（無資料）")`);
  sh.getRange('S6:S').setNumberFormat('yyyy/mm/dd');
  sh.getRange('V2').setValue('手術前沒上討論會的已開刀個案（期間內，依手術日新到舊）').setFontWeight('bold');
  const mh = ['手術日', '姓名', '病歷號', '手術', '主治'];
  sh.getRange(3, 22, 1, mh.length).setValues([mh]);
  csHeader_(sh.getRange(3, 22, 1, mh.length));
  sh.getRange('V4').setFormula(csStatMissedFormula_());
  sh.getRange('V4:V').setNumberFormat('yyyy/mm/dd');
  [[5, 4, 8], [5, 13, 2], [5, 16, 2], [5, 19, 2]].forEach(([r, c, n]) => csHeader_(sh.getRange(r, c, 1, n)));

  // 明細（隱藏）
  const bh = ['討論日', '月份', '癌別', '狀態', '主治', '報告醫師', '病歷號'];
  sh.getRange(3, 28, 1, bh.length).setValues([bh]);
  sh.getRange('AB4').setFormula(csStatBaseFormula_());
  sh.getRange('AB4:AB').setNumberFormat('yyyy/mm/dd');
  if (sh.getMaxColumns() >= 34) sh.hideColumns(28, 7);

  sh.setFrozenRows(5);
  csWidths_(sh, [190, 70, 20, 80, 60, 60, 70, 60, 60, 60, 60, 20, 90, 50, 20, 90, 50, 20, 100, 50, 20, 90, 80, 90, 200, 80]);
}

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
  const v = sh ? sh.getRange('B1:B4').getValues().map(r => r[0]) : [];
  const dv = String(v[3] || '').trim();
  return {
    driveId: (dv.match(/folders\/([\w-]{20,})/) || dv.match(/^([\w-]{25,})$/) || [])[1] || '',
    docUrl: /docs\.google\.com\/document/.test(String(v[0] || '')) ? String(v[0]) : '',
    docTab: v[0] !== undefined && /docs\.google/.test(String(v[0] || '')) && v[1] ? String(v[1]) : CS_DOC_TAB,
    autoAdd: v[2] === false ? false : true,
  };
}

function csHelp_(sh, st) {
  csReset_(sh);
  st = st || { docUrl: '', docTab: CS_DOC_TAB, autoAdd: true, driveId: '' };
  sh.getRange('A1:B4').setValues([
    ['交班單 Google 文件網址', st.docUrl],
    ['交班單裡的分頁名稱', st.docTab],
    ['刀表新個案自動加入追蹤', st.autoAdd],
    ['討論會 Drive 資料夾', st.driveId ? 'https://drive.google.com/drive/folders/' + st.driveId : ''],
  ]);
  sh.getRange('B3').insertCheckboxes().setValue(st.autoAdd);
  sh.getRange('A1:A4').setFontWeight('bold');
  sh.getRange('B1:B4').setBackground('#FFF2CC');
  sh.getRange('C4').setValue('每小時讀這個資料夾裡新的「給放射科」名單與「癌症討論會」紀錄單，更新個案追蹤與達文西術前（只補空白欄位）').setFontColor('#888888');
  sh.getRange('C1').setValue('「注意事項」會自動寫進這份文件的這個分頁（整個分頁會被覆蓋，請用專用分頁）').setFontColor('#888888');
  const rows = [
    ['項目', '內容'],
    ['平常只看三頁', '① 下次討論會：上面是已排入這次的人，下面是還沒討論、可以挑的個案（⚠ 是最晚這次就要討論）。\n② 注意事項：已討論、還沒開刀的人，討論結論與會後待辦，會自動同步到交班單。\n③ 個案追蹤：唯一要手動填的地方。其他分頁都是自動產生，已隱藏（左下「所有工作表」可打開）。'],
    ['達文西術前', '個案追蹤裡預定達文西／機器手臂手術、還沒開的病人會自動列進來（依手術日排序，開完的移到下面並變灰）。灰底 A–H 欄自動更新；白底欄位（PSA、切片、MRI、體積、ADT、過去手術、IIEF-5、備註）請手動補，第一次列入時會從診斷欄自動帶入能讀到的值。完成度只算 RARP。切片到手術不足 6 週會標紅。'],
    ['從 Drive 自動更新', '每小時（或入口按「從 Drive 更新」）讀 B4 資料夾裡新增或修改過的 docx：「給放射科」名單（姓名、病歷號、計畫手術與日期、主治、報告醫師、Discussion）與「癌症討論會紀錄單」（PSA、切片、MRI、體積、bone scan、risk、預定手術）。新病人加入個案追蹤；已在表上的只補空白欄位、更新討論會日期與狀態（待討論→已排入／已討論-待手術），不覆蓋已填的內容。'],
    ['統計', '依「個案追蹤」自動計算：已討論人數、場數、每月依癌別、依主治／報告醫師、討論後結果、手術前沒上討論會的個案。B2、D2 可設定期間。癌別依預定手術／診斷的關鍵字自動判斷。'],
    ['一個個案的流程', '刀表匯入後自動加入（狀態＝待討論）→ 排進討論會：填 F 討論會日期，狀態改「已排入」→ 討論完：填 K 討論結論、O 會後待辦（P 負責人、Q 期限），狀態改「已討論-待手術」→ 開完刀：狀態改「已開刀」、填 L 病理。不用討論的改「不需手術」。'],
    ['最晚討論日', '手術日前一個週四（自動計算）。'],
    ['匯入新刀表', '用調度板匯入刀表時會自動送到這裡；也可以把 HIS 手術工作表文字貼在隱藏的「刀表匯入」A2。符合主刀與關鍵字（隱藏的「刀表候選」B2、B3）的重大手術會自動加入個案追蹤。'],
    ['同步到交班單', '編輯「個案追蹤」後幾秒內、以及每小時，會把「注意事項」寫進上面設定的交班單分頁。交班單那一頁不要手動改，會被覆蓋。'],
    ['公式從哪裡來', '所有公式與格式由 GitHub repo ttimcom-hue/Uro-OR 的 cancer/CancerSheet.gs 產生，會自動更新。'],
  ];
  sh.getRange(6, 1, rows.length, 2).setValues(rows);
  csHeader_(sh.getRange(6, 1, 1, 2));
  sh.getRange(7, 1, rows.length - 1, 1).setFontWeight('bold');
  sh.getRange(7, 2, rows.length - 1, 1).setWrap(true).setVerticalAlignment('top');
  sh.getRange(7, 1, rows.length - 1, 1).setVerticalAlignment('top');
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
    let dv = 0;
    try { dv = csSyncDv_(ss); } catch (e) { console.log('達文西術前同步失敗', e); }
    let doc;
    try { doc = csSyncDoc_(ss, st); } catch (e) { doc = '同步交班單失敗：' + (e && e.message || e); }
    const msg = (added ? '自動加入 ' + added + ' 位新個案；' : '') + (dv ? '達文西術前新增 ' + dv + ' 位；' : '') + doc;
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

// ---------- 從 Drive 讀討論會檔案 ----------
// 讀 說明!B4 資料夾（含子資料夾）裡近 CS_DRIVE_DAYS 天新增／修改的 docx：
//   「…給放射科.docx」：Combined Conference 名單表格（姓名病歷號年齡｜病史｜影像＋Plan｜主治＋報告醫師；下一列 Discussion）
//   「…癌症討論會….docx」：泌尿生殖癌團隊會議個案討論紀錄單（開會日期、主治、姓名、病歷號、Investigation 表、@Surgery）
// 已讀過且沒改過的檔案記在指令碼屬性 ccDriveSeen2（讀失敗的不記，下次重試）。寫入規則：新病人加一列；舊病人只補空白欄位。
const CS_DRIVE_DAYS = 45;
const CS_DOCX = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';

function cancerDriveScan() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return '忙碌中，稍後再試';
  try {
    const ss = SpreadsheetApp.openById(CANCER_ID);
    const st = csSettings_(ss);
    if (!st.driveId) return '未設定討論會 Drive 資料夾（說明!B4）';
    const props = PropertiesService.getScriptProperties();
    const seen = JSON.parse(props.getProperty('ccDriveSeen2') || '{}');
    const since = Utilities.formatDate(new Date(Date.now() - CS_DRIVE_DAYS * 864e5), 'UTC', "yyyy-MM-dd'T'HH:mm:ss");
    const files = [];
    [`mimeType = '${CS_DOCX}' and (title contains '放射科' or title contains '癌症討論會' or title contains '紀錄')`,
      `mimeType = 'application/pdf' and (title contains '討論會' or title contains 'GU' or title contains '紀錄')`].forEach(q => {
      const it = DriveApp.searchFiles(`modifiedDate > '${since}' and ` + q);
      while (it.hasNext()) {
        const f = it.next(), m = f.getLastUpdated().getTime();
        if (seen[f.getId()] === m || !csInFolder_(f, st.driveId)) continue;
        files.push(f);
      }
    });
    const pts = [];
    let bad = 0;
    files.forEach(f => {
      try {
        if (f.getMimeType() === 'application/pdf') pts.push(...csParseRecordText_(csPdfText_(f), f));
        else pts.push(...csParseMeeting_(csDocx_(f.getBlob()), f));
        seen[f.getId()] = f.getLastUpdated().getTime();   // 讀成功才記住，失敗的下次重試
      } catch (e) { bad++; console.log('讀不到 ' + f.getName(), e); }
    });
    let msg = '讀了 ' + (files.length - bad) + ' 個新檔案' + (bad ? `（${bad} 個讀不到，下次重試）` : '');
    if (pts.length) {
      const r = csApplyMeeting_(ss, pts);
      SpreadsheetApp.flush();
      csSyncDv_(ss);
      const dv = csDvFill_(ss, pts);
      msg += `：新增 ${r.add} 位、更新 ${r.upd} 位` + (dv ? `、達文西術前補 ${dv} 格` : '');
    }
    const keep = Object.entries(seen).sort((a, b) => b[1] - a[1]).slice(0, 300);
    props.setProperty('ccDriveSeen2', JSON.stringify(Object.fromEntries(keep)));
    console.log(msg);
    return msg;
  } finally { lock.releaseLock(); }
}

function csInFolder_(f, root) {
  let lvl = [f];
  for (let d = 0; d < 6 && lvl.length; d++) {
    const nx = [];
    for (const x of lvl) {
      const it = x.getParents();
      while (it.hasNext()) { const p = it.next(); if (p.getId() === root) return true; nx.push(p); }
    }
    lvl = nx;
  }
  return false;
}

// docx → [{p:段落}|{row:[儲存格文字]}]
function csDocx_(blob) {
  const part = Utilities.unzip(blob.setContentType('application/zip')).find(b => b.getName() === 'word/document.xml');
  if (!part) return [];
  const x = part.getDataAsString('UTF-8')
    .replace(/<w:tbl(?:\s[^>]*)?>/g, '\n').replace(/<\/w:tbl>/g, '\n')
    .replace(/<w:(?:tab|br|cr)\/>/g, ' ').replace(/<\/w:p>/g, '¶').replace(/<\/w:tc>/g, '\t').replace(/<\/w:tr>/g, '\n')
    .replace(/<[^>]+>/g, '')
    .replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');
  const items = [];
  x.split('\n').forEach(line => {
    if (line.includes('\t')) {
      const cells = line.split('\t'); cells.pop();
      items.push({ row: cells.map(c => c.replace(/¶/g, ' ').replace(/\s+/g, ' ').trim()) });
    } else line.split('¶').forEach(t => { t = t.replace(/\s+/g, ' ').trim(); if (t) items.push({ p: t }); });
  });
  return items;
}

function csYmd_(y, m, d) { y = +y; return y > 2000 ? new Date(y, +m - 1, +d) : null; }
function csMeetDate_(text, f) {
  let m = text.match(/(20\d\d)[-\/](\d{1,2})[-\/](\d{1,2})\s*0?7[:：]\d\d/) || text.match(/開會日期[:：]\s*(\d{4})\/(\d{1,2})\/(\d{1,2})/);
  if (m) return csYmd_(m[1], m[2], m[3]);
  m = f.getName().match(/(20\d\d)(\d\d)(\d\d)/);
  if (m) return csYmd_(m[1], m[2], m[3]);
  const ps = f.getParents();
  if (ps.hasNext()) {
    const n = ps.next().getName().match(/^(\d\d)(\d\d)$/);
    if (n) { const y = new Date().getFullYear(); let d = new Date(y, +n[1] - 1, +n[2]); if (d - new Date() > 200 * 864e5) d = new Date(y - 1, +n[1] - 1, +n[2]); return d; }
  }
  return null;
}
function csPlanOp_(plan) {
  const m = plan.match(/(20\d\d)\/(\d{1,2})\/(\d{1,2})/);
  return {
    date: m ? csYmd_(m[1], m[2], m[3]) : '',
    op: plan.replace(/(20\d\d)\/\d{1,2}\/\d{1,2}/, '').replace(/\b0?8:?30\b/, '08:30').replace(/\b12:?00\b/, '12:00').replace(/\s+/g, ' ').trim(),
  };
}
const csFirst_ = (s, re) => ((String(s || '').match(re) || [])[1] || '').trim();

function csParseMeeting_(items, f) {
  const text = items.map(i => i.p || i.row.join(' | ')).join('\n');
  const date = csMeetDate_(text, f), src = f.getName().replace(/\.docx$/i, '');
  const out = [];
  if (/討論紀錄單/.test(text)) {
    text.split(/泌尿生殖癌團隊會議個案討論紀錄單/).slice(1).forEach(b => {
      const mrn = csFirst_(b, /病歷號碼[:：]\s*(\d{8})/);
      if (!mrn) return;
      const bd = b.match(/開會日期[:：]\s*(\d{4})\/(\d{1,2})\/(\d{1,2})/);
      const p = { date: (bd && csYmd_(bd[1], bd[2], bd[3])) || date, mrn, src,
        name: csFirst_(b, /姓名[:：]\s*([^\s|]+)/), vs: csFirst_(b, /主治醫師[:：]\s*([^|\n]+)/).replace(/(醫師|主任|部長)/g, '').replace(/\s+/g, ''),
        rep: '', disc: '', opDate: '', op: '', dv: {} };
      const pmh = csFirst_(b, /Past medical history[:：]\s*([^\n]+)/), cf = csFirst_(b, /Clinical findings[:：]\s*([^\n]+)/);
      let mri = '', risk = '', bone = '', tx = '';
      b.split('\n').filter(l => l.includes(' | ')).forEach(l => {
        const c = l.split(' | ').map(x => x.trim()), k = (c[0] || '').toLowerCase(), dt = c[1] || '', res = c.slice(2).join(' ');
        if (/^psa/.test(k)) { const v = csFirst_(res, /PSA[:：]?\s*([\d.]+)/i); const fp = csFirst_(res, /\(([\d.]+\s*%)\)/); if (v) p.dv.psa = v + (dt ? `（${dt}）` : '') + (fp ? `，fPSA ${fp}` : ''); }
        else if (/biopsy/.test(k)) {
          const gs = res.match(/(\d)\s*\+\s*(\d)\s*=\s*\d+/);
          let pos = 0, tot = 0; (res.match(/\b(\d{1,2})\/(\d{1,2})\b/g) || []).forEach(x => { const [a, z] = x.split('/'); pos += +a; tot += +z; });
          const vol = csFirst_(res, /([\d.]+)\s*cc/i);
          if (gs) p.dv.gs = `GS ${gs[1]}+${gs[2]}` + (tot ? `，${pos}/${tot} cores` : '');
          if (dt) p.dv.bx = dt.replace(/-/g, '/');
          if (vol) p.dv.vol = vol + '（切片時 TRUS）';
        }
        else if (/mri/.test(k)) mri = `${dt.replace(/-/g, '/')}：${res}`.slice(0, 200);
        else if (/trus|echo/.test(k) && !p.dv.vol) { const vol = csFirst_(res, /([\d.]+)\s*cc/i); if (vol) p.dv.vol = vol + '（TRUS）'; }
        else if (/bone scan/.test(k)) bone = /no |negative|low probability/i.test(res) ? 'bone scan(-)' : 'bone scan 見報告';
        else if (/risk group/.test(k)) risk = c[1] || '';
        else if (/@\s*surgery/.test(k)) { const m = dt.match(/(20\d\d)[-\/](\d{1,2})[-\/](\d{1,2})/); p.opDate = m ? csYmd_(m[1], m[2], m[3]) : ''; p.op = res; }
        else if (/^treatment/.test(k)) tx = c[1] || '';
      });
      if (mri) p.dv.mri = mri;
      if (/TURP|TUEP|HIFU|hernia|TURBT/i.test(pmh)) p.dv.hx = pmh;
      p.dv.note = [risk && risk + ' risk', bone, pmh && 'PMH：' + pmh].filter(Boolean).join('；');
      p.diag = [cf, csFirst_(mri, /(y?c?T[1-4x][a-c]?\s*N[0-9x])/i), p.dv.gs].filter(Boolean).join('，').slice(0, 250);
      if (tx) p.disc = tx;
      out.push(p);
    });
    return out;
  }
  const rows = items.filter(i => i.row).map(i => i.row);
  rows.forEach((r, k) => {
    const m = (r[0] || '').match(/^(.{2,6}?)\s*(\d{8})\s*(\d{1,3})\s*y\/o\s*([MF])/i);
    if (!m) return;
    const hist = r[1] || '', img = r[2] || '', vc = r[3] || '';
    const plan = csFirst_(img, /Plan[:：]\s*(.+)$/i), io = img.replace(/Plan[:：].*$/i, '').trim();
    const po = csPlanOp_(plan);
    const nx = rows[k + 1] || [];
    const disc = /^Discussion/i.test(nx[0] || '') ? nx.join(' ').replace(/^Discussion[:：]?\s*/i, '').trim() : '';
    const cT = csFirst_(io, /(y?c?T[1-4x][a-c]?\s*N[0-9x])/i);
    const gs = csFirst_(hist, /Gleason score\s*[=:：]?\s*(\d\s*\+\s*\d)/i).replace(/\s/g, '');
    const psa = csFirst_(hist + ' ' + io, /PSA[^0-9]{0,15}([\d.]+)\s*ng/i);
    const mri = (io.match(/(20\d\d\/\d\d\/\d\d)\s*Prostate MRI(.*?)(?=20\d\d\/\d\d\/\d\d|$)/i) || []);
    out.push({
      date, mrn: m[2], name: m[1].trim(), src, disc,
      vs: csFirst_(vc, /^(.+?)(部長|主任|醫師)/), rep: csFirst_(vc, /報告醫師\s*(.+?)醫師/),
      opDate: po.date, op: po.op,
      diag: [hist, cT].filter(Boolean).join('，').slice(0, 250),
      dv: {
        psa, gs: gs ? 'GS ' + gs : '', mri: mri[0] ? `${mri[1]}：${mri[2].trim()}`.slice(0, 200) : '',
        vol: csFirst_(io, /size[:：]?\s*([\d.]+)\s*(?:ml|cc)/i), adt: /ADT/i.test(hist) ? '有（' + csFirst_(hist, /((?:s\/p\s*)?(?:neoadjuvant|neo-)?\s*ADT[^,，;；]*)/i) + '，藥物／日期待補）' : '',
        hx: (hist.match(/s\/p\s*(TURP|TUEP|TUIP|HIFU|TURBT|RARC)[^,，;；]*/gi) || []).join('；'),
      },
    });
  });
  return out;
}

// 寫進個案追蹤：新病人加一列；舊病人只補空白、更新討論會日期與狀態
function csApplyMeeting_(ss, pts) {
  const trk = ss.getSheetByName(CS.TRK), n = trk.getLastRow();
  const v = n > 1 ? trk.getRange(2, 1, n - 1, 17).getValues() : [];
  const idx = new Map();
  v.forEach((r, i) => { const k = String(r[2]).trim().replace(/^'/, ''); if (k) idx.set(k, i); });
  let last = 1;
  v.forEach((r, i) => { if (String(r[1]).trim() || String(r[2]).trim()) last = i + 2; });
  const today = csToday_(), md = d => Utilities.formatDate(d, 'Asia/Taipei', 'M/d');
  let add = 0, upd = 0;
  pts.sort((a, b) => (a.date || 0) - (b.date || 0)).forEach(p => {
    const i = idx.get(p.mrn);
    if (i === undefined) {
      const st = p.date && p.date > today ? '已排入' : (p.opDate && p.opDate < today ? '待確認' : '已討論-待手術');
      const r0 = ++last;
      trk.getRange(r0, 1, 1, 6).setValues([[st, p.name, "'" + p.mrn, p.opDate || '', p.op || '', p.date || '']]);
      trk.getRange(r0, 7, 1, 2).setFormulas([[csTrackerG_(r0), csTrackerH_(r0)]]);
      trk.getRange(r0, 9, 1, 6).setValues([[p.vs || '', p.diag || '', p.disc || '', '', p.rep || '', 'Drive 匯入：' + p.src]]);
      const row = new Array(17).fill(''); row[0] = st; row[2] = p.mrn; row[3] = p.opDate; row[5] = p.date;
      v.push(row); idx.set(p.mrn, v.length - 1); add++;
      return;
    }
    const r = v[i], row = i + 2, set = (c, val) => { if (val === '' || val == null) return; r[c - 1] = val; trk.getRange(row, c).setValue(val); ch = true; };
    let ch = false;
    const blank = c => String(r[c - 1] instanceof Date ? 'd' : r[c - 1] || '').trim() === '';
    if (p.date) {
      if (blank(6)) set(6, p.date);
      else if (r[5] instanceof Date && p.date > r[5]) { set(6, p.date); set(14, String(r[13] || '') + '；' + md(p.date) + ' 再討論'); }
      const st = String(r[0]);
      if (st === '待討論' || (st === '已排入' && p.date <= today)) set(1, p.date > today ? '已排入' : '已討論-待手術');
    }
    if (blank(4) && p.opDate) set(4, p.opDate);
    if (blank(5) && p.op) set(5, p.op);
    if (blank(9) && p.vs) set(9, p.vs);
    if ((blank(10) || /資料待補/.test(String(r[9]))) && p.diag) set(10, p.diag);
    if (blank(11) && p.disc) set(11, p.disc);
    if (blank(13) && p.rep) set(13, p.rep);
    if (ch) upd++;
  });
  return { add, upd };
}

// 「達文西術前」只補空白欄位（或把自動帶入的簡短版換成檔案裡較完整的）；回傳補了幾格
function csDvFill_(ss, pts) {
  const sh = ss.getSheetByName(CS.DV);
  if (!sh || sh.getLastRow() < 2) return 0;
  const v = sh.getRange(2, 1, sh.getLastRow() - 1, CS_DV_N).getValues();
  const idx = new Map(v.map((r, i) => [String(r[3]).trim().replace(/^'/, ''), i]));
  const col = { psa: 9, gs: 10, bx: 11, mri: 13, vol: 14, adt: 15, hx: 16, note: 18 };
  let n = 0;
  pts.forEach(p => {
    const i = idx.get(p.mrn);
    if (i === undefined || !p.dv) return;
    Object.entries(col).forEach(([k, c]) => {
      const val = p.dv[k];
      if (!val) return;
      const cur = String(v[i][c - 1] instanceof Date ? 'd' : v[i][c - 1] || '').trim();
      // 空白、或只是自動帶入的簡短版（新內容完整包含舊內容）才寫；手動填的不覆蓋
      if (cur && !(k === 'note' && cur === '自動帶入（請核對）') && !(String(val).length > cur.length && String(val).includes(cur))) return;
      const w = k === 'bx' ? (csYmd_(...(String(val).split('/'))) || val) : val;
      sh.getRange(i + 2, c).setValue(w); v[i][c - 1] = w; n++;
    });
  });
  return n;
}

// ---------- PDF ----------
// Apps Script 不能直接讀 PDF：用 Drive API 複製成 Google 文件（自動轉文字）→ 讀內文 → 立刻永久刪除暫存檔。
// 需要 drive.file 權限（只能建立／刪除程式自己建的檔案）。
function csPdfText_(f) {
  // 上傳一份 PDF 並要求 Drive 轉成 Google 文件（新檔是程式建立的，drive.file 權限即可），讀完刪除
  const h = { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, b = 'cc' + Date.now();
  const meta = JSON.stringify({ name: '暫存-討論會轉文字（會自動刪除）', mimeType: 'application/vnd.google-apps.document' });
  const body = Utilities.newBlob(`--${b}\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n${meta}\r\n--${b}\r\nContent-Type: application/pdf\r\n\r\n`).getBytes()
    .concat(f.getBlob().getBytes()).concat(Utilities.newBlob(`\r\n--${b}--`).getBytes());
  const r = UrlFetchApp.fetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&ocrLanguage=zh-TW&fields=id', {
    method: 'post', contentType: 'multipart/related; boundary=' + b, payload: body, headers: h, muteHttpExceptions: true });
  if (r.getResponseCode() >= 300) throw new Error('PDF 轉文字失敗 ' + r.getResponseCode() + ' ' + r.getContentText().slice(0, 200));
  const id = JSON.parse(r.getContentText()).id;
  try { return DocumentApp.openById(id).getBody().getText(); }
  finally { UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + id, { method: 'delete', headers: h, muteHttpExceptions: true }); }
}

// 紀錄單轉成純文字後（PDF）表格會攤平，改用整段文字比對。中英文版紀錄單都適用
function csParseRecordText_(text, f) {
  const src = f.getName().replace(/\.(pdf|docx)$/i, ''), fdate = csMeetDate_(text, f);
  const D = '(20\\d\\d)[.\\/-](\\d{1,2})[.\\/-](\\d{1,2})';
  const dt = (b, re) => { const m = b.match(re); return m ? csYmd_(m[1], m[2], m[3]) : ''; };
  const ds = d => d ? Utilities.formatDate(d, 'Asia/Taipei', 'yyyy/MM/dd') : '';
  const out = [];
  text.split(/泌尿生殖癌團隊會議個案討論紀錄單|Case Discussion Record/i).forEach(b => {
    const mrn = csFirst_(b, /(?:病歷號碼|Medical Record Number)[:：]\s*(\d{8})/i);
    if (!mrn) return;
    const one = b.replace(/\s+/g, ' ');
    const p = { mrn, src, rep: '', disc: '', dv: {},
      date: dt(one, new RegExp('(?:開會日期|Meeting Date)[:：]\\s*' + D, 'i')) || fdate,
      name: csFirst_(one, /(?:姓名|Name)[:：]\s*(.+?)\s*(?:性別|Gender)/i).replace(/\s+/g, ''),
      vs: csFirst_(one, /(?:主治醫師|Attending Physician)[:：]\s*(.+?)\s*(?:姓名|Name[:：])/i).replace(/(醫師|主任|部長|\s)/g, ''),
    };
    // PSA：取最後一筆（排除 Free PSA）
    // PSA：有多筆就寫「最高 → 最近」（例如 ADT 前後）；排除 Free PSA
    const ps = []; let m; const re = /(?<!Free )PSA\s*(?:[:：]\s*)?(?:(20\d\d[.\/-]\d{1,2}[.\/-]\d{1,2})\s*)?([\d.]+)\s*(?:ng)?/gi;
    while ((m = re.exec(one))) if (+m[2] > 0) ps.push(m[2] + (m[1] ? `（${m[1].replace(/[.-]/g, '/')}）` : ''));
    if (ps.length) {
      const top = ps.reduce((a, b) => parseFloat(b) > parseFloat(a) ? b : a), last = ps[ps.length - 1];
      p.dv.psa = top === last ? last : `${top} → ${last}`;
    }
    const gs = [...one.matchAll(/(?:Gleason[’'`s]*\s*score|GS)\s*[:：=]?\s*(\d)\s*\+\s*(\d)/gi)].map(x => +x[1] + +x[2] * 0.1).sort((a, b) => b - a)[0];
    if (gs) p.dv.gs = 'GS ' + Math.floor(gs) + '+' + Math.round((gs % 1) * 10);
    const bx = dt(one, new RegExp('(?:TRUS )?biopsy\\s*' + D, 'i'));
    if (bx) p.dv.bx = ds(bx);
    const mri = one.match(new RegExp('(?:Prostate )?MRI\\s*' + D + '\\s*(.{0,180}?)(?=\\s(?:Bone scan|Chi Mei|Risk group|Plan|@|Investigation)|$)', 'i'));
    if (mri) p.dv.mri = `${mri[1]}/${mri[2]}/${mri[3]}：${mri[4].trim()}`;
    const vol = csFirst_(one, /(?:TRUS echo\s*\S+\s*(?:About\s*)?|Prostate[:：]\s*)([\d.]+)\s*(?:cc|ml)/i);
    if (vol) p.dv.vol = vol + '（TRUS）';
    const adt = one.match(new RegExp('ADT\\s*' + D + '\\s*([A-Za-z ]{3,30})', 'i'));
    if (adt) p.dv.adt = `${adt[4].trim()}（${adt[1]}/${adt[2]}/${adt[3]} 起）`;
    const pmh = csFirst_(one, /Past medical history[:：]\s*(.+?)\s*(?:Presenting|Clinical findings|ECOG|Investigation|Date Result|$)/i);
    if (/TURP|TUEP|HIFU|hernia|TURBT/i.test(pmh)) p.dv.hx = pmh;
    const risk = csFirst_(one, /Risk group\s*[:：]?\s*((?:Very |Unfavorable |Favorable )?(?:low|intermediate|high)(?: risk)?)/i);
    const bone = /Bone scan[^@]{0,120}?(no (obvious )?(bone )?metast|negative|low probability)/i.test(one) ? 'bone scan(-)' : '';
    const cT = csFirst_(one, /\b(y?c?T[1-4x][a-c]?\s*N[0-9x](?:\s*M[0-9x])?)\b/i);
    p.dv.note = [risk && risk.replace(/ risk$/i, '') + ' risk', bone, cT, pmh && 'PMH：' + pmh].filter(Boolean).join('；').slice(0, 200);
    const pl = one.match(new RegExp('(?:@\\s*Plan|@\\s*Surgery|Plan)\\s*' + D + '(?:\\s*' + D.replace(/\(/g, '(?:') + ')?\\s*(\\d\\d:\\d\\d)?\\s*(.{0,80}?)(?=\\s奇美|\\s泌尿生殖|$)', 'i'));
    p.opDate = pl ? csYmd_(pl[1], pl[2], pl[3]) : '';
    p.op = pl ? ((pl[5] || '') + ' ' + pl[4]).trim() : '';
    if (pl && !pl[4]) p.op = (pl[5] || '').trim();
    const stop = '\\s*(?:ECOG|Investigation|Date\\b|Past|Clinical findings|$)';
    const cf = csFirst_(one, new RegExp('Clinical findings[:：]\\s*(.*?)' + stop, 'i')) || csFirst_(one, new RegExp('Presenting symptoms[:：]\\s*(.*?)' + stop, 'i'));
    p.diag = [cf, cT, p.dv.gs].filter(Boolean).join('，').slice(0, 250);
    out.push(p);
  });
  return out;
}

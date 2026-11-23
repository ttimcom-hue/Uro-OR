// 用 JS 重現試算表「刀表解析」＋「刀表候選」的規則，方便改規則前先測試。
// 用法：node cancer/parse.js cancer/sample-or.txt [主刀] [已在追蹤表的病歷號,逗號分隔]
const fs = require('fs');

const KEYWORDS = /davinci|prostatectomy|nephr\w*ectomy|nephrouret|\bRNU\b|cystectomy|RARC|adrenal|thrombect|retroperitoneal/i;

function rocToDate(n) {
  n = Number(n);
  const y = Math.floor(n / 10000) + 1911, m = Math.floor(n / 100) % 100, d = n % 100;
  return `${y}-${String(m).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
}

// 對應 刀表解析 A–K 欄
function parseOR(text) {
  const L = String(text).replace(/\r/g, '').split('\n').map(s => s.trim());
  const at = i => (i < L.length ? L[i] : '');
  const out = [];
  L.forEach((line, s) => {
    if (!/^\d{8}$/.test(line)) return;                       // A：起始列＝8 位數病歷號
    const d = /^1\d{6}$/.test(at(s + 3)) ? s + 3 : s + 4;    // B：日期列（有床號 +3，沒床號 +4）
    const loc = at(d + 3) === '' ? d + 5 : d + 4;            // C：位置列（有「Y」癌症註記 +4）
    out.push({
      mrn: line, name: at(s + 1), date: rocToDate(at(d)), time: at(d + 1),
      surgeon: at(d + 2), cancer: loc === d + 4 ? at(d + 3) : '',
      op: at(loc + 1), source: at(loc + 3),
    });
  });
  return out;
}

// 對應 刀表候選：主刀＋關鍵字篩選，依手術日排序，去重複
function candidates(rows, surgeon = '黃冠華', tracked = new Set()) {
  const seen = new Set();
  return rows
    .filter(r => r.surgeon.includes(surgeon) && KEYWORDS.test(r.op))
    .map(r => ({ status: tracked.has(r.mrn) ? '已在追蹤表' : '新個案', ...r }))
    .filter(r => { const k = [r.mrn, r.date, r.op].join('|'); if (seen.has(k)) return false; seen.add(k); return true; })
    .sort((a, b) => a.date.localeCompare(b.date));
}

module.exports = { parseOR, candidates, rocToDate };

if (require.main === module) {
  const [file, surgeon, tracked] = process.argv.slice(2);
  if (!file) { console.error('用法：node cancer/parse.js <刀表.txt> [主刀] [病歷號,...]'); process.exit(1); }
  const rows = parseOR(fs.readFileSync(file, 'utf8'));
  const cand = candidates(rows, surgeon || '黃冠華', new Set((tracked || '').split(',').filter(Boolean)));
  console.log(`解析 ${rows.length} 筆手術，符合 ${cand.length} 筆`);
  console.table(cand.map(r => ({ 比對: r.status, 姓名: r.name, 病歷號: r.mrn, 手術日: r.date, 時: r.time, 手術: r.op, 癌: r.cancer, 來源: r.source })));
}

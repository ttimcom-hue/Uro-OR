// 測試「圖片辨識」的議程解析（CancerSheet.gs 的 csParseAgendaText_）
// 用法：node cancer/ocr-test.js [OCR 文字檔]，預設 cancer/sample-ocr.txt
// sample-ocr.txt 是用「假資料」做的議程圖片交給 Google Drive OCR 後的真實輸出（姓名、病歷號、醫師都是假的）。
const fs = require('fs'), vm = require('vm'), path = require('path');
const ctx = { console, Utilities: { formatDate: d => d.toISOString().slice(0, 10) } };
vm.createContext(ctx);
vm.runInContext(fs.readFileSync(path.join(__dirname, 'CancerSheet.gs'), 'utf8') + '\nthis.parse = csParseAgendaText_;', ctx);
const file = process.argv[2] || path.join(__dirname, 'sample-ocr.txt');
const f = { getName: () => 'IMG_0001.jpg', getParents: () => ({ hasNext: () => false }) };
const pts = ctx.parse(fs.readFileSync(file, 'utf8'), f);
const ds = d => d ? `${d.getFullYear()}/${d.getMonth() + 1}/${d.getDate()}` : '';
pts.forEach(p => console.log([ds(p.date), p.name, p.mrn, ds(p.opDate), p.op, p.vs, p.rep].join(' | ')));
if (!process.argv[2]) {
  const want = [
    ['2026/11/26', '王大明', '12345678', '2026/12/1', 'MR-Fusion biopsy', '陳一二', '李三四'],
    ['2026/11/26', '林小華', '23456789', '2026/12/3', '08:30 RARP', '陳一二', '張五六'],
    ['2026/11/26', '黃志強', '34567890', '2026/12/8', '08:30 RARP', '陳一二', '趙七八'],
  ];
  const got = pts.map(p => [ds(p.date), p.name, p.mrn, ds(p.opDate), p.op, p.vs, p.rep]);
  const ok = JSON.stringify(got) === JSON.stringify(want);
  console.log(ok ? '✓ 通過' : '✗ 不符，預期：\n' + want.map(r => r.join(' | ')).join('\n'));
  process.exit(ok ? 0 : 1);
}

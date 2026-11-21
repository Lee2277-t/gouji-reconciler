'use strict';
/** 第五轮审查的回归：结论计数、合计行、0 余额未映射、文字格、合并单元格、首屏顺序 */
const path = require('path');
const fs = require('fs');
const vm = require('vm');
const XLSX = require(path.join(__dirname, '..', 'vendor', 'xlsx.full.min.js'));
const Core = require(path.join(__dirname, '..', 'core.js'));

let count = 0;
const fails = [];
function ok(cond, msg) { count++; if (!cond) fails.push(msg); }
function eq(a, b, msg) {
  count++;
  if (JSON.stringify(a) !== JSON.stringify(b)) fails.push(msg + ' — 期望 ' + JSON.stringify(b) + '，实际 ' + JSON.stringify(a));
}

console.log('\n[1 结论行不能把单边重复计数]');
const agg = Core.aggregateItems([
  { code: '1002', name: '银行存款', debitCents: 100000, creditCents: 0 },
  { code: '9999', name: '自定义科目', debitCents: 50000, creditCents: 0 }
]);
const cmp = Core.compareItems(agg, [{ item: '货币资金', cents: 100000 }]);
const out = Core.buildWorkbook(XLSX, {
  client: '', dateText: '2026-11-22',
  bs: { label: 'a', accountCount: 2, leafCount: 2, balanceDiffCents: 0, identityFailed: 0, identityChecked: 2, inferred: 0, missing: 0, skippedRows: 0, level: 'leaf' },
  target: { label: 'b', blockName: '资产', colLabel: '审计前' },
  agg: agg, cmp: cmp
});
const gj = XLSX.utils.sheet_to_json(out.wb.Sheets['勾稽表'], { header: 1, defval: null });
const concl = String((gj.find(r => r[0] === '结论') || [])[4] || '');
ok(/未勾平/.test(concl), '有未映射科目时结论是未勾平');
ok(!/未映射\/读不出/.test(concl), '不该再用含糊的「未映射/读不出」合计');
ok(/未映射科目 1 个/.test(concl), '未映射要单独数：' + concl);
ok(!/单边项目/.test(concl), '这张表没有单边项目，不该出现单边');

console.log('\n[2 合计行不算「有行没计入」]');
const withTotal = [
  ['科目编码', '科目名称', '期末余额'],
  ['1002', '银行存款', 1000],
  ['4001', '实收资本', -1000],
  ['', '合计', 0]
];
const h2 = Core.detectBalanceHeader(withTotal);
const p2 = Core.parseBalanceRows(withTotal, h2, { firstExcelRow: 1 });
eq(p2.problems.length, 0, '合计行不该触发任何警告');
ok(p2.skipped.some(s => s.reason === '汇总行'), '合计行标记为汇总行');
eq(p2.skipped.filter(s => s.reason !== '汇总行' && s.reason !== '空行').length, 0, '没有真正丢行');

console.log('\n[3 合并单元格：有金额但没编码名称的行不能静默消失]');
const merged = [
  ['科目编码', '科目名称', '期末余额'],
  ['1122', '应收账款', 100],
  ['', '', 300]
];
const p3 = Core.parseBalanceRows(merged, Core.detectBalanceHeader(merged), { firstExcelRow: 1 });
ok(p3.problems.some(p => /四位数字/.test(p)), '要报出来');
ok(p3.skipped.some(s => s.reason === '缺科目编码和名称'), '标记为缺编码名称');

console.log('\n[4 零余额的未映射科目不算缺口]');
const agg4 = Core.aggregateItems([
  { code: '1002', name: '银行存款', debitCents: 100000, creditCents: 0 },
  { code: '9999', name: '停用科目', debitCents: 0, creditCents: 0 }
]);
const cmp4 = Core.compareItems(agg4, [{ item: '货币资金', cents: 100000 }]);
eq(agg4.unmapped.length, 1, '它确实在未映射清单里');
eq(cmp4.clean, true, '但余额为 0 不该阻止「全部勾平」');

console.log('\n[5 期末借贷格写文字：取不到，不能当 0，也不能重复计数]');
const txt = [
  ['科目编码', '科目名称', '期末借方', '期末贷方'],
  ['1122', '应收账款', '待补', 0],
  ['4001', '实收资本', 0, 1000]
];
const p5 = Core.parseBalanceRows(txt, Core.detectBalanceHeader(txt), { firstExcelRow: 1 });
eq(p5.accounts[0].closingCents, null, '读不出就是 null');
const msg = p5.problems.join('；');
ok(/文字/.test(msg), '要报出来');
ok(!/2 行/.test(msg), '一行不能报成两行：' + msg);

console.log('\n[6 摘要口径：明细表有差异时首屏差异区不能说「分毫不差」]');
function makeEl(id) {
  const h = {};
  return {
    id, innerHTML: '', value: '', textContent: '', disabled: false, style: {}, files: null, options: [], selectedIndex: -1, _h: h,
    classList: {
      _s: new Set(),
      add(c) { this._s.add(c); }, remove(c) { this._s.delete(c); },
      contains(c) { return this._s.has(c); },
      toggle(c, f) { if (f === undefined) { this._s.has(c) ? this._s.delete(c) : this._s.add(c); } else if (f) this._s.add(c); else this._s.delete(c); }
    },
    addEventListener(t, fn) { (h[t] = h[t] || []).push(fn); },
    click() { (h.click || []).forEach(fn => fn({ target: this })); },
    appendChild() {}, setAttribute() {}, querySelector() { return null; }
  };
}
const els = {};
const doc = {
  getElementById(id) { return els[id] || (els[id] = makeEl(id)); },
  body: makeEl('body'), createElement: () => makeEl('x'), addEventListener() {}
};
doc.body.classList = makeEl('body').classList;
const store = {};
const ls = { getItem(k) { return store[k] === undefined ? null : store[k]; }, setItem(k, v) { store[k] = String(v); }, removeItem(k) { delete store[k]; } };
let nextPayload = null;
function FR() { this.onload = null; this.onerror = null; this.result = null; }
FR.prototype.readAsArrayBuffer = function () {
  const self = this;
  setImmediate(() => { self.result = nextPayload; self.onload({ target: { result: nextPayload } }); });
};
const sandbox = {
  window: { localStorage: ls, document: doc }, document: doc, localStorage: ls,
  XLSX, Core, FileReader: FR, Uint8Array, setImmediate,
  console, Date, JSON, Math, String, Number, Object, Array, parseInt, parseFloat, isNaN, Promise
};
sandbox.window.window = sandbox.window;
vm.createContext(sandbox);
vm.runInContext(fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8'), sandbox, { filename: 'app.js' });

function buf(rows, name) {
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), name || 'S');
  return new Uint8Array(XLSX.write(wb, { type: 'array', bookType: 'xlsx' }));
}
async function load(id, data) {
  nextPayload = data;
  const el = doc.getElementById(id);
  el.files = [{ name: 'x.xlsx' }];
  (el._h.change || []).forEach(fn => fn({ target: el }));
  await new Promise(r => setTimeout(r, 70));
}

(async function () {
  await load('bsFile', buf([
    ['科目编码', '科目名称', '期初余额', '本期借方', '本期贷方', '期末余额'],
    ['1002', '银行存款', 1000, 0, 0, 1000],
    ['1122', '应收账款', 300, 0, 0, 300],
    ['4001', '实收资本', 1300, 0, 0, 1300]
  ]));
  await load('stFile', buf([
    ['项目', '期末余额'],
    ['货币资金', 1000],
    ['应收账款', 300],
    ['实收资本', -1300]
  ]));
  ok(/分毫不差|全部勾平/.test(doc.getElementById('diffBox').innerHTML), '先确认两边是对上的');
  ok(!/分毫不差/.test(doc.getElementById('detailBox').innerHTML) || doc.getElementById('detailBox').innerHTML === '',
    '还没载入明细表时明细区是空的');

  // 再载入一张差 99 万的明细表：首屏差异区必须立刻反映，不能还挂着「分毫不差」
  // 明细表按「一个项目一张表」组织，工作表名就是项目名
  await load('dtFile', buf([
    ['摘要', '金额'],
    ['甲公司', 629300],
    ['乙公司', 670000],
    ['合  计', 1299300]
  ], '应收账款'));
  ok(!/分毫不差/.test(doc.getElementById('diffBox').innerHTML),
    '明细表有差异时，差异区不能说「所有项目都分毫不差」');


console.log('\n[7 试算表分区标题不算坏行]');
const stSec = [
  ['项目', '期末余额'],
  ['流动资产：', ''],
  ['货币资金', 1000],
  ['流动负债：', ''],
  ['实收资本', -1000]
];
const hsec = Core.detectStatementHeader(stSec);
const bsec = Core.parseStatementBlock(stSec, hsec, 0, 0, { firstExcelRow: 1 });
const secBad = (bsec.skipped || []).filter(s => s.reason !== '空行' && s.reason !== '分区标题');
eq(secBad.length, 0, '「流动资产：」这种分区标题不该算坏行');
eq(bsec.items.length, 2, '两个项目都读到了');

console.log('\n[8 没有金额的行不报警（续表/分隔不影响合计）]');
const noMoney = [
  ['科目编码', '科目名称', '期末余额'],
  ['1002', '银行存款', 1000],
  ['—', '续表', ''],
  ['4001', '实收资本', -1000]
];
const pnm = Core.parseBalanceRows(noMoney, Core.detectBalanceHeader(noMoney), { firstExcelRow: 1 });
eq(pnm.problems.length, 0, '没有金额的行不报警');

console.log('\n[9 明细页与主表结论不能互相矛盾]');
const out9 = Core.buildWorkbook(XLSX, {
  client: '某公司', dateText: '2026-11-22',
  bs: { label: 'a', accountCount: 2, leafCount: 2, balanceDiffCents: 0, identityFailed: 0, identityChecked: 2, inferred: 0, missing: 0, skippedRows: 0, level: 'leaf' },
  target: { label: 'b', blockName: '资产', colLabel: '审计前' },
  agg: { items: [], unmapped: [], missing: [] },
  cmp: { diffs: [], matched: [], onlyAgg: [], onlyTarget: [], targetTotals: [], diffTotalCents: 0, unbalancedCents: 0, singleSideAbsCents: 0, singleSideCount: 0, clean: true, gapCount: 0 },
  detail: { diffs: [], matched: [], failed: [], onlyDetail: [], onlyAgg: [{ item: '应收账款', cents: 100 }] }
});
const gj9 = XLSX.utils.sheet_to_json(out9.wb.Sheets['勾稽表'], { header: 1, defval: null });
const c9 = String((gj9.find(r => r[0] === '结论') || [])[4] || '');
const dt9 = XLSX.utils.sheet_to_json(out9.wb.Sheets['明细表勾稽'], { header: 1, defval: null });
const c9d = String((dt9.find(r => r[0] === '结论') || [])[4] || '');
ok(/未勾平/.test(c9), '明细表里缺项目时，主表结论不能写全部勾平');
ok(/未勾平/.test(c9d), '明细页结论也要说未勾平');
ok(!/全部勾平/.test(c9) && !/明细表全部勾平/.test(c9d), '两页不能一边说勾平一边不勾平');
ok(/被审计单位/.test(JSON.stringify(dt9[0])), '明细表勾稽页也要有抬头');

console.log('\n[10 一行文字只报一次（本期/期末/两处都写）]');
[
  ['本期借方文字', [['科目编码','科目名称','方向','期初余额','本期借方','本期贷方','期末余额'],['1002','银行存款','借',1000,'见附注',0,1200]]],
  ['期末余额文字', [['科目编码','科目名称','方向','期初余额','本期借方','本期贷方','期末余额'],['1002','银行存款','借',1000,0,0,'见附注']]],
  ['同一行两处文字', [['科目编码','科目名称','方向','期初余额','本期借方','期末余额'],['1002','银行存款','借',1000,'见附注','见附注']]]
].forEach(function (c) {
  const p = Core.parseBalanceRows(c[1], Core.detectBalanceHeader(c[1]), { firstExcelRow: 1 });
  const m = /(\d+) 行的期末余额或发生额/.exec(p.problems.join(''));
  ok(m && m[1] === '1', c[0] + '：一行只能报一行（实际 ' + (m ? m[1] : '没报') + '）');
});

console.log('\n[11 没有金额的行不进 skippedRows]');
const noMoney2 = [
  ['科目编码', '科目名称', '期末余额'],
  ['1002', '银行存款', 1000],
  ['—', '', ''],
  ['4001', '实收资本', -1000]
];
const pnm2 = Core.parseBalanceRows(noMoney2, Core.detectBalanceHeader(noMoney2), { firstExcelRow: 1 });
const moneyRows = (pnm2.skipped || []).filter(r => r.money === true && r.reason !== '汇总行');
eq(moneyRows.length, 0, '没有金额的行不该算「有行没计入」');
eq(pnm2.problems.length, 0, '也不该产生警告');

  console.log('\n[12 界面：明细表只覆盖部分项目时首屏不能说「分毫不差」]');
  await load('tFile', buf([
    ['摘要', '金额'],
    ['（合计）', 1000]
  ], '货币资金'));
  ok(!/分毫不差/.test(doc.getElementById('diffBox').innerHTML),
    '明细表缺项目时，界面结论也要说未勾平');
  ok(/明细表/.test(doc.getElementById('diffBox').innerHTML),
    '并且要说清是明细表的问题');

console.log('\n[13 对方表读不到金额的行，要进底稿结论]');
const out13 = Core.buildWorkbook(XLSX, {
  client: '', dateText: '2026-11-22',
  bs: { label: 'a', accountCount: 2, leafCount: 2, balanceDiffCents: 0, identityFailed: 0, identityChecked: 2, inferred: 0, missing: 0, skippedRows: 0, level: 'leaf' },
  target: { label: 'b', blockName: '资产', colLabel: '审计前' },
  targetSkipped: 1,
  agg: { items: [], unmapped: [], missing: [] },
  cmp: { diffs: [], matched: [], onlyAgg: [], onlyTarget: [], targetTotals: [], diffTotalCents: 0, unbalancedCents: 0, singleSideAbsCents: 0, singleSideCount: 0, clean: true, gapCount: 0 }
});
const gj13 = XLSX.utils.sheet_to_json(out13.wb.Sheets['勾稽表'], { header: 1, defval: null });
const c13 = String((gj13.find(r => r[0] === '结论') || [])[4] || '');
ok(/未勾平/.test(c13), '对方表有行读不到时，底稿不能写全部勾平：' + c13);
ok(/对方表/.test(c13), '并且要说明是对方表的问题');

console.log('\n[14 明细表：0 余额的项目不该造成「未勾平」]');
const rec14 = Core.reconcileDetail([
  { name: '应收账款', rows: [['项目', '金额'], ['合计', 300]], total: 300 }
], { items: [{ item: '应收账款', cents: 300, accounts: [] }, { item: '长期待摊费用', cents: 0, accounts: [] }] }, { col: 1 });
eq((rec14.onlyAgg || []).length, 0, '0 余额的项目不算「明细表里没有」');

  console.log('\n[15 非 A1 起点 + #REF! 错误格：界面与底稿要一致]');
  const wsErr = {
    '!ref': 'B1:C3',
    B1: { t: 's', v: '项目' }, C1: { t: 's', v: '金额' },
    B2: { t: 's', v: '货币资金' }, C2: { t: 'n', v: 1000 },
    B3: { t: 's', v: '其他非流动资产' }, C3: { t: 'e', v: 23, w: '#REF!' }
  };
  await load('stFile', new Uint8Array(XLSX.write({ SheetNames: ['ST'], Sheets: { ST: wsErr } }, { type: 'array', bookType: 'xlsx' })));
  const stText = doc.getElementById('stSummary').innerHTML;
  ok(/没有读到金额/.test(stText), '错误格要报到界面：' + stText.replace(/<[^>]+>/g, '').slice(0, 80));
  ok(/还没勾平/.test(doc.getElementById('diffBox').innerHTML), '界面结论不能说勾平');
  let written = null;
  const realWrite = sandbox.XLSX.writeFile;
  sandbox.XLSX.writeFile = function (wb) { written = wb; };
  doc.getElementById('genBtn').click();
  sandbox.XLSX.writeFile = realWrite;
  ok(written, '导出要真的写出工作簿（genInfo：' + doc.getElementById('genInfo').textContent + '）');
  if (written) {
    const gj15 = XLSX.utils.sheet_to_json(written.Sheets['勾稽表'], { header: 1, defval: null });
    const c15 = String((gj15.find(r => r[0] === '结论') || [])[4] || '');
    ok(/未勾平/.test(c15), '底稿结论也要说未勾平：' + c15);
    ok(/对方表/.test(c15), '底稿要指明是对方表的问题');
    const help15 = XLSX.utils.sheet_to_json(written.Sheets['说明'], { header: 1, defval: null });
    ok(help15.some(r => /对方表没读到/.test(String(r[0]))), '说明页要留痕');
  }
  console.log('\n' + '='.repeat(56));
  if (fails.length) {
    console.log('失败 ' + fails.length + ' / ' + count + ' 项：');
    fails.forEach((f, i) => console.log('  ' + (i + 1) + '. ' + f));
    process.exit(1);
  }
  console.log('第五轮修复回归全部通过：' + count + ' 项断言');
  console.log('='.repeat(56));
})();

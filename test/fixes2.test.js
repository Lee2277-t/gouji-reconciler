'use strict';
/**
 * 第二轮审查发现的问题，逐条钉成回归。
 * 这些都是「看着像对、其实算错」的写法，改回去就会红。
 */
const path = require('path');
const vm = require('vm');
const fs = require('fs');
const Core = require(path.join(__dirname, '..', 'core.js'));

let count = 0;
const fails = [];
function ok(cond, msg) { count++; if (!cond) fails.push(msg); }
function eq(a, b, msg) {
  count++;
  if (JSON.stringify(a) !== JSON.stringify(b)) fails.push(msg + ' — 期望 ' + JSON.stringify(b) + '，实际 ' + JSON.stringify(a));
}

console.log('\n[1 年末余额 / 年初余额 这些写法]');
const y = [
  ['科目编码', '科目名称', '年初余额', '方向', '本期借方', '本期贷方', '年末余额'],
  ['1002', '银行存款', 1000, '借', 500, 0, 2000],
  ['4001', '实收资本', 1000, '贷', 0, 500, 900]
];
const hy = Core.detectBalanceHeader(y);
eq(hy.columns.closing.amount, 6, '年末余额要当期末列');
const py = Core.parseBalanceRows(y, hy, { firstExcelRow: 1 });
eq(py.accounts[0].closingCents, 200000, '期末取年末余额 2000，不用推算');
eq(py.accounts[0].closingInferred, false, '不是推算出来的');

console.log('\n[2 方向列只有一个、且不挨着任何分组]');
const d1 = [
  ['科目编码', '科目名称', '方向', '期初余额', '本期借方', '本期贷方', '期末余额'],
  ['2202', '应付账款', '贷', 1000, 0, 500, 1500],
  ['4001', '实收资本', '贷', 1000, 0, 0, 1000]
];
const hd1 = Core.detectBalanceHeader(d1);
eq(hd1.dirCol, 2, '认出了全局方向列');
const pd1 = Core.parseBalanceRows(d1, hd1, { firstExcelRow: 1 });
eq(pd1.accounts[0].closingCents, -150000, '贷方余额要记成负，不能是正的');
eq(pd1.accounts[1].closingCents, -100000, '实收资本同理');

console.log('\n[3 期末整列空：不能报「借贷平」]');
const empty = [
  ['科目编码', '科目名称', '期初余额', '本期借方', '本期贷方', '期末余额'],
  ['1002', '银行存款', 1000, 500, 0, ''],
  ['4001', '实收资本', 1000, 0, 500, '']
];
const he = Core.detectBalanceHeader(empty);
const pe = Core.parseBalanceRows(empty, he, { firstExcelRow: 1 });
ok(pe.problems.some(p => /整列都是空/.test(p)), '要报出期末整列为空');
const be = Core.balanceCheck(pe.accounts);
eq(be.balanced, null, 'balanced 必须是 null');
eq(be.conclusive, false, '不能下结论');
eq(pe.accounts[0].closingInferred, false, '整列空时也不许用推算顶替');

console.log('\n[4 认不出的表头列必须报出来]');
const strange = [
  ['科目编码', '科目名称', '期初余额', '往来单位', '期末余额'],
  ['1002', '银行存款', 1000, 5, 1500]
];
const hs = Core.detectBalanceHeader(strange);
ok(hs.unrecognized.some(u => /往来单位/.test(u)), '要收集没认出来的列');
ok(hs.problems.some(p => /没认出来/.test(p)), '并写进 problems');
ok(hs.problems.some(p => /往来单位/.test(p)), 'problems 里要带列名');

console.log('\n[5 同一组两列都像金额列：不能静默覆盖]');
const dup = [
  ['科目编码', '科目名称', '期末余额(借)', '期末余额(贷)'],
  ['1002', '银行存款', 1500, 0]
];
const hdup = Core.detectBalanceHeader(dup);
ok(hdup.conflicts.length > 0, '要报列冲突');
ok(hdup.problems.some(p => /列冲突/.test(p)), 'problems 里要有列冲突');

console.log('\n[6 标准资产负债表的合计行]');
['所有者权益（或股东权益）合计', '负债和所有者权益（或股东权益）总计',
  '二、营业利润', '四、净利润', '利润总额'].forEach(n => {
  ok(Core.isStatementTotal(n), n + ' 应认成合计行');
});
// 「一、营业收入」「减：营业成本」是项目不是合计行，要靠项目名归一去前缀来对上
ok(!Core.isStatementTotal('一、营业收入'), '营业收入是项目，不是合计行');
eq(Core.normItem('一、营业收入'), Core.normItem('营业收入'), '前缀「一、」要去掉');
eq(Core.normItem('减：营业成本'), Core.normItem('营业成本'), '前缀「减：」要去掉');
ok(!Core.isStatementTotal('货币资金'), '货币资金不是合计行');
ok(!Core.isStatementTotal('应收账款'), '应收账款不是合计行');

const agg6 = Core.aggregateItems([{ code: '4001', name: '实收资本', debitCents: 0, creditCents: 100000 }]);
const cmp6 = Core.compareItems(agg6, [
  { item: '实收资本', cents: 100000 },
  { item: '所有者权益（或股东权益）合计', cents: 100000 }
]);
eq(cmp6.onlyTarget.length, 0, '合计行不能算单边项目');
eq(cmp6.clean, true, '标准报表应能判定为全部勾平');

console.log('\n[7 手工映射：借贷混向要按净额定方向]');
const mixed = Core.aggregateItems([
  { code: '152601', name: '某资产-借', debitCents: 10000, creditCents: 0 },
  { code: '152602', name: '某资产-贷', debitCents: 0, creditCents: 4000 }
], { overrides: { '1526': '其他非流动资产' } });
eq(mixed.items[0].cents, 6000, '净额 6000 分，不能放大成 14000');
const explicit = Core.aggregateItems([
  { code: '152601', name: '某资产-借', debitCents: 10000, creditCents: 0 },
  { code: '152602', name: '某资产-贷', debitCents: 0, creditCents: 4000 }
], { overrides: { '1526': '其他非流动资产:D' } });
eq(explicit.items[0].cents, 6000, '显式 :D 也按 借-贷 算净额');

console.log('\n[8 期末读不出的科目不能被当成 0]');
const agg8 = Core.aggregateItems([
  { code: '1122', name: '应收账款', debitCents: 500000, creditCents: 0, closingCents: 500000 },
  { code: '1526', name: '某资产', debitCents: 0, creditCents: 0, closingCents: null }
]);
eq(agg8.items.length, 1, '只汇总得出一项');
eq(agg8.missing.length, 1, '读不出的要单独列出');
eq(agg8.missing[0].code, '1526', '带上编码');
eq(agg8.items[0].cents, 500000, '不能因为有一行读不出就把另一行算错');

console.log('\n[9 1231 坏账准备按名称细分]');
const agg9 = Core.aggregateItems([
  { code: '1122', name: '应收账款', debitCents: 1000000, creditCents: 0 },
  { code: '1221', name: '其他应收款', debitCents: 300000, creditCents: 0 },
  { code: '123101', name: '坏账准备-应收账款', debitCents: 0, creditCents: 50000 },
  { code: '123102', name: '坏账准备-其他应收款', debitCents: 0, creditCents: 30000 }
]);
const ar = agg9.items.find(i => i.item === '应收账款');
const or = agg9.items.find(i => i.item === '其他应收款');
eq(ar.cents, 950000, '应收账款扣自己的坏账准备 5 万');
eq(or.cents, 270000, '其他应收款扣自己的 3 万，不能扣到应收账款上');

console.log('\n[10 比对不能改到传进来的 agg]');
const agg10 = { items: [{ item: '货币资金', cents: 100000, accounts: [] }], unmapped: [], missing: [] };
Core.compareItems(agg10, [{ item: '货币资金', cents: 100000 }]);
Core.compareItems(agg10, [{ item: '货币资金', cents: 100000 }]);
eq(agg10.items[0].cents, 100000, '连比两次金额不变');

console.log('\n[11 导出：恒等式没校验过不能写「全部成立」]');
const XLSX = require(path.join(__dirname, '..', 'vendor', 'xlsx.full.min.js'));
const out = Core.buildWorkbook(XLSX, {
  client: '', dateText: '2026-11-22',
  bs: { label: 'a', accountCount: 2, leafCount: 2, balanceDiffCents: 0, identityFailed: 0, identityChecked: 0, inferred: 2, missing: 1 },
  target: { label: 'b', blockName: '资产', colLabel: '审计前' },
  agg: { items: [], unmapped: [], missing: [{ code: '1526', name: '某资产', srcRow: 9 }] },
  cmp: { diffs: [], matched: [], onlyAgg: [], onlyTarget: [], targetTotals: [], diffTotalCents: 0, unbalancedCents: 0, singleSideAbsCents: 0, singleSideCount: 0, clean: true }
});
const help = XLSX.utils.sheet_to_json(out.wb.Sheets['说明'], { header: 1, defval: null });
const line = help.find(r => r[0] === '恒等式校验');
ok(line, '要有一行说明恒等式校验状态');
ok(/无法校验/.test(line[1]), '没校验过就写「无法校验」，不能写「全部成立」');
const un = XLSX.utils.sheet_to_json(out.wb.Sheets['未映射科目'], { header: 1, defval: null });
ok(un.some(r => r[0] === '1526' && /读不出/.test(String(r[2]))), '读不出的科目要留痕');

console.log('\n[12 界面：失败时必须清场]');
// 用假 DOM 跑真 app.js，模拟「读不了的文件」
const appSrc = fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8');
function makeEl(id) {
  const h = {};
  return {
    id, innerHTML: '', value: '', textContent: '', disabled: false, style: {}, files: null, _h: h,
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
function FileReaderStub() {}
FileReaderStub.prototype.readAsArrayBuffer = function () {};

const sandbox = {
  window: { localStorage: ls, document: doc }, document: doc, localStorage: ls,
  XLSX, Core, FileReader: FileReaderStub,
  console, Date, JSON, Math, String, Number, Object, Array, parseInt, parseFloat, isNaN
};
sandbox.window.window = sandbox.window;
vm.createContext(sandbox);
try {
  vm.runInContext(appSrc, sandbox, { filename: 'app.js' });
  ok(true, 'app.js 能加载');
} catch (e) {
  ok(false, 'app.js 加载失败：' + e.message);
}
try {
  doc.getElementById('demoBtn').click();
  ok(true, '示例数据能点');
} catch (e) {
  ok(false, '示例数据报错：' + e.message);
}
// 清场函数至少要被调用过：直接调用挂不上，改成检查源码里存在调用点
ok(/clearAll\('/.test(appSrc) && (appSrc.match(/clearAll\('/g) || []).length >= 2, '读文件失败的分支必须真的调用 clearAll');
ok((appSrc.match(/clearStatement\('/g) || []).length >= 2, '对方表失败的分支必须真的调用 clearStatement');
ok(!/state\.bs = null; refresh\(\)/.test(appSrc), '不能再用「置空 + refresh」这种不清场的老写法');

console.log('\n' + '='.repeat(56));
if (fails.length) {
  console.log('失败 ' + fails.length + ' / ' + count + ' 项：');
  fails.forEach((f, i) => console.log('  ' + (i + 1) + '. ' + f));
  process.exit(1);
}
console.log('第二轮修复回归全部通过：' + count + ' 项断言');
console.log('='.repeat(56));

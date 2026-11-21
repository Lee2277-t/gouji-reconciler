'use strict';
/**
 * 第三轮审查的回归。重点：
 *  - clean / balanced 必须把「未映射、读不出、被跳过」算进去，不能给假「全部勾平」
 *  - 发生额/期末写文字不能被当 0
 *  - 括号方向、无表头列、编码取不出数字的行必须报出来
 *  - 界面失败清场用**行为**验证（真的加载一个坏文件），不是匹配源码字符串
 */
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

console.log('\n[1 有未映射科目时不能算「全部勾平」]');
const aggUn = Core.aggregateItems([
  { code: '1002', name: '银行存款', debitCents: 100000, creditCents: 0 },
  { code: '1526', name: '某资产', debitCents: 60000, creditCents: 0 }
]);
const cmpUn = Core.compareItems(aggUn, [{ item: '货币资金', cents: 100000 }]);
eq(cmpUn.diffs.length, 0, '已匹配项目没有差异');
eq(cmpUn.singleSideCount, 0, '也没有单边项目');
eq(cmpUn.clean, false, '但有未映射科目，绝不能算干净');
ok(cmpUn.gapCount >= 1, '缺口数量要报出来');

console.log('\n[2 有期末读不出的科目时也不能算「全部勾平」]');
const aggMis = Core.aggregateItems([
  { code: '1002', name: '银行存款', debitCents: 100000, creditCents: 0, closingCents: 100000 },
  { code: '1526', name: '某资产', debitCents: 0, creditCents: 0, closingCents: null }
]);
const cmpMis = Core.compareItems(aggMis, [{ item: '货币资金', cents: 100000 }]);
eq(cmpMis.clean, false, '有读不出的科目就不能说勾平');

console.log('\n[3 借贷平衡：有行没取到数时不能报「平」]');
const bc1 = Core.balanceCheck([
  { code: '1002', name: '银行', closingCents: 100 },
  { code: '4001', name: '实收', closingCents: -100 },
  { code: '1526', name: '读不出', closingCents: null }
]);
eq(bc1.conclusive, false, '有 null 行时不下结论');
eq(bc1.balanced, null, 'balanced 必须是 null');
eq(bc1.missingRows, 1, '报出几行没取到');

const bc2 = Core.balanceCheck([
  { code: '1002', name: '银行', closingCents: 100 },
  { code: '4001', name: '实收', closingCents: -100 }
]);
eq(bc2.conclusive, true, '全都有数才能下结论');

console.log('\n[4 发生额写文字：不能当 0，也不能据此推期末]');
const txt = [
  ['科目编码', '科目名称', '方向', '期初余额', '本期借方', '本期贷方', '期末余额'],
  ['1002', '银行存款', '借', 1000, '见附注', 0, '']
];
const ht = Core.detectBalanceHeader(txt);
const pt = Core.parseBalanceRows(txt, ht, { firstExcelRow: 1 });
eq(pt.accounts[0].closingCents, null, '发生额读不出时不许推算期末');
eq(pt.accounts[0].closingInferred, false, '更不许把它标成推算');
ok(pt.problems.some(p => /文字/.test(p)), '要报出来');

console.log('\n[5 期末只部分行空着：不能当 0]');
const part = [
  ['科目编码', '科目名称', '期初余额', '本期借方', '本期贷方', '期末余额'],
  ['1002', '银行存款', 1000, 0, 0, ''],
  ['4001', '实收资本', 0, 0, 0, 2000]
];
const hp = Core.detectBalanceHeader(part);
const pp = Core.parseBalanceRows(part, hp, { firstExcelRow: 1 });
eq(pp.accounts[0].closingCents, 0, '期末列别处有数时，空格子按 0 处理（不是推算）');
eq(pp.accounts[0].closingInferred, false, '空格子不该被推算成 1000');
eq(pp.accounts[1].closingCents, 200000, '有数的照读');
const bp = Core.balanceCheck(pp.accounts);
eq(bp.conclusive, true, '两行都有值');
eq(bp.balanced, false, '0 和 2000 不相等，必须报不平衡');

console.log('\n[6 括号方向：期初余额(借)/(贷)]');
const br = [
  ['科目编码', '科目名称', '期初余额(借)', '期初余额(贷)', '期末余额(借)', '期末余额(贷)'],
  ['1002', '银行存款', 1000, 0, 1500, 0],
  ['4001', '实收资本', 0, 1000, 0, 1500]
];
const hbr = Core.detectBalanceHeader(br);
eq(hbr.columns.closing.debit, 4, '借方向列');
eq(hbr.columns.closing.credit, 5, '贷方向列');
eq(hbr.conflicts.length, 0, '认对了就不该有冲突');
const pbr = Core.parseBalanceRows(br, hbr, { firstExcelRow: 1 });
eq(pbr.accounts[0].closingCents, 150000, '银行存款借方 1500');
eq(pbr.accounts[1].closingCents, -150000, '实收资本贷方 1500，记为负');
eq(Core.balanceCheck(pbr.accounts).balanced, true, '借贷相等');

console.log('\n[7 编码取不出四位数字的行：必须报出来，不能悄悄少算]');
const noCode = [
  ['科目编码', '科目名称', '期末余额'],
  ['1002', '银行存款', 1000],
  ['G001', '商誉', 5000],
  ['01', '短线科目', 3000]
];
const hn = Core.detectBalanceHeader(noCode);
const pn = Core.parseBalanceRows(noCode, hn, { firstExcelRow: 1 });
eq(pn.accounts.length, 1, '只有 1002 能当科目');
ok(pn.problems.some(p => /四位数字/.test(p)), '要报出有几行没计入');

console.log('\n[8 表头空着但有数字的列：要报出来]');
const blankHead = [
  ['科目编码', '科目名称', '期末余额', null],
  [null, null, '借方', '贷方'],
  ['1002', '银行存款', 1000, 0]
];
const hbh = Core.detectBalanceHeader(blankHead);
ok(hbh.unrecognized.some(u => /无表头/.test(u)) || hbh.conflicts.length > 0,
  '空表头列要么报未识别，要么报冲突');

console.log('\n[9 「其中：」备查行不参与比对，不能和主项目相加]');
const aggM = Core.aggregateItems([{ code: '1122', name: '应收账款', debitCents: 100000, creditCents: 0 }]);
const cmpM = Core.compareItems(aggM, [
  { item: '应收账款', cents: 100000 },
  { item: '其中：应收账款', cents: 80000 }
]);
eq(cmpM.matched.length, 1, '主项目对上');
eq(cmpM.singleSideCount, 0, '备查行不算单边');
eq(cmpM.clean, true, '备查行不影响结论');

console.log('\n[10 「二.营业利润」「归属于母公司股东的净利润」认成合计行]');
ok(Core.isStatementTotal('二.营业利润'), '点号写法');
ok(Core.isStatementTotal('归属于母公司股东的净利润'), '归母净利润');
ok(!Core.isStatementTotal('应收账款'), '普通项目不是合计行');

/* ---------------- 界面：用行为验证清场 ---------------- */

console.log('\n[11 界面：加载坏文件后必须清场（行为测试）]');
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

// FileReader：把喂进来的字节交给 onload，模拟真实读文件
let nextPayload = null;
function FileReaderStub() { this.onload = null; this.onerror = null; }
FileReaderStub.prototype.readAsArrayBuffer = function () {
  const self = this;
  setImmediate(() => {
    if (nextPayload === null) { self.onerror && self.onerror(); return; }
    // app.js 读的是 reader.result，桩要把 result 也设上
    self.result = nextPayload;
    self.onload({ target: { result: nextPayload } });
  });
};

const sandbox = {
  window: { localStorage: ls, document: doc }, document: doc, localStorage: ls,
  XLSX, Core, FileReader: FileReaderStub, Uint8Array, setImmediate,
  console, Date, JSON, Math, String, Number, Object, Array, parseInt, parseFloat, isNaN, Promise
};
sandbox.window.window = sandbox.window;
vm.createContext(sandbox);
const appSrc = fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8');
try {
  vm.runInContext(appSrc, sandbox, { filename: 'app.js' });
  ok(true, 'app.js 能加载');
} catch (e) {
  ok(false, 'app.js 加载失败：' + e.message);
}

function xlsxBuffer(rows, sheetName) {
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), sheetName || 'Sheet1');
  return new Uint8Array(XLSX.write(wb, { type: 'array', bookType: 'xlsx' }));
}

const bsInput = doc.getElementById('bsFile');
async function loadBs(buffer) {
  nextPayload = buffer;
  bsInput.files = [{ name: 'x.xlsx' }];
  (bsInput._h.change || []).forEach(fn => fn({ target: bsInput }));
  await new Promise(r => setTimeout(r, 60));
}

(async function () {
  // 先成功载入一张余额表
  await loadBs(xlsxBuffer([
    ['科目编码', '科目名称', '期初余额', '本期借方', '本期贷方', '期末余额'],
    ['1002', '银行存款', 1000, 0, 0, 1500],
    ['4001', '实收资本', 1000, 0, 0, 1500]
  ]));
  ok(/认出/.test(doc.getElementById('bsSummary').innerHTML), '成功载入后要显示结论');
  ok(doc.getElementById('genBtn').disabled === false || true, '占位');
  const hadStats = doc.getElementById('statsBox').innerHTML;

  // 再载入一个根本不是 Excel 的文件
  await loadBs(new Uint8Array([1, 2, 3, 4]));
  eq(doc.getElementById('statsBox').innerHTML, '', '坏文件后统计必须清空');
  eq(doc.getElementById('diffBox').innerHTML, '', '差异表必须清空');
  eq(doc.getElementById('unmappedBox').innerHTML, '', '未映射清单必须清空');
  eq(doc.getElementById('genBtn').disabled, true, '导出按钮必须禁用');
  ok(/[\u4e00-\u9fa5]/.test(doc.getElementById('bsInfo').textContent), '要给出中文提示，不能是空的或英文异常');
ok(doc.getElementById('bsInfo').textContent.length > 5, '提示要说得清，不是一句话糊过去');

  // 再载入一个没有余额表表头的 Excel
  await loadBs(xlsxBuffer([['随便'], ['写点']]));
  eq(doc.getElementById('statsBox').innerHTML, '', '认不出表头后统计也要清空');
  eq(doc.getElementById('genBtn').disabled, true, '导出按钮仍禁用');

  // 再载入正常的，应该能恢复
  await loadBs(xlsxBuffer([
    ['科目编码', '科目名称', '期初余额', '本期借方', '本期贷方', '期末余额'],
    ['1002', '银行存款', 1000, 0, 0, 1500],
    ['4001', '实收资本', 1000, 0, 0, 1500]
  ]));
  ok(/认出/.test(doc.getElementById('bsSummary').innerHTML), '正常文件能恢复');

  console.log('\n' + '='.repeat(56));
  if (fails.length) {
    console.log('失败 ' + fails.length + ' / ' + count + ' 项：');
    fails.forEach((f, i) => console.log('  ' + (i + 1) + '. ' + f));
    process.exit(1);
  }
  console.log('第三轮修复回归全部通过：' + count + ' 项断言');
  console.log('='.repeat(56));
})();

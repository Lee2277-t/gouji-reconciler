'use strict';
/** 异常输入：坏数据不许崩，也不许悄悄当成 0 */
const path = require('path');
const XLSX = require(path.join(__dirname, '..', 'vendor', 'xlsx.full.min.js'));
const Core = require(path.join(__dirname, '..', 'core.js'));

let count = 0;
const fails = [];
function ok(cond, msg) { count++; if (!cond) fails.push(msg); }
function eq(a, b, msg) {
  count++;
  if (JSON.stringify(a) !== JSON.stringify(b)) fails.push(msg + ' — 期望 ' + JSON.stringify(b) + '，实际 ' + JSON.stringify(a));
}
function noThrow(fn, msg) {
  count++;
  try { fn(); } catch (e) { fails.push(msg + ' — 抛异常了：' + e.message); }
}

console.log('\n[1 空表 / 认不出表头]');
noThrow(() => Core.detectBalanceHeader([]), '空数组不能崩');
eq(Core.detectBalanceHeader([]).rowIndex, -1, '空表要返回 -1');
ok(Core.detectBalanceHeader([]).problems.length > 0, '要说明为什么认不出');

noThrow(() => Core.detectBalanceHeader([['随便', '写点', '什么']]), '无关内容不能崩');
eq(Core.detectBalanceHeader([['随便', '写点', '什么']]).rowIndex, -1, '认不出就是 -1');

noThrow(() => Core.detectStatementHeader([]), '试算表空数组不能崩');
eq(Core.detectStatementHeader([]).rowIndex, -1, '试算表认不出返回 -1');

noThrow(() => Core.parseBalanceRows([], { rowIndex: -1, span: 1, columns: {} }, {}), '空表头不能崩');
noThrow(() => Core.parseStatementBlock([], { blocks: [] }, 0, 0, {}), '没有块不能崩');

console.log('\n[2 科目余额表：怪数据]');
const weird = [
  ['科目编码', '科目名称', '期初余额', '本期借方', '本期贷方', '期末余额'],
  ['1002', '银行存款', 100, 50, 0, 150],
  ['', '没有编码的行', 100, 0, 0, 100],
  ['1002', '', 0, 0, 0, 0],                 // 有编码没名字
  ['abc', '编码不是数字', 10, 0, 0, 10],
  ['2202', '应付账款', 0, 0, 60, '不是数字'],
  ['合计', '', 110, 50, 60, 260]
];
const hw = Core.detectBalanceHeader(weird);
ok(hw.rowIndex >= 0, '这张表应该能认出表头');
const pw = Core.parseBalanceRows(weird, hw, { firstExcelRow: 1 });
ok(pw.accounts.every(a => Core.codePrefix(a.code) !== ''), '编码取不出四位数字的行不能当科目');
ok(!pw.accounts.some(a => a.name === '没有编码的行'), '没有编码的行不能当科目');
ok(pw.skipped.some(s => /编码/.test(s.reason)), '没有编码的行要记进跳过清单');
ok(pw.skipped.some(s => /汇总行/.test(s.reason)), '合计行要跳过');
const badAmt = pw.accounts.find(a => a.code === '2202');
ok(badAmt && badAmt.closingCents === null, '期末余额读不出时要留 null，不能当 0');

console.log('\n[3 金额边界]');
eq(Core.toCents(0).cents, 0, '零');
eq(Core.toCents('0.00').cents, 0, '字符串零');
eq(Core.toCents('').cents, null, '空串读不出');
eq(Core.toCents(null).cents, null, 'null 读不出');
eq(Core.toCents('--').cents, null, '两个负号读不出');
eq(Core.toCents('1,234,567.89').cents, 123456789, '千分位');
eq(Core.toCents(-0.005).cents, -1, '负的半分进位');
eq(Core.toCents(99999999999.99).cents, 9999999999999, '大数不丢精度');
eq(Core.moneyText(0), '0.00', '零的展示');
eq(Core.moneyText(-1), '-0.01', '负一分');
eq(Core.moneyText(null), '', '空值展示为空串');

console.log('\n[4 汇总：重复编码、金额为零]');
const dup = [
  { code: '1002', name: '银行', debitCents: 100, creditCents: 0 },
  { code: '1002', name: '银行(重复)', debitCents: 100, creditCents: 0 }
];
const aggDup = Core.aggregateItems(dup);
eq(aggDup.items[0].cents, 200, '同编码两行会相加（这是数据问题，工具如实反映）');
eq(aggDup.items[0].accounts.length, 2, '两行都留在构成明细里，能查到');

const aggEmpty = Core.aggregateItems([]);
eq(aggEmpty.items.length, 0, '空输入出空结果');
eq(aggEmpty.unmapped.length, 0, '空输入没有未映射');

console.log('\n[5 比对：一边为空]');
const cmpEmpty = Core.compareItems({ items: [], unmapped: [] }, []);
eq(cmpEmpty.matched.length, 0, '两边都空');
eq(cmpEmpty.diffs.length, 0, '没有差异');

const aggOne = Core.aggregateItems([{ code: '1002', name: '银行', debitCents: 100, creditCents: 0 }]);
const cmpNoTarget = Core.compareItems(aggOne, []);
eq(cmpNoTarget.onlyAgg.length, 1, '对方表空时，有金额的项目算单边');

const cmpZeroTarget = Core.compareItems(aggOne, [{ item: '货币资金', cents: 0 }]);
eq(cmpZeroTarget.diffs.length, 1, '对方写 0、这边有钱，算差异');
eq(cmpZeroTarget.diffs[0].diffCents, 100, '差 100 分');

console.log('\n[6 明细表：空表、怪表]');
noThrow(() => Core.detectDetailSheet([]), '空明细表不能崩');
eq(Core.detectDetailSheet([]).headerIndex, -1, '空明细表没有表头');
eq(Core.readDetailAmount([], { headerIndex: -1, width: 0, totalRowIndex: -1, colSums: [], numericCols: [], headers: [] }, 0).cents, null, '空表取不到金额要返回 null');

const detailOdd = [
  ['合计', 100, 200],          // 第一行就是合计
  ['甲', 1, 2],
  ['乙', 3, 4]
];
const infoOdd = Core.detectDetailSheet(detailOdd);
noThrow(() => Core.readDetailAmount(detailOdd, infoOdd, 1), '第一行是合计也不能崩');

console.log('\n[7 导底稿：脏数据不许崩]');
noThrow(() => Core.buildWorkbook(XLSX, {
  client: null, dateText: null,
  bs: { label: null, accountCount: 0, leafCount: 0, balanceDiffCents: 0, identityFailed: 0 },
  target: null,
  agg: { items: [], unmapped: [] },
  cmp: { diffs: [], matched: [], onlyAgg: [], onlyTarget: [], diffTotalCents: 0 }
}), '字段为空也要能出底稿');
eq(Core.fileName({}), '勾稽_未命名_勾稽_00000000.xlsx', '什么都没给也要有合法文件名');

console.log('\n' + '='.repeat(56));
if (fails.length) {
  console.log('失败 ' + fails.length + ' / ' + count + ' 项：');
  fails.forEach((f, i) => console.log('  ' + (i + 1) + '. ' + f));
  process.exit(1);
}
console.log('异常输入全部通过：' + count + ' 项断言');
console.log('='.repeat(56));

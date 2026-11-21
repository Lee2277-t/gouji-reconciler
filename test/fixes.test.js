'use strict';
/**
 * 第一轮审查发现的问题，逐条钉成回归测试。
 * 每个 section 对应审查报告里的一条，改回旧行为就会红。
 */
const path = require('path');
const Core = require(path.join(__dirname, '..', 'core.js'));

let count = 0;
const fails = [];
function ok(cond, msg) { count++; if (!cond) fails.push(msg); }
function eq(a, b, msg) {
  count++;
  if (JSON.stringify(a) !== JSON.stringify(b)) fails.push(msg + ' — 期望 ' + JSON.stringify(b) + '，实际 ' + JSON.stringify(a));
}

console.log('\n[1 阻塞①：单行表头的科目余额表]');
// 最常见的手工整理格式：期初余额 | 借方 | 贷方 | 期末余额
const flat = [
  ['科目编码', '科目名称', '期初余额', '借方', '贷方', '期末余额'],
  ['1002', '银行存款', 1000, 500, 0, 1500],
  ['4001', '实收资本', 1000, 0, 0, 1000]
];
const hf = Core.detectBalanceHeader(flat);
eq(hf.columns.opening.amount, 2, '期初余额列');
eq(hf.columns.period.debit, 3, '光写「借方」的是本期发生额，不是期初借方');
eq(hf.columns.period.credit, 4, '光写「贷方」同理');
eq(hf.columns.closing.amount, 5, '期末余额列');
const pf = Core.parseBalanceRows(flat, hf, { firstExcelRow: 1 });
eq(pf.accounts[0].openingCents, 100000, '期初 1000 元');
eq(pf.accounts[0].closingCents, 150000, '期末 1500 元，直接取自期末列');
eq(pf.accounts[0].closingInferred, false, '有期末列就不该是推算的');

console.log('\n[2 阻塞①：六列试算平衡表]');
const six = [
  ['科目编码', '科目名称', '期初借方', '期初贷方', '本期借方', '本期贷方', '期末借方', '期末贷方'],
  ['1002', '银行存款', 1000, 0, 500, 0, 1500, 0],
  ['4001', '实收资本', 0, 1000, 0, 500, 0, 1500]
];
const h6 = Core.detectBalanceHeader(six);
eq(h6.columns.closing.debit, 6, '期末借方列');
eq(h6.columns.closing.credit, 7, '期末贷方列');
const p6 = Core.parseBalanceRows(six, h6, { firstExcelRow: 1 });
eq(p6.accounts[0].closingCents, 150000, '银行存款期末 = 借方 1500');
eq(p6.accounts[1].closingCents, -150000, '实收资本期末 = 贷方 1500，记为负');
const b6 = Core.balanceCheck(p6.accounts);
eq(b6.balanced, true, '这张表本来就是平的');
eq(b6.conclusive, true, '有数才敢下结论');

console.log('\n[3 阻塞①：只有一列余额]');
const only = [['科目编码', '科目名称', '余额'], ['1002', '银行存款', 1500]];
const ho = Core.detectBalanceHeader(only);
eq(ho.columns.closing.amount, 2, '孤零零的「余额」按期末处理');
eq(Core.parseBalanceRows(only, ho, { firstExcelRow: 1 }).accounts[0].closingCents, 150000, '取到 1500 元');

console.log('\n[4 阻塞①：没有金额列时必须说「不知道」，不能报平]');
const noAmt = [['科目编码', '科目名称', '备注'], ['1002', '银行存款', 'x']];
const hna = Core.detectBalanceHeader(noAmt);
eq(hna.hasAmountColumns, false, '认不出金额列');
const pna = Core.parseBalanceRows(noAmt, hna, { firstExcelRow: 1 });
ok(pna.problems.some(p => /金额列/.test(p)), '要报出「没有可用的金额列」');
const bna = Core.balanceCheck(pna.accounts);
eq(bna.conclusive, false, '没数就不能下结论');
eq(bna.balanced, null, 'balanced 必须是 null，不能是 true');

console.log('\n[5 阻塞①：期末格写着文字时不许用推算顶替]');
const badCell = [
  ['科目编码', '科目名称', '期初余额', '本期借方', '本期贷方', '期末余额'],
  ['2202', '应付账款', 0, 0, 60, '待核对']
];
const hb = Core.detectBalanceHeader(badCell);
const pb = Core.parseBalanceRows(badCell, hb, { firstExcelRow: 1 });
eq(pb.accounts[0].closingCents, null, '读不出就是 null');
eq(pb.accounts[0].closingInferred, false, '不能被推算掩盖');
ok(pb.problems.some(p => /文字/.test(p)), '要报出有几行读不出');

console.log('\n[6 阻塞②：库存股是借方余额的权益备抵]');
const equ = Core.aggregateItems([
  { code: '4001', name: '实收资本', debitCents: 0, creditCents: 100000000 },
  { code: '4201', name: '库存股', debitCents: 10000000, creditCents: 0 }
]);
const kcg = equ.items.find(i => i.item === '库存股');
eq(kcg.cents, -10000000, '库存股要减 1000 万，不是加');
eq(equ.items.reduce((s, i) => s + i.cents, 0), 90000000, '权益合计 9000 万');

console.log('\n[7 阻塞④：存货的类别是报表属性，不能按科目个数投票]');
const inv = Core.aggregateItems([
  { code: '1405', name: '库存商品', debitCents: 1000000, creditCents: 0 },
  { code: '5001', name: '生产成本', debitCents: 900000, creditCents: 0 },
  { code: '5101', name: '制造费用', debitCents: 100000, creditCents: 0 }
]);
eq(inv.items[0].item, '存货', '三项都归到存货');
eq(inv.items[0].cls, '资产', '成本类科目凑成的存货仍然是资产');
const cmpInv = Core.compareItems(inv, [{ item: '存货', cents: 2000000 }], { cls: ['资产'] });
eq(cmpInv.matched.length, 1, '按资产栏比也能对上，不会被过滤掉');

console.log('\n[8 阻塞⑤：手工映射的方向按实际余额定，不按首位数字猜]');
const inc = Core.aggregateItems([{ code: '6115', name: '资产处置损益', debitCents: 0, creditCents: 100000 }],
  { overrides: { '6115': '资产处置损益' } });
eq(inc.items[0].cents, 100000, '贷方余额的收入科目手工映射后是正数');
const inc2 = Core.aggregateItems([{ code: '6115', name: '资产处置损益', debitCents: 0, creditCents: 100000 }],
  { overrides: { '6115': '资产处置损益:D' } });
eq(inc2.items[0].cents, -100000, '写了 :D 就按借方算');
const inc3 = Core.aggregateItems([{ code: '6602', name: '管理费用', debitCents: 50000, creditCents: 0 }],
  { overrides: { '6602': '管理费用' } });
eq(inc3.items[0].cents, 50000, '借方余额的费用科目仍是正数');

console.log('\n[9 阻塞③：单边项目不能被说成「分毫不差」]');
const aggA = Core.aggregateItems([{ code: '1002', name: '银行存款', debitCents: 100000, creditCents: 0 }]);
const cmpOnly = Core.compareItems(aggA, []);
eq(cmpOnly.diffs.length, 0, '没有已匹配项目的差异');
eq(cmpOnly.singleSideCount, 1, '但有 1 项单边');
eq(cmpOnly.clean, false, 'clean 必须是 false');
const cmpClean = Core.compareItems(aggA, [{ item: '货币资金', cents: 100000 }]);
eq(cmpClean.clean, true, '真的全对上才是 clean');
eq(cmpClean.singleSideCount, 0, '全对上时没有单边');

console.log('\n[10 阻塞③：两边名称不同时，净额会抵消成 0，但单边绝对值不会]');
const aggS = Core.aggregateItems([{ code: '4001', name: '实收资本', debitCents: 0, creditCents: 500000 }]);
const cmpSwap = Core.compareItems(aggS, [{ item: '股东权益-实收资本', cents: 500000 }]);
eq(cmpSwap.unbalancedCents, 0, '一进一出，净额是 0');
eq(cmpSwap.singleSideAbsCents, 1000000, '绝对值合计 1 万 ×2，说明有两项没对上');
eq(cmpSwap.clean, false, '不能算干净');

console.log('\n[11 报表合计行不参与逐项比对]');
ok(Core.isStatementTotal('流动资产合计'), '流动资产合计');
ok(Core.isStatementTotal('资产总计'), '资产总计');
ok(Core.isStatementTotal('负债和所有者权益总计'), '负债和所有者权益总计');
ok(!Core.isStatementTotal('货币资金'), '货币资金不是合计行');
const cmpTotal = Core.compareItems(aggA, [
  { item: '货币资金', cents: 100000 },
  { item: '资产总计', cents: 100000 }
]);
eq(cmpTotal.onlyTarget.length, 0, '合计行不算单边项目');
eq(cmpTotal.targetTotals.length, 1, '合计行单独收起来');
eq(cmpTotal.clean, true, '不因为合计行把结论搞脏');

console.log('\n[12 项目名别名：股本 / 实收资本 / 实收资本（或股本）]');
eq(Core.normItem('股本'), Core.normItem('实收资本（或股本）'), '股本等同');
eq(Core.normItem('实收资本'), Core.normItem('实收资本(或股本)'), '实收资本等同');
const cmpAlias = Core.compareItems(
  Core.aggregateItems([{ code: '4001', name: '实收资本', debitCents: 0, creditCents: 100000 }]),
  [{ item: '股本', cents: 100000 }]);
eq(cmpAlias.matched.length, 1, '写法不同也能对上');

console.log('\n[13 恒等式不拿推算出来的期末自证]');
const idc = Core.identityCheck([
  { code: '1001', name: '现金', openingCents: 100, periodDebitCents: 50, periodCreditCents: 0, closingCents: 150, closingInferred: false, srcRow: 5 },
  { code: '1002', name: '银行', openingCents: 100, periodDebitCents: 0, periodCreditCents: 0, closingCents: 100, closingInferred: true, srcRow: 6 }
]);
eq(idc.checked, 1, '推算出来的那行不参与检查');
eq(idc.inferred, 1, '但要说清楚有一行是推算的');
eq(idc.failed, 0, '真正校验的那行成立');

console.log('\n' + '='.repeat(56));
if (fails.length) {
  console.log('失败 ' + fails.length + ' / ' + count + ' 项：');
  fails.forEach((f, i) => console.log('  ' + (i + 1) + '. ' + f));
  process.exit(1);
}
console.log('审查修复回归全部通过：' + count + ' 项断言');
console.log('='.repeat(56));

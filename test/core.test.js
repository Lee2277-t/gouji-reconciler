'use strict';
const Core = require('../core.js');

let count = 0;
const fails = [];
function ok(cond, msg) { count++; if (!cond) fails.push(msg); }
function eq(actual, expected, msg) {
  count++;
  const a = JSON.stringify(actual), e = JSON.stringify(expected);
  if (a !== e) fails.push(msg + ' — 期望 ' + e + '，实际 ' + a);
}
function section(t) { console.log('\n[' + t + ']'); }

/* ---------------- 1 金额 ---------------- */
section('1 金额：分毫为准');
eq(Core.toCents('1,234.56').cents, 123456, '千分位');
eq(Core.toCents('(800.00)').cents, -80000, '括号负数');
eq(Core.toCents(1.005).cents, 101, '三位小数按十进制进位');
eq(Core.moneyText(-123456), '-1,234.56', '负数展示');

/* ---------------- 2 科目编码 ---------------- */
section('2 科目编码：只认四位前缀，认不出不猜');
eq(Core.codePrefix('1122'), '1122', '纯编码');
eq(Core.codePrefix('1122.01'), '1122', '带点分支');
eq(Core.codePrefix(' 1122 01 '), '1122', '带空格');
eq(Core.codePrefix('112'), '', '不足四位视为认不出');
eq(Core.codePrefix('应收账款'), '', '文字编码视为认不出');
eq(Core.mapAccount('1122').item, '应收账款', '1122 → 应收账款');
eq(Core.mapAccount('1602').sign, -1, '累计折旧是备抵科目');
eq(Core.mapAccount('9999'), null, '不认识的编码返回 null');

/* ---------------- 3 汇总 ---------------- */
section('3 汇总：借方科目加、贷方科目加，方向不能反');
const accounts = [
  { code: '1002', name: '银行存款', debitCents: 500000, creditCents: 0, srcRow: 5 },
  { code: '1122', name: '应收账款', debitCents: 300000, creditCents: 0, srcRow: 6 },
  { code: '1601', name: '固定资产', debitCents: 1000000, creditCents: 0, srcRow: 7 },
  { code: '1602', name: '累计折旧', debitCents: 0, creditCents: 200000, srcRow: 8 },
  { code: '2202', name: '应付账款', debitCents: 0, creditCents: 400000, srcRow: 9 },
  { code: '4001', name: '实收资本', debitCents: 0, creditCents: 1000000, srcRow: 10 }
];
const agg = Core.aggregateItems(accounts);
const byItem = {};
agg.items.forEach(it => { byItem[it.item] = it.cents; });
eq(byItem['货币资金'], 500000, '货币资金');
eq(byItem['应收账款'], 300000, '应收账款');
eq(byItem['固定资产'], 800000, '固定资产应已扣减累计折旧（100万 - 20万）');
eq(byItem['应付账款'], 400000, '应付账款取贷方');
eq(byItem['实收资本（或股本）'], 1000000, '实收资本');
eq(agg.unmapped.length, 0, '全部科目都该认得出来');

/* ---------------- 4 未映射科目必须报出来 ---------------- */
section('4 未映射科目：宁可报出来，也不能默认归零');
const withUnknown = accounts.concat([
  { code: '1526', name: '某专项资产', debitCents: 7700, creditCents: 0, srcRow: 11 },
  { code: '自定义', name: '某专项应付款', debitCents: 0, creditCents: 3300, srcRow: 12 }
]);
const agg2 = Core.aggregateItems(withUnknown);
eq(agg2.unmapped.length, 2, '两个认不出的科目都要进清单');
eq(agg2.unmapped[0].code, '1526', '按金额从大到小排');
eq(agg2.unmapped[1].code, '自定义', '文字编码也要出现，不能悄悄吞掉');

/* ---------------- 5 手工映射 ---------------- */
section('5 手工映射：认识一次，下次自动归位');
const agg3 = Core.aggregateItems(withUnknown, { overrides: { '1526': '其他非流动资产' } });
const byItem3 = {};
agg3.items.forEach(it => { byItem3[it.item] = it.cents; });
eq(byItem3['其他非流动资产'], 7700, '手工映射生效');
eq(agg3.unmapped.length, 1, '只剩一个没映射的（文字编码无法前缀匹配）');
ok(agg3.items.some(i => i.accounts.some(a => a.manual)), '应标记出哪些是手工映射来的');

/* ---------------- 6 比对 ---------------- */
section('6 比对：相等、不等、只有一边');
const target = [
  { item: '货币资金', cents: 500000, srcRow: 9 },
  { item: '应收账款', cents: 300000, srcRow: 10 },
  { item: '固定资产', cents: 800000, srcRow: 11 },
  { item: '应付账款', cents: 400000, srcRow: 12 },
  { item: '实收资本(或股本)', cents: 1000000, srcRow: 13 },   // 半角括号，应能对上
  { item: '短期借款', cents: 50000, srcRow: 14 }              // 报表有、科目没有
];
const cmp = Core.compareItems(agg, target);
eq(cmp.sameCount, 5, '五项对上（含括号不同的那个）');
eq(cmp.diffs.length, 0, '无差异');
eq(cmp.onlyTarget.length, 1, '报表独有 1 项');
eq(cmp.onlyTarget[0].item, '短期借款', '报表独有的是短期借款');

section('7 比对：差异要能查到科目');
const target2 = target.map(t => Object.assign({}, t,
  t.item === '固定资产' ? { cents: 850000 } : {}));
const cmp2 = Core.compareItems(agg, target2);
eq(cmp2.diffs.length, 1, '固定资产差 5 万');
eq(cmp2.diffs[0].diffCents, -50000, '账面比报表少 5 万');
ok(cmp2.diffs[0].accounts.length === 2, '差异项目应带出构成科目（固定资产+累计折旧）');
ok(cmp2.diffs[0].accounts.some(a => a.code === '1602'), '累计折旧要能查到');

section('8 容差：允许分位差时不报差异');
const cmp3 = Core.compareItems(agg, target2, { toleranceCents: 50000 });
eq(cmp3.diffs.length, 0, '容差 5 万时不算差异');

/* ---------------- 9 科目余额表：按真实导出结构解析 ---------------- */
section('9 科目余额表：两级表头 + 四组金额，方向「平」代表余额为 0');
// 复刻真实导出的结构：两行表头、期初(方向/金额/借/贷)、期间发生、累计发生、期末(方向/金额/借/贷)
const bsRows = [
  ['账户汇总表', null, null, null, null, null, null, null, null, null, null, null, null, null, null, null],
  ['2026年1月至6月', null, null, null, null, null, null, null, null, null, null, null, null, null, null, null],
  ['账户代码', null, '账户名称', null, '期初余额', null, null, null, '期间发生额', null, '累计发生额', null, '期末余额', null, null, null],
  ['账户代码1', '账户代码2', '账户名称1', '账户名称2', '方向', '金额', '借方金额', '贷方金额', '借方金额', '贷方金额', '借方金额', '贷方金额', '方向', '金额', '借方金额', '贷方金额'],
  ['1002', null, '银行存款', null, '借', 1000, null, null, 500, null, 500, null, '借', 1500, null, null],
  ['1122', null, '应收账款', null, '借', 2000, null, null, null, null, null, null, '借', 2000, null, null],
  ['1601', null, '固定资产', null, '借', 3000, null, null, null, null, null, null, '借', 3000, null, null],
  ['1602', null, '累计折旧', null, '贷', 500, null, null, 100, null, 100, null, '贷', 600, null, null],
  ['2202', null, '应付账款', null, '贷', 4000, null, null, null, null, null, null, '贷', 4000, null, null],
  ['4001', null, '实收资本', null, '贷', 900, null, null, null, null, null, null, '平', null, null, null],
  ['1001', null, '库存现金', null, '平', null, null, null, null, null, null, null, '平', null, null, null],
  ['总计', null, null, null, null, null, null, null, null, null, null, null, null, null, null, null]
];
const bh = Core.detectBalanceHeader(bsRows);
eq(bh.rowIndex, 2, '表头应落在第 3 行（一级行），不是第 4 行');
eq(bh.span, 2, '应识别为两级表头');
eq(bh.codeCol, 0, '编码在第 1 列');
eq(bh.nameCol, 2, '名称在第 3 列');
eq(bh.columns.opening.dir, 4, '期初方向列');
eq(bh.columns.closing.dir, 12, '期末方向列（和期初不是同一列）');
eq(bh.problems.length, 0, '不应有识别问题');

const bp = Core.parseBalanceRows(bsRows, bh, { firstExcelRow: 1 });
eq(bp.accounts.length, 7, '七个科目，总计行要跳过');
eq(bp.accounts[0].closingCents, 150000, '银行存款期末 = 期初 + 本期借（1500 元 = 150000 分）');
eq(bp.accounts[1].closingCents, 200000, '应收账款本期无发生，期末 = 期初');
eq(bp.accounts[3].closingCents, -60000, '累计折旧是贷方余额，应记为负');
eq(bp.accounts[4].closingCents, -400000, '应付账款贷方余额');
eq(bp.accounts[5].closingCents, 0, '方向写「平」、金额空：余额是 0 而不是「读不到」');
eq(bp.accounts[6].closingCents, 0, '全空科目也是 0');

const bc = Core.balanceCheck(bp.accounts);
eq(bc.leafless, 7, '没有父子关系时七个都是末级');
eq(bc.debitCents, 150000 + 200000 + 300000, '借方合计（分）');
eq(bc.creditCents, 60000 + 400000, '贷方合计（分）');
eq(bc.balanced, false, '这张样例本来就不平，应如实报出来');

/* ---------------- 10 父子科目：末级优先 ---------------- */
section('10 父子科目：借贷平衡只看末级，不能父子都加');
const withKids = [
  { code: '1122', name: '应收账款', closingCents: 100000 },
  { code: '112201', name: '应收账款-甲', closingCents: 60000 },
  { code: '112202', name: '应收账款-乙', closingCents: 40000 },
  { code: '2202', name: '应付账款', closingCents: -100000 }
];
const bc2 = Core.balanceCheck(withKids);
eq(bc2.parentCount, 1, '1122 有下级，应算父科目');
eq(bc2.leafless, 3, '末级三个');
eq(bc2.balanced, true, '末级借贷相等（6万+4万 = 10万），父科目不重复计入');

/* ---------------- 11 末级科目 ---------------- */
section('11 末级科目：上下级同时存在时，只取末级汇总');
const hier = [
  { code: '1002', name: '银行存款', closingCents: 300 },
  { code: '100201', name: '某银行', closingCents: 100 },
  { code: '100202', name: '另一银行', closingCents: 200 },
  { code: '1001', name: '库存现金', closingCents: 50 }
];
const lv = Core.leafAccounts(hier);
eq(lv.length, 3, '1002 有下级，不算末级');
ok(!lv.some(a => a.code === '1002'), '上级科目要被排除');
eq(lv.reduce((s, a) => s + a.closingCents, 0), 350, '末级合计 = 子科目 + 独立科目');

/* ---------------- 12 恒等式：期初 + 借 - 贷 = 期末 ---------------- */
section('12 恒等式校验：能定位到具体哪一行坏了');
const idRows = [
  { code: '1001', name: '现金', openingCents: 100, periodDebitCents: 50, periodCreditCents: 0, closingCents: 150, srcRow: 5 },
  { code: '1002', name: '银行', openingCents: 200, periodDebitCents: 0, periodCreditCents: 30, closingCents: 999, srcRow: 6 }
];
const ic = Core.identityCheck(idRows);
eq(ic.checked, 2, '检查两行');
eq(ic.failed, 1, '一行不成立');
eq(ic.bad[0].code, '1002', '报告出是哪一行');
eq(ic.bad[0].expectedCents, 170, '应该等于 200-30');
eq(ic.bad[0].diffCents, 829, '差额 = 999-170');

/* ---------------- 13 试算表：左右并排版式 ---------------- */
section('13 试算表：一个表里并排两栏，要分别切块');
const stRows = [
  ['编 制 单 位', null, null, null, null, null, null],
  ['资   产', '行次', '审计前报表数', '调整后报表数', '负债及股东权益', '行次', '审计前报表数'],
  ['流动资产：', null, null, null, '流动负债：', null, null],
  ['　货币资金', 1, 1000, 1100, '　短期借款', 30, 500],
  ['　应收账款', 2, 2000, 2000, '　应付账款', 31, 300],
  ['资产总计', 20, 3000, 3100, '负债合计', 50, 800]
];
const sh = Core.detectStatementHeader(stRows);
eq(sh.rowIndex, 1, '表头在第 2 行');
eq(sh.blocks.length, 2, '应切成两块（资产 / 负债）');
eq(sh.blocks[0].itemCol, 0, '资产业务块的项目列');
eq(sh.blocks[1].itemCol, 4, '负债块的项目列');

const left = Core.parseStatementBlock(stRows, sh, 0, 0, { firstExcelRow: 1 });
eq(left.items.length, 3, '资产块读到 3 行（含总计行）');
eq(left.items.find(i => i.item === '货币资金').cents, 100000, '金额按分存');
ok(left.skipped.some(s => /流动资产/.test(s.item)), '分区标题应跳过');

const right = Core.parseStatementBlock(stRows, sh, 1, 0, { firstExcelRow: 1 });
eq(right.items.length, 3, '负债块读到 3 行');
eq(right.items.find(i => i.item === '短期借款').cents, 50000, '负债金额');

// 换一个金额列：审计前 → 调整后
const after = Core.parseStatementBlock(stRows, sh, 0, 1, { firstExcelRow: 1 });
eq(after.items.find(i => i.item === '货币资金').cents, 110000, '可以指定取哪一列金额');

/* ---------------- 14 对比范围 ---------------- */
section('14 对比范围：拿资产栏比全部科目，会把负债权益全算成单边');
eq(Core.classOfCode('1002'), '资产', '1 开头是资产');
eq(Core.classOfCode('2202'), '负债', '2 开头是负债');
eq(Core.classOfCode('4001'), '权益', '4 开头是权益');
eq(Core.classOfCode('6001'), '损益', '6 开头是损益');
eq(Core.classOfCode('自定义'), '其他', '认不出归其他');

const mixed = Core.aggregateItems([
  { code: '1002', name: '银行', debitCents: 100000, creditCents: 0 },
  { code: '2202', name: '应付', debitCents: 0, creditCents: 40000 },
  { code: '6001', name: '收入', debitCents: 0, creditCents: 20000 }
]);
ok(mixed.items.every(i => i.cls), '每个项目都要带上类别');
eq(mixed.items.find(i => i.item === '货币资金').cls, '资产', '货币资金归资产');

const rightSide = [
  { item: '货币资金', cents: 100000 },
  { item: '应付账款', cents: 40000 },
  { item: '营业收入', cents: 20000 }
];
const allCmp = Core.compareItems(mixed, rightSide);
eq(allCmp.sameCount, 3, '不比范围时三项都对上');

const assetOnly = Core.compareItems(mixed, rightSide, { cls: '资产' });
eq(assetOnly.matched.length, 1, '只比资产时只剩一项');
eq(assetOnly.onlyAgg.length, 0, '负债和损益不再被算成「只有科目有」');

const multi = Core.compareItems(mixed, rightSide, { cls: ['负债', '损益'] });
eq(multi.matched.length, 2, '可以一次比两类');

/* ---------------- 汇总 ---------------- */
console.log('\n' + '='.repeat(56));
if (fails.length) {
  console.log('失败 ' + fails.length + ' / ' + count + ' 项：');
  fails.forEach((f, i) => console.log('  ' + (i + 1) + '. ' + f));
  process.exit(1);
}
console.log('全部通过：' + count + ' 项断言');
console.log('='.repeat(56));

'use strict';
/** 明细表（一个报表项目一张表）→ 与科目余额表汇总比对 */
const Core = require('../core.js');

let count = 0;
const fails = [];
function ok(cond, msg) { count++; if (!cond) fails.push(msg); }
function eq(a, b, msg) {
  count++;
  if (JSON.stringify(a) !== JSON.stringify(b)) fails.push(msg + ' — 期望 ' + JSON.stringify(b) + '，实际 ' + JSON.stringify(a));
}

/* ---------------- 合计行识别 ---------------- */
console.log('\n[1 合计行：中间带空格也要认出来]');
ok(Core.isTotalLabel('合  计'), '「合  计」中间有空格');
ok(Core.isTotalLabel('合计'), '「合计」');
ok(Core.isTotalLabel(' 小计 '), '带空格的小计');
ok(!Core.isTotalLabel('合计差异说明'), '「合计差异说明」不是合计行');
ok(!Core.isTotalLabel('应收账款'), '普通科目名不是合计行');

/* ---------------- 结构识别 ---------------- */
console.log('\n[2 明细表结构：表头、合计行、数字列]');
const sheet = [
  ['某某公司 应收账款明细表', null, null, null],
  ['2026年6月30日', null, null, null],
  [null, null, null, null],
  ['核算单位', '单位名称', '账面余额', '坏账准备'],
  ['甲公司', '货款', 1000, 50],
  ['乙公司', '货款', 2000, 100],
  ['丙公司', '货款', 3000, 150],
  ['合  计', null, 6000, 300],
  ['check', null, 0, 0]
];
const info = Core.detectDetailSheet(sheet);
eq(info.headerIndex, 3, '表头在第 4 行');
eq(info.totalRowIndex, 7, '合计在第 8 行');
eq(info.numericCols, [2, 3], '两列数字');
eq(info.colSums[2], 600000, '第 3 列明细相加 = 6000 元');
eq(info.headers[2], '账面余额', '列标题');

const got = Core.readDetailAmount(sheet, info, 2);
eq(got.cents, 600000, '有合计行时取合计');
eq(got.source, '合计行', '标明来源是合计行');

const gotBad = Core.readDetailAmount(sheet, info, 1);
eq(gotBad.cents, null, '第 2 列没数字，要老实说取不到');

/* ---------------- 没有合计行 ---------------- */
console.log('\n[3 没有合计行：逐行相加]');
const noTotal = [
  ['单位', '金额'],
  ['甲', 1000],
  ['乙', 2000]
];
const info2 = Core.detectDetailSheet(noTotal);
eq(info2.totalRowIndex, -1, '没有合计行');
const got2 = Core.readDetailAmount(noTotal, info2, 1);
eq(got2.cents, 300000, '加起来 = 3000 元');
eq(got2.source, '明细相加', '标明来源是逐行相加');

/* ---------------- 与科目余额表比对 ---------------- */
console.log('\n[4 三级勾稽：明细表 ↔ 科目余额表]');
const agg = Core.aggregateItems([
  { code: '1122', name: '应收账款', debitCents: 550000, creditCents: 0 },
  { code: '1221', name: '其他应收款', debitCents: 80000, creditCents: 0 }
]);
const sheets = [
  { name: '应收账款', rows: sheet },                       // 6000 元 vs 5500 元 → 差 500
  { name: '其他应收款', rows: [['单位', '备注', '金额'], ['甲', '', 800]] },  // 800 元 vs 800 元 → 对上
  { name: '预付账款', rows: [['单位', '备注', '金额'], ['乙', '', 300]] }     // 科目表里没有
];
const rec = Core.reconcileDetail(sheets, agg, { col: 2 });
eq(rec.matched.length, 1, '一项对上');
eq(rec.diffs.length, 1, '一项对不上');
eq(rec.diffs[0].item, '应收账款', '差异项是应收账款');
eq(rec.diffs[0].diffCents, 50000, '差 500 元 = 50000 分');
eq(rec.onlyDetail.length, 1, '明细表有、科目表没有的');
eq(rec.onlyDetail[0].item, '预付账款', '是预付账款');

// 第 3 列（坏账准备）用在「其他应收款」那张表上取不到，要报出来而不是当 0
const rec2 = Core.reconcileDetail(sheets, agg, { col: 3 });
ok(rec2.failed.length >= 2, '取不到金额的表要列进 failed');
ok(!rec2.diffs.some(d => d.diffCents === 0 && d.aggCents === 0), '不能把取不到当成 0 来算');

/* ---------------- 列的选择 ---------------- */
console.log('\n[5 选哪一列由人定，工具不猜]');
const multi = [
  ['单位', '合并数', '其中：本公司', '抵消'],
  ['甲', 5000, 3000, -2000],
  ['乙', 3000, 2000, -1000]
];
const info3 = Core.detectDetailSheet(multi);
eq(Core.readDetailAmount(multi, info3, 1).cents, 800000, '取第 2 列（合并数）');
eq(Core.readDetailAmount(multi, info3, 2).cents, 500000, '取第 3 列（本公司）');
eq(Core.readDetailAmount(multi, info3, 3).cents, -300000, '取第 4 列（抵消，负数）');

console.log('\n' + '='.repeat(56));
if (fails.length) {
  console.log('失败 ' + fails.length + ' / ' + count + ' 项：');
  fails.forEach((f, i) => console.log('  ' + (i + 1) + '. ' + f));
  process.exit(1);
}
console.log('明细表勾稽全部通过：' + count + ' 项断言');
console.log('='.repeat(56));

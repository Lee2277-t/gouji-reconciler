'use strict';
/** 导出装配测试：底稿的四张表、金额格式、文件名 */
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

const box = {
  client: '某公司', dateText: '2026-11-22',
  bs: { label: '余额表.xlsx · Sheet1', accountCount: 12, leafCount: 9, balanceDiffCents: 12300, identityFailed: 2 },
  target: { label: '试算表.xls · 本年试算表', blockName: '资   产', colLabel: '审计前报表数' },
  agg: {
    items: [{ item: '货币资金', cents: 100000000 }],
    unmapped: [
      { code: '1526', name: '某专项资产', signedCents: 6000000, srcRow: 13 },
      { code: '9999', name: '自己也说不清的科目', signedCents: -150000, srcRow: 20 }
    ]
  },
  cmp: {
    diffs: [{ item: '应收账款', leftCents: 30000000, rightCents: 25000000, accounts: [{ code: '1122', name: '应收账款' }] }],
    matched: [{ item: '货币资金', leftCents: 100000000, rightCents: 100000000, accounts: [] }],
    onlyAgg: [{ item: '其他流动资产', cents: 500000 }],
    onlyTarget: [{ item: '其他非流动资产', cents: 6000000 }],
    diffTotalCents: 5000000
  }
};

const sheets = Core.buildSheets(box);
eq(sheets.length, 4, '四张表');
eq(sheets.map(s => s.name), ['勾稽表', '未映射科目', '单边项目', '说明'], '表名与顺序');

/* 勾稽表 */
const gj = sheets[0].rows;
ok(/勾稽表$/.test(gj[0][0]), '第一行是标题');
eq(gj[0][4], '被审计单位：某公司', '抬头带单位名');
eq(gj[3], ['序号', '项目', '科目余额表汇总', '对方表', '差异', '构成科目数'], '表头列名');
// 差异行在前，对上的在后
eq(gj[4][1], '应收账款', '先列差异项');
eq(gj[4][2], 300000, '金额以元为单位（分/100）');
eq(gj[4][3], 250000, '对方表金额');
eq(gj[4][4], 50000, '差异 = 科目 - 对方');
eq(gj[5][1], '货币资金', '对上的一项排在后面');
ok(gj[gj.length - 1][0] === '差异合计', '最后一行是差异合计');
eq(gj[gj.length - 1][4], 50000, '合计金额');

/* 未映射科目 */
const wm = sheets[1].rows;
ok(/未映射科目/.test(wm[0][0]), '标题说明为什么重要');
eq(wm[1], ['科目编码', '科目名称', '期末余额', '原始行号'], '未映射列表头');
eq(wm[2][0], '1526', '编码');
eq(wm[2][2], 60000, '有余额的科目按元展示');
eq(wm[3][2], -1500, '负数（贷方）照实写');
eq(sheets[1].moneyCols, [2], '金额列要套货币格式');

/* 单边项目 */
const db = sheets[2].rows;
eq(db[1], ['方向', '项目', '金额'], '单边项目表头');
ok(db.some(r => r[0] === '只有科目余额表有' && r[1] === '其他流动资产'), '科目侧独有');
ok(db.some(r => r[0] === '只有对方表有' && r[1] === '其他非流动资产'), '报表侧独有');

/* 说明 */
const sm = sheets[3].rows;
const flat = sm.map(r => r.join('|')).join('\n');
ok(/勾稽助手 v/.test(flat), '写明工具版本');
ok(/借贷差额/.test(flat), '写明借贷差额');
ok(/123\.00/.test(flat), '差额 12300 分要写成 123.00 元');
ok(/本月|不含损益|内部往来/.test(flat), '不平的时候要给原因提示');
ok(/恒等式/.test(flat), '写明恒等式检查结果');
ok(/不代替审计判断/.test(flat), '免责一句');

/* 文件名 */
eq(Core.fileName(box), '勾稽_某公司_审计前报表数_20261122.xlsx', '文件名规则');
eq(Core.fileName({ client: 'a/b:c', dateText: '2026-11-22', target: { colLabel: '审计前' } }),
  '勾稽_a-b-c_审计前_20261122.xlsx', '文件名里的非法字符要替换');

/* 真的装配成工作簿 */
const out = Core.buildWorkbook(XLSX, box);
eq(out.sheetNames, ['勾稽表', '未映射科目', '单边项目', '说明'], '工作簿页签');
const ws = out.wb.Sheets['勾稽表'];
const c = ws[XLSX.utils.encode_cell({ r: 4, c: 2 })];
ok(c && c.t === 'n' && c.z === '#,##0.00', '金额单元格套了货币格式');
const cntCell = ws[XLSX.utils.encode_cell({ r: 4, c: 5 })];
ok(cntCell && cntCell.z !== '#,##0.00', '「构成科目数」不能被套成金额');
ok(out.fileName.endsWith('.xlsx'), '文件名后缀');

/* 空数据不能崩 */
try {
  const empty = Core.buildWorkbook(XLSX, { client: '', dateText: '', bs: {}, target: {}, agg: { items: [], unmapped: [] }, cmp: { diffs: [], matched: [], onlyAgg: [], onlyTarget: [], diffTotalCents: 0 } });
  ok(empty.sheetNames.length === 4, '没有数据也能出一份空底稿');
} catch (e) {
  ok(false, '空数据装配失败：' + e.message);
}

console.log('\n' + '='.repeat(56));
if (fails.length) {
  console.log('失败 ' + fails.length + ' / ' + count + ' 项：');
  fails.forEach((f, i) => console.log('  ' + (i + 1) + '. ' + f));
  process.exit(1);
}
console.log('导出装配全部通过：' + count + ' 项断言');
console.log('='.repeat(56));

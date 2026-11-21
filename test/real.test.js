'use strict';
/**
 * 真实文件回归：拿本机真实底稿跑一遍，只验「结构认得出、数出得来、不崩」。
 *
 * 客户路径不进仓库 —— 通过环境变量给：
 *   $env:GOUJI_REAL_FILES = "余额表.xlsx;明细表.xlsx;试算表.xls"
 * 没设置就跳过（CI/别人 clone 下来不会红）。
 *
 * 断言只看结构和自洽性，不看具体金额，所以不会把客户数据带进测试里。
 */
const fs = require('fs');
const path = require('path');
const XLSX = require(path.join(__dirname, '..', 'vendor', 'xlsx.full.min.js'));
const Core = require(path.join(__dirname, '..', 'core.js'));

let count = 0;
const fails = [];
function ok(cond, msg) { count++; if (!cond) fails.push(msg); }

const raw = process.env.GOUJI_REAL_FILES || '';
const files = raw.split(';').map(s => s.trim()).filter(Boolean);

if (!files.length) {
  console.log('跳过：没有设置 GOUJI_REAL_FILES');
  console.log('（想跑真实文件回归，先设置这个环境变量，指向本机的科目余额表/明细表/试算表）');
  process.exit(0);
}

const missing = files.filter(f => !fs.existsSync(f));
if (missing.length) {
  console.log('跳过：这些文件不存在 —— ' + missing.join('、'));
  process.exit(0);
}

function readBook(file) {
  return XLSX.read(fs.readFileSync(file), { type: 'buffer' });
}
function rowsOf(wb, name) {
  return XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, raw: true, defval: null, blankrows: true });
}

let sawBalance = 0, sawStatement = 0, sawDetail = 0;

files.forEach(file => {
  const label = path.basename(file);
  let wb;
  try {
    wb = readBook(file);
    ok(true, label + ' 能读');
  } catch (e) {
    ok(false, label + ' 读不了：' + e.message);
    return;
  }

  let hit = false;

  // 当作科目余额表试
  for (const name of wb.SheetNames) {
    const rows = rowsOf(wb, name);
    const h = Core.detectBalanceHeader(rows);
    if (h.rowIndex < 0) continue;
    const p = Core.parseBalanceRows(rows, h, { firstExcelRow: 1 });
    if (p.accounts.length < 5) continue;

    sawBalance++; hit = true;
    ok(p.accounts.length > 0, label + ' / ' + name + '：认出 ' + p.accounts.length + ' 个科目');
    ok(p.accounts.every(a => Core.codePrefix(a.code) !== ''), label + '：每个科目都有可识别的编码');
    ok(h.codeCol >= 0 && h.nameCol >= 0, label + '：编码列和名称列都认出来了');

    const leaves = Core.leafAccounts(p.accounts);
    ok(leaves.length > 0 && leaves.length <= p.accounts.length, label + '：末级科目数合理（' + leaves.length + '/' + p.accounts.length + '）');

    const ic = Core.identityCheck(p.accounts);
    // 恒等式应该绝大多数成立；成立率异常低说明列读错了
    const ratio = ic.checked ? (ic.checked - ic.failed) / ic.checked : 1;
    ok(ratio >= 0.9, label + '：恒等式成立率 ' + (ratio * 100).toFixed(1) + '%（应 ≥ 90%）');

    const bc = Core.balanceCheck(p.accounts);
    ok(typeof bc.debitCents === 'number' && typeof bc.creditCents === 'number', label + '：借贷合计算得出来');

    const agg = Core.aggregateItems(leaves);
    ok(agg.items.length > 0, label + '：汇总出 ' + agg.items.length + ' 个报表项目');
    ok(agg.items.every(i => typeof i.cents === 'number' && isFinite(i.cents)), label + '：每个项目金额都是有限数');
    ok(agg.unmapped.every(u => u.code !== undefined), label + '：未映射科目都带编码（' + agg.unmapped.length + ' 个）');
    break;
  }

  // 当作试算表/报表试
  if (!hit) {
    for (const name of wb.SheetNames) {
      const rows = rowsOf(wb, name);
      const h = Core.detectStatementHeader(rows);
      if (h.rowIndex < 0 || !h.blocks.length) continue;
      const block = Core.parseStatementBlock(rows, h, 0, 0, { firstExcelRow: 1 });
      if (block.items.length < 3) continue;

      sawStatement++; hit = true;
      ok(block.items.length > 0, label + ' / ' + name + '：读到 ' + block.items.length + ' 个项目');
      ok(h.blocks.length >= 1, label + '：切成 ' + h.blocks.length + ' 栏');
      break;
    }
  }

  // 当作明细表试
  if (!hit) {
    let sheets = 0, withNum = 0;
    for (const name of wb.SheetNames) {
      const rows = rowsOf(wb, name);
      const info = Core.detectDetailSheet(rows);
      sheets++;
      if (!info.numericCols.length) continue;
      withNum++;
      const got = Core.readDetailAmount(rows, info, info.numericCols[0]);
      ok(got.cents === null || isFinite(got.cents), label + ' / ' + name + '：取出的金额是有限数');
    }
    if (withNum > 0) {
      sawDetail++;
      ok(withNum <= sheets, label + '：' + withNum + '/' + sheets + ' 张表带数字');
    }
  }
});

ok(sawBalance + sawStatement + sawDetail > 0, '至少认出一张表（余额表 ' + sawBalance + ' / 试算表 ' + sawStatement + ' / 明细表 ' + sawDetail + '）');

console.log('\n' + '='.repeat(56));
if (fails.length) {
  console.log('失败 ' + fails.length + ' / ' + count + ' 项：');
  fails.forEach((f, i) => console.log('  ' + (i + 1) + '. ' + f));
  process.exit(1);
}
console.log('真实文件回归全部通过：' + count + ' 项断言');
console.log('='.repeat(56));

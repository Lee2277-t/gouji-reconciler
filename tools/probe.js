/**
 * 结构探针：只看「表长什么样」，不打印任何金额或客户文字。
 * 数字一律显示为 n，长字符串显示为 s(长度)，避免把客户数据带进终端。
 *
 * 用法： node tools/probe.js <xlsx|xls> [工作表名] [--cols N] [--rows N]
 */
'use strict';
const fs = require('fs');
const path = require('path');
const XLSX = require('../vendor/xlsx.full.min.js');

const args = process.argv.slice(2);
const file = args[0];
const sheetArg = args[1] && !args[1].startsWith('--') ? args[1] : null;
const opt = (name, dflt) => {
  const i = args.indexOf('--' + name);
  return i >= 0 && args[i + 1] ? parseInt(args[i + 1], 10) : dflt;
};
const MAX_ROWS = opt('rows', 12);
const MAX_COLS = opt('cols', 16);

if (!file) {
  console.error('用法: node tools/probe.js <xlsx|xls> [工作表] [--rows N] [--cols N]');
  process.exit(2);
}

/** 一个格子脱敏后的样子 */
function mask(v) {
  if (v === null || v === undefined || v === '') return '·';
  if (typeof v === 'number') return 'n';
  if (v instanceof Date) return 'D';
  const s = String(v);
  if (s.length > 8) return 's' + s.length;
  // 短文本多半是表头或科目名，值得看见；但只保留前 6 个字
  return s.slice(0, 6);
}

const wb = XLSX.read(fs.readFileSync(file), { type: 'buffer', cellDates: false });
console.log('文件: ' + path.basename(file));
console.log('工作表(' + wb.SheetNames.length + '): ' + wb.SheetNames.join(' | '));

const targets = sheetArg ? [sheetArg] : wb.SheetNames.slice(0, 1);

for (const name of targets) {
  const ws = wb.Sheets[name];
  if (!ws) { console.log('\n[缺工作表] ' + name); continue; }
  const ref = ws['!ref'] || 'A1:A1';
  const range = XLSX.utils.decode_range(ref);
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: true });

  console.log('\n' + '='.repeat(60));
  console.log('[' + name + '] ' + ref + '  ' + (range.e.r + 1) + ' 行 × ' + (range.e.c + 1) + ' 列');
  // 合并单元格是两级表头的信号
  const merges = ws['!merges'] || [];
  if (merges.length) {
    console.log('合并单元格 ' + merges.length + ' 处，前 5: ' +
      merges.slice(0, 5).map(m => XLSX.utils.encode_cell(m.s) + ':' + XLSX.utils.encode_cell(m.e)).join(', '));
  }

  for (let r = 0; r < Math.min(rows.length, MAX_ROWS); r++) {
    const row = (rows[r] || []).slice(0, MAX_COLS).map(mask);
    while (row.length && row[row.length - 1] === '·') row.pop();
    console.log(String(r + 1).padStart(4) + ' ' + (row.join(' | ') || '（空行）'));
  }

  // 每列的「填充率」和「类型」：判断哪列是编码、哪列是金额
  const colInfo = [];
  for (let c = 0; c <= range.e.c; c++) {
    let num = 0, txt = 0, empty = 0;
    for (let r = 0; r < rows.length; r++) {
      const v = (rows[r] || [])[c];
      if (v === null || v === undefined || v === '') empty++;
      else if (typeof v === 'number') num++;
      else txt++;
    }
    colInfo.push({ c, num, txt, empty });
  }
  const filled = colInfo.filter(x => x.num + x.txt > 0);
  console.log('列概况: ' + filled.map(x =>
    XLSX.utils.encode_col(x.c) + '(数字' + x.num + '/文本' + x.txt + ')').join(' '));
}

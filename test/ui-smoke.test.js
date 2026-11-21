'use strict';
/**
 * 界面冒烟测试：用假 DOM 加载真实的 app.js，点「示例数据」，
 * 检查页面上的统计、差异表、未映射清单是否按预期渲染。
 * 不测样式，只测「点了以后有没有把对的东西显示出来」。
 */
const fs = require('fs');
const path = require('path');
const vm = require('vm');

let count = 0;
const fails = [];
function ok(cond, msg) { count++; if (!cond) fails.push(msg); }

/* ---------------- 假 DOM ---------------- */

function makeEl(id) {
  const handlers = {};
  const el = {
    id,
    innerHTML: '',
    value: '',
    textContent: '',
    disabled: false,
    style: {},
    files: null,
    _handlers: handlers,
    classList: {
      _set: new Set(),
      add(c) { this._set.add(c); },
      remove(c) { this._set.delete(c); },
      contains(c) { return this._set.has(c); },
      toggle(c, force) {
        if (force === undefined) { this._set.has(c) ? this._set.delete(c) : this._set.add(c); }
        else if (force) this._set.add(c); else this._set.delete(c);
      }
    },
    addEventListener(type, fn) { (handlers[type] = handlers[type] || []).push(fn); },
    click() { (handlers.click || []).forEach(fn => fn({ target: el })); },
    appendChild() {},
    setAttribute() {},
    querySelector() { return null; }
  };
  return el;
}

const els = {};
const document = {
  getElementById(id) { return els[id] || (els[id] = makeEl(id)); },
  body: makeEl('body'),
  createElement(tag) { return makeEl(tag); },
  addEventListener() {}
};

// body.classList 要为 mode-simple 服务
document.body.classList = {
  _set: new Set(['mode-simple']),
  add(c) { this._set.add(c); },
  remove(c) { this._set.delete(c); },
  contains(c) { return this._set.has(c); },
  toggle(c, force) {
    if (force === undefined) { this._set.has(c) ? this._set.delete(c) : this._set.add(c); }
    else if (force) this._set.add(c); else this._set.delete(c);
  }
};

const store = {};
const localStorage = {
  getItem(k) { return Object.prototype.hasOwnProperty.call(store, k) ? store[k] : null; },
  setItem(k, v) { store[k] = String(v); },
  removeItem(k) { delete store[k]; }
};

const XLSX = require(path.join(__dirname, '..', 'vendor', 'xlsx.full.min.js'));
const Core = require(path.join(__dirname, '..', 'core.js'));

// FileReader 桩：这个测试不真的读文件
function FileReader() {}
FileReader.prototype.readAsArrayBuffer = function () {};

const sandbox = {
  window: { localStorage, document },
  document, localStorage, XLSX, Core, FileReader,
  console, Date, JSON, Math, String, Number, Object, Array, parseInt, parseFloat, isNaN
};
sandbox.window.window = sandbox.window;
vm.createContext(sandbox);

/* ---------------- 加载 app.js ---------------- */
const appSrc = fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8');
try {
  vm.runInContext(appSrc, sandbox, { filename: 'app.js' });
  ok(true, 'app.js 能加载');
} catch (e) {
  ok(false, 'app.js 加载失败：' + e.message);
}

/* ---------------- 点示例数据 ---------------- */

try {
  document.getElementById('stScope').value = 'auto';
document.getElementById('demoBtn').click();
  ok(true, '点「示例数据」不报错');
} catch (e) {
  ok(false, '点「示例数据」报错：' + e.message);
}

const stats = document.getElementById('statsBox').innerHTML;
const diff = document.getElementById('diffBox').innerHTML;
const unmapped = document.getElementById('unmappedBox').innerHTML;
const bsSum = document.getElementById('bsSummary').innerHTML;
const stSum = document.getElementById('stSummary').innerHTML;

/* ---------------- 断言 ---------------- */

ok(/认出/.test(bsSum), '科目余额表要显示「认出 N 个科目」');
ok(/借贷平/.test(bsSum) || /差/.test(bsSum), '要显示借贷是否平');
ok(/读到/.test(stSum), '试算表要显示「读到 N 个项目」');

ok(/对上/.test(stats), '统计里要有「对上」');
ok(/未映射科目/.test(stats), '统计里要有「未映射科目」');
ok(/差异金额/.test(stats), '统计里要有差异金额');

// 示例数据：货币资金 100 万应能对上（1002 有下级，只算末级才不会翻倍）
ok(!/货币资金/.test(diff) || !/class="num"><b>/.test(diff),
  '货币资金应能对上，不该出现在差异表里');

// 示例里 1526 故意没映射，应该出现在未映射清单
ok(/1526/.test(unmapped), '未映射科目里要出现 1526');
ok(/某专项资产/.test(unmapped), '未映射清单要带科目名称');
ok(/其他非流动资产/.test(diff), '试算表独有的项目要能被识别出来');

// 导出按钮在两张表都就绪后应该可用
ok(document.getElementById('genBtn').disabled === false, '两张表选好后「生成勾稽底稿」应可点');

/* ---------------- 导出装配能跑 ---------------- */

try {
  const out = Core.buildWorkbook(XLSX, {
    client: '', dateText: '2026-11-22',
    bs: { label: 'a', accountCount: 12, leafCount: 8, balanceDiffCents: 0, identityFailed: 0 },
    target: { label: 'b', blockName: '资   产', colLabel: '审计前报表数' },
    agg: { items: [{ item: '货币资金', cents: 100000000 }], unmapped: [{ code: '1526', name: '某专项资产', signedCents: 6000000, srcRow: 13 }] },
    cmp: { diffs: [], matched: [{ item: '货币资金', leftCents: 100000000, rightCents: 100000000 }], onlyAgg: [], onlyTarget: [{ item: '其他非流动资产', cents: 6000000 }], diffTotalCents: 0 }
  });
  ok(out.sheetNames.length === 4, '底稿应有 4 张表，实际 ' + out.sheetNames.length);
  ok(/勾稽_/.test(out.fileName), '文件名要以「勾稽_」开头');
} catch (e) {
  ok(false, '底稿装配失败：' + e.message);
}

/* ---------------- 汇总 ---------------- */
console.log('\n' + '='.repeat(56));
if (fails.length) {
  console.log('失败 ' + fails.length + ' / ' + count + ' 项：');
  fails.forEach((f, i) => console.log('  ' + (i + 1) + '. ' + f));
  process.exit(1);
}
console.log('界面冒烟全部通过：' + count + ' 项断言');
console.log('='.repeat(56));

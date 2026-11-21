/**
 * 勾稽助手 · app.js
 * 界面层：读文件、认列、调 core、渲染、导出。业务判断一律在 core.js 里。
 */
'use strict';

(function () {

  var $ = function (id) { return document.getElementById(id); };

  var state = {
    bs: null,        // {fileName, wb, sheetName, rows, header, parsed, level}
    st: null,        // {fileName, wb, sheetName, rows, header, blockIndex, colIndex, parsed}
    agg: null,
    cmp: null,
    overrides: {},
    verdict: null
  };

  var MAP_KEY = 'gouji.map.v1';

  function esc(s) {
    return String(s === null || s === undefined ? '' : s)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }
  function money(cents) { return Core.moneyText(cents); }

  /* ================= 简易 / 高级 ================= */

  function isSimple() { return document.body.classList.contains('mode-simple'); }

  function setMode(simple) {
    document.body.classList.toggle('mode-simple', simple);
    $('modeSimple').classList.toggle('on', simple);
    $('modeExpert').classList.toggle('on', !simple);
    $('modeHint').textContent = simple
      ? '两步选文件，第三步看差异。认错了列再切高级。'
      : '所有识别结果都露出来了，可以逐列核对。';
    try { window.localStorage.setItem('gouji.mode', simple ? 'simple' : 'expert'); } catch (e) { /* 忽略 */ }
    renderMapPanel();
  }

  /* ================= 读文件 ================= */

  function readWorkbook(file) {
    return new Promise(function (resolve) {
      var reader = new FileReader();
      reader.onload = function () {
        try {
          var wb = XLSX.read(new Uint8Array(reader.result), { type: 'array', cellDates: false });
          resolve({ fileName: file.name, wb: wb, ok: true });
        } catch (err) {
          resolve({ fileName: file.name, ok: false, error: err.message });
        }
      };
      reader.onerror = function () { resolve({ fileName: file.name, ok: false, error: '读不了这个文件' }); };
      reader.readAsArrayBuffer(file);
    });
  }

  function sheetRows(wb, name) {
    return XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, raw: true, defval: null, blankrows: true });
  }

  /** 给每张表打分，挑最像科目余额表的那张 */
  function pickBalanceSheet(wb) {
    var best = null;
    wb.SheetNames.forEach(function (name) {
      var rows = sheetRows(wb, name);
      var h = Core.detectBalanceHeader(rows);
      var score = h.rowIndex >= 0 ? (100 - h.problems.length * 30) : 0;
      var n = h.rowIndex >= 0 ? rows.length : 0;
      if (!best || score + n / 1000 > best.score) best = { name: name, rows: rows, header: h, score: score };
    });
    return best;
  }

  function pickStatementSheet(wb) {
    var best = null;
    wb.SheetNames.forEach(function (name) {
      var rows = sheetRows(wb, name);
      var h = Core.detectStatementHeader(rows);
      if (h.rowIndex < 0) return;
      var items = 0;
      try { items = Core.parseStatementBlock(rows, h, 0, 0).items.length; } catch (e) { items = 0; }
      var score = items;
      if (!best || score > best.score) best = { name: name, rows: rows, header: h, score: score };
    });
    return best;
  }

  /* ================= 失败时清场 ================= */

  /** 读文件失败/认不出表头：把上一步的结果全部清掉，别让人对着旧数字做判断 */
  function clearAll(msg) {
    state.bs = null;
    state.agg = null;
    state.cmp = null;
    state.detailCmp = null;
    $('bsInfo').textContent = msg || '';
    $('bsSummary').classList.add('hidden');
    $('bsPanel').classList.add('hidden');
    $('bsMapTable').classList.add('hidden');
    $('statsBox').innerHTML = '';
    $('diffBox').innerHTML = '';
    $('unmappedBox').innerHTML = '';
    $('detailBox').innerHTML = '';
    $('genBtn').disabled = true;
    $('genInfo').textContent = '';
  }

  function clearStatement(msg) {
    state.st = null;
    state.cmp = null;
    $('stInfo').textContent = msg || '';
    $('stSummary').classList.add('hidden');
    $('stPanel').classList.add('hidden');
    $('diffBox').innerHTML = '';
    $('statsBox').innerHTML = '';
    $('genBtn').disabled = true;
  }

  /* ================= 科目余额表 ================= */

  function loadBalanceSheet(file) {
    readWorkbook(file).then(function (res) {
      if (!res.ok) { clearAll('读不了：' + res.error); return; }
      var pick = pickBalanceSheet(res.wb);
      if (!pick || pick.header.rowIndex < 0) {
        clearAll('这个文件里找不到科目余额表的表头（要能认出「科目编码」「科目名称」「余额」）。判定不了的表不能硬算，上一份结果已清空。');
        return;
      }
      var opts = [];
      res.wb.SheetNames.forEach(function (n) {
        var h = Core.detectBalanceHeader(sheetRows(res.wb, n));
        opts.push({ name: n, ok: h.rowIndex >= 0 });
      });
      $('bsSheet').innerHTML = opts.map(function (o) {
        return '<option value="' + esc(o.name) + '"' + (o.name === pick.name ? ' selected' : '') + '>' +
          esc(o.name) + (o.ok ? '' : '（不像余额表）') + '</option>';
      }).join('');

      state.bs = {
        fileName: res.fileName, wb: res.wb, sheetName: pick.name,
        rows: pick.rows, header: pick.header, parsed: null, level: 'leaf'
      };
      // 面板一直带 hidden，之前没有任何地方移除它，等于换工作表/改表头行/改口径全用不了
      $('bsPanel').classList.remove('hidden');
      $('bsHeaderRow').value = pick.header.rowIndex + 1;
      $('bsInfo').textContent = res.fileName + ' · ' + pick.name;
      parseBalance();
      refresh();
    });
  }

  function parseBalance() {
    if (!state.bs) return;
    var h = state.bs.header;
    state.bs.parsed = Core.parseBalanceRows(state.bs.rows, h, { firstExcelRow: 1 });
    var accounts = state.bs.parsed.accounts;
    var leaves = Core.leafAccounts(accounts);
    var bc = Core.balanceCheck(accounts);
    var ic = Core.identityCheck(accounts);
    state.bs.stats = {
      accounts: accounts.length,
      leaves: leaves.length,
      bc: bc,
      ic: ic,
      // 有内容但没计入科目的行（编码取不出四位数字）：必须进结论，否则金额凭空消失
      skippedRows: (state.bs.parsed.skipped || []).filter(function (r) {
        return (r.code && Core.str(r.code).trim()) || (r.name && Core.str(r.name).trim());
      }).length,
      problems: (state.bs.parsed.problems || []).concat(state.bs.header.problems || [])
    };
    renderBsMap();
  }

  function currentAccounts() {
    if (!state.bs || !state.bs.parsed) return [];
    var all = state.bs.parsed.accounts;
    var level = state.bs.level || 'leaf';
    if (level === 'all') return all;
    if (level === 'leaf') return Core.leafAccounts(all);
    // 只算上级：被别的编码引用为前缀的那些
    var codes = all.map(function (a) { return Core.str(a.code).replace(/[^\d]/g, ''); });
    return all.filter(function (a, i) {
      var c = codes[i];
      if (!c) return false;
      return codes.some(function (o, j) { return j !== i && o && o !== c && o.indexOf(c) === 0; });
    });
  }

  function renderBsMap() {
    var t = $('bsMapTable');
    if (!state.bs || !state.bs.parsed) { t.innerHTML = ''; t.classList.add('hidden'); return; }
    t.classList.remove('hidden');
    var h = state.bs.header;
    var cols = [];
    [['opening', '期初余额'], ['period', '期间发生额'], ['year', '累计发生额'], ['closing', '期末余额']]
      .forEach(function (g) {
        var grp = h.columns[g[0]] || {};
        ['dir', 'amount', 'debit', 'credit'].forEach(function (k) {
          if (grp[k] !== undefined) cols.push({ group: g[1], kind: k, col: grp[k] });
        });
      });
    var kindName = { dir: '方向', amount: '金额', debit: '借方', credit: '贷方' };
    t.innerHTML = '<tr><th>编码列</th><th>名称列</th><th>表头行</th><th>识别到的金额列</th></tr>' +
      '<tr><td>' + (h.codeCol + 1) + ' 列</td><td>' + (h.nameCol + 1) + ' 列</td>' +
      '<td>第 ' + (h.rowIndex + 1) + ' 行' + (h.span > 1 ? '（两级表头）' : '') + '</td>' +
      '<td>' + cols.map(function (c) {
        return c.group + '·' + kindName[c.kind] + '=第' + (c.col + 1) + '列';
      }).join('；') + '</td></tr>';
  }

  /* ================= 试算表 / 报表 ================= */

  function loadStatement(file) {
    readWorkbook(file).then(function (res) {
      if (!res.ok) { clearStatement('读不了：' + res.error); return; }
      var pick = pickStatementSheet(res.wb);
      if (!pick) {
        clearStatement('这个文件里找不到能当试算表/报表的表（要能认出「项目」和「金额」列）。');
        return;
      }
      var opts = [];
      res.wb.SheetNames.forEach(function (n) {
        var h = Core.detectStatementHeader(sheetRows(res.wb, n));
        opts.push({ name: n, ok: h.rowIndex >= 0 });
      });
      $('stSheet').innerHTML = opts.map(function (o) {
        return '<option value="' + esc(o.name) + '"' + (o.name === pick.name ? ' selected' : '') + '>' +
          esc(o.name) + (o.ok ? '' : '（不像报表）') + '</option>';
      }).join('');

      state.st = {
        fileName: res.fileName, wb: res.wb, sheetName: pick.name,
        rows: pick.rows, header: pick.header, blockIndex: 0, colIndex: 0, parsed: null
      };
      $('stInfo').textContent = res.fileName + ' · ' + pick.name;
      $('stPanel').classList.remove('hidden');
      renderStatementSelectors();
      parseStatement();
      refresh();
    });
  }

  function renderStatementSelectors() {
    var h = state.st.header;
    $('stBlock').innerHTML = h.blocks.map(function (b, i) {
      return '<option value="' + i + '">' + esc(b.name) + '</option>';
    }).join('');
    $('stBlock').value = String(state.st.blockIndex);
    renderColSelect();
  }

  function renderColSelect() {
    var h = state.st.header;
    var b = h.blocks[state.st.blockIndex] || { amountCols: [] };
    $('stCol').innerHTML = b.amountCols.map(function (a, i) {
      return '<option value="' + i + '">' + esc(a.label) + '</option>';
    }).join('');
    if (state.st.colIndex >= b.amountCols.length) state.st.colIndex = 0;
    $('stCol').value = String(state.st.colIndex);
  }

  function parseStatement() {
    if (!state.st) return;
    var r = Core.parseStatementBlock(state.st.rows, state.st.header,
      state.st.blockIndex, state.st.colIndex, { firstExcelRow: 1 });
    state.st.parsed = r;
  }

  /** 对比范围：auto 时按这一栏的标题猜，省得用户自己想 */
  function resolveScope() {
    var v = ($('stScope') && $('stScope').value) || 'auto';
    if (v === 'all') return '';
    if (v === 'auto') {
      var name = ((state.st && state.st.header.blocks[state.st.blockIndex]) || {}).name || '';
      var t = Core.norm(name);
      if (/资产/.test(t) && !/负债/.test(t)) return ['资产'];
      if (/负债|权益/.test(t)) return ['负债', '权益'];
      if (/利润|损益/.test(t)) return ['损益'];
      return '';
    }
    return [v];
  }

  /* ================= 明细表 ================= */

  function loadDetail(file) {
    readWorkbook(file).then(function (res) {
      if (!res.ok) { $('dtInfo').textContent = '读不了：' + res.error; return; }
      setDetail(res.fileName, res.wb);
    });
  }

  /** 认明细表：只留带数字的工作表 */
  function setDetail(fileName, wb) {
    var sheets = [];
    wb.SheetNames.forEach(function (name) {
      var rows = sheetRows(wb, name);
      var info = Core.detectDetailSheet(rows);
      if (info.numericCols.length) sheets.push({ name: name, rows: rows, info: info });
    });
    if (!sheets.length) {
      $('dtInfo').textContent = '这个文件里没有带数字的工作表，不像明细表。';
      state.dt = null; refresh(); return;
    }
    state.dt = { fileName: fileName, sheets: sheets, colIndex: null };
    $('dtPanel').classList.remove('hidden');
    $('dtInfo').textContent = fileName + ' · ' + sheets.length + ' 张明细表';
    renderDetailCols();
    refresh();
  }

  /** 列选项：把所有表里出现过的数字列并起来，带表头文字和合计，方便认 */
  function renderDetailCols() {
    var cols = {};
    state.dt.sheets.forEach(function (s) {
      s.info.numericCols.forEach(function (c) {
        if (!cols[c]) cols[c] = { col: c, label: (s.info.headers[c] || ''), sample: s.info.colSums[c] };
      });
    });
    var list = Object.keys(cols).map(function (k) { return cols[k]; })
      .sort(function (a, b) { return a.col - b.col; });
    $('dtCol').innerHTML = list.map(function (c) {
      return '<option value="' + c.col + '">第 ' + (c.col + 1) + ' 列' +
        (c.label ? ' [' + esc(c.label) + ']' : '') + '</option>';
    }).join('');
    if (state.dt.colIndex !== null && cols[state.dt.colIndex]) {
      $('dtCol').value = String(state.dt.colIndex);
    } else if (list.length) {
      // 默认列不能取"第一个有数字的列"——那张表里第一个数字列往往是序号或账号，
      // 示例数据就会因此算出 −98 万的假差异。先按表头关键词认，认不出取绝对值最大的列。
      var kw = list.filter(function (c) { return /金额|余额|账面|价值|余额数|数$/.test(c.label || ''); });
      var pick = kw.length ? kw[0] : list.reduce(function (a, b) {
        return Math.abs(b.sample) > Math.abs(a.sample) ? b : a;
      }, list[0]);
      state.dt.colIndex = pick.col;
      $('dtCol').value = String(pick.col);
    }
  }

  function renderDetail() {
    var box = $('detailBox');
    if (!state.dt || !state.agg) { box.innerHTML = ''; return; }
    var rec = Core.reconcileDetail(state.dt.sheets, state.agg, { col: state.dt.colIndex });
    state.detailCmp = rec;

    var html = '<div class="stats">' +
      '<div class="cell ok"><div class="k">对上</div><div class="v">' + rec.matched.length + '<span class="u"> 项</span></div></div>' +
      '<div class="cell ' + (rec.diffs.length ? 'bad' : '') + '"><div class="k">对不上</div><div class="v">' + rec.diffs.length + '<span class="u"> 项</span></div></div>' +
      '<div class="cell ' + (rec.failed.length ? 'warn' : '') + '"><div class="k">取不到金额</div><div class="v">' + rec.failed.length + '<span class="u"> 张表</span></div></div>' +
      '</div>';

    var show = rec.diffs.concat(rec.matched).slice(0, 60);
    if (show.length) {
      html += '<div class="scroll"><table class="data"><thead><tr>' +
        '<th>工作表（报表项目）</th><th class="num">明细表合计</th><th class="num">科目余额表汇总</th>' +
        '<th class="num">差异</th><th>取数</th></tr></thead><tbody>';
      show.forEach(function (d) {
        var cls = Math.abs(d.diffCents) > 0 ? ' class="diff"' : '';
        html += '<tr' + cls + '><td>' + esc(d.item) + '</td>' +
          '<td class="num">' + money(d.detailCents) + '</td>' +
          '<td class="num">' + (d.aggCents === null ? '—' : money(d.aggCents)) + '</td>' +
          '<td class="num"><b>' + (d.aggCents === null ? '—' : money(d.diffCents)) + '</b></td>' +
          '<td class="muted">' + esc(d.source) + (d.totalRow ? '（第 ' + d.totalRow + ' 行）' : '') + '</td></tr>';
      });
      html += '</tbody></table></div>';
    }
    if (rec.onlyAgg.length) {
      html += '<h3>科目余额表有、明细表没有的（' + rec.onlyAgg.length + ' 项）</h3>' +
        '<div class="note warn">这些项目在明细表里找不到同名的工作表。可能是没做明细，也可能是工作表名不一样。</div>' +
        '<div class="scroll"><table class="data"><thead><tr><th>项目</th><th class="num">科目余额表汇总</th></tr></thead><tbody>' +
        rec.onlyAgg.slice(0, 30).map(function (x) {
          return '<tr><td>' + esc(x.item) + '</td><td class="num">' + money(x.cents) + '</td></tr>';
        }).join('') + '</tbody></table></div>';
    }
    if (rec.failed.length) {
      html += '<p class="hint">这几张表在这一列上没有数字：' +
        rec.failed.slice(0, 8).map(function (f) { return esc(f.sheet); }).join('、') + '</p>';
    }
    box.innerHTML = html;
  }

  /* ================= 刷新与渲染 ================= */

  function refresh() {
    if (!state.bs || !state.bs.parsed) {
      $('bsSummary').classList.add('hidden');
      // 没有可用的余额表：把下游全部清掉，别让上一份的差异表留在屏幕上
      state.agg = null;
      state.cmp = null;
      $('statsBox').innerHTML = '';
      $('diffBox').innerHTML = '';
      $('unmappedBox').innerHTML = '';
      $('genBtn').disabled = true;
      return;
    }
    var s = state.bs.stats;
    var sum = $('bsSummary');
    sum.classList.remove('hidden');
    var bc = s.bc;
    var levelNote = { leaf: '只算末级', top: '只算上级', all: '全部科目' }[state.bs.level] || '只算末级';

    // 一个期末余额都没取到：不能说「借贷平」，那是 0=0 造成的假象
    if (!bc.conclusive) {
      sum.className = 'note bad';
      sum.innerHTML = '认出 <b>' + s.accounts + '</b> 个科目，但<b>一个期末余额都没取到</b>——' +
        '这张表的金额列没认出来，不能用来勾稽。' +
        (s.problems && s.problems.length ? '<br>原因：' + esc(s.problems.join('；')) : '') +
        '<br>请在高级里改「表头所在行」，或确认这张表有没有「期末余额/借方/贷方」这类列。';
      $('genBtn').disabled = true;
      // 结论都下不了，下面那些差异表也必须清掉，否则红字上面还挂着绿色的「分毫不差」
      state.agg = null;
      state.cmp = null;
      $('statsBox').innerHTML = '';
      $('diffBox').innerHTML = '';
      $('unmappedBox').innerHTML = '';
      $('detailBox').innerHTML = '';
      return;
    }

    sum.className = 'note ' + (bc.balanced ? 'ok' : 'warn');
    sum.innerHTML = '认出 <b>' + s.accounts + '</b> 个科目（末级 ' + s.leaves + ' 个，' + levelNote + '）' +
      '；借方合计 <b>' + money(bc.debitCents) + '</b>，贷方 <b>' + money(bc.creditCents) + '</b>' +
      (bc.balanced ? '，借贷平' : '，<b>差 ' + money(bc.diffCents) + '</b>') +
      (s.ic.failed ? '；另有 <b>' + s.ic.failed + '</b> 行「期初+本期借-本期贷≠期末」' : '') +
      (s.ic.inferred ? '；<b>' + s.ic.inferred + '</b> 行期末余额是推算的（这些行不参与恒等式校验）' : '') +
      (s.problems && s.problems.length ? '；注意：' + esc(s.problems.join('；')) : '') +
      '。';

    // 恒等式对不上的行要指到具体哪一行，只说「有 1 行」没法查
    if (s.ic.failed && (s.ic.bad || []).length) {
      sum.innerHTML += '<div class="scroll" style="margin-top:8px"><table class="data"><thead><tr>' +
        '<th>期初+本期借-本期贷≠期末的科目</th><th class="num">原始行号</th></tr></thead><tbody>' +
        s.ic.bad.slice(0, 15).map(function (b) {
          return '<tr><td>' + esc(b.code + ' ' + b.name) + '</td><td class="num">' + esc(b.srcRow) + '</td></tr>';
        }).join('') +
        '</tbody></table></div>' +
        (s.ic.bad.length > 15 ? '<p class="muted">只列前 15 行，完整的在导出的底稿里。</p>' : '');
    }
    if (s.skippedRows) {
      sum.innerHTML += '<div class="note bad" style="margin-top:8px">原表有 <b>' + s.skippedRows +
        '</b> 行有内容但科目编码里取不出四位数字，这些行没有计入——金额不会凭空消失，但也没有归到科目上。请回原表核对。</div>';
    }

    var accounts = currentAccounts();
    state.agg = Core.aggregateItems(accounts, { overrides: state.overrides });

    if (state.st && state.st.parsed) {
      state.cmp = Core.compareItems(state.agg, state.st.parsed.items, { cls: resolveScope() });
      var a = state.st.header.blocks[state.st.blockIndex] || {};
      var c = (a.amountCols || [])[state.st.colIndex] || {};
      state.stSummaryKey = a.name + ' · ' + c.label;
      var ss = $('stSummary');
      ss.classList.remove('hidden');
      ss.className = 'note ok';
      ss.innerHTML = '读到 <b>' + state.st.parsed.items.length + '</b> 个项目（' +
        esc(state.st.sheetName) + ' · ' + esc(a.name || '') + ' · ' + esc(c.label || '') + '）';
    } else {
      state.cmp = null;
    }

    renderStats();
    renderDiff();
    renderUnmapped();
    renderDetail();
    renderMapPanel();
    $('genBtn').disabled = !(state.agg && state.cmp);
  }

  function cell(k, v, u, cls) {
    return '<div class="cell ' + (cls || '') + '"><div class="k">' + esc(k) + '</div>' +
      '<div class="v">' + esc(v) + '<span class="u">' + esc(u || '') + '</span></div></div>';
  }

  function renderStats() {
    var box = $('statsBox');
    if (!state.agg || !state.cmp) { box.innerHTML = ''; return; }
    var cmp = state.cmp;
    var diffAbs = cmp.diffs.reduce(function (s, d) { return s + Math.abs(d.diffCents); }, 0);
    box.innerHTML = '<div class="stats">' +
      cell('对上', cmp.matched.length, ' 项', 'ok') +
      cell('对不上', cmp.diffs.length, ' 项', cmp.diffs.length ? 'bad' : '') +
      cell('单边项目', cmp.onlyAgg.length + cmp.onlyTarget.length, ' 项', (cmp.onlyAgg.length + cmp.onlyTarget.length) ? 'warn' : '') +
      cell('未映射科目', state.agg.unmapped.length, ' 个', state.agg.unmapped.length ? 'warn' : '') +
      '</div>' +
      '<div class="stats">' +
      cell('差异金额（绝对值合计）', money(diffAbs), ' 元', diffAbs ? 'bad' : '') +
      cell('差异净额', money(cmp.diffTotalCents), ' 元', '') +
      cell('单边差额', money(cmp.unbalancedCents), ' 元', cmp.unbalancedCents ? 'warn' : '') +
      cell('覆盖报表项目', (function () { var sc = resolveScope(); return sc ? state.agg.items.filter(function (i) { return sc.indexOf(i.cls) >= 0; }).length : state.agg.items.length; })(), ' 个', '') +
      '</div>';
  }

  function renderDiff() {
    var box = $('diffBox');
    if (!state.cmp) { box.innerHTML = ''; return; }
    var cmp = state.cmp;
    var skippedN = (state.bs.stats && state.bs.stats.skippedRows) || 0;
    var detailBad = state.detailCmp && ((state.detailCmp.diffs || []).length || (state.detailCmp.failed || []).length);
    if (!cmp.diffs.length && !cmp.matched.length && !cmp.singleSideCount) { box.innerHTML = '<p class="muted">没有可比的项目。</p>'; return; }

    var html = '<h3>差异（' + cmp.diffs.length + ' 项）</h3>';
    if (!cmp.diffs.length && cmp.singleSideCount) {
      html += '<div class="note warn">已匹配的项目都一致，但还有 ' + cmp.singleSideCount +
        ' 项只在一边出现（见下）。这不等于勾平——项目名对不上或报表漏行都会这样。</div>';
    }
    if (cmp.diffs.length) {
      html += '<div class="scroll"><table class="data"><thead><tr>' +
        '<th>项目</th><th class="num">科目余额表</th><th class="num">对方表</th><th class="num">差异</th><th>构成科目</th>' +
        '</tr></thead><tbody>';
      cmp.diffs.forEach(function (d) {
        html += '<tr class="diff"><td>' + esc(d.item) + '</td>' +
          '<td class="num">' + money(d.leftCents) + '</td>' +
          '<td class="num">' + money(d.rightCents) + '</td>' +
          '<td class="num"><b>' + money(d.diffCents) + '</b></td>' +
          '<td>' + (d.accounts || []).slice(0, 4).map(function (a) {
            return esc(a.code + ' ' + a.name);
          }).join('<br>') + ((d.accounts || []).length > 4 ? '<br>…共 ' + d.accounts.length + ' 个' : '') + '</td></tr>';
      });
      html += '</tbody></table></div>';
    } else if (cmp.clean && !skippedN && !detailBad) {
      html += '<div class="note ok">所有项目都分毫不差。</div>';
    }

    if (cmp.matched.length) {
      html += '<details><summary class="muted">对上的 ' + cmp.matched.length + ' 项（点开看）</summary>' +
        '<div class="scroll"><table class="data"><thead><tr><th>项目</th><th class="num">科目余额表</th><th class="num">对方表</th></tr></thead><tbody>' +
        cmp.matched.map(function (d) {
          return '<tr><td>' + esc(d.item) + '</td><td class="num">' + money(d.leftCents) +
            '</td><td class="num">' + money(d.rightCents) + '</td></tr>';
        }).join('') + '</tbody></table></div></details>';
    }

    // 单边项目：一边有一边没有。不是「差异」，但同样会导致两边合计对不上，必须露出来
    var oneSide = cmp.onlyAgg.length + cmp.onlyTarget.length;
    if (oneSide) {
      html += '<h3>单边项目（' + oneSide + ' 项）</h3>' +
        '<div class="note warn">一边有一边没有：要么项目名称写法不同，要么对方表漏了这一行。' +
        '这类项目不算进「差异」，但会让两边合计对不上。</div>' +
        '<div class="scroll"><table class="data"><thead><tr>' +
        '<th>哪一边</th><th>项目</th><th class="num">金额</th></tr></thead><tbody>';
      cmp.onlyAgg.forEach(function (x) {
        html += '<tr><td>只有科目余额表有</td><td>' + esc(x.item) + '</td><td class="num">' + money(x.cents) + '</td></tr>';
      });
      cmp.onlyTarget.forEach(function (x) {
        html += '<tr><td>只有对方表有</td><td>' + esc(x.item) + '</td><td class="num">' + money(x.cents) + '</td></tr>';
      });
      html += '</tbody></table></div>';
    }
    if ((cmp.targetTotals || []).length) {
      html += '<p class="hint">报表里的 ' + cmp.targetTotals.length + ' 个合计行（' +
        cmp.targetTotals.slice(0, 4).map(function (t) { return esc(t.item); }).join('、') +
        '…）不参与逐项比对。</p>';
    }
    box.innerHTML = html;
  }

  function renderUnmapped() {
    var box = $('unmappedBox');
    if (!state.agg) { box.innerHTML = ''; return; }
    var un = state.agg.unmapped;
    var miss = state.agg.missing || [];

    // 期末读不出的科目：既没进项目、也没被当成 0，必须单独说
    var missHtml = '';
    if (miss.length) {
      missHtml = '<h3>期末余额读不出的科目（' + miss.length + ' 个）</h3>' +
        '<div class="note bad">这些科目<b>没有计入任何项目</b>——不是当 0 处理，是读不出来。' +
        '请回原表看这几行的期末余额格。</div>' +
        '<div class="scroll"><table class="data"><thead><tr>' +
        '<th>科目编码</th><th>科目名称</th><th class="num">原始行号</th></tr></thead><tbody>' +
        miss.slice(0, 30).map(function (m) {
          return '<tr><td>' + esc(m.code) + '</td><td>' + esc(m.name) +
            '</td><td class="num">' + esc(m.srcRow) + '</td></tr>';
        }).join('') + '</tbody></table></div>';
    }

    if (!un.length) {
      box.innerHTML = missHtml + (miss.length ? '' : '<div class="note ok">所有科目都归到报表项目了。</div>');
      return;
    }
    var withMoney = un.filter(function (u) { return (u.signedCents || 0) !== 0; });
    var html = '<h3>未映射科目（' + un.length + ' 个，其中有余额的 ' + withMoney.length + ' 个）</h3>' +
      '<div class="note warn">这些科目没有归到任何报表项目，所以汇总一定对不上。' +
      '在下面的框里给它们指定项目，指定过的会记住。</div>' +
      '<div class="scroll"><table class="data"><thead><tr>' +
      '<th>科目编码</th><th>科目名称</th><th class="num">期末余额</th><th class="num">原始行号</th></tr></thead><tbody>';
    un.slice(0, 50).forEach(function (u) {
      html += '<tr><td>' + esc(u.code) + '</td><td>' + esc(u.name) + '</td>' +
        '<td class="num">' + money(u.signedCents) + '</td><td class="num">' + esc(u.srcRow) + '</td></tr>';
    });
    html += '</tbody></table></div>';
    if (un.length > 50) html += '<p class="muted">只列前 50 个，完整的在导出的底稿里。</p>';
    box.innerHTML = missHtml + html;
  }

  function renderMapPanel() {
    var el = $('mapText');
    if (!el) return;
    el.value = Object.keys(state.overrides).map(function (k) {
      return k + '=' + state.overrides[k];
    }).join('\n');
    $('mapHint').textContent = Object.keys(state.overrides).length
      ? '已记住 ' + Object.keys(state.overrides).length + ' 条'
      : '还没有手工映射';
  }

  /* ================= 导出 ================= */

  function buildBox() {
    var today = new Date();
    function p2(n) { return String(n).padStart(2, '0'); }
    return {
      client: ($('clientName') && $('clientName').value.trim()) || '',
      dateText: today.getFullYear() + '-' + p2(today.getMonth() + 1) + '-' + p2(today.getDate()),
      bs: {
        label: state.bs.fileName + ' · ' + state.bs.sheetName,
        accountCount: state.bs.stats.accounts,
        leafCount: state.bs.stats.leaves,
        balanceDiffCents: state.bs.stats.bc.diffCents,
        identityFailed: state.bs.stats.ic.failed,
        identityChecked: state.bs.stats.ic.checked,
        inferred: state.bs.stats.ic.inferred,
        identityBad: (state.bs.stats.ic.bad || []).map(function (b) {
          return { code: b.code, name: b.name, srcRow: b.srcRow };
        }),
        skippedRows: state.bs.stats.skippedRows,
        level: state.bs.level || 'leaf',
        missing: (state.agg.missing || []).length
      },
      target: {
        label: state.st.fileName + ' · ' + state.st.sheetName,
        blockName: (state.st.header.blocks[state.st.blockIndex] || {}).name || '',
        colLabel: (((state.st.header.blocks[state.st.blockIndex] || {}).amountCols || [])[state.st.colIndex] || {}).label || ''
      },
      agg: state.agg,
      cmp: state.cmp,
      overrides: state.overrides,
      detail: (state.dt && state.detailCmp) ? {
        label: state.dt.fileName,
        col: state.dt.colIndex,
        colLabel: (function () {
          var o = $('dtCol').options[$('dtCol').selectedIndex];
          return o ? o.textContent : '';
        })(),
        diffs: state.detailCmp.diffs,
        matched: state.detailCmp.matched,
        onlyDetail: state.detailCmp.onlyDetail,
        onlyAgg: state.detailCmp.onlyAgg,
        failed: state.detailCmp.failed,
        diffTotalCents: state.detailCmp.diffTotalCents
      } : null
    };
  }

  /* ================= 事件 ================= */

  $('modeSimple').addEventListener('click', function () { setMode(true); });
  $('modeExpert').addEventListener('click', function () { setMode(false); });

  $('bsFile').addEventListener('change', function (e) {
    var f = e.target.files && e.target.files[0];
    if (f) loadBalanceSheet(f);
  });
  $('stFile').addEventListener('change', function (e) {
    var f = e.target.files && e.target.files[0];
    if (f) loadStatement(f);
  });

  $('bsSheet').addEventListener('change', function () {
    if (!state.bs) return;
    state.bs.sheetName = $('bsSheet').value;
    state.bs.rows = sheetRows(state.bs.wb, state.bs.sheetName);
    state.bs.header = Core.detectBalanceHeader(state.bs.rows);
    $('bsHeaderRow').value = state.bs.header.rowIndex + 1;
    parseBalance(); refresh();
  });

  $('bsHeaderRow').addEventListener('change', function () {
    if (!state.bs) return;
    var want = parseInt($('bsHeaderRow').value, 10) - 1;
    if (!(want >= 0)) return;
    var h = Core.detectBalanceHeader(state.bs.rows);
    h.rowIndex = want;
    state.bs.header = h;
    parseBalance(); refresh();
  });

  $('bsRedetect').addEventListener('click', function () {
    if (!state.bs) return;
    state.bs.header = Core.detectBalanceHeader(state.bs.rows);
    $('bsHeaderRow').value = state.bs.header.rowIndex + 1;
    parseBalance(); refresh();
  });

  $('bsLevel').addEventListener('change', function () {
    if (!state.bs) return;
    state.bs.level = $('bsLevel').value;
    refresh();
  });

  $('stSheet').addEventListener('change', function () {
    if (!state.st) return;
    state.st.sheetName = $('stSheet').value;
    state.st.rows = sheetRows(state.st.wb, state.st.sheetName);
    state.st.header = Core.detectStatementHeader(state.st.rows);
    state.st.blockIndex = 0; state.st.colIndex = 0;
    renderStatementSelectors();
    parseStatement(); refresh();
  });

  $('stBlock').addEventListener('change', function () {
    if (!state.st) return;
    state.st.blockIndex = parseInt($('stBlock').value, 10) || 0;
    state.st.colIndex = 0;
    renderColSelect();
    parseStatement(); refresh();
  });

  $('dtFile').addEventListener('change', function (e) {
    var f = e.target.files && e.target.files[0];
    if (f) loadDetail(f);
  });

  $('dtCol').addEventListener('change', function () {
    if (!state.dt) return;
    state.dt.colIndex = parseInt($('dtCol').value, 10);
    renderDetail();
  });

  $('stScope').addEventListener('change', function () {
    refresh();
  });

  $('stCol').addEventListener('change', function () {
    if (!state.st) return;
    state.st.colIndex = parseInt($('stCol').value, 10) || 0;
    parseStatement(); refresh();
  });

  $('mapApply').addEventListener('click', function () {
    var text = $('mapText').value || '';
    var map = {};
    text.split(/\r?\n/).forEach(function (line) {
      var m = /^\s*([0-9A-Za-z]{4,})\s*[=＝:：]\s*(.+?)\s*$/.exec(line);
      if (m) map[m[1]] = m[2];
    });
    state.overrides = map;
    try { window.localStorage.setItem(MAP_KEY, JSON.stringify(map)); } catch (e) { /* 忽略 */ }
    refresh();
  });

  $('mapForget').addEventListener('click', function () {
    state.overrides = {};
    try { window.localStorage.removeItem(MAP_KEY); } catch (e) { /* 忽略 */ }
    refresh();
  });

  $('genBtn').addEventListener('click', function () {
    if (!state.agg || !state.cmp) return;
    try {
      var out = Core.buildWorkbook(XLSX, buildBox());
      XLSX.writeFile(out.wb, out.fileName);
      $('genInfo').textContent = '已生成：' + out.fileName;
    } catch (err) {
      $('genInfo').textContent = '生成失败：' + err.message;
    }
  });

  /* ================= 示例数据 ================= */

  $('demoBtn').addEventListener('click', function () {
    // 造一张迷你科目余额表 + 一张试算表，结构照真实导出，金额是编的
    var bs = [
      ['账户汇总表', null, null, null, null, null, null, null, null, null, null, null, null, null, null, null],
      ['2026年1月至6月', null, null, null, null, null, null, null, null, null, null, null, null, null, null, null],
      ['账户代码', null, '账户名称', null, '期初余额', null, null, null, '期间发生额', null, '累计发生额', null, '期末余额', null, null, null],
      ['账户代码1', '账户代码2', '账户名称1', '账户名称2', '方向', '金额', '借方金额', '贷方金额', '借方金额', '贷方金额', '借方金额', '贷方金额', '方向', '金额', '借方金额', '贷方金额'],
      ['1002', null, '银行存款', null, '借', 1000000, null, null, null, null, null, null, '借', 1000000, null, null],
      ['1002', '01', '工行基本户', null, '借', 600000, null, null, null, null, null, null, '借', 600000, null, null],
      ['1002', '02', '建行一般户', null, '借', 400000, null, null, null, null, null, null, '借', 400000, null, null],
      ['1122', null, '应收账款', null, '借', 300000, null, null, null, null, null, null, '借', 300000, null, null],
      ['1601', null, '固定资产', null, '借', 500000, null, null, null, null, null, null, '借', 500000, null, null],
      ['1602', null, '累计折旧', null, '贷', 100000, null, null, null, null, null, null, '贷', 100000, null, null],
      ['2202', null, '应付账款', null, '贷', 200000, null, null, null, null, null, null, '贷', 200000, null, null],
      ['4001', null, '实收资本', null, '贷', 1560000, null, null, null, null, null, null, '贷', 1560000, null, null],
      ['1526', null, '某专项资产', null, '借', 60000, null, null, null, null, null, null, '借', 60000, null, null]
    ];
    var st = [
      ['试 算 表'],
      ['资   产', '行次', '审计前报表数', '负债及股东权益', '行次', '审计前报表数'],
      ['流动资产：', null, null, '流动负债：', null, null],
      ['　货币资金', 1, 1000000, '　应付账款', 30, 200000],
      ['　应收账款', 2, 300000, null, null, null],
      ['　固定资产', 3, 400000, null, null, null],
      ['非流动资产：', null, null, '所有者权益：', null, null],
      ['　其他非流动资产', 4, 60000, '　实收资本', 40, 1560000],
      ['资产总计', 9, 1760000, '负债和所有者权益总计', 50, 1760000]
    ];
    function toWb(rows, name) {
      var wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), name);
      return wb;
    }
    var wbBs = toWb(bs, '科目余额表');
    var wbSt = toWb(st, '试算表');

    state.bs = {
      fileName: '示例-科目余额表.xlsx', wb: wbBs, sheetName: '科目余额表',
      rows: bs, header: Core.detectBalanceHeader(bs), parsed: null, level: 'leaf'
    };
    parseBalance();
    $('bsSheet').innerHTML = '<option>科目余额表</option>';
    $('bsHeaderRow').value = state.bs.header.rowIndex + 1;

    state.st = {
      fileName: '示例-试算表.xlsx', wb: wbSt, sheetName: '试算表',
      rows: st, header: Core.detectStatementHeader(st), blockIndex: 0, colIndex: 0, parsed: null
    };
    $('stPanel').classList.remove('hidden');
    renderStatementSelectors();
    parseStatement();

    // 明细表：一个报表项目一张表，合计行中间带空格（真实导出就是这样）
    var wbDt = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wbDt, XLSX.utils.aoa_to_sheet([
      ['某某公司 货币资金明细表'],
      ['2026年6月30日'],
      ['单位', '账号', '金额'],
      ['工行基本户', '6222', 600000],
      ['建行一般户', '4321', 400000],
      ['合  计', null, 1000000]
    ]), '货币资金');
    XLSX.utils.book_append_sheet(wbDt, XLSX.utils.aoa_to_sheet([
      ['某某公司 应收账款明细表'],
      ['客户', '款项性质', '账面余额'],
      ['甲公司', '货款', 200000],
      ['乙公司', '货款', 100000],
      ['合  计', null, 300000]
    ]), '应收账款');
    setDetail('示例-明细表.xlsx', wbDt);

    refresh();
  });

  /* ================= 启动 ================= */

  $('verText').textContent = 'v' + Core.VERSION;
  try {
    var saved = window.localStorage.getItem(MAP_KEY);
    if (saved) state.overrides = JSON.parse(saved) || {};
  } catch (e) { /* 忽略 */ }
  (function initMode() {
    var m = null;
    try { m = window.localStorage.getItem('gouji.mode'); } catch (e) { /* 忽略 */ }
    setMode(m ? m === 'simple' : true);
  })();
  renderMapPanel();

})();

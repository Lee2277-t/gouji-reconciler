/**
 * 勾稽助手 · core.js
 *
 * 把「科目余额表」按科目编码汇总成「报表项目」，再与试算表/报表逐项比对。
 * 设计要点：
 *   1. 金额一律按「分」整数；展示才除 100。
 *   2. 映射表只认它认识的编码；认不出的科目必须报出来，绝不默认归零或猜。
 *   3. 每个报表项目都记住「由哪些科目汇总而来」，差异可以一路查到科目。
 *
 * 许可：AGPL-3.0
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.Core = factory();
}(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  var VERSION = '0.1.0';

  /* ============================================================
   * 一、基础工具
   * ============================================================ */

  function str(v) { return (v === null || v === undefined) ? '' : String(v); }

  function norm(v) {
    return str(v)
      .replace(/[\s\u3000]/g, '')
      .replace(/[（）()【】\[\]〔〕:：、,，*\/]/g, '')
      .toLowerCase();
  }

  /**
   * 金额 → 分（整数）。支持数字、"1,234.56"、"(800.00)"、"500.00-"、"1,000元"、全角括号。
   * @returns {{cents:number|null, reason:string|null, extra:boolean}}
   */
  function toCents(v) {
    if (v === null || v === undefined) return { cents: null, reason: 'empty', extra: false };

    if (typeof v === 'number') {
      if (!isFinite(v)) return { cents: null, reason: 'bad', extra: false };
      var sNum = String(v);
      if (/e/i.test(sNum)) return toCents(v.toFixed(2));
      var dot = sNum.indexOf('.');
      if (dot >= 0 && sNum.length - dot - 1 > 6) return toCents(v.toFixed(2));
      return toCents(sNum);
    }

    var raw = str(v).trim();
    if (!raw) return { cents: null, reason: 'empty', extra: false };

    var s = raw.replace(/[\s\u3000,，¥￥$]/g, '').replace(/(元|人民币|rmb)/gi, '');
    if (!s) return { cents: null, reason: 'empty', extra: false };
    s = s.replace(/（/g, '(').replace(/）/g, ')').replace(/[－—–]/g, '-');

    var neg = false;
    if (/^\(.*\)$/.test(s)) { neg = true; s = s.slice(1, -1); }
    if (/^[-−－—]/.test(s)) { neg = !neg; s = s.slice(1); }
    if (/[-−－—]$/.test(s)) { neg = !neg; s = s.slice(0, -1); }
    if (/^\+/.test(s)) { s = s.slice(1); }

    if (!/^\d*\.?\d*$/.test(s) || s === '' || s === '.') return { cents: null, reason: 'bad', extra: false };

    var m = /^(\d*)(?:\.(\d*))?$/.exec(s);
    if (!m) return { cents: null, reason: 'bad', extra: false };

    var ip = m[1] || '0';
    var frac = m[2] || '';
    var extra = false;
    var roundUp = false;
    if (frac.length > 2) {
      extra = true;
      roundUp = (frac.charCodeAt(2) - 48) >= 5;
      frac = frac.slice(0, 2);
    }
    frac = (frac + '00').slice(0, 2);

    var cents = Number(ip) * 100 + Number(frac);
    if (roundUp) cents += 1;
    if (!isFinite(cents)) return { cents: null, reason: 'bad', extra: false };
    return { cents: neg ? -cents : cents, reason: null, extra: extra };
  }

  function moneyText(cents) {
    if (cents === null || cents === undefined || !isFinite(cents)) return '';
    var neg = cents < 0;
    var a = Math.abs(Math.round(cents));
    var ip = Math.floor(a / 100);
    var fp = a % 100;
    var s = String(ip).replace(/\B(?=(\d{3})+(?!\d))/g, ',') + '.' + String(fp).padStart(2, '0');
    return (neg ? '-' : '') + s;
  }

  /* ============================================================
   * 二、科目编码 → 报表项目
   *
   * dir: 'D' 表示该科目余额在借方时对报表项目是正贡献
   *      'C' 表示贷方余额为正贡献
   * sign: -1 表示备抵科目（累计折旧、减值准备等），汇总时要减掉
   *
   * 只收录《企业会计准则》通用科目。非常用科目、企业自定义科目一律不猜，
   * 由 unmapped 报出来让人手工指定 —— 猜错比不猜更危险。
   * ============================================================ */

  var CHART = [
    /* ---- 资产 ---- */
    ['1001', '货币资金', 'D'], ['1002', '货币资金', 'D'], ['1012', '货币资金', 'D'],
    ['1031', '其他流动资产', 'D'],                       // 存出保证金
    ['1101', '交易性金融资产', 'D'], ['1103', '衍生金融资产', 'D'],
    ['1111', '其他流动资产', 'D'],                       // 买入返售金融资产
    ['1201', '其他应收款', 'D'],                          // 应收代位追偿款
    ['1321', '其他流动资产', 'D'],                       // 代理业务资产
    ['1121', '应收票据', 'D'], ['1122', '应收账款', 'D'], ['1123', '预付款项', 'D'],
    ['1131', '应收股利', 'D'], ['1132', '应收利息', 'D'], ['1221', '其他应收款', 'D'],
    ['1231', '应收账款', 'C', -1],                       // 坏账准备：余额在贷方，对资产是减项
    ['1401', '存货', 'D'], ['1402', '存货', 'D'], ['1403', '存货', 'D'], ['1404', '存货', 'D'],
    ['1405', '存货', 'D'], ['1406', '存货', 'D'], ['1407', '存货', 'D'], ['1408', '存货', 'D'],
    ['1411', '存货', 'D'], ['1421', '存货', 'D'], ['1431', '存货', 'D'], ['1441', '存货', 'D'],
    ['1451', '存货', 'D'], ['1461', '存货', 'D'],
    ['1471', '存货', 'C', -1],                            // 存货跌价准备：余额在贷方
    ['1501', '债权投资', 'D'], ['1502', '其他债权投资', 'D'], ['1503', '其他权益工具投资', 'D'],
    ['1511', '长期股权投资', 'D'], ['1521', '投资性房地产', 'D'],
    ['1522', '投资性房地产', 'C', -1], ['1523', '投资性房地产', 'C', -1],
    ['1531', '长期应收款', 'D'], ['1541', '其他非流动资产', 'D'],   // 存出资本保证金
    ['1601', '固定资产', 'D'], ['1602', '固定资产', 'C', -1], ['1603', '固定资产', 'C', -1],
    ['1604', '在建工程', 'D'], ['1605', '工程物资', 'D'], ['1606', '固定资产清理', 'D'],
    ['1611', '使用权资产', 'D'], ['1621', '生产性生物资产', 'D'],
    ['1622', '生产性生物资产', 'C', -1], ['1623', '生产性生物资产', 'D'],
    ['1701', '无形资产', 'D'], ['1702', '无形资产', 'C', -1], ['1703', '无形资产', 'C', -1],
    ['1704', '开发支出', 'D'], ['1711', '商誉', 'D'], ['1801', '长期待摊费用', 'D'],
    ['1811', '递延所得税资产', 'D'], ['1821', '递延所得税资产', 'D'],
    ['1901', '其他流动资产', 'D'],

    /* ---- 负债 ---- */
    ['2001', '短期借款', 'C'], ['2002', '其他流动负债', 'C'],      // 存入保证金
    ['2101', '交易性金融负债', 'C'],
    ['2201', '应付票据', 'C'], ['2202', '应付账款', 'C'], ['2203', '预收款项', 'C'],
    ['2205', '合同负债', 'C'], ['2211', '应付职工薪酬', 'C'], ['2221', '应交税费', 'C'],
    ['2231', '应付利息', 'C'], ['2232', '应付股利', 'C'], ['2241', '其他应付款', 'C'],
    ['2401', '递延收益', 'C'],
    ['2501', '长期借款', 'C'], ['2502', '应付债券', 'C'],
    ['2601', '保险合同准备金', 'C'], ['2602', '保险合同准备金', 'C'],
    ['2701', '长期应付款', 'C'], ['2702', '长期应付款', 'D', -1],   // 未确认融资费用（减项）
    ['2711', '专项应付款', 'C'],
    ['2801', '预计负债', 'C'], ['2901', '递延所得税负债', 'C'],

    /* ---- 所有者权益 ---- */
    ['4001', '实收资本（或股本）', 'C'], ['4002', '资本公积', 'C'], ['4003', '其他综合收益', 'C'],
    ['4101', '盈余公积', 'C'], ['4103', '未分配利润', 'C'], ['4104', '未分配利润', 'C'],
    ['4201', '库存股', 'C', -1],

    /* ---- 成本类：期末通常并入存货 ---- */
    ['5001', '存货', 'D'], ['5101', '存货', 'D'], ['5201', '存货', 'D'],
    ['5301', '开发支出', 'D'], ['5401', '存货', 'D'],

    /* ---- 损益 ---- */
    ['6001', '营业收入', 'C'], ['6011', '营业收入', 'C'], ['6051', '营业收入', 'C'],
    ['6061', '汇兑收益', 'C'], ['6101', '公允价值变动收益', 'C'], ['6111', '投资收益', 'C'],
    ['6301', '营业外收入', 'C'],
    ['6401', '营业成本', 'D'], ['6402', '营业成本', 'D'], ['6405', '税金及附加', 'D'],
    ['6601', '销售费用', 'D'], ['6602', '管理费用', 'D'], ['6603', '财务费用', 'D'],
    ['6701', '资产减值损失', 'D'], ['6702', '信用减值损失', 'D'],
    ['6711', '营业外支出', 'D'], ['6801', '所得税费用', 'D'],
    ['6901', '未分配利润', 'C']                      // 以前年度损益调整，直接进未分配利润
  ];

  var CHART_MAP = {};
  CHART.forEach(function (r) {
    var entry = { code: r[0], item: r[1], dir: r[2], sign: r[3] === undefined ? 1 : r[3] };
    if (!CHART_MAP[r[0]]) CHART_MAP[r[0]] = entry;
  });

  /**
   * 从科目编码（可能带空格、点、横线，如 "1122.01" "1122 01"）取出四位前缀。
   * 取不出就返回 ''，由上层报「编码不认识」。
   */
  function codePrefix(code) {
    var s = str(code).replace(/[^\d]/g, '');
    if (s.length < 4) return '';
    return s.slice(0, 4);
  }

  /** 科目 → 报表项目；认不出返回 null（不猜） */
  function mapAccount(code) {
    var p = codePrefix(code);
    if (!p) return null;
    return CHART_MAP[p] || null;
  }

  /**
   * 科目余额 → 该科目对报表项目的「带符号贡献」（分）。
   * dir=D：借-贷；dir=C：贷-借。再乘备抵科目的 sign。
   */
  function contribution(account, dir) {
    var dr = account.debitCents || 0;
    var cr = account.creditCents || 0;
    var v = dir === 'C' ? (cr - dr) : (dr - cr);
    return v;
  }

  /* ============================================================
   * 三、读科目余额表
   *
   * 真实结构（金蝶/用友/ERP 导出都差不多）：
   *   第 1 行  账户汇总表
   *   第 2 行  2026年1月至6月
   *   第 3 行  账户代码 | 账户名称 | 期初余额 | 期间发生额 | 累计发生额     ← 一级
   *   第 4 行  代码1|代码2|名称1|名称2|方向|金额|借方|贷方|借方|贷方|借方|贷方 ← 二级
   *   第 5 行起 数据
   * ============================================================ */

  // 一级表头 → 内部字段前缀
  var GROUP_ALIASES = [
    { key: 'opening', words: ['期初余额', '期初', '年初余额', '上期期末', '期初数'] },
    { key: 'period', words: ['期间发生额', '本期发生额', '本期发生', '本期数', '发生额', '本期'] },
    { key: 'year', words: ['累计发生额', '本年累计', '本年发生额', '累计发生', '年度累计'] },
    { key: 'closing', words: ['期末余额', '期末', '本期期末', '期末数'] }
  ];

  function pickGroup(text) {
    var t = norm(text);
    if (!t) return '';
    var best = '', bestLen = 0;
    GROUP_ALIASES.forEach(function (g) {
      g.words.forEach(function (w) {
        var n = norm(w);
        if (t.indexOf(n) >= 0 && n.length > bestLen) { best = g.key; bestLen = n.length; }
      });
    });
    return best;
  }

  /** 二级表头 → 借/贷/金额/方向 */
  function pickLeaf(text) {
    var t = norm(text);
    if (!t) return '';
    if (t === '借方' || t === '借方金额' || t === '借') return 'debit';
    if (t === '贷方' || t === '贷方金额' || t === '贷') return 'credit';
    if (t === '方向' || t === '借贷方向' || t === '借贷') return 'dir';
    if (t === '金额' || t === '余额' || t === '余额金额') return 'amount';
    return '';
  }

  /**
   * 找科目余额表的表头，并拼出「一级.二级」的列映射。
   * 支持一级表头只在第一列写、右边合并单元格留空的情况。
   * @returns {{rowIndex, span, columns:{opening:{debit,credit,amount,dir}, ...}, problems:[]}}
   */
  function detectBalanceHeader(rows) {
    var limit = Math.min(rows.length, 30);
    var problems = [];

    function scoreRow(r) {
      var cells = rows[r] || [];
      var s = 0;
      for (var c = 0; c < cells.length; c++) {
        var t = norm(cells[c]);
        if (!t) continue;
        if (/账户代码|科目编码|科目代码|账户编号|科目编号|编码/.test(t)) s += 3;
        if (/账户名称|科目名称|科目|名称/.test(t)) s += 3;
        if (pickGroup(cells[c])) s += 2;
        if (/借方|贷方|方向|余额/.test(t)) s += 1;
      }
      return s;
    }

    // 判断某行是不是「二级表头」：几乎全是 借/贷/金额/方向 这种短词
    function leafRatio(r) {
      var cells = rows[r] || [];
      var total = 0, hits = 0;
      for (var i = 0; i < cells.length; i++) {
        if (!str(cells[i]).trim()) continue;
        total++;
        if (pickLeaf(cells[i])) hits++;
      }
      return total >= 2 ? hits / total : 0;
    }
    function groupHits(r) {
      var cells = rows[r] || [];
      var n = 0;
      for (var i = 0; i < cells.length; i++) if (pickGroup(cells[i])) n++;
      return n;
    }

    var bestR = -1, bestScore = 0, bestSpan = 1;
    for (var r = 0; r < limit; r++) {
      var sc = scoreRow(r);
      // 一级行单独看得分不高（「期初余额」这种只算 2 分），
      // 但它的下一行如果是满屏 借/贷，就该把这两行当成一个两级表头整体来判
      var isPair = groupHits(r) >= 2 && leafRatio(r + 1) >= 0.5;
      var total = sc + (isPair ? 10 : 0);
      if (total > bestScore) { bestScore = total; bestR = r; bestSpan = isPair ? 2 : 1; }
    }
    if (bestR < 0 || bestScore < 4) {
      return { rowIndex: -1, span: 1, columns: {}, problems: ['找不到科目余额表的表头（需要有「科目编码」「科目名称」「余额」这类列）'] };
    }

    var leafRow = rows[bestR + 1] || [];
    var span = bestSpan;

    var columns = { opening: {}, period: {}, year: {}, closing: {} };
    var codeCol = -1, codeCol2 = -1, nameCol = -1, dirCol = -1, amountCol = -1;
    var topRow = rows[bestR] || [];
    var subRow = span === 2 ? leafRow : [];

    // 一级表头向右填充（合并单元格只在首列有值）
    var lastGroup = '';
    var colCount = Math.max(topRow.length, subRow.length);
    for (var c2 = 0; c2 < colCount; c2++) {
      var g = pickGroup(topRow[c2]);
      if (g) lastGroup = g;

      var leaf = span === 2 ? pickLeaf(subRow[c2]) : '';
      if (!leaf && span === 1) leaf = pickLeaf(topRow[c2]);

      var tc = norm(topRow[c2]);
      var ts = span === 2 ? norm(subRow[c2]) : '';
      // 编码/名称常常拆成两列且写在二级表头（账户代码1 / 账户代码2），
      // 一级行只有「账户代码」一个词，所以两级都要看
      if (/账户代码|科目编码|科目代码|账户编号|科目编号|编码/.test(tc) || /代码|编码/.test(ts)) {
        if (codeCol < 0) codeCol = c2;
        else if (codeCol2 < 0 && c2 > codeCol) codeCol2 = c2;
      } else if (nameCol < 0 && (/账户名称|科目名称|名称/.test(tc) || /名称/.test(ts))) {
        nameCol = c2;
      }

      // 方向列可能每组各有一个（期初有、期末也有），必须按组记，不能只留全局一个
      if (leaf === 'dir') {
        if (lastGroup) columns[lastGroup].dir = c2;
        if (dirCol < 0) dirCol = c2;
        continue;
      }
      if (leaf === 'amount' && !lastGroup) { amountCol = c2; continue; }
      if (!lastGroup || !leaf) continue;

      if (!columns[lastGroup][leaf] && columns[lastGroup][leaf] !== 0) {
        columns[lastGroup][leaf] = c2;
      }
    }

    if (codeCol < 0) problems.push('没找到科目编码列');
    if (nameCol < 0) problems.push('没找到科目名称列');

    var hasAnyAmount = ['opening', 'period', 'year', 'closing'].some(function (k) {
      return columns[k].debit !== undefined || columns[k].credit !== undefined ||
             columns[k].amount !== undefined;
    });
    if (!hasAnyAmount) problems.push('没找到金额列（借方/贷方/金额都没有）');

    return {
      rowIndex: bestR, span: span, codeCol: codeCol, codeCol2: codeCol2, nameCol: nameCol,
      dirCol: dirCol, amountCol: amountCol, columns: columns, problems: problems
    };
  }

  /** Excel 行号从 1 开始；rows[0] 是第 firstExcelRow 行 */
  function parseBalanceRows(rows, header, opts) {
    opts = opts || {};
    var firstExcelRow = opts.firstExcelRow || 1;
    var span = header.span || 1;
    var cols = header.columns || {};
    var accounts = [];
    var skipped = [];
    var problems = [];

    function centsAt(group, kind, row) {
      var c = cols[group] && cols[group][kind];
      if (c === undefined) return null;
      var v = toCents(row[c]);
      return v.cents;
    }

    for (var r = header.rowIndex + span; r < rows.length; r++) {
      var row = rows[r] || [];
      var excelRow = firstExcelRow + r;

      var code1 = header.codeCol >= 0 ? str(row[header.codeCol]).trim() : '';
      var code2 = header.codeCol2 >= 0 ? str(row[header.codeCol2]).trim() : '';
      // 编码可能拆成两列（4001 + 01），拼起来才是真正的科目编码；只有第一段的是上级科目
      var code = code2 ? (code1 + code2) : code1;
      var name = header.nameCol >= 0 ? str(row[header.nameCol]).trim() : '';

      // 全空行
      if (!code && !name && !row.join('').trim()) { skipped.push({ srcRow: excelRow, reason: '空行', code: '', name: '' }); continue; }

      // 合计/汇总行：没有编码但有「合计」字样
      if (!code && /合计|总计|小计|汇总/.test(name)) {
        skipped.push({ srcRow: excelRow, reason: '汇总行', code: code, name: name });
        continue;
      }
      if (!code || codePrefix(code) === '') {
        // 有名称没编码、或编码里没有四位数字（如「总计」）：是标题/说明行，不是科目
        skipped.push({ srcRow: excelRow, reason: name ? '无科目编码' : '空行', code: code, name: name });
        continue;
      }

      function dirOf(group) {
        var c = cols[group] && cols[group].dir;
        if (c === undefined) return '';
        return str(row[c]).trim();
      }
      /** 某组有没有出现（有方向列也算，因为「方向=平」就代表余额为 0） */
      function hasGroup(group) {
        var g = cols[group] || {};
        return g.amount !== undefined || g.debit !== undefined ||
               g.credit !== undefined || g.dir !== undefined;
      }
      /**
       * 一组的带符号余额（借方为正）。
       * 关键：金额空着不代表读不到 —— 国产 ERP 常把「余额为 0」写成 方向=平、金额空。
       * 但方向写着借/贷、金额却空的情况多半是导出没填，这时返回 null，交给上层去推。
       */
      function signed(group) {
        if (!hasGroup(group)) return null;
        var amt = centsAt(group, 'amount', row);
        var dr = centsAt(group, 'debit', row);
        var cr = centsAt(group, 'credit', row);
        if (amt !== null) return /贷/.test(dirOf(group)) ? -Math.abs(amt) : Math.abs(amt);
        if (dr !== null || cr !== null) return (dr || 0) - (cr || 0);
        var d = dirOf(group);
        if (d && !/借|贷/.test(d)) return 0;      // 「平」「无」等 = 零余额
        if (!d) return 0;                          // 没有方向列，也没有金额 = 空行
        return null;                               // 写着借/贷却没金额：推出来再说
      }

      var closingInferred = false;
      var opDir = dirOf('opening');
      var clDir = dirOf('closing');
      var opening = signed('opening');

      var pDr = centsAt('period', 'debit', row);
      var pCr = centsAt('period', 'credit', row);
      var yDr = centsAt('year', 'debit', row);
      var yCr = centsAt('year', 'credit', row);

      // 期末：有期末数就用；没填就用 期初 + 期间借 - 期间贷
      var closing = signed('closing');
      if (closing === null && opening !== null) {
        closing = opening + (pDr || 0) - (pCr || 0);
        closingInferred = true;
      }

      accounts.push({
        srcRow: excelRow,
        code: code,
        codeLevel1: code1,
        isDetail: !!code2,
        name: name,
        rawDir: opDir,
        openingCents: opening,
        periodDebitCents: pDr,
        periodCreditCents: pCr,
        yearDebitCents: yDr,
        yearCreditCents: yCr,
        closingCents: closing,
        closingInferred: closingInferred,
        // 汇总时按「借方余额为正」统一：借方科目看 (借-贷)，贷方科目自然为负
        debitCents: closing !== null && closing > 0 ? closing : 0,
        creditCents: closing !== null && closing < 0 ? -closing : 0
      });
    }

    // 一笔都没有，说明表头认错了
    if (!accounts.length) problems.push('表头下面没有读到科目行');

    return { accounts: accounts, skipped: skipped, problems: problems, header: header };
  }

  /* ============================================================
   * 四、汇总：科目余额表 → 报表项目
   * ============================================================ */

  /** 报表项目按科目首位数字归类：资产 / 负债 / 权益 / 成本 / 损益 */
  var CLASSES = { '1': '资产', '2': '负债', '3': '共同', '4': '权益', '5': '成本', '6': '损益' };

  function classOfCode(code) {
    var p = codePrefix(code);
    return p ? (CLASSES[p.charAt(0)] || '其他') : '其他';
  }

  /**
   * @param accounts [{code,name,debitCents,creditCents,srcRow}]
   * @param opts { overrides: {code4: itemName} }  手工映射（优先于内置）
   * @returns {{items:[{item,cents,cls,accounts:[]}], unmapped:[], total:number}}
   */
  function aggregateItems(accounts, opts) {
    opts = opts || {};
    var overrides = opts.overrides || {};
    var byItem = {};
    var order = [];
    var unmapped = [];
    var total = 0;

    accounts.forEach(function (a) {
      var ov = overrides[codePrefix(a.code)];
      var entry = null;
      if (ov) {
        // 手工映射：方向跟科目大类走（1/5/6 借方，2/3/4 贷方，6xxx 收入类特殊）
        var p = codePrefix(a.code);
        var dir = (p.charAt(0) === '2' || p.charAt(0) === '4') ? 'C'
          : (p.charAt(0) === '3' ? 'C' : 'D');
        entry = { item: ov, dir: dir, sign: 1, manual: true };
      } else {
        entry = mapAccount(a.code);
      }

      if (!entry) {
        var v0 = (a.debitCents || 0) - (a.creditCents || 0);
        unmapped.push({
          code: a.code, name: a.name, srcRow: a.srcRow,
          cents: Math.abs(v0), signedCents: v0
        });
        total += v0;
        return;
      }

      var v = contribution(a, entry.dir) * entry.sign;
      if (!byItem[entry.item]) {
        byItem[entry.item] = { item: entry.item, cents: 0, accounts: [], clsCount: {} };
        order.push(entry.item);
      }
      var cls = classOfCode(a.code);
      byItem[entry.item].clsCount[cls] = (byItem[entry.item].clsCount[cls] || 0) + 1;
      byItem[entry.item].cents += v;
      byItem[entry.item].accounts.push({
        code: a.code, name: a.name, cents: v,
        srcRow: a.srcRow, manual: !!entry.manual
      });
      total += v;
    });

    var items = order.map(function (k) {
      var it = byItem[k];
      it.accounts.sort(function (x, y) { return Math.abs(y.cents) - Math.abs(x.cents); });
      // 同一个项目可能由跨类科目凑成（如「未分配利润」含损益类），取占比最大的那一类
      var bestCls = '', bestN = -1;
      Object.keys(it.clsCount).forEach(function (c) {
        if (it.clsCount[c] > bestN) { bestN = it.clsCount[c]; bestCls = c; }
      });
      it.cls = bestCls;
      delete it.clsCount;
      return it;
    });
    items.sort(function (x, y) { return Math.abs(y.cents) - Math.abs(x.cents); });
    unmapped.sort(function (x, y) { return Math.abs(y.signedCents) - Math.abs(x.signedCents); });

    return { items: items, unmapped: unmapped, total: total };
  }

  /* ============================================================
   * 五、比对：汇总结果 vs 试算表/报表
   * ============================================================ */

  /** 项目名归一：全角括号、空格、冒号都不算差别 */
  function normItem(name) {
    return norm(name).replace(/^[-—－]+/, '');
  }

  /**
   * @param agg  aggregateItems 的结果
   * @param target [{item, cents}]  试算表或报表里的项目
   * @param opts { toleranceCents, cls: 只比这一类（字符串或数组） }
   * @returns {{matched:[], diffs:[], onlyAgg:[], onlyTarget:[], sameCount:number}}
   */
  function compareItems(agg, target, opts) {
    opts = opts || {};
    var tol = opts.toleranceCents || 0;
    // 只比某一类：拿「资产」栏去比全部科目，会把负债权益全算成单边项目，没有意义
    var scope = opts.cls || '';

    var right = {};
    var rightOrder = [];
    target.forEach(function (t) {
      var k = normItem(t.item);
      if (!k) return;
      if (right[k] === undefined) {
        right[k] = { item: t.item, cents: 0, rows: [] };
        rightOrder.push(k);
      }
      right[k].cents += (t.cents || 0);
      right[k].rows.push(t.srcRow);
    });

    var left = {};
    var leftOrder = [];
    agg.items.forEach(function (it) {
      if (scope) {
        var keep = Array.isArray(scope) ? scope.indexOf(it.cls) >= 0 : it.cls === scope;
        if (!keep) return;
      }
      var k = normItem(it.item);
      if (left[k] === undefined) { left[k] = it; leftOrder.push(k); }
      else { left[k].cents += it.cents; }     // 同名项目合并（如「存货」来自多组科目）
    });

    var matched = [], diffs = [], onlyAgg = [], onlyTarget = [];
    leftOrder.forEach(function (k) {
      var L = left[k];
      var R = right[k];
      if (!R) {
        if (!L.cents && !opts.keepZero) return;
        onlyAgg.push({ item: L.item, cents: L.cents, accounts: L.accounts });
        return;
      }
      var d = L.cents - R.cents;
      var row = {
        item: L.item, leftCents: L.cents, rightCents: R.cents, diffCents: d,
        accounts: L.accounts, targetRows: R.rows
      };
      if (Math.abs(d) <= tol) matched.push(row);
      else diffs.push(row);
    });
    rightOrder.forEach(function (k) {
      if (left[k]) return;
      // 报表模板里大量零值行（拆出资金、衍生金融资产…），没金额就不算差异
      if (!right[k].cents && !opts.keepZero) return;
      onlyTarget.push({ item: right[k].item, cents: right[k].cents, rows: right[k].rows });
    });

    diffs.sort(function (x, y) { return Math.abs(y.diffCents) - Math.abs(x.diffCents); });
    onlyAgg.sort(function (x, y) { return Math.abs(y.cents) - Math.abs(x.cents); });
    onlyTarget.sort(function (x, y) { return Math.abs(y.cents) - Math.abs(x.cents); });

    return {
      matched: matched, diffs: diffs, onlyAgg: onlyAgg, onlyTarget: onlyTarget,
      sameCount: matched.length,
      diffTotalCents: diffs.reduce(function (s, x) { return s + x.diffCents; }, 0),
      // 单边项目也是差异：左边有右边没有，说明要么项目名对不上，要么报表漏了
      unbalancedCents: onlyAgg.reduce(function (s, x) { return s + x.cents; }, 0) -
                       onlyTarget.reduce(function (s, x) { return s + x.cents; }, 0)
    };
  }

  /**
   * 取末级科目：没有任何别的科目以它的编码为前缀。
   * 科目余额表里上级科目和下级科目同时存在，两边都加会把金额翻倍。
   * 报表项目应该由末级科目汇总而来。
   */
  function leafAccounts(accounts) {
    var codes = accounts.map(function (a) { return str(a.code).replace(/[^\d]/g, ''); });
    return accounts.filter(function (a, i) {
      var c = codes[i];
      if (!c) return true;
      for (var j = 0; j < codes.length; j++) {
        if (j === i) continue;
        var o = codes[j];
        if (o && o !== c && o.indexOf(c) === 0) return false;
      }
      return true;
    });
  }

  /**
   * 借贷平衡自检：一张正常的科目余额表，末级科目借方合计应等于贷方合计。
   * 不平就说明要么漏了科目、要么某列读错了 —— 先别急着往下勾稽。
   */
  function balanceCheck(accounts) {
    var codes = accounts.map(function (a) { return str(a.code).replace(/[^\d]/g, ''); });
    var leaves = accounts.filter(function (a, i) {
      var c = codes[i];
      if (!c) return true;
      return !codes.some(function (o, j) { return j !== i && o && o !== c && o.indexOf(c) === 0; });
    });
    var dr = 0, cr = 0;
    leaves.forEach(function (a) {
      var v = a.closingCents;
      if (v === null || v === undefined) return;
      if (v > 0) dr += v; else cr += -v;
    });
    return {
      leafless: leaves.length,
      parentCount: accounts.length - leaves.length,
      debitCents: dr,
      creditCents: cr,
      diffCents: dr - cr,
      balanced: dr === cr
    };
  }

  /**
   * 逐科目验算：期初 + 本期借方 - 本期贷方 = 期末。
   * 这条恒等式不成立，说明这张导出本身有问题（列错位、口径混用），
   * 比「借贷合计不平」更能定位到具体哪一行。
   */
  function identityCheck(accounts) {
    var bad = [];
    var checked = 0;
    accounts.forEach(function (a) {
      if (a.openingCents === null || a.closingCents === null) return;
      checked++;
      var calc = a.openingCents + (a.periodDebitCents || 0) - (a.periodCreditCents || 0);
      if (Math.abs(calc - a.closingCents) > 1) {
        bad.push({
          code: a.code, name: a.name, srcRow: a.srcRow,
          openingCents: a.openingCents, periodDebitCents: a.periodDebitCents,
          periodCreditCents: a.periodCreditCents, closingCents: a.closingCents,
          expectedCents: calc, diffCents: a.closingCents - calc
        });
      }
    });
    return { checked: checked, bad: bad, failed: bad.length };
  }

  /* ============================================================
   * 五、读试算表 / 报表
   *
   * 两种版式都要认：
   *   1) 竖清单列：项目 | 金额
   *   2) 左右并排：资产 | 行次 | 金额 || 负债及所有者权益 | 行次 | 金额（大华/事务所模板常见）
   * ============================================================ */

  var ITEM_HEAD_RE = /^(项目|资产|负债|所有者权益|股东权益|负债及|负债和|利润表项目|科目|内容|摘要)$/;

  function looksLikeItemHead(text) {
    var t = norm(text);
    if (!t) return false;
    if (/^项目$|^资产$|^负债|^所有者权益$|^股东权益$|^负债及|^负债和|^科目$|^内容$/.test(t)) return true;
    return false;
  }

  function looksLikeAmountHead(text) {
    var t = norm(text);
    if (!t) return false;
    return /审计前|审计后|调整后|审定|未审|账面|期末余额|期初余额|本期|本年|金额|余额|借方|贷方/.test(t);
  }

  /**
   * 找试算表/报表的表头行，并把「项目列 → 它右边的金额列」切成若干块。
   * @returns {{rowIndex, blocks:[{itemCol, name, amountCols:[{col,label}]}], problems:[]}}
   */
  function detectStatementHeader(rows) {
    var limit = Math.min(rows.length, 40);
    var problems = [];

    // 表头行：同一行里出现 ≥1 个项目列 且 ≥1 个金额列
    var bestR = -1, bestScore = 0;
    for (var r = 0; r < limit; r++) {
      var cells = rows[r] || [];
      var items = 0, amts = 0;
      for (var c = 0; c < cells.length; c++) {
        if (!str(cells[c]).trim()) continue;
        if (looksLikeItemHead(cells[c])) items++;
        else if (looksLikeAmountHead(cells[c])) amts++;
      }
      var sc = items * 3 + amts;
      if (items >= 1 && amts >= 1 && sc > bestScore) { bestScore = sc; bestR = r; }
    }
    if (bestR < 0) {
      return { rowIndex: -1, blocks: [], problems: ['找不到试算表/报表的表头（需要有「项目」「金额」这类列）'] };
    }

    var head = rows[bestR] || [];
    // 项目列的位置
    var itemCols = [];
    for (var c2 = 0; c2 < head.length; c2++) {
      if (looksLikeItemHead(head[c2])) itemCols.push(c2);
    }
    if (!itemCols.length) problems.push('没找到项目列');

    // 每个项目列到下一个项目列之间的金额列，归这块
    var blocks = itemCols.map(function (col, i) {
      var next = i + 1 < itemCols.length ? itemCols[i + 1] : head.length;
      var amountCols = [];
      for (var c3 = col + 1; c3 < next; c3++) {
        var label = str(head[c3]).trim();
        if (!label) continue;
        if (/行次|行号|附注|备注/.test(norm(label))) continue;
        amountCols.push({ col: c3, label: label });
      }
      return {
        itemCol: col,
        name: str(head[col]).trim() || ('第 ' + (col + 1) + ' 列'),
        amountCols: amountCols
      };
    });

    return { rowIndex: bestR, blocks: blocks, problems: problems };
  }

  /**
   * 读出一个块里的项目与某一列金额。
   * @param colIndex 用第几个金额列（默认第 1 个）
   */
  function parseStatementBlock(rows, header, blockIndex, colIndex, opts) {
    opts = opts || {};
    var firstExcelRow = opts.firstExcelRow || 1;
    var block = header.blocks[blockIndex];
    if (!block) return { items: [], skipped: [], problems: ['没有这一块'] };
    var ac = block.amountCols[colIndex || 0];
    if (!ac) return { items: [], skipped: [], problems: ['这一块没有金额列'] };

    var items = [], skipped = [];
    for (var r = header.rowIndex + 1; r < rows.length; r++) {
      var row = rows[r] || [];
      var name = str(row[block.itemCol]).trim();
      var excelRow = firstExcelRow + r;
      if (!name) { continue; }
      // 分区标题（「流动资产：」「流动负债：」）没有金额，跳过
      var parsed = toCents(row[ac.col]);
      if (parsed.cents === null) {
        if (/[:：]$/.test(name) || /^(流动资产|非流动资产|流动负债|非流动负债|所有者权益|股东权益)/.test(norm(name))) {
          skipped.push({ srcRow: excelRow, reason: '分区标题', item: name });
        } else if (parsed.reason === 'bad') {
          skipped.push({ srcRow: excelRow, reason: '金额读不出', item: name });
        }
        continue;
      }
      items.push({
        item: name.replace(/^[\s\u3000]+/, ''),
        cents: parsed.cents,
        srcRow: excelRow,
        col: ac.col
      });
    }
    return { items: items, skipped: skipped, problems: [] };
  }

  /* ============================================================
   * 六、输出：勾稽底稿
   * ============================================================ */

  /**
   * @param box {
   *   client, dateText,
   *   bs: {label, accountCount, leafCount, balanceDiffCents, identityFailed, skipped},
   *   target: {label, blockName, colLabel},
   *   agg, cmp, overrides
   * }
   */
  function buildSheets(box) {
    var sheets = [];

    /* ---- 勾稽表 ---- */
    var rows = [];
    rows.push(['科目余额表 ↔ ' + (box.target ? box.target.label : '试算表') + ' 勾稽表', '', '', '', '被审计单位：' + (box.client || '')]);
    rows.push(['编制日期：' + (box.dateText || ''), '', '', '', '金额单位：元']);
    rows.push([]);
    rows.push(['序号', '项目', '科目余额表汇总', '对方表', '差异', '构成科目数']);
    var seq = 0;
    var moneyCols = [2, 3, 4];

    function pushRow(item, leftCents, rightCents, extra) {
      seq++;
      rows.push([seq, item, leftCents / 100, rightCents === null ? '' : rightCents / 100,
        (leftCents - (rightCents || 0)) / 100, extra || '']);
    }

    if (box.cmp) {
      box.cmp.diffs.forEach(function (d) { pushRow(d.item, d.leftCents, d.rightCents, d.accounts ? d.accounts.length : ''); });
      box.cmp.matched.forEach(function (d) { pushRow(d.item, d.leftCents, d.rightCents, d.accounts ? d.accounts.length : ''); });
    }
    rows.push([]);
    rows.push(['差异合计', '', '', '', ((box.cmp ? box.cmp.diffTotalCents : 0) / 100), '']);
    sheets.push({ name: '勾稽表', rows: rows, moneyCols: moneyCols });

    /* ---- 未映射科目 ---- */
    var un = [];
    un.push(['未映射科目（这些科目没有归到任何报表项目，必须先指定，否则汇总一定对不上）']);
    un.push(['科目编码', '科目名称', '期末余额', '原始行号']);
    (box.agg ? box.agg.unmapped : []).forEach(function (u) {
      un.push([u.code, u.name, (u.signedCents || 0) / 100, u.srcRow]);
    });
    if (!un.length) un.push(['（无）']);
    sheets.push({ name: '未映射科目', rows: un, moneyCols: [2] });

    /* ---- 单边项目 ---- */
    var one = [];
    one.push(['只在一边出现的项目（要么项目名对不上，要么对方表漏了）']);
    one.push(['方向', '项目', '金额']);
    if (box.cmp) {
      box.cmp.onlyAgg.forEach(function (x) { one.push(['只有科目余额表有', x.item, (x.cents || 0) / 100]); });
      box.cmp.onlyTarget.forEach(function (x) { one.push(['只有对方表有', x.item, (x.cents || 0) / 100]); });
    }
    if (one.length === 2) one.push(['（无）']);
    sheets.push({ name: '单边项目', rows: one, moneyCols: [2] });

    /* ---- 说明 ---- */
    var help = [];
    help.push(['说明', '']);
    help.push(['工具版本', '勾稽助手 v' + VERSION]);
    help.push(['科目余额表', (box.bs && box.bs.label) || '']);
    help.push(['对方表', (box.target && box.target.label) || '']);
    if (box.target) help.push(['取数列', box.target.blockName + ' · ' + box.target.colLabel]);
    if (box.bs) {
      help.push(['科目数', box.bs.accountCount]);
      help.push(['末级科目数', box.bs.leafCount]);
      help.push(['借贷差额', moneyText(box.bs.balanceDiffCents),
        box.bs.balanceDiffCents === 0 ? '平' :
        '末级科目借贷不平。可能原因：这张表不含损益类科目，或含内部往来未抵销——先确认口径再勾稽']);
      help.push(['恒等式不平的行数', box.bs.identityFailed,
        box.bs.identityFailed ? '这些科目「期初+本期借-本期贷≠期末」，多半是列错位，需回原表核对' : '全部成立']);
    }
    help.push([]);
    help.push(['怎么用这张底稿']);
    help.push(['1', '先看「未映射科目」：有余额的科目没归到报表项目，汇总一定对不上，先在上一步给它指定项目。']);
    help.push(['2', '再看「勾稽表」里差异不为 0 的项目，右侧「构成科目数」告诉你有几个科目凑成这个数，回工具里能展开看明细。']);
    help.push(['3', '最后看「单边项目」：一边有一边没有，通常是项目名称写法不同（如「应收票据」对「应收票据净额」）。']);
    help.push([]);
    help.push(['口径']);
    help.push(['· 金额一律按「分」整数计算，本表以元展示，保留两位小数。']);
    help.push(['· 只汇总末级科目，上级科目不重复计入。']);
    help.push(['· 备抵科目（累计折旧、减值准备、库存股）已按负向计入对应项目。']);
    help.push(['· 未映射科目不计入任何项目，单独列出——本工具不猜科目归属。']);
    help.push(['· 差异 = 科目余额表汇总 − 对方表。正数=科目侧多，负数=科目侧少。']);
    help.push([]);
    help.push(['本表由程序生成，只做数据比对，不代替审计判断。']);
    sheets.push({ name: '说明', rows: help, moneyCols: [] });

    return sheets;
  }

  /** 数据来源 → 文件名 */
  function fileName(box) {
    var who = String((box && box.client) || '').trim() || '未命名';
    var d = String((box && box.dateText) || '').replace(/[^\d]/g, '') || '00000000';
    var p = box && box.target && box.target.colLabel ? box.target.colLabel : '勾稽';
    var s = '勾稽_' + who + '_' + p + '_' + d + '.xlsx';
    return s.replace(/[\\/:*?"<>|]/g, '-');
  }

  function applyMoneyFormat(XLSX, ws, rowCount, cols) {
    for (var r = 0; r < rowCount; r++) {
      for (var i = 0; i < cols.length; i++) {
        var cell = ws[XLSX.utils.encode_cell({ r: r, c: cols[i] })];
        if (cell && cell.t === 'n') cell.z = '#,##0.00';
      }
    }
  }

  /** 装配工作簿：界面与测试共用这一个入口 */
  function buildWorkbook(XLSX, box) {
    var sheets = buildSheets(box);
    var wb = XLSX.utils.book_new();
    sheets.forEach(function (s) {
      var ws = XLSX.utils.aoa_to_sheet(s.rows);
      applyMoneyFormat(XLSX, ws, s.rows.length, s.moneyCols || []);
      var ref = ws['!ref'];
      if (ref) {
        var range = XLSX.utils.decode_range(ref);
        ws['!cols'] = [];
        for (var c = 0; c <= range.e.c; c++) ws['!cols'].push({ wch: c === 1 ? 30 : 16 });
      }
      XLSX.utils.book_append_sheet(wb, ws, s.name);
    });
    return { wb: wb, fileName: fileName(box), sheetNames: sheets.map(function (s) { return s.name; }) };
  }

  return {
    VERSION: VERSION,
    toCents: toCents, moneyText: moneyText, norm: norm, str: str,
    CHART: CHART, CHART_MAP: CHART_MAP,
    detectBalanceHeader: detectBalanceHeader, parseBalanceRows: parseBalanceRows,
    balanceCheck: balanceCheck, identityCheck: identityCheck, leafAccounts: leafAccounts,
    detectStatementHeader: detectStatementHeader, parseStatementBlock: parseStatementBlock,
    buildSheets: buildSheets, buildWorkbook: buildWorkbook, fileName: fileName,
    codePrefix: codePrefix, mapAccount: mapAccount,
    contribution: contribution, aggregateItems: aggregateItems,
    normItem: normItem, compareItems: compareItems, classOfCode: classOfCode, CLASSES: CLASSES
  };
}));

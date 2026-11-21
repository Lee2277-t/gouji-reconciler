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

  var VERSION = '0.7.6';

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
    // 余额表里用「-」「—」表示 0 是极常见的写法，不能当成读不出
    if (typeof v === 'string' && /^[-—–－]$/.test(v.trim())) return { cents: 0, text: true };
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
    // 2601/2602 只有保险公司才用（未到期责任准备金/保险责任准备金）；
    // 非保险客户常把 2601 当「长期借款」，硬映射会静默算错，所以故意不内置，让它进未映射清单
    ['2701', '长期应付款', 'C'], ['2702', '长期应付款', 'D', -1],   // 未确认融资费用（减项）
    ['2711', '专项应付款', 'C'],
    ['2801', '预计负债', 'C'], ['2901', '递延所得税负债', 'C'],

    /* ---- 所有者权益 ---- */
    ['4001', '实收资本（或股本）', 'C'], ['4002', '资本公积', 'C'], ['4003', '其他综合收益', 'C'],
    ['4101', '盈余公积', 'C'], ['4103', '未分配利润', 'C'], ['4104', '未分配利润', 'C'],
    ['4201', '库存股', 'D', -1],                     // 库存股是借方余额的权益备抵

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
   * 报表项目属于哪一栏。
   * 这是**报表属性**，不能由「凑成它的科目有几个」推断 ——
   * 存货常由 1403/1405 加 5001/5101 凑成，按科目个数投票会把它归成「成本」，
   * 于是在资产栏比对时整行被过滤掉，差异凭空消失。
   */
  var ITEM_CLASS = {
    '货币资金': '资产', '交易性金融资产': '资产', '衍生金融资产': '资产', '应收票据': '资产',
    '应收账款': '资产', '预付款项': '资产', '应收股利': '资产', '应收利息': '资产',
    '其他应收款': '资产', '存货': '资产', '债权投资': '资产', '其他债权投资': '资产',
    '其他权益工具投资': '资产', '长期股权投资': '资产', '投资性房地产': '资产',
    '长期应收款': '资产', '固定资产': '资产', '在建工程': '资产', '工程物资': '资产',
    '固定资产清理': '资产', '使用权资产': '资产', '生产性生物资产': '资产',
    '无形资产': '资产', '开发支出': '资产', '商誉': '资产', '长期待摊费用': '资产',
    '递延所得税资产': '资产', '其他流动资产': '资产', '其他非流动资产': '资产',

    '短期借款': '负债', '交易性金融负债': '负债', '应付票据': '负债', '应付账款': '负债',
    '预收款项': '负债', '合同负债': '负债', '应付职工薪酬': '负债', '应交税费': '负债',
    '应付利息': '负债', '应付股利': '负债', '其他应付款': '负债', '递延收益': '负债',
    '长期借款': '负债', '应付债券': '负债', '长期应付款': '负债', '专项应付款': '负债',
    '预计负债': '负债', '递延所得税负债': '负债', '其他流动负债': '负债',
    '保险合同准备金': '负债',

    '实收资本（或股本）': '权益', '资本公积': '权益', '其他综合收益': '权益',
    '盈余公积': '权益', '未分配利润': '权益', '库存股': '权益',

    '营业收入': '损益', '营业成本': '损益', '税金及附加': '损益', '销售费用': '损益',
    '管理费用': '损益', '财务费用': '损益', '资产减值损失': '损益',
    '信用减值损失': '损益', '投资收益': '损益', '公允价值变动收益': '损益',
    '汇兑收益': '损益', '营业外收入': '损益', '营业外支出': '损益', '所得税费用': '损益'
  };

  function itemClass(item, fallbackCode) {
    return ITEM_CLASS[item] || (fallbackCode ? classOfCode(fallbackCode) : '其他');
  }

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
    { key: 'opening', words: ['期初余额', '期初', '年初余额', '年初数', '年初', '上期期末', '期初数'] },
    { key: 'period', words: ['期间发生额', '本期发生额', '本期发生', '本期数', '本期', '发生额'] },
    { key: 'year', words: ['累计发生额', '本年累计', '本年发生额', '累计发生', '年度累计', '本年'] },
    { key: 'closing', words: ['期末余额', '期末', '本期期末', '年末余额', '年末数', '年末', '期末数'] }
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

  /**
   * 二级表头 → 借/贷/金额/方向。
   * 用「包含」而不是全等：扁平表头里 「期初借方」「本期贷方」「期末余额」 都是合成词，
   * 全等匹配会把它们整列丢掉，于是期初期末全空、还照样报「借贷平」。
   */
  function pickLeaf(text) {
    var raw = str(text);
    // 括号里写方向的（期初余额(借)）：norm 会把括号整个剥掉，必须在原文上看
    if (/[（(]\s*借\s*[）)]/.test(raw)) return 'debit';
    if (/[（(]\s*贷\s*[）)]/.test(raw)) return 'credit';
    var t = norm(text);
    if (!t) return '';
    if (/方向/.test(t) || /^(借贷|借贷方)$/.test(t)) return 'dir';
    // 「期初余额（借）」这种把方向塞在括号里的，按借方列处理
    if (/借方|\(借\)$/.test(t) || t === '借') return 'debit';
    if (/贷方|\(贷\)$/.test(t) || t === '贷') return 'credit';
    if (/金额|余额/.test(t)) return 'amount';
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
    var unrecognized = [];
    var conflictCols = [];
    var colCount = Math.max(topRow.length, subRow.length);
    for (var c2 = 0; c2 < colCount; c2++) {
      var g = pickGroup(topRow[c2]);
      if (g) lastGroup = g;

      var leaf = span === 2 ? pickLeaf(subRow[c2]) : '';
      if (!leaf && span === 1) leaf = pickLeaf(topRow[c2]);

      // 单行表头里光写着「借方/贷方」、自己不带分组词的，是本期发生额。
      // 不这样处理，它会被 lastGroup 挂到前面的「期初余额」上，期初期末一起错。
      if (span === 1 && (leaf === 'debit' || leaf === 'credit') && !g) {
        lastGroup = 'period';
      }

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
        if (lastGroup) {
          if (columns[lastGroup].dir !== undefined && columns[lastGroup].dir !== c2) {
            // 同一组两个方向列，后一个会盖掉前一个，必须说出来
            conflictCols.push((c2 + 1) + ' 列的方向列与 ' + (columns[lastGroup].dir + 1) + ' 列重复');
          }
          columns[lastGroup].dir = c2;
        }
        if (dirCol < 0) dirCol = c2;
        continue;
      }
      if (leaf === 'amount' && !lastGroup) { amountCol = c2; continue; }
      if (!lastGroup || !leaf) {
        // 有表头文字但一列都没认出来 —— 这一列的数据会被整列丢掉。
        // 静默丢列是最危险的：少算还照样报「借贷平」。收集起来报给使用者。
        var rawHead = span === 2 ? str(subRow[c2] || topRow[c2]) : str(topRow[c2]);
        if (!str(rawHead).trim()) {
          // 整列一个数都没有的，是格式残留的空列，不是「没认出来」
          var hasData = false;
          for (var rr = bestR + span; rr < Math.min(rows.length, bestR + span + 200); rr++) {
            if (str((rows[rr] || [])[c2]).trim() !== '') { hasData = true; break; }
          }
          if (!hasData) continue;
          rawHead = '（无表头）';
        }
        if (str(rawHead).trim() && !/账户代码|科目编码|科目代码|账户编号|科目编号|编码|名称/.test(norm(rawHead))) {
          unrecognized.push((c2 + 1) + ' 列「' + str(rawHead).trim() + '」');
        }
        continue;
      }

      if (columns[lastGroup][leaf] === undefined) {
        columns[lastGroup][leaf] = c2;
      } else if (columns[lastGroup][leaf] !== c2) {
        // 同组同类列出现两次（如「期末余额(借)」「期末余额(贷)」都被认成 amount），
        // 后一个会被丢掉，金额直接少一半
        conflictCols.push((c2 + 1) + ' 列与 ' + (columns[lastGroup][leaf] + 1) + ' 列都像「' + leaf + '」，只用了前者');
      }
    }

    if (codeCol < 0) problems.push('没找到科目编码列');
    if (nameCol < 0) problems.push('没找到科目名称列');

    // 单行表头里只写了「余额」「方向」、没有分组词的（科目编码|科目名称|余额），
    // 默认按期末余额处理 —— 这是最常见的简表写法
    var anyGroup = ['opening', 'period', 'year', 'closing'].some(function (k) {
      return Object.keys(columns[k]).length > 0;
    });
    if (!anyGroup && amountCol >= 0) {
      columns.closing.amount = amountCol;
      if (dirCol >= 0) columns.closing.dir = dirCol;
    }

    var hasAnyAmount = ['opening', 'period', 'year', 'closing'].some(function (k) {
      return columns[k].debit !== undefined || columns[k].credit !== undefined ||
             columns[k].amount !== undefined;
    });
    if (!hasAnyAmount) problems.push('没找到金额列（借方/贷方/金额都没有）');
    if (unrecognized.length) problems.push('这些列表头没认出来，整列没有参与计算：' + unrecognized.slice(0, 6).join('、') + (unrecognized.length > 6 ? ' 等' : ''));
    if (conflictCols.length) problems.push('列冲突：' + conflictCols.slice(0, 4).join('；'));

    return {
      rowIndex: bestR, span: span, codeCol: codeCol, codeCol2: codeCol2, nameCol: nameCol,
      dirCol: dirCol, amountCol: amountCol, columns: columns, problems: problems,
      hasAmountColumns: hasAnyAmount, unrecognized: unrecognized, conflicts: conflictCols
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
    var badClosing = 0;
    var noCodeRows = 0;

    function centsAt(group, kind, row) {
      var c = cols[group] && cols[group][kind];
      if (c === undefined) return null;
      var v = toCents(row[c]);
      return v.cents;
    }

    // 先看期末这一组到底有没有填过数：整列空的表不能按「每行都是 0」处理
    var closingGroupPre = cols.closing || {};
    var closingColumnEmpty = false;
    if (closingGroupPre.amount !== undefined || closingGroupPre.debit !== undefined || closingGroupPre.credit !== undefined) {
      var filledPre = 0;
      for (var rp = header.rowIndex + span; rp < rows.length; rp++) {
        var rowp = rows[rp] || [];
        [closingGroupPre.amount, closingGroupPre.debit, closingGroupPre.credit].forEach(function (cc) {
          if (cc !== undefined && str(rowp[cc]).trim() !== '') filledPre++;
        });
      }
      closingColumnEmpty = (filledPre === 0);
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

      // 合计/汇总行：编码列或名称列写着「合计」的都算（有的表把合计写在编码列）
      if (/合计|总计|小计|汇总/.test(code) || /合计|总计|小计|汇总/.test(name)) {
        skipped.push({ srcRow: excelRow, reason: '汇总行', code: code, name: name });
        continue;
      }
      if (!code || codePrefix(code) === '') {
        // 有名称没编码、或编码里没有四位数字（如「总计」）：是标题/说明行，不是科目
        // 合计/总计行是故意跳过的，不算「丢了」
        var isSummaryRow = isTotalLabel(code) || isTotalLabel(name);
        // 有金额、却没有编码也没有名称：合并单元格展开或辅助核算行，
        // 这种行不能当空行静默丢掉，否则金额凭空消失
        var hasMoney = ['opening', 'period', 'year', 'closing'].some(function (g) {
          var g2 = cols[g] || {};
          return [g2.amount, g2.debit, g2.credit].some(function (cc) {
            if (cc === undefined) return false;
            var n = toCents(row[cc]);
            return n.cents !== null && n.cents !== 0;
          });
        });
        skipped.push({
          srcRow: excelRow, code: code, name: name, money: hasMoney,
          reason: isSummaryRow ? '汇总行'
            : (hasMoney ? '缺科目编码和名称'
              : ((code || name) ? '无科目编码' : '空行'))
        });
        // 只有带金额的行才会让合计对不上，没有金额的行（续表、分隔）不报
        if (!isSummaryRow && hasMoney) noCodeRows++;
        continue;
      }

      function dirOf(group) {
        var c = cols[group] && cols[group].dir;
        // 组内没有方向列时退回全局方向列：
        // 「科目编码|科目名称|方向|期初余额|本期借方|本期贷方|期末余额」只写了一个方向列，
        // 它管的是整行的余额方向，不归任何一组。
        if (c === undefined) c = header.dirCol;
        if (c === undefined || c < 0) return '';
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
        var hasDir = (cols[group] && cols[group].dir !== undefined) ||
                     (header.dirCol !== undefined && header.dirCol >= 0);
        var amtCol = cols[group] && cols[group].amount;
        var amtRaw = amtCol === undefined ? '' : str(row[amtCol]).trim();
        var amt = centsAt(group, 'amount', row);
        var dr = centsAt(group, 'debit', row);
        var cr = centsAt(group, 'credit', row);

        if (amt !== null) {
          // 有方向列才按方向定正负；没有方向列时金额本身带符号，取绝对值会把负数抹平
          if (hasDir) return /贷/.test(dirOf(group)) ? -Math.abs(amt) : Math.abs(amt);
          return amt;
        }
        // 格子里写了东西但读不成数字 —— 这是「读不出」，不是「零」
        if (amtRaw !== '') return null;
        if (dr !== null || cr !== null) return (dr || 0) - (cr || 0);
        var d = dirOf(group);
        // 有方向列、写着「平」——这才是确定的零
        if (d && !/借|贷/.test(d)) return 0;
        // 既没金额、又没方向列：
        // 期末列别处有数、这格空着 —— 这是「余额为 0」的常见写法；
        // 整列都空 —— 那是列没填，属于「不知道」。
        if (!d) return (hasDir || !closingColumnEmpty) ? 0 : null;
        return null;                               // 写着借/贷却没金额：推出来再说
      }

      var closingInferred = false;
      var opDir = dirOf('opening');
      var clDir = dirOf('closing');
      var opening = signed('opening');

      var pDr = centsAt('period', 'debit', row);
      var pCr = centsAt('period', 'credit', row);
      // 发生额格写了文字（「见附注」「待补」）：这行不能算，也不能拿 0 去推期末
      var periodBad = false;
      ['debit', 'credit'].forEach(function (k) {
        var cc = cols.period && cols.period[k];
        if (cc === undefined) return;
        var rawp = str(row[cc]).trim();
        if (rawp !== '' && toCents(rawp).cents === null) periodBad = true;
      });
      // 发生额写文字：只置标志，最后统一计一次

      // 期末的借/贷格写文字，同样不能当 0
      var closingBad = false;
      ['debit', 'credit', 'amount'].forEach(function (k) {
        var cc = cols.closing && cols.closing[k];
        if (cc === undefined) return;
        var rawc = str(row[cc]).trim();
        if (rawc === '') return;
        var cc2 = toCents(rawc);
        if (cc2.cents === null && !cc2.text) closingBad = true;
      });
      // 期末写文字在下面统一计数，这里不要再加一次
      var yDr = centsAt('year', 'debit', row);
      var yCr = centsAt('year', 'credit', row);

      // 期末：有期末数就用；没填就用 期初 + 期间借 - 期间贷。
      // 注意：只有「确实有期间发生额列」时才敢推算。少了这个条件，
      // 一张只认出期初的表会把期初当期末，而借贷两边同时错、看上去还是平的。
      var closing = closingColumnEmpty ? null : signed('closing');
      var clRaw = '';
      if (cols.closing && cols.closing.amount !== undefined) clRaw = str(row[cols.closing.amount]).trim();
      var clWritten = clRaw !== '' && toCents(clRaw).cents === null;   // 写了东西但不是数字
      if (closingBad || clWritten) {
        // 期末格写的是文字，不是空白。这时候推算等于把"读不出"粉饰成"读出来了"，
        // 宁可留空并报出来，让人回原表看那一格。
        // 一行只计一次：closingBad 和 clWritten 说的是同一件事
        closing = null;
      } else if (closing === null && opening !== null && hasGroup('period') && !closingColumnEmpty && !periodBad) {
        closing = opening + (pDr || 0) - (pCr || 0);
        closingInferred = true;
      }

      // 一行最多算一次「写的是文字」：发生额、期末余额可能同时写文字，
      // 那是同一行的同一个问题，不能按格子数报成人行
      if (periodBad || closingBad || clWritten) badClosing++;

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
    // 金额列一个都没认出来，或者一个期末余额都算不出来：这张表不能用来勾稽
    if (header.hasAmountColumns === false) problems.push('没有可用的金额列，这张表不能用来勾稽');
    if (badClosing) problems.push(badClosing + ' 行的期末余额或发生额写的是文字、读不成数字，这些行按「取不到」处理，没有推算');
    if (noCodeRows) problems.push(noCodeRows + ' 行有内容但科目编码里取不出四位数字，这些行没有计入（金额没有消失，只是没归到科目上，请核对）');
    var withClosing = accounts.filter(function (a) { return a.closingCents !== null; }).length;
    // 期末列整列都是空格子：每家 ERP 都可能这样（导出了列但没填），
    // 这时每行都算成 0，借贷「平」得毫无意义
    if (closingColumnEmpty) problems.push('「期末余额」这一组整列都是空的，没有取到任何期末数');
    if (accounts.length && !withClosing) problems.push('所有科目的期末余额都取不到，不能用来勾稽；请检查表头行是否选对了');

    return {
      accounts: accounts, skipped: skipped, problems: problems, header: header,
      withClosing: withClosing,
      inferred: accounts.filter(function (a) { return a.closingInferred; }).length
    };
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
    // 手工映射的方向要按「这个编码下所有科目合起来的净额」定一次。
    // 逐科目各取各的方向会把借贷混向的往来放大两倍贷方金额（净 6000 算成 14000）。
    var ovDirMap = {};
    Object.keys(overrides).forEach(function (k) {
      var net = 0;
      accounts.forEach(function (a) {
        if (codePrefix(a.code) === k) net += (a.creditCents || 0) - (a.debitCents || 0);
      });
      ovDirMap[k] = net >= 0 ? 'C' : 'D';
    });
    var byItem = {};
    var order = [];
    var unmapped = [];
    var missing = [];
    var total = 0;

    accounts.forEach(function (a) {
      var ovRaw = overrides[codePrefix(a.code)];
      var entry = null;
      if (ovRaw) {
        // 手工映射的方向不能用首位数字猜：6115 资产处置损益、6117 其他收益 都是贷方科目，
        // 猜成借方会让金额变成相反数。优先用使用者显式写的 :C / :D，其次看科目实际余额方向。
        var ovItem = ovRaw, ovDir = '';
        var mOv = /^(.*?)\s*[:：]\s*([CDcd])\s*$/.exec(ovRaw);
        if (mOv) { ovItem = mOv[1]; ovDir = mOv[2].toUpperCase(); }
        var dir2 = ovDir || ovDirMap[codePrefix(a.code)] || 'D';
        entry = { item: ovItem, dir: dir2, sign: 1, manual: true, manualDir: ovDir };
      } else {
        entry = mapAccount(a.code);
        // 1231 是「坏账准备」总账科目，二级可能挂在不同往来上。
        // 全部扣在应收账款上，会让应收账款少、其他应收款多，金额还互相抵消看不出来。
        if (entry && entry.code === '1231' && /其他应收/.test(str(a.name))) {
          entry = { code: '1231', item: '其他应收款', dir: 'C', sign: -1 };
        } else if (entry && entry.code === '1231' && /应收票据/.test(str(a.name))) {
          entry = { code: '1231', item: '应收票据', dir: 'C', sign: -1 };
        }
      }

      // 只对「从余额表解析出来的行」判这一条：手工构造的科目可能只给借贷金额、没有期末字段
      if (Object.prototype.hasOwnProperty.call(a, 'closingCents') && a.closingCents === null) {
        // 期末余额读不出来：既不能当 0 汇总（那等于替真实数据判了 0），
        // 也不能悄悄丢掉。单独列出来，让人回原表看那一行。
        missing.push({ code: a.code, name: a.name, srcRow: a.srcRow });
        return;
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
        byItem[entry.item] = { item: entry.item, cents: 0, accounts: [], cls: itemClass(entry.item, a.code) };
        order.push(entry.item);
      }
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
            return it;
    });
    items.sort(function (x, y) { return Math.abs(y.cents) - Math.abs(x.cents); });
    unmapped.sort(function (x, y) { return Math.abs(y.signedCents) - Math.abs(x.signedCents); });

    missing.sort(function (x, y) { return (x.srcRow || 0) - (y.srcRow || 0); });
    return { items: items, unmapped: unmapped, missing: missing, total: total };
  }

  /* ============================================================
   * 五、比对：汇总结果 vs 试算表/报表
   * ============================================================ */

  /** 同一项目在不同报表里的写法不一样，先归一再比 */
  var ITEM_ALIAS = {
    '实收资本': '实收资本或股本',
    '股本': '实收资本或股本',
    '实收资本净额': '实收资本或股本',
    '应付工资': '应付职工薪酬',
    '应交税金': '应交税费',
    '主营业务收入': '营业收入',
    '主营业务成本': '营业成本',
    '其他业务收入': '营业收入',
    '其他业务成本': '营业成本'
  };

  /** 项目名归一：全角括号、空格、冒号都不算差别 */
  function normItem(name) {
    // 报表里常带「一、」「减：」「其中：」这类前缀，去掉才能和科目侧的项目名对上
    var raw = str(name).trim()
      .replace(/^[一二三四五六七八九十]+[、.．]/, '')
      .replace(/^(减|加)[:：]/, '')
      .replace(/^[-—－]+/, '');
    var k = norm(raw);
    return ITEM_ALIAS[k] || k;
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
        right[k] = { item: t.item, cents: 0, rows: [], isTotal: !!t.isTotal || isStatementTotal(t.item) };
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
      if (left[k] === undefined) { left[k] = { item: it.item, cents: it.cents, accounts: it.accounts, cls: it.cls }; leftOrder.push(k); }
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
    var targetTotals = [];
    rightOrder.forEach(function (k) {
      if (left[k]) return;
      if (right[k].isTotal) { targetTotals.push({ item: right[k].item, cents: right[k].cents }); return; }
      // 报表模板里大量零值行（拆出资金、衍生金融资产…），没金额就不算差异
      if (!right[k].cents && !opts.keepZero) return;
      onlyTarget.push({ item: right[k].item, cents: right[k].cents, rows: right[k].rows });
    });

    diffs.sort(function (x, y) { return Math.abs(y.diffCents) - Math.abs(x.diffCents); });
    onlyAgg.sort(function (x, y) { return Math.abs(y.cents) - Math.abs(x.cents); });
    onlyTarget.sort(function (x, y) { return Math.abs(y.cents) - Math.abs(x.cents); });

    return {
      matched: matched, diffs: diffs, onlyAgg: onlyAgg, onlyTarget: onlyTarget,
      targetTotals: targetTotals,
      sameCount: matched.length,
      diffTotalCents: diffs.reduce(function (s, x) { return s + x.diffCents; }, 0),
      // 单边项目也是差异：左边有右边没有，说明要么项目名对不上，要么报表漏了
      unbalancedCents: onlyAgg.reduce(function (s, x) { return s + x.cents; }, 0) -
                       onlyTarget.reduce(function (s, x) { return s + x.cents; }, 0),
      // 单边项目绝对值之和：两边互相抵消时 unbalancedCents 会是 0，
      // 但「有 3 项对不上名字」这件事并没有消失，所以还要有这个数
      singleSideAbsCents: onlyAgg.reduce(function (s, x) { return s + Math.abs(x.cents); }, 0) +
                          onlyTarget.reduce(function (s, x) { return s + Math.abs(x.cents); }, 0),
      singleSideCount: onlyAgg.length + onlyTarget.length,
      // 「干净」= 没有差异、没有单边，**并且**没有未映射科目、没有期末读不出的科目。
      // 只看差异和单边，会在有明显缺口时说「全部勾平」——这是最误导人的假结论。
      clean: diffs.length === 0 && onlyAgg.length === 0 && onlyTarget.length === 0
        && !((agg.unmapped || []).filter(function (u) { return (u.signedCents || 0) !== 0; }).length)
        && !(agg.missing && agg.missing.length),
      gapCount: (onlyAgg.length + onlyTarget.length +
        ((agg.unmapped || []).filter(function (u) { return (u.signedCents || 0) !== 0; }).length) +
        ((agg.missing && agg.missing.length) || 0))
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
   *
   * 特别地：如果连一个期末余额都没取到，balanced 必须是 null（不知道），
   * 不能因为「0 = 0」就给出「借贷平」这个结论 —— 那是最危险的假阳性。
   */
  function balanceCheck(accounts) {
    var codes = accounts.map(function (a) { return str(a.code).replace(/[^\d]/g, ''); });
    var leaves = accounts.filter(function (a, i) {
      var c = codes[i];
      if (!c) return true;
      return !codes.some(function (o, j) { return j !== i && o && o !== c && o.indexOf(c) === 0; });
    });
    var dr = 0, cr = 0, valued = 0;
    leaves.forEach(function (a) {
      var v = a.closingCents;
      if (v === null || v === undefined) return;
      valued++;
      if (v > 0) dr += v; else cr += -v;
    });
    // 有行没取到期末余额时，「借贷平」只是「已取到的部分平」，不能当结论
    // 有部分行没取到数时仍然给结论，但把缺的行数报出来（clean 会因此不成立）
    var conclusive = valued > 0;
    return {
      leafless: leaves.length,
      parentCount: accounts.length - leaves.length,
      valued: valued,
      debitCents: dr,
      creditCents: cr,
      diffCents: dr - cr,
      balanced: conclusive ? (dr === cr) : null,
      conclusive: conclusive,
      missingRows: leaves.length - valued
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
    var inferred = 0;
    accounts.forEach(function (a) {
      if (a.openingCents === null || a.closingCents === null) return;
      // 期末是推算出来的，这条恒等式对它恒成立，拿它当检验等于自己证明自己
      if (a.closingInferred) { inferred++; return; }
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
    return { checked: checked, bad: bad, failed: bad.length, inferred: inferred };
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

  /** 报表里的小计/合计行：不是项目，是表内加总 */
  function isStatementTotal(name) {
    // 「其中：xx」是备查行，和主项目并列相加会算重，先按原始文字判掉
    if (/^其中/.test(str(name).replace(/[\s\u3000]/g, ''))) return true;
    var t = str(name)
      .replace(/[\s\u3000:：]/g, '')
      .replace(/[（(][^）)]*[）)]/g, '')          // 去掉「（或股东权益）」这类括注
      .replace(/^[一二三四五六七八九十]+[、.．]/, '')  // 去掉「一、」「二.」
      .replace(/^(减|加|其中|内中)[:：]?/, '');    // 去掉「减：」「其中：」
    if (/^(合计|总计|小计)$/.test(t)) return true;
    if (/(合计|总计|小计)$/.test(t)) return true;
    // 利润表的小计行不带「合计」二字
    if (/^(营业总收入|营业总成本|营业利润|利润总额|净利润|持续经营净利润|终止经营净利润|综合收益总额|其他综合收益的税后净额|归属于母公司股东的净利润|归属于母公司所有者的净利润|少数股东损益)$/.test(t)) return true;
    // 「其中：应收账款」是备查行，与主项目并列相加会算重
    if (/^其中/.test(t)) return true;
    return false;
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
      if (!name) {
        // 有金额、没项目名：不能静默丢，否则这笔钱在底稿里彻底消失
        var amtNoName = toCents(row[ac.col]);
        if (amtNoName.cents !== null && amtNoName.cents !== 0) {
          skipped.push({ srcRow: excelRow, reason: '无项目名', money: true });
        }
        continue;
      }
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
      var clean = name.replace(/^[\s\u3000]+/, '');
      items.push({
        item: clean,
        cents: parsed.cents,
        srcRow: excelRow,
        col: ac.col,
        // 「流动资产合计」「资产总计」这类不是报表项目，是表内小计。
        // 混进比对会变成一堆假「单边项目」，把真正的信号盖住。
        isTotal: isStatementTotal(clean)
      });
    }
    return { items: items, skipped: skipped, problems: [] };
  }

  /* ============================================================
   * 七、输出：勾稽底稿
   * ============================================================ */

  /**
   * @param box {
   *   client, dateText,
   *   bs: {label, accountCount, leafCount, balanceDiffCents, identityFailed, skipped},
   *   target: {label, blockName, colLabel},
   *   agg, cmp, overrides
   * }
   */
  var LEVEL_TEXT = { leaf: '只汇总末级科目', top: '只汇总上级科目', all: '全部科目都算（可能重复计）' };

  function sheetHead(box, title) {
    return [
      [title, '', '', '', '被审计单位：' + (box.client || '（未填）'), ''],
      ['编制日期：' + (box.dateText || '') + ' · 金额单位：元', '', '', '', '', '']
    ];
  }

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

    // 单边项目也写进这张表。只写「差异合计 0.00」而不提单边，
    // 复核会以为已经勾平了 —— 而项目名对不上恰恰是最常见的情况。
    if (box.cmp && (box.cmp.onlyAgg.length || box.cmp.onlyTarget.length)) {
      rows.push([]);
      rows.push(['单边项目（不参与上面的差异计算，但同样要处理）', '', '', '', '', '']);
      rows.push(['方向', '项目', '金额', '', '', '']);
      box.cmp.onlyAgg.forEach(function (x) {
        rows.push(['只有科目余额表有', x.item, (x.cents || 0) / 100, '', '', '']);
      });
      box.cmp.onlyTarget.forEach(function (x) {
        rows.push(['只有对方表有', x.item, (x.cents || 0) / 100, '', '', '']);
      });
    }

    rows.push([]);
    rows.push(['差异合计（已匹配项目）', '', '', '', ((box.cmp ? box.cmp.diffTotalCents : 0) / 100), '']);
    rows.push(['单边差额（两边对冲后）', '', '', '', ((box.cmp ? box.cmp.unbalancedCents : 0) / 100), '']);
    rows.push(['单边项目绝对值合计', '', '', '', ((box.cmp ? box.cmp.singleSideAbsCents : 0) / 100), '']);
    // 总结论要把「被跳过的行」「明细表这一级」也算进去，否则分工表说勾平、
    // 同一本底稿的明细表页却写着差异 10,000，自相矛盾
    var skippedN = (box.bs && box.bs.skippedRows) || 0;
    var detailBad = box.detail && (
      (box.detail.diffs && box.detail.diffs.length) ||
      (box.detail.failed && box.detail.failed.length) ||
      (box.detail.onlyDetail && box.detail.onlyDetail.length) ||
      (box.detail.onlyAgg && box.detail.onlyAgg.length));
    var targetBad = box.targetSkipped || 0;
    var overallClean = box.cmp && box.cmp.clean && !skippedN && !detailBad && !targetBad;
    var why = [];
    if (box.cmp) {
      if ((box.cmp.diffs || []).length) why.push('项目差异 ' + box.cmp.diffs.length + ' 项');
      if (box.cmp.singleSideCount) why.push('单边项目 ' + box.cmp.singleSideCount + ' 项');
      // 未映射/读不出单独数，不要用 gapCount（它把单边也算进去了，会重复）
      var unN = ((box.agg && box.agg.unmapped) || []).filter(function (u) { return (u.signedCents || 0) !== 0; }).length;
      var missN = ((box.agg && box.agg.missing) || []).length;
      if (unN) why.push('未映射科目 ' + unN + ' 个');
      if (missN) why.push('期末读不出 ' + missN + ' 个');
    }
    if (skippedN) why.push('原表有 ' + skippedN + ' 行没有计入');
    if (targetBad) why.push('对方表有 ' + targetBad + ' 行没读到金额');
    if (detailBad) why.push('明细表勾稽有差异');
    if (box.bs && box.bs.identityFailed) why.push('另有 ' + box.bs.identityFailed + ' 行恒等式不平');
    var conclText = overallClean ? '全部勾平' : ('未勾平：' + why.join('，'));
    if (overallClean && box.bs && box.bs.identityFailed) {
      conclText += '（另有 ' + box.bs.identityFailed + ' 行恒等式不平，见说明页）';
    }
    rows.push(['结论', '', '', '', conclText, '']);
    sheets.push({ name: '勾稽表', rows: rows, moneyCols: moneyCols });

    /* ---- 明细表（三级勾稽） ---- */
    if (box.detail) {
      var dt = [];
      dt.push.apply(dt, sheetHead(box, '明细表 ↔ 科目余额表 勾稽表'));
      dt.push(['明细表', (box.detail.label || '') ]);
      dt.push(['取数列', box.detail.colLabel || ('第 ' + ((box.detail.col || 0) + 1) + ' 列')]);
      dt.push([]);
      dt.push(['工作表（报表项目）', '明细表合计', '科目余额表汇总', '差异', '取数方式', '原始行号']);
      var dtMoney = [1, 2, 3];
      var pushDt = function (d) {
        dt.push([
          d.item, (d.detailCents || 0) / 100,
          d.aggCents === null || d.aggCents === undefined ? '' : d.aggCents / 100,
          d.aggCents === null || d.aggCents === undefined ? '' : (d.diffCents || 0) / 100,
          d.source + (d.totalRow ? '（第 ' + d.totalRow + ' 行）' : ''), ''
        ]);
      };
      box.detail.diffs.forEach(pushDt);
      (box.detail.matched || []).forEach(pushDt);
      if ((box.detail.onlyDetail || []).length) {
        dt.push([]);
        dt.push(['明细表有、科目余额表没有的', '', '', '', '', '']);
        box.detail.onlyDetail.forEach(function (d) {
          dt.push([d.item, (d.detailCents || 0) / 100, '', '', d.source || '', '']);
        });
      }
      if (box.detail.onlyAgg && box.detail.onlyAgg.length) {
        dt.push([]);
        dt.push(['科目余额表有、明细表没有的', '', '', '', '', '']);
        box.detail.onlyAgg.forEach(function (x) {
          dt.push([x.item, '', (x.cents || 0) / 100, '', '', '']);
        });
      }
      if ((box.detail.failed || []).length) {
        dt.push([]);
        dt.push(['这一列上取不到金额的表', '', '', '', '', '']);
        box.detail.failed.forEach(function (f) { dt.push([f.sheet, '', '', '', f.reason || '', '']); });
      }
      dt.push([]);
      dt.push(['差异合计', '', '', ((box.detail.diffTotalCents || 0) / 100), '', '']);
      // 这一页也要有自己的结论，否则上面几页说勾平、这里写着差异，没人知道该信哪个
      var dClean = !(box.detail.diffs && box.detail.diffs.length)
        && !(box.detail.failed && box.detail.failed.length)
        && !(box.detail.onlyDetail && box.detail.onlyDetail.length)
        && !(box.detail.onlyAgg && box.detail.onlyAgg.length);
      dt.push([]);
      dt.push(['结论', '', '', '', dClean ? '明细表全部勾平'
        : ('未勾平：' + ((box.detail.diffs || []).length) + ' 项差异、' + ((box.detail.failed || []).length) + ' 项取不到金额、' +
          ((box.detail.onlyAgg || []).length) + ' 项明细表里没有、' + ((box.detail.onlyDetail || []).length) + ' 项科目余额表里没有'), '']);
      sheets.push({ name: '明细表勾稽', rows: dt, moneyCols: dtMoney });
    }

    /* ---- 未映射科目 ---- */
    var un = [];
    un.push.apply(un, sheetHead(box, '未映射科目'));
    un.push(['未映射科目（这些科目没有归到任何报表项目，必须先指定，否则汇总一定对不上）']);
    un.push(['科目编码', '科目名称', '期末余额', '原始行号']);
    (box.agg ? box.agg.unmapped : []).forEach(function (u) {
      un.push([u.code, u.name, (u.signedCents || 0) / 100, u.srcRow]);
    });
    if (!box.agg || !box.agg.unmapped.length) un.push(['（无）']);
    // 期末余额读不出的科目：既没进项目、也没被当成 0，必须留痕
    if (box.agg && box.agg.missing && box.agg.missing.length) {
      un.push([]);
      un.push(['期末余额读不出、没有计入任何项目的科目']);
      un.push(['科目编码', '科目名称', '', '原始行号']);
      box.agg.missing.forEach(function (m) {
        un.push([m.code, m.name, '（读不出）', m.srcRow]);
      });
    }
    // 恒等式对不上的行也要能定位到具体科目，README 一直这么承诺
    if (box.bs && box.bs.identityBad && box.bs.identityBad.length) {
      un.push([]);
      un.push(['期初+本期借-本期贷≠期末的科目']);
      un.push(['科目编码', '科目名称', '', '原始行号']);
      box.bs.identityBad.forEach(function (b) {
        un.push([b.code, b.name, '（恒等式不平）', b.srcRow]);
      });
    }
    sheets.push({ name: '未映射科目', rows: un, moneyCols: [2] });

    /* ---- 单边项目 ---- */
    var one = [];
    one.push.apply(one, sheetHead(box, '单边项目'));
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
    help.push.apply(help, sheetHead(box, '说明'));
    help.push(['说明', '']);
    help.push(['工具版本', '勾稽助手 v' + VERSION]);
    help.push(['科目余额表', (box.bs && box.bs.label) || '']);
    help.push(['对方表', (box.target && box.target.label) || '']);
    if (box.target) help.push(['取数列', box.target.blockName + ' · ' + box.target.colLabel]);
    if (box.bs) {
      help.push(['科目数', box.bs.accountCount]);
      help.push(['末级科目数', box.bs.leafCount]);
      help.push(['汇总口径', LEVEL_TEXT[box.bs.level] || LEVEL_TEXT.leaf,
        '上下级同时存在时只算末级，避免重复计']);
      help.push(['借贷差额', moneyText(box.bs.balanceDiffCents),
        box.bs.balanceDiffCents === 0 ? '平' :
        '末级科目借贷不平。可能原因：这张表不含损益类科目，或含内部往来未抵销——先确认口径再勾稽']);
      var checkedN = box.bs.identityChecked;
      if (checkedN === 0) {
        help.push(['恒等式校验', '无法校验', '所有能比的期末余额都是推算出来的（推算值与公式天然一致），这类表不能拿恒等式当证据']);
      } else {
        help.push(['恒等式不平的行数', box.bs.identityFailed,
          box.bs.identityFailed ? '这些科目「期初+本期借-本期贷≠期末」，多半是列错位，需回原表核对'
            : ('参与校验 ' + checkedN + ' 行，全部成立')]);
      }
      if (box.bs.inferred) help.push(['期末余额为推算值的行数', box.bs.inferred, '这些行没有期末数可读，按「期初+本期借-本期贷」推出，不参与恒等式校验']);
      if (box.targetSkipped) help.push(['对方表没读到金额的行数', box.targetSkipped, '项目名读到了、金额是文字或 Excel 错误值（#REF! 等），这些行没有参与比对，见「单边项目」页']);
      if (box.bs.missing) help.push(['期末余额读不出的科目', box.bs.missing, '这些科目没有计入任何项目，在「未映射科目」页末尾列出']);
    }
    help.push([]);
    help.push(['怎么用这张底稿']);
    help.push(['1', '先看「未映射科目」：有余额的科目没归到报表项目，汇总一定对不上，先在上一步给它指定项目。']);
    help.push(['2', '再看「勾稽表」里差异不为 0 的项目，右侧「构成科目数」告诉你有几个科目凑成这个数，回工具里能展开看明细。']);
    help.push(['3', '最后看「单边项目」：一边有一边没有，通常是项目名称写法不同（如「应收票据」对「应收票据净额」）。']);
    help.push([]);
    help.push(['口径']);
    help.push(['· 金额一律按「分」整数计算，本表以元展示，保留两位小数。']);
    help.push(['· ' + (LEVEL_TEXT[box.level] || LEVEL_TEXT.leaf) + '，上级科目不重复计入。']);
    help.push(['· 备抵科目（累计折旧、减值准备、库存股）已按负向计入对应项目。']);
    help.push(['· 未映射科目不计入任何项目，单独列出——本工具不猜科目归属。']);
    help.push(['· 差异 = 科目余额表汇总 − 对方表。正数=科目侧多，负数=科目侧少。']);
    help.push(['· 单边项目不计入差异合计，但同样要处理；勾稽表末尾三行把两类分开列。']);
    help.push(['· 报表里的「流动资产合计」「资产总计」等合计行不参与逐项比对（它们是表内加总，不是项目）。']);
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

  /* ============================================================
   * 六、读明细表（一个报表项目一张工作表）
   *
   * 真实结构：`主要科目明细-XX.xlsx` 里每张表对应一个报表项目，
   * 上面几行是标题，中间一行是表头，最后有一行「合  计」（中间带空格），
   * 合计行下面还可能跟一行 check。
   * ============================================================ */

  var TOTAL_RE = /^(合计|小计|总计|合\s*计|小\s*计|总\s*计)$/;

  function isTotalLabel(text) {
    var t = str(text).replace(/[\s\u3000]/g, '');
    return /^(合计|小计|总计)$/.test(t);
  }

  function isCheckLabel(text) {
    var t = norm(text);
    return /^check|核对|勾稽|验证/.test(t);
  }

  /**
   * 认出一张明细表的结构：表头在第几行、有几列、有没有合计行。
   * @returns {{headerIndex, width, totalRowIndex, colSums:[], numericCols:[]}}
   */
  function detectDetailSheet(rows) {
    var width = 0;
    rows.forEach(function (r) { if (r && r.length > width) width = r.length; });

    // 表头：前 15 行里文字单元格最多的那一行（明细表的表头全是中文列名）
    var headerIndex = -1, bestTexts = 0;
    var limit = Math.min(rows.length, 15);
    for (var r = 0; r < limit; r++) {
      var texts = 0, nums = 0;
      (rows[r] || []).forEach(function (v) {
        if (typeof v === 'number') nums++;
        else if (str(v).trim()) texts++;
      });
      if (texts >= 3 && texts > bestTexts && texts > nums) { bestTexts = texts; headerIndex = r; }
    }

    // 合计行：表头下面第一行带「合计」字样的
    var totalRowIndex = -1;
    for (var r2 = (headerIndex >= 0 ? headerIndex + 1 : 0); r2 < rows.length; r2++) {
      var row = rows[r2] || [];
      for (var c = 0; c < row.length; c++) {
        if (isTotalLabel(row[c])) { totalRowIndex = r2; break; }
      }
      if (totalRowIndex >= 0) break;
    }

    // 每列求和：跳过表头以上、合计行、check 行
    var colSums = new Array(width).fill(0);
    var colCount = new Array(width).fill(0);
    for (var r3 = (headerIndex >= 0 ? headerIndex + 1 : 0); r3 < rows.length; r3++) {
      if (r3 === totalRowIndex) continue;
      var row3 = rows[r3] || [];
      if (isCheckLabel(row3[0])) continue;
      var hasTotal = row3.some(isTotalLabel);
      if (hasTotal) continue;
      for (var c3 = 0; c3 < width; c3++) {
        var v = toCents(row3[c3]);
        if (v.cents !== null && v.cents !== 0) { colSums[c3] += v.cents; colCount[c3]++; }
      }
    }

    var numericCols = [];
    for (var i = 0; i < width; i++) if (colCount[i] > 0) numericCols.push(i);

    // 列标题：表头行取文字；表头若是两级（下一行也全是文字），拼起来
    var headers = [];
    for (var h = 0; h < width; h++) {
      var t1 = str((rows[headerIndex] || [])[h]).trim();
      var t2 = str((rows[headerIndex + 1] || [])[h]).trim();
      var nextIsText = (rows[headerIndex + 1] || []).filter(function (v) {
        return typeof v !== 'number' && str(v).trim();
      }).length >= 3;
      headers.push(nextIsText && t2 && t2 !== t1 ? (t1 ? t1 + '·' + t2 : t2) : t1);
    }

    return {
      headerIndex: headerIndex,
      width: width,
      totalRowIndex: totalRowIndex,
      colSums: colSums,
      numericCols: numericCols,
      headers: headers
    };
  }

  /**
   * 从一张明细表里取出金额：优先用合计行的值，没有合计行就把明细加起来。
   * @param col 用第几列（0 起）；不给就挑绝对值最大的数字列
   */
  function readDetailAmount(rows, info, col) {
    info = info || detectDetailSheet(rows);
    var useCol = col;
    if (useCol === undefined || useCol === null) {
      var best = -1, bestAbs = -1;
      info.numericCols.forEach(function (c) {
        var v = Math.abs(info.colSums[c]);
        if (v > bestAbs) { bestAbs = v; best = c; }
      });
      useCol = best;
    }
    if (useCol < 0 || useCol >= info.width) {
      return { cents: null, col: useCol, source: 'none', reason: '这一列没有数字' };
    }

    if (info.totalRowIndex >= 0) {
      var tr = rows[info.totalRowIndex] || [];
      var tv = toCents(tr[useCol]);
      if (tv.cents !== null) {
        return { cents: tv.cents, col: useCol, source: '合计行', totalRow: info.totalRowIndex + 1 };
      }
    }
    if (info.colSums[useCol] === 0 && info.numericCols.indexOf(useCol) < 0) {
      return { cents: null, col: useCol, source: 'none', reason: '这一列没有数字' };
    }
    return { cents: info.colSums[useCol], col: useCol, source: '明细相加' };
  }

  /**
   * 明细表 ↔ 科目余额表汇总：按「工作表名 = 报表项目名」对上。
   * @param sheets [{name, rows}]
   * @param agg aggregateItems 的结果
   * @param opts { col: 用到第几列, getInfo: 可选 }
   */
  function reconcileDetail(sheets, agg, opts) {
    opts = opts || {};
    var left = {};
    agg.items.forEach(function (it) { left[normItem(it.item)] = it; });

    var matched = [], diffs = [], onlyDetail = [], onlyAgg = [], failed = [];
    var used = {};

    sheets.forEach(function (s) {
      var info = opts.getInfo ? opts.getInfo(s) : detectDetailSheet(s.rows);
      var got = readDetailAmount(s.rows, info, opts.col);
      if (got.cents === null) {
        failed.push({ sheet: s.name, reason: got.reason || '取不到金额' });
        return;
      }
      var key = normItem(s.name);
      var target = left[key];
      used[key] = true;
      var row = {
        item: s.name, detailCents: got.cents,
        aggCents: target ? target.cents : null,
        source: got.source, col: got.col,
        totalRow: got.totalRow || null,
        accounts: target ? target.accounts : []
      };
      if (!target) { onlyDetail.push(row); return; }
      row.diffCents = got.cents - target.cents;
      if (Math.abs(row.diffCents) <= (opts.toleranceCents || 0)) matched.push(row);
      else diffs.push(row);
    });

    agg.items.forEach(function (it) {
      if (!used[normItem(it.item)]) {
        // 0 余额的项目没有明细表很正常，不能因此判「未勾平」；
        // 而且全零的明细表会被当成「没有数字列」丢掉，这个缺口永远消不掉
        if ((it.cents || 0) !== 0) onlyAgg.push({ item: it.item, cents: it.cents, accounts: it.accounts });
      }
    });

    diffs.sort(function (a, b) { return Math.abs(b.diffCents) - Math.abs(a.diffCents); });
    onlyDetail.sort(function (a, b) { return Math.abs(b.detailCents) - Math.abs(a.detailCents); });
    onlyAgg.sort(function (a, b) { return Math.abs(b.cents) - Math.abs(a.cents); });

    return {
      matched: matched, diffs: diffs, onlyDetail: onlyDetail, onlyAgg: onlyAgg, failed: failed,
      sameCount: matched.length,
      diffTotalCents: diffs.reduce(function (s, d) { return s + d.diffCents; }, 0)
    };
  }

  return {
    VERSION: VERSION,
    toCents: toCents, moneyText: moneyText, norm: norm, str: str,
    CHART: CHART, CHART_MAP: CHART_MAP,
    detectBalanceHeader: detectBalanceHeader, parseBalanceRows: parseBalanceRows,
    balanceCheck: balanceCheck, identityCheck: identityCheck, leafAccounts: leafAccounts,
    detectStatementHeader: detectStatementHeader, isStatementTotal: isStatementTotal, parseStatementBlock: parseStatementBlock,
    detectDetailSheet: detectDetailSheet, readDetailAmount: readDetailAmount, reconcileDetail: reconcileDetail, isTotalLabel: isTotalLabel,
    buildSheets: buildSheets, buildWorkbook: buildWorkbook, fileName: fileName,
    codePrefix: codePrefix, mapAccount: mapAccount,
    contribution: contribution, aggregateItems: aggregateItems,
    normItem: normItem, compareItems: compareItems, classOfCode: classOfCode, CLASSES: CLASSES,
    ITEM_CLASS: ITEM_CLASS, itemClass: itemClass
  };
}));

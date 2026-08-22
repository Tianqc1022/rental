/* utils.js — 通用工具：日期、金额、HTML 转义、uid */
'use strict';
var utils = (function () {

  function pad2(n) { return n < 10 ? '0' + n : '' + n; }

  /** 解析 "YYYY-MM-DD" 为本地时区正午的 Date（避免 UTC 偏移错一天） */
  function parseDate(str) {
    if (!str) return null;
    var p = String(str).split('-').map(Number);
    if (p.length !== 3 || isNaN(p[0]) || isNaN(p[1]) || isNaN(p[2])) return null;
    return new Date(p[0], p[1] - 1, p[2], 12, 0, 0, 0);
  }

  /** 日期转整数 yyyymmdd，便于比较/求差 */
  function dayKey(d) { return d.getFullYear() * 10000 + (d.getMonth() + 1) * 100 + d.getDate(); }

  /** 某年某月的天数（m 为 1~12） */
  function lastDayOfMonth(y, m) { return new Date(y, m, 0).getDate(); }

  /** 整日差值：diffDays(a, b) = b - a 的天数（正午相减，精确） */
  function diffDays(a, b) { return Math.round((b.getTime() - a.getTime()) / 86400000); }

  function fmtDate(d) { return d.getFullYear() + '-' + pad2(d.getMonth() + 1) + '-' + pad2(d.getDate()); }
  function todayStr() { return fmtDate(new Date()); }

  /** "2026-08-10" → "8月10日" */
  function fmtDateCN(str) {
    var d = parseDate(str);
    return d ? (d.getMonth() + 1) + '月' + d.getDate() + '日' : '';
  }

  /** "2026-08" */
  function monthKeyOf(d) { return d.getFullYear() + '-' + pad2(d.getMonth() + 1); }

  function uid(prefix) { return prefix + '_' + Date.now() + '_' + Math.floor(Math.random() * 10000); }

  /** 金额格式化为千分位，如 3500 → 3,500.00 */
  function fmtMoney(n) {
    var v = Number(n) || 0;
    var s = v.toFixed(2);
    var dot = s.indexOf('.');
    var intPart = dot > -1 ? s.slice(0, dot) : s;
    intPart = intPart.replace(/\B(?=(\d{3})+(?!\d))/g, ',');
    return dot > -1 ? intPart + s.slice(dot) : intPart;
  }

  /** XSS 转义 */
  function escapeHtml(s) {
    if (s === null || s === undefined) return '';
    return String(s).replace(/[&<>"']/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
    });
  }

  return {
    pad2: pad2,
    parseDate: parseDate,
    dayKey: dayKey,
    lastDayOfMonth: lastDayOfMonth,
    diffDays: diffDays,
    fmtDate: fmtDate,
    todayStr: todayStr,
    fmtDateCN: fmtDateCN,
    monthKeyOf: monthKeyOf,
    uid: uid,
    fmtMoney: fmtMoney,
    escapeHtml: escapeHtml
  };
})();

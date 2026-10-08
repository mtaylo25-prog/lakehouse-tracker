// Shared number formatter (JavaScript). build/fmt.py is the identical Python version.
// Rules: USD with $ and thousands separators, negatives in parentheses, no hyphens or dashes anywhere.
'use strict';
function roundHalfUp(x, d) { const p = Math.pow(10, d); return Math.floor(Math.abs(x) * p + 0.5 + 1e-9) / p; }
function group(n) { let s = String(Math.trunc(n)); const out = []; while (s.length > 3) { out.unshift(s.slice(-3)); s = s.slice(0, -3); } out.unshift(s); return out.join(','); }
function fixed(x, d) {
  const v = roundHalfUp(x, d);
  if (d === 0) return group(v);
  let whole = Math.floor(v + 1e-9);
  let frac = Math.round((v - whole) * Math.pow(10, d));
  if (frac >= Math.pow(10, d)) { whole += 1; frac -= Math.pow(10, d); }
  return group(whole) + '.' + String(frac).padStart(d, '0');
}
function wrap(s, neg) { return neg ? '(' + s + ')' : s; }
function isNum(x) { return typeof x === 'number' && !Number.isNaN(x); }
function num(x, d = 0) { if (!isNum(x)) return 'n/a'; return wrap(fixed(x, d), x < 0 && roundHalfUp(x, d) !== 0); }
function money(x, d = 0) { if (!isNum(x)) return 'n/a'; return wrap('$' + fixed(x, d), x < 0 && roundHalfUp(x, d) !== 0); }
function money_k(x) { if (!isNum(x)) return 'n/a'; const v = x / 1000; return wrap('$' + fixed(v, 0) + 'K', x < 0 && roundHalfUp(v, 0) !== 0); }
function money_m(x) { if (!isNum(x)) return 'n/a'; const v = x / 1000000; return wrap('$' + fixed(v, 2) + 'M', x < 0 && roundHalfUp(v, 2) !== 0); }
function money_m_short(x) { if (!isNum(x)) return 'n/a'; const v = x / 1000000; if (Math.abs(v - Math.round(v)) < 1e-9) return wrap('$' + String(Math.abs(Math.round(v))) + 'M', x < 0); return money_m(x); }
function krw_m_short(x) { if (!isNum(x)) return 'n/a'; const v = x / 1000000; if (Math.abs(v - Math.round(v)) < 1e-9) return wrap('KRW ' + String(Math.abs(Math.round(v))) + 'M', x < 0); return krw_m(x); }
function money_auto(x) { if (!isNum(x)) return 'n/a'; const a = Math.abs(x); if (a >= 1000000) return money_m(x); if (a >= 10000) return money_k(x); return money(x); }
function pct(x, d = 1) { if (!isNum(x)) return 'n/a'; const v = x * 100; return wrap(fixed(v, d) + '%', x < 0 && roundHalfUp(v, d) !== 0); }
function pts(x, d = 1) { if (!isNum(x)) return 'n/a'; return wrap(fixed(x * 100, d) + ' points', x < 0); }
function month_label(m) { if (m === null || m === undefined || m === '' || m === 'none') return 'none'; const n = Number(m); if (Number.isNaN(n)) return String(m); return 'Month ' + String(Math.trunc(n)); }
function krw(x) { if (!isNum(x)) return 'n/a'; return wrap('KRW ' + fixed(x, 0), x < 0); }
function krw_m(x) { if (!isNum(x)) return 'n/a'; return wrap('KRW ' + fixed(x / 1000000, 1) + 'M', x < 0); }
function units(x) { return num(x, 0); }
module.exports = { num, money, money_k, money_m, money_m_short, money_auto, pct, pts, month_label, krw, krw_m, krw_m_short, units };
if (require.main === module) {
  const tests = [1166340, 153020, -19636, 0.131, -0.055, 1537200, 0, 8, null, 457804, 4576228, 0.177, 44.99, -8.96, 1234567.891, 999.5, 1049.5, 0.0005];
  const f = (fn) => tests.map((t) => (isNum(t) ? fn(t) : null));
  const out = { money_m: f(money_m), money_k: f(money_k), money: f(money), money2: tests.map((t) => (isNum(t) ? money(t, 2) : null)), pct: f(pct), num: f(num),
    month_label: [8, null, 'none', 5, 12].map(month_label), krw: [krw(24400000), krw_m(24400000)], money_auto: f(money_auto) };
  console.log(JSON.stringify(out));
}

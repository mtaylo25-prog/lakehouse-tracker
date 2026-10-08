"""Shared number formatter (Python). build/fmt.js is the identical JavaScript version.
Rules: USD with $ and thousands separators, negatives in parentheses, no hyphens or dashes anywhere.
money_m(1166340) -> "$1.17M"; money_k(153020) -> "$153K"; money(1166340) -> "$1,166,340";
pct(0.131) -> "13.1%"; month_label(8) -> "Month 8"; month_label(None) -> "none".
Rounding is half up on the absolute value using the same float arithmetic in both languages.
"""
import math

def _round_half_up(x, d):
    p = 10 ** d
    return math.floor(abs(x) * p + 0.5 + 1e-9) / p

def _group(n):
    # n is a non negative integer
    s = str(int(n))
    out = []
    while len(s) > 3:
        out.insert(0, s[-3:]); s = s[:-3]
    out.insert(0, s)
    return ",".join(out)

def _fixed(x, d):
    v = _round_half_up(x, d)
    if d == 0:
        return _group(v)
    whole = math.floor(v + 1e-9)
    frac = int(round((v - whole) * (10 ** d)))
    if frac >= 10 ** d:
        whole += 1; frac -= 10 ** d
    return _group(whole) + "." + str(frac).rjust(d, "0")

def _wrap(s, neg):
    return "(" + s + ")" if neg else s

def num(x, d=0):
    """Plain number with thousands separators and d decimals, negatives in parentheses."""
    if x is None: return "n/a"
    return _wrap(_fixed(x, d), x < 0 and _round_half_up(x, d) != 0)

def money(x, d=0):
    if x is None: return "n/a"
    return _wrap("$" + _fixed(x, d), x < 0 and _round_half_up(x, d) != 0)

def money_k(x):
    if x is None: return "n/a"
    v = x / 1000.0
    return _wrap("$" + _fixed(v, 0) + "K", x < 0 and _round_half_up(v, 0) != 0)

def money_m(x):
    if x is None: return "n/a"
    v = x / 1000000.0
    return _wrap("$" + _fixed(v, 2) + "M", x < 0 and _round_half_up(v, 2) != 0)

def money_m_short(x):
    """$1M for round millions, else money_m."""
    if x is None: return "n/a"
    v = x / 1000000.0
    if abs(v - round(v)) < 1e-9: return _wrap("$" + str(int(abs(round(v)))) + "M", x < 0)
    return money_m(x)

def krw_m_short(x):
    if x is None: return "n/a"
    v = x / 1000000.0
    if abs(v - round(v)) < 1e-9: return _wrap("KRW " + str(int(abs(round(v)))) + "M", x < 0)
    return krw_m(x)

def money_auto(x):
    """$1.17M above 1M, $153K above 10K, else $9,500."""
    if x is None: return "n/a"
    a = abs(x)
    if a >= 1000000: return money_m(x)
    if a >= 10000: return money_k(x)
    return money(x)

def pct(x, d=1):
    if x is None: return "n/a"
    v = x * 100.0
    return _wrap(_fixed(v, d) + "%", x < 0 and _round_half_up(v, d) != 0)

def pts(x, d=1):
    """Percentage points, e.g. 0.015 -> "1.5 points"."""
    if x is None: return "n/a"
    return _wrap(_fixed(x * 100.0, d) + " points", x < 0)

def month_label(m):
    if m is None or m == "" or m == "none": return "none"
    try:
        return "Month " + str(int(m))
    except (TypeError, ValueError):
        return str(m)

def krw(x):
    if x is None: return "n/a"
    return _wrap("KRW " + _fixed(x, 0), x < 0)

def krw_m(x):
    if x is None: return "n/a"
    return _wrap("KRW " + _fixed(x / 1000000.0, 1) + "M", x < 0)

def units(x):
    return num(x, 0)

if __name__ == "__main__":
    import json, sys
    tests = [1166340, 153020, -19636, 0.131, -0.055, 1537200, 0, 8, None, 457804, 4576228, 0.177, 44.99, -8.96, 1234567.891, 999.5, 1049.5, 0.0005]
    out = {"money_m": [money_m(t) if isinstance(t, (int, float)) else None for t in tests],
           "money_k": [money_k(t) if isinstance(t, (int, float)) else None for t in tests],
           "money": [money(t) if isinstance(t, (int, float)) else None for t in tests],
           "money2": [money(t, 2) if isinstance(t, (int, float)) else None for t in tests],
           "pct": [pct(t) if isinstance(t, (int, float)) else None for t in tests],
           "num": [num(t) if isinstance(t, (int, float)) else None for t in tests],
           "month_label": [month_label(t) for t in [8, None, "none", 5, 12]],
           "krw": [krw(24400000), krw_m(24400000)],
           "money_auto": [money_auto(t) if isinstance(t, (int, float)) else None for t in tests]}
    print(json.dumps(out))

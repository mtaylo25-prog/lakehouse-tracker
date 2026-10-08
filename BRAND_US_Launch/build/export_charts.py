#!/usr/bin/env python3
"""Exports chart PNGs (200 dpi) from build/figures.json into build/charts/.
White versions for documents (with titles), dark versions (suffix _dark.png, deck palette, no titles) for the decks.
No hyphens or dashes in any label: negatives render in parentheses. Rerun: python3 build/export_charts.py
"""
import json, os, sys
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.ticker import FuncFormatter
HERE = os.path.dirname(os.path.abspath(__file__)); ROOT = os.path.dirname(HERE)
sys.path.insert(0, HERE)
import fmt
F = json.load(open(os.path.join(HERE, "figures.json")))
OUTD = os.path.join(HERE, "charts"); os.makedirs(OUTD, exist_ok=True)
SCN = ["Low", "Medium", "High"]
CHANNELS = [("amazon_us", "Amazon US"), ("ebay_walmart", "eBay and Walmart"), ("b2b", "B2B installers"), ("dtc", "DTC site"), ("amazon_canada", "Amazon Canada"), ("pro_line", "Pro line")]

def theme(dark):
    if dark:
        return dict(bg="#0A0F1E", card="#141C36", text="#F3F6FF", muted="#9AA5C4", grid="#243052",
                    cat=["#2DE2FF", "#FFB547", "#8B5CF6", "#3DDC97", "#FF4FD8", "#C8D0E8"], scen={"Low": "#FFB547", "Medium": "#2DE2FF", "High": "#3DDC97"},
                    good="#3DDC97", warn="#FFB547", bad="#FF4FD8", accent="#2DE2FF", second="#8B5CF6")
    return dict(bg="#FFFFFF", card="#F4F6FA", text="#1B2235", muted="#5B6478", grid="#E1E5EE",
                cat=["#1193B5", "#B8700C", "#6D3FD0", "#0F8A5F", "#B8299A", "#7A7A1C"], scen={"Low": "#B8700C", "Medium": "#1193B5", "High": "#0F8A5F"},
                good="#0F8A5F", warn="#B8700C", bad="#B8299A", accent="#1193B5", second="#6D3FD0")

def money_tick(x, pos=None):
    return fmt.money_auto(x)

def setup(T, w=10, h=5.6):
    plt.rcParams.update({"font.family": "DejaVu Sans", "font.size": 11, "axes.unicode_minus": False, "figure.dpi": 100, "text.parse_math": False})
    fig, ax = plt.subplots(figsize=(w, h))
    fig.patch.set_facecolor(T["bg"]); ax.set_facecolor(T["bg"])
    for sp in ("top", "right"): ax.spines[sp].set_visible(False)
    for sp in ("left", "bottom"): ax.spines[sp].set_color(T["grid"])
    ax.tick_params(colors=T["muted"], labelsize=10)
    ax.yaxis.grid(True, color=T["grid"], linewidth=0.8); ax.set_axisbelow(True)
    ax.xaxis.label.set_color(T["muted"]); ax.yaxis.label.set_color(T["muted"])
    return fig, ax

def finish(fig, ax, T, name, dark, title, legend=True, ncol=3, loc="upper left"):
    if title and not dark:
        ax.set_title(title, color=T["text"], fontsize=13, loc="left", pad=14)
    if legend:
        anchor = {"upper left": (0, 1.02 if (dark or not title) else 1.0), "lower right": (1.0, 0.0)}[loc]
        lg = ax.legend(frameon=False, ncol=ncol, loc=loc, bbox_to_anchor=anchor, labelcolor=T["text"], fontsize=10)
        for t in lg.get_texts(): t.set_color(T["text"])
    fig.tight_layout()
    path = os.path.join(OUTD, name + ("_dark.png" if dark else ".png"))
    fig.savefig(path, dpi=200, facecolor=T["bg"]); plt.close(fig)
    return path

made = []
for dark in (False, True):
    T = theme(dark)
    # 1. revenue and net by month per scenario
    for s in SCN:
        mo = F["scenarios"][s]["monthly"]
        x = [m["month"] for m in mo]; rev = [m["revenue"] for m in mo]; net = [m["net"] for m in mo]
        fig, ax = setup(T)
        ax.plot(x, rev, color=T["accent"], linewidth=2, label="Revenue, all channels")
        ax.plot(x, net, color=T["second"], linewidth=2, label="Net, all channels")
        ax.axhline(0, color=T["muted"], linewidth=0.8)
        ax.yaxis.set_major_formatter(FuncFormatter(money_tick))
        ax.set_xticks(range(1, 25, 1)); ax.set_xticklabels([str(i) if i % 3 == 1 or i == 24 else "" for i in range(1, 25)])
        ax.set_xlabel("Model month (Year 1 is Months 1 to 12, Year 2 is Months 13 to 24)")
        ax.annotate(fmt.money_auto(rev[-1]), (24, rev[-1]), xytext=(6, 0), textcoords="offset points", color=T["text"], fontsize=10, va="center")
        ax.annotate(fmt.money_auto(net[-1]), (24, net[-1]), xytext=(6, 0), textcoords="offset points", color=T["text"], fontsize=10, va="center")
        ax.set_xlim(0.5, 26.5)
        made.append(finish(fig, ax, T, f"revenue_net_monthly_{s.lower()}", dark, f"{s} case: revenue and net by month, Years 1 and 2", ncol=2))
    # 2. annual revenue by channel (Medium) stacked
    A = F["scenarios"]["Medium"]["annual"]
    years = [f"Year {y}" for y in range(1, 6)]
    fig, ax = setup(T)
    bottom = [0] * 5
    for i, (ch, label) in enumerate(CHANNELS):
        vals = [A[f"Y{y}"]["revenue_by_channel"][ch] for y in range(1, 6)]
        ax.bar(years, vals, bottom=bottom, color=T["cat"][i], label=label, width=0.62, edgecolor=T["bg"], linewidth=1.5)
        bottom = [b + v for b, v in zip(bottom, vals)]
    for i, tot in enumerate(bottom):
        ax.annotate(fmt.money_auto(tot), (i, tot), xytext=(0, 5), textcoords="offset points", ha="center", color=T["text"], fontsize=10)
    ax.yaxis.set_major_formatter(FuncFormatter(money_tick)); ax.set_ylim(0, max(bottom) * 1.18)
    made.append(finish(fig, ax, T, "annual_revenue_by_channel_medium", dark, "Medium case: annual revenue by channel, Years 1 to 5", ncol=3))
    # 2b. three band comparison (revenue grouped bars) and net
    for metric, label, name in (("revenue", "revenue", "annual_revenue_three_bands"), ("net", "net", "annual_net_three_bands")):
        fig, ax = setup(T)
        w = 0.26
        for j, s in enumerate(SCN):
            vals = [F["scenarios"][s]["annual"][f"Y{y}"][metric] for y in range(1, 6)]
            xs = [i + (j - 1) * w for i in range(5)]
            ax.bar(xs, vals, width=w - 0.03, color=T["scen"][s], label=s)
            for xx, v in zip(xs, vals):
                ax.annotate(fmt.money_auto(v), (xx, v), xytext=(0, 4 if v >= 0 else -12), textcoords="offset points", ha="center", color=T["text"], fontsize=8)
        ax.set_xticks(range(5)); ax.set_xticklabels(years)
        ax.axhline(0, color=T["muted"], linewidth=0.8)
        ax.yaxis.set_major_formatter(FuncFormatter(money_tick))
        lo = min(min(F["scenarios"][s]["annual"][f"Y{y}"][metric] for y in range(1, 6)) for s in SCN); hi = max(max(F["scenarios"][s]["annual"][f"Y{y}"][metric] for y in range(1, 6)) for s in SCN)
        # negative labels sit below the bar, so the axis floor is pushed down enough that they clear the tick text
        floor = min(0, lo) - (0.08 * (hi - min(0, lo)) if lo < 0 else 0)
        ax.set_ylim(floor, hi * 1.18)
        made.append(finish(fig, ax, T, name, dark, f"Annual {label} by scenario, Years 1 to 5 (Low, Medium, High)", ncol=3))
    # 3. cash balance monthly, three scenarios, minimum marked
    fig, ax = setup(T)
    for s in SCN:
        S = F["scenarios"][s]
        x = list(range(0, 25)); y = [S["cash_balance_month0"]] + [m["cash_balance"] for m in S["monthly"]]
        ax.plot(x, y, color=T["scen"][s], linewidth=2, label=s)
        mi = min(range(25), key=lambda i: y[i])
        ax.plot([mi], [y[mi]], marker="o", markersize=8, color=T["scen"][s], markeredgecolor=T["bg"], markeredgewidth=1.5)
        off = {"Low": (-12, -16, "right"), "Medium": (0, -18, "center"), "High": (14, 26, "left")}[s]
        ax.annotate(f"{s} minimum {fmt.money_auto(y[mi])} at Month {mi}", (mi, y[mi]), xytext=(off[0], off[1]), textcoords="offset points", ha=off[2], color=T["text"], fontsize=9)
    ax.axhline(0, color=T["muted"], linewidth=0.8)
    ax.yaxis.set_major_formatter(FuncFormatter(money_tick))
    ax.set_xticks(range(0, 25, 2)); ax.set_xlabel("Model month (Month 0 is the pre launch purchase)")
    ax.set_ylim(ax.get_ylim()[0] * 1.25, ax.get_ylim()[1] * 1.1)
    made.append(finish(fig, ax, T, "cash_balance_monthly", dark, "Cash balance by month before new funding, minimum marked (peak cash need)", ncol=3))
    # 4. tornado
    tor = F["sensitivity"]["tornado"]; base = F["sensitivity"]["base_y5_net"]
    fig, ax = setup(T, 10, 5.2)
    labels = [t["driver"] for t in tor][::-1]
    lows = [t["low_minus_base"] for t in tor][::-1]; highs = [t["high_minus_base"] for t in tor][::-1]
    lolab = [t["low_label"] for t in tor][::-1]; hilab = [t["high_label"] for t in tor][::-1]
    yy = range(len(labels))
    ax.barh(yy, lows, color=T["warn"], height=0.6, label="Downside setting")
    ax.barh(yy, highs, color=T["good"], height=0.6, label="Upside setting")
    span = max(abs(min(lows)), abs(max(highs)))
    for i in yy:
        if abs(lows[i]) > 0.45 * span:
            ax.annotate(f"{lolab[i]}: {fmt.money_auto(lows[i])}", (lows[i], i), xytext=(6, 0), textcoords="offset points", ha="left", va="center", color="#1B2235", fontsize=8.5)
        else:
            ax.annotate(f"{lolab[i]}: {fmt.money_auto(lows[i])}", (lows[i], i), xytext=(-4, 0), textcoords="offset points", ha="right", va="center", color=T["text"], fontsize=8.5)
        ax.annotate(f"{hilab[i]}: {fmt.money_auto(highs[i])}", (highs[i], i), xytext=(4, 0), textcoords="offset points", ha="left", va="center", color=T["text"], fontsize=8.5)
    ax.set_yticks(list(yy)); ax.set_yticklabels(labels, color=T["text"])
    ax.axvline(0, color=T["muted"], linewidth=0.8)
    ax.xaxis.set_major_formatter(FuncFormatter(money_tick)); ax.yaxis.grid(False); ax.xaxis.grid(True, color=T["grid"], linewidth=0.8)
    ax.set_xlim(-span * 1.25, span * 1.75)
    ax.set_xlabel(f"Change in Year 5 net vs the Medium base of {fmt.money_auto(base)}")
    made.append(finish(fig, ax, T, "sensitivity_tornado", dark, "Sensitivity of Year 5 net (Medium case): one way, largest swing first", ncol=1, loc="lower right"))
    # 5. duty waterfall (hero net per unit)
    wf = F["duty"]["waterfall"]
    fig, ax = setup(T, 9, 5)
    cols = []
    for i, w in enumerate(wf):
        lab = w["label"]
        c = T["accent"] if "net" in lab.lower() and "uplift" not in lab.lower() and "penalty" not in lab.lower() else (T["good"] if "uplift" in lab.lower() else T["bad"])
        ax.bar(i, w["value"], bottom=w["base"], color=c, width=0.6)
        top = w["base"] + w["value"]
        txt = fmt.money(w["value"], 2) if ("net" in lab.lower() and "uplift" not in lab.lower() and "penalty" not in lab.lower()) else ("plus " if "uplift" in lab.lower() else "minus ") + fmt.money(abs(w["value"]), 2)
        ax.annotate(txt, (i, top), xytext=(0, 5), textcoords="offset points", ha="center", color=T["text"], fontsize=10)
    ax.set_xticks(range(len(wf))); ax.set_xticklabels([w["label"] for w in wf], color=T["text"], fontsize=9.5)
    ax.yaxis.set_major_formatter(FuncFormatter(lambda x, p: fmt.money(x, 2)))
    ax.set_ylim(0, max(w["base"] + w["value"] for w in wf) * 1.22)
    rates = {d: F["duty"][d]["rate"] for d in ("copper", "fiber", "stress")}
    ax.set_xlabel(f"Hero 8K 10M at $44.99, Medium base. Duty: copper {fmt.pct(rates['copper'])}, fiber {fmt.pct(rates['fiber'])}, stress {fmt.pct(rates['stress'])}")
    made.append(finish(fig, ax, T, "duty_waterfall", dark, "Duty scenario waterfall: hero net per unit", legend=False))
    # 6. eighteen month scenarios: cumulative net (Amazon US) with breakeven markers
    fig, ax = setup(T)
    for s in SCN:
        S = F["scenarios"][s]; mo = S["monthly"][:18]
        x = [m["month"] for m in mo]; y = [m["amazon_cum_net"] for m in mo]
        ax.plot(x, y, color=T["scen"][s], linewidth=2, label=f"{s}: revenue {fmt.money_auto(S['m18']['revenue'])}, net {fmt.money_auto(S['m18']['net'])}")
        be = S["m18"]["breakeven_month"]
        if be:
            ax.plot([be], [y[be - 1]], marker="o", markersize=8, color=T["scen"][s], markeredgecolor=T["bg"], markeredgewidth=1.5)
            ax.annotate(f"breakeven Month {be}", (be, y[be - 1]), xytext=(0, 10), textcoords="offset points", ha="center", color=T["text"], fontsize=9)
        ax.annotate(fmt.money_auto(y[-1]), (18, y[-1]), xytext=(6, 0), textcoords="offset points", color=T["text"], fontsize=9.5, va="center")
    ax.axhline(0, color=T["muted"], linewidth=0.8)
    ax.yaxis.set_major_formatter(FuncFormatter(money_tick))
    ax.set_xticks(range(1, 19)); ax.set_xlabel("Model month"); ax.set_xlim(0.5, 20.5)
    ax.set_ylabel("Cumulative net, Amazon US (August basis)")
    made.append(finish(fig, ax, T, "eighteen_month_scenarios", dark, "Eighteen month Amazon model: cumulative net by scenario", ncol=1))
print("\n".join(os.path.relpath(p, ROOT) for p in made))

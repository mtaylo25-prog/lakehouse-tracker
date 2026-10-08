"""Recalculates the workbook with LibreOffice and checks the calibration targets (spec 2.3 and 2.5), the tie out
between the annual and monthly tabs, the Scenario Summary engines against the main tabs, the pure Python replica and the realism rules.
Usage: python3 build/verify_model.py [workbook]   (build_model.py calls it after saving)."""
import json, os, subprocess, sys
HERE = os.path.dirname(os.path.abspath(__file__)); ROOT = os.path.dirname(HERE)
sys.path.insert(0, HERE)
import model_core as mc
RECALC = ["python3", "/mnt/skills/public/xlsx/scripts/recalc.py"]

def recalc(path, timeout=180):
    out = subprocess.run(RECALC + [path, str(timeout)], capture_output=True, text=True)
    try:
        return json.loads(out.stdout)
    except json.JSONDecodeError:
        return {"error": out.stdout + out.stderr}

class Reader:
    def __init__(self, path):
        from openpyxl import load_workbook
        self.wb = load_workbook(path, data_only=True)
        self.names = json.load(open(os.path.join(HERE, "figures_map.json")))
    def __call__(self, name):
        ref = self.names[name]; sh, a = ref.rsplit("!", 1); sh = sh.strip("'").replace("''", "'"); a = a.replace("$", "")
        ws = self.wb[sh]
        if ":" in a:
            return [c.value for row in ws[a] for c in row]
        return ws[a].value

def within(a, b, tol):
    if b == 0: return abs(a) <= 1e-6
    return abs(a - b) / abs(b) <= tol

def verify(path):
    v = Reader(path)
    fails = []; notes = []
    def check(cond, msg):
        (notes if cond else fails).append(("PASS " if cond else "FAIL ") + msg)
    # 1. August reconciliation (spec 2.3) from the Scenario Summary engines
    targets = {"Medium": (1166340, 153020, 8, 1537200, 0.172, 0.01), "Low": (633156, -19636, None, None, None, 0.02), "High": (1850935, 457804, 5, 2439648, None, 0.02)}
    brief = {"Low": (633000, -19000), "Medium": (1170000, 153000), "High": (1850000, 455000)}
    for sname, (rev, net, be, rr, em, tol) in targets.items():
        xr = v(f"ss_{sname}_m18_revenue"); xn = v(f"ss_{sname}_m18_net"); xb = v(f"ss_{sname}_m18_breakeven_month")
        check(within(xr, rev, tol), f"{sname} 18 month revenue {xr:,.0f} vs spec {rev:,.0f} (tol {tol:.0%})")
        if sname == "High":
            check(within(xn, brief["High"][1], tol), f"High 18 month net {xn:,.0f} vs brief 455,000 (tol {tol:.0%}); the original spec figure 457,804 assumed 7% ads in Months 13 to 18, below the 8% floor; spec 2.3 now reads 447,155")
            check(within(xn, 447155, 0.001), f"High 18 month net {xn:,.0f} equals the corrected spec 2.3 figure 447,155")
        else:
            check(within(xn, net, tol), f"{sname} 18 month net {xn:,.0f} vs spec {net:,.0f} (tol {tol:.0%})")
        check((xb == be) if be else (xb == "none"), f"{sname} breakeven month {xb} vs spec {be}")
        if rr: check(within(v(f"ss_{sname}_m18_run_rate"), rr, tol), f"{sname} run rate {v(f'ss_{sname}_m18_run_rate'):,.0f} vs spec {rr:,.0f}")
        if em: check(abs(v(f"ss_{sname}_m18_exit_margin") - em) <= 0.002, f"{sname} exit margin {v(f'ss_{sname}_m18_exit_margin'):.3f} vs spec {em}")
    # 2. Year 5 bands (spec 2.5)
    bands = {"Medium": (4.0e6, 5.0e6, 0.10, 0.15), "High": (8.0e6, 10.0e6, None, None), "Low": (0.7e6, 1.2e6, None, 0.05)}
    for sname, (lo, hi, mlo, mhi) in bands.items():
        r5 = v(f"ss_{sname}_annual_Y5_revenue"); m5 = v(f"ss_{sname}_annual_Y5_net_margin")
        check(lo <= r5 <= hi, f"{sname} Year 5 revenue {r5:,.0f} in band {lo:,.0f} to {hi:,.0f}")
        if mlo is not None: check(mlo <= m5 <= mhi, f"{sname} Year 5 net margin {m5:.1%} in band {mlo:.0%} to {mhi:.0%}")
        if sname == "Low": check(m5 <= mhi, f"Low Year 5 net margin {m5:.1%} near zero (at or below 5%)")
    check(v("ss_Low_gate_month12") == "KILL", f"Low Month 12 gate is {v('ss_Low_gate_month12')} (must be KILL)")
    # 3. Annual ties to monthly (active)
    for k in ["annual_check_revenue_Y1", "annual_check_revenue_Y2", "annual_check_net_Y1", "annual_check_net_Y2"]:
        check(abs(v(k)) < 0.01, f"{k} = {v(k)}")
    # 4. Scenario Summary Medium engine equals the main tabs (selector at 2)
    if v("Scenario") == 2:
        for k in ["m18_revenue", "m18_net", "annual_Y1_revenue", "annual_Y5_revenue", "annual_Y5_net", "cash_peak_need", "cash_funding_gap"]:
            a = v(k); b = v(f"ss_Medium_{k}")
            check(within(a, b, 1e-6), f"main tab {k} {a:,.2f} equals Scenario Summary Medium {b:,.2f}")
    # 5. Pure Python replica
    for s, sname in enumerate(mc.SCEN_NAMES):
        R = mc.evaluate(s)
        pairs = [("m18_revenue", R["m18"]["revenue"]), ("m18_net", R["m18"]["net"]), ("annual_Y5_revenue", R["annual"][5]["revenue"]), ("annual_Y5_net", R["annual"][5]["net"]), ("cash_peak_need", R["cash"]["peak_need"]), ("cash_funding_gap", R["cash"]["funding_gap"])]
        for k, pyv in pairs:
            check(within(v(f"ss_{sname}_{k}"), pyv, 0.001), f"python replica {sname} {k}: workbook {v(f'ss_{sname}_{k}'):,.0f} vs python {pyv:,.0f}")
    # 6. Realism rules on the drivers
    for key in ["ads_p1", "ads_p2", "ads_p3", "ads_p4", "ads_p5", "ads_y3", "ads_y4", "ads_y5"]:
        check(min(mc.D[key]["vals"]) >= 0.08, f"ad spend never below 8%: {key} {mc.D[key]['vals']}")
    for key in ["returns_y1", "returns_y2"]:
        check(min(mc.D[key]["vals"]) >= 0.02, f"returns never below 2%: {key} {mc.D[key]['vals']}")
    check(min(mc.CH[k] for k in ("ebay_returns", "b2b_returns", "dtc_returns", "ca_returns", "pro_amz_returns")) >= 0.02, "channel returns never below 2%")
    # company level ad spend floor (all channels) and net margin compression, every scenario year, from the replica
    for s, sname in enumerate(mc.SCEN_NAMES):
        A = mc.evaluate(s)["annual"]
        blended = [A[y]["ads_t"] / A[y]["revenue"] for y in range(1, 6)]
        check(min(blended) >= 0.08, f"{sname} blended ad spend (all channels) never below 8% of revenue, Years 1 to 5: {[round(b*100,2) for b in blended]}")
        margins = [A[y]["margin"] for y in range(2, 6)]
        if sname != "Low":
            check(all(margins[i] - margins[i + 1] >= 0.01 for i in range(3)), f"{sname} net margin compresses at least 1 point each year, Years 2 to 5: {[round(m*100,2) for m in margins]}")
        else:
            check(A[5]["net"] <= 0, f"Low Year 5 net {A[5]['net']:,.0f} is at or below zero")
    check(mc.D["lost_q_factor"]["vals"][0] < 1 and mc.D["lost_q_start"]["vals"][0] <= 22, "Low case includes a lost quarter (three months at 10% of trend)")
    for sname in ["Medium", "High"]:
        c2 = v(f"ss_{sname}_annual_Y2_net"); 
    # compression: Amazon US contribution margin after ads falls at least 1 point a year on average from Year 2 to Year 5, and never rises
    for s, sname in enumerate(mc.SCEN_NAMES):
        R = mc.evaluate(s); A = R["annual"]
        con = [A[y]["amz"]["con"] / A[y]["amz"]["rev"] for y in range(2, 6)]
        if sname == "Low":
            check(con[3] < con[1], f"Low Amazon contribution margin falls Year 3 to Year 5: {[round(c*100,1) for c in con]}")
        else:
            check((con[0] - con[3]) / 3 >= 0.01 and all(con[i + 1] <= con[i] for i in range(3)), f"{sname} Amazon contribution margin after ads compresses at least 1 point a year, Years 2 to 5: {[round(c*100,1) for c in con]}")
    # 7. hero economics
    check(abs(v("ue_hero_net") - 8.0) < 0.01 and abs(v("ue_hero_net_pct") - 0.178) < 0.001, f"hero net {v('ue_hero_net'):.3f} (17.8%)")
    check(abs(v("duty_fiber_landed_pct_medium") - 0.336) < 0.005, f"Medium plus Fiber landed {v('duty_fiber_landed_pct_medium'):.3%} near 33.6%")
    check(v("cost_base_alibaba_low_net") < 0, f"Alibaba cost base hero net {v('cost_base_alibaba_low_net'):.2f} is negative at $44.99")
    for line in notes + fails: print(line)
    print(f"verify: {len(notes)} passed, {len(fails)} failed")
    return not fails

if __name__ == "__main__":
    path = sys.argv[1] if len(sys.argv) > 1 else os.path.join(ROOT, "output", "03_Financial_Model_5Y.xlsx")
    if "--recalc" in sys.argv:
        print(json.dumps(recalc(path)))
    sys.exit(0 if verify(path) else 1)

"""Single source of drivers for the 5 year model plus a pure Python evaluator.
The workbook formulas are the authority; this evaluator writes the inputs and cross checks
the recalculated workbook. Every driver carries a basis tag (Evidence based or Management estimate).
No hyphens in any label string.
"""
from __future__ import annotations
import math

def xround(x, d=0):
    """Excel style ROUND (half away from zero)."""
    p = 10 ** d
    return math.floor(abs(x) * p + 0.5 + 1e-9) / p * (1 if x >= 0 else -1)

EB = "Evidence based"
ME = "Management estimate"

SCEN_NAMES = ["Low", "Medium", "High"]
DUTY_NAMES = ["Copper", "Fiber", "Stress"]

# ---------------------------------------------------------------- SKUs (spec 2.1)
SKUS = [
    # id, spec (internal), length m, H2 cost KRW, US list, US sale (ASP), weight lb, FBA fee, launch
    ("8K_7M",  "HDMI 2.1 8K hybrid AOC", 7,  10525, 49.99, 39.99, 0.35, 4.90, "Yes"),
    ("8K_10M", "HDMI 2.1 8K hybrid AOC (hero)", 10, 11470, 59.99, 44.99, 0.45, 4.90, "Yes"),
    ("8K_15M", "HDMI 2.1 8K hybrid AOC", 15, 13376, 69.99, 54.99, 0.60, 5.50, "Yes"),
    ("8K_20M", "HDMI 2.1 8K hybrid AOC", 20, 15266, 79.99, 64.99, 0.75, 5.50, "Yes"),
    ("8K_25M", "HDMI 2.1 8K hybrid AOC", 25, 17172, 89.99, 74.99, 0.90, 6.60, "No"),
    ("8K_50M", "HDMI 2.1 8K hybrid AOC", 50, 27349, 149.99, 119.99, 1.90, 8.40, "No"),
    ("4K_15M", "HDMI 2.0 4K hybrid AOC", 15, 12680, 44.99, 34.99, 0.60, 5.50, "No"),
    ("4K_20M", "HDMI 2.0 4K hybrid AOC", 20, 14122, 49.99, 39.99, 0.75, 5.50, "No"),
    ("4K_30M", "HDMI 2.0 4K hybrid AOC", 30, 19542, 64.99, 49.99, 1.10, 6.60, "Yes"),
    ("4K_40M", "HDMI 2.0 4K hybrid AOC", 40, 22609, 74.99, 59.99, 1.50, 8.40, "No"),
]
HERO = "8K_10M"

# ---------------------------------------------------------------- global inputs
FX_DEFAULT = 1430
FX_SCEN = [1300, 1430, 1550]
FOB_FACTOR = 0.92          # removes Korea inbound freight and handling (ME)
AIR_PER_KG = 8.50
OCEAN_PER_KG = 1.10
LB_TO_KG = 0.4536
MPF_HMF_OCEAN = 0.0035
BROKER_PER_UNIT = 0.40
INBOUND_PER_UNIT = 0.50
QC_PER_UNIT = 0.25
ALIBABA_LOW, ALIBABA_HIGH = 17.0, 22.0

# duty table (spec 2.4): name, HTS, MFN, 301 list, 301 forced labor, pending 301 add
DUTY = [
    ("Copper", "8544.42.90", 0.026, 0.25, 0.125, 0.0, EB),
    ("Fiber",  "8544.70.00", 0.0,   0.0,  0.125, 0.0, "Evidence based rates; List 3 coverage of 8544.70.00 is a broker check [PLACEHOLDER]"),
    ("Stress", "8544.42.90 plus pending second Section 301 action", 0.026, 0.25, 0.125, 0.10, ME),
]
DUTY_FIBER_LIST3_CHECK = ("Fiber if also on List 3", "8544.70.00", 0.0, 0.25, 0.125, 0.0)

# August hero unit economics (EB)
HERO_AUG = dict(price=44.99, landed_pct=0.38, referral=0.15, fba=5.50, ads_pct=0.12, returns_storage_pct=0.05, korea_gross=1.86)

# ---------------------------------------------------------------- 18 month block (spec 2.3)
MEDIUM_UNITS_18 = [150, 240, 360, 500, 650, 820, 1000, 1200, 1400, 1600, 1800, 2050, 2250, 2450, 2600, 2750, 2900, 3050]
SHARED = dict(referral=0.15, fba=5.50, storage=0.02, fixed_m=850.0, seller_fee=39.99, launch=9500.0)

# scenario drivers: value per scenario [Low, Medium, High], with basis
D = {}
def drv(key, label, unit, vals, basis, note=""):
    D[key] = dict(label=label, unit=unit, vals=list(vals), basis=basis, note=note)

drv("unit_mult", "Unit ramp multiplier on the Medium ramp, Months 1 to 18", "x", [0.60, 1.0, 1.55], ME, "Low 0.60 and High 1.55 per the build spec; Medium ramp is the August deck")
drv("asp_y12", "ASP Years 1 and 2", "$", [38.0, 42.0, 43.0], EB, "August model Medium $42; Low and High calibrated per spec 2.3")
drv("landed_base", "Landed COGS % of revenue, base (copper duty, air then ocean)", "%", [0.43, 0.38, 0.32], EB, "August model 38% pending the customs ruling; Low 43%; High 32% assumes fiber ruling, ocean and volume pricing")
drv("returns_y1", "Returns % Year 1", "%", [0.04, 0.03, 0.025], EB, "August model 3%")
drv("returns_y2", "Returns % Year 2 onward (from Month 19)", "%", [0.04, 0.025, 0.02], ME, "never below 2%")
drv("ads_p1", "Ad spend % Months 1 to 3", "%", [0.37, 0.30, 0.26], EB, "August model 30%; Low plus 7 points; High minus 4 points")
drv("ads_p2", "Ad spend % Months 4 to 6", "%", [0.29, 0.22, 0.18], EB, "")
drv("ads_p3", "Ad spend % Months 7 to 12", "%", [0.22, 0.15, 0.11], EB, "")
drv("ads_p4", "Ad spend % Months 13 to 18", "%", [0.18, 0.11, 0.08], EB, "August model 11% for Months 13 to 18; Low plus 7 points; High minus 3 points (the 8% floor)")
drv("ads_p5", "Ad spend % Months 19 to 24", "%", [0.18, 0.11, 0.095], ME, "Low and Medium carry the Months 13 to 18 rate; High steps to its Year 3 rate so the Year 2 ad spend across all channels stays above the 8% floor (8.1% blended; at 8% Amazon only it diluted to 7.4%)")
drv("ads_y3", "Ad spend % Year 3", "%", [0.13, 0.115, 0.095], ME, "spec starting values 12, 9.5, 9.0; raised so competition compresses margin at least 1 point a year; never below 8%")
drv("ads_y4", "Ad spend % Year 4", "%", [0.135, 0.115, 0.095], ME, "spec starting values 12, 9.0, 8.5; see Year 3 note")
drv("ads_y5", "Ad spend % Year 5", "%", [0.165, 0.115, 0.095], ME, "spec starting values 12, 8.5, 8.0; see Year 3 note; never below 8%; Low 16.5% so the Low Year 5 net is at or below zero (the listing buys rank it cannot earn from reviews)")
drv("y2_start", "Year 2 driver start month (ocean landed cost, fixed cost step up, compliance, FBA fee inflation, Year 2 returns and the Months 19 to 24 ad rate begin here)", "month", [19, 19, 19], EB, "Month 19 is the first month after the August 18 month horizon; the compliance charge is spread over the months from here to Month 24")
drv("m19_add", "Months 19 to 24 units added per month on top of Month 18", "units", [0, 100, 160], ME, "Low flat at Month 18")
drv("lost_q_factor", "Lost quarter factor (listing suppression or account health event)", "x", [0.10, 1.0, 1.0], ME, "Low case only: units at 10% of trend for three months")
drv("lost_q_start", "Lost quarter start month (three months)", "month", [19, 99, 99], ME, "Spec 2.5 said Months 15 to 17; moved to Months 19 to 21 so the Low 18 month block reproduces the August Low case, per the brief")
drv("g_y3", "Amazon US unit growth Year 3", "%", [0.10, 0.30, 0.35], ME, "spec starting values Low 5, Medium 30, High 35; Low raised 5 points (within the ±5 bound) so Year 5 lands in the 0.7 to 1.2M band")
drv("g_y4", "Amazon US unit growth Year 4", "%", [0.08, 0.20, 0.25], ME, "Low raised 5 points within bound")
drv("g_y5", "Amazon US unit growth Year 5", "%", [0.05, 0.12, 0.15], ME, "Low raised 5 points within bound")
drv("asp_chg", "ASP change per year from Year 3 (price competition)", "%", [-0.04, -0.04, -0.03], ME, "spec starting values minus 3, minus 2, minus 1.5 per year; deepened so contribution margin falls at least 1 point a year overall and Year 5 net lands in band")
drv("roadmap_uplift", "Roadmap SKU mix uplift to ASP in Year 3 (HDMI 2.2 16K, DP 2.1, USB4 AOC)", "%", [0.0, 0.03, 0.03], ME, "Low has no capital for roadmap SKUs")
drv("landed_chg_y2", "Landed COGS % change vs base, Year 2 (from Month 19)", "pts", [0.0, -0.015, -0.010], ME, "ocean freight and volume")
drv("landed_chg_y3", "Landed COGS % change vs base, Year 3", "pts", [0.0, -0.015, -0.015], ME, "spec starting value Medium minus 2.5; held at the Year 2 benefit")
drv("landed_chg_y4", "Landed COGS % change vs base, Year 4", "pts", [0.0, -0.010, -0.010], ME, "spec starting value Medium minus 2.0; creep back starts in Year 4 so the Medium net margin compresses at least 1 point from Year 3 to Year 4")
drv("landed_chg_y5", "Landed COGS % change vs base, Year 5", "pts", [0.0, -0.005, 0.0], ME, "creep back toward base (factory price increases, freight volatility); spec starting values minus 1.5 and minus 0.5")
drv("fba_infl", "FBA fee inflation per year from Year 2", "%", [0.03, 0.03, 0.03], ME, "")
drv("fixed_y2", "Fixed costs per month Year 2 (from Month 19)", "$", [1200.0, 1500.0, 1950.0], ME, "High 1.3 times Medium")
drv("fixed_y3", "Fixed costs per month Year 3", "$", [1200.0, 4000.0, 5200.0], ME, "Medium adds Mark consulting retainer 2,000, virtual assistant 1,000, tools and insurance 1,000; Low frozen")
drv("fixed_y4", "Fixed costs per month Year 4", "$", [1200.0, 7000.0, 9100.0], ME, "")
drv("fixed_y5", "Fixed costs per month Year 5", "$", [1200.0, 10000.0, 13000.0], ME, "")
drv("ebay_pct_y2", "eBay and Walmart overflow, % of Amazon US revenue, Year 2", "%", [0.08, 0.08, 0.08], ME, "")
drv("ebay_pct_y3", "eBay and Walmart overflow, % of Amazon US revenue, Year 3", "%", [0.08, 0.10, 0.10], ME, "")
drv("ebay_pct_y4", "eBay and Walmart overflow, % of Amazon US revenue, Year 4", "%", [0.08, 0.12, 0.12], ME, "")
drv("ebay_pct_y5", "eBay and Walmart overflow, % of Amazon US revenue, Year 5", "%", [0.08, 0.12, 0.12], ME, "")
drv("b2b_y2", "B2B installer and AV integrator revenue, Year 2", "$", [0.0, 120000.0, 216000.0], ME, "High 1.8 times Medium")
drv("b2b_y3", "B2B installer and AV integrator revenue, Year 3", "$", [60000.0, 350000.0, 630000.0], ME, "Low 60K flat from Year 3")
drv("b2b_y4", "B2B installer and AV integrator revenue, Year 4", "$", [60000.0, 600000.0, 1080000.0], ME, "")
drv("b2b_y5", "B2B installer and AV integrator revenue, Year 5", "$", [60000.0, 850000.0, 1530000.0], ME, "")
drv("dtc_y3", "DTC site revenue, Year 3", "$", [0.0, 100000.0, 180000.0], ME, "")
drv("dtc_y4", "DTC site revenue, Year 4", "$", [0.0, 250000.0, 450000.0], ME, "")
drv("dtc_y5", "DTC site revenue, Year 5", "$", [0.0, 400000.0, 720000.0], ME, "")
drv("ca_y4", "Amazon Canada revenue, Year 4", "$", [0.0, 120000.0, 216000.0], ME, "")
drv("ca_y5", "Amazon Canada revenue, Year 5", "$", [0.0, 250000.0, 450000.0], ME, "")
drv("pro_y4", "Pro line revenue, Year 4", "$", [0.0, 150000.0, 270000.0], ME, "")
drv("pro_y5", "Pro line revenue, Year 5", "$", [0.0, 350000.0, 630000.0], ME, "")
drv("cover", "Inventory cover target, months of forward units", "months", [3.0, 2.5, 2.0], ME, "")
drv("ocean_from", "Ocean freight from Month (before that air)", "month", [99, 7, 7], ME, "after the Day 90 gate for proven SKUs; Low stays on air")
drv("lead_air", "Air freight lead time, months (order to arrival)", "months", [1, 1, 1], ME, "spec 2.6")
drv("lead_ocean", "Ocean freight lead time, months (order to arrival)", "months", [2, 2, 2], ME, "spec 2.6")
drv("inv_days", "Inventory days of landed COGS, Years 3 to 5", "days", [135, 95, 85], ME, "spec 2.6 starting values 90, 75, 60; raised to the days implied by the Month 24 stock (cover months plus one lead time in transit) so Year 3 does not release cash artificially")
drv("fixed_y1", "Fixed costs per month Year 1 (software, insurance, misc), Months 1 to 18", "$", [850.0, 850.0, 850.0], EB, "August model")

# channel economics (same for every scenario)
CH = dict(
    ebay_referral=0.13, ebay_mcf=6.50, ebay_ads=0.05, ebay_returns=0.03,
    b2b_price_factor=0.60, b2b_3pl=4.0, b2b_commission=0.05, b2b_marketing=0.03, b2b_returns=0.02,   # spec 2.5 started at 1.5%; the brief's floor (returns never below 2%) wins
    dtc_payment=0.03, dtc_ship=11.0, dtc_ads=0.15, dtc_returns=0.03,
    ca_asp_factor=0.92, ca_fba=6.50, ca_referral=0.15, ca_ads=0.12, ca_returns=0.03, ca_duty=0.0,
    pro_asp=120.0, pro_landed=0.40, pro_b2b_share=0.60, pro_fba=8.40, pro_amz_ads=0.12, pro_amz_returns=0.03, pro_storage=0.02,
)
COMPLIANCE = dict(form5472=800.0, wyoming=60.0, bond=500.0, insurance=1200.0, trademark_y5=0.0)
COMPLIANCE_Y1_NOTE = "Year 1 compliance is inside the August fixed 850 per month and the 9,500 launch costs"
OPENING_CAPITAL = 39000.0
DEPOSIT_PCT = 0.30
PAYOUT_LAG_DAYS = 14
RECV_DAYS_ANNUAL = 14
FUNDING_CUSHION = 1.2

# Plan B (spec 2.7)
PLANB = dict(b2b=[120000.0, 300000.0, 450000.0], korea_deals=4, korea_units=1400, korea_gp_unit=2.50, korea_growth=0.15, fixed_m=[850.0, 1200.0, 1200.0])

# Gate thresholds (spec 2.7): key, label, unit, value. Inputs on Assumptions; the Funding tab and the Scenario Summary engines read them.
GATES = [
    ("g_d90_conv_pass", "Day 90: unit session conversion, pass at or above (3 of 5 launch SKUs; observed)", "%", 0.08),
    ("g_d90_conv_kill", "Day 90: hero conversion, kill below (observed)", "%", 0.05),
    ("g_d90_tacos_pass", "Day 90: TACOS in Month 3, pass at or below", "%", 0.30),
    ("g_d90_tacos_kill", "Day 90: TACOS in Month 3, kill above", "%", 0.40),
    ("g_d90_ret_pass", "Day 90: return rate, pass at or below", "%", 0.04),
    ("g_d90_ret_kill", "Day 90: return rate, kill above", "%", 0.07),
    ("g_d90_cpu_pass", "Day 90: contribution per unit before advertising, pass at or above", "$", 11.0),
    ("g_d90_cpu_kill", "Day 90: hero contribution per unit before advertising, kill below", "$", 8.0),
    ("g_d90_run_pass", "Day 90: cash runway after funding, months, pass at or above", "months", 4.0),
    ("g_d90_run_kill", "Day 90: cash runway, months, kill below", "months", 2.0),
    ("g_d90_ip_days", "Day 90: suppressed listing or IP complaint open more than this many days is a kill (observed)", "days", 14),
    ("g_m12_rev_pass", "Month 12: monthly revenue, pass at or above", "$", 70000.0),
    ("g_m12_rev_kill", "Month 12: monthly revenue, kill below", "$", 50000.0),
    ("g_m12_cum_pass", "Month 12: cumulative net, pass at or above", "$", 0.0),
    ("g_m12_cum_kill", "Month 12: cumulative net, kill below", "$", -25000.0),
    ("g_m12_tacos_pass", "Month 12: TACOS, pass at or below", "%", 0.15),
    ("g_m12_tacos_kill", "Month 12: TACOS, kill above", "%", 0.25),
    ("g_m12_ret_pass", "Month 12: return rate, pass at or below", "%", 0.03),
    ("g_m12_ret_kill", "Month 12: return rate, kill above", "%", 0.05),
    ("g_m12_gap_kill", "Month 12: funding gap above this with three negative operating cash flow months is a kill", "$", 50000.0),
]

# August deal totals (evidence based) read from inputs/august_deal_totals.csv so the FX tab does not carry literals
def _load_deal_totals():
    import csv, os
    path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "inputs", "august_deal_totals.csv")
    out = {}
    with open(path, newline="") as f:
        for row in csv.DictReader(f):
            out[row["metric"]] = row
    def val(k): return float(out[k]["value"])
    return dict(units=val("units_sold"), cost_krw=val("total_cost"), revenue_krw=val("revenue_ex_vat"), gross_profit_krw=val("gross_profit"), gross_margin=val("gross_margin"), source=out["total_cost"]["source"])
DEAL = _load_deal_totals()

# ---------------------------------------------------------------- helpers
def duty_rate(i):
    d = DUTY[i]
    return d[2] + d[3] + d[4] + d[5]

def hero_fob_usd(fx=FX_DEFAULT):
    krw = [s for s in SKUS if s[0] == HERO][0][3]
    return krw / fx * FOB_FACTOR

def hero_fob_share(fx=FX_DEFAULT):
    return hero_fob_usd(FX_DEFAULT) / HERO_AUG["price"]

def landed_active(s, duty=0, fx=FX_DEFAULT, freight_delta=0.0):
    share = hero_fob_share()
    return D["landed_base"]["vals"][s] + (duty_rate(duty) - duty_rate(0)) * share + freight_delta + share * (FX_DEFAULT / fx - 1)

def units_24(s):
    v = D["unit_mult"]["vals"][s]
    u = [xround(x * v) for x in MEDIUM_UNITS_18]
    add = D["m19_add"]["vals"][s]
    for k in range(1, 7):
        u.append(u[17] + add * k)
    lq = D["lost_q_factor"]["vals"][s]; st = D["lost_q_start"]["vals"][s]
    for m in range(1, 25):
        if st <= m <= st + 2:
            u[m - 1] = xround(u[m - 1] * lq)
    return u

def ads_pct_month(s, m):
    if m <= 3: return D["ads_p1"]["vals"][s]
    if m <= 6: return D["ads_p2"]["vals"][s]
    if m <= 12: return D["ads_p3"]["vals"][s]
    if m < D["y2_start"]["vals"][s]: return D["ads_p4"]["vals"][s]
    return D["ads_p5"]["vals"][s]

def evaluate(s, duty=0, fx=FX_DEFAULT):
    """Pure Python replica of the workbook engine for scenario index s (0 Low, 1 Medium, 2 High)."""
    V = lambda k: D[k]["vals"][s]
    asp = V("asp_y12")
    base = landed_active(s, duty, fx)
    units = units_24(s)
    fob_u = hero_fob_usd(fx)
    R = dict(month=[])
    rows = {k: [] for k in ["units", "asp", "rev", "cogs", "ref", "fba", "ret", "sto", "ads", "contrib", "fixed", "fee", "launch", "comp", "net", "cum", "margin", "tacos", "cpu_b", "cpu_a",
                            "eb_rev", "eb_units", "eb_cogs", "eb_ref", "eb_mcf", "eb_ads", "eb_ret", "eb_con",
                            "b2b_rev", "b2b_units", "b2b_cogs", "b2b_3pl", "b2b_com", "b2b_mkt", "b2b_ret", "b2b_con",
                            "tot_rev", "tot_net", "tot_cum", "landed_pct", "ads_pct", "ret_pct", "fba_fee"]}
    cum = 0.0; tcum = 0.0
    y2s = int(V("y2_start")); comp_months = 24 - y2s + 1
    for m in range(1, 25):
        u = units[m - 1]
        y2 = m >= y2s
        lp = base + (V("landed_chg_y2") if y2 else 0.0)
        ap = ads_pct_month(s, m)
        rp = V("returns_y2") if y2 else V("returns_y1")
        fee = SHARED["fba"] * ((1 + V("fba_infl")) if y2 else 1.0)
        rev = u * asp
        cogs = rev * lp; ref = rev * SHARED["referral"]; fba = u * fee; ret = rev * rp; sto = rev * SHARED["storage"]; ads = rev * ap
        contrib = rev - cogs - ref - fba - ret - sto - ads
        fixed = V("fixed_y2") if y2 else V("fixed_y1")
        feem = SHARED["seller_fee"]
        launch = SHARED["launch"] if m == 1 else 0.0
        comp = (sum(v for k, v in COMPLIANCE.items() if k != "trademark_y5") / comp_months) if y2 else 0.0
        net = contrib - fixed - feem - launch - comp
        cum += net
        # Year 2 channels (Months 13 to 24)
        if m >= 13:
            eb_rev = rev * V("ebay_pct_y2"); eb_units = eb_rev / asp
            eb_cogs = eb_rev * lp; eb_ref = eb_rev * CH["ebay_referral"]; eb_mcf = eb_units * CH["ebay_mcf"]; eb_ads = eb_rev * CH["ebay_ads"]; eb_ret = eb_rev * CH["ebay_returns"]
            eb_con = eb_rev - eb_cogs - eb_ref - eb_mcf - eb_ads - eb_ret
            b2b_rev = V("b2b_y2") / 12.0; b2b_units = b2b_rev / (asp * CH["b2b_price_factor"])
            b2b_cogs = b2b_units * asp * lp; b2b_3pl = b2b_units * CH["b2b_3pl"]; b2b_com = b2b_rev * CH["b2b_commission"]; b2b_mkt = b2b_rev * CH["b2b_marketing"]; b2b_ret = b2b_rev * CH["b2b_returns"]
            b2b_con = b2b_rev - b2b_cogs - b2b_3pl - b2b_com - b2b_mkt - b2b_ret
        else:
            eb_rev = eb_units = eb_cogs = eb_ref = eb_mcf = eb_ads = eb_ret = eb_con = 0.0
            b2b_rev = b2b_units = b2b_cogs = b2b_3pl = b2b_com = b2b_mkt = b2b_ret = b2b_con = 0.0
        tot_rev = rev + eb_rev + b2b_rev
        tot_net = net + eb_con + b2b_con
        tcum += tot_net
        vals = dict(units=u, asp=asp, rev=rev, cogs=cogs, ref=ref, fba=fba, ret=ret, sto=sto, ads=ads, contrib=contrib, fixed=fixed, fee=feem, launch=launch, comp=comp,
                    net=net, cum=cum, margin=(net / rev if rev else 0.0), tacos=(ads / rev if rev else 0.0), cpu_b=((contrib + ads) / u if u else 0.0), cpu_a=(contrib / u if u else 0.0),
                    eb_rev=eb_rev, eb_units=eb_units, eb_cogs=eb_cogs, eb_ref=eb_ref, eb_mcf=eb_mcf, eb_ads=eb_ads, eb_ret=eb_ret, eb_con=eb_con,
                    b2b_rev=b2b_rev, b2b_units=b2b_units, b2b_cogs=b2b_cogs, b2b_3pl=b2b_3pl, b2b_com=b2b_com, b2b_mkt=b2b_mkt, b2b_ret=b2b_ret, b2b_con=b2b_con,
                    tot_rev=tot_rev, tot_net=tot_net, tot_cum=tcum, landed_pct=lp, ads_pct=ap, ret_pct=rp, fba_fee=fee)
        for k in rows: rows[k].append(vals[k])
    R.update(rows)
    # 18 month block (Amazon US, August basis)
    rev18 = sum(rows["rev"][:18]); net18 = sum(rows["net"][:18])
    be = next((m for m in range(1, 19) if rows["cum"][m - 1] >= 0), None)
    R["m18"] = dict(revenue=rev18, net=net18, net_margin=net18 / rev18, breakeven_month=be, run_rate=rows["rev"][17] * 12, exit_margin=rows["margin"][17], units=sum(units[:18]))
    # ---------------- annual
    A = {}
    def ysum(key, lo, hi): return sum(rows[key][lo - 1:hi])
    for y in (1, 2):
        lo, hi = (1, 12) if y == 1 else (13, 24)
        A[y] = dict(
            amz=dict(units=ysum("units", lo, hi), rev=ysum("rev", lo, hi), cogs=ysum("cogs", lo, hi), fees=ysum("ref", lo, hi), ful=ysum("fba", lo, hi), ads=ysum("ads", lo, hi), ret=ysum("ret", lo, hi) + ysum("sto", lo, hi), con=ysum("contrib", lo, hi)),
            ebay=dict(units=ysum("eb_units", lo, hi), rev=ysum("eb_rev", lo, hi), cogs=ysum("eb_cogs", lo, hi), fees=ysum("eb_ref", lo, hi), ful=ysum("eb_mcf", lo, hi), ads=ysum("eb_ads", lo, hi), ret=ysum("eb_ret", lo, hi), con=ysum("eb_con", lo, hi)),
            b2b=dict(units=ysum("b2b_units", lo, hi), rev=ysum("b2b_rev", lo, hi), cogs=ysum("b2b_cogs", lo, hi), fees=ysum("b2b_com", lo, hi), ful=ysum("b2b_3pl", lo, hi), ads=ysum("b2b_mkt", lo, hi), ret=ysum("b2b_ret", lo, hi), con=ysum("b2b_con", lo, hi)),
            dtc=dict(units=0, rev=0, cogs=0, fees=0, ful=0, ads=0, ret=0, con=0), ca=dict(units=0, rev=0, cogs=0, fees=0, ful=0, ads=0, ret=0, con=0), pro=dict(units=0, rev=0, cogs=0, fees=0, ful=0, ads=0, ret=0, con=0),
            fixed=ysum("fixed", lo, hi), fee=ysum("fee", lo, hi), launch=ysum("launch", lo, hi), comp=ysum("comp", lo, hi),
            asp=asp, landed=(base if y == 1 else (ysum("cogs", lo, hi) / ysum("rev", lo, hi))))
    prev_units = A[2]["amz"]["units"]; prev_asp = asp
    comp_y = sum(v for k, v in COMPLIANCE.items() if k != "trademark_y5")
    for y in (3, 4, 5):
        g = V(f"g_y{y}"); units_y = prev_units * (1 + g); prev_units = units_y
        asp_y = prev_asp * (1 + V("asp_chg")) * ((1 + V("roadmap_uplift")) if y == 3 else 1.0); prev_asp = asp_y
        lp = base + V(f"landed_chg_y{y}")
        fee = SHARED["fba"] * (1 + V("fba_infl")) ** (y - 1)
        ap = V(f"ads_y{y}"); rp = V("returns_y2")
        rev = units_y * asp_y
        amz = dict(units=units_y, rev=rev, cogs=rev * lp, fees=rev * SHARED["referral"], ful=units_y * fee, ads=rev * ap, ret=rev * rp + rev * SHARED["storage"])
        amz["con"] = rev - sum(amz[k] for k in ("cogs", "fees", "ful", "ads", "ret"))
        eb_rev = rev * V(f"ebay_pct_y{y}"); ebu = eb_rev / asp_y
        ebay = dict(units=ebu, rev=eb_rev, cogs=eb_rev * lp, fees=eb_rev * CH["ebay_referral"], ful=ebu * CH["ebay_mcf"], ads=eb_rev * CH["ebay_ads"], ret=eb_rev * CH["ebay_returns"])
        ebay["con"] = eb_rev - sum(ebay[k] for k in ("cogs", "fees", "ful", "ads", "ret"))
        b_rev = V(f"b2b_y{y}"); bu = b_rev / (asp_y * CH["b2b_price_factor"])
        b2b = dict(units=bu, rev=b_rev, cogs=bu * asp_y * lp, fees=b_rev * CH["b2b_commission"], ful=bu * CH["b2b_3pl"], ads=b_rev * CH["b2b_marketing"], ret=b_rev * CH["b2b_returns"])
        b2b["con"] = b_rev - sum(b2b[k] for k in ("cogs", "fees", "ful", "ads", "ret"))
        d_rev = V(f"dtc_y{y}"); du = d_rev / asp_y
        dtc = dict(units=du, rev=d_rev, cogs=d_rev * lp, fees=d_rev * CH["dtc_payment"], ful=du * CH["dtc_ship"], ads=d_rev * CH["dtc_ads"], ret=d_rev * CH["dtc_returns"])
        dtc["con"] = d_rev - sum(dtc[k] for k in ("cogs", "fees", "ful", "ads", "ret"))
        c_rev = V(f"ca_y{y}") if y >= 4 else 0.0; cu = c_rev / (asp_y * CH["ca_asp_factor"]) if c_rev else 0.0
        ca = dict(units=cu, rev=c_rev, cogs=cu * asp_y * lp * (1 + CH["ca_duty"]), fees=c_rev * CH["ca_referral"], ful=cu * CH["ca_fba"], ads=c_rev * CH["ca_ads"], ret=c_rev * CH["ca_returns"])
        ca["con"] = c_rev - sum(ca[k] for k in ("cogs", "fees", "ful", "ads", "ret"))
        p_rev = V(f"pro_y{y}") if y >= 4 else 0.0; pu = p_rev / CH["pro_asp"]
        pb = CH["pro_b2b_share"]; pa = 1 - pb
        pro = dict(units=pu, rev=p_rev, cogs=p_rev * CH["pro_landed"],
                   fees=p_rev * pb * CH["b2b_commission"] + p_rev * pa * SHARED["referral"],
                   ful=pu * pb * CH["b2b_3pl"] + pu * pa * CH["pro_fba"],
                   ads=p_rev * pb * CH["b2b_marketing"] + p_rev * pa * CH["pro_amz_ads"],
                   ret=p_rev * pb * CH["b2b_returns"] + p_rev * pa * (CH["pro_amz_returns"] + CH["pro_storage"]))
        pro["con"] = p_rev - sum(pro[k] for k in ("cogs", "fees", "ful", "ads", "ret"))
        A[y] = dict(amz=amz, ebay=ebay, b2b=b2b, dtc=dtc, ca=ca, pro=pro, fixed=V(f"fixed_y{y}") * 12, fee=SHARED["seller_fee"] * 12, launch=0.0,
                    comp=comp_y + (COMPLIANCE["trademark_y5"] if y == 5 else 0.0), asp=asp_y, landed=lp)
    for y in A:
        a = A[y]
        chans = [a[c] for c in ("amz", "ebay", "b2b", "dtc", "ca", "pro")]
        a["revenue"] = sum(c["rev"] for c in chans); a["cogs_t"] = sum(c["cogs"] for c in chans); a["fees_t"] = sum(c["fees"] for c in chans)
        a["ful_t"] = sum(c["ful"] for c in chans); a["ads_t"] = sum(c["ads"] for c in chans); a["ret_t"] = sum(c["ret"] for c in chans)
        a["con_t"] = sum(c["con"] for c in chans)
        a["net"] = a["con_t"] - a["fixed"] - a["fee"] - a["launch"] - a["comp"]
        a["margin"] = a["net"] / a["revenue"] if a["revenue"] else 0.0
    R["annual"] = A
    # ---------------- cash (monthly 0 to 24, then annual 3 to 5)
    cover = V("cover"); ocean_from = V("ocean_from"); g3 = V("g_y3")
    tot_units = [rows["units"][m - 1] + rows["eb_units"][m - 1] + rows["b2b_units"][m - 1] for m in range(1, 25)]
    fut = tot_units[23] * (1 + g3)
    ext = tot_units + [fut] * 6  # months 25 to 30
    landed_u = [rows["landed_pct"][m - 1] * asp for m in range(1, 25)] + [(base + V("landed_chg_y2")) * asp] * 6
    nonfob_u = [lu - fob_u for lu in landed_u]
    N = 26
    target = [cover * sum(ext[m:m + 3]) / 3.0 for m in range(1, N + 1)]  # ext index m = month m+1
    closing = [0.0] * (N + 1); arrivals = [0.0] * (N + 1)
    for m in range(1, N + 1):
        opening = closing[m - 1]; sales = ext[m - 1]
        arrivals[m] = max(0.0, target[m - 1] - (opening - sales))
        closing[m] = opening + arrivals[m] - sales
    lead = [0] + [int(V("lead_ocean")) if m >= ocean_from else int(V("lead_air")) for m in range(1, N + 1)]
    ordm = [0] + [m - lead[m] for m in range(1, N + 1)]
    shipm = [0] + [m - 1 for m in range(1, N + 1)]
    fobv = [0.0] + [arrivals[m] * fob_u for m in range(1, N + 1)]
    nonfobv = [0.0] + [arrivals[m] * nonfob_u[m - 1] for m in range(1, N + 1)]
    cash = {}
    bal = OPENING_CAPITAL; balances = []; recv_prev = 0.0; opcf = []
    lag = PAYOUT_LAG_DAYS / 30.0
    inv_cash = []; cash_in = []; cash_out = []; onhand = []; intransit = []; recv = []
    for m in range(0, 25):
        dep = DEPOSIT_PCT * sum(fobv[a] for a in range(1, N + 1) if ordm[a] == m)
        balp = (1 - DEPOSIT_PCT) * sum(fobv[a] for a in range(1, N + 1) if shipm[a] == m)
        nf = nonfobv[m] if 1 <= m <= 24 else 0.0
        inv = dep + balp + nf
        if m == 0:
            proceeds = 0.0; opex = 0.0; r = 0.0
        else:
            i = m - 1
            proceeds = (rows["rev"][i] - rows["ref"][i] - rows["fba"][i] - rows["ret"][i] - rows["sto"][i] - rows["ads"][i]
                        + rows["eb_rev"][i] - rows["eb_ref"][i] - rows["eb_mcf"][i] - rows["eb_ads"][i] - rows["eb_ret"][i]
                        + rows["b2b_rev"][i] - rows["b2b_3pl"][i] - rows["b2b_com"][i] - rows["b2b_mkt"][i] - rows["b2b_ret"][i])
            opex = rows["fixed"][i] + rows["fee"][i] + rows["launch"][i] + rows["comp"][i]
            r = proceeds * lag
        cin = proceeds - r + recv_prev
        cout = inv + opex
        bal = bal + cin - cout
        balances.append(bal); opcf.append(cin - cout); inv_cash.append(inv); cash_in.append(cin); cash_out.append(cout); recv.append(r)
        onhand.append(closing[m] * landed_u[m - 1] if m >= 1 else 0.0)
        it = sum(fobv[a] * (DEPOSIT_PCT if (ordm[a] <= m < shipm[a]) else (1.0 if shipm[a] <= m < a else 0.0)) for a in range(1, N + 1))
        intransit.append(it)
        recv_prev = r
    # annual Y3 to Y5
    inv_prev = onhand[24] + intransit[24]; recv_prevy = recv[24]; balance = balances[24]
    ann_bal = {}
    for y in (3, 4, 5):
        a = A[y]
        inv_y = a["cogs_t"] * V("inv_days") / 365.0
        recv_y = a["revenue"] * RECV_DAYS_ANNUAL / 365.0
        balance = balance + a["net"] - (inv_y - inv_prev) - (recv_y - recv_prevy)
        ann_bal[y] = dict(balance=balance, inventory=inv_y, receivable=recv_y)
        inv_prev, recv_prevy = inv_y, recv_y
    allbal = balances + [ann_bal[y]["balance"] for y in (3, 4, 5)]
    minbal = min(allbal); imin = allbal.index(minbal)
    min_label = imin if imin <= 24 else f"Year {imin - 22}"
    peak_need = max(0.0, -minbal)
    gap = FUNDING_CUSHION * peak_need
    first_neg = next((m for m, b in enumerate(balances) if b < 0), None)
    need_by = (first_neg - 1) if first_neg is not None else None
    burn = sum(opcf[4:7]) / 3.0
    runway = (balances[3] / -burn) if burn < 0 else "funded"
    R["cash"] = dict(balances=balances, opcf=opcf, inv_cash=inv_cash, cash_in=cash_in, cash_out=cash_out, onhand=onhand, intransit=intransit, recv=recv,
                     arrivals=arrivals, closing=closing, target=target, min_balance=minbal, min_month=min_label, peak_need=peak_need, funding_gap=gap,
                     need_by=need_by, runway=runway, annual=ann_bal, tot_units=tot_units)
    return R

if __name__ == "__main__":
    for s in range(3):
        R = evaluate(s)
        m = R["m18"]; A = R["annual"]; C = R["cash"]
        print(f"== {SCEN_NAMES[s]}: 18m rev {m['revenue']:,.0f} net {m['net']:,.0f} ({m['net_margin']:.1%}) BE {m['breakeven_month']} runrate {m['run_rate']:,.0f} exit {m['exit_margin']:.1%}")
        for y in range(1, 6):
            a = A[y]
            print(f"   Y{y} rev {a['revenue']:,.0f} net {a['net']:,.0f} ({a['margin']:.1%}) amz {a['amz']['rev']:,.0f} amz con% {a['amz']['con']/a['amz']['rev']:.1%} fixed {a['fixed']:,.0f}")
        print(f"   cash min {C['min_balance']:,.0f} at {C['min_month']} peak need {C['peak_need']:,.0f} gap {C['funding_gap']:,.0f} need by {C['need_by']} runway {C['runway']}")
        print("   balances:", [round(b / 1000) for b in C["balances"]])

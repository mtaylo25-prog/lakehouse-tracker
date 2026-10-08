#!/usr/bin/env python3
"""Pulls every figure the decks and documents need from the recalculated workbook into build/figures.json.
Sets the Scenario selector to 1, 2, 3 in turn on temporary copies (build/qa/model/), recalculates each with LibreOffice,
reads values back with openpyxl data_only=True. The delivered workbook is never modified (it stays at Scenario 2 Medium, Duty 1 Copper, FX 1430).
Also writes build/figures_keys.json (JSON key path -> defined name and selector state) and build/figures_readme.md.
Rerun: python3 build/pull_figures.py
"""
import json, os, shutil, subprocess, sys
from openpyxl import load_workbook
HERE = os.path.dirname(os.path.abspath(__file__)); ROOT = os.path.dirname(HERE)
sys.path.insert(0, HERE)
import model_core as mc
from verify_model import recalc, Reader
WB = os.path.join(ROOT, "output", "03_Financial_Model_5Y.xlsx")
QA = os.path.join(HERE, "qa", "model"); os.makedirs(QA, exist_ok=True)
NAMES = json.load(open(os.path.join(HERE, "figures_map.json")))
FACTS = json.load(open(os.path.join(HERE, "facts.json")))
KEYS = {}   # json path -> {"name": defined name, "scenario": selector or "any"}

def key(path, name, scenario="any"):
    KEYS[path] = {"name": name, "ref": NAMES[name], "scenario": scenario}

def num(x):
    if isinstance(x, (int, float)): return x
    return None

def pull_scenario(sel):
    """Return the scenario block for selector value sel (1 Low, 2 Medium, 3 High)."""
    sname = mc.SCEN_NAMES[sel - 1]
    tmp = os.path.join(QA, f"model_scenario_{sel}.xlsx")
    shutil.copy(WB, tmp)
    wb = load_workbook(tmp)
    wb["Assumptions"]["C3"] = sel; wb["Assumptions"]["C4"] = 1; wb["Assumptions"]["C5"] = mc.FX_DEFAULT
    wb.save(tmp)
    res = recalc(tmp)
    assert res.get("status") == "success" and res.get("total_errors") == 0, res
    v = Reader(tmp)
    assert v("Scenario") == sel and v("ScenarioName") == sname, (v("Scenario"), v("ScenarioName"))
    S = {}
    be = v("m18_breakeven_month")
    S["m18"] = dict(revenue=v("m18_revenue"), net=v("m18_net"), net_margin=v("m18_net_margin"), breakeven_month=(be if isinstance(be, (int, float)) else None),
                    run_rate=v("m18_run_rate"), exit_margin=v("m18_exit_margin"), units=v("m18_units"))
    for k, nm in [("revenue", "m18_revenue"), ("net", "m18_net"), ("net_margin", "m18_net_margin"), ("breakeven_month", "m18_breakeven_month"), ("run_rate", "m18_run_rate"), ("exit_margin", "m18_exit_margin"), ("units", "m18_units")]:
        key(f"scenarios.{sname}.m18.{k}", nm, sel)
    be24 = v("m24_breakeven_month")
    S["m24"] = dict(revenue=v("m24_revenue"), net=v("m24_net"), breakeven_month=(be24 if isinstance(be24, (int, float)) else None))
    for k, nm in [("revenue", "m24_revenue"), ("net", "m24_net"), ("breakeven_month", "m24_breakeven_month")]: key(f"scenarios.{sname}.m24.{k}", nm, sel)
    rows = {k: v(nm) for k, nm in [("units", "monthly_units"), ("revenue", "monthly_revenue"), ("net", "monthly_net"), ("cum_net", "monthly_cum_net"), ("tacos", "monthly_tacos"),
                                   ("contribution_before_ads", "monthly_contribution_before_ads"), ("contribution_after_ads", "monthly_contribution_after_ads"),
                                   ("amazon_revenue", "monthly_amazon_revenue"), ("amazon_net", "monthly_amazon_net"), ("amazon_cum_net", "monthly_amazon_cum_net"), ("amazon_net_margin", "monthly_amazon_net_margin"),
                                   ("asp", "monthly_asp"), ("landed_pct", "monthly_landed_pct"), ("ads_pct", "monthly_ads_pct"), ("returns_pct", "monthly_returns_pct"),
                                   ("ebay_walmart_revenue", "monthly_ebay_walmart_revenue"), ("b2b_revenue", "monthly_b2b_revenue"),
                                   ("inventory_on_hand", "monthly_inventory_on_hand"), ("inventory_in_transit", "monthly_inventory_in_transit"), ("inventory_cash_out", "monthly_inventory_cash_out"),
                                   ("arrivals", "monthly_arrivals"), ("closing_stock", "monthly_closing_stock"), ("operating_cash_flow", "monthly_operating_cash_flow"), ("net_cash_flow", "monthly_net_cash_flow")]}
    bal = v("monthly_cash_balance"); balf = v("monthly_cash_balance_after_funding")
    monthly = []
    for m in range(1, 25):
        d = {"month": m}
        for k, arr in rows.items():
            d[k] = arr[m - 1] if len(arr) == 24 else arr[m]
        d["cash_balance"] = bal[m]; d["cash_balance_after_funding"] = balf[m]
        monthly.append(d)
    S["monthly"] = monthly
    S["cash_balance_month0"] = bal[0]
    for k, nm in [("units", "monthly_units"), ("revenue", "monthly_revenue"), ("net", "monthly_net"), ("cum_net", "monthly_cum_net"), ("cash_balance", "monthly_cash_balance"), ("tacos", "monthly_tacos"), ("contribution_before_ads", "monthly_contribution_before_ads"), ("contribution_after_ads", "monthly_contribution_after_ads"), ("amazon_revenue", "monthly_amazon_revenue"), ("amazon_net", "monthly_amazon_net"), ("cash_balance_after_funding", "monthly_cash_balance_after_funding"), ("inventory_on_hand", "monthly_inventory_on_hand"), ("inventory_in_transit", "monthly_inventory_in_transit")]:
        key(f"scenarios.{sname}.monthly[].{k}", nm, sel)
    annual = {}
    for y in range(1, 6):
        a = {k: v(f"annual_Y{y}_{nm}") for k, nm in [("revenue", "revenue"), ("net", "net"), ("net_margin", "net_margin"), ("cogs", "cogs"), ("fees", "fees"), ("fulfillment", "fulfillment"), ("advertising", "advertising"), ("returns", "returns"), ("contribution", "contribution"), ("fixed", "fixed"), ("seller_fee", "seller_fee"), ("launch", "launch"), ("compliance", "compliance"), ("cum_net", "cum_net"), ("amazon_units", "amazon_units"), ("asp", "asp"), ("landed_pct", "landed_pct"), ("ads_pct", "ads_pct"), ("amazon_contribution_pct", "amazon_contribution_pct")]}
        a["revenue_by_channel"] = {ch: v(f"annual_Y{y}_revenue_by_channel_{ch}") for ch in ["amazon_us", "ebay_walmart", "b2b", "dtc", "amazon_canada", "pro_line"]}
        a["contribution_by_channel"] = {ch: v(f"annual_Y{y}_contribution_by_channel_{ch}") for ch in ["amazon_us", "ebay_walmart", "b2b", "dtc", "amazon_canada", "pro_line"]}
        a["fees_and_fulfillment"] = a["fees"] + a["fulfillment"]
        annual[f"Y{y}"] = a
        for k in ["revenue", "net", "net_margin", "cogs", "fees", "fulfillment", "advertising", "returns", "contribution", "fixed", "seller_fee", "launch", "compliance", "cum_net", "amazon_units", "asp", "landed_pct", "ads_pct", "amazon_contribution_pct"]:
            key(f"scenarios.{sname}.annual.Y{y}.{k}", f"annual_Y{y}_{k}", sel)
        for ch in ["amazon_us", "ebay_walmart", "b2b", "dtc", "amazon_canada", "pro_line"]:
            key(f"scenarios.{sname}.annual.Y{y}.revenue_by_channel.{ch}", f"annual_Y{y}_revenue_by_channel_{ch}", sel)
    S["annual"] = annual
    nb = v("cash_funding_needed_by_month"); rw = v("cash_runway_months_day90"); rwu = v("cash_runway_months_day90_unfunded")
    S["cash"] = dict(peak_need=v("cash_peak_need"), peak_need_month=v("cash_peak_need_month"), peak_need_month_label=v("cash_peak_need_month_label"), min_balance=v("cash_min_balance"),
                     funding_gap=v("cash_funding_gap"), funding_needed_by_month=(nb if isinstance(nb, (int, float)) else None),
                     runway_months_day90=(rw if isinstance(rw, (int, float)) else rw), runway_months_day90_unfunded=(rwu if isinstance(rwu, (int, float)) else rwu),
                     opening_capital=v("cash_opening_capital"), burn_months_4_to_6=v("cash_burn_months_4_to_6"),
                     year_end_balance={f"Y{y}": v(f"cash_Y{y}_balance") for y in (3, 4, 5)}, year_end_balance_after_funding={f"Y{y}": v(f"cash_Y{y}_balance_after_funding") for y in (3, 4, 5)},
                     cash_balance_month24=bal[24], cash_balance_month24_after_funding=balf[24])
    for k, nm in [("peak_need", "cash_peak_need"), ("peak_need_month", "cash_peak_need_month"), ("peak_need_month_label", "cash_peak_need_month_label"), ("min_balance", "cash_min_balance"), ("funding_gap", "cash_funding_gap"), ("funding_needed_by_month", "cash_funding_needed_by_month"), ("runway_months_day90", "cash_runway_months_day90"), ("runway_months_day90_unfunded", "cash_runway_months_day90_unfunded"), ("opening_capital", "cash_opening_capital")]:
        key(f"scenarios.{sname}.cash.{k}", nm, sel)
    def gate(prefix, metrics):
        jname = {"d90": "day90", "m12": "month12"}[prefix]     # JSON key name; the cell names keep the short prefix
        rname = f"gate_{jname}_result"
        out = {"metrics": {}, "result": v(rname)}
        for mk, label, pass_t, kill_t in metrics:
            out["metrics"][mk] = dict(label=label, value=v(f"gate_{prefix}_{mk}_value"), pass_threshold=pass_t, kill_trigger=kill_t, passes=v(f"gate_{prefix}_{mk}_pass"), kills=v(f"gate_{prefix}_{mk}_kill"))
            key(f"scenarios.{sname}.gates.{jname}.metrics.{mk}.value", f"gate_{prefix}_{mk}_value", sel)
        key(f"scenarios.{sname}.gates.{jname}.result", rname, sel)
        return out
    S["gates"] = {
        "day90": gate("d90", [("tacos", "TACOS in Month 3", "at or below 30%", "above 40%"), ("ret", "Return rate in Month 3", "at or below 4%", "above 7%"), ("cpu", "Contribution per unit before advertising (hero basis)", "at or above $11", "below $8"), ("runway", "Cash runway after funding, months", "at or above 4", "below 2")]),
        "month12": gate("m12", [("rev", "Monthly revenue in Month 12", "at or above $70,000", "below $50,000"), ("cum", "Cumulative net at Month 12", "at or above $0", "below ($25,000)"), ("tacos", "TACOS in Month 12", "at or below 15%", "above 25%"), ("ret", "Return rate in Month 12", "at or below 3%", "above 5%"), ("opcf", "Operating cash flow, lower of Months 11 and 12", "positive in Months 11 and 12", "negative Months 10 to 12 with a funding gap above $50,000")]),
    }
    S["gates"]["day90"]["observed"] = ["Unit session conversion at or above 8% on at least 3 of the 5 launch SKUs (kill: hero below 5%)", "No suppressed listing and no IP complaint open more than 14 days"]
    S["gates"]["day90"]["result_rule"] = "KILL if any kill trigger fires; PASS if every modeled pass threshold is met; REVIEW otherwise"
    S["month12_revenue"] = monthly[11]["amazon_revenue"]; S["month12_cum_net"] = monthly[11]["amazon_cum_net"]
    # product liability insurance trigger: first month with Amazon US revenue at or above facts.amazon.insurance_trigger_monthly_sales_usd (single source for the deck and the SOPs)
    _thr = FACTS["amazon"]["insurance_trigger_monthly_sales_usd"]
    _hit = next((d for d in monthly if d["amazon_revenue"] >= _thr), None)
    S["insurance_trigger_month"] = _hit["month"] if _hit else None
    S["insurance_trigger_margin_usd"] = (_hit["amazon_revenue"] - _thr) if _hit else None
    S["sensitivity_base"] = dict(m18_net=v("sens_base_m18_net"), y5_net=v("sens_base_y5_net"), y5_revenue=v("sens_base_y5_revenue"))
    return sname, S, v

def pull_variant(sel, duty, fx, tag):
    """Full recalc of a temporary copy at (Scenario, DutyScenario, FXRate); returns the headline figures the documents quote."""
    tmp = os.path.join(QA, f"model_variant_{tag}.xlsx")
    shutil.copy(WB, tmp)
    wb = load_workbook(tmp)
    wb["Assumptions"]["C3"] = sel; wb["Assumptions"]["C4"] = duty; wb["Assumptions"]["C5"] = fx
    wb.save(tmp)
    res = recalc(tmp)
    assert res.get("status") == "success" and res.get("total_errors") == 0, res
    v = Reader(tmp)
    assert v("Scenario") == sel and v("DutyScenario") == duty and v("FXRate") == fx
    be = v("m18_breakeven_month"); nb = v("cash_funding_needed_by_month")
    return dict(scenario=mc.SCEN_NAMES[sel - 1], duty_scenario=mc.DUTY_NAMES[duty - 1], fx=fx, landed_pct=v("LandedPctActive"), duty_rate=v("DutyRateActive"),
                m18_revenue=v("m18_revenue"), m18_net=v("m18_net"), m18_net_margin=v("m18_net_margin"), m18_breakeven_month=(be if isinstance(be, (int, float)) else None),
                y5_revenue=v("annual_Y5_revenue"), y5_net=v("annual_Y5_net"), y5_net_margin=v("annual_Y5_net_margin"),
                cum_net_5y=v("annual_Y5_cum_net"), peak_need=v("cash_peak_need"), peak_need_month=v("cash_peak_need_month"), funding_gap=v("cash_funding_gap"),
                funding_needed_by_month=(nb if isinstance(nb, (int, float)) else None), hero_net_per_unit=v("duty_" + mc.DUTY_NAMES[duty - 1].lower() + "_hero_net"))

def main():
    F = {"meta": {"fx": mc.FX_DEFAULT, "scenario_default": "Medium", "duty_default": "Copper", "built": "October 8, 2026", "workbook": "output/03_Financial_Model_5Y.xlsx",
                  "fx_scenarios": mc.FX_SCEN, "month_rule": "Month 1 is the first month with inventory live on Amazon (assumed early 2027); no calendar months",
                  "m18_basis": "18 month figures are Amazon US only (August basis): revenue and net after fixed costs, seller fee and launch costs, Months 1 to 18",
                  "monthly_basis": "monthly revenue and net are all channels (Amazon US plus eBay and Walmart and B2B from Month 13); amazon_revenue and amazon_net are Amazon US only",
                  "tab_name_note": "The duty tab is named Landed Cost & Tariff Scenarios (30 characters) because Excel limits sheet names to 31 characters"}}
    F["scenarios"] = {}
    readers = {}
    for sel in (1, 2, 3):
        sname, S, v = pull_scenario(sel)
        F["scenarios"][sname] = S; readers[sname] = v
    v = readers["Medium"]
    # unit economics (scenario independent)
    hero = {k: v(f"ue_hero_{nm}") for k, nm in [("price", "price"), ("list", "list"), ("landed_cost", "landed_cost"), ("landed_pct", "landed_pct"), ("referral", "referral"), ("fba", "fba"), ("ads", "ads"), ("returns_storage", "returns_storage"), ("net", "net"), ("net_pct", "net_pct"), ("korea_wholesale_gross", "korea_wholesale_gross"), ("contribution_before_ads", "contribution_before_ads")]}
    for k in hero: key(f"unit_economics.hero.{k}", f"ue_hero_{k}")
    hero["fob_usd"] = v("HeroFobUsd"); hero["fob_share_of_price"] = v("HeroFobShare"); key("unit_economics.hero.fob_usd", "HeroFobUsd"); key("unit_economics.hero.fob_share_of_price", "HeroFobShare")
    hero["contribution_before_ads_at_asp42"] = F["scenarios"]["Medium"]["monthly"][2]["contribution_before_ads"]
    hdr = ["sku", "spec_internal", "length_m", "h2_cost_krw", "fob_usd", "us_list", "us_sale_price", "packaged_lb", "kg", "air_freight", "ocean_freight", "duty_copper", "duty_fiber", "duty_stress", "duty_active", "mpf_hmf_ocean", "broker_bond", "fba_inbound_prep", "qc_labels",
           "landed_air", "landed_ocean", "landed_air_pct", "landed_ocean_pct", "fba_fee", "referral", "ads", "returns_storage", "net_air", "net_air_pct", "net_ocean", "net_ocean_pct", "launch_sku"]
    tbl = v("sku_table"); n = len(hdr)
    per_sku = [dict(zip(hdr, tbl[i:i + n])) for i in range(0, len(tbl), n)]
    key("unit_economics.per_sku", "sku_table")
    cost_bases = {k: {"landed": v(f"cost_base_{k}_landed"), "net": v(f"cost_base_{k}_net")} for k in ["h2", "august", "alibaba_low", "alibaba_high"]}
    for k in cost_bases:
        key(f"unit_economics.cost_bases.{k}.landed", f"cost_base_{k}_landed"); key(f"unit_economics.cost_bases.{k}.net", f"cost_base_{k}_net")
    F["unit_economics"] = {"hero": hero, "per_sku": per_sku, "cost_bases": cost_bases,
                           "note": "Hero block is the August model (evidence based). Per SKU rows are the bottom up H2 cost base at the active duty scenario (Copper) and FX 1430; air freight for launch lots."}
    F["duty"] = {d: {"rate": v(f"duty_{d}_rate"), "landed_pct_medium": v(f"duty_{d}_landed_pct_medium"), "hero_net": v(f"duty_{d}_hero_net")} for d in ["copper", "fiber", "stress"]}
    F["duty"]["fiber_list3_check"] = {"rate": v("duty_fiber_list3_check_rate")}
    F["duty"]["active_landed_pct"] = {"Low": v("LandedPctLow"), "Medium": v("LandedPctMedium"), "High": v("LandedPctHigh")}
    for d in ["copper", "fiber", "stress"]:
        for k in ["rate", "landed_pct_medium", "hero_net"]: key(f"duty.{d}.{k}", f"duty_{d}_{k}")
    key("duty.fiber_list3_check.rate", "duty_fiber_list3_check_rate")
    wf = v("waterfall_block"); F["duty"]["waterfall"] = [{"label": wf[i], "base": wf[i + 1], "value": wf[i + 2]} for i in range(0, len(wf), 3)]; key("duty.waterfall", "waterfall_block")
    ts = v("tariff_stack_block"); F["duty"]["tariff_stack"] = [{"scenario": ts[i], "mfn": ts[i + 1], "section_301_list": ts[i + 2], "section_301_forced_labor": ts[i + 3], "pending_second_action": ts[i + 4], "total": ts[i + 5]} for i in range(0, len(ts), 6)]; key("duty.tariff_stack", "tariff_stack_block")
    tb = v("tornado_block"); tor = [{"driver": tb[i], "low_label": tb[i + 1], "low_value": tb[i + 2], "high_label": tb[i + 3], "high_value": tb[i + 4], "base": tb[i + 5], "swing": tb[i + 6], "low_minus_base": tb[i + 7], "high_minus_base": tb[i + 8]} for i in range(0, len(tb), 9)]
    tb18 = v("tornado18_block"); tor18 = [{"driver": tb18[i], "low_label": tb18[i + 1], "low_value": tb18[i + 2], "high_label": tb18[i + 3], "high_value": tb18[i + 4], "base": tb18[i + 5], "swing": tb18[i + 6], "low_minus_base": tb18[i + 7], "high_minus_base": tb18[i + 8]} for i in range(0, len(tb18), 9)]
    st = v("sensitivity_table"); table = [{"driver": st[i], "setting": st[i + 1], "setting_value": st[i + 2], "landed_pct_change": st[i + 3], "m18_net": st[i + 4], "y5_net": st[i + 5], "m18_net_vs_base": st[i + 6], "y5_net_vs_base": st[i + 7]} for i in range(0, len(st), 8)]
    F["sensitivity"] = {"base_metric": "Year 5 net, Medium", "base_m18_net": v("sens_base_m18_net"), "base_y5_net": v("sens_base_y5_net"), "base_y5_revenue": v("sens_base_y5_revenue"), "tornado": tor, "tornado_m18": tor18, "table": table,
                        "note": "Live formulas (compact one line P&L applied to the Monthly and Annual tabs). Factory cost Alibaba case = midpoint of the $17 to $22 benchmark with duty on the higher FOB."}
    key("sensitivity.tornado", "tornado_block"); key("sensitivity.tornado_m18", "tornado18_block"); key("sensitivity.table", "sensitivity_table"); key("sensitivity.base_y5_net", "sens_base_y5_net"); key("sensitivity.base_m18_net", "sens_base_m18_net")
    lb = v("levers_block"); F["levers"] = [{"n": lb[i], "lever": lb[i + 1], "assumption_moved": lb[i + 2], "from": lb[i + 3], "to": lb[i + 4], "owner": lb[i + 5], "evidence": lb[i + 6], "model_cell": lb[i + 7], "model_effect": lb[i + 8], "effect_basis": lb[i + 9]} for i in range(0, len(lb), 10)]; key("levers", "levers_block")
    F["gate_thresholds"] = {k[2:]: {"label": lbl, "unit": unit, "value": v("gate_threshold_" + k[2:])} for k, lbl, unit, val in mc.GATES}
    for k, lbl, unit, val in mc.GATES: key(f"gate_thresholds.{k[2:]}.value", "gate_threshold_" + k[2:])
    F["august_deal"] = {k: v("august_deal_" + k) for k in ["units", "cost_krw", "revenue_krw", "gross_profit_krw", "gross_margin"]}
    for k in F["august_deal"]: key(f"august_deal.{k}", "august_deal_" + k)
    F["model_notes"] = model_notes()
    F["kill_criteria"] = {
        "day90": [{"metric": "Unit session conversion on at least 3 of the 5 launch SKUs", "pass": "at or above 8%", "kill": "hero below 5%", "source": "observed"},
                  {"metric": "TACOS in Month 3, falling month over month", "pass": "at or below 30%", "kill": "above 40%", "source": "model"},
                  {"metric": "Return rate", "pass": "at or below 4%", "kill": "above 7%", "source": "model"},
                  {"metric": "Contribution per unit before advertising on at least 3 of 5 SKUs", "pass": "at or above $11 (Medium hero $12.14)", "kill": "hero below $8", "source": "model"},
                  {"metric": "Cash runway after funding the next purchase order", "pass": "at or above 4 months", "kill": "below 2 months", "source": "model"},
                  {"metric": "Suppressed listing or IP complaint", "pass": "none open", "kill": "any open more than 14 days", "source": "observed"}],
        "month12": [{"metric": "Monthly revenue in Month 12", "pass": "at or above $70,000 (Medium $86,100; Low $46,740)", "kill": "below $50,000", "source": "model"},
                    {"metric": "Cumulative net", "pass": "at or above $0", "kill": "below ($25,000)", "source": "model"},
                    {"metric": "TACOS", "pass": "at or below 15%", "kill": "above 25%", "source": "model"},
                    {"metric": "Return rate", "pass": "at or below 3%", "kill": "above 5%", "source": "model"},
                    {"metric": "Operating cash flow", "pass": "positive two consecutive months", "kill": "three consecutive negative months with a funding gap above $50,000", "source": "model"}],
        "kill_means": "No further capital, sell down inventory, delist losers, move to Plan B",
        "results": {s: {"day90": F["scenarios"][s]["gates"]["day90"]["result"], "month12": F["scenarios"][s]["gates"]["month12"]["result"]} for s in mc.SCEN_NAMES}}
    F["plan_b"] = {f"Y{y}": {"revenue": v(f"planb_Y{y}_revenue"), "net": v(f"planb_Y{y}_net"), "b2b_revenue": v(f"planb_Y{y}_b2b_revenue"), "korea_gross_profit": v(f"planb_Y{y}_korea_gross_profit")} for y in (1, 2, 3)}
    for y in (1, 2, 3):
        for k in ["revenue", "net", "b2b_revenue", "korea_gross_profit"]: key(f"plan_b.Y{y}.{k}", f"planb_Y{y}_{k}")
    # scenario summary cross check (read from the delivered workbook at its default state)
    vd = Reader(WB)
    assert vd("Scenario") == 2 and vd("DutyScenario") == 1 and vd("FXRate") == mc.FX_DEFAULT, "delivered workbook must be at Scenario 2, Duty 1, FX 1430"
    F["scenario_summary"] = {s: {k: vd(f"ss_{s}_{k}") for k in ["m18_revenue", "m18_net", "m18_net_margin", "m18_breakeven_month", "m18_run_rate", "m18_exit_margin", "annual_Y1_revenue", "annual_Y1_net", "annual_Y2_revenue", "annual_Y2_net", "annual_Y3_revenue", "annual_Y3_net", "annual_Y4_revenue", "annual_Y4_net", "annual_Y5_revenue", "annual_Y5_net", "annual_Y5_net_margin", "cash_peak_need", "cash_peak_need_month", "cash_peak_need_month_label", "cash_funding_gap", "cash_funding_needed_by_month", "cash_runway_months_day90", "gate_day90", "gate_month12"]} for s in mc.SCEN_NAMES}
    for s in mc.SCEN_NAMES:
        for k in F["scenario_summary"][s]: key(f"scenario_summary.{s}.{k}", f"ss_{s}_{k}", "any (Scenario Summary engines)")
    # cross check: selector pulls equal the Scenario Summary engines
    mism = []
    for s in mc.SCEN_NAMES:
        for k, path in [("m18_revenue", ("m18", "revenue")), ("m18_net", ("m18", "net")), ("annual_Y5_revenue", ("annual", "Y5", "revenue")), ("annual_Y5_net", ("annual", "Y5", "net")), ("cash_peak_need", ("cash", "peak_need")), ("cash_funding_gap", ("cash", "funding_gap"))]:
            a = F["scenario_summary"][s][k]; b = F["scenarios"][s]
            for p in path: b = b[p]
            if abs(a - b) > max(1.0, 0.0005 * abs(b)): mism.append((s, k, a, b))
    F["meta"]["scenario_summary_cross_check"] = "all match within 0.05%" if not mism else f"MISMATCH {mism}"
    F["meta"]["selector_state_delivered"] = {"Scenario": vd("Scenario"), "DutyScenario": vd("DutyScenario"), "FXRate": vd("FXRate")}
    # full recalc variants: duty and FX effects on the Medium case (exact; the Sensitivity tab rows are compact approximations)
    base = pull_variant(2, 1, mc.FX_DEFAULT, "medium_base")
    F["full_recalc"] = {"note": "Each block is a full LibreOffice recalc of the workbook at the stated selectors (Scenario 2 Medium unless stated). Quote duty and FX effects from here, not from the compact Sensitivity rows.",
                        "medium_base": base,
                        "duty": {"copper": base, "fiber": pull_variant(2, 2, mc.FX_DEFAULT, "medium_fiber"), "stress": pull_variant(2, 3, mc.FX_DEFAULT, "medium_stress")},
                        "fx": {str(fx): (base if fx == mc.FX_DEFAULT else pull_variant(2, 1, fx, f"medium_fx{fx}")) for fx in mc.FX_SCEN}}
    for blk in ("duty", "fx"):
        for k, d in F["full_recalc"][blk].items():
            d["m18_net_vs_base"] = d["m18_net"] - base["m18_net"]; d["y5_net_vs_base"] = d["y5_net"] - base["y5_net"]; d["landed_pct_vs_base"] = d["landed_pct"] - base["landed_pct"]
    # headline convenience block
    F["headline"] = {s: {"m18_revenue": F["scenarios"][s]["m18"]["revenue"], "m18_net": F["scenarios"][s]["m18"]["net"], "m18_net_margin": F["scenarios"][s]["m18"]["net_margin"], "breakeven_month": F["scenarios"][s]["m18"]["breakeven_month"],
                         "run_rate": F["scenarios"][s]["m18"]["run_rate"], "exit_margin": F["scenarios"][s]["m18"]["exit_margin"],
                         "annual_revenue": [F["scenarios"][s]["annual"][f"Y{y}"]["revenue"] for y in range(1, 6)], "annual_net": [F["scenarios"][s]["annual"][f"Y{y}"]["net"] for y in range(1, 6)],
                         "y5_net_margin": F["scenarios"][s]["annual"]["Y5"]["net_margin"], "peak_need": F["scenarios"][s]["cash"]["peak_need"], "peak_need_month": F["scenarios"][s]["cash"]["peak_need_month"],
                         "funding_gap": F["scenarios"][s]["cash"]["funding_gap"], "funding_needed_by_month": F["scenarios"][s]["cash"]["funding_needed_by_month"],
                         "gate_day90": F["scenarios"][s]["gates"]["day90"]["result"], "gate_month12": F["scenarios"][s]["gates"]["month12"]["result"]} for s in mc.SCEN_NAMES}
    F["insurance_trigger_month"] = {s: F["scenarios"][s]["insurance_trigger_month"] for s in mc.SCEN_NAMES}
    with open(os.path.join(HERE, "figures.json"), "w") as f: json.dump(F, f, indent=1)
    with open(os.path.join(HERE, "figures_keys.json"), "w") as f: json.dump(KEYS, f, indent=1)
    write_readme(F)
    print("figures.json written;", F["meta"]["scenario_summary_cross_check"])
    for s in mc.SCEN_NAMES:
        h = F["headline"][s]
        print(f"{s}: 18m rev {h['m18_revenue']:,.0f} net {h['m18_net']:,.0f} ({h['m18_net_margin']:.1%}) BE {h['breakeven_month']} | Y1..Y5 rev {[round(x/1000) for x in h['annual_revenue']]}K net {[round(x/1000) for x in h['annual_net']]}K | peak need {h['peak_need']:,.0f} at {h['peak_need_month']} gap {h['funding_gap']:,.0f} by M{h['funding_needed_by_month']} | gates {h['gate_day90']}/{h['gate_month12']}")

def model_notes():
    """Deviations from BUILD_SPEC and the choices behind them, for 11_Decision_Log.md, the README tie out and the band slide. Numbers here are driver inputs (model_core.py), not results."""
    D = mc.D
    return [
        {"id": "ads_months_19_to_24", "where": "Assumptions: Ad spend % Months 13 to 18 and Months 19 to 24", "default": f"Months 13 to 18 at Low {D['ads_p4']['vals'][0]:.0%}, Medium {D['ads_p4']['vals'][1]:.0%}, High {D['ads_p4']['vals'][2]:.0%}; Months 19 to 24 at Low {D['ads_p5']['vals'][0]:.0%}, Medium {D['ads_p5']['vals'][1]:.0%}, High {D['ads_p5']['vals'][2]:.1%}",
         "why": "The brief's floor (ad spend never below 8% of revenue) is applied at the company level across all channels. With High at 8% on Amazon only, the eBay and Walmart (5%) and B2B (3%) mix diluted Year 2 to 7.4%. High steps to its Year 3 rate from Month 19, which keeps the 18 month block exactly as the August model and lifts High Year 2 to 8.1%.", "decides": "Mark", "by": "before the deck is sent"},
        {"id": "high_18_month_net", "where": "Scenario Summary: High 18 month net", "default": "The workbook High 18 month net is $447K (24.2%). The spec sentence said $457,804 (24.7%), which needs 7% ads in Months 13 to 18, below the 8% floor; the brief says $455K (24.6%). The workbook is 1.7% from the brief, inside the 2% tolerance.", "why": "The 8% floor wins over the spec figure. Decks quote $447K and 24.2% from figures.json, never $457,804 or $455K.", "decides": "Mark", "by": "done"},
        {"id": "low_lost_quarter_months", "where": f"Assumptions: lost quarter start month (Low = {D['lost_q_start']['vals'][0]})", "default": "Low lost quarter in Months 19 to 21 at 10% of trend (spec 2.5 said Months 15 to 17)", "why": "Spec 2.3 also requires the Low 18 month block to reproduce the brief's $633K revenue and ($19K) net, which a Months 15 to 17 suppression (about $169K of revenue) makes impossible. The brief only requires one lost quarter in Low. Do not move it back without re calibrating the Low 18 month block.", "decides": "Mark", "by": "done"},
        {"id": "low_year2_to_year3_margin", "where": "Scenario Summary: Low Year 2 and Year 3 net margin", "default": "Low net margin rises from Year 2 to Year 3", "why": "Year 2 carries the lost quarter (Months 19 to 21), so the Year 2 to Year 3 step is a recovery, not better economics. From Year 3 the Low margin compresses every year and Year 5 net is at or below zero.", "decides": "none", "by": "none"},
        {"id": "inventory_days_years_3_to_5", "where": "Assumptions: inventory days Years 3 to 5", "default": f"Low {D['inv_days']['vals'][0]}, Medium {D['inv_days']['vals'][1]}, High {D['inv_days']['vals'][2]} days of landed COGS (spec 2.6 started at 90, 75, 60)", "why": "Raised to the days implied by the Month 24 stock (cover months plus one lead time in transit) so Year 3 does not release cash artificially. Only the Year 3 to 5 balances change; peak need and funding gap sit in Months 7 to 24 in every scenario.", "decides": "Mark", "by": "before the bank meeting"},
        {"id": "calibration_beyond_spec", "where": "Assumptions rows for ASP change, ad spend Years 3 to 5, landed change Years 3 to 5, Low unit growth", "default": f"Medium ASP change {D['asp_chg']['vals'][1]:.0%} per year (spec minus 2%); Medium ad spend Years 3 to 5 {D['ads_y3']['vals'][1]:.1%} flat (spec 9.5, 9.0, 8.5); Medium landed change Y3 {D['landed_chg_y3']['vals'][1]*100:.1f}, Y4 {D['landed_chg_y4']['vals'][1]*100:.1f}, Y5 {D['landed_chg_y5']['vals'][1]*100:.1f} points (spec minus 2.5, minus 2.0, minus 1.5); Low ad spend Years 3 to 5 {D['ads_y3']['vals'][0]:.0%}, {D['ads_y4']['vals'][0]:.1%}, {D['ads_y5']['vals'][0]:.1%} (spec 12 flat); High ad spend {D['ads_y3']['vals'][2]:.1%} flat (spec 9.0, 8.5, 8.0); High ASP change {D['asp_chg']['vals'][2]:.0%} (spec minus 1.5%); Low unit growth {D['g_y3']['vals'][0]:.0%}, {D['g_y4']['vals'][0]:.0%}, {D['g_y5']['vals'][0]:.0%} (spec 5, 3, 0; within the plus or minus 5 bound)",
         "why": "The Year 5 bands are reached with heavier price and ad compression than the spec started from, so that margin compresses at least 1 point a year (brief realism rule) and the Low case stalls near $1M with a non positive Year 5 net. Every row carries a Note on Assumptions.", "decides": "Mark", "by": "before the deck is sent"},
        {"id": "medium_year4_landed_change", "where": "Assumptions: landed COGS % change Year 4 (Medium)", "default": f"Medium {D['landed_chg_y4']['vals'][1]*100:.1f} points (spec minus 2.0; fix round 1 used minus 1.5)", "why": "Creep back starts in Year 4 so the Medium net margin step from Year 3 to Year 4 exceeds 1 point (it was 0.96).", "decides": "Mark", "by": "done"},
        {"id": "b2b_returns_floor", "where": "Assumptions: B2B returns %", "default": f"{mc.CH['b2b_returns']:.1%} in every scenario (spec 2.5 channel economics said 1.5%)", "why": "The brief's realism rule (returns never below 2%) wins. The pro line B2B share inherits the same rate. Costs about $4K of Medium Year 5 net.", "decides": "none", "by": "done"},
        {"id": "low_year5_ads", "where": "Assumptions: ad spend % Year 5 (Low)", "default": f"{D['ads_y5']['vals'][0]:.1%}", "why": "Spec 2.5 asks for a Low Year 5 net that is negative or near zero. At 14% the Low Year 5 net was about $14K positive; at this rate it is at or below zero while cash before funding stays negative at every year end and the Month 12 gate reads KILL.", "decides": "Mark", "by": "before the deck is sent"},
        {"id": "compliance_from_month_19", "where": "Monthly PL: compliance row", "default": "Form 5472 preparation, the Wyoming report, the bond renewal and the product liability premium are charged from the Year 2 start month (Month 19); Year 1 sits inside the August $850 per month fixed cost and the $9,500 launch costs", "why": "Keeps the 18 month reconciliation exact. It understates Months 4 to 18 by at most about $1,500. The compliance calendar must still trigger the $1M product liability policy in the first month sales pass $10K (Medium Month 4).", "decides": "Lucy", "by": "Month 4"},
        {"id": "levers_and_plan_b_pinned_to_medium", "where": "Levers tab and Plan B tab", "default": "Both read the Medium column of Assumptions, not the Active column", "why": "Spec 2.8 quotes Medium From and To values and spec 2.7 defines Plan B as a fixed side case; pinning them means the lever table and Plan B read the same whatever the selector shows. No P&L, cash or gate figure depends on them. The landed % on Plan B still follows the duty and FX selectors.", "decides": "none", "by": "done"},
        {"id": "gate_review_outcome", "where": "Funding and Breakeven: gate results", "default": "Gate results are PASS, REVIEW or KILL", "why": "Spec 2.7 defines PASS and KILL only. REVIEW (no kill trigger fired, but not every pass threshold met) means fix and retest within 30 days. Low reads REVIEW at Day 90 and KILL at Month 12.", "decides": "none", "by": "done"},
        {"id": "low_peak_need_ordering", "where": "Cash Flow and Working Capital: Low peak need", "default": "The order rule keeps buying through the Low lost quarter to rebuild cover for the recovery month", "why": "The Low peak need at Month 24 therefore assumes purchases continue during the suppression. An operator would pause orders until the listing is restored; the modeled figure is conservative.", "decides": "none", "by": "done"},
        {"id": "runway_depends_on_funding_size", "where": "Funding and Breakeven: Day 90 runway", "default": "Runway after funding = cash after the funding draw divided by the Months 4 to 6 burn; the draw is 1.2 times the peak need", "why": "A more profitable case with a smaller draw can show a shorter runway. If the Day 90 runway test is marginal, size the line at the larger of the funding gap and four months of burn.", "decides": "Mark", "by": "at the Day 90 gate"},
        {"id": "duty_tab_name", "where": "Workbook tab 5", "default": "Landed Cost & Tariff Scenarios (30 characters)", "why": "Spec 3 asks for Landed Cost and Tariff Scenarios (32 characters) but Excel caps sheet names at 31. The tab's A1 title carries the full wording. No number is affected.", "decides": "none", "by": "done"},
    ]

def write_readme(F):
    L = ["# figures.json: every key, its meaning and the workbook cell it comes from", "",
         "Pulled by build/pull_figures.py from output/03_Financial_Model_5Y.xlsx after LibreOffice recalculation. Scenario blocks are pulled by setting the Scenario selector (Assumptions!C3) to 1, 2 and 3 on temporary copies; the delivered workbook stays at Scenario 2 (Medium), Duty 1 (Copper), FX 1430. Every defined name is listed in build/figures_map.json (name to sheet!cell) and every JSON key path in build/figures_keys.json (path to name, cell and selector state). Numbers are raw; format with build/fmt.py or build/fmt.js.", "",
         "## Conventions", "",
         "Model months: Month 1 is the first month with inventory live on Amazon (assumed early 2027). Year 1 is Months 1 to 12, Year 2 is Months 13 to 24.",
         "18 month block (m18): Amazon US only, the August basis, Months 1 to 18: revenue, net after fixed costs, seller fee and launch costs, net margin, cumulative breakeven month (null if never), Month 18 run rate (Month 18 revenue times 12), exit margin (Month 18 net margin), units.",
         "Monthly rows: revenue and net are all channels (Amazon US plus the eBay and Walmart overflow and B2B rows from Month 13); amazon_revenue and amazon_net are Amazon US only; cash_balance is before new funding, cash_balance_after_funding adds the line of credit or investor amount in the month it is needed.",
         "Annual: revenue, net, net margin, COGS, channel fees (referral, payment, commission), fulfillment (FBA, MCF, 3PL, shipping), advertising, returns and storage, contribution, fixed, seller fee, launch, compliance, revenue_by_channel and contribution_by_channel for amazon_us, ebay_walmart, b2b, dtc, amazon_canada, pro_line.",
         "Cash: peak_need = the negative minimum cash balance before funding (Months 0 to 24 and the Year 3 to 5 year ends); peak_need_month (number; 36, 48, 60 mean a year end) and peak_need_month_label; funding_gap = 1.2 times peak need; funding_needed_by_month = the month before the balance first goes negative (null if never); runway_months_day90 = cash at end of Month 3 after funding divided by the average burn of Months 4 to 6 (the string funded if there is no burn); runway_months_day90_unfunded = the same before funding.",
         "Gates: day90 and month12 carry each modeled metric with its value, pass threshold, kill trigger and the PASS or KILL booleans, plus result = PASS, REVIEW or KILL (KILL if any kill trigger fires, PASS if every modeled threshold is met, REVIEW otherwise). Conversion and listing suppression are observed at the gate, not modeled.", "",
         "## Key list", "", "| JSON key path | Defined name | Workbook cell | Selector state |", "|---|---|---|---|"]
    for path, d in KEYS.items():
        L.append(f"| {path} | {d['name']} | {d['ref']} | {d['scenario']} |")
    L += ["", "## Keys without a single cell", "",
          "meta: build metadata (fx, scenario_default, duty_default, built, workbook, fx_scenarios, selector state, scenario summary cross check).",
          "scenarios.{S}.gates.*.pass_threshold, kill_trigger, label: text from BUILD_SPEC 2.7 written by pull_figures.py.",
          "kill_criteria: the Day 90 and Month 12 criteria lists (text from BUILD_SPEC 2.7) plus results per scenario (from gate_day90_result and gate_month12_result).",
          "headline: a convenience copy of the scenario headline numbers (same cells as scenarios.{S}.m18, annual, cash, gates).",
          "unit_economics.per_sku: one object per row of the SKU Unit Economics table (sku_table), columns in table order.",
          "duty.waterfall and duty.tariff_stack: chart data blocks from the Landed Cost & Tariff Scenarios tab.",
          "sensitivity.table: the one way sensitivity table rows (driver, setting, setting value, landed % change, 18 month net, Year 5 net, deltas).",
          "scenario_summary: the Scenario Summary tab (three parallel live engines) read from the delivered workbook; used to cross check the selector pulls.",
          "full_recalc: Medium case recalculated at each duty scenario (copper, fiber, stress) and each FX scenario (1300, 1430, 1550): landed %, 18 month revenue and net, breakeven, Year 5 revenue and net, peak need, funding gap, and the deltas vs the Medium base. These are exact full model values; the Sensitivity tab rows are compact approximations.",
          "scenarios.{S}.insurance_trigger_month and insurance_trigger_margin_usd, insurance_trigger_month: the first month whose Amazon US revenue (monthly_amazon_revenue) is at or above facts.amazon.insurance_trigger_monthly_sales_usd, and the amount it passes the threshold by; derived by pull_figures.py, the single source for the deck and the SOPs.",
          "gate_thresholds: the Day 90 and Month 12 thresholds from the Gate thresholds block on Assumptions (one cell each; the gate formulas and their wording read these cells).",
          "august_deal: the August 2026 deal totals (units, cost, revenue, gross profit in KRW, gross margin) from the yellow input cells on the FX tab, sourced from inputs/august_deal_totals.csv.",
          "model_notes: every deviation from BUILD_SPEC and the choice behind it (id, where, default, why, decides, by), written by pull_figures.py from model_core.py for 11_Decision_Log.md, the README tie out and the band slide."]
    with open(os.path.join(HERE, "figures_readme.md"), "w") as f: f.write("\n".join(L) + "\n")

if __name__ == "__main__":
    main()

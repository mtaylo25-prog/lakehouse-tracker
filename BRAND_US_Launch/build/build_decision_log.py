#!/usr/bin/env python3
"""Builds output/11_Decision_Log.md from build/figures.json, build/facts.json and inputs/DECISIONS.md.
Every model number comes from figures.json through fmt.py. No hyphens or dashes in the delivered text.
Rerun: python3 build/build_decision_log.py
"""
import json, sys
from pathlib import Path

ROOT = Path("/home/user/lakehouse-tracker/BRAND_US_Launch")
sys.path.insert(0, str(ROOT / "build"))
import fmt  # noqa: E402

F = json.load(open(ROOT / "build" / "figures.json"))
X = json.load(open(ROOT / "build" / "facts.json"))

S = F["scenarios"]
M, L, H = S["Medium"], S["Low"], S["High"]
sens = {(r["driver"], r["setting"]): r for r in F["sensitivity"]["table"]}
duty = F["duty"]
hero = F["unit_economics"]["hero"]
cb = F["unit_economics"]["cost_bases"]
sku = {r["sku"]: r for r in F["unit_economics"]["per_sku"]}
gt = F["gate_thresholds"]
notes = {n["id"]: n for n in F["model_notes"]}
levers = {l["n"]: l for l in F["levers"]}
fx = F["meta"]["fx"]
built = F["meta"]["built"]
opening = M["cash"]["opening_capital"]
ent = X["entity"]
ins_month = F["insurance_trigger_month"]


def mk(x):
    return fmt.money_k(x)


def mm(x):
    return fmt.money_m(x)


def m2(x):
    return fmt.money(x, 2)


def pct(x):
    """Whole percentages print without a decimal (30%), others with one (13.1%)."""
    v = round(x * 1000) / 10.0
    return fmt.pct(x, 0) if abs(v - round(v)) < 1e-9 else fmt.pct(x, 1)


def krwm(usd, rate):
    return fmt.krw_m(usd * rate)


def delta(x):
    """Signed delta in $K with the word plus or minus, never a bare sign."""
    if x >= 0:
        return "plus " + mk(x)
    return "minus " + mk(abs(x))


def deltam(x):
    if x >= 0:
        return "plus " + mm(x)
    return "minus " + mm(abs(x))


def cell(s):
    return s.replace("|", "/").replace("\n", " ")


def table(headers, rows):
    out = ["| " + " | ".join(headers) + " |", "|" + "---|" * len(headers)]
    for r in rows:
        out.append("| " + " | ".join(cell(str(c)) for c in r) + " |")
    return "\n".join(out)


# Sensitivity shortcuts (Medium base)
def sv(driver, setting):
    return sens[(driver, setting)]


fx1300, fx1550 = sv("FX (KRW per USD)", "1,300"), sv("FX (KRW per USD)", "1,550")
dfib, dstr = sv("Duty rate", "Fiber 12.5%"), sv("Duty rate", "Stress 50.1%")
ali, f25, f10 = sv("Factory cost (FOB)", "Alibaba benchmark"), sv("Factory cost (FOB)", "plus 25%"), sv("Factory cost (FOB)", "minus 10%")
aspm, aspp = sv("ASP", "minus 10%"), sv("ASP", "plus 10%")
adm, adp = sv("Ad spend", "minus 3 points"), sv("Ad spend", "plus 3 points")
ret2, ret5 = sv("Return rate", "2%"), sv("Return rate", "5%")

m18_net = M["m18"]["net"]
y5_net = M["annual"]["Y5"]["net"]
gap = M["cash"]["funding_gap"]
peak = M["cash"]["peak_need"]
peak_m = M["cash"]["peak_need_month"]
need_by = M["cash"]["funding_needed_by_month"]
first_neg = next(r for r in M["monthly"] if r["cash_balance"] < 0)
total_capital = opening + gap

# Withholding illustration if Mark were an owner (37% on Lucy's share; shown at 100% share as the ceiling)
wh_y1 = 0.37 * M["annual"]["Y1"]["net"]
wh_y5 = 0.37 * y5_net
retainer_y = 2000 * 12
wy_5y = ent["formation_cost_usd"] + 4 * ent["annual_cost_usd"]
de_5y = 5 * ent["delaware_annual_cost_usd"]

hero_h2 = sku["8K_10M"]
w4k = sku["4K_30M"]
s25 = sku["8K_25M"]

doc = []
A = doc.append

A("# 11 Decision Log: open decisions, the defaults used, and what changes if they go the other way")
A("")
A(f"Built {built} from inputs/DECISIONS.md, build/BUILD_SPEC.md, build/figures.json (pulled from output/03_Financial_Model_5Y.xlsx) and build/facts.json. Every model number below is the Medium scenario at the delivered selector state (Scenario Medium, Duty Copper, FX {fmt.num(fx)}) unless the row says otherwise. Model months: Month 1 is the first month with inventory live on Amazon (assumed early 2027). Status is Open on every row until Lucy or Mark writes the decision into inputs/DECISIONS.md and the build is rerun.")
A("")
A("How to read the columns: Default used is what every file in output/ assumes today. What changes is quantified from the Sensitivity and Scenario Summary tabs of the workbook where the model can measure it, and labeled as a direction only where it cannot. By when is relative to the capital transfer (the Korean overseas direct investment report must be filed before capital moves) or to the launch week numbers in 07_90_Day_Launch_Plan.xlsx.")
A("")
A("Acronyms used here: ODI (overseas direct investment report, 해외직접투자 신고), LLC (limited liability company), EIN (Employer Identification Number), CBP (US Customs and Border Protection), HTS (Harmonized Tariff Schedule), FOB (free on board, the price at the factory port), COGS (cost of goods sold), ASP (average selling price), FBA (Fulfillment by Amazon), TACOS (total advertising cost of sales), PE (permanent establishment), LOC (line of credit), SKU (stock keeping unit), FX (foreign exchange rate, KRW per USD), QC (quality control), 3PL (third party logistics warehouse), ATC (HDMI Authorized Test Center), FCC (Federal Communications Commission), CL3 (in wall cable fire rating), 세무사 (Korean licensed tax accountant).")
A("")

# ---------------------------------------------------------------- Section 1
A("## 1. The five open decisions")
A("")
hdr = ["#", "Decision", "Default used", "Why", "What changes if it goes the other way (quantified)", "Who decides", "By when", "Status"]
rows = []

rows.append([
    1, "US brand name",
    "BRAND, the literal placeholder in every file. BRAND LLC is the entity name.",
    "No name is chosen. The Korean deck used MBF, which is not cleared.",
    f"No model number moves. Everything else on the critical path waits for it: the trademark filing (about {fmt.money(X['trademark']['uspto_fee_per_class_usd'])} per class plus attorney fees, placeholder) in weeks 1 to 2, the LLC filing (exact legal name match across state, IRS, bank and Amazon), the Amazon Brand Registry application (gates A+ content, Brand Store, video ads and hijacker protection, which levers 5 and 6 in the Levers tab depend on), the packaging artwork and the listing titles in 09_Amazon_Listing_Kit.docx. If the name is MBF, a USPTO clearance search is still required [PLACEHOLDER: attorney clearance result]. Every delivered file needs a search and replace of BRAND once the name is final.",
    "Lucy, with Mark's input and the trademark attorney's clearance",
    "Week 1, before the LLC articles are filed",
    "Open",
])

rows.append([
    2, "Ownership of the US LLC",
    "Lucy (Hyejin Yoo) owns 100%. Mark Taylor is paid under a written consulting or commission agreement, no equity.",
    "Keeps the LLC a foreign owned single member LLC (Form 5472 with a pro forma Form 1120) and keeps the treaty position clean: no US agent who negotiates or signs contracts, so no PE.",
    f"If Mark takes any stake the LLC becomes a partnership: Form 1065 replaces Form 5472 and 37% withholding applies to Lucy's share of profit. At the Medium Year 1 net of {mk(M['annual']['Y1']['net'])} that is up to {mk(wh_y1)} withheld; at the Year 5 net of {mk(y5_net)} up to {mk(wh_y5)} (ceiling shown at a 100% share; multiply by her actual share). Mark acting as an owner who signs contracts also risks a US PE, which would move the whole profit into US tax. The model's Mark retainer of $2,000 per month from Year 3 (Medium, {mk(retainer_y)} per year) would become a profit share instead. The second alternative, Mark's own US company as distributor and importer of record with Lucy at the wholesale layer, keeps Lucy at about the Korea wholesale margin ({m2(hero['korea_wholesale_gross'])} gross per hero unit) instead of the Amazon net of {m2(hero['net'])} per unit, and moves the customs bond, insurance, Amazon account and Form 5472 off her list.",
    "Lucy and Mark, then the 세무사 and a US CPA on the tax side",
    "Week 1, before the LLC articles and operating agreement, and before the ODI report (the report states the stake)",
    "Open",
])

rows.append([
    3, "Funding source",
    f"Self funded by Lucy to {mk(opening)} (the August ask). Any need above that is a dated line of credit or investor amount in the model.",
    f"The August deck asked for about {mk(X['august']['funding_ask_usd'])}. The working capital model shows that is not enough on its own.",
    f"Medium: the cash balance first goes negative in {fmt.month_label(first_neg['month'])}, balance {mk(first_neg['cash_balance'])}, peak need is {mk(peak)} in {fmt.month_label(peak_m)}, so the funding gap is {mk(gap)} (1.2 times peak) and must be in place by {fmt.month_label(need_by)}. Low: peak {mk(L['cash']['peak_need'])} in {fmt.month_label(L['cash']['peak_need_month'])}, gap {mk(L['cash']['funding_gap'])} by {fmt.month_label(L['cash']['funding_needed_by_month'])}. High: peak {mk(H['cash']['peak_need'])} in {fmt.month_label(H['cash']['peak_need_month'])}, gap {mk(H['cash']['funding_gap'])} by {fmt.month_label(H['cash']['funding_needed_by_month'])}. Other way A, Lucy funds the whole Medium need up front: {mk(opening)} plus {mk(gap)} = {mk(total_capital)} of owner capital, {krwm(total_capital, fx)} at {fmt.num(fx)}, all declared on one ODI report, and no lender. Other way B, a US line of credit: no lender is identified [PLACEHOLDER: lender or investor and terms], a new LLC with no credit history may not get one by {fmt.month_label(need_by)}, so the realistic form is a second owner capital transfer (a second ODI remittance report) or a Korean loan to Lucy. Doing nothing: with {mk(opening)} only, runway at the Day 90 gate before funding is {fmt.num(M['cash']['runway_months_day90_unfunded'])} months, which fires the kill trigger (below {fmt.num(gt['d90_run_kill']['value'])} months); after funding it is {fmt.num(M['cash']['runway_months_day90'], 1)} months.",
    "Lucy (amount and source), Mark (lender search if any)",
    f"Amount before the capital transfer (the ODI report states it). Funds in the LLC account by {fmt.month_label(need_by)}, which is about week 14 if Month 1 starts in week 10",
    "Open",
])

rows.append([
    4, "FX rate",
    f"{fmt.num(fx)} KRW per USD, an input cell (Assumptions!C5, named FXRate). Scenarios at 1,300 and 1,550.",
    "Mark's October 2026 working rate. The only verified cost data (the August deal) is in KRW, so FX moves landed COGS through the hero FOB share of price.",
    f"At 1,300 (stronger won): landed COGS rises {fmt.pts(fx1300['landed_pct_change'])}, 18 month net {mk(fx1300['m18_net'])} ({delta(fx1300['m18_net_vs_base'])} versus {mk(m18_net)}), Year 5 net {mk(fx1300['y5_net'])} ({delta(fx1300['y5_net_vs_base'])}); Lucy's KRW cost of the {mk(opening)} opening capital falls to {krwm(opening, 1300)} and of the {mk(gap)} funding gap to {krwm(gap, 1300)}. At 1,550 (weaker won): landed COGS falls {fmt.pts(abs(fx1550['landed_pct_change']))}, 18 month net {mk(fx1550['m18_net'])} ({delta(fx1550['m18_net_vs_base'])}), Year 5 net {mk(fx1550['y5_net'])} ({delta(fx1550['y5_net_vs_base'])}); the opening capital costs {krwm(opening, 1550)} and the gap {krwm(gap, 1550)}. At {fmt.num(fx)} the capital is {krwm(opening, fx)} and the gap {krwm(gap, fx)}. The KRW figures hold the USD need constant; the model's own gap moves a few thousand dollars with FX. A USD FOB quote from the factory removes the FX effect on COGS entirely and leaves only the capital conversion.",
    "Lucy (the rate is set by the market on the transfer date; she decides the model input)",
    "Set the input to the bank's rate on the day of the capital transfer; review at each PO",
    "Open",
])

rows.append([
    5, "Duty scenario for the base case",
    f"Copper: HTS 8544.42.90, {pct(X['duty']['copper']['mfn'])} MFN plus {fmt.pct(X['duty']['copper']['section_301_list'], 0)} Section 301 List 3 plus {pct(X['duty']['copper']['section_301_forced_labor'])} Section 301 forced labor = {pct(duty['copper']['rate'])}. Landed COGS {fmt.pct(duty['copper']['landed_pct_medium'], 0)} of revenue, hero net {m2(duty['copper']['hero_net'])} per unit.",
    "The brief's conservative default pending the CBP binding ruling. The 12.5% applies under every classification.",
    f"Fiber (HTS 8544.70.00, {pct(duty['fiber']['rate'])}): landed COGS {pct(duty['fiber']['landed_pct_medium'])} ({fmt.pts(abs(dfib['landed_pct_change']))} better), hero net {m2(duty['fiber']['hero_net'])} per unit, 18 month net {mk(dfib['m18_net'])} ({delta(dfib['m18_net_vs_base'])}), Year 5 net {mk(dfib['y5_net'])} ({delta(dfib['y5_net_vs_base'])}). Stress (copper plus the pending second Section 301 action, {pct(duty['stress']['rate'])}): landed COGS {pct(duty['stress']['landed_pct_medium'])}, hero net {m2(duty['stress']['hero_net'])}, 18 month net {mk(dstr['m18_net'])} ({delta(dstr['m18_net_vs_base'])}), Year 5 net {mk(dstr['y5_net'])} ({delta(dstr['y5_net_vs_base'])}). Check row: if 8544.70.00 is also on List 3 the rate is {pct(duty['fiber_list3_check']['rate'])}, almost the copper case, so the fiber upside depends on the broker confirming List 3 coverage [PLACEHOLDER: broker confirmation]. The brief states the duty ruling is the single variable separating Low from Medium.",
    "CBP decides the classification; the broker files; Mark owns the request",
    "Ruling request filed in week 1; answer in 30 to 90 days, so before the Day 90 gate (end of Month 3). Model on Copper until the ruling letter is in hand",
    "Open",
])
A(table(hdr, rows))
A("")

# ---------------------------------------------------------------- Section 2
A("## 2. Additional defaults a decision could change")
A("")
A("These are the other defaults in inputs/DECISIONS.md and build/BUILD_SPEC.md that one decision by Lucy, Mark, the factory, the broker or a regulator would move. Same columns.")
A("")
rows = []

rows.append([
    "D1", "H2 Company legal form",
    "Korean sole proprietorship (registration number 599-64-00496, middle digits 64). The LLC is owned by Lucy personally.",
    "Brief. Flagged for the 세무사.",
    "If Lucy incorporates H2 first and the Korean corporation owns the LLC: the ODI report is filed by the corporation, the LLC's owner on Form 5472 becomes the corporation, Korean corporate tax applies to repatriated profit instead of personal income tax, and the Form W-8BEN-E for Amazon names the corporation. No model number moves. Direction only: the 세무사 decides which is cheaper after Korean tax on the Medium Year 5 net of " + mk(y5_net) + " (" + krwm(y5_net, fx) + " at " + fmt.num(fx) + ").",
    "Lucy with her 세무사",
    "Before the ODI report, so before the capital transfer",
    "Open",
])

rows.append([
    "D2", "Mark's role and agreement terms",
    "Advisor and US operations support under a consulting or commission agreement. Not an owner, not an agent who signs contracts. The model carries $2,000 per month from Year 3 (Medium) inside fixed costs.",
    "Keeps the treaty position (no dependent agent, no PE).",
    f"A retainer from Year 1 instead of Year 3 adds {mk(retainer_y)} per year of fixed cost: Medium Year 1 net {mk(M['annual']['Y1']['net'])} would fall to about {mk(M['annual']['Y1']['net'] - retainer_y)} and the 18 month net {mk(m18_net)} to about {mk(m18_net - 18 * 2000)}, and the peak cash need of {mk(peak)} rises by roughly the same {mk(18 * 2000)} (direction only, the model has no Year 1 retainer row). A commission on revenue instead of a retainer moves with the scenario: 2% of Medium Year 5 revenue is {mk(0.02 * M['annual']['Y5']['revenue'])}. Any wording that lets Mark negotiate or sign contracts for BRAND LLC risks a US PE (see decision 2). Agreement terms are a placeholder [PLACEHOLDER: Mark's agreement terms].",
    "Lucy and Mark",
    "Week 2, before Mark contacts the broker, bank or Amazon on the LLC's behalf",
    "Open",
])

rows.append([
    "D3", "Entity state",
    f"Wyoming single member LLC: {fmt.money(ent['formation_cost_usd'])} to form, about {fmt.money(ent['annual_cost_usd'])} per year, no state income tax.",
    "Brief. Delaware is only worth it for raising investment.",
    f"Delaware: about {fmt.money(ent['delaware_annual_cost_usd'])} per year. Over 5 years Wyoming costs about {fmt.money(wy_5y)} and Delaware about {fmt.money(de_5y)} plus a formation fee [PLACEHOLDER: Delaware formation and registered agent quote], so Delaware is about {fmt.money(de_5y - wy_5y)} more before agent fees. The difference is inside the compliance line of the model ({fmt.money(M['annual']['Y5']['compliance'])} per year). It only matters if an outside investor requires Delaware, which decision 3 does not assume.",
    "Lucy",
    "Week 1, with the LLC filing",
    "Open",
])

rows.append([
    "D4", "First SKUs",
    "5 launch SKUs: 8K 7M, 8K 10M (hero), 8K 15M, 8K 20M, 4K 30M. Air DDP at 100 to 300 units each.",
    "The 8K short runs were 53% of units and 54% of gross profit in the August deal (management estimate for the US mix).",
    f"Per SKU net at the H2 cost base on air freight (SKU Unit Economics tab): 8K 10M {m2(hero_h2['net_air'])} ({pct(hero_h2['net_air_pct'])}), 8K 20M {m2(sku['8K_20M']['net_air'])} ({pct(sku['8K_20M']['net_air_pct'])}), 4K 30M {m2(w4k['net_air'])} ({pct(w4k['net_air_pct'])}). The 4K line has the weakest economics (landed {pct(w4k['landed_air_pct'])} of price on air), so swapping 4K 30M for 8K 25M (net {m2(s25['net_air'])}, {pct(s25['net_air_pct'])}) raises blended margin and raises cash tied up per SKU (landed {m2(s25['landed_air'])} times 200 units is about {mk(s25['landed_air'] * 200)} per lot versus {mk(w4k['landed_air'] * 200)} for 4K 30M). Adding a sixth SKU adds one more lot of that size to inventory in transit, funded by the business.",
    "Lucy and Mark, after the factory confirms the per SKU FOB",
    "Week 4, before the test inventory PO",
    "Open",
])

rows.append([
    "D5", "Factory cost base",
    f"H2 cost base: August deal KRW cost divided by FX times 0.92 (FOB factor, management estimate). Hero FOB {m2(hero['fob_usd'])}, landed {m2(cb['h2']['landed'])} on air. The P&L uses the August assumption of {pct(hero['landed_pct'])} of revenue (hero landed {m2(cb['august']['landed'])}).",
    "The August deal is the only verified cost. The Alibaba benchmark is higher and is flagged as the largest unverified assumption in the plan.",
    f"At the Alibaba benchmark ($17 to $22 FOB for a comparable 30 ft AOC) the hero landed cost is {m2(cb['alibaba_low']['landed'])} to {m2(cb['alibaba_high']['landed'])} and the hero loses {m2(abs(cb['alibaba_low']['net']))} to {m2(abs(cb['alibaba_high']['net']))} per unit at $44.99; 18 month net becomes {mk(ali['m18_net'])} ({delta(ali['m18_net_vs_base'])}) and Year 5 net {mm(ali['y5_net'])} ({deltam(ali['y5_net_vs_base'])}). Factory cost plus 25%: 18 month net {mk(f25['m18_net'])} ({delta(f25['m18_net_vs_base'])}), Year 5 {mk(f25['y5_net'])}. Factory cost minus 10% (volume tiers, lever 7): 18 month net {mk(f10['m18_net'])} ({delta(f10['m18_net_vs_base'])}), Year 5 {mk(f10['y5_net'])}. A written USD FOB quote per SKU from the factory settles this and removes the FX effect in decision 4.",
    "The factory sets the price; Lucy negotiates",
    "Week 2, before the ruling request (the FOB is the customs value) and before the test inventory PO in week 4",
    "Open",
])

rows.append([
    "D6", "Freight mode",
    f"Air DDP at {fmt.money(X['logistics']['air_per_kg_usd'], 2)} per kg for the first lots (Months 1 to 6), ocean at {fmt.money(X['logistics']['ocean_per_kg_usd'], 2)} per kg for proven SKUs after the Day 90 gate (Months 7 onward, Medium and High). Low stays on air.",
    "Brief: ocean freight only on proven SKUs.",
    f"Hero landed cost {m2(hero_h2['landed_air'])} on air versus {m2(hero_h2['landed_ocean'])} on ocean (H2 base). Lever 2 is worth up to {mk(levers[2]['model_effect'])} on the 18 month net if every unit moved by ocean. Staying on air keeps lead time at 1 month and the in transit inventory lower; moving to ocean adds a second month of lead time and one more month of inventory funded by the business, which is why the peak cash need sits in {fmt.month_label(peak_m)}.",
    "Lucy with the broker and 3PL",
    "Day 90 gate (end of Month 3), first ocean PO in Month 4 for Month 7 arrival",
    "Open",
])

rows.append([
    "D7", "Retail prices and ASP",
    f"Hero sale price $44.99 (list $59.99), blended ASP $42 in Year 1 and 2, then minus 4% per year from Year 3 with a 3% roadmap mix uplift in Year 3 (Medium, Assumptions rows 40 and 41; the model deepened the spec starting value of minus 2% so margin compresses at least 1 point a year). Prices per SKU are management estimates pending the Amazon rate card and fee preview.",
    "August model.",
    f"ASP minus 10%: 18 month net {mk(aspm['m18_net'])} ({delta(aspm['m18_net_vs_base'])}), Year 5 net {mk(aspm['y5_net'])} ({delta(aspm['y5_net_vs_base'])}). ASP plus 10%: 18 month net {mk(aspp['m18_net'])} ({delta(aspp['m18_net_vs_base'])}), Year 5 net {mk(aspp['y5_net'])} ({delta(aspp['y5_net_vs_base'])}). Lever 10 tests plus or minus 5% with coupons.",
    "Mark proposes, Lucy approves",
    "Week 6, before listings go live (weeks 6 to 9); retest at the Day 90 gate",
    "Open",
])

rows.append([
    "D8", "Advertising spend phases",
    "Medium: 30% of revenue Months 1 to 3, 22% Months 4 to 6, 15% Months 7 to 12, 11% Months 13 to 18 and 19 to 24, then 11.5% in Years 3 to 5 (Assumptions rows 42 to 44; the model holds Medium above the spec starting values of 9.5%, 9.0%, 8.5% for the same margin compression reason). Floor 8% at the company level.",
    "August model; the 8% floor is a brief realism rule. High runs 9.5% in Months 19 to 24 so Year 2 stays above the floor (model note ads_months_19_to_24).",
    f"Every phase minus 3 points: 18 month net {mk(adm['m18_net'])} ({delta(adm['m18_net_vs_base'])}), Year 5 net {mk(adm['y5_net'])} ({delta(adm['y5_net_vs_base'])}). Plus 3 points: 18 month net {mk(adp['m18_net'])} ({delta(adp['m18_net_vs_base'])}), Year 5 net {mk(adp['y5_net'])} ({delta(adp['y5_net_vs_base'])}). TACOS in Month 3 is also a gate metric: pass at or below {pct(gt['d90_tacos_pass']['value'])}, kill above {pct(gt['d90_tacos_kill']['value'])}.",
    "Mark (campaign owner)",
    "Weekly from Month 1; phase review at the Day 90 gate",
    "Open",
])

rows.append([
    "D9", "Return rate",
    "3% in Year 1, 2.5% from Year 2 (Medium). Never below 2%.",
    "August model; lever 3 (direction proof connector) targets 2%.",
    f"At 2%: 18 month net {mk(ret2['m18_net'])} ({delta(ret2['m18_net_vs_base'])}), Year 5 net {mk(ret2['y5_net'])} ({delta(ret2['y5_net_vs_base'])}). At 5%: 18 month net {mk(ret5['m18_net'])} ({delta(ret5['m18_net_vs_base'])}), Year 5 net {mk(ret5['y5_net'])} ({delta(ret5['y5_net_vs_base'])}). Gate: pass at or below {pct(gt['d90_ret_pass']['value'])} at Day 90, kill above {pct(gt['d90_ret_kill']['value'])}; at Month 12 pass at or below {pct(gt['m12_ret_pass']['value'])}, kill above {pct(gt['m12_ret_kill']['value'])}.",
    "Lucy and the factory (connector design), Mark (listing clarity)",
    "Packaging artwork frozen by week 4; measured monthly",
    "Open",
])

rows.append([
    "D10", "Working capital terms",
    "Inventory cover 2.5 months of forward forecast (Medium; 3.0 Low, 2.0 High), PO deposit 30% at order and 70% at shipment, lead time 1 month air then 2 months ocean, Amazon payout lag 14 days, inventory in transit funded by the business. Years 3 to 5 inventory days 95 (Medium), 135 (Low), 85 (High) of landed COGS (model note inventory_days_years_3_to_5).",
    "Spec section 2.6, management estimates. The brief says working capital is what kills cable businesses.",
    f"These terms set the peak cash need of {mk(peak)} in {fmt.month_label(peak_m)} and the funding gap of {mk(gap)}. Direction only (no sensitivity row in the model): a factory that accepts 30% at order and 70% at 30 days after shipment, or a cover target of 2.0 months, lowers the peak need; a 100% prepayment factory or 3.0 months of cover raises it. Rerun the model with the Assumptions cells changed to size it.",
    "Lucy (factory terms), Mark (cover target and reorder points)",
    "Factory terms with the week 4 PO; cover target reviewed at the Day 90 gate",
    "Open",
])

rows.append([
    "D11", "Product liability insurance",
    f"$1M policy once sales pass {fmt.money(X['amazon']['insurance_trigger_monthly_sales_usd'])} in a month; the model books {fmt.money(1200)} per year (management estimate) from Month {ins_month['Medium']} in Medium (Month {ins_month['High']} High, Month {ins_month['Low']} Low).",
    "Amazon requirement in the brief; the premium is an estimate.",
    "A quoted premium replaces the estimate inside fixed costs. Direction only: each $1,000 per year of premium above the estimate reduces every annual net by $1,000. A refusal to insure a foreign owned LLC [PLACEHOLDER: insurance quote] would stop sales above $10K per month, which the Medium case reaches in Month 2.",
    "Lucy, with Mark sourcing quotes",
    f"Quote in week 3; policy bound before Month {ins_month['Medium']}",
    "Open",
])

rows.append([
    "D12", "Bank choice",
    "USD account for the LLC at Mercury or Relay; Payoneer or WorldFirst as the fallback.",
    "Brief. Amazon pays into the LLC account; the legal name must match the state filing and the EIN exactly.",
    "No model number moves. A fallback provider changes payout timing and fees (not modeled; direction only) and may not accept the ODI related capital transfer, so the bank must be chosen before the ODI report names the receiving account.",
    "Lucy, with Mark opening the application",
    "Weeks 2 to 3, after the EIN and before the capital transfer",
    "Open",
])

rows.append([
    "D13", "Gate thresholds",
    f"Day 90: conversion at or above {pct(gt['d90_conv_pass']['value'])} on 3 of 5 SKUs, TACOS at or below {pct(gt['d90_tacos_pass']['value'])}, return rate at or below {pct(gt['d90_ret_pass']['value'])}, contribution before advertising at or above {fmt.money(gt['d90_cpu_pass']['value'])} on 3 of 5 SKUs, runway at or above {fmt.num(gt['d90_run_pass']['value'])} months. Month 12: revenue at or above {mk(gt['m12_rev_pass']['value'])} per month, cumulative net at or above $0, TACOS at or below {pct(gt['m12_tacos_pass']['value'])}, return rate at or below {pct(gt['m12_ret_pass']['value'])}. Kill triggers in the Funding and Breakeven tab.",
    "Spec section 2.7. The thresholds live in one block on Assumptions so every file reads the same numbers.",
    f"Modeled results today: Medium Day 90 {M['gates']['day90']['result']} and Month 12 {M['gates']['month12']['result']}; High {H['gates']['day90']['result']} and {H['gates']['month12']['result']}; Low {L['gates']['day90']['result']} at Day 90 (TACOS {pct(L['gates']['day90']['metrics']['tacos']['value'])} and hero contribution {m2(L['gates']['day90']['metrics']['cpu']['value'])} miss the pass line without firing a kill) and {L['gates']['month12']['result']} at Month 12. Loosening the Month 12 revenue pass line below {mk(gt['m12_rev_pass']['value'])} would let the Low case (Month 12 revenue {mk(L['monthly'][11]['revenue'])}) continue to absorb capital; the brief wants the Low case to trigger the kill.",
    "Mark and Lucy",
    "Before Month 1; written into the Assumptions block, not changed after the gate date is known",
    "Open",
])

rows.append([
    "D14", "Low case lost quarter",
    "Months 19 to 21 at 10% of trend (listing suppression) in the Low scenario only (model note low_lost_quarter_months).",
    "The brief requires one lost quarter in Low; the spec's original Months 15 to 17 made the Low 18 month reconciliation impossible.",
    f"Moving it into Months 1 to 18 changes the Low 18 month block ({mk(L['m18']['revenue'])} revenue, {mk(L['m18']['net'])} net) and breaks the tie to the brief's ($19K); keep it in Months 19 to 21 unless the Low block is recalibrated. Low Year 2 net {mk(L['annual']['Y2']['net'])} already carries it.",
    "Mark",
    "Done in this build; reopen only with a model recalibration",
    "Open",
])

rows.append([
    "D15", "Channel expansion timing and Temu",
    "eBay and Walmart overflow and B2B installer from Year 2, DTC site and roadmap SKUs from Year 3, Amazon Canada and the pro line from Year 4. Temu is not in the base case.",
    "Brief channel order. Channel revenues in Years 2 to 5 are management estimates.",
    f"Medium Year 5 revenue {mm(M['annual']['Y5']['revenue'])} is {mm(M['annual']['Y5']['revenue_by_channel']['amazon_us'])} Amazon US and {mm(M['annual']['Y5']['revenue'] - M['annual']['Y5']['revenue_by_channel']['amazon_us'])} other channels (B2B {mk(M['annual']['Y5']['revenue_by_channel']['b2b'])}, DTC {mk(M['annual']['Y5']['revenue_by_channel']['dtc'])}, eBay and Walmart {mk(M['annual']['Y5']['revenue_by_channel']['ebay_walmart'])}, pro line {mk(M['annual']['Y5']['revenue_by_channel']['pro_line'])}, Amazon Canada {mk(M['annual']['Y5']['revenue_by_channel']['amazon_canada'])}). Delaying B2B by a year removes its Year 2 revenue of {mk(M['annual']['Y2']['revenue_by_channel']['b2b'])} and its contribution. Adding Temu through its local seller program after the Amazon price floor is set is optional and not modeled.",
    "Lucy and Mark at the Month 12 gate",
    "Month 12 gate for Year 2 channels; Year 2 close for Year 3",
    "Open",
])

rows.append([
    "D16", "Launch month",
    "Month 1 = the first month with inventory live on Amazon, assumed early 2027. No calendar months in the model.",
    "The build sequence needs 8 to 10 weeks before the first sale.",
    "A later Month 1 shifts every model month and every compliance due date that is relative to it (insurance, the Day 90 gate, the Month 12 gate, the first Form 5472 for the LLC's first tax year). No model number moves. A Month 1 in the fourth quarter would meet peak season demand with no reviews; not modeled.",
    "Lucy",
    "Week 8, when the test inventory ships",
    "Open",
])

rows.append([
    "D17", "4K line category name and certification",
    "The 4K line is labeled Premium High Speed HDMI Cable, pending certification (placeholder). The 8K line is Ultra High Speed HDMI Cable. No HDMI version numbers on packaging or in promotion.",
    "HDMI naming rule in the brief; the factory's ATC report decides the category each cable may carry.",
    "If the factory cannot show an ATC report for a category, that line cannot carry the category name and the listing kit titles and packaging for it must change; HDMI LA runs seizures. If the 4K line is not certified, the 4K 30M launch SKU drops out (see D4) and the other 4K lengths stay off the roadmap until it is.",
    "The factory provides; Lucy verifies; Mark updates the listing kit",
    "Week 2, before packaging artwork and before the test inventory PO",
    "Open",
])

A(table(hdr, rows))
A("")

# ---------------------------------------------------------------- Section 3
A("## 3. Data Mark and Lucy should provide to tighten the plan")
A("")
A("These inputs are management estimates or placeholders in build/figures.json and the SKU table of build/BUILD_SPEC.md section 2.1. Each one replaces an estimate with a fact when it arrives. Put the file in inputs/ and rerun the build.")
A("")
hdr3 = ["#", "Data item", "Who provides", "Used where in the package", "Current basis in the build", "By when"]
rows3 = [
    [1, "USD FOB quote by SKU from the factory (all 10 SKUs, with the volume tiers)", "Lucy, from the factory",
     "SKU Unit Economics tab, Landed Cost & Tariff Scenarios tab, customs value on the ruling request, decisions 4 and 5",
     f"Management estimate: August KRW cost divided by FX times 0.92 (hero FOB {m2(hero['fob_usd'])})", "Week 2"],
    [2, "Packaged weight and box dimensions by SKU", "Lucy, from the factory",
     "Air and ocean freight per unit, FBA size tier and fee, lever 8",
     "Management estimate: packaged weights 0.35 lb to 1.9 lb in the SKU table; no dimensions", "Week 2"],
    [3, "Factory HDMI Adopter certificate and ATC test reports per cable category", "Lucy, from the factory",
     "Compliance checklist, listing kit category names, packaging, D17",
     "[PLACEHOLDER: HDMI Adopter certificate number from factory] and [PLACEHOLDER: ATC test report]", "Week 2"],
    [4, "Factory FCC test report for the active heads", "Lucy, from the factory",
     "Compliance checklist, listing kit FCC statement", "[PLACEHOLDER: FCC test report]", "Week 2"],
    [5, "CL3 documentation", "Lucy, from the factory",
     "In wall claim in the listing kit, differentiators, pro line", "[PLACEHOLDER: CL3 documentation]; no in wall claim until it arrives", "Week 4"],
    [6, "August deal per SKU unit mix and sell prices (the csv carries implied sell prices only)", "Lucy, from H2 Company records",
     "SKU mix weights for the launch SKUs, Plan B Korea wholesale pace, D4",
     "Evidence based deal totals; per SKU sell prices implied from the 20.3% deal margin", "Week 1"],
    [7, "Actual Amazon fee preview per SKU (referral, FBA fee by size tier, storage) from Seller Central", "Mark",
     "FBA fee column in the SKU table, Monthly PL Y1 Y2, D7",
     f"Management estimate: FBA {fmt.money(sku['8K_7M']['fba_fee'], 2)} to {fmt.money(sku['8K_50M']['fba_fee'], 2)} per unit, referral 15%", "Week 4, once the account is open"],
    [8, "Product liability insurance quote for a foreign owned LLC", "Mark sources, Lucy signs",
     "Fixed costs, compliance checklist, D11", "Management estimate $1,200 per year", "Week 3"],
    [9, "Customs broker quote: entry fee, continuous bond, ruling request fee, and the List 3 check on 8544.70.00", "Mark",
     "Broker and bond per unit, duty scenario table, decision 5",
     f"Management estimate: {fmt.money(X['logistics']['broker_entry_fee_usd'])} per entry, {fmt.money(X['logistics']['broker_bond_annual_usd'])} per year bond, {fmt.money(X['logistics']['broker_per_unit_usd'], 2)} per unit", "Week 1"],
    [10, "Bank choice and the receiving account for the capital transfer", "Lucy, with Mark",
     "ODI report, one page summaries, D12", "Mercury or Relay assumed; fallback Payoneer or WorldFirst", "Weeks 2 to 3"],
    [11, "Brand name shortlist (3 names) for the attorney clearance search", "Lucy and Mark",
     "Trademark filing, LLC name, Brand Registry, every file (BRAND placeholder), decision 1", "BRAND placeholder", "Week 1"],
    [12, "Mark's agreement terms (retainer or commission, scope, no authority to sign)", "Lucy and Mark",
     "Fixed costs from Year 3, treaty position, decision 2 and D2", "$2,000 per month from Year 3 (Medium), management estimate", "Week 2"],
    [13, "FX rate and date of the capital transfer, and the amount declared on the ODI report", "Lucy, from her foreign exchange bank",
     "FX input cell, one page summaries, decisions 3 and 4", f"{fmt.num(fx)} KRW per USD, {mk(opening)} opening capital", "At the capital transfer"],
    [14, "Lender or investor terms if the funding gap is not self funded", "Lucy, Mark",
     "Funding plan rows in Funding and Breakeven, decision 3", "[PLACEHOLDER: lender or investor and terms]", f"{fmt.month_label(need_by)}"],
    [15, "3PL quote (receiving, storage, pick pack ship, FBA replenishment)", "Mark",
     "Storage and inbound %, B2B fulfillment $4 per unit, D6", "Management estimates", "Week 4"],
    [16, "Market size or category sales benchmark for long HDMI on Amazon (a sourced figure)", "Mark",
     "Deck slide 6 segment sizes", "Placeholders; no market size is stated anywhere in the package", "Before the deck is sent to a lender or investor"],
]
A(table(hdr3, rows3))
A("")

# ---------------------------------------------------------------- Section 4
A("## 4. How to record a decision")
A("")
A("1. Edit the Default used cell in inputs/DECISIONS.md (or add a row for a D item).")
A("2. For a model input (FX, duty scenario, ASP, ad phases, returns, cover months, PO terms), change the yellow input cell on the Assumptions tab of output/03_Financial_Model_5Y.xlsx or the matching value in build/build_model.py, then rerun build/build_model.py, the recalc, build/pull_figures.py and build/tie_out.py so every file picks up the new figures.")
A("3. Change this row's Status from Open to Decided with the date and the name of the person who decided.")
A("")
A(f"Reference figures at the delivered selector state (Medium, Copper, FX {fmt.num(fx)}): 18 month revenue {mm(M['m18']['revenue'])}, net {mk(m18_net)} ({pct(M['m18']['net_margin'])}), cumulative breakeven {fmt.month_label(M['m18']['breakeven_month'])}; Year 5 revenue {mm(M['annual']['Y5']['revenue'])}, net {mk(y5_net)} ({pct(M['annual']['Y5']['net_margin'])}); peak cash need {mk(peak)} in {fmt.month_label(peak_m)}; funding gap {mk(gap)} by {fmt.month_label(need_by)}.")
A("")

out = ROOT / "output" / "11_Decision_Log.md"
out.write_text("\n".join(doc) + "\n", encoding="utf-8")
print("wrote", out, len(doc), "lines")

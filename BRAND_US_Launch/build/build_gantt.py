#!/usr/bin/env python3
"""Build output/07_90_Day_Launch_Plan.xlsx (week by week Gantt, Costs, Gate).
Every model number comes from build/figures.json through build/fmt.py; static facts from build/facts.json;
working capital inputs from build/sops_data.json. Weeks follow the brief's critical path (weeks 1 to 22).
Cost basis: Evidence based (brief, August model, verified fact) or Management estimate.
No hyphens, en dashes or em dashes in any text (official form and registration numbers keep their format).
Usage: python3 build_gantt.py            (writes the workbook)
       python3 build_gantt.py --patch-fonts   (after $RECALC: force every font back to Arial in styles.xml)
"""
import json, os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import fmt as F
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName

ROOT = os.environ.get("ROOT", "/home/user/lakehouse-tracker/BRAND_US_Launch")
BUILD = os.path.join(ROOT, "build"); OUT = os.path.join(ROOT, "output")
PATH = os.path.join(OUT, "07_90_Day_Launch_Plan.xlsx")

def patch_fonts(path):
    """After $RECALC (LibreOffice store) cells with Hangul carry a fallback font name. Rewrite xl/styles.xml so every font is Arial."""
    import zipfile, shutil, re
    tmp = path + ".tmp"
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "xl/styles.xml":
                txt = data.decode("utf8")
                names = set(re.findall(r'<name val="([^"]+)"', txt))
                txt = re.sub(r'<name val="(?!Arial")[^"]+"', '<name val="Arial"', txt)
                print("styles.xml font names before patch:", sorted(names))
                data = txt.encode("utf8")
            zout.writestr(item, data)
    shutil.move(tmp, path)

if "--patch-fonts" in sys.argv:
    patch_fonts(PATH); sys.exit(0)

facts = json.load(open(os.path.join(BUILD, "facts.json")))
fig = json.load(open(os.path.join(BUILD, "figures.json")))
sd = json.load(open(os.path.join(BUILD, "sops_data.json")))

EB, ME = "Evidence based", "Management estimate"
FONT, SZ = "Arial", 10
FILL_IN = PatternFill("solid", fgColor="FFFF00"); BLUE = "0000FF"; GREEN = "008000"; GREY = "595959"
FILL_HDR = PatternFill("solid", fgColor="D9E1F2"); FILL_SEC = PatternFill("solid", fgColor="EDEDED"); FILL_TOT = PatternFill("solid", fgColor="E2EFDA")
FILL_GATE = PatternFill("solid", fgColor="FFE699")
FMT_MONEY = '$#,##0;($#,##0);0'; FMT_MONEY2 = '$#,##0.00;($#,##0.00);0'; FMT_PCT = '0.0%'; FMT_NUM1 = '0.0'
THIN = Side(style="thin", color="BFBFBF"); BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
BUILT = fig["meta"]["built"]
PH = lambda s: "[PLACEHOLDER: " + s + "]"
money = F.money

# ---------------- model and fact numbers used in the plan ----------------
M = fig["scenarios"]["Medium"]; L = fig["scenarios"]["Low"]; H = fig["scenarios"]["High"]
GT = fig["gate_thresholds"]; KC = fig["kill_criteria"]
aug = facts["august"]; m18 = aug["m18_assumptions"]
OPEN_CAP = M["cash"]["opening_capital"]                      # 39000
LAUNCH_ONE = m18["launch_one_time"]                          # 9500
FIXED_M = m18["fixed_monthly"]                               # 850
SELLER_FEE = m18["seller_fee_monthly"]                       # 39.99
WC = sd["wc_inputs"]; INS_YEAR = WC["insurance_per_year"]["Medium"]; DEP = WC["po_deposit"]["Medium"]
INS_TRIG = facts["amazon"]["insurance_trigger_monthly_sales_usd"]; INS_COVER = facts["amazon"]["product_liability_insurance_usd"]
INS_MONTH = sd["insurance_trigger_month"]["Medium"]
TM_FEE = facts["trademark"]["uspto_fee_per_class_usd"]; FORM_FEE = facts["entity"]["formation_cost_usd"]; BOND = facts["logistics"]["broker_bond_annual_usd"]
PEN_5472 = facts["entity"]["form_5472_penalty_usd"]
hero = fig["unit_economics"]["hero"]; FOB = hero["fob_usd"]
mon = M["monthly"]
ASP1 = mon[0]["asp"]; LANDED1 = mon[0]["landed_pct"] * ASP1; FEES_PER_UNIT = LANDED1 - FOB
M0_OUT = OPEN_CAP - M["cash_balance_month0"]                 # Month 0 inventory cash out (lot 1 FOB)
INV_OUT = [M0_OUT] + [mon[i]["inventory_cash_out"] for i in range(3)]   # Months 0 to 3
ARR = [mon[i]["arrivals"] for i in range(4)]                 # units arriving Months 1 to 4
# Lot k arrives in Month k (air, one month lead): FOB paid the month before (30% deposit at order, 70% at shipment); freight, duty and fees paid on arrival
LOT_FOB = [ARR[i] * FOB for i in range(4)]
LOT_FEES = [ARR[i] * FEES_PER_UNIT for i in range(4)]
# consistency with the model's monthly inventory cash out (Months 0 to 3)
chk = [LOT_FOB[0], LOT_FEES[0] + LOT_FOB[1], LOT_FEES[1] + LOT_FOB[2], LOT_FEES[2] + LOT_FOB[3]]
for a, b in zip(chk, INV_OUT):
    assert abs(a - b) < 1.0, (a, b)
INV_TOTAL = sum(INV_OUT)
DUTY_RATE = fig["duty"]["copper"]["rate"]
units_per_sku = lambda u: u / 5.0
GAP = M["cash"]["funding_gap"]; GAP_MONTH = M["cash"]["funding_needed_by_month"]; PEAK = M["cash"]["peak_need"]; PEAK_MONTH = M["cash"]["peak_need_month"]
CASH_M3 = mon[2]["cash_balance"]; CASH_M3_F = mon[2]["cash_balance_after_funding"]
D90 = M["gates"]["day90"]; D90L = L["gates"]["day90"]; D90H = H["gates"]["day90"]
M12 = M["gates"]["month12"]; M12L = L["gates"]["month12"]; M12H = H["gates"]["month12"]

# Month to week mapping (the brief: listings weeks 6 to 9, validation weeks 10 to 22; Day 90 gate at week 22)
MONTH_WEEKS = {0: (1, 9), 1: (10, 13), 2: (14, 17), 3: (18, 22)}
def wk_label(a, b): return "week %d" % a if a == b else "weeks %d to %d" % (a, b)

# ---------------- task list ----------------
OWNERS = ["Lucy", "Mark", "attorney", "broker", "factory", "3PL", "VA"]
OWNER_FILL = {"Lucy": "9BE7F5", "Mark": "C9B8F7", "attorney": "F7C6E8", "broker": "FFD28F", "factory": "B7E4C7", "3PL": "CFE2F3", "VA": "E6E0B0"}
COST_TYPES = ["Launch budget", "Inventory", "Fixed cost", "Insurance", "No cash cost"]
TASKS = []   # dict rows or ("section", title)
def sec(t): TASKS.append(("section", t))
def T(id_, task, owner, support, start, end, dep, cost, ctype, deliverable, basis):
    assert owner in OWNERS and basis in (EB, ME) and ctype in COST_TYPES and 1 <= start <= end <= 22, id_
    TASKS.append({"id": id_, "task": task, "owner": owner, "support": support, "start": start, "end": end, "dep": dep, "cost": round(float(cost), 2), "ctype": ctype, "deliv": deliverable, "basis": basis})

sec("Weeks 1 to 4: brand, entity, Korea filing, bank, customs (brief critical path items 1 to 5)")
T("A1", "Decide the US brand name so the legal name BRAND LLC is final before any filing", "Lucy", "Mark", 1, 1, "", 0, "No cash cost",
  "Name recorded in 11_Decision_Log.md; the same spelling is then used at the state, the IRS, the bank and Amazon", EB)
T("A2", "Trademark clearance search and USPTO (United States Patent and Trademark Office) application in one class through a US licensed attorney", "attorney", "Lucy", 1, 2, "A1", 1200, "Launch budget",
  "USPTO serial number " + PH("serial number") + ". USPTO fee " + money(TM_FEE) + " per class (brief); attorney fee " + money(1200 - TM_FEE) + " is a management estimate " + PH("attorney quote"), ME)
T("A3", "Wyoming single member LLC: articles of organization, registered agent, operating agreement naming Lucy as 100% owner", "attorney", "Lucy", 1, 2, "A1", 700, "Launch budget",
  "Certificate of organization. State fee " + money(FORM_FEE) + " (brief); registered agent and operating agreement " + money(600) + " are management estimates", ME)
T("A4", "EIN (Employer Identification Number) by phone or fax with Form SS-4; no SSN (Social Security Number) needed", "attorney", "Lucy", 2, 2, "A3", 0, "No cash cost",
  "IRS (Internal Revenue Service) EIN confirmation letter " + PH("EIN") + ". The EIN is the importer of record number and the tax ID for the bank and Amazon", EB)
T("A5", "Written consulting or commission agreement between BRAND LLC and Mark Taylor (no equity, no authority to negotiate or sign)", "attorney", "Lucy, Mark", 2, 3, "A3", 500, "Launch budget",
  "Signed agreement. Keeps the treaty position: no dependent agent in the US. Fee is a management estimate " + PH("attorney quote"), ME)
T("A6", "Document pack with the exact legal name: certificate, EIN letter, operating agreement, owner ID, US mailing address", "Lucy", "Mark", 2, 3, "A4", 250, "Launch budget",
  "One folder used for the bank, Amazon and the broker. Notary, mailbox and courier costs are a management estimate", ME)
T("A7", "Korean overseas direct investment (ODI) report at the foreign exchange bank, before any capital moves", "Lucy", "Korean tax accountant (세무사)", 1, 2, "A3", 300, "Launch budget",
  "Bank acceptance of the ODI report. Hard stop: no transfer before acceptance. 세무사 support fee is a management estimate " + PH("세무사 quote"), ME)
T("A8", "USD bank account for BRAND LLC (Mercury or Relay; Payoneer or WorldFirst as fallback)", "Lucy", "Mark", 2, 3, "A4, A6", 0, "No cash cost",
  "Account open under the exact legal name and EIN", EB)
T("A9", "Capital transfer of the opening capital " + money(OPEN_CAP) + " (KRW to USD at the FX input), then the Korean remittance and securities acquisition reports", "Lucy", "Korean tax accountant (세무사)", 3, 3, "A7, A8", 0, "No cash cost",
  "Funds in the LLC account; bank acknowledgments of the follow up reports. Opening capital is the August funding ask (DECISIONS.md decision 3)", EB)
T("A10", "US CPA engaged: treaty memo (no permanent establishment from FBA stock), Form 5472 calendar, California nexus question", "Lucy", "Mark", 3, 4, "A4", 0, "No cash cost",
  "Engagement letter and memo. Annual Form 5472 preparation " + money(WC["form_5472_prep"]["Medium"]) + " sits in the model from Year 2; penalty for a missed filing " + money(PEN_5472), ME)
T("B1", "Customs broker engaged: power of attorney, importer of record set up on the EIN", "broker", "Mark", 1, 1, "", 150, "Launch budget",
  "Signed broker agreement. Setup fee is a management estimate " + PH("broker quote"), ME)
T("B2", "Continuous customs bond (" + money(BOND) + " per year)", "broker", "Lucy", 2, 2, "A4, B1", BOND, "Launch budget",
  "Bond number " + PH("bond number") + ". Required for formal entry; de minimis is gone", EB)
T("B3", "CBP (US Customs and Border Protection) binding ruling request filed in week 1: optical fiber cable 8544.70.00 versus electric conductor 8544.42.90", "broker", "Mark, factory", 1, 1, "B1", 800, "Launch budget",
  "Ruling request number " + PH("request number") + ". Preparation fee is a management estimate. Answer takes 30 to 90 days", ME)
T("B4", "Ruling answer window (30 to 90 days): track, answer CBP questions, keep the copper rate (" + F.pct(DUTY_RATE) + ") in every budget until the letter arrives", "broker", "Mark", 2, 14, "B3", 0, "No cash cost",
  "Ruling letter " + PH("ruling number") + ". A fiber ruling moves landed COGS from " + F.pct(fig["levers"][0]["from"]) + " to " + F.pct(fig["levers"][0]["to"]) + " (Levers tab, lever 1)", EB)

sec("Weeks 2 to 9: Amazon account, product compliance, test inventory, logistics (critical path items 6 and 7)")
T("C1", "Amazon Professional seller account: identity verification, tax interview as a foreign owned entity, bank link", "Lucy", "Mark", 3, 4, "A6, A8", 0, "No cash cost",
  "Seller Central account active. The " + money(SELLER_FEE, 2) + " monthly fee is counted from Month 1 in task F2", EB)
T("C2", "Amazon Brand Registry with the pending trademark application", "Mark", "Lucy", 4, 4, "A2, C1", 0, "No cash cost",
  "Brand Registry approval: unlocks A+ content, Brand Store, video ads and hijacker protection", EB)
T("C3", "Product liability insurance quote (" + money(INS_COVER) + " cover, required once sales pass " + money(INS_TRIG) + " in a month) and binder deposit", "Mark", "Lucy", 4, 5, "A8", 300, "Launch budget",
  "Quote on file " + PH("insurer quote") + ". Deposit is a management estimate inside the launch budget", ME)
T("D1", "Factory document pack: HDMI (High Definition Multimedia Interface) Adopter certificate and ATC (HDMI Authorized Test Center) test report, FCC (Federal Communications Commission) test report for the active heads, CL3 (in wall cable fire rating, Class 3) documentation", "factory", "Lucy", 2, 4, "", 0, "No cash cost",
  "Hard stop before the purchase order. " + PH("HDMI Adopter certificate number") + ", " + PH("FCC report number") + ", " + PH("CL3 report"), EB)
T("D2", "Packaging artwork: Ultra High Speed HDMI Cable category name (no version numbers), Prop 65 label, SOURCE and DISPLAY end markings, QR (quick response) code to the setup video, BRAND placeholder", "Lucy", "factory, Mark", 3, 5, "A1, D1", 400, "Launch budget",
  "Approved artwork proof. Artwork fee is a management estimate", ME)
T("D3", "Samples of the 5 launch SKUs shipped to Memphis for photography, direction proof test and packaging check", "factory", "Mark", 3, 4, "", 500, "Launch budget",
  "Samples received and tested. Sample and courier cost is a management estimate", ME)
T("D4", "Lot 1 purchase order: 5 launch SKUs, about %s units (about %s per SKU, inside the brief's 100 to 300). FOB %s: %s deposit at order, %s at shipment" % (F.units(ARR[0]), F.units(units_per_sku(ARR[0])), money(LOT_FOB[0]), F.pct(DEP, 0), F.pct(1 - DEP, 0)),
  "Lucy", "factory", 4, 7, "A9, D1, D2", LOT_FOB[0], "Inventory",
  "Signed PO and payment receipts. Units and FOB are the Medium Month 1 arrivals from the model (Cash Flow and Working Capital tab)", EB)
T("D5", "3PL onboarding: account, receiving rules, FBA prep profile, Amazon shipping plan template", "3PL", "Mark", 5, 6, "C1", 400, "Launch budget",
  "3PL agreement " + PH("3PL quote") + ". Setup fee is a management estimate", ME)
T("D6", "Third party QC (quality control) inspection before departure and FNSKU (Fulfillment Network Stock Keeping Unit, Amazon's barcode) labels applied at the factory (" + money(facts["logistics"]["qc_label_per_unit_usd"], 2) + " per unit in the landed cost)", "factory", "Lucy", 7, 7, "D4", 0, "No cash cost",
  "QC report " + PH("inspection report") + ". Per unit cost is inside the lot 1 arrival charges (D7)", EB)
T("D7", "Lot 1 air DDP (delivered duty paid) shipment, formal entry by the broker, duty at the copper rate, 3PL receiving and FBA inbound prep: " + money(LOT_FEES[0]) + " on arrival", "broker", "factory, 3PL", 8, 8, "B2, D5, D6", LOT_FEES[0], "Inventory",
  "Entry summary and 3PL receiving report. Freight, duty, broker, inbound and QC per unit from the model (landed cost less FOB). Test inventory lands in week 8 (brief critical path item 7)", EB)
T("D8", "First FBA inbound shipment from the 3PL; stock checked in before Month 1", "3PL", "Mark", 8, 9, "D7", 0, "No cash cost",
  "FBA inventory available in Seller Central", EB)

sec("Weeks 5 to 9: listings and content (critical path item 8)")
T("E1", "Photography and video for the 7 image slots and the video slot (09_Amazon_Listing_Kit.docx image brief)", "Mark", "VA", 5, 7, "D3", 2300, "Launch budget",
  "Final images and video files. Studio cost is a management estimate " + PH("studio quote"), ME)
T("E2", "Listing copy: one parent ASIN per line with children by length, titles under 200 characters, five bullets, backend keywords, HDMI naming rule", "VA", "Mark", 6, 7, "C2", 500, "Launch budget",
  "Listing flat file uploaded. VA hours are a management estimate", ME)
T("E3", "A+ content (7 modules) and Brand Store", "VA", "Mark", 7, 9, "C2, E1, E2", 300, "Launch budget",
  "A+ pages approved by Amazon. VA hours are a management estimate", ME)
T("E4", "Keyword research, PPC (pay per click advertising) campaign structure (auto plus manual), negative keyword list, ad budget caps", "Mark", "VA", 8, 9, "E2", 150, "Launch budget",
  "Campaigns built and paused until stock is live. Keyword tool subscription is a management estimate", ME)
T("E5", "Spec Match finder page and trade up registration form (off Amazon email list), QR setup video live", "Mark", "VA", 8, 10, "E1", 250, "Launch budget",
  "Pages live and linked from the insert card. Site tool cost is a management estimate", ME)
T("E6", "Listings live with the price ladder (list price versus sale price) and launch coupon; PPC campaigns switched on", "Mark", "Lucy", 9, 10, "D8, E3, E4", 0, "No cash cost",
  "Month 1 starts (first month with inventory live on Amazon)", EB)

sec("Weeks 10 to 22: 90 day validation, Months 1 to 3 (critical path item 9)")
T("F1", "Month 1 selling: Medium plan %s units, revenue %s; Vine enrollment and insert card for review velocity; daily listing and account health check" % (F.units(mon[0]["units"]), money(mon[0]["revenue"])),
  "Mark", "Lucy, VA", 10, 13, "E6", 0, "No cash cost",
  "Weekly KPI sheet: conversion, TACOS, return rate, contribution per SKU. Advertising is in the P&L at " + F.pct(mon[0]["ads_pct"], 0) + " of revenue, not in this cost column", EB)
T("F2", "Fixed costs Months 1 to 3: software, insurance and misc %s per month plus the Professional seller fee %s per month" % (money(FIXED_M), money(SELLER_FEE, 2)),
  "Lucy", "", 10, 22, "C1", 3 * (FIXED_M + SELLER_FEE), "Fixed cost",
  "Paid monthly from the LLC account. August model fixed cost lines", EB)
T("F3", "Lot 2 purchase order in Month 1 (arrives Month 2): about %s units, FOB %s (%s deposit, %s at shipment)" % (F.units(ARR[1]), money(LOT_FOB[1]), F.pct(DEP, 0), F.pct(1 - DEP, 0)),
  "Lucy", "factory", 10, 12, "D4, E6", LOT_FOB[1], "Inventory",
  "PO and payment receipts, placed once listings are live (E6). Order up to level from the reorder table (05_Operating_Processes.docx section 5)", EB)
T("F4", "Lot 2 arrival: air freight, duty, broker, inbound and QC " + money(LOT_FEES[1]), "broker", "factory, 3PL", 14, 14, "F3", LOT_FEES[1], "Inventory",
  "Entry summary and 3PL receiving report", EB)
T("F5", "Month 2 selling: Medium plan %s units, revenue %s; first keyword pruning and negative keyword pass; coupon review" % (F.units(mon[1]["units"]), money(mon[1]["revenue"])),
  "Mark", "VA", 14, 17, "F1", 0, "No cash cost",
  "Weekly KPI sheet; TACOS must be falling month over month by the gate", EB)
T("F6", "Lot 3 purchase order in Month 2 (arrives Month 3): about %s units, FOB %s" % (F.units(ARR[2]), money(LOT_FOB[2])), "Lucy", "factory", 14, 16, "F3, F4", LOT_FOB[2], "Inventory",
  "PO and payment receipts, placed once lot 2 has arrived (F4) and the Month 1 sales are known", EB)
T("F7", "Funding line in place by Month %d: the model draws %s (1.2 times the peak need of %s in Month %d) so the cash balance never goes negative" % (GAP_MONTH, F.money_k(GAP), F.money_k(PEAK), PEAK_MONTH),
  "Lucy", "Mark", 12, 17, "A9, A10", 0, "No cash cost",
  "Signed line of credit or investor agreement " + PH("lender or investor") + ". Korean change report before the money moves. Amount from the Funding and Breakeven tab (Medium)", EB)
T("F8", "Lot 3 arrival: air freight, duty, broker, inbound and QC " + money(LOT_FEES[2]), "broker", "factory, 3PL", 18, 18, "F6", LOT_FEES[2], "Inventory",
  "Entry summary and 3PL receiving report", EB)
T("F9", "Month 3 selling: Medium plan %s units, revenue %s; hold ad caps; collect the gate evidence" % (F.units(mon[2]["units"]), money(mon[2]["revenue"])), "Mark", "VA", 18, 22, "F5", 0, "No cash cost",
  "Weekly KPI sheet and the Gate tab filled in with actuals", EB)
T("F10", "Lot 4 purchase order in Month 3 (arrives Month 4, after the gate): about %s units, FOB %s; sized down if the gate reads REVIEW or KILL" % (F.units(ARR[3]), money(LOT_FOB[3])), "Lucy", "factory", 18, 21, "F6, F8", LOT_FOB[3], "Inventory",
  "PO and payment receipts, placed once lot 3 has arrived (F8). Arrival charges fall in Month 4, outside this plan", EB)
INS_W = MONTH_WEEKS[INS_MONTH - 1][1]   # bind in the last week of the month before sales pass the trigger
T("F11", "Bind the %s product liability policy (%s per year) before Month %d, when Medium monthly sales pass %s" % (money(INS_COVER), money(INS_YEAR), INS_MONTH, money(INS_TRIG)), "Mark", "Lucy", INS_W - 1, INS_W, "C3", INS_YEAR, "Insurance",
  "Policy certificate " + PH("policy number") + ". Premium is the model input (management estimate)", ME)
T("F12", "Weekly operating cadence: Monday KPI review (Mark prepares, Lucy decides), inventory position, account health, returns log", "Mark", "Lucy", 10, 22, "F1", 0, "No cash cost",
  "13 weekly notes in the shared folder (05_Operating_Processes.docx section 7)", EB)

sec("Week 22: Day 90 gate (critical path items 9 and 10)")
T("G1", "Day 90 gate pack: actuals for every criterion on the Gate tab, cash runway after funding the next PO, listing and IP status", "Mark", "VA", 21, 22, "F9, F12", 0, "No cash cost",
  "GATE: Gate tab result PASS, REVIEW or KILL. Medium model result at Day 90: " + D90["result"], EB)
T("G2", "Gate decision by Lucy: PASS scales winners on ocean freight and delists losers; REVIEW fixes and retests within 30 days; KILL stops further capital and moves to Plan B", "Lucy", "Mark", 22, 22, "G1", 0, "No cash cost",
  "GATE: signed decision note in 11_Decision_Log.md. Ocean PO is sized at the gate from the reorder table; not in this budget", EB)

ids = [t["id"] for t in TASKS if isinstance(t, dict)]
assert len(ids) == len(set(ids))
for t in TASKS:
    if isinstance(t, dict) and t["dep"]:
        for d in t["dep"].split(", "):
            assert d in ids, (t["id"], d)
            prev = next(x for x in TASKS if isinstance(x, dict) and x["id"] == d)
            # same rule as the workbook's Dependency check: a predecessor may overlap, but it never starts after, or ends after, the task that depends on it
            assert prev["start"] <= t["start"] and prev["end"] <= t["end"], (t["id"], d, prev["start"], prev["end"], t["start"], t["end"])
        assert len(t["dep"].split(", ")) <= 4, t["id"]
MAX_DEPS = 4   # dependency slots parsed from the Depends on column by formula
N_TASKS = len(ids)
LAUNCH_SUM = round(sum(t["cost"] for t in TASKS if isinstance(t, dict) and t["ctype"] == "Launch budget"), 2)
assert abs(LAUNCH_SUM - LAUNCH_ONE) < 0.01, LAUNCH_SUM

# ---------------- workbook helpers ----------------
wb = Workbook()
def W(ws, r, c, v, bold=False, color="000000", fill=None, fmt=None, wrap=False, size=None, italic=False, align=None, border=False, valign="top"):
    cell = ws.cell(row=r, column=c, value=v)
    cell.font = Font(name=FONT, size=size or SZ, bold=bold, italic=italic, color=color)
    if fill: cell.fill = fill
    if fmt: cell.number_format = fmt
    cell.alignment = Alignment(wrap_text=wrap, vertical=valign, horizontal=align)
    if border: cell.border = BOX
    return cell
def INP(ws, r, c, v, fmt=None, border=True, align=None):
    return W(ws, r, c, v, color=BLUE, fill=FILL_IN, fmt=fmt, border=border, align=align)
def name(n, ref):
    wb.defined_names[n] = DefinedName(n, attr_text=ref)

# ---------------- Cover ----------------
cv = wb.active; cv.title = "Cover"
CV_A, CV_B = 48, 100
cv.column_dimensions["A"].width = CV_A; cv.column_dimensions["B"].width = CV_B
def cv_row(k, v, bold_k=True):
    """Label in A, wrapped text in B, row height from the text length at the column B width."""
    global r
    W(cv, r, 1, k, bold=bold_k, wrap=True); W(cv, r, 2, v, wrap=True)
    cv.row_dimensions[r].height = max(15, 13 * (len(v) // (CV_B - 8) + 1) + 2); r += 1
r = 1
W(cv, r, 1, "BRAND LLC (limited liability company): 90 day launch plan (weeks 1 to 22)", bold=True, size=14); r += 1
W(cv, r, 1, "Built " + BUILT + ". Weeks follow the critical path in BUILD_BRIEF.md. Every model number comes from build/figures.json (Medium case, copper duty, FX " + F.num(fig["meta"]["fx"]) + " KRW per USD).", italic=True, color=GREY); r += 2
rows = [
    ("Tabs", "Gantt: one row per task with ID, Task, Owner, Support, Start week, End week, Depends on, Cost, Cost type, Gate or deliverable, Basis, Status, then the week grid 1 to 22 shaded by conditional formatting from the Start and End columns. Costs: SUMIFS by owner, by cost type, by week and by owner and week, with a check against the model. Gate: the Day 90 criteria with the Medium model values beside each threshold and an input column for the actuals."),
    ("How to use", "Lucy and Mark update the Status column (dropdown) and the Current week cell on the Gantt tab every Monday. Change Start or End weeks when a task slips; the grid, the Costs tab and the Dependency check recalculate. Edit the Depends on column as plain text (task IDs separated by a comma and a space, up to four); the check reads it. At week 22 fill the yellow Actual cells on the Gate tab; the result formula reads PASS, REVIEW or KILL."),
    ("Dependency check", "Column O on the Gantt tab flags a task (1 per problem, 0 = ok) when a task it depends on starts after it starts, ends after it ends, or has an ID that is not in the list. Tasks may overlap their predecessors (for example the bank application starts while the document pack is finished). The helper columns to the right of the week grid (Dep 1 to Dep 4 and their flags) parse the Depends on text; they are formulas and need no editing."),
    ("Weeks and months", "Month 1 is the first month with inventory live on Amazon (assumed early 2027). Weeks 1 to 9 are pre launch (Month 0), Month 1 is weeks 10 to 13, Month 2 is weeks 14 to 17, Month 3 is weeks 18 to 22. The Day 90 gate is the end of Month 3, week 22. No calendar dates are assigned. The weeks match the brief's critical path: broker and binding ruling in week 1, test inventory landed by week 8, listings weeks 6 to 9, validation weeks 10 to 22."),
    ("Owners", "Lucy (owner, decides and pays), Mark (US operations support under a consulting agreement, prepares), attorney (US licensed), broker (customs broker), factory (COMQ supplier in China), 3PL (third party logistics warehouse in the US), VA (virtual assistant). The Korean tax accountant (세무사) and the US CPA (certified public accountant) appear in the Support column."),
    ("Cost column", "USD cash out in the week range of the task. Launch budget tasks sum to the model's one time launch cost of " + money(LAUNCH_ONE) + " (evidence based total, management estimate split). Inventory rows are the Medium case purchase orders for Months 0 to 3 from the model's Cash Flow and Working Capital tab. Advertising is not a task cost: it sits in the P&L as a percent of revenue."),
    ("Legend", "Yellow fill with blue text = input. Black = formula or fixed text. Green = reads another tab. Grid colors by owner: Lucy, Mark, attorney, broker, factory, 3PL, VA (see the swatches below). A Done task shows a green grid. Amber rows are gate milestones."),
    ("Basis", "Evidence based = brief section 1, the August model, a verified regulatory fact or a model output. Management estimate = a cost or duration not yet quoted; the deliverable text names the placeholder to fill."),
    ("Rule", "No hyphens or dashes in any text. Official form and registration numbers keep their format (Form SS-4, 599-64-00496)."),
]
for k, v in rows: cv_row(k, v)
r += 1
W(cv, r, 1, "Owner swatches", bold=True); r += 1
for o in OWNERS:
    W(cv, r, 1, o, fill=PatternFill("solid", fgColor=OWNER_FILL[o]), border=True); r += 1
W(cv, r, 1, "Done", fill=PatternFill("solid", fgColor="C6EFCE"), border=True); r += 1
W(cv, r, 1, "Gate milestone", fill=FILL_GATE, border=True); r += 2
W(cv, r, 1, "Headline numbers (Medium)", bold=True); r += 1
for k, v in [("Opening capital", money(OPEN_CAP)), ("Cash out in this plan, weeks 1 to 22 (Costs tab)", "=PlanCashOut"),
             ("Funding gap in the model", F.money_k(GAP) + " needed by " + F.month_label(GAP_MONTH) + " (1.2 times the peak need of " + F.money_k(PEAK) + " in " + F.month_label(PEAK_MONTH) + ")"),
             ("Cash at end of Month 3", money(CASH_M3) + " before funding; " + money(CASH_M3_F) + " after funding"),
             ("Day 90 gate, model result", D90["result"] + " (Low " + D90L["result"] + ", High " + D90H["result"] + ")")]:
    W(cv, r, 1, k); W(cv, r, 2, v, fmt=(FMT_MONEY if str(v).startswith("=") else None), color=(GREEN if str(v).startswith("=") else "000000"), align="left"); r += 1
r += 1
W(cv, r, 1, "Glossary (every acronym used in this workbook)", bold=True); r += 1
GLOSSARY = [
    ("3PL", "third party logistics warehouse"), ("ACOS", "advertising cost of sales (ad spend divided by ad attributed sales)"), ("AOC", "active optical cable"),
    ("ASIN", "Amazon Standard Identification Number"), ("ASP", "average selling price"), ("ATC", "HDMI Authorized Test Center"), ("CBP", "US Customs and Border Protection"),
    ("CL3", "in wall cable fire rating (Class 3)"), ("COGS", "cost of goods sold"), ("CPA", "US certified public accountant"), ("DDP", "delivered duty paid"),
    ("EIN", "Employer Identification Number"), ("FBA", "Fulfillment by Amazon"), ("FCC", "Federal Communications Commission"), ("FNSKU", "Fulfillment Network Stock Keeping Unit (Amazon's barcode)"),
    ("FOB", "free on board (the price at the factory's port)"), ("FX", "foreign exchange rate (KRW per USD)"), ("HDMI", "High Definition Multimedia Interface"), ("HTS", "Harmonized Tariff Schedule"),
    ("IP", "intellectual property"), ("IRS", "Internal Revenue Service"), ("KPI", "key performance indicator"), ("LLC", "limited liability company"),
    ("ODI", "overseas direct investment report (해외직접투자 신고)"), ("PO", "purchase order"), ("PPC", "pay per click advertising"), ("Prop 65", "California Proposition 65 warning label"),
    ("QC", "quality control"), ("QR", "quick response code"), ("SKU", "stock keeping unit"), ("SSN", "Social Security Number"),
    ("TACOS", "total advertising cost of sales (ad spend divided by total sales)"), ("USPTO", "United States Patent and Trademark Office"), ("VA", "virtual assistant"), ("세무사", "Korean licensed tax accountant"),
]
GLOSSARY_LINES = []
for i in range(0, len(GLOSSARY), 5):
    GLOSSARY_LINES.append(". ".join(k + ": " + v for k, v in GLOSSARY[i:i + 5]) + ".")
for i, line in enumerate(GLOSSARY_LINES, 1): cv_row("Glossary %d of %d" % (i, len(GLOSSARY_LINES)), line, bold_k=False)
COVER_LAST = r - 1
cv.print_area = "A1:B%d" % COVER_LAST
cv.page_setup.orientation = "portrait"; cv.page_setup.paperSize = 1; cv.page_setup.fitToWidth = 1; cv.page_setup.fitToHeight = 1; cv.sheet_properties.pageSetUpPr.fitToPage = True
cv.page_margins.left = cv.page_margins.right = 0.4; cv.page_margins.top = cv.page_margins.bottom = 0.4

# ---------------- Gantt ----------------
ws = wb.create_sheet("Gantt")
HDR = ["ID", "Task", "Owner", "Support", "Start week", "End week", "Depends on", "Cost", "Cost type", "Gate or deliverable", "Basis", "Status", "Weeks", "Cost per week", "Dependency check"]
WIDTHS = [5, 58, 9, 16, 7, 7, 10, 10, 13, 52, 12, 11, 6, 9, 12]
NFIX = len(HDR)                      # fixed columns before the week grid
GRID0 = NFIX + 1                     # first week column index (P)
NWEEKS = 22
for i, w in enumerate(WIDTHS, 1): ws.column_dimensions[get_column_letter(i)].width = w
for wk in range(NWEEKS): ws.column_dimensions[get_column_letter(GRID0 + wk)].width = 3.4
HELP0 = GRID0 + NWEEKS + 1           # first helper column (dependency slots parsed from Depends on), one blank column after the grid
for k in range(MAX_DEPS):
    ws.column_dimensions[get_column_letter(HELP0 + k)].width = 6
    ws.column_dimensions[get_column_letter(HELP0 + MAX_DEPS + k)].width = 6
ws.column_dimensions[get_column_letter(GRID0 + NWEEKS)].width = 2
W(ws, 1, 1, "Gantt: 90 day launch plan, weeks 1 to 22 (Day 90 gate at week 22)", bold=True, size=12)
W(ws, 2, 1, "Grid shading is conditional formatting from Start week and End week, colored by Owner. Change Start, End, Depends on, Status or the Current week; nothing else needs editing. Acronyms are defined at first use and in the Cover glossary.", italic=True, color=GREY)
W(ws, 3, 1, "Current week", bold=True); INP(ws, 3, 3, 1, fmt="0"); name("CurrentWeek", "Gantt!$C$3")
W(ws, 3, 4, "Opening capital", bold=True); W(ws, 3, 5, OPEN_CAP, fmt=FMT_MONEY, color=GREEN); name("OpeningCapital", "Gantt!$E$3")
W(ws, 3, 7, "Total cost", bold=True)
HROW = 5
for i, h in enumerate(HDR, 1): W(ws, HROW, i, h, bold=True, fill=FILL_HDR, wrap=True, border=True, align="center")
for wk in range(1, NWEEKS + 1): W(ws, HROW, GRID0 + wk - 1, wk, bold=True, fill=FILL_HDR, border=True, align="center", fmt="0")
for k in range(MAX_DEPS):
    W(ws, HROW, HELP0 + k, "Dep %d" % (k + 1), bold=True, fill=FILL_HDR, border=True, align="center", wrap=True)
    W(ws, HROW, HELP0 + MAX_DEPS + k, "Flag %d" % (k + 1), bold=True, fill=FILL_HDR, border=True, align="center", wrap=True)
W(ws, HROW - 1, HELP0, "Helper columns: dependency IDs parsed from Depends on (formulas, no editing). Flag = that task starts after this one starts, ends after this one ends, or its ID is unknown.", color=GREY)
W(ws, HROW - 1, GRID0, "Week", bold=True); W(ws, HROW - 1, GRID0 + 9, "Month 1", bold=True, color=GREY); W(ws, HROW - 1, GRID0 + 13, "Month 2", bold=True, color=GREY); W(ws, HROW - 1, GRID0 + 17, "Month 3", bold=True, color=GREY); W(ws, HROW - 1, GRID0 + 21, "Gate", bold=True, color="9C5700")
ws.row_dimensions[HROW].height = 30
FIRST = HROW + 1
r = FIRST
STATUSES = ["Not started", "In progress", "Done", "Blocked"]
dv = DataValidation(type="list", formula1='"' + ",".join(STATUSES) + '"', allow_blank=True, showErrorMessage=True, showDropDown=False,
                    errorTitle="Status", error="Pick one of: " + ", ".join(STATUSES) + ". The Done count, the green grid and the Costs tab read the exact word Done.", errorStyle="stop")
ws.add_data_validation(dv)
dvo = DataValidation(type="list", formula1='"' + ",".join(OWNERS) + '"', allow_blank=False, showErrorMessage=True, showDropDown=False,
                     errorTitle="Owner", error="Pick one of: " + ", ".join(OWNERS) + ". The grid color and the Costs tab read the exact owner name.", errorStyle="stop")
ws.add_data_validation(dvo)
GRID_L = get_column_letter(GRID0); GRID_R = get_column_letter(GRID0 + NWEEKS - 1)
task_rows = []
for t in TASKS:
    if isinstance(t, tuple):
        W(ws, r, 1, t[1], bold=True, fill=FILL_SEC)
        for c in range(2, GRID0 + NWEEKS): ws.cell(row=r, column=c).fill = FILL_SEC
        r += 1; continue
    task_rows.append(r)
    W(ws, r, 1, t["id"], bold=True, border=True)
    W(ws, r, 2, t["task"], wrap=True, border=True)
    INP(ws, r, 3, t["owner"]); dvo.add(ws.cell(row=r, column=3))
    W(ws, r, 4, t["support"], wrap=True, border=True)
    INP(ws, r, 5, t["start"], fmt="0", align="center"); INP(ws, r, 6, t["end"], fmt="0", align="center")
    W(ws, r, 7, t["dep"], border=True, align="center")
    INP(ws, r, 8, t["cost"], fmt=FMT_MONEY2 if t["cost"] % 1 else FMT_MONEY)
    W(ws, r, 9, t["ctype"], border=True)
    W(ws, r, 10, t["deliv"], wrap=True, border=True, fill=(FILL_GATE if t["deliv"].startswith("GATE") else None))
    W(ws, r, 11, t["basis"], border=True)
    INP(ws, r, 12, "Not started"); dv.add(ws.cell(row=r, column=12))
    W(ws, r, 13, "=F%d-E%d+1" % (r, r), fmt="0", border=True, align="center")
    W(ws, r, 14, "=IF(M%d<=0,0,H%d/M%d)" % (r, r, r), fmt=FMT_MONEY2, border=True)
    # dependency check driven by the Depends on text in column G: the helper columns parse up to MAX_DEPS IDs and flag each one
    # (1 = that predecessor starts after this task starts, ends after this task ends, or its ID is not in column A; 0 = ok); column O sums the flags
    LASTROW_REF = FIRST + len(TASKS) - 1
    ids_rng = "$A$%d:$A$%d" % (FIRST, LASTROW_REF); st_rng = "$E$%d:$E$%d" % (FIRST, LASTROW_REF); en_rng = "$F$%d:$F$%d" % (FIRST, LASTROW_REF)
    for k in range(MAX_DEPS):
        slot = get_column_letter(HELP0 + k); flag = get_column_letter(HELP0 + MAX_DEPS + k)
        W(ws, r, HELP0 + k, '=TRIM(MID(SUBSTITUTE($G%d&",",",",REPT(" ",99)),%d,99))' % (r, k * 99 + 1), border=True, align="center", color=GREY)
        W(ws, r, HELP0 + MAX_DEPS + k, '=IF(%s%d="",0,IFERROR(IF(OR(INDEX(%s,MATCH(%s%d,%s,0))>$E%d,INDEX(%s,MATCH(%s%d,%s,0))>$F%d),1,0),1))'
          % (slot, r, st_rng, slot, r, ids_rng, r, en_rng, slot, r, ids_rng, r), fmt="0", border=True, align="center", color=GREY)
    flags = "%s%d:%s%d" % (get_column_letter(HELP0 + MAX_DEPS), r, get_column_letter(HELP0 + 2 * MAX_DEPS - 1), r)
    # a fifth or later ID in Depends on cannot be parsed, so it counts as one more flag
    W(ws, r, 15, '=SUM(%s)+IF(LEN($G%d)-LEN(SUBSTITUTE($G%d,",",""))>=%d,1,0)' % (flags, r, r, MAX_DEPS), fmt="0", border=True, align="center")
    for wk in range(NWEEKS):
        c = ws.cell(row=r, column=GRID0 + wk); c.border = BOX
    lines = max(len(t["task"]) // 56, len(t["deliv"]) // 50) + 1
    ws.row_dimensions[r].height = max(15, 13 * lines + 2)
    r += 1
LAST = r - 1
W(ws, 3, 8, "=SUM(H%d:H%d)" % (FIRST, LAST), fmt=FMT_MONEY, bold=True); name("GanttTotalCost", "Gantt!$H$3")
W(ws, 3, 9, "Tasks", bold=True); W(ws, 3, 10, "=COUNTA(A%d:A%d)-%d" % (FIRST, LAST, sum(1 for t in TASKS if isinstance(t, tuple))), fmt="0", bold=True, align="center")
W(ws, 3, 11, "Dependency errors (0 = ok)", bold=True, wrap=True); W(ws, 3, 13, "=SUM(O%d:O%d)" % (FIRST, LAST), fmt="0", bold=True); name("DependencyErrors", "Gantt!$M$3")
W(ws, 3, 14, "Done", bold=True); W(ws, 3, 15, '=COUNTIF(L%d:L%d,"Done")' % (FIRST, LAST), fmt="0", bold=True)
name("GanttCost", "Gantt!$H$%d:$H$%d" % (FIRST, LAST)); name("GanttOwner", "Gantt!$C$%d:$C$%d" % (FIRST, LAST)); name("GanttStart", "Gantt!$E$%d:$E$%d" % (FIRST, LAST))
name("GanttEnd", "Gantt!$F$%d:$F$%d" % (FIRST, LAST)); name("GanttType", "Gantt!$I$%d:$I$%d" % (FIRST, LAST)); name("GanttBasis", "Gantt!$K$%d:$K$%d" % (FIRST, LAST))
name("GanttCostPerWeek", "Gantt!$N$%d:$N$%d" % (FIRST, LAST)); name("GanttStatus", "Gantt!$L$%d:$L$%d" % (FIRST, LAST))
# conditional formatting on the grid
rng = "%s%d:%s%d" % (GRID_L, FIRST, GRID_R, LAST)
hdr_ref = "%s$%d" % (GRID_L, HROW)
active = "AND(%s$%d>=$E%d,%s$%d<=$F%d)" % (GRID_L, HROW, FIRST, GRID_L, HROW, FIRST)
ws.conditional_formatting.add(rng, FormulaRule(formula=['AND(%s$%d>=$E%d,%s$%d<=$F%d,$L%d="Done")' % (GRID_L, HROW, FIRST, GRID_L, HROW, FIRST, FIRST)], fill=PatternFill("solid", fgColor="C6EFCE"), stopIfTrue=True))
for o in OWNERS:
    ws.conditional_formatting.add(rng, FormulaRule(formula=['AND(%s$%d>=$E%d,%s$%d<=$F%d,$C%d="%s")' % (GRID_L, HROW, FIRST, GRID_L, HROW, FIRST, FIRST, o)], fill=PatternFill("solid", fgColor=OWNER_FILL[o]), stopIfTrue=True))
ws.conditional_formatting.add(rng, FormulaRule(formula=[active], fill=PatternFill("solid", fgColor="BFBFBF"), stopIfTrue=True))
ws.conditional_formatting.add(rng, FormulaRule(formula=["%s$%d=$C$3" % (GRID_L, HROW)], fill=PatternFill("solid", fgColor="FFF2CC")))
# a Blocked status tints the task text
ws.conditional_formatting.add("A%d:L%d" % (FIRST, LAST), FormulaRule(formula=['$L%d="Blocked"' % FIRST], font=Font(name=FONT, size=SZ, color="9C0006")))
ws.conditional_formatting.add("A%d:L%d" % (FIRST, LAST), FormulaRule(formula=['$L%d="Done"' % FIRST], font=Font(name=FONT, size=SZ, color="006100")))
ws.conditional_formatting.add("O%d:O%d" % (FIRST, LAST), FormulaRule(formula=["O%d>0" % FIRST], fill=PatternFill("solid", fgColor="FFC7CE")))
ws.freeze_panes = "C%d" % FIRST
ws.auto_filter.ref = "A%d:%s%d" % (HROW, get_column_letter(NFIX), LAST)
r = LAST + 2
W(ws, r, 1, "Notes", bold=True); r += 1
notes = [
    "Advertising (" + F.pct(m18["ads_phases"][0][2], 0) + " of revenue in Months 1 to 3, Medium) is paid from Amazon proceeds inside the P&L and is not a task cost here.",
    "Inventory rows D4, D7, F3, F4, F6, F8 and F10 equal the model's inventory cash out for Months 0 to 3 (" + money(INV_TOTAL) + "); the Costs tab checks this.",
    "The ocean freight purchase order after a PASS is sized at the gate from the reorder table and is outside this budget (critical path item 10).",
    "Korean follow up reports (remittance, securities acquisition, annual business performance) and the overseas financial account report are on 06_Compliance_Checklist.xlsx.",
]
for n_ in notes: W(ws, r, 1, n_, wrap=False, color=GREY); r += 1
NOTES_LAST = r - 1
# print: tabloid (11 x 17 in) landscape, one page wide so the task text and the week grid stay on the same page; the Notes block prints; the helper columns do not
ws.print_area = "A1:%s%d" % (GRID_R, NOTES_LAST)
ws.page_setup.orientation = "landscape"; ws.page_setup.paperSize = 3; ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0; ws.sheet_properties.pageSetUpPr.fitToPage = True
ws.print_title_rows = "%d:%d" % (HROW, HROW)
ws.page_margins.left = ws.page_margins.right = 0.25; ws.page_margins.top = ws.page_margins.bottom = 0.4

# ---------------- Costs ----------------
cs = wb.create_sheet("Costs")
for col, w in zip("ABCDEFGHIJKLMNOPQRSTUVWXYZ", [34, 12, 12, 12] + [9] * 22): cs.column_dimensions[col].width = w
W(cs, 1, 1, "Costs: cash out by owner, by cost type, by week and by owner and week (SUMIFS over the Gantt tab)", bold=True, size=12)
W(cs, 2, 1, "Every cell here is a live formula over the Gantt tab except the model reference values, which are written from build/figures.json (Medium case) for the check rows.", italic=True, color=GREY)
r = 4
W(cs, r, 1, "By owner", bold=True, fill=FILL_SEC)
for i, h in enumerate(["Tasks", "Cost", "Share"], 2): W(cs, r, i, h, bold=True, fill=FILL_HDR, border=True, align="center")
r += 1; OW0 = r
for o in OWNERS:
    W(cs, r, 1, o, border=True)
    W(cs, r, 2, '=COUNTIF(GanttOwner,A%d)' % r, fmt="0", color=GREEN, border=True)
    W(cs, r, 3, '=SUMIFS(GanttCost,GanttOwner,A%d)' % r, fmt=FMT_MONEY, color=GREEN, border=True)
    W(cs, r, 4, '=IF($C$%d=0,0,C%d/$C$%d)' % (OW0 + len(OWNERS), r, OW0 + len(OWNERS)), fmt=FMT_PCT, border=True); r += 1
OWT = r
W(cs, r, 1, "Total", bold=True, fill=FILL_TOT, border=True); W(cs, r, 2, "=SUM(B%d:B%d)" % (OW0, r - 1), fmt="0", bold=True, fill=FILL_TOT, border=True)
W(cs, r, 3, "=SUM(C%d:C%d)" % (OW0, r - 1), fmt=FMT_MONEY, bold=True, fill=FILL_TOT, border=True); W(cs, r, 4, "=IF(C%d=0,0,1)" % r, fmt=FMT_PCT, bold=True, fill=FILL_TOT, border=True); r += 1
W(cs, r, 1, "Check: owner total less Gantt total (0 = ok)", italic=True, color=GREY); W(cs, r, 3, "=C%d-GanttTotalCost" % OWT, fmt=FMT_MONEY, color=GREEN, border=True); r += 2

W(cs, r, 1, "By cost type, with the model reference", bold=True, fill=FILL_SEC)
for i, h in enumerate(["Tasks", "Cost", "Model reference (Medium)", "Check (0 = ok)", "Reference source"], 2): W(cs, r, i, h, bold=True, fill=FILL_HDR, border=True, align="center", wrap=True)
cs.row_dimensions[r].height = 30; r += 1; CT0 = r
refs = {"Launch budget": (LAUNCH_ONE, "One time launch costs, Month 1 (August model, evidence based)"),
        "Inventory": (INV_TOTAL, "Inventory cash out Months 0 to 3, Cash Flow and Working Capital tab (lot FOB plus arrival charges)"),
        "Fixed cost": (3 * (FIXED_M + SELLER_FEE), "Fixed " + money(FIXED_M) + " plus seller fee " + money(SELLER_FEE, 2) + " per month, Months 1 to 3"),
        "Insurance": (INS_YEAR, "Product liability premium per year, model input from Month " + str(INS_MONTH)),
        "No cash cost": (0, "Tasks with no cash cost in the plan window")}
for ct in COST_TYPES:
    W(cs, r, 1, ct, border=True)
    W(cs, r, 2, '=COUNTIF(GanttType,A%d)' % r, fmt="0", color=GREEN, border=True)
    W(cs, r, 3, '=SUMIFS(GanttCost,GanttType,A%d)' % r, fmt=FMT_MONEY, color=GREEN, border=True)
    W(cs, r, 4, refs[ct][0], fmt=FMT_MONEY, border=True)
    W(cs, r, 5, "=ROUND(C%d-D%d,0)" % (r, r), fmt=FMT_MONEY, border=True)
    W(cs, r, 6, refs[ct][1], wrap=True, border=True); cs.merge_cells("F%d:N%d" % (r, r)); cs.row_dimensions[r].height = 28; r += 1
CTT = r
W(cs, r, 1, "Total cash out, weeks 1 to 22", bold=True, fill=FILL_TOT, border=True); W(cs, r, 2, "=SUM(B%d:B%d)" % (CT0, r - 1), fmt="0", bold=True, fill=FILL_TOT, border=True)
W(cs, r, 3, "=SUM(C%d:C%d)" % (CT0, r - 1), fmt=FMT_MONEY, bold=True, fill=FILL_TOT, border=True); W(cs, r, 4, "=SUM(D%d:D%d)" % (CT0, r - 1), fmt=FMT_MONEY, bold=True, fill=FILL_TOT, border=True)
W(cs, r, 5, "=ROUND(C%d-D%d,0)" % (r, r), fmt=FMT_MONEY, bold=True, fill=FILL_TOT, border=True); name("PlanCashOut", "Costs!$C$%d" % r); r += 1
W(cs, r, 1, "Opening capital (Gantt tab)", border=True); W(cs, r, 3, "=OpeningCapital", fmt=FMT_MONEY, color=GREEN, border=True); r += 1
W(cs, r, 1, "Plan cash out less opening capital (before Amazon proceeds)", border=True, wrap=True); W(cs, r, 3, "=C%d-OpeningCapital" % CTT, fmt=FMT_MONEY, border=True)
W(cs, r, 6, "Amazon proceeds in Months 1 to 3 are not in this plan. The model's cash balance at the end of Month 3 is " + money(CASH_M3) + " before funding and " + money(CASH_M3_F) + " after the " + F.money_k(GAP) + " draw in " + F.month_label(GAP_MONTH) + ".", wrap=True); cs.merge_cells("F%d:N%d" % (r, r)); cs.row_dimensions[r].height = 40; r += 2

W(cs, r, 1, "By basis", bold=True, fill=FILL_SEC)
for i, h in enumerate(["Tasks", "Cost", "Share"], 2): W(cs, r, i, h, bold=True, fill=FILL_HDR, border=True, align="center")
r += 1; B0 = r
for b in (EB, ME):
    W(cs, r, 1, b, border=True); W(cs, r, 2, '=COUNTIF(GanttBasis,A%d)' % r, fmt="0", color=GREEN, border=True)
    W(cs, r, 3, '=SUMIFS(GanttCost,GanttBasis,A%d)' % r, fmt=FMT_MONEY, color=GREEN, border=True); W(cs, r, 4, "=IF(GanttTotalCost=0,0,C%d/GanttTotalCost)" % r, fmt=FMT_PCT, border=True); r += 1
r += 1

W(cs, r, 1, "By week: cost of every task spread evenly over its weeks (Cost per week column), SUMIFS on Start week and End week", bold=True, fill=FILL_SEC); r += 1
WK = r
W(cs, r, 1, "Week", bold=True, fill=FILL_HDR, border=True)
for wk in range(1, NWEEKS + 1): W(cs, r, 1 + wk, wk, bold=True, fill=FILL_HDR, border=True, align="center", fmt="0")
W(cs, r, 2 + NWEEKS, "Total", bold=True, fill=FILL_HDR, border=True, align="center"); r += 1
W(cs, r, 1, "Cash out in the week", border=True)
for wk in range(1, NWEEKS + 1):
    cl = get_column_letter(1 + wk)
    W(cs, r, 1 + wk, '=SUMIFS(GanttCostPerWeek,GanttStart,"<="&%s$%d,GanttEnd,">="&%s$%d)' % (cl, WK, cl, WK), fmt=FMT_MONEY, color=GREEN, border=True)
W(cs, r, 2 + NWEEKS, "=SUM(B%d:%s%d)" % (r, get_column_letter(1 + NWEEKS), r), fmt=FMT_MONEY, bold=True, border=True); WKROW = r; r += 1
W(cs, r, 1, "Cumulative cash out", border=True)
for wk in range(1, NWEEKS + 1):
    cl = get_column_letter(1 + wk)
    W(cs, r, 1 + wk, "=SUM($B$%d:%s%d)" % (WKROW, cl, WKROW), fmt=FMT_MONEY, border=True)
r += 1
W(cs, r, 1, "Opening capital less cumulative cash out", border=True)
for wk in range(1, NWEEKS + 1):
    cl = get_column_letter(1 + wk)
    W(cs, r, 1 + wk, "=OpeningCapital-%s%d" % (cl, r - 1), fmt=FMT_MONEY, border=True)
cs.conditional_formatting.add("B%d:%s%d" % (r, get_column_letter(1 + NWEEKS), r), FormulaRule(formula=["B%d<0" % r], fill=PatternFill("solid", fgColor="FFC7CE")))
r += 1
W(cs, r, 1, "Check: week total less Gantt total (0 = ok)", italic=True, color=GREY); W(cs, r, 2 + NWEEKS, "=%s%d-GanttTotalCost" % (get_column_letter(2 + NWEEKS), WKROW), fmt=FMT_MONEY, color=GREEN, border=True); r += 1
W(cs, r, 1, "Month in the model", border=True)
for wk in range(1, NWEEKS + 1):
    mth = next(k for k, (a, b) in MONTH_WEEKS.items() if a <= wk <= b)
    W(cs, r, 1 + wk, ("Pre" if mth == 0 else "M%d" % mth), border=True, align="center", color=GREY)
r += 2

W(cs, r, 1, "By owner and week (SUMIFS on Owner, Start week and End week)", bold=True, fill=FILL_SEC); r += 1
OWK = r
W(cs, r, 1, "Owner", bold=True, fill=FILL_HDR, border=True)
for wk in range(1, NWEEKS + 1): W(cs, r, 1 + wk, wk, bold=True, fill=FILL_HDR, border=True, align="center", fmt="0")
W(cs, r, 2 + NWEEKS, "Total", bold=True, fill=FILL_HDR, border=True, align="center"); r += 1
OWK0 = r
for o in OWNERS:
    W(cs, r, 1, o, border=True, fill=PatternFill("solid", fgColor=OWNER_FILL[o]))
    for wk in range(1, NWEEKS + 1):
        cl = get_column_letter(1 + wk)
        W(cs, r, 1 + wk, '=SUMIFS(GanttCostPerWeek,GanttOwner,$A%d,GanttStart,"<="&%s$%d,GanttEnd,">="&%s$%d)' % (r, cl, OWK, cl, OWK), fmt=FMT_MONEY, color=GREEN, border=True)
    W(cs, r, 2 + NWEEKS, "=SUM(B%d:%s%d)" % (r, get_column_letter(1 + NWEEKS), r), fmt=FMT_MONEY, bold=True, border=True); r += 1
W(cs, r, 1, "Total", bold=True, fill=FILL_TOT, border=True)
for wk in range(1, NWEEKS + 2):
    cl = get_column_letter(1 + wk)
    W(cs, r, 1 + wk, "=SUM(%s%d:%s%d)" % (cl, OWK0, cl, r - 1), fmt=FMT_MONEY, bold=True, fill=FILL_TOT, border=True)
r += 1
W(cs, r, 1, "Check: owner and week total less Gantt total (0 = ok)", italic=True, color=GREY); W(cs, r, 2 + NWEEKS, "=%s%d-GanttTotalCost" % (get_column_letter(2 + NWEEKS), r - 1), fmt=FMT_MONEY, color=GREEN, border=True); r += 2

W(cs, r, 1, "Open tasks by owner (Status not Done)", bold=True, fill=FILL_SEC); r += 1
for o in OWNERS:
    W(cs, r, 1, o, border=True); W(cs, r, 2, '=COUNTIFS(GanttOwner,A%d,GanttStatus,"<>Done")' % r, fmt="0", color=GREEN, border=True); r += 1
cs.freeze_panes = "B4"
cs.page_setup.orientation = "landscape"; cs.page_setup.fitToWidth = 1; cs.page_setup.fitToHeight = 0; cs.sheet_properties.pageSetUpPr.fitToPage = True

# ---------------- Gate ----------------
g = wb.create_sheet("Gate")
for col, w in zip("ABCDEFGHIJKLMN", [5, 54, 20, 20, 10, 10, 14, 11, 11, 11, 13, 12, 19, 40]): g.column_dimensions[col].width = w
W(g, 1, 1, "Gate: Day 90 criteria (end of Month 3, week 22) with the model values beside each threshold", bold=True, size=12)
W(g, 2, 1, "Result rule: " + D90["result_rule"] + ". REVIEW means fix and retest within 30 days. KILL means: " + KC["kill_means"] + ".", italic=True, color=GREY, wrap=True)
g.merge_cells("A2:N2"); g.row_dimensions[2].height = 28
W(g, 3, 1, "Model values are the Medium case from build/figures.json (Low and High shown for reference). Fill the yellow Actual cells at week 22; the Result column and the overall result are formulas.", italic=True, color=GREY, wrap=True)
g.merge_cells("A3:N3"); g.row_dimensions[3].height = 28
GH = ["#", "Criterion", "Pass threshold", "Kill trigger", "Pass value", "Kill value", "Test direction", "Medium model (Month 3)", "Low model", "High model", "Medium model result", "Actual at Day 90", "Actual result", "Source and basis"]
r = 5
for i, h in enumerate(GH, 1): W(g, r, i, h, bold=True, fill=FILL_HDR, wrap=True, border=True, align="center")
g.row_dimensions[r].height = 42; r += 1; G0 = r

def mres(metric):
    return "PASS" if metric["passes"] else ("KILL" if metric["kills"] else "REVIEW")
d90m = D90["metrics"]; d90l = D90L["metrics"]; d90h = D90H["metrics"]
# rows: (criterion, pass text, kill text, pass value, kill value, direction, med, low, high, med result, fmt, source, result formula builder)
def f_above(rr, c="L"): return '=IF(%s%d="","",IF(%s%d<F%d,"KILL",IF(%s%d>=E%d,"PASS","REVIEW")))' % (c, rr, c, rr, rr, c, rr, rr)
def f_below(rr, c="L"): return '=IF(%s%d="","",IF(%s%d>F%d,"KILL",IF(%s%d<=E%d,"PASS","REVIEW")))' % (c, rr, c, rr, rr, c, rr, rr)
GATE_ROWS = [
    ("Unit session conversion: number of the 5 launch SKUs at or above " + F.pct(GT["d90_conv_pass"]["value"], 0), "at or above 3 of 5 SKUs", "none (see the hero row)", 3, None, "at or above", "not modeled", "not modeled", "not modeled", "n/a", "0",
     "Observed in Seller Central (Business Reports). Evidence based threshold", lambda rr: '=IF(L%d="","",IF(L%d>=E%d,"PASS","REVIEW"))' % (rr, rr, rr)),
    ("Hero SKU (8K 10M) unit session conversion", "at or above " + F.pct(GT["d90_conv_pass"]["value"], 0), "below " + F.pct(GT["d90_conv_kill"]["value"], 0), GT["d90_conv_pass"]["value"], GT["d90_conv_kill"]["value"], "at or above", "not modeled", "not modeled", "not modeled", "n/a", FMT_PCT,
     "Observed in Seller Central. Evidence based threshold", f_above),
    ("TACOS in Month 3 (ad spend divided by total revenue)", d90m["tacos"]["pass_threshold"], d90m["tacos"]["kill_trigger"], GT["d90_tacos_pass"]["value"], GT["d90_tacos_kill"]["value"], "at or below", d90m["tacos"]["value"], d90l["tacos"]["value"], d90h["tacos"]["value"], mres(d90m["tacos"]), FMT_PCT,
     "Model: Monthly PL Y1 Y2, Month 3 ad spend phase rate. Evidence based (August model)", f_below),
    ("TACOS falling month over month: Month 2 actual (enter here) must be above the Month 3 actual in the row above", "Month 3 below Month 2", "none", None, None, "falling", "flat (phase rate %s in Months 1 to 3)" % F.pct(d90m["tacos"]["value"], 0), "flat", "flat", "n/a", FMT_PCT,
     "Observed from the advertising reports. The model holds one phase rate for Months 1 to 3, so this test applies to actuals only", None),
    ("Return rate in Month 3", d90m["ret"]["pass_threshold"], d90m["ret"]["kill_trigger"], GT["d90_ret_pass"]["value"], GT["d90_ret_kill"]["value"], "at or below", d90m["ret"]["value"], d90l["ret"]["value"], d90h["ret"]["value"], mres(d90m["ret"]), FMT_PCT,
     "Model: returns % assumption. Evidence based (August model)", f_below),
    ("Contribution per unit before advertising on at least 3 of 5 SKUs (hero basis in the model)", d90m["cpu"]["pass_threshold"], d90m["cpu"]["kill_trigger"] + " (hero)", GT["d90_cpu_pass"]["value"], GT["d90_cpu_kill"]["value"], "at or above", d90m["cpu"]["value"], d90l["cpu"]["value"], d90h["cpu"]["value"], mres(d90m["cpu"]), FMT_MONEY2,
     "Model: ASP less landed COGS, referral, FBA, returns, storage (Monthly PL Y1 Y2). Enter the hero actual", f_above),
    ("Cash runway after funding the next purchase order, months", d90m["runway"]["pass_threshold"] + " months", d90m["runway"]["kill_trigger"] + " months", GT["d90_run_pass"]["value"], GT["d90_run_kill"]["value"], "at or above", d90m["runway"]["value"], d90l["runway"]["value"], d90h["runway"]["value"], mres(d90m["runway"]), FMT_NUM1,
     "Model: cash after the funding draw divided by the Months 4 to 6 burn (Cash Flow and Working Capital). Before funding the Medium runway is " + F.num(M["cash"]["runway_months_day90_unfunded"], 1) + " months", f_above),
    ("Suppressed listing or intellectual property complaint: days the longest one has been open (0 = none)", "none open", "any open more than " + str(GT["d90_ip_days"]["value"]) + " days", 0, GT["d90_ip_days"]["value"], "at or below", "not modeled", "not modeled", "not modeled", "n/a", "0",
     "Observed in Account Health. Evidence based threshold", lambda rr: '=IF(L%d="","",IF(L%d>F%d,"KILL",IF(L%d<=E%d,"PASS","REVIEW")))' % (rr, rr, rr, rr, rr)),
]
for i, row in enumerate(GATE_ROWS, 1):
    crit, pt, kt, pv, kv, dirn, med, low, high, medres, fm, src, fb = row
    W(g, r, 1, i, border=True, align="center", fmt="0"); W(g, r, 2, crit, wrap=True, border=True)
    W(g, r, 3, pt, wrap=True, border=True); W(g, r, 4, kt, wrap=True, border=True)
    W(g, r, 5, pv if pv is not None else "n/a", fmt=fm, border=True, align="center"); W(g, r, 6, kv if kv is not None else "n/a", fmt=fm, border=True, align="center")
    W(g, r, 7, dirn, border=True, align="center")
    for c, v in ((8, med), (9, low), (10, high)):
        W(g, r, c, v, fmt=fm, border=True, align="center", wrap=True, color=(GREEN if isinstance(v, (int, float)) else "000000"))
    # Medium model result: the same formula as the Actual result, reading the model value in H, so it follows a threshold edit
    if isinstance(med, (int, float)) and fb is not None:
        W(g, r, 11, fb(r, "H"), border=True, align="center", bold=True)
        assert medres in ("PASS", "REVIEW", "KILL")
    else:
        W(g, r, 11, medres, border=True, align="center", bold=True)
    INP(g, r, 12, None, fmt=fm, align="center")
    if fb is None:
        W(g, r, 13, '=IF(OR(L%d="",L%d=""),"",IF(L%d<L%d,"PASS","REVIEW"))' % (r - 1, r, r - 1, r), border=True, align="center", bold=True)
    else:
        W(g, r, 13, fb(r), border=True, align="center", bold=True)
    W(g, r, 14, src, wrap=True, border=True)
    g.row_dimensions[r].height = 44; r += 1
G1 = r - 1
W(g, r, 2, "Overall Day 90 result", bold=True, fill=FILL_TOT, border=True)
W(g, r, 11, '=IF(COUNTIF(K%d:K%d,"KILL")>0,"KILL",IF(COUNTIF(K%d:K%d,"REVIEW")>0,"REVIEW","PASS"))' % (G0, G1, G0, G1), bold=True, fill=FILL_TOT, border=True, align="center")
name("GateModelResult", "Gate!$K$%d" % r)
W(g, r, 12, '=COUNTBLANK(L%d:L%d)' % (G0, G1), fmt="0", fill=FILL_TOT, border=True, align="center")
W(g, r, 13, '=IF(COUNTBLANK(L%d:L%d)>0,"Enter every actual",IF(COUNTIF(M%d:M%d,"KILL")>0,"KILL",IF(COUNTIF(M%d:M%d,"PASS")=ROWS(M%d:M%d),"PASS","REVIEW")))' % (G0, G1, G0, G1, G0, G1, G0, G1), bold=True, fill=FILL_TOT, border=True, align="center")
W(g, r, 14, "Column L counts the empty Actual cells. Low model result " + D90L["result"] + ", High " + D90H["result"] + ".", wrap=True, fill=FILL_TOT, border=True)
name("GateResult", "Gate!$M$%d" % r); g.row_dimensions[r].height = 30; r += 1
rng = "K%d:M%d" % (G0, r - 1)
g.conditional_formatting.add(rng, FormulaRule(formula=['K%d="PASS"' % G0], fill=PatternFill("solid", fgColor="C6EFCE"), font=Font(name=FONT, size=SZ, bold=True, color="006100")))
g.conditional_formatting.add(rng, FormulaRule(formula=['K%d="KILL"' % G0], fill=PatternFill("solid", fgColor="FFC7CE"), font=Font(name=FONT, size=SZ, bold=True, color="9C0006")))
g.conditional_formatting.add(rng, FormulaRule(formula=['K%d="REVIEW"' % G0], fill=PatternFill("solid", fgColor="FFEB9C"), font=Font(name=FONT, size=SZ, bold=True, color="9C5700")))
r += 1

W(g, r, 2, "Month 3 model context (Medium, from build/figures.json)", bold=True, fill=FILL_SEC); W(g, r, 3, "Medium", bold=True, fill=FILL_HDR, border=True, align="center"); W(g, r, 4, "Low", bold=True, fill=FILL_HDR, border=True, align="center"); W(g, r, 5, "High", bold=True, fill=FILL_HDR, border=True, align="center"); r += 1
ctx = [
    ("Units sold in Month 3", "units", "0"), ("Revenue in Month 3", "revenue", FMT_MONEY), ("Net in Month 3 (after fixed costs and the seller fee)", "net", FMT_MONEY),
    ("Cumulative net, Months 1 to 3 (includes the " + money(LAUNCH_ONE) + " launch costs)", "cum_net", FMT_MONEY), ("Contribution per unit after advertising, Month 3", "contribution_after_ads", FMT_MONEY2),
    ("Cash balance at end of Month 3 before funding", "cash_balance", FMT_MONEY), ("Cash balance at end of Month 3 after funding", "cash_balance_after_funding", FMT_MONEY),
]
for label, key, fm in ctx:
    W(g, r, 2, label, border=True, wrap=True)
    for c, S in ((3, M), (4, L), (5, H)): W(g, r, c, S["monthly"][2][key], fmt=fm, border=True, align="center", color=GREEN)
    r += 1
for label, vals, fm in [("Funding gap in the model (1.2 times the peak cash need)", [S["cash"]["funding_gap"] for S in (M, L, H)], FMT_MONEY),
                        ("Funding needed by month", ["Month %d" % S["cash"]["funding_needed_by_month"] for S in (M, L, H)], None),
                        ("Peak cash need and month", [F.money_k(S["cash"]["peak_need"]) + " in Month " + str(S["cash"]["peak_need_month"]) for S in (M, L, H)], None),
                        ("18 month revenue and net (Amazon US)", [F.money_m(S["m18"]["revenue"]) + " and " + F.money_k(S["m18"]["net"]) for S in (M, L, H)], None),
                        ("Cumulative breakeven month", [F.month_label(S["m18"]["breakeven_month"]) for S in (M, L, H)], None)]:
    W(g, r, 2, label, border=True, wrap=True)
    for c, v in zip((3, 4, 5), vals): W(g, r, c, v, fmt=fm, border=True, align="center", color=GREEN, wrap=True)
    r += 1
r += 1
W(g, r, 2, "Reference: Month 12 gate (outside this 90 day plan; the same PASS, REVIEW, KILL rule)", bold=True, fill=FILL_SEC); r += 1
for i, h in enumerate(["#", "Criterion", "Pass threshold", "Kill trigger", "Pass value", "Kill value", "Test direction", "Medium model (Month 12)", "Low model", "High model", "Medium model result"], 1): W(g, r, i, h, bold=True, fill=FILL_HDR, wrap=True, border=True, align="center")
g.row_dimensions[r].height = 42; r += 1; M0 = r
m12m = M12["metrics"]; m12l = M12L["metrics"]; m12h = M12H["metrics"]
M12_ROWS = [("rev", GT["m12_rev_pass"]["value"], GT["m12_rev_kill"]["value"], "at or above", FMT_MONEY), ("cum", GT["m12_cum_pass"]["value"], GT["m12_cum_kill"]["value"], "at or above", FMT_MONEY),
            ("tacos", GT["m12_tacos_pass"]["value"], GT["m12_tacos_kill"]["value"], "at or below", FMT_PCT), ("ret", GT["m12_ret_pass"]["value"], GT["m12_ret_kill"]["value"], "at or below", FMT_PCT),
            ("opcf", 0, None, "at or above", FMT_MONEY)]
for i, (k, pv, kv, dirn, fm) in enumerate(M12_ROWS, 1):
    W(g, r, 1, i, border=True, align="center", fmt="0"); W(g, r, 2, m12m[k]["label"], wrap=True, border=True)
    W(g, r, 3, m12m[k]["pass_threshold"], wrap=True, border=True); W(g, r, 4, m12m[k]["kill_trigger"], wrap=True, border=True)
    W(g, r, 5, pv, fmt=fm, border=True, align="center"); W(g, r, 6, kv if kv is not None else "gap above " + money(GT["m12_gap_kill"]["value"]), fmt=fm, border=True, align="center", wrap=True)
    W(g, r, 7, dirn, border=True, align="center")
    for c, S in ((8, m12m), (9, m12l), (10, m12h)): W(g, r, c, S[k]["value"], fmt=fm, border=True, align="center", color=GREEN)
    if kv is None:   # operating cash flow: the kill trigger needs three negative months and a funding gap, which this one cell cannot test
        W(g, r, 11, '=IF(H%d>=E%d,"PASS","REVIEW")' % (r, r), border=True, align="center", bold=True)
    else:
        W(g, r, 11, (f_above if dirn == "at or above" else f_below)(r, "H"), border=True, align="center", bold=True)
    assert mres(m12m[k]) in ("PASS", "REVIEW", "KILL")
    g.row_dimensions[r].height = 40; r += 1
M1 = r - 1
W(g, r, 2, "Overall Month 12 result", bold=True, fill=FILL_TOT, border=True)
W(g, r, 11, '=IF(COUNTIF(K%d:K%d,"KILL")>0,"KILL",IF(COUNTIF(K%d:K%d,"REVIEW")>0,"REVIEW","PASS"))' % (M0, M1, M0, M1), bold=True, fill=FILL_TOT, border=True, align="center")
W(g, r, 12, "Low " + M12L["result"] + ", High " + M12H["result"], wrap=True); r += 1
rng = "K%d:K%d" % (M0, r - 1)
g.conditional_formatting.add(rng, FormulaRule(formula=['K%d="PASS"' % M0], fill=PatternFill("solid", fgColor="C6EFCE"), font=Font(name=FONT, size=SZ, bold=True, color="006100")))
g.conditional_formatting.add(rng, FormulaRule(formula=['K%d="KILL"' % M0], fill=PatternFill("solid", fgColor="FFC7CE"), font=Font(name=FONT, size=SZ, bold=True, color="9C0006")))
r += 1
W(g, r, 2, "Glossary (the same list is on the Cover tab)", bold=True); r += 1
for line in GLOSSARY_LINES:
    W(g, r, 2, line, color=GREY, wrap=True); g.merge_cells("B%d:N%d" % (r, r)); g.row_dimensions[r].height = 28; r += 1
GATE_LAST = r - 1
g.print_area = "A1:N%d" % GATE_LAST
g.freeze_panes = "C6"
g.page_setup.orientation = "landscape"; g.page_setup.fitToWidth = 1; g.page_setup.fitToHeight = 1; g.sheet_properties.pageSetUpPr.fitToPage = True
g.page_margins.left = g.page_margins.right = 0.3; g.page_margins.top = g.page_margins.bottom = 0.3

os.makedirs(OUT, exist_ok=True)
wb.save(PATH)
print(json.dumps({"file": PATH, "tasks": N_TASKS, "sections": sum(1 for t in TASKS if isinstance(t, tuple)), "gantt_first_row": FIRST, "gantt_last_row": LAST,
                  "launch_budget_sum": LAUNCH_SUM, "inventory_sum": round(INV_TOTAL, 2), "fixed_sum": round(3 * (FIXED_M + SELLER_FEE), 2), "insurance": INS_YEAR,
                  "total_cost": round(sum(t["cost"] for t in TASKS if isinstance(t, dict)), 2), "lot_units": [round(a, 1) for a in ARR], "lot_fob": [round(x, 2) for x in LOT_FOB], "lot_fees": [round(x, 2) for x in LOT_FEES],
                  "evidence_based": sum(1 for t in TASKS if isinstance(t, dict) and t["basis"] == EB), "management_estimate": sum(1 for t in TASKS if isinstance(t, dict) and t["basis"] == ME)}))

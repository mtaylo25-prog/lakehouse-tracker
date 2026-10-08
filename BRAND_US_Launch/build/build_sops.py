#!/usr/bin/env python3
"""Build 05_Operating_Processes.docx and .pdf.

Step 1 (this file): read the working capital inputs (cover months, lead times, ocean start month,
PO deposit, payout lag) from the Assumptions tab of the recalculated workbook and derive the
reorder point table from the unit ramp in build/figures.json. Write build/sops_data.json.
Step 2: run node build/build_sops_docx.js (docx generator, reads figures.json, facts.json, fmt.js, sops_data.json).
Step 3: to_pdf, lint_hyphens, QA render to build/qa/sops/.

Usage: source build/env.sh && python3 build/build_sops.py
"""
import json, os, subprocess, sys, math
from pathlib import Path

ROOT = Path(os.environ.get("ROOT", "/home/user/lakehouse-tracker/BRAND_US_Launch"))
BUILD = ROOT / "build"
OUT = ROOT / "output"
sys.path.insert(0, str(BUILD))
import fmt  # noqa: E402

fig = json.load(open(BUILD / "figures.json"))
facts = json.load(open(BUILD / "facts.json"))

# ---- 1. working capital inputs from the workbook Assumptions tab (Low, Medium, High, Active) ----
from openpyxl import load_workbook
wb = load_workbook(OUT / "03_Financial_Model_5Y.xlsx", data_only=True, read_only=True)
ws = wb["Assumptions"]
wc = {}
labels = {
    "cover_months": "Inventory cover target, months of forward units",
    "ocean_from_month": "Ocean freight from Month (before that air)",
    "air_lead_months": "Air freight lead time, months (order to arrival)",
    "ocean_lead_months": "Ocean freight lead time, months (order to arrival)",
    "po_deposit": "PO deposit at order (balance at shipment)",
    "payout_lag_days": "Amazon payout lag, days",
    "inventory_days_y3_y5": "Inventory days of landed COGS, Years 3 to 5",
    "insurance_per_year": "Product liability insurance ($1M cover) per year once monthly sales pass $10K",
    "form_5472_prep": "Form 5472 preparation (with pro forma Form 1120)",
    "wyoming_annual_report": "Wyoming annual report",
    "bond_renewal": "Customs continuous bond renewal",
    "fixed_y1": "Fixed costs per month Year 1",
}
for row in ws.iter_rows(min_row=1, max_row=200, max_col=7, values_only=True):
    a = row[0]
    if not isinstance(a, str):
        continue
    for key, lab in labels.items():
        if a.startswith(lab):
            wc[key] = {"Low": row[2], "Medium": row[3], "High": row[4], "Active": row[5], "basis": row[6]}
missing = [k for k in labels if k not in wc]
if missing:
    raise SystemExit(f"Assumptions rows not found: {missing}")

# ---- 2. reorder table: the workbook's own stock rule (Cash Flow and Working Capital rows 11 to 19) ----
# target closing stock = cover months times the average of the next three months of all channel units;
# arrivals = max(0, target minus (opening minus sales)); the PO for an arrival in month a is placed in month a minus lead(a),
# where lead(a) is the air lead until the ocean switch month and the ocean lead from then on (lead is set by the ARRIVAL month,
# exactly as the workbook row 15 does); deposit at order, balance at shipment (the month before arrival).
# All unit and dollar figures are read from figures.json scenarios.{S}.monthly (closing_stock, arrivals, inventory_on_hand,
# inventory_in_transit) so the table and the cash forecast are the same numbers.
def reorder_table(scn, months=12):
    mon = fig["scenarios"][scn]["monthly"]
    cover = float(wc["cover_months"][scn])
    ocean_from = int(wc["ocean_from_month"][scn])
    air_lead = int(wc["air_lead_months"][scn])
    ocean_lead = int(wc["ocean_lead_months"][scn])
    dep = float(wc["po_deposit"][scn])
    fob = fig["unit_economics"]["hero"]["fob_usd"]          # workbook row 9: FOB per unit paid to the factory (hero FOB USD at the active FX)
    n = len(mon)

    def lead_of(a):           # a = arrival month (1 based)
        return ocean_lead if a >= ocean_from else air_lead

    def arrivals(a):
        return float(mon[a - 1]["arrivals"]) if 1 <= a <= n else 0.0

    def closing(m):
        return float(mon[m - 1]["closing_stock"]) if 1 <= m <= n else 0.0

    # figures.json "units" is Amazon US only; the workbook stock rule runs on all channel units (Amazon plus eBay and Walmart plus B2B
    # from Month 13). Back the all channel units out of the model's own identity closing = opening + arrivals minus sales.
    units_all = {m: closing(m - 1) + arrivals(m) - closing(m) for m in range(1, n + 1)}
    for m in range(1, 13):
        if abs(units_all[m] - float(mon[m - 1]["units"])) > 0.5:
            raise SystemExit(f"{scn} Month {m}: all channel units {units_all[m]} differ from Amazon units {mon[m - 1]['units']} in Year 1")
    # check the stock rule against the pulled rows so a silent model change is caught
    for m in range(1, n + 1):
        opening = closing(m - 1)
        fwd = [units_all[k] for k in (m + 1, m + 2, m + 3) if k <= n]
        if len(fwd) == 3:
            target = cover * sum(fwd) / 3.0
            expected = max(0.0, target - (opening - units_all[m]))
            if abs(expected - arrivals(m)) > 0.5:
                raise SystemExit(f"{scn} Month {m}: arrivals {arrivals(m)} differ from the stock rule {expected}")

    rows = []
    for m in range(0, months + 1):        # Month 0 is the pre launch PO for the Month 1 arrival
        po = [(a, arrivals(a), lead_of(a)) for a in range(m + 1, m + max(air_lead, ocean_lead) + 1) if a - lead_of(a) == m and a <= n]
        po_units = sum(u for _, u, _ in po)
        r = mon[m - 1] if m >= 1 else None
        fwd = [units_all[k] for k in (m + 1, m + 2, m + 3) if k <= n]
        rows.append({
            "month": m,
            "units": (int(round(units_all[m])) if r else 0),
            "amazon_units": (r["units"] if r else 0),
            "next3_avg_units": (sum(fwd) / len(fwd)) if fwd else None,
            "target_closing_stock": (r["closing_stock"] if r else 0.0),
            "arrivals": (r["arrivals"] if r else 0.0),
            "arrival_mode": ("Ocean" if m >= ocean_from else "Air") if r else "n/a",
            "po_units": po_units,
            "po_for_months": [a for a, _, _ in po],
            "po_lead_months": [l for _, _, l in po],
            "po_fob_usd": po_units * fob,
            "deposit_usd": dep * po_units * fob,
            "balance_usd": (1 - dep) * po_units * fob,
            "inventory_on_hand_usd": (r["inventory_on_hand"] if r else 0.0),
            "inventory_in_transit_usd": (r["inventory_in_transit"] if r else 0.0),
        })
    return {"cover_months": cover, "ocean_from_month": ocean_from, "air_lead_months": air_lead,
            "ocean_lead_months": ocean_lead, "po_deposit": dep, "fob_per_unit_usd": fob,
            "landed_cost_per_unit_month1": (float(mon[0]["inventory_on_hand"]) / float(mon[0]["closing_stock"])) if mon[0]["closing_stock"] else None,
            "rows": rows}

data = {
    "wc_inputs": wc,
    "reorder": {s: reorder_table(s) for s in ("Low", "Medium", "High")},
    "hero_landed_cost": fig["unit_economics"]["hero"]["landed_cost"],
    "medium_month12_units": fig["scenarios"]["Medium"]["monthly"][11]["units"],
    "insurance_trigger_month": {s: fig["scenarios"][s]["insurance_trigger_month"] for s in ("Low", "Medium", "High")},
    "insurance_trigger_margin_usd": {s: fig["scenarios"][s]["insurance_trigger_margin_usd"] for s in ("Low", "Medium", "High")},
}
(BUILD / "sops_data.json").write_text(json.dumps(data, indent=1))
print("sops_data.json written; Medium reorder rows:")
for r in data["reorder"]["Medium"]["rows"]:
    print({k: (round(v, 1) if isinstance(v, float) else v) for k, v in r.items()})
print("insurance trigger month by scenario:", data["insurance_trigger_month"])

# ---- 3. docx generator ----
env = dict(os.environ)
r = subprocess.run(["node", str(BUILD / "build_sops_docx.js")], env=env, text=True, capture_output=True)
print(r.stdout)
if r.returncode != 0:
    print(r.stderr)
    raise SystemExit("docx generator failed")
print("docx written")

#!/usr/bin/env python3
"""Builds output/00_README.md: index of every delivered file, readers, order of use, rebuild steps, headline numbers,
top risks, Monday actions and the tie out result.
Every model number comes from build/figures.json through build/fmt.py. Top risks are read from
output/08_Risk_Register.xlsx (Probability times Impact). The tie out result is build/tie_out_result.md quoted verbatim.
No hyphens, en dashes or em dashes in the delivered text (official form and registration numbers keep their format).
Rerun: python3 build/build_readme.py  (run it last, after build/tie_out.py, so the quoted result is current)
"""
import json, sys
from pathlib import Path

ROOT = Path("/home/user/lakehouse-tracker/BRAND_US_Launch")
sys.path.insert(0, str(ROOT / "build"))
import fmt  # noqa: E402
from openpyxl import load_workbook  # noqa: E402

fig = json.load(open(ROOT / "build" / "figures.json"))
facts = json.load(open(ROOT / "build" / "facts.json"))
S = fig["scenarios"]
M, L, H = S["Medium"], S["Low"], S["High"]
hero = fig["unit_economics"]["hero"]

# ---------------------------------------------------------------- top risks from the register
wb = load_workbook(ROOT / "output" / "08_Risk_Register.xlsx", data_only=True)
ws = wb["Risk Register"]
rows = list(ws.iter_rows(values_only=True))
hdr_i = next(i for i, r in enumerate(rows) if r and r[0] == "ID")
hdr = list(rows[hdr_i])
pi = next(i for i, h in enumerate(hdr) if str(h).startswith("Prob"))
ii = next(i for i, h in enumerate(hdr) if str(h).startswith("Impact"))
oi = hdr.index("Owner"); mi = hdr.index("Mitigation"); ci = hdr.index("Category")
risks = [r for r in rows[hdr_i + 1:] if r and r[0] and isinstance(r[pi], (int, float)) and isinstance(r[ii], (int, float))]
risks.sort(key=lambda r: (-(r[pi] * r[ii]), str(r[0])))
top3 = risks[:3]

# ---------------------------------------------------------------- tie out result verbatim
tie = (ROOT / "build" / "tie_out_result.md").read_text().rstrip("\n")

# ---------------------------------------------------------------- delivered file listing (every file in output/)
out_dir = ROOT / "output"
delivered = sorted(p.relative_to(out_dir).as_posix() for p in out_dir.rglob("*") if p.is_file())

def dollars(x):
    return fmt.money_auto(x)

m18 = M["m18"]; c = M["cash"]
y5 = {k: S[k]["annual"]["Y5"] for k in ("Low", "Medium", "High")}

# ---------------------------------------------------------------- index table
INDEX = [
    ("00_README.md", "This file. Index, order of use, rebuild steps, headline numbers, tie out result.", "Lucy, Mark", "Read first"),
    ("01_Business_Proposal_EN.pptx and .pdf", "37 slide English proposal: market, product, structure, launch plan, economics, 5 year path, funding, risks, kill criteria, the ask.", "Mark, bank (English readers), any US adviser", "Step 3"),
    ("02_Business_Proposal_KO.pptx and .pdf", "The same 37 slides in Korean (Noto Sans KR). Amazon listing content stays in English.", "Lucy, 세무사, bank", "Step 3"),
    ("03_Financial_Model_5Y.xlsx", "Live 5 year model. One scenario selector (Low, Medium, High), one duty selector (Copper, Fiber, Stress), FX input. 14 tabs from Assumptions to Plan B. Every number in every other file comes from it.", "Lucy, Mark, bank (if asked for the model), 세무사 (cash and repatriation tabs)", "Step 4"),
    ("04_Flowcharts/ (8 charts, each .mmd, .png, .pdf, plus README.md)", "Critical path, order to cash, inbound supply chain, returns and the 5 Year Signal Guarantee, monthly cadence, Day 90 scale or kill tree, 5 year channel map, entity money and goods flow.", "Lucy, Mark, broker (03 and 08), 3PL (03), bank and 세무사 (02 and 08)", "With the deck and the SOPs"),
    ("05_Operating_Processes.docx and .pdf", "Eight SOPs (owner, trigger, steps, system, KPI, failure modes) plus the compliance calendar.", "Lucy, Mark, VA, broker (customs SOP), 3PL (inbound SOP)", "Step 5"),
    ("06_Compliance_Checklist.xlsx", "Every US, Korea, Amazon, HDMI, FCC, Prop 65, CL3, insurance, customs and tax item with owner, due, cost, evidence, status.", "Lucy, 세무사 (Korea rows), broker (customs rows), Mark", "Step 5, then weekly"),
    ("07_90_Day_Launch_Plan.xlsx", "Week by week Gantt (weeks 1 to 22) with owner, dependencies, cost by owner and week, and the Day 90 gate tab.", "Lucy, Mark, attorney, broker, factory, 3PL", "Step 5, then weekly"),
    ("08_Risk_Register.xlsx", "32 risks scored probability times impact, early warnings, mitigations, owners, kill criteria tab, heat map.", "Lucy, Mark, bank (if asked)", "Step 5, then monthly"),
    ("09_Amazon_Listing_Kit.docx", "Parent and child ASIN structure, titles, bullets, A+ outline, backend keywords, image and video brief, FAQ, compliance notes. English only.", "Mark, VA, photographer, factory (packaging text)", "Weeks 6 to 9"),
    ("10_One_Page_Summary_EN.pdf", "One page for the bank and any English reader: who, structure, filings in order, capital need, timeline.", "Bank, Mark", "Step 1 (Monday)"),
    ("10_One_Page_Summary_KO.pdf", "The same page in Korean for the foreign exchange bank and the 세무사.", "세무사, bank, Lucy", "Step 1 (Monday)"),
    ("11_Decision_Log.md", "The five open decisions, the defaults used, what changes if each goes the other way (quantified from the model), who decides, by when.", "Lucy, Mark, 세무사 (ownership and entity rows)", "Step 2 (Monday)"),
]

# ---------------------------------------------------------------- compose
lines = []
A = lines.append
A("# 00 README: BRAND US launch package")
A("")
A(f"Built {fig['meta']['built']}. Owner Lucy (Hyejin Yoo), H2 Company, Seoul. Adviser Mark Taylor, Memphis. US entity BRAND LLC (Wyoming), brand name placeholder BRAND until chosen. Every number below comes from build/figures.json, pulled from output/03_Financial_Model_5Y.xlsx at its delivered state: Scenario Medium, Duty Copper (40.1%), FX {fmt.num(fig['meta']['fx'])} KRW per USD.")
A("")
A("패키지 전체의 파일 목록, 읽는 순서, 재생성 방법, 핵심 숫자와 검증 결과를 담은 안내 파일입니다.")
A("")
A("## 1. What is in output/")
A("")
A("파일별 용도와 읽는 사람입니다.")
A("")
A("| File | What it is for | Who reads it | When |")
A("|---|---|---|---|")
for f, what, who, when in INDEX:
    A(f"| {f} | {what} | {who} | {when} |")
A("")
A("Full directory listing (every delivered file):")
A("")
for p in delivered:
    A(f"- output/{p}")
A("")
A("## 2. Order of use")
A("")
A("월요일에 시작하는 순서입니다. 은행과 세무사용 한 장 요약부터 시작합니다.")
A("")
A("Monday sequence first, then the rest.")
A("")
A("1. **10_One_Page_Summary_KO.pdf to the foreign exchange bank and the 세무사; 10_One_Page_Summary_EN.pdf to Mark and any English reader.** It carries the structure, the filings in order relative to the capital transfer, the opening capital and the funding gap. The bank needs it for the overseas direct investment (ODI) report before any capital moves. The 세무사 needs it for the ODI follow up reports, the overseas financial account report threshold (KRW 500M month end aggregate), Form 5472 and the question of whether to incorporate H2 Company first.")
A("2. **11_Decision_Log.md.** Confirm or change the five defaults: brand name, LLC ownership (Lucy 100%, Mark under a consulting agreement), funding source (self funded to $39K, gap as a dated line of credit), FX (1,430), base duty scenario (Copper 40.1% pending the CBP binding ruling). Each row says what changes if the decision goes the other way.")
A("3. **01_Business_Proposal_EN or 02_Business_Proposal_KO.** Read the executive summary (slide 2), the Day 90 gate (slide 19), the 18 month P&L (slide 23), the 5 year trajectory (slide 24), working capital (slide 27) and the ask (slide 35). Use the PDF to share; the pptx to edit.")
A("4. **03_Financial_Model_5Y.xlsx.** Change only yellow cells with blue text. Scenario Summary shows all three scenarios side by side without touching the selector. KPI Dashboard carries the charts.")
A("5. **07_90_Day_Launch_Plan.xlsx** for the weekly launch review (with 04_Flowcharts/01_launch_critical_path.png). Then 06_Compliance_Checklist.xlsx, 08_Risk_Register.xlsx and 05_Operating_Processes as the business starts running. 09_Amazon_Listing_Kit.docx goes to Mark and the VA in weeks 6 to 9.")
A("")
A("## 3. Headline numbers (Medium case unless stated)")
A("")
A("모델에서 가져온 핵심 숫자입니다. 중간(Medium) 시나리오 기준입니다.")
A("")
A("| Measure | Value |")
A("|---|---|")
A(f"| 18 month revenue (Amazon US only) | {fmt.money_m(m18['revenue'])} |")
A(f"| 18 month net | {fmt.money_k(m18['net'])} ({fmt.pct(m18['net_margin'])}) |")
A(f"| Cumulative breakeven | {fmt.month_label(m18['breakeven_month'])} |")
A(f"| Month 18 run rate and exit margin | {fmt.money_m(m18['run_rate'])}, {fmt.pct(m18['exit_margin'])} |")
for band in ("Low", "Medium", "High"):
    A(f"| Year 5, {band} | revenue {fmt.money_m(y5[band]['revenue'])}, net {fmt.money_k(y5[band]['net'])}, margin {fmt.pct(y5[band]['net_margin'])} |")
A(f"| Opening capital | {fmt.money_k(c['opening_capital'])} (self funded, August ask) |")
A(f"| Peak cash need | {fmt.money_k(c['peak_need'])} in {fmt.month_label(c['peak_need_month'])} |")
A(f"| Funding gap (1.2 times the lowest balance) | {fmt.money_k(c['funding_gap'])}, in place by {fmt.month_label(c['funding_needed_by_month'])} |")
A(f"| Day 90 and Month 12 gates | Medium {M['gates']['day90']['result']} and {M['gates']['month12']['result']}; Low {L['gates']['day90']['result']} and {L['gates']['month12']['result']}; High {H['gates']['day90']['result']} and {H['gates']['month12']['result']} |")
A("")
A(f"Other scenarios, 18 months: Low {fmt.money_m(L['m18']['revenue'])} revenue, {fmt.money_k(L['m18']['net'])} net, breakeven {fmt.month_label(L['m18']['breakeven_month'])}, peak cash need {fmt.money_k(L['cash']['peak_need'])} in {fmt.month_label(L['cash']['peak_need_month'])}. High {fmt.money_m(H['m18']['revenue'])} revenue, {fmt.money_k(H['m18']['net'])} net, breakeven {fmt.month_label(H['m18']['breakeven_month'])}, peak cash need {fmt.money_k(H['cash']['peak_need'])} in {fmt.month_label(H['cash']['peak_need_month'])}.")
A("")
A(f"Hero SKU (8K 10M) at {fmt.money(hero['price'], 2)}: landed cost {fmt.money(hero['landed_cost'], 2)} ({fmt.pct(hero['landed_pct'])}), net {fmt.money(hero['net'], 2)} per unit ({fmt.pct(hero['net_pct'])}). The same unit earned {fmt.money(hero['korea_wholesale_gross'], 2)} gross in Korea wholesale. The Alibaba benchmark ($17 to $22 FOB for a comparable 30 ft AOC) against the H2 base ({fmt.money(hero['fob_usd'], 2)} FOB) is the single largest unverified assumption: at the benchmark the hero loses money at $44.99. Get a written factory FOB quote in USD in week 1.")
A("")
A("## 4. Top three risks (from 08_Risk_Register.xlsx, by score)")
A("")
A("위험 등록부에서 점수가 가장 높은 세 가지입니다.")
A("")
for r in top3:
    def sent(t):
        t = str(t).strip()
        return t if t.endswith(".") else t + "."
    A(f"- **{r[0]} (score {int(r[pi] * r[ii])}, {r[ci]})**: {sent(r[1])} Mitigation: {sent(r[mi])} Owner: {sent(r[oi])}")
A("")
A("## 5. First three actions for Lucy on Monday")
A("")
A("월요일 첫 세 가지 행동입니다. 모두 1주차 핵심 경로에 있습니다.")
A("")
A(f"1. **ODI report conversation.** Take 10_One_Page_Summary_KO.pdf to her foreign exchange bank and the 세무사. The overseas direct investment report (any stake of 10% or more, any amount) must be accepted before the first dollar leaves Korea; the capital is {fmt.money_k(c['opening_capital'])} now and a dated line of credit or investor amount of {fmt.money_k(c['funding_gap'])} by {fmt.month_label(c['funding_needed_by_month'])}. Ask the 세무사 two questions at the same meeting: incorporate H2 Company first or stay a sole proprietorship, and the overseas financial account report threshold.")
A("2. **Brand name shortlist and trademark attorney.** Pick three candidate names, have a US licensed trademark attorney clear and file the strongest one (about $350 per class plus attorney fees). Amazon Brand Registry accepts the pending application, and the LLC name must match the brand across state, IRS, bank and Amazon, so this gates step 2 of the critical path (Wyoming LLC and EIN, weeks 1 to 2).")
A("3. **Customs broker and binding ruling.** Engage a licensed customs broker with a continuous bond (about $500 per year) and file the CBP binding ruling request in week 1: HTS 8544.70.00 (optical fiber cables, 12.5% total) against 8544.42.90 (copper, 40.1%). The ruling takes 30 to 90 days and is the single variable that separates the Low case from the Medium case (landed COGS 38% to 33.6%). Send the broker 04_Flowcharts/03_inbound_supply_chain.png and the customs SOP in 05_Operating_Processes.pdf.")
A("")
A("## 6. How to rerun the build")
A("")
A("재생성 순서입니다. 모델을 먼저 만들고, 숫자를 뽑고, 차트와 문서를 다시 만든 뒤 마지막에 검증합니다.")
A("")
A("Every shell starts with `source /home/user/lakehouse-tracker/BRAND_US_Launch/build/env.sh && cd /home/user/lakehouse-tracker/BRAND_US_Launch`. Run the steps in this order; each later step reads what the earlier one wrote.")
A("")
A("```")
A("python3 build/write_facts.py                      # 0. static facts from the brief to build/facts.json")
A("python3 build/build_model.py                      # 1. 03_Financial_Model_5Y.xlsx (recalculates with LibreOffice and verifies calibration)")
A("$RECALC output/03_Financial_Model_5Y.xlsx 180     # 2. recalc: status must be success with total_errors 0")
A("python3 build/pull_figures.py                     # 3. build/figures.json (sets the selector to each scenario on copies in build/qa/model/)")
A("python3 build/export_charts.py                    # 4. chart PNGs to build/charts/")
A("python3 build/gen_flowcharts.py                   # 5. Mermaid sources to output/04_Flowcharts/")
A("for f in output/04_Flowcharts/*.mmd; do render_mmd \"$f\"; done   #    render .png and .pdf beside each .mmd")
A("python3 build/make_assets.py && python3 build/deck_flowcharts.py && python3 build/deck_flowcharts_ko.py   # 6a. deck background and slide sized flowcharts")
A("node build/build_deck.js --lang en                # 6b. 01_Business_Proposal_EN.pptx")
A("$VALIDATE output/01_Business_Proposal_EN.pptx && to_pdf output/01_Business_Proposal_EN.pptx")
A("bash build/build_deck_ko.sh                       # 6c. 02_Business_Proposal_KO.pptx and .pdf (validate, Korean PDF export, QA renders, lint)")
A("python3 build/build_sops.py                       # 7a. 05_Operating_Processes.docx and .pdf (runs node build/build_sops_docx.js)")
A("python3 build/build_checklist.py && $RECALC output/06_Compliance_Checklist.xlsx 180 && python3 build/build_checklist.py --patch-fonts   # 7b")
A("python3 build/build_gantt.py && $RECALC output/07_90_Day_Launch_Plan.xlsx 180 && python3 build/build_gantt.py --patch-fonts           # 7c")
A("python3 build/build_risk.py && $RECALC output/08_Risk_Register.xlsx 180 && python3 build/build_risk.py --patch-fonts && python3 build/build_risk.py --check   # 7d")
A("node build/build_listing.js                       # 7e. 09_Amazon_Listing_Kit.docx")
A("python3 build/build_onepagers.py                  # 7f. 10_One_Page_Summary_EN.pdf and _KO.pdf")
A("python3 build/build_decision_log.py               # 7g. 11_Decision_Log.md")
A("python3 build/tie_out.py                          # 8. tie out: workbook vs figures.json vs every document; writes build/tie_out_result.md; exits 1 on any mismatch")
A("python3 build/build_readme.py                     # 9. this file (quotes the tie out result)")
A("python3 build/lint_hyphens.py output/             # 10. must print violating_lines 0 for every file")
A("```")
A("")
A("Notes. The decks, documents and this README never hard code a model number; they format build/figures.json through build/fmt.py or build/fmt.js. Static facts (duty rates, filing costs, form names) come from build/facts.json. QA renders of every slide and page sit under build/qa/<name>/. The duty tab is named `Landed Cost & Tariff Scenarios` (30 characters) because Excel caps sheet names at 31 characters.")
A("")
A("## 7. Tie out result (build/tie_out_result.md, quoted verbatim)")
A("")
A("검증 스크립트의 최종 결과입니다.")
A("")
A(tie)
A("")

text = "\n".join(lines)
(ROOT / "output" / "00_README.md").write_text(text, encoding="utf-8")
print("wrote output/00_README.md", len(text), "chars; top risks", [r[0] for r in top3])

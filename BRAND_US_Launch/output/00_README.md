# 00 README: BRAND US launch package

Built October 8, 2026. Owner Lucy (Hyejin Yoo), H2 Company, Seoul. Adviser Mark Taylor, Memphis. US entity BRAND LLC (Wyoming), brand name placeholder BRAND until chosen. Every number below comes from build/figures.json, pulled from output/03_Financial_Model_5Y.xlsx at its delivered state: Scenario Medium, Duty Copper (40.1%), FX 1,430 KRW per USD.

패키지 전체의 파일 목록, 읽는 순서, 재생성 방법, 핵심 숫자와 검증 결과를 담은 안내 파일입니다.

## 1. What is in output/

파일별 용도와 읽는 사람입니다.

| File | What it is for | Who reads it | When |
|---|---|---|---|
| 00_README.md | This file. Index, order of use, rebuild steps, headline numbers, tie out result. | Lucy, Mark | Read first |
| 01_Business_Proposal_EN.pptx and .pdf | 37 slide English proposal: market, product, structure, launch plan, economics, 5 year path, funding, risks, kill criteria, the ask. | Mark, bank (English readers), any US adviser | Step 3 |
| 02_Business_Proposal_KO.pptx and .pdf | The same 37 slides in Korean (Noto Sans KR). Amazon listing content stays in English. | Lucy, 세무사, bank | Step 3 |
| 03_Financial_Model_5Y.xlsx | Live 5 year model. One scenario selector (Low, Medium, High), one duty selector (Copper, Fiber, Stress), FX input. 14 tabs from Assumptions to Plan B. Every number in every other file comes from it. | Lucy, Mark, bank (if asked for the model), 세무사 (cash and repatriation tabs) | Step 4 |
| 04_Flowcharts/ (8 charts, each .mmd, .png, .pdf, plus README.md) | Critical path, order to cash, inbound supply chain, returns and the 5 Year Signal Guarantee, monthly cadence, Day 90 scale or kill tree, 5 year channel map, entity money and goods flow. | Lucy, Mark, broker (03 and 08), 3PL (03), bank and 세무사 (02 and 08) | With the deck and the SOPs |
| 05_Operating_Processes.docx and .pdf | Eight SOPs (owner, trigger, steps, system, KPI, failure modes) plus the compliance calendar. | Lucy, Mark, VA, broker (customs SOP), 3PL (inbound SOP) | Step 5 |
| 06_Compliance_Checklist.xlsx | Every US, Korea, Amazon, HDMI, FCC, Prop 65, CL3, insurance, customs and tax item with owner, due, cost, evidence, status. | Lucy, 세무사 (Korea rows), broker (customs rows), Mark | Step 5, then weekly |
| 07_90_Day_Launch_Plan.xlsx | Week by week Gantt (weeks 1 to 22) with owner, dependencies, cost by owner and week, and the Day 90 gate tab. | Lucy, Mark, attorney, broker, factory, 3PL | Step 5, then weekly |
| 08_Risk_Register.xlsx | 32 risks scored probability times impact, early warnings, mitigations, owners, kill criteria tab, heat map. | Lucy, Mark, bank (if asked) | Step 5, then monthly |
| 09_Amazon_Listing_Kit.docx | Parent and child ASIN structure, titles, bullets, A+ outline, backend keywords, image and video brief, FAQ, compliance notes. English only. | Mark, VA, photographer, factory (packaging text) | Weeks 6 to 9 |
| 10_One_Page_Summary_EN.pdf | One page for the bank and any English reader: who, structure, filings in order, capital need, timeline. | Bank, Mark | Step 1 (Monday) |
| 10_One_Page_Summary_KO.pdf | The same page in Korean for the foreign exchange bank and the 세무사. | 세무사, bank, Lucy | Step 1 (Monday) |
| 11_Decision_Log.md | The five open decisions, the defaults used, what changes if each goes the other way (quantified from the model), who decides, by when. | Lucy, Mark, 세무사 (ownership and entity rows) | Step 2 (Monday) |

Full directory listing (every delivered file):

- output/00_README.md
- output/01_Business_Proposal_EN.pdf
- output/01_Business_Proposal_EN.pptx
- output/02_Business_Proposal_KO.pdf
- output/02_Business_Proposal_KO.pptx
- output/03_Financial_Model_5Y.xlsx
- output/04_Flowcharts/01_launch_critical_path.mmd
- output/04_Flowcharts/01_launch_critical_path.pdf
- output/04_Flowcharts/01_launch_critical_path.png
- output/04_Flowcharts/02_order_to_cash_amazon.mmd
- output/04_Flowcharts/02_order_to_cash_amazon.pdf
- output/04_Flowcharts/02_order_to_cash_amazon.png
- output/04_Flowcharts/03_inbound_supply_chain.mmd
- output/04_Flowcharts/03_inbound_supply_chain.pdf
- output/04_Flowcharts/03_inbound_supply_chain.png
- output/04_Flowcharts/04_returns_and_signal_guarantee.mmd
- output/04_Flowcharts/04_returns_and_signal_guarantee.pdf
- output/04_Flowcharts/04_returns_and_signal_guarantee.png
- output/04_Flowcharts/05_monthly_operating_cadence.mmd
- output/04_Flowcharts/05_monthly_operating_cadence.pdf
- output/04_Flowcharts/05_monthly_operating_cadence.png
- output/04_Flowcharts/06_day90_scale_or_kill.mmd
- output/04_Flowcharts/06_day90_scale_or_kill.pdf
- output/04_Flowcharts/06_day90_scale_or_kill.png
- output/04_Flowcharts/07_five_year_channel_map.mmd
- output/04_Flowcharts/07_five_year_channel_map.pdf
- output/04_Flowcharts/07_five_year_channel_map.png
- output/04_Flowcharts/08_entity_money_goods_flow.mmd
- output/04_Flowcharts/08_entity_money_goods_flow.pdf
- output/04_Flowcharts/08_entity_money_goods_flow.png
- output/04_Flowcharts/README.md
- output/05_Operating_Processes.docx
- output/05_Operating_Processes.pdf
- output/06_Compliance_Checklist.xlsx
- output/07_90_Day_Launch_Plan.xlsx
- output/08_Risk_Register.xlsx
- output/09_Amazon_Listing_Kit.docx
- output/10_One_Page_Summary_EN.pdf
- output/10_One_Page_Summary_KO.pdf
- output/11_Decision_Log.md

## 2. Order of use

월요일에 시작하는 순서입니다. 은행과 세무사용 한 장 요약부터 시작합니다.

Monday sequence first, then the rest.

1. **10_One_Page_Summary_KO.pdf to the foreign exchange bank and the 세무사; 10_One_Page_Summary_EN.pdf to Mark and any English reader.** It carries the structure, the filings in order relative to the capital transfer, the opening capital and the funding gap. The bank needs it for the overseas direct investment (ODI) report before any capital moves. The 세무사 needs it for the ODI follow up reports, the overseas financial account report threshold (KRW 500M month end aggregate), Form 5472 and the question of whether to incorporate H2 Company first.
2. **11_Decision_Log.md.** Confirm or change the five defaults: brand name, LLC ownership (Lucy 100%, Mark under a consulting agreement), funding source (self funded to $39K, gap as a dated line of credit), FX (1,430), base duty scenario (Copper 40.1% pending the CBP binding ruling). Each row says what changes if the decision goes the other way.
3. **01_Business_Proposal_EN or 02_Business_Proposal_KO.** Read the executive summary (slide 2), the Day 90 gate (slide 19), the 18 month P&L (slide 23), the 5 year trajectory (slide 24), working capital (slide 27) and the ask (slide 35). Use the PDF to share; the pptx to edit.
4. **03_Financial_Model_5Y.xlsx.** Change only yellow cells with blue text. Scenario Summary shows all three scenarios side by side without touching the selector. KPI Dashboard carries the charts.
5. **07_90_Day_Launch_Plan.xlsx** for the weekly launch review (with 04_Flowcharts/01_launch_critical_path.png). Then 06_Compliance_Checklist.xlsx, 08_Risk_Register.xlsx and 05_Operating_Processes as the business starts running. 09_Amazon_Listing_Kit.docx goes to Mark and the VA in weeks 6 to 9.

## 3. Headline numbers (Medium case unless stated)

모델에서 가져온 핵심 숫자입니다. 중간(Medium) 시나리오 기준입니다.

| Measure | Value |
|---|---|
| 18 month revenue (Amazon US only) | $1.17M |
| 18 month net | $153K (13.1%) |
| Cumulative breakeven | Month 8 |
| Month 18 run rate and exit margin | $1.54M, 17.2% |
| Year 5, Low | revenue $0.77M, net ($2K), margin (0.3%) |
| Year 5, Medium | revenue $4.58M, net $613K, margin 13.4% |
| Year 5, High | revenue $8.30M, net $1,838K, margin 22.2% |
| Opening capital | $39K (self funded, August ask) |
| Peak cash need | $96K in Month 12 |
| Funding gap (1.2 times the lowest balance) | $116K, in place by Month 2 |
| Day 90 and Month 12 gates | Medium PASS and PASS; Low REVIEW and KILL; High PASS and PASS |

Other scenarios, 18 months: Low $0.63M revenue, ($20K) net, breakeven none, peak cash need $114K in Month 24. High $1.85M revenue, $447K net, breakeven Month 5, peak cash need $34K in Month 7.

Hero SKU (8K 10M) at $44.99: landed cost $17.10 (38.0%), net $8.00 per unit (17.8%). The same unit earned $1.86 gross in Korea wholesale. The Alibaba benchmark ($17 to $22 FOB for a comparable 30 ft AOC) against the H2 base ($7.38 FOB) is the single largest unverified assumption: at the benchmark the hero loses money at $44.99. Get a written factory FOB quote in USD in week 1.

## 4. Top three risks (from 08_Risk_Register.xlsx, by score)

위험 등록부에서 점수가 가장 높은 세 가지입니다.

- **R07 (score 20, Cash and FX)**: Working capital runs out before the ramp pays back: Medium peak cash need $96K in Month 12 against $39K opening capital, funding gap $116K needed by Month 2; Low peak $114K in Month 24. Mitigation: Commit the funding line (retained earnings first, then the line of credit or investor amount) before Month 2, not when the balance turns negative. Hold inventory cover at 2.5 months, order in smaller lots during the ramp, and run the monthly cash check in the close. Runway below 2 months is a Day 90 kill trigger. Owner: Lucy, Mark.
- **R03 (score 15, Product and compliance)**: Factory cost gap: the Alibaba benchmark for a comparable 30 ft AOC is $17 to $22 FOB against the H2 base of $7.38 FOB for the hero; at the benchmark the hero loses ($1.61) to ($8.61) per unit at $44.99 and Medium 18 month net turns to ($287K). Mitigation: Get the factory's written USD FOB quote for the five launch SKUs before the LLC is formed. The single largest unverified assumption in the plan: if the quote is more than 25% above the H2 base, rerun SKU Unit Economics and reprice or stop before the first PO. Owner: Lucy.
- **R01 (score 12, Tariff and customs)**: Pending second Section 301 action adds 10 points of duty on China origin goods (stress case 50.1% against the copper base 40.1%). Mitigation: Stress case is already modeled: Medium 18 month net falls from $153K to $134K, Year 5 net from $613K to $534K (full model recalc at the stress rate). Hold the price ladder test (lever 10) in reserve, push the fiber ruling (lever 1), buy no more than the 2.5 month cover before a rate change date, and ask the broker about an earlier entry date for goods already on the water. Owner: Broker, Mark.

## 5. First three actions for Lucy on Monday

월요일 첫 세 가지 행동입니다. 모두 1주차 핵심 경로에 있습니다.

1. **ODI report conversation.** Take 10_One_Page_Summary_KO.pdf to her foreign exchange bank and the 세무사. The overseas direct investment report (any stake of 10% or more, any amount) must be accepted before the first dollar leaves Korea; the capital is $39K now and a dated line of credit or investor amount of $116K by Month 2. Ask the 세무사 two questions at the same meeting: incorporate H2 Company first or stay a sole proprietorship, and the overseas financial account report threshold.
2. **Brand name shortlist and trademark attorney.** Pick three candidate names, have a US licensed trademark attorney clear and file the strongest one (about $350 per class plus attorney fees). Amazon Brand Registry accepts the pending application, and the LLC name must match the brand across state, IRS, bank and Amazon, so this gates step 2 of the critical path (Wyoming LLC and EIN, weeks 1 to 2).
3. **Customs broker and binding ruling.** Engage a licensed customs broker with a continuous bond (about $500 per year) and file the CBP binding ruling request in week 1: HTS 8544.70.00 (optical fiber cables, 12.5% total) against 8544.42.90 (copper, 40.1%). The ruling takes 30 to 90 days and is the single variable that separates the Low case from the Medium case (landed COGS 38% to 33.6%). Send the broker 04_Flowcharts/03_inbound_supply_chain.png and the customs SOP in 05_Operating_Processes.pdf.

## 6. How to rerun the build

재생성 순서입니다. 모델을 먼저 만들고, 숫자를 뽑고, 차트와 문서를 다시 만든 뒤 마지막에 검증합니다.

Every shell starts with `source /home/user/lakehouse-tracker/BRAND_US_Launch/build/env.sh && cd /home/user/lakehouse-tracker/BRAND_US_Launch`. Run the steps in this order; each later step reads what the earlier one wrote.

```
python3 build/write_facts.py                      # 0. static facts from the brief to build/facts.json
python3 build/build_model.py                      # 1. 03_Financial_Model_5Y.xlsx (recalculates with LibreOffice and verifies calibration)
$RECALC output/03_Financial_Model_5Y.xlsx 180     # 2. recalc: status must be success with total_errors 0
python3 build/pull_figures.py                     # 3. build/figures.json (sets the selector to each scenario on copies in build/qa/model/)
python3 build/export_charts.py                    # 4. chart PNGs to build/charts/
python3 build/gen_flowcharts.py                   # 5. Mermaid sources to output/04_Flowcharts/
for f in output/04_Flowcharts/*.mmd; do render_mmd "$f"; done   #    render .png and .pdf beside each .mmd
python3 build/make_assets.py && python3 build/deck_flowcharts.py && python3 build/deck_flowcharts_ko.py   # 6a. deck background and slide sized flowcharts
node build/build_deck.js --lang en                # 6b. 01_Business_Proposal_EN.pptx
$VALIDATE output/01_Business_Proposal_EN.pptx && to_pdf output/01_Business_Proposal_EN.pptx
bash build/build_deck_ko.sh                       # 6c. 02_Business_Proposal_KO.pptx and .pdf (validate, Korean PDF export, QA renders, lint)
python3 build/build_sops.py                       # 7a. 05_Operating_Processes.docx and .pdf (runs node build/build_sops_docx.js)
python3 build/build_checklist.py && $RECALC output/06_Compliance_Checklist.xlsx 180 && python3 build/build_checklist.py --patch-fonts   # 7b
python3 build/build_gantt.py && $RECALC output/07_90_Day_Launch_Plan.xlsx 180 && python3 build/build_gantt.py --patch-fonts           # 7c
python3 build/build_risk.py && $RECALC output/08_Risk_Register.xlsx 180 && python3 build/build_risk.py --patch-fonts && python3 build/build_risk.py --check   # 7d
node build/build_listing.js                       # 7e. 09_Amazon_Listing_Kit.docx
python3 build/build_onepagers.py                  # 7f. 10_One_Page_Summary_EN.pdf and _KO.pdf
python3 build/build_decision_log.py               # 7g. 11_Decision_Log.md
python3 build/tie_out.py                          # 8. tie out: workbook vs figures.json vs every document; writes build/tie_out_result.md; exits 1 on any mismatch
python3 build/build_readme.py                     # 9. this file (quotes the tie out result)
python3 build/lint_hyphens.py output/             # 10. must print violating_lines 0 for every file
```

Notes. The decks, documents and this README never hard code a model number; they format build/figures.json through build/fmt.py or build/fmt.js. Static facts (duty rates, filing costs, form names) come from build/facts.json. QA renders of every slide and page sit under build/qa/<name>/. The duty tab is named `Landed Cost & Tariff Scenarios` (30 characters) because Excel caps sheet names at 31 characters.

## 7. Tie out result (build/tie_out_result.md, quoted verbatim)

검증 스크립트의 최종 결과입니다.

# Tie out result: PASS

Run October 08, 2026 20:47 by build/tie_out.py. 383 checks, 383 passed, 0 failed. Tolerance 0.5%. Workbook: output/03_Financial_Model_5Y.xlsx at Scenario 2 (Medium), Duty 1 (Copper), FX 1430.
Hyphen lint: 32 files, 0 violating lines.

| Deliverable | Checks | Passed | Failed | Details |
|---|---|---|---|---|
| 03_Financial_Model_5Y.xlsx | 12 | 12 | 0 | all pass |
| 01_Business_Proposal_EN.pptx | 84 | 84 | 0 | all pass |
| 02_Business_Proposal_KO.pptx | 84 | 84 | 0 | all pass |
| 05_Operating_Processes.docx | 62 | 62 | 0 | all pass |
| 05_Operating_Processes.pdf | 4 | 4 | 0 | all pass |
| 09_Amazon_Listing_Kit.docx | 14 | 14 | 0 | all pass |
| 10_One_Page_Summary_EN.pdf | 36 | 36 | 0 | all pass |
| 10_One_Page_Summary_KO.pdf | 36 | 36 | 0 | all pass |
| 06_Compliance_Checklist.xlsx | 12 | 12 | 0 | all pass |
| 07_90_Day_Launch_Plan.xlsx | 12 | 12 | 0 | all pass |
| 08_Risk_Register.xlsx | 17 | 17 | 0 | all pass |
| 00_README.md | 4 | 4 | 0 | all pass |
| 11_Decision_Log.md | 5 | 5 | 0 | all pass |
| output/ (all files) | 1 | 1 | 0 | all pass |

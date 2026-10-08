# BUILD SPEC: shared contract for every build agent

Read this after BUILD_BRIEF.md and inputs/DECISIONS.md. The brief says what to build. This file says how, so every file agrees with every other file. Where this file and the brief conflict, the brief wins and you report the conflict in your result.

## 0. Paths and tools

- ROOT = /home/user/lakehouse-tracker/BRAND_US_Launch. Inputs in ROOT/inputs, scripts and data pulls in ROOT/build, deliverables in ROOT/output.
- Start every shell with: `source /home/user/lakehouse-tracker/BRAND_US_Launch/build/env.sh` (sets NODE_PATH, $RECALC, $SOFFICE, $VALIDATE, $APPLY_THEME, render_mmd, to_pdf, to_jpg).
- Excel: Python openpyxl with live formulas. Recalculate with `$RECALC file.xlsx 180` (JSON result; status must be success with total_errors 0). Never use XLOOKUP, FILTER, UNIQUE, SORT, SEQUENCE, LET, LAMBDA. Use INDEX, MATCH, SUMIFS, SUMPRODUCT, IFERROR, CHOOSE, OFFSET.
- PowerPoint: node pptxgenjs (global). After writeFile run `$VALIDATE deck.pptx`. Convert to PDF with `to_pdf deck.pptx`, render pages with `to_jpg deck.pdf dir`, then look at every page image with the Read tool.
- Word: node docx (global). Convert to PDF with `to_pdf file.docx`, render with `to_jpg`, inspect.
- PDF one pagers: build as docx with node docx then `to_pdf`, or with reportlab. Korean text needs font "Noto Sans KR" (installed at /usr/share/fonts/truetype/notokr/NotoSansKR-VF.ttf). For reportlab register that TTF explicitly.
- Flowcharts: Mermaid source (.mmd) rendered with `render_mmd file.mmd` (uses mmdc with build/puppeteer.json). Produces .png and .pdf beside the source.
- Charts PNG: matplotlib (installed). Font: DejaVu Sans for EN, Noto Sans KR for KO labels.
- Text lint: `python3 $BUILD/lint_hyphens.py <files or dirs>` must print violating_lines 0 for every file you deliver. It understands md, mmd, csv, docx, pptx, xlsx, pdf.
- Skill references if you need detail: /mnt/skills/public/xlsx/SKILL.md, /mnt/skills/public/pptx/SKILL.md, /mnt/skills/public/docx/SKILL.md, /mnt/skills/public/pdf/SKILL.md.
- The machine has 4 CPUs. Give LibreOffice conversions a 180 second timeout and do not run more than one soffice job at a time inside your own agent.

## 1. Standing rules (apply to every file)

1. No hyphens, en dashes or em dashes in any text, including chart labels, table cells, slide text, Mermaid node labels, file names shown to readers, and JSON strings that end up in a document. Write "long run" not "long-run", "5 year" not "5-year", "co owner" not "co-owner", "ecommerce", "ebook", "in wall" not "in-wall", "third party" not "third-party". Date ranges read "weeks 1 to 2", "2027 to 2031". Negative money reads ($19K). Official form and registration numbers keep their format: Form SS-4, W-8BEN-E, 599-64-00496, EIN format 12-3456789, C-TPAT. Mermaid arrow syntax (-->) is code, not text, and is fine.
2. Phone numbers use dots: 901.356.3328. No phone number appears unless the brief gives it.
3. The US brand name is the literal placeholder BRAND (all caps) everywhere. The LLC is "BRAND LLC". The Korean company is "H2 Company" (에이치투 컴퍼니 is not confirmed; write H2 Company in Korean text too).
4. Placeholders are explicit and labeled: `[PLACEHOLDER: HDMI Adopter certificate number from factory]`. Never invent certification numbers, review counts, star ratings, supplier quotes, market sizes, or customer names. A market size or benchmark that is not in the brief is a labeled management estimate or a placeholder with the source still to be found.
5. Every assumption is tagged as one of: `Evidence based` (brief section 1, the August deal, the August model, a verified regulatory fact) or `Management estimate`. The workbook carries a Basis column; decks and documents say "estimate" where the number is one.
6. Every number in a deck, document, one pager or README comes from build/figures.json, which is pulled from the recalculated workbook. Generators load the JSON; they never hard code a model number. Static facts from the brief (tariff rates, fees, form names, costs of filings) live in build/facts.json so they are also single sourced.
7. Plain language, short sentences, lead with the number or the action. Define each acronym on first use in each document and in the glossary. No corporate filler ("leverage", "synergy", "partner with", "excited").
8. Audience split: Lucy reads Korean and basic English. 세무사 (Korean tax accountant) and bank read Korean. Mark reads everything. Amazon listing content stays in English in every language version.
9. Currency: USD with $ and thousands separators. KRW with "KRW" prefix and thousands separators (KRW 24,400,000 or KRW 24.4M). Millions as $1.17M, thousands as $153K in decks; full figures in the workbook.
10. Dates: October 8, 2026 style. Month numbers in the model are "Month 1" to "Month 24"; Year 1 to Year 5. Month 1 is the first month with inventory live on Amazon (assumed early 2027). Do not assign calendar months to model months.
11. Deliver only the files the brief names, with the exact names, into ROOT/output. Keep generators, JSON, scratch renders and QA images in ROOT/build (QA renders under ROOT/build/qa/).

## 2. Master data (used by the model, decks and documents)

### 2.1 SKUs

FX default 1,430 KRW per USD (input cell). FOB China (customs value) = H2 landed Korea cost in KRW divided by FX times 0.92 (management estimate: removes Korea inbound freight and handling). US retail prices and FBA fees are management estimates pending the Amazon rate card check. Weights are packaged weights, management estimates.

| SKU | Spec | Length m | H2 cost KRW (evidence) | US list | US sale price (ASP input) | Packaged weight lb | FBA fee estimate $ | Launch SKU |
|---|---|---|---|---|---|---|---|---|
| 8K_7M | HDMI 2.1 8K hybrid AOC | 7 | 10,525 | 49.99 | 39.99 | 0.35 | 4.90 | Yes |
| 8K_10M | HDMI 2.1 8K hybrid AOC (hero) | 10 | 11,470 | 59.99 | 44.99 | 0.45 | 4.90 | Yes |
| 8K_15M | HDMI 2.1 8K hybrid AOC | 15 | 13,376 | 69.99 | 54.99 | 0.60 | 5.50 | Yes |
| 8K_20M | HDMI 2.1 8K hybrid AOC | 20 | 15,266 | 79.99 | 64.99 | 0.75 | 5.50 | Yes |
| 8K_25M | HDMI 2.1 8K hybrid AOC | 25 | 17,172 | 89.99 | 74.99 | 0.90 | 6.60 | No |
| 8K_50M | HDMI 2.1 8K hybrid AOC | 50 | 27,349 | 149.99 | 119.99 | 1.90 | 8.40 | No |
| 4K_15M | HDMI 2.0 4K hybrid AOC | 15 | 12,680 | 44.99 | 34.99 | 0.60 | 5.50 | No |
| 4K_20M | HDMI 2.0 4K hybrid AOC | 20 | 14,122 | 49.99 | 39.99 | 0.75 | 5.50 | No |
| 4K_30M | HDMI 2.0 4K hybrid AOC | 30 | 19,542 | 64.99 | 49.99 | 1.10 | 6.60 | Yes |
| 4K_40M | HDMI 2.0 4K hybrid AOC | 40 | 22,609 | 74.99 | 59.99 | 1.50 | 8.40 | No |

Naming rule on packaging and listings: never "HDMI 2.1" or "HDMI 2.0" on the box or in promotion. Use "Ultra High Speed HDMI Cable" (the 8K line) and "Premium High Speed HDMI Cable" (the 4K line, pending certification; placeholder). Titles may say 48Gbps and 8K. The spec column above is internal only.

Hero SKU unit economics (August model, evidence based): price $44.99, landed cost $17.10 (38%), referral $6.75 (15%), FBA $5.50, advertising $5.40 (12%), returns and storage reserve $2.25 (5%), net $8.00 (17.8%). The Korea wholesale channel earned $1.86 gross on the same unit.

### 2.2 Cost build per unit (SKU Unit Economics tab)

- Air freight DDP estimate: $8.50 per kg packaged (management estimate for consolidated express, first lots). Ocean LCL plus drayage: $1.10 per kg (management estimate). 1 lb = 0.4536 kg.
- Duty = duty rate (section 2.4) times FOB. Merchandise processing fee and harbor maintenance fee: add 0.35% of FOB for ocean only (management estimate).
- Broker entry fee $150 per entry and continuous bond $500 per year, amortized: $0.40 per unit at launch volumes (management estimate).
- FBA inbound placement and prep: $0.50 per unit. Third party QC and FNSKU labeling at factory: $0.25 per unit (management estimates).
- Three cost bases to show side by side: (a) H2 cost base (table above), (b) August model assumption (38% of ASP for the hero, evidence based as the August deck's assumption), (c) Alibaba benchmark FOB $17 to $22 for a comparable 30 ft (9 m) AOC (evidence based quote, flagged). The gap between (a) and (c) is the single largest unverified assumption in the plan: at the Alibaba cost the hero SKU loses money at $44.99. State this plainly wherever unit economics appear.

### 2.3 Eighteen month Amazon model (reconciliation target; evidence based)

Shared: referral 15%, FBA $5.50 per unit, storage and inbound 2% of revenue, fixed software insurance misc $850 per month, Professional seller fee $39.99 per month, one time launch costs $9,500 in Month 1 (trademark filing, LLC, EIN, broker setup, ruling request, photography, insurance deposit, samples).

Medium (August deck, evidence based): ASP $42, landed COGS 38%, returns 3%, ad spend 30% of revenue months 1 to 3, 22% months 4 to 6, 15% months 7 to 12, 11% months 13 to 18. Units by month: 150, 240, 360, 500, 650, 820, 1000, 1200, 1400, 1600, 1800, 2050, 2250, 2450, 2600, 2750, 2900, 3050. Results the model must reproduce: revenue $1,166,340 (brief says $1.17M), net $153,020 (13.1%), cumulative breakeven Month 8, Month 18 run rate $1,537,200 (brief says $1.5M), exit margin 17.2%.

Low (calibrated to the brief's results, management estimate): units = Medium times 0.60 rounded, ASP $38, landed COGS 43%, returns 4%, ad spend 7 points above Medium in every phase (37, 29, 22, 18). Reproduces revenue $633,156 (brief $633K), net ($19,636) (brief ($19K)), no cumulative breakeven in 18 months.

High (calibrated, management estimate): units = Medium times 1.55 rounded, ASP $43, landed COGS 32% (fiber ruling plus ocean freight plus volume pricing), returns 2.5%, ad spend 4 points below Medium in Months 1 to 12 and at the 8% floor in Months 13 to 18 (26, 18, 11, 8, never below 8). Reproduces revenue $1,851,021 (brief $1.85M; Excel ROUND half up on 232.5 and 4262.5 units adds $86 to a pure decimal build), net $447,155 (24.2%; brief $455K, 24.6%, inside the 2% tolerance), cumulative breakeven Month 5, Month 18 run rate $2,439,648 (brief $2.4M). Amended in fix round 2: the earlier figure $457,804 (24.7%) needed 7% ads in Months 13 to 18, which breaks the 8% floor, so it is not reproducible. Decks quote $447K and 24.2% from figures.json. Months 19 to 24 carry a separate ad spend row (High 9.5%, its Year 3 rate) so High Year 2 ad spend across all channels stays above 8% (8.1% blended).

Tolerance for reconciliation: within 1% of the brief's figures for Medium; within 2% for Low and High. Report the exact model values; the decks quote the model, not the brief.

### 2.4 Duty scenarios (Landed Cost and Tariff Scenarios tab)

| Scenario | HTS | MFN | Section 301 list rate | Section 301 forced labor (late July 2026) | Total | Basis |
|---|---|---|---|---|---|---|
| Copper (default) | 8544.42.90 | 2.6% | 25% (List 3) | 12.5% | 40.1% | Evidence based (brief) |
| Fiber | 8544.70.00 | 0% | 0% assumed | 12.5% | 12.5% | Evidence based rates; List 3 coverage of 8544.70.00 is a broker check (placeholder) |
| Stress | 8544.42.90 plus pending second Section 301 action | 2.6% | 25% | 12.5% + 10 points | 50.1% | Management estimate per brief |
| Fiber if also on List 3 | 8544.70.00 | 0% | 25% | 12.5% | 37.5% | Shown as a check row only, not a selectable scenario |

P&L landed COGS % for the selected scenario = scenario base landed % (2.3) + (selected duty rate minus copper rate) times hero FOB share of ASP (from the SKU tab, about 16%) + freight mode delta + FX delta. Medium plus Fiber must land near 33.6%, matching the August lever "38% to 33%". FX delta = hero FOB share times (1430 divided by selected FX, minus 1), applied because the only verified cost data is in KRW; flag that a USD factory quote removes this effect.

### 2.5 Five year drivers (Annual P&L tab). Starting values; calibrate within the stated bounds so Year 5 lands in band

Targets: Medium Year 5 revenue $4.0M to $5.0M, net 10% to 15%. High Year 5 $8M to $10M. Low Year 5 near $1M (0.7 to 1.2) with negative or near zero net and the Month 12 kill criteria triggered. Year 1 = Months 1 to 12 of the monthly model, Amazon US only. Year 2 = Months 13 to 24 (Months 19 to 24 extend the ramp) plus eBay and Walmart overflow plus B2B installer. Year 3 adds DTC site and the roadmap SKUs (HDMI 2.2 16K, DP 2.1, USB4 AOC: modeled as a mix uplift to ASP of +3% in Year 3, management estimate). Years 4 and 5 add Amazon Canada and the pro line.

Amazon US drivers:

| Driver | Low | Medium | High | Bounds |
|---|---|---|---|---|
| Months 19 to 24 units | flat at Month 18 | Month 18 + 100 per month | Month 18 + 160 per month | |
| Lost quarter | Months 19, 20, 21 units at 10% of trend (listing suppression). Amended in fix round 2: the starting value was Months 15 to 17, but that makes the Low 18 month reconciliation in 2.3 impossible (about $169K of revenue lost). The brief only requires one lost quarter in Low. Do not move it back without re calibrating the Low 18 month block | none | none | Low must include it |
| Unit growth Y3, Y4, Y5 | +5%, +3%, 0% | +30%, +20%, +12% | +35%, +25%, +15% | adjust ±5 points |
| ASP Y1 and Y2, then yearly change | $38, then minus 3% per year | $42, then minus 2% per year | $43, then minus 1.5% per year | compression at least 1 point of margin per year overall |
| Ad spend % Y3, Y4, Y5 | 12, 12, 12 | 9.5, 9.0, 8.5 | 9.0, 8.5, 8.0 | never below 8% |
| Landed COGS % change vs base | 0 | Y2 minus 1.5, Y3 minus 2.5, Y4 minus 2.0, Y5 minus 1.5 (ocean, volume, then creep) | Y2 minus 1.0, Y3 minus 1.5, Y4 minus 1.0, Y5 minus 0.5 | |
| Returns | 4% flat | 3% Y1, 2.5% Y2 onward | 2.5% Y1, 2% Y2 onward | never below 2% |
| FBA fee inflation | +3% per year from Y2 | +3% per year | +3% per year | realism |
| Referral | 15% | 15% | 15% | |
| Storage and inbound | 2% | 2% | 2% | |

Fixed costs (software, insurance, people, compliance) per month by year, Medium: Y1 $850 + $39.99 seller fee; Y2 $1,500; Y3 $4,000 (adds Mark consulting retainer $2,000, virtual assistant $1,000, tools and insurance $1,000); Y4 $7,000; Y5 $10,000. High: 1.3 times Medium from Y2. Low: Y2 $1,200 then frozen. Add yearly compliance: Form 5472 preparation $800, Wyoming annual report $60, continuous bond $500, trademark maintenance when due (placeholder $0 until year 5), product liability insurance $1,200 per year once monthly sales pass $10K (Medium from Month 4).

Other channels (revenue, Medium; High = 1.8 times; Low as stated):

| Channel | Y2 | Y3 | Y4 | Y5 | Economics | Low case |
|---|---|---|---|---|---|---|
| eBay and Walmart overflow | 8% of Amazon US revenue | 10% | 12% | 12% | referral 13%, multi channel fulfillment $6.50 per unit, ads 5%, returns 3%, same landed COGS | 8% of Amazon |
| B2B installer and AV integrator | $120K | $350K | $600K | $850K | price 60% of retail, landed COGS, 3PL pick pack ship $4 per unit, sales commission 5%, marketing 3%, returns 2% (amended in fix round 2 from 1.5%: the brief's floor, returns never below 2%, wins) | $60K from Y3, flat |
| DTC site | 0 | $100K | $250K | $400K | retail price, payment 3%, fulfillment and shipping $11 per unit, ads 15%, returns 3% | none |
| Amazon Canada | 0 | 0 | $120K | $250K | ASP 8% below US, FBA $6.50, referral 15%, ads 12%, returns 3%, duty on US to Canada transfer placeholder 0 | none |
| Pro line (CL3, plenum, conduit pull, long run fiber KVM) | 0 | 0 | $150K | $350K | ASP $120, landed COGS 40%, sold 60% B2B and 40% Amazon with those channels' cost structures | none |

Label every driver with Basis = Management estimate except the ones traced to the August model.

### 2.6 Working capital and cash (Cash Flow and Working Capital tab, monthly for Months 1 to 24, annual for Years 3 to 5)

Inputs: opening capital $39,000 (DECISIONS, self funded); inventory cover target 2.5 months of forward unit forecast (Medium), 3.0 (Low, slower sell through), 2.0 (High); PO deposit 30% at order, 70% at shipment; lead time 1 month air (Months 1 to 6), 2 months ocean after the 90 day gate for proven SKUs (Months 7 onward, Medium and High; Low stays air); Amazon payout lag 14 days, so the receivable at month end = 14/30 of the month's net Amazon proceeds; in transit inventory is funded by the business and shown separately from on hand inventory. Years 3 to 5: inventory days 75 (Medium), 90 (Low), 60 (High) of landed COGS; receivable 14 days of revenue.

Outputs: monthly cash balance, minimum cash balance and its month (peak cash need), runway in months at the Day 90 gate (cash at end of Month 3 divided by average monthly operating cash burn of Months 4 to 6 if negative, else "funded"), funding gap = max(0, 1.2 times negative minimum balance) and the month it must be in place (the month before the balance first goes negative). Funding plan rows: retained earnings first, then a line of credit or investor amount with the date. Show the result for all three scenarios in Scenario Summary.

### 2.7 Kill criteria (Funding and Breakeven tab, Risk Register, decks, SOPs, decision tree)

Gate results are PASS, REVIEW or KILL: KILL if any kill trigger fires, PASS if every modeled pass threshold is met, REVIEW otherwise (fix and retest within 30 days; added in fix round 2 because Low misses the Day 90 pass thresholds without firing a kill trigger). The thresholds live in one Gate thresholds block on Assumptions (figures.json gate_thresholds) and every gate formula and its wording reads them.

Day 90 gate (end of Month 3). Pass requires all of: unit session conversion at or above 8% on at least 3 of the 5 launch SKUs; TACOS (total advertising cost of sales, ad spend divided by total revenue) at or below 30% in Month 3 and falling month over month; return rate at or below 4%; contribution per unit before advertising at or above $11 on at least 3 of 5 SKUs (Medium hero: $12.14); cash runway at or above 4 months after funding the next purchase order; no suppressed listing and no intellectual property complaint open more than 14 days. Kill triggers (any one): hero conversion below 5%; TACOS above 40% in Month 3; return rate above 7%; hero contribution before advertising below $8; runway below 2 months.

Month 12 gate. Pass: monthly revenue at or above $70K (Medium Month 12 is $86,100; Low is $46,740); cumulative net at or above $0; TACOS at or below 15%; return rate at or below 3%; operating cash flow positive two consecutive months. Kill triggers (any one): revenue below $50K per month; cumulative net below ($25K); TACOS above 25%; return rate above 5%; three consecutive months of negative operating cash flow with a funding gap above $50K. Kill means: no further capital, sell down inventory, delist losers, move to Plan B.

Plan B (its own small P&L in Funding and Breakeven): B2B installer distribution in the US (Year 1 $120K revenue, Year 2 $300K, Year 3 $450K at the B2B economics above, 3PL only, no Amazon fees) plus continued Korea wholesale (August deal pace: 4 deals per year of 1,400 units at $2.50 gross profit per unit growing 15% per year, management estimate). Plan B is a side business: show its net honestly.

### 2.8 Twelve levers (Levers tab, deck slide, SOPs)

| # | Lever | Assumption it moves | From | To | Owner | Evidence |
|---|---|---|---|---|---|---|
| 1 | CBP binding ruling: fiber classification 8544.70.00 | Duty rate, landed COGS % | 38% | 33.6% | Broker, Mark | August deck lever, brief |
| 2 | Air to ocean freight after the Day 90 gate | Freight per unit, landed COGS % | air $8.50 per kg | ocean $1.10 per kg | Lucy, 3PL | August deck lever |
| 3 | Direction proof connector design | Return rate | 3% | 2% | Lucy, factory | August deck lever |
| 4 | Keyword pruning and negative keywords | Ad spend % | phase rate | minus 2 points | Mark, VA | August deck lever |
| 5 | Review velocity (Vine, insert card, Brand Registry follow up) | Conversion, unit ramp | base | +10% units | Mark | August deck lever |
| 6 | Length variations under one parent ASIN | Conversion, review pooling, ad efficiency | separate ASINs | one parent | Mark | August deck lever |
| 7 | Factory volume pricing tiers | FOB cost | H2 base | minus 5% at 5,000 units per quarter | Lucy | Management estimate |
| 8 | Packaging to a smaller FBA size tier | FBA fee per unit | $5.50 | $4.75 on 7M and 10M | Lucy, factory | Management estimate |
| 9 | 3PL feeding FBA, avoid aged inventory surcharges | Storage and inbound % | 2% | 1.5% | 3PL, Lucy | Management estimate |
| 10 | Price ladder and coupons (list vs sale price) | ASP | $42 | ±5% test | Mark | Management estimate |
| 11 | Off Amazon email list (Spec Match finder, trade up registration) | DTC revenue, ad % | 0 | DTC channel from Year 3 | Mark | Management estimate |
| 12 | B2B installer program (AV integrators, custom installers) | B2B channel revenue | 0 | Year 2 $120K | Mark, Lucy | Management estimate |

### 2.9 Static facts (build/facts.json)

The model builder writes build/facts.json from the brief: duty table (2.4), entity costs (Wyoming $100 to form, $60 per year; Delaware about $300 per year), Form 5472 penalty $25,000, trademark about $350 per class plus attorney fees, continuous bond about $500 per year, product liability insurance $1M required above $10K monthly sales, Korean overseas financial account report threshold KRW 500M month end aggregate, Korean overseas direct investment report (any stake 10% or more, any amount, before capital moves), the treaty position (1976 US Korea tax treaty, Article 9(3)(b), no PE from FBA storage), de minimis gone, IEEPA tariffs struck down February 20, 2026, Section 122 10% ran February 24 to July 24, 2026, HDMI naming rule, August funding ask $39K, August deal totals. Documents read facts.json for these.

## 3. Workbook structure (03_Financial_Model_5Y.xlsx)

Exact tab names in this order: `Cover`, `Assumptions`, `FX`, `SKU Unit Economics`, `Landed Cost & Tariff Scenarios`, `Monthly PL Y1 Y2`, `Annual PL Y1 to Y5`, `Cash Flow and Working Capital`, `Funding and Breakeven`, `Sensitivity`, `KPI Dashboard`, `Levers`, `Scenario Summary`, `Plan B`. (Amended in fix round 2: the duty tab was specified as `Landed Cost and Tariff Scenarios`, 32 characters, but Excel caps sheet names at 31, so the 30 character ampersand form is the delivered name; the tab's A1 title carries the full wording. Documents that name the tab use the 30 character form.)

- Cover: title, how to use, legend (yellow fill with blue text = input; black = formula; green = link to another sheet), the scenario selector description, build date, the no hyphen note is not needed here.
- Assumptions: Low, Medium, High in three adjacent columns for every driver, and a fourth "Active" column pulled by `=INDEX(Low:High, selector)` style formulas (use CHOOSE or INDEX with the selector cell). Selector cells on Assumptions, named ranges: `Scenario` (1 Low, 2 Medium, 3 High; default 2), `DutyScenario` (1 Copper, 2 Fiber, 3 Stress; default 1), `FXRate` (default 1430). Every other sheet reads the Active column, never the Low, Medium or High columns directly. Put a data validation list on the selectors and a text cell that shows the active scenario name.
- FX: the FX input, the three FX scenarios, and KRW to USD conversion of the capital, the August deal, and profit repatriation.
- SKU Unit Economics: one row per SKU (2.1) with FOB KRW, FOB USD, weight, air freight, ocean freight, duty by each of the three scenarios, fees, FBA, referral, ads, returns, net per unit and net %, for both air and ocean, plus the three cost bases (2.2). Hero row highlighted.
- Landed Cost and Tariff Scenarios: the duty table (2.4) and the landed % build (2.4 formula) for the active scenario, and the waterfall data for the chart (copper vs fiber vs stress net per hero unit).
- Monthly PL Y1 Y2: Months 1 to 24, Amazon US. Rows: units, ASP, revenue, landed COGS, referral, FBA, returns, storage and inbound, advertising, fixed, seller fee, launch costs, net, cumulative net, net margin, TACOS, contribution per unit before ads, contribution per unit after ads. Months 13 to 24 also carry the eBay and Walmart overflow and B2B rows so Year 2 ties to the annual tab.
- Annual PL Y1 to Y5: by channel (Amazon US, eBay and Walmart, B2B installer and AV integrator, DTC site, Amazon Canada, Pro line) with revenue, COGS, channel fees, fulfillment, advertising, returns, contribution, then fixed costs, compliance, net, net margin; Year 1 and Year 2 must equal the sums from the monthly tab (add a check row that shows 0).
- Cash Flow and Working Capital: section 2.6, monthly 1 to 24 then annual Y3 to Y5. Peak need and month, runway, funding gap and date.
- Funding and Breakeven: opening capital, cumulative breakeven month (formula), funding plan rows, Day 90 and Month 12 gate tables with the model's own values next to each threshold and a PASS or KILL formula, Plan B P&L (2.7).
- Sensitivity: one way sensitivity of Year 5 net (and 18 month net) to duty rate (12.5, 40.1, 50.1), ASP (minus 10%, base, plus 10%), ad spend (minus 3, base, plus 3 points), return rate (2, 3, 5%), FX (1300, 1430, 1550), factory cost (minus 10%, base, plus 25%, Alibaba). Implement with formulas that recompute a compact one line P&L from the Assumptions cells plus the delta, or with the full model if you can wire it; if any cell is computed by Python and pasted as a value, mark the block "values pasted by build script, rerun build to refresh" in a note cell and report it. Tornado data block for the chart.
- KPI Dashboard: openpyxl native charts: revenue and net by month (Months 1 to 24, line), annual revenue by channel (stacked column, Years 1 to 5), cash balance by month with the minimum marked (line plus a marker series), sensitivity tornado (bar), duty scenario waterfall (stacked column with invisible base). Plus a KPI table: 18 month revenue and net, breakeven month, Year 5 revenue and net, peak cash need, funding gap.
- Levers: table 2.8.
- Scenario Summary: all three scenarios side by side computed without changing the selector (either by three parallel compact calculations, or by the build script writing the three result sets as values with a note; prefer live). Must include: 18 month revenue, net, margin, breakeven month, run rate, exit margin; Year 1 to 5 revenue and net; Year 5 margin; peak cash need and month; funding gap and month; gate outcomes.
- Plan B: the small P&L from 2.7.

Conventions: Arial 10. Inputs: yellow fill (FFFF00) with blue font (0000FF). Formulas black. Cross sheet links green (008000). Currency format `$#,##0;($#,##0);0` (zero shows as 0, not as a dash, because of the no hyphen rule). Percent `0.0%`. Years as text. Freeze panes under headers. Column widths readable. Named ranges for every figure figures.json pulls (list them in build/figures_map.json as name: sheet!cell).

## 4. figures.json (build/figures.json)

Pulled by build/pull_figures.py from the recalculated workbook, after setting the selector to each scenario in turn (recalc three times) or from the Scenario Summary tab. Required keys:

```
meta: { fx, scenario_default: "Medium", duty_default: "Copper", built: "October 8, 2026", workbook: "output/03_Financial_Model_5Y.xlsx" }
scenarios.{Low|Medium|High}: {
  m18: { revenue, net, net_margin, breakeven_month (number or null), run_rate, exit_margin, units },
  monthly: [ {month, units, revenue, net, cum_net, cash_balance, tacos, contribution_before_ads, contribution_after_ads} x 24 ],
  annual: { Y1..Y5: { revenue, net, net_margin, revenue_by_channel: {amazon_us, ebay_walmart, b2b, dtc, amazon_canada, pro_line}, cogs, fees, advertising, fixed } },
  cash: { peak_need, peak_need_month, min_balance, funding_gap, funding_needed_by_month, runway_months_day90, opening_capital },
  gates: { day90: { metrics: {...value and threshold...}, result }, month12: {...} }
}
unit_economics: { hero: { price, landed_cost, landed_pct, referral, fba, ads, returns_storage, net, net_pct, korea_wholesale_gross }, per_sku: [ ... ], cost_bases: { h2, august, alibaba_low, alibaba_high } }
duty: { copper: {rate, landed_pct_medium, hero_net}, fiber: {...}, stress: {...}, fiber_list3_check: {rate} }
sensitivity: { base_metric: "Year 5 net, Medium", tornado: [ {driver, low_label, low_value, high_label, high_value, base} ] }
levers: [ ... 12 ... ]
kill_criteria: { day90: [...], month12: [...] }
plan_b: { Y1..Y3: {revenue, net} }
```

Numbers are raw (no formatting). build/fmt.js and build/fmt.py provide the one formatter every generator uses: money_m (1166340 → "$1.17M"), money_k (153020 → "$153K"), pct (0.131 → "13.1%"), neg money in parentheses.

## 5. Decks (01 EN, 02 KO)

One generator, two languages: `build/build_deck.js --lang en|ko` reading build/figures.json, build/facts.json, and build/strings/en.json or ko.json. Output ROOT/output/01_Business_Proposal_EN.pptx and 02_Business_Proposal_KO.pptx, PDFs beside them via to_pdf.

Design: dark fiber optic aesthetic. Background 0A0F1E, card surface 141C36, fiber cyan 2DE2FF, violet 8B5CF6, magenta FF4FD8 (sparingly), text F3F6FF, muted 9AA5C4, green 3DDC97 for good, amber FFB547 for caution. Motif: thin glowing fiber strands as a background image (generate a PNG with PIL: dark gradient plus a few bright thin curved lines) on title and section slides only; content slides use the plain dark background with rounded cards. No accent bars, no underlines beneath titles. Fonts: EN heading and body Calibri (Arial fallback), KO heading and body "Noto Sans KR". Title 32 to 36 pt, section header 20 to 24, body 14 to 16, captions 11. Layout LAYOUT_WIDE (13.33 x 7.5 in), 0.5 in margins. Charts native pptxgenjs (line, bar, stacked column) fed from figures.json, same palette, labels 11 pt, no chart titles duplicated by slide titles. The tornado and waterfall are bar charts built from the JSON blocks. Flowchart PNGs from output/04_Flowcharts placed as images with a short caption.

Slide list (37 slides; EN and KO identical in structure):
1 Title. 2 Executive summary (five numbers). 3 What we are deciding. 4 Section: Market and why fiber. 5 Why copper fails past 3 m at 48Gbps. 6 Who buys long HDMI (segments; placeholders for sizes). 7 Price ladder and positioning (sell the signal, not the cable). 8 Section: Product. 9 Line and roadmap. 10 Differentiators. 11 Brand architecture and the HDMI naming rule. 12 Section: Structure and compliance. 13 US entity and tax (Wyoming LLC, Form 5472, treaty, no PE). 14 Korea filings (overseas direct investment, follow ups, 세무사 items). 15 Product compliance (HDMI Adopter and ATC, FCC, Prop 65, CL3). 16 Tariff stack (duty waterfall chart). 17 Section: Launch. 18 Critical path (flowchart image). 19 Day 90 gate and kill criteria. 20 Section: Economics. 21 Hero unit economics waterfall. 22 Factory cost cross check (H2 vs August vs Alibaba). 23 Eighteen month P&L (chart plus three scenario table). 24 Five year trajectory (annual revenue by channel chart, three bands table). 25 What must be true for each band. 26 Channel expansion roadmap (flowchart image). 27 Working capital: cash by month with peak need marked (chart). 28 Funding plan. 29 Sensitivity tornado. 30 Twelve levers. 31 Marketing playbook. 32 Operating model and team. 33 Risk register summary (top 8). 34 Kill criteria and Plan B (decision tree image plus Plan B P&L). 35 The ask. 36 First 30 days. 37 Glossary.

Speaker notes on every content slide (two to four sentences). Section slides may be lighter.

QA: $VALIDATE, markitdown content check, to_pdf, to_jpg, then read every page image and fix overflow, overlap, cut text, Korean tofu boxes, label collisions. Re render after fixes.

## 6. Flowcharts (output/04_Flowcharts)

Files (mmd plus png plus pdf each): `01_launch_critical_path`, `02_order_to_cash_amazon`, `03_inbound_supply_chain`, `04_returns_and_signal_guarantee`, `05_monthly_operating_cadence`, `06_day90_scale_or_kill`, `07_five_year_channel_map`, `08_entity_money_goods_flow`. Mermaid flowchart or gantt syntax; `01` carries owner and week on every node ("Wyoming LLC and EIN. Owner Lucy, attorney. Weeks 1 to 2"); `06` uses the kill criteria numbers from 2.7; `08` shows H2 Company (Seoul) to BRAND LLC (Wyoming) to factory (China) to US customs (LLC importer of record via broker, continuous bond) to 3PL and FBA to customer, with money flows (ODI report at the foreign exchange bank, capital, factory payments, Amazon payouts, owner draw, Mark consulting fee, Form 5472) in a different edge style. Use `%%{init: {'theme':'base','themeVariables': {...}}}%%` with the deck palette on white background for print, 2400 px wide. Node labels contain no hyphens. Verify every PNG is legible at the rendered size (read it) and that no label is clipped.

## 7. Documents

- 05_Operating_Processes.docx and .pdf: US Letter, Arial 11, headings, one SOP per section with a table: Owner, Trigger, Steps (numbered), System, KPI, Failure modes. Sections: purchasing and QC; customs and duty management; FBA and 3PL inbound; listing creation and PPC management; inventory planning and reorder points; returns and warranty (5 Year Signal Guarantee, replacement shipped before the return arrives); monthly close and KPI review; compliance calendar (Form 5472, Wyoming annual report, Korean follow up reports, insurance renewal, trademark deadlines, Amazon account health). Reference the flowchart file names. Numbers from figures.json (reorder points from cover months and ramp; gate thresholds).
- 06_Compliance_Checklist.xlsx: columns Area, Item, Jurisdiction, Owner, Due (relative: "Week 1", "Before capital transfer", "Annually by April 15" style), Estimated cost, Evidence required, Status (dropdown: Not started, In progress, Done, Blocked), Notes, Basis. Areas: US entity, US tax, Korea foreign exchange, Korea tax, Amazon account, Brand and trademark, HDMI licensing, FCC, Prop 65, CL3, insurance, customs, product safety and labeling, data and privacy (placeholder). Freeze header, autofilter, a summary block with counts by status (COUNTIF formulas).
- 07_90_Day_Launch_Plan.xlsx: Gantt tab with one row per task (about 40 tasks), columns ID, Task, Owner (Lucy, Mark, attorney, broker, factory, 3PL, VA), Start week, End week, Depends on, Cost, Gate or deliverable, Basis; week columns 1 to 22 shaded by formula (conditional formatting from start and end), a Costs tab summing by owner and week (SUMIFS), a Gate tab with the Day 90 criteria and the model values beside them. Weeks match the brief's critical path.
- 08_Risk_Register.xlsx: columns ID, Risk, Category, Probability (1 to 5), Impact (1 to 5), Score (formula), Early warning signal, Mitigation, Owner, Trigger for action, Kill criterion (Yes or No), Basis. At least 25 risks covering tariff, classification, factory cost gap, HDMI licensing seizure, account suspension, IP complaint, cash, FX, returns, counterfeits, Amazon fee changes, key person (Lucy, Mark), Korea filings, US tax filings, California nexus, product liability, shipping damage, demand, competition, review manipulation accusations, listing suppression, inventory aging, 3PL failure, bank account closure, treaty position (dependent agent). A Kill Criteria tab with the Day 90 and Month 12 tables from figures.json. Heat map tab (COUNTIFS by probability and impact).
- 09_Amazon_Listing_Kit.docx: parent and child ASIN structure (one parent per line, children by length), titles (under 200 characters, category name not version number, 48Gbps and 8K allowed), five bullets each, A+ content outline (7 modules), backend keywords (under 250 bytes, no brand names of competitors), image brief for 7 slots and the video slot, FAQ (directionality, in wall CL3, projector and AV receiver setups, Xbox and PS5 120Hz, eARC, length choice), compliance notes (Prop 65 statement, FCC statement placeholder, HDMI naming). All English. Placeholders for ASINs, FNSKUs, certificate numbers.
- 10_One_Page_Summary_EN.pdf and 10_One_Page_Summary_KO.pdf: one page each, for the bank and the 세무사: who (Lucy, H2 Company, BRAND LLC Wyoming), structure diagram (small), filings in order with dates relative to the capital transfer, capital need (opening capital and funding gap from figures.json), timeline, contacts placeholder. Korean version in Noto Sans KR; keep form names and the LLC name in English.
- 11_Decision_Log.md: the five decisions from DECISIONS.md plus every additional default in this spec that a decision could change, each with: default used, what changes if it goes the other way (quantified from the model where possible: for example FX 1,300 vs 1,550 moves 18 month net by $X), who decides, by when.
- 00_README.md: index of every file with purpose, reader (Lucy, 세무사, bank, broker, Mark), order of use, how to rerun the build (scripts in build/), and the tie out result.

## 8. Tie out (build/tie_out.py)

Reads the recalculated workbook (data_only) at its default selector state and build/figures.json, re derives every figure from the workbook through figures_map.json and asserts equality within 0.5%. Extracts deck text with markitdown for both decks and asserts the formatted headline figures (18 month revenue and net for all three scenarios, breakeven month, Year 5 revenue and net Medium and High, peak cash need, funding gap, hero unit economics) appear in both decks, formatted by fmt.py. Checks docx and pdf one pagers for the capital need and funding gap strings. Runs lint_hyphens.py over output/. Prints a table and exits 1 on any mismatch. The README quotes its final result.

## 9. Glossary (shared wording; every document that defines acronyms uses these)

AOC: active optical cable. HDMI: High Definition Multimedia Interface. ASP: average selling price. COGS: cost of goods sold. FOB: free on board (the price at the factory's port). DDP: delivered duty paid. FBA: Fulfillment by Amazon. FNSKU: Fulfillment Network Stock Keeping Unit (Amazon's barcode). ASIN: Amazon Standard Identification Number. ACOS: advertising cost of sales (ad spend divided by ad attributed sales). TACOS: total advertising cost of sales (ad spend divided by total sales). PPC: pay per click advertising. 3PL: third party logistics warehouse. HTS: Harmonized Tariff Schedule. CBP: US Customs and Border Protection. MFN: most favored nation duty rate. Section 301: US trade law tariffs on China origin goods. IEEPA: International Emergency Economic Powers Act. EIN: Employer Identification Number. LLC: limited liability company. IRS: Internal Revenue Service. PE: permanent establishment. ODI: overseas direct investment report (해외직접투자 신고). 세무사: Korean licensed tax accountant. FTA: Foreign Exchange Transactions Act (외국환거래법). USPTO: United States Patent and Trademark Office. ATC: HDMI Authorized Test Center. HDMI LA: HDMI Licensing Administrator. FCC: Federal Communications Commission. Prop 65: California Proposition 65 warning label. CL3: in wall cable fire rating (Class 3). KVM: keyboard, video and mouse extension. DTC: direct to consumer website. LOC: line of credit. QC: quality control. MOQ: minimum order quantity. PO: purchase order. SKU: stock keeping unit. UPH: units per hour. KPI: key performance indicator. SOP: standard operating procedure. VA: virtual assistant. eARC: enhanced audio return channel. Gbps: gigabits per second.

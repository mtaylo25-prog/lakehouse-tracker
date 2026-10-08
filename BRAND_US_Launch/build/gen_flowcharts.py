#!/usr/bin/env python3
"""Writes the eight Mermaid flowchart sources into output/04_Flowcharts.
Content comes from BUILD_BRIEF.md sections 1 and 2 and BUILD_SPEC.md sections 2.6, 2.7, 2.8 and 6.
Render afterwards with: for f in output/04_Flowcharts/*.mmd; do render_mmd "$f"; done
"""
from pathlib import Path

ROOT = Path("/home/user/lakehouse-tracker/BRAND_US_Launch")
OUT = ROOT / "output" / "04_Flowcharts"
OUT.mkdir(parents=True, exist_ok=True)

# Deck palette on a white print background
INIT = ("%%%%{init: {'theme':'base','themeVariables': {"
        "'fontFamily':'Liberation Sans, Noto Sans KR, sans-serif','fontSize':'18px',"
        "'primaryColor':'#141C36','primaryTextColor':'#141C36','primaryBorderColor':'#2DE2FF',"
        "'lineColor':'#141C36','secondaryColor':'#EEF2FF','tertiaryColor':'#F5F7FF',"
        "'tertiaryTextColor':'#141C36','edgeLabelBackground':'#FFFFFF',"
        "'clusterBkg':'#F5F7FF','clusterBorder':'#9AA5C4','titleColor':'#141C36','textColor':'#141C36','labelTextColor':'#141C36'},"
        "'flowchart': {'nodeSpacing':%d,'rankSpacing':%d,'curve':'%s','padding':14,'htmlLabels':true,'wrappingWidth':1200}}}%%%%")

CLASSES = """
classDef step fill:#141C36,stroke:#2DE2FF,stroke-width:2px,color:#F3F6FF
classDef dec fill:#FFB547,stroke:#141C36,stroke-width:2px,color:#141C36
classDef good fill:#3DDC97,stroke:#141C36,stroke-width:2px,color:#141C36
classDef kill fill:#FF4FD8,stroke:#141C36,stroke-width:2px,color:#141C36
classDef money fill:#8B5CF6,stroke:#141C36,stroke-width:2px,color:#F3F6FF
classDef ext fill:#FFFFFF,stroke:#141C36,stroke-width:2px,color:#141C36
classDef legend fill:#FFFFFF,stroke:#9AA5C4,stroke-width:1px,color:#141C36
"""


import textwrap, re as _re
def _wrap_label(m):
    q, body = m.group(1), m.group(2)
    width = 110 if body.startswith(("Legend", "Acronyms")) else (70 if body.startswith(("Month 12 gate","KILL.","REVIEW.","SCALE.","Day 90 gate at")) else 34)
    parts = []
    for seg in body.split("<br/>"):
        parts += textwrap.wrap(seg, width) or [seg]
    return q + "<br/>".join(parts) + '"'

def wrap_labels(body):
    return _re.sub(r'(\[\"|\{\"|\(\")(.*?)"', _wrap_label, body)

def write(name, body, node=40, rank=70, curve="basis"):
    body = wrap_labels(body)
    text = INIT % (node, rank, curve) + "\n" + body.strip() + "\n" + CLASSES
    (OUT / f"{name}.mmd").write_text(text)
    print("wrote", name)


# 01 Launch critical path: every node carries owner and weeks
write("01_launch_critical_path", """
flowchart LR
  subgraph P1 ["Weeks 1 to 4: set up"]
    direction TB
    A1["1. Brand name and trademark filing (US licensed attorney)<br/>Owner: Lucy, attorney<br/>Weeks 1 to 2"]
    A2["2. Wyoming LLC and EIN (Form SS-4)<br/>Exact legal name match: state, IRS, bank, Amazon<br/>Owner: Lucy, attorney<br/>Weeks 1 to 2"]
    A5["5. Customs broker engaged, continuous bond, binding ruling filed<br/>Owner: Mark, broker<br/>Week 1 (ruling answer takes 30 to 90 days)"]
    A3["3. Korean ODI report at the foreign exchange bank<br/>Owner: Lucy, Korean tax accountant (세무사)<br/>Week 2, before any capital moves"]
    A4["4. USD bank account for BRAND LLC (Mercury or Relay)<br/>Owner: Lucy, Mark<br/>Weeks 2 to 3"]
    A4b["Capital transfer to the LLC account<br/>Owner: Lucy<br/>Week 3, only after the ODI report"]
    A6["6. Amazon Professional account, tax interview, Brand Registry<br/>Owner: Lucy, Mark<br/>Weeks 3 to 4"]
  end
  subgraph P2 ["Weeks 4 to 9: inventory and content"]
    direction TB
    A7["7. Test inventory: 5 hero SKUs, 100 to 300 units each, air DDP, QC before departure<br/>Owner: Lucy, factory, broker, 3PL<br/>Weeks 4 to 8"]
    A8["8. Listings, images, A+ content, video, PPC campaigns<br/>Owner: Mark, VA<br/>Weeks 6 to 9"]
  end
  subgraph P3 ["Weeks 10 to 22: validate"]
    direction TB
    A9["9. 90 day validation: conversion, TACOS, return rate, contribution per SKU<br/>Owner: Mark, Lucy<br/>Weeks 10 to 22"]
    G{"Day 90 gate decision<br/>Owner: Lucy decides, Mark prepares<br/>Week 22"}
  end
  subgraph P4 ["Week 22 onward"]
    direction TB
    A10["10. Scale winners on ocean freight, delist losers<br/>Owner: Lucy, Mark, 3PL<br/>Week 22 onward"]
    PB["Plan B: B2B installer distribution and Korea wholesale<br/>Owner: Lucy, Mark<br/>Week 22 onward"]
  end
  LEG["Legend. LLC: limited liability company. EIN: Employer Identification Number. ODI: overseas direct investment report. SKU: stock keeping unit. DDP: delivered duty paid. QC: quality control. 3PL: third party logistics warehouse. VA: virtual assistant. PPC: pay per click ads. TACOS: total advertising cost of sales. B2B: business to business."]

  A1 --> A6
  A2 --> A4
  A2 --> A6
  A3 --> A4b
  A4 --> A4b
  A4b --> A7
  A5 --> A7
  A6 --> A8
  A7 --> A8
  A8 --> A9
  A9 --> G
  G -->|"Pass"| A10
  G -->|"Kill"| PB
  A5 -.->|"Ruling result sets the duty rate for ocean lots"| A10

  class A1,A2,A3,A4,A4b,A5,A6,A7,A8,A9 step
  class G dec
  class A10 good
  class PB kill
  class LEG legend
""", node=30, rank=60)

# 02 Order to cash on Amazon
write("02_order_to_cash_amazon", """
flowchart TD
  subgraph R1 ["1. Order and delivery"]
    direction LR
    S["Shopper session from Amazon search or a PPC ad"] --> C["Customer places the order on Amazon.com"]
    C --> PAY["Amazon charges the card and collects marketplace sales tax"]
    PAY --> FBA["FBA picks, packs and ships from an Amazon warehouse"]
    FBA --> DEL["Delivery to the customer"]
    DEL --> OK{"Kept or returned?"}
  end
  subgraph R2 ["2. Settlement and payout (Amazon side)"]
    direction LR
    RET["Returned in the window: refund issued, unit back to FBA (see chart 04)"] --> SET["Amazon settlement: sale price minus referral 15%, FBA fee, storage, ad spend, refunds"]
    SET --> LAG["Disbursement 14 days after the sale. The payout lag is funded by the business"]
  end
  subgraph R3 ["3. Uses of cash (BRAND LLC side)"]
    direction LR
    BANK["BRAND LLC USD bank account (Mercury or Relay)"] --> U1["Factory PO deposits and balances"]
    BANK --> U2["Ad spend (PPC budget buys the next sessions), software, insurance, 3PL and broker invoices"]
    BANK --> U3["Mark consulting fee under the written agreement"]
    BANK --> U4["Retained earnings first, then owner draw to Lucy, reported on Form 5472"]
    BANK --> CLOSE["Monthly close and KPI review (chart 05)<br/>Owner: Mark, VA"]
  end
  R1 -->|"Kept: straight to settlement. Returned: refund first"| R2
  R2 -->|"Payout lands in the LLC account"| R3
  LEG["Legend. PPC: pay per click advertising. FBA: Fulfillment by Amazon. PO: purchase order. 3PL: third party logistics warehouse. KPI: key performance indicator. VA: virtual assistant. Referral 15% and the 14 day payout lag are August model assumptions."]

  class S,C,PAY,FBA,DEL,RET,SET,LAG,CLOSE step
  class OK dec
  class BANK money
  class U1,U2,U3,U4 ext
  class LEG legend
""")

# 03 Inbound supply chain
write("03_inbound_supply_chain", """
flowchart TD
  subgraph R1 ["1. China: order, production, QC"]
    direction LR
    FC["Forecast and reorder point: 2.5 months of forward units (Medium)<br/>Owner: Mark"] --> PO["PO to the factory, 30% deposit paid<br/>Owner: Lucy"]
    PO --> PROD["Production. FNSKU labels and Prop 65 label applied at the factory<br/>Owner: factory"]
    PROD --> DOCS["Documents on file before shipment: HDMI Adopter status and ATC test report, FCC test report, CL3 document<br/>Owner: Lucy"]
    DOCS --> QC["Third party QC inspection before departure<br/>Owner: Lucy"]
    QC --> QCD{"QC pass?"}
    QCD -->|"Fail"| RW["Rework or reject the lot. Hold the 70% balance"]
    RW --> QC
    QCD -->|"Pass"| BAL["70% balance paid at shipment<br/>Owner: Lucy"]
  end
  subgraph R2 ["2. Freight"]
    direction LR
    MODE{"SKU proven at the Day 90 gate?"} -->|"No"| AIR["Air freight DDP, about 1 month lead time. First lots 100 to 300 units per SKU"]
    MODE -->|"Yes"| OCEAN["Ocean LCL plus drayage, about 2 months lead time. Proven SKUs only"]
  end
  subgraph R3 ["3. United States: customs and warehouses"]
    direction LR
    ENTRY["Formal customs entry filed by the broker. BRAND LLC is importer of record. Continuous bond. HTS 8544.42 or 8544.70 per the binding ruling. Duty paid. No de minimis<br/>Owner: broker, Mark"] --> TPL["3PL receives bulk inventory, counts and inspects<br/>Owner: 3PL"]
    TPL --> FBAIN["FBA inbound shipment plan, placement and prep, carton labels<br/>Owner: Mark, 3PL"]
    FBAIN --> LIVE["Inventory live on Amazon, listing active. Sell through feeds the next forecast in step 1"]
    TPL --> MCF["3PL also ships eBay, Walmart and B2B orders from Year 2"]
  end
  R1 -->|"Goods released by the factory"| R2
  R2 -->|"Arrival at the US airport or port"| R3
  LEG["Legend. PO: purchase order. FNSKU: Amazon barcode. ATC: HDMI Authorized Test Center. FCC: Federal Communications Commission. CL3: in wall cable rating. QC: quality control. DDP: delivered duty paid. LCL: less than container load. HTS: Harmonized Tariff Schedule. 3PL: third party logistics warehouse. FBA: Fulfillment by Amazon. B2B: business to business."]

  class FC,PO,PROD,DOCS,QC,BAL,AIR,OCEAN,ENTRY,TPL,FBAIN step
  class QCD,MODE dec
  class RW kill
  class LIVE good
  class MCF ext
  class LEG legend
""")

# 04 Returns and the 5 Year Signal Guarantee
write("04_returns_and_signal_guarantee", """
flowchart TD
  subgraph R1 ["1. Triage"]
    direction LR
    ISSUE["Customer reports no signal, flicker or a fault"] --> SELF["Self help: QR code setup video. Direction check: SOURCE end at the player, DISPLAY end at the TV<br/>Owner: VA"]
    SELF --> FIX{"Resolved?"}
    FIX -->|"Yes"| CLOSE["Close the case. Review follow up through Brand Registry"]
    FIX -->|"No"| WIN{"Inside the Amazon return window?"}
  end
  subgraph R2 ["2. Amazon return path"]
    direction LR
    AR["Amazon return: refund to the customer, unit back to FBA, graded by Amazon"] --> RES["Returns reserve absorbs the cost. Base 3%, target 2% with the direction proof design"]
  end
  subgraph R3 ["3. 5 Year Signal Guarantee path"]
    direction LR
    GUAR["Guarantee claim through the registration list (Spec Match finder, trade up registration)"] --> VER["Verify registration, serial and symptoms<br/>Owner: Mark, VA"]
    VER --> SHIP["Replacement shipped first, from the 3PL or FBA multi channel fulfillment, before the faulty unit comes back"]
    SHIP --> LABEL["Prepaid return label sent to the customer"]
  end
  subgraph R4 ["4. Receiving, root cause, feedback"]
    direction LR
    RCV["Faulty unit received at the 3PL and inspected<br/>Owner: 3PL"] --> ROOT["Root cause log: direction error, connector damage, dead head electronics, in wall pull damage"]
    ROOT --> CLAIM["Factory warranty credit claim on confirmed defects<br/>Owner: Lucy"]
    ROOT --> KPI["Return rate KPI each month. Day 90: pass at or below 4%, kill above 7%. Month 12: pass at or below 3%, kill above 5%<br/>Owner: Mark"]
    KPI --> DESIGN["Design feedback to the factory: color coded ends, printed arrows, detachable head<br/>Owner: Lucy"]
  end
  R1 -->|"Inside the window: yes"| R2
  R1 -->|"Inside the window: no"| R3
  R2 -->|"Unsellable units removed to the 3PL"| R4
  R3 -->|"Faulty unit comes back"| R4
  LEG["Legend. QR: quick response code. VA: virtual assistant. FBA: Fulfillment by Amazon. 3PL: third party logistics warehouse. KPI: key performance indicator. Thresholds come from the kill criteria in the risk register."]

  class ISSUE,SELF,AR,GUAR,VER,SHIP,LABEL,RCV,ROOT,CLAIM,DESIGN step
  class FIX,WIN dec
  class CLOSE good
  class RES,KPI ext
  class LEG legend
""")

# 05 Monthly operating cadence
write("05_monthly_operating_cadence", """
flowchart TD
  subgraph W1 ["Week 1: close and review"]
    direction LR
    C1["Days 1 to 3: reconcile Amazon settlements, 3PL, broker and ad invoices<br/>Owner: VA, Mark"] --> C2["Days 3 to 5: monthly close in USD with a KRW view at the FX input<br/>Owner: Mark, Korean tax accountant (세무사)"]
    C2 --> C3["Day 5: KPI review call. Revenue, units, conversion, TACOS, return rate, contribution per unit, cash balance, runway<br/>Owner: Lucy, Mark"]
  end
  subgraph W2 ["Week 2: inventory and cash"]
    direction LR
    I1["Forward forecast, cover target and reorder point by SKU (chart 03)<br/>Owner: Mark"] --> I2["PO decision and 30% deposit. Air for new SKUs, ocean for proven SKUs<br/>Owner: Lucy"]
    I2 --> I3["Cash forecast 3 months ahead, funding gap check against the line of credit plan<br/>Owner: Mark"]
  end
  subgraph W3 ["Week 3: listings and ads"]
    direction LR
    P1["PPC: keyword pruning, negative keywords, bid and budget review<br/>Owner: Mark, VA"] --> P2["Listings: Q and A, review follow up, Vine, insert card results, A+ updates<br/>Owner: VA"]
    P2 --> P3["Price ladder and coupon test (list vs sale price)<br/>Owner: Mark"]
  end
  subgraph W4 ["Week 4: compliance and plan"]
    direction LR
    K1["Account health, IP complaints, listing suppressions, insurance threshold ($10K sales month)<br/>Owner: Mark"] --> K2["Filing calendar: Form 5472, Wyoming annual report, Korean follow up reports, trademark dates<br/>Owner: Lucy, attorney, 세무사"]
    K2 --> K3["Next month plan and owner actions agreed, then back to Week 1<br/>Owner: Lucy, Mark"]
  end
  W1 --> W2
  W2 --> W3
  W3 --> W4
  LEG["Legend. KPI: key performance indicator. TACOS: total advertising cost of sales. FX: foreign exchange rate. SKU: stock keeping unit. PO: purchase order. PPC: pay per click advertising. VA: virtual assistant. IP: intellectual property. 3PL: third party logistics warehouse."]

  class C1,C2,I1,I2,I3,P1,P2,P3,K1,K2 step
  class C3,K3 dec
  class LEG legend
""")

# 06 Day 90 scale or kill decision tree with the exact thresholds from spec 2.7
write("06_day90_scale_or_kill", """
flowchart TD
  START["Day 90 gate at the end of Month 3<br/>Owner: Lucy decides, Mark prepares the numbers<br/>Step 1: check the five kill triggers. Step 2: check the six pass tests"]
  START --> K1["Kill trigger 1<br/>Hero SKU conversion below 5%?"]
  START --> K2["Kill trigger 2<br/>TACOS above 40% in Month 3?"]
  START --> K3["Kill trigger 3<br/>Return rate above 7%?"]
  START --> K4["Kill trigger 4<br/>Hero contribution before advertising below $8 per unit?"]
  START --> K5["Kill trigger 5<br/>Cash runway below 2 months?"]
  K1 -->|"Yes"| KILL
  K2 -->|"Yes"| KILL
  K3 -->|"Yes"| KILL
  K4 -->|"Yes"| KILL
  K5 -->|"Yes"| KILL
  K1 -->|"No"| CLEAR
  K2 -->|"No"| CLEAR
  K3 -->|"No"| CLEAR
  K4 -->|"No"| CLEAR
  K5 -->|"No"| CLEAR
  KILL["KILL. Any one trigger kills. No further capital. Sell down inventory, delist losers, move to Plan B: B2B installer distribution plus Korea wholesale"]
  CLEAR["No kill trigger hit. Now all six pass tests must pass"]
  CLEAR --> P1["Pass test 1<br/>Conversion at or above 8% on at least 3 of the 5 launch SKUs?"]
  CLEAR --> P2["Pass test 2<br/>TACOS at or below 30% in Month 3 and falling month over month?"]
  CLEAR --> P3["Pass test 3<br/>Return rate at or below 4%?"]
  CLEAR --> P4["Pass test 4<br/>Contribution before advertising at or above $11 per unit on at least 3 of 5 SKUs?"]
  CLEAR --> P5["Pass test 5<br/>Cash runway at or above 4 months after funding the next PO?"]
  CLEAR --> P6["Pass test 6<br/>No suppressed listing and no IP complaint open more than 14 days?"]
  P1 -->|"No"| HOLD
  P2 -->|"No"| HOLD
  P3 -->|"No"| HOLD
  P4 -->|"No"| HOLD
  P5 -->|"No"| HOLD
  P6 -->|"No"| HOLD
  P1 -->|"Yes"| ALL
  P2 -->|"Yes"| ALL
  P3 -->|"Yes"| ALL
  P4 -->|"Yes"| ALL
  P5 -->|"Yes"| ALL
  P6 -->|"Yes"| ALL
  HOLD["REVIEW. Fix the failed test and retest in 30 days. No new capital beyond the committed PO. One retest only, then pass or kill"]
  ALL["All six pass"]
  ALL --> SCALE["SCALE. Ocean freight PO on proven SKUs, delist losers, keep air for new SKUs, next review at Month 12"]
  HOLD -.->|"Retest once"| START
  SCALE --> M12["Month 12 gate. Pass needs all of: revenue at or above $70K per month, cumulative net at or above $0, TACOS at or below 15%, return rate at or below 3%, operating cash flow positive two months in a row.<br/>Kill on any one of: revenue below $50K per month, cumulative net below ($25K), TACOS above 25%, return rate above 5%, three months of negative operating cash flow with a funding gap above $50K"]
  LEG["Legend. TACOS: total advertising cost of sales (ad spend divided by total revenue). SKU: stock keeping unit. PO: purchase order. IP: intellectual property. Contribution: price minus landed cost, referral, FBA fee and returns, before advertising. Hero SKU: 8K 10M. Thresholds from the kill criteria in the risk register."]

  class START,CLEAR step
  class K1,K2,K3,K4,K5,P1,P2,P3,P4,P5,P6 dec
  class KILL kill
  class SCALE,ALL good
  class HOLD ext
  class M12 step
  class LEG legend
""", node=24, rank=70, curve="linear")

# 07 Five year channel expansion map
write("07_five_year_channel_map", """
flowchart TD
  subgraph Y1 ["Year 1: Amazon US only"]
    direction LR
    A1["Amazon US. 5 hero SKUs under one parent ASIN. FBA. Brand Registry"] --> G1{"Day 90 gate (chart 06)"}
    G1 -->|"Pass"| G2{"Month 12 gate"}
    G1 -->|"Kill"| PB["Plan B if Amazon fails: B2B installer distribution plus Korea wholesale. Own small P&L"]
    G2 -->|"Kill"| PB
    A2["Korea wholesale continues at H2 Company in every year"] --> PB
  end
  subgraph Y2 ["Year 2: overflow and installers"]
    direction LR
    B1["Amazon US: full length ladder, ocean freight on proven SKUs"]
    B2["eBay and Walmart Marketplace as overflow. Multi channel fulfillment, near zero extra effort"]
    B3["B2B installer and AV integrator program. 3PL ships, price 60% of retail"]
    B1 ~~~ B2 ~~~ B3
  end
  subgraph Y3 ["Year 3: own channel and new products"]
    direction LR
    C1["DTC site fed by the off Amazon email list (Spec Match finder, trade up registration)"]
    C2["Roadmap SKUs: HDMI 2.2 16K, DP 2.1, USB4 AOC"]
    C1 ~~~ C2
  end
  subgraph Y4 ["Year 4: Canada and the pro line"]
    direction LR
    D1["Amazon Canada, stock transferred from the US"]
    D2["Pro line: CL3, plenum, conduit pull, long run fiber KVM. Sold 60% B2B, 40% Amazon"]
    D1 ~~~ D2
  end
  subgraph Y5 ["Year 5: scale"]
    direction LR
    E1["Scale every channel. Medium band target: a multi million dollar business with 10% to 15% net (see the model)"]
    E2["Temu local seller program: optional, only after the Amazon price floor holds. Not in the base case"]
    E1 ~~~ E2
  end
  Y1 -->|"Month 12 gate passed"| Y2
  Y2 -->|"Amazon US and B2B profitable"| Y3
  Y3 -->|"DTC list and roadmap SKUs live"| Y4
  Y4 --> Y5
  LEG["Legend. ASIN: Amazon Standard Identification Number. FBA: Fulfillment by Amazon. B2B: business to business. AV: audio video. 3PL: third party logistics warehouse. DTC: direct to consumer website. DP: DisplayPort. AOC: active optical cable. CL3: in wall cable rating. KVM: keyboard, video and mouse extension. P&L: profit and loss."]

  class A1,A2,B1,B2,B3,C1,C2,D1,D2 step
  class G1,G2 dec
  class E1 good
  class E2 ext
  class PB kill
  class LEG legend
""")

# 08 Entity, money and goods flow. Edge styles: goods thick navy, money dotted violet, filings dashed amber.
edges = []  # (src, arrow, label, dst, kind)
def e(src, dst, kind, label=""):
    edges.append((src, dst, kind, label))

# goods
e("FACT", "CBP", "goods", "Air DDP first lots, ocean LCL on proven SKUs")
e("CBP", "TPL", "goods", "Released entry, bulk to the 3PL")
e("TPL", "FBA", "goods", "FBA inbound shipments")
e("FBA", "CUST", "goods", "Amazon delivers the order")
e("TPL", "B2B", "goods", "B2B and overflow orders from Year 2")
e("FACT", "H2", "goods", "Korea wholesale lots")
e("H2", "DIO", "goods", "Korea wholesale sales (August deal pattern)")
# money
e("LUCY", "FXB", "money", "Capital to the US, after the ODI report")
e("FXB", "BANK", "money", "Capital received by the LLC")
e("BANK", "FACT", "money", "PO deposit 30%, balance 70% at shipment")
e("BANK", "BROKER", "money", "Duty, entry fee, continuous bond")
e("AMZ", "BANK", "money", "Payouts 14 days after sale, net of fees")
e("BANK", "TPL", "money", "Storage, pick pack, prep")
e("BANK", "MARK", "money", "Consulting fee under the written agreement")
e("BANK", "LUCY", "money", "Owner draw after retained earnings, reported on Form 5472")
e("CUST", "AMZ", "money", "Pays Amazon, sales tax collected by Amazon")
# filings
e("LUCY", "FXB", "filing", "ODI report before capital moves, then follow up reports")
e("LUCY", "SEMU", "filing", "Overseas financial account check, incorporate H2 question")
e("LLC", "IRS", "filing", "Form 5472 with pro forma 1120 each year")
e("LLC", "WY", "filing", "Annual report")
e("LLC", "USPTO", "filing", "Trademark via a US licensed attorney")
e("BROKER", "CBP", "filing", "Binding ruling request, formal entries, LLC as importer of record")
e("LLC", "AMZ", "filing", "Seller account, tax interview, Brand Registry")
e("LUCY", "LLC", "filing", "100% owner, single member LLC")
# legend
e("LG1", "LG2", "goods", "Goods")
e("LM1", "LM2", "money", "Money")
e("LF1", "LF2", "filing", "Filings and agreements")

arrow = {"goods": "==>", "money": "-.->", "filing": "-->"}
lines = ["flowchart TD"]
lines += [
    '  LUCY["Lucy (Hyejin Yoo), owner, Seoul"]',
    '  H2["H2 Company, Seoul. Korean sole proprietorship (599-64-00496)"]',
    '  DIO["Korean wholesale customers (Dioptek)"]',
    '  FXB["Foreign exchange bank, Seoul: ODI report and follow up reports"]',
    '  SEMU["Korean tax accountant (세무사)"]',
    '  FACT["Factory, China (HDMI Adopter, ATC and FCC reports)"]',
    '  LLC["BRAND LLC, Wyoming single member LLC, owned 100% by Lucy. No US office, no employees, no signing agent"]',
    '  BANK["LLC USD bank account, US (Mercury or Relay)"]',
    '  IRS["IRS"]',
    '  WY["Wyoming Secretary of State"]',
    '  USPTO["USPTO"]',
    '  BROKER["Customs broker, US. Continuous bond"]',
    '  CBP["US Customs (CBP): formal entry, duty under HTS 8544.42 or 8544.70"]',
    '  TPL["3PL warehouse, US"]',
    '  FBA["Amazon FBA warehouses, US (no permanent establishment under the treaty)"]',
    '  AMZ["Amazon marketplace"]',
    '  CUST["US customer"]',
    '  B2B["B2B installers, eBay and Walmart buyers, US"]',
    '  MARK["Mark Taylor, Memphis: consulting agreement, no equity, signs no contracts"]',
    '  subgraph LEGEND ["Legend"]',
    '    direction TB',
    '    LG1["Start"]',
    '    LG2["Goods: thick navy line"]',
    '    LM1["Start"]',
    '    LM2["Money: dotted violet line"]',
    '    LF1["Start"]',
    '    LF2["Filings and agreements: dashed amber line"]',
    '  end',
    '  LEG["Acronyms. LLC: limited liability company. ODI: overseas direct investment report. IRS: Internal Revenue Service. USPTO: US Patent and Trademark Office. CBP: US Customs and Border Protection. HTS: Harmonized Tariff Schedule. 3PL: third party logistics warehouse. FBA: Fulfillment by Amazon. B2B: business to business. DDP: delivered duty paid. LCL: less than container load. PO: purchase order."]',
]
for src, dst, kind, label in edges:
    lab = f'|"{label}"|' if label else ""
    lines.append(f"  {src} {arrow[kind]}{lab} {dst}")
style = {
    "goods": "stroke:#141C36,stroke-width:4px",
    "money": "stroke:#8B5CF6,stroke-width:3px",
    "filing": "stroke:#D98F1C,stroke-width:3px,stroke-dasharray:9 6",
}
for kind in ("goods", "money", "filing"):
    idx = [str(i) for i, ed in enumerate(edges) if ed[2] == kind]
    lines.append(f"  linkStyle {','.join(idx)} {style[kind]}")
lines += [
    "  class LUCY,H2,DIO,FACT,LLC,TPL,FBA,AMZ,CUST,B2B,DIO step",
    "  class BANK money",
    "  class FXB,SEMU,IRS,WY,USPTO,BROKER,CBP,MARK ext",
    "  class LG1,LM1,LF1,LG2,LM2,LF2,LEG legend",
]
write("08_entity_money_goods_flow", "\n".join(lines), node=30, rank=80)

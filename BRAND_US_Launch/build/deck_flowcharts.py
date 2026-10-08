#!/usr/bin/env python3
"""Deck specific renders of flowcharts 01, 06 and 07 for the proposal decks (EN and KO share them).

The delivered flowcharts in output/04_Flowcharts are print renders at 2400 px wide. Placed on a slide they
scale down and their 18 px labels fall to about 4 to 6 pt. The deck needs variants laid out for the slide
slot (wide and short), with terse labels and no legend (the glossary slide carries the acronyms), so the
labels read at 9 pt or more on the page. Content is identical to the delivered charts: same steps, owners,
weeks and thresholds, taken from the same sources (BUILD_BRIEF section 1 and BUILD_SPEC 2.7 and 6).

Writes build/assets/deck_flow/<name>.mmd, .svg, .png and dims.json (pixel size, natural size, font px).
Run: source build/env.sh && python3 build/deck_flowcharts.py
"""
import json
import re
import subprocess
import textwrap
from pathlib import Path

ROOT = Path("/home/user/lakehouse-tracker/BRAND_US_Launch")
BUILD = ROOT / "build"
OUT = BUILD / "assets" / "deck_flow"
OUT.mkdir(parents=True, exist_ok=True)
MMDC_CFG = BUILD / "puppeteer.json"
FONT_PX = 30
SIZE = 2600  # PNG max dimension

INIT = ("%%%%{init: {'theme':'base','themeVariables': {"
        "'fontFamily':'Liberation Sans, Noto Sans KR, sans-serif','fontSize':'%dpx',"
        "'primaryColor':'#141C36','primaryTextColor':'#141C36','primaryBorderColor':'#2DE2FF',"
        "'lineColor':'#141C36','secondaryColor':'#EEF2FF','tertiaryColor':'#F5F7FF',"
        "'tertiaryTextColor':'#141C36','edgeLabelBackground':'#FFFFFF',"
        "'clusterBkg':'#F5F7FF','clusterBorder':'#9AA5C4','titleColor':'#141C36','textColor':'#141C36','labelTextColor':'#141C36'},"
        "'flowchart': {'nodeSpacing':%d,'rankSpacing':%d,'curve':'%s','padding':%d,'htmlLabels':true,'wrappingWidth':2000}}}%%%%")

CLASSES = """
classDef step fill:#141C36,stroke:#2DE2FF,stroke-width:3px,color:#F3F6FF
classDef dec fill:#FFB547,stroke:#141C36,stroke-width:3px,color:#141C36
classDef good fill:#3DDC97,stroke:#141C36,stroke-width:3px,color:#141C36
classDef kill fill:#FF4FD8,stroke:#141C36,stroke-width:3px,color:#141C36
classDef money fill:#8B5CF6,stroke:#141C36,stroke-width:3px,color:#F3F6FF
classDef ext fill:#FFFFFF,stroke:#141C36,stroke-width:3px,color:#141C36
"""

# Deck config: same theme as build/mermaid_config.json with the larger font and a wide wrapping width so the
# explicit <br/> breaks below decide the line length.
DECK_CFG = OUT / "mermaid_config_deck.json"
base_cfg = json.loads((BUILD / "mermaid_config.json").read_text())
base_cfg["themeVariables"]["fontSize"] = f"{FONT_PX}px"
base_cfg["flowchart"]["wrappingWidth"] = 2000
base_cfg["themeCSS"] += " .cluster-label span { font-size: %dpx; }" % (FONT_PX + 2)
DECK_CFG.write_text(json.dumps(base_cfg, indent=1))


def wrap_labels(body, width):
    def _w(m):
        q, text = m.group(1), m.group(2)
        w = int(width * 1.6) if text.startswith(("KILL", "SCALE", "HOLD", "REVIEW", "Plan B")) else width
        parts = []
        for seg in text.split("<br/>"):
            parts += textwrap.wrap(seg, w) or [seg]
        return q + "<br/>".join(parts) + '"'
    out = []
    for line in body.splitlines():
        out.append(line if line.strip().startswith("subgraph") else re.sub(r'(\[\"|\{\"|\(\")(.*?)"', _w, line))
    return "\n".join(out)


def render(name, body, width=24, node=30, rank=60, curve="basis", padding=14):
    text = INIT % (FONT_PX, node, rank, curve, padding) + "\n" + wrap_labels(body, width).strip() + "\n" + CLASSES
    labels = re.findall(r'(?:\[\"|\{\"|\(\"|\|\")(.*?)"', body)
    bad = [l for l in labels if re.search(r"[-\u2013\u2014]", l)]
    if bad:
        raise SystemExit("hyphen in label text: " + " | ".join(bad))
    mmd = OUT / f"{name}.mmd"
    mmd.write_text(text)
    svg = OUT / f"{name}.svg"
    png = OUT / f"{name}.png"
    common = ["mmdc", "-q", "-p", str(MMDC_CFG), "-c", str(DECK_CFG), "-i", str(mmd), "-b", "white", "-s", "1"]
    subprocess.run(common + ["-o", str(svg)], check=True)
    subprocess.run(common + ["-o", str(png), "--size", str(SIZE)], check=True)
    s = svg.read_text()
    vb = re.search(r'viewBox="([-\d.]+) ([-\d.]+) ([\d.]+) ([\d.]+)"', s)
    nat_w, nat_h = float(vb.group(3)), float(vb.group(4))
    from PIL import Image
    im = Image.open(png)
    w, h = im.size
    # If the PNG is taller than the slot it is cropped later by the deck; record natural dims so the deck can
    # compute the label size: pt = FONT_PX * (placed_in / nat_w_in_px) * 72 where placed width maps to nat_w.
    info = {"w": w, "h": h, "natural_w": nat_w, "natural_h": nat_h, "font_px": FONT_PX}
    print(name, info, "label pt at 12.13 in:", round(FONT_PX * 12.13 / nat_w * 72, 1), "aspect", round(nat_w / nat_h, 2))
    return info


dims = {}

# 01 Launch critical path, deck variant: same ten steps, owners and weeks; terse labels; no legend; the ruling
# note sits in step 10 instead of on an edge label (edge labels widen the layout).
dims["01"] = render("01_launch_critical_path_deck", """
flowchart LR
  subgraph P1 ["Weeks 1 to 4: set up"]
    A1["1. Brand name and trademark<br/>Lucy, attorney<br/>Weeks 1 to 2"]
    A2["2. Wyoming LLC and EIN, exact legal name match<br/>Lucy, attorney<br/>Weeks 1 to 2"]
    A5["5. Broker, bond, binding ruling filed<br/>Mark, broker<br/>Week 1"]
    A3["3. Korean ODI report at the bank<br/>Lucy, tax accountant<br/>Week 2, before capital moves"]
    A4["4. USD bank account<br/>Lucy, Mark<br/>Weeks 2 to 3"]
    A6["6. Amazon account and Brand Registry<br/>Lucy, Mark<br/>Weeks 3 to 4"]
    CAP["Capital transfer<br/>Lucy<br/>Week 3, after the ODI report"]
    A2 --> A4 --> CAP
    A3 --> CAP
    A1 --> A6
    A2 --> A6
    A5 ~~~ A4
  end
  subgraph P2 ["Weeks 4 to 9: inventory<br/>and content"]
    B7["7. Test inventory: 5 hero SKUs, 100 to 300 units each, air DDP, QC before departure<br/>Lucy, factory, broker, 3PL<br/>Weeks 4 to 8"]
    B8["8. Listings, images, A+ content, video, PPC<br/>Mark, VA<br/>Weeks 6 to 9"]
    B7 --> B8
  end
  subgraph P3 ["Weeks 10 to 22: validate"]
    C9["9. 90 day validation: conversion, TACOS, return rate, contribution per SKU<br/>Mark, Lucy<br/>Weeks 10 to 22"]
    G{{"Day 90 gate<br/>Lucy decides<br/>Week 22"}}
    C9 --> G
  end
  subgraph P4 ["Week 22 onward"]
    S10["10. Scale winners on ocean freight at the ruled duty rate, delist losers<br/>Lucy, Mark, 3PL"]
    PB["Plan B: B2B installers and Korea wholesale<br/>Lucy, Mark"]
  end
  CAP --> B7
  A6 --> B8
  B8 --> C9
  G -->|"Pass"| S10
  G -->|"Kill"| PB
  class A1,A2,A3,A4,A5,A6,B7,B8,C9,S10 step
  class CAP money
  class G dec
  class PB kill
""", width=22, node=20, rank=44, padding=10)

# 06 Day 90 scale or kill, deck variant: five kill triggers and six pass tests in two columns each, edges from
# the groups instead of from every test, Month 12 gate folded into the SCALE node (its thresholds are on the
# gate slide and in the delivered chart 06).
dims["06"] = render("06_day90_scale_or_kill_deck", """
flowchart LR
  subgraph KT ["Day 90 gate at the end of Month 3<br/>Lucy decides, Mark prepares the numbers<br/>Step 1: five kill triggers, any one kills"]
    direction LR
    K1["1. Hero SKU conversion below 5%?"]
    K2["2. TACOS above 40% in Month 3?"]
    K3["3. Return rate above 7%?"]
    K4["4. Hero contribution before advertising below $8 per unit?"]
    K5["5. Cash runway below 2 months?"]
    K1 ~~~ K2 ~~~ K3
    K4 ~~~ K5
  end
  subgraph PT ["Step 2: six pass tests<br/>All must pass"]
    direction LR
    P1["1. Conversion at or above 8% on 3 of 5 launch SKUs?"]
    P2["2. TACOS at or below 30% in Month 3 and falling month over month?"]
    P3["3. Return rate at or below 4%?"]
    P4["4. Contribution before advertising at or above $11 per unit on 3 of 5 SKUs?"]
    P5["5. Cash runway at or above 4 months after funding the next PO?"]
    P6["6. No suppressed listing and no IP complaint open more than 14 days?"]
    P1 ~~~ P2 ~~~ P3
    P4 ~~~ P5 ~~~ P6
  end
  KT -->|"Any yes"| KILL["KILL. No further capital. Sell down inventory, delist losers, move to Plan B: B2B installers plus Korea wholesale"]
  KT -->|"All no"| PT
  PT -->|"All yes"| SCALE["SCALE. Ocean freight PO on proven SKUs, delist losers, keep air for new SKUs. Next review: Month 12 gate"]
  PT -->|"Any no"| HOLD["REVIEW. Fix the failed test, retest in 30 days. No new capital beyond the committed PO. One retest only, then pass or kill"]
  class K1,K2,K3,K4,K5,P1,P2,P3,P4,P5,P6 dec
  class KILL kill
  class SCALE good
  class HOLD ext
""", width=22, node=12, rank=50, curve="linear", padding=10)

# 07 Five year channel map, deck variant: years left to right, channels stacked, gates and Plan B in Year 1,
# the gate between years written in the year title instead of on an edge label.
dims["07"] = render("07_five_year_channel_map_deck", """
flowchart LR
  subgraph Y1 ["Year 1: Amazon US only"]
    direction TB
    A1["Amazon US. 5 hero SKUs under one parent ASIN. FBA. Brand Registry"]
    G1{{"Day 90 gate"}}
    G2{{"Month 12 gate"}}
    PB["Plan B if Amazon fails: B2B installer distribution plus Korea wholesale"]
    A1 --> G1
    G1 -->|"Pass"| G2
    G1 -->|"Kill"| PB
    G2 -->|"Kill"| PB
  end
  subgraph Y2 ["Year 2: overflow<br/>and installers"]
    direction TB
    B1["Amazon US: full length ladder, ocean freight on proven SKUs"]
    B2["eBay and Walmart as overflow. Multi channel fulfillment, near zero extra effort"]
    B3["B2B installer and AV integrator program. 3PL ships, price 60% of retail"]
    B1 ~~~ B2 ~~~ B3
  end
  subgraph Y3 ["Year 3: own channel<br/>and new products"]
    direction TB
    C1["DTC site fed by the off Amazon email list (Spec Match finder, trade up registration)"]
    C2["Roadmap SKUs: HDMI 2.2 16K, DP 2.1, USB4 AOC"]
    C1 ~~~ C2
  end
  subgraph Y4 ["Year 4: Canada<br/>and the pro line"]
    direction TB
    D1["Amazon Canada, stock transferred from the US"]
    D2["Pro line: CL3, plenum, conduit pull, long run fiber KVM. Sold 60% B2B, 40% Amazon"]
    D1 ~~~ D2
  end
  subgraph Y5 ["Year 5: scale"]
    direction TB
    E1["Scale every channel. Medium band target: 10% to 15% net (see the model)"]
    E2["Temu local seller program: optional, only after the Amazon price floor holds. Not in the base case"]
    E1 ~~~ E2
  end
  Y1 --> Y2 --> Y3 --> Y4 --> Y5
  class A1,B1,B2,B3,C1,C2,D1,D2 step
  class G1,G2 dec
  class E1 good
  class E2 ext
  class PB kill
""", width=25, node=20, rank=50, padding=10)

(OUT / "dims.json").write_text(json.dumps(dims, indent=1))
print("wrote", OUT / "dims.json")

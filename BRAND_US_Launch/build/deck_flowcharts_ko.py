#!/usr/bin/env python3
"""Korean renders of the three deck flowcharts (01, 06, 07) for 02_Business_Proposal_KO.pptx.

Same layout, steps, owners, weeks and thresholds as build/deck_flowcharts.py (which the EN deck uses); only the
label wording is Korean. Numbers and form names are unchanged. Hangul glyphs are about twice as wide as Latin
glyphs, so the wrap width is lower. Labels contain no hyphens or dashes.

Writes build/assets/deck_flow/<name>_deck_ko.mmd, .svg, .png and dims_ko.json.
Run: source build/env.sh && python3 build/deck_flowcharts_ko.py
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
SIZE = 2600

INIT = ("%%%%{init: {'theme':'base','themeVariables': {"
        "'fontFamily':'Noto Sans KR, Liberation Sans, sans-serif','fontSize':'%dpx',"
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

DECK_CFG = OUT / "mermaid_config_deck_ko.json"
base_cfg = json.loads((BUILD / "mermaid_config.json").read_text())
base_cfg["fontFamily"] = "Noto Sans KR, Liberation Sans, sans-serif"
base_cfg["themeVariables"]["fontFamily"] = "Noto Sans KR, Liberation Sans, sans-serif"
base_cfg["themeVariables"]["fontSize"] = f"{FONT_PX}px"
base_cfg["flowchart"]["wrappingWidth"] = 2000
base_cfg["themeCSS"] = (base_cfg["themeCSS"].replace("'Liberation Sans', 'Noto Sans KR'", "'Noto Sans KR', 'Liberation Sans'")
                        + " .cluster-label span { font-size: %dpx; } .label, .nodeLabel, .edgeLabel, p, span { word-break: keep-all; }" % (FONT_PX + 2))
DECK_CFG.write_text(json.dumps(base_cfg, indent=1, ensure_ascii=False))


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


def render(name, body, width=13, node=30, rank=60, curve="basis", padding=14):
    text = INIT % (FONT_PX, node, rank, curve, padding) + "\n" + wrap_labels(body, width).strip() + "\n" + CLASSES
    labels = re.findall(r'(?:\[\"|\{\"|\(\"|\|\")(.*?)"', body)
    bad = [l for l in labels if re.search(r"[-–—]", l)]
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
    info = {"w": w, "h": h, "natural_w": nat_w, "natural_h": nat_h, "font_px": FONT_PX}
    print(name, info, "label pt at 12.13 in:", round(FONT_PX * 12.13 / nat_w * 72, 1), "aspect", round(nat_w / nat_h, 2))
    return info


dims = {}

dims["01"] = render("01_launch_critical_path_deck_ko", """
flowchart LR
  subgraph P1 ["1주차에서 4주차: 준비"]
    A1["1. 브랜드명과 상표<br/>Lucy, 변호사<br/>1주차에서 2주차"]
    A2["2. Wyoming LLC와 EIN, 법인명 정확히 일치<br/>Lucy, 변호사<br/>1주차에서 2주차"]
    A5["5. 관세사, 연간 보증, 품목분류 사전심사 신청<br/>Mark, 관세사<br/>1주차"]
    A3["3. 외국환은행에 해외직접투자 신고<br/>Lucy, 세무사<br/>2주차, 자본 이체 전"]
    A4["4. USD 은행 계좌<br/>Lucy, Mark<br/>2주차에서 3주차"]
    A6["6. 아마존 계정과 Brand Registry<br/>Lucy, Mark<br/>3주차에서 4주차"]
    CAP["자본 이체<br/>Lucy<br/>3주차, 해외직접투자 신고 후"]
    A2 --> A4 --> CAP
    A3 --> CAP
    A1 --> A6
    A2 --> A6
    A5 ~~~ A4
  end
  subgraph P2 ["4주차에서 9주차:<br/>재고와 콘텐츠"]
    B7["7. 테스트 재고: 주력 5개 SKU, 각 100에서 300개, 항공 DDP, 출발 전 QC<br/>Lucy, 공장, 관세사, 3PL<br/>4주차에서 8주차"]
    B8["8. 리스팅, 이미지, A+ 콘텐츠, 영상, PPC<br/>Mark, VA<br/>6주차에서 9주차"]
    B7 --> B8
  end
  subgraph P3 ["10주차에서 22주차: 검증"]
    C9["9. 90일 검증: 전환율, TACOS, 반품률, SKU별 공헌이익<br/>Mark, Lucy<br/>10주차에서 22주차"]
    G{{"90일 게이트<br/>Lucy 결정<br/>22주차"}}
    C9 --> G
  end
  subgraph P4 ["22주차 이후"]
    S10["10. 판정된 관세율로 해상 운송, 승자 SKU 확대, 패자 SKU 삭제<br/>Lucy, Mark, 3PL"]
    PB["Plan B: B2B 설치업체와 한국 도매<br/>Lucy, Mark"]
  end
  CAP --> B7
  A6 --> B8
  B8 --> C9
  G -->|"통과"| S10
  G -->|"중단"| PB
  class A1,A2,A3,A4,A5,A6,B7,B8,C9,S10 step
  class CAP money
  class G dec
  class PB kill
""", width=13, node=20, rank=44, padding=10)

dims["06"] = render("06_day90_scale_or_kill_deck_ko", """
flowchart LR
  subgraph KT ["3개월차 말 90일 게이트<br/>Lucy 결정, Mark 수치 준비<br/>1단계: 중단 트리거 5개, 하나라도 해당되면 중단"]
    direction LR
    K1["1. 주력 SKU 전환율 5% 미만?"]
    K2["2. 3개월차 TACOS 40% 초과?"]
    K3["3. 반품률 7% 초과?"]
    K4["4. 주력 SKU 광고 전 공헌이익 개당 $8 미만?"]
    K5["5. 현금 런웨이 2개월 미만?"]
    K1 ~~~ K2 ~~~ K3
    K4 ~~~ K5
  end
  subgraph PT ["2단계: 통과 테스트 6개<br/>모두 통과해야 함"]
    direction LR
    P1["1. 출시 SKU 5개 중 3개 전환율 8% 이상?"]
    P2["2. 3개월차 TACOS 30% 이하, 매월 하락?"]
    P3["3. 반품률 4% 이하?"]
    P4["4. 5개 중 3개 SKU 광고 전 공헌이익 개당 $11 이상?"]
    P5["5. 다음 PO 자금 확보 후 현금 런웨이 4개월 이상?"]
    P6["6. 리스팅 정지 없음, 14일 초과 IP 신고 없음?"]
    P1 ~~~ P2 ~~~ P3
    P4 ~~~ P5 ~~~ P6
  end
  KT -->|"하나라도 예"| KILL["KILL (중단). 추가 자본 없음. 재고 소진, 패자 SKU 삭제, Plan B 전환: B2B 설치업체와 한국 도매"]
  KT -->|"모두 아니오"| PT
  PT -->|"모두 예"| SCALE["SCALE (확대). 검증된 SKU는 해상 운송 PO, 패자 SKU 삭제, 신규 SKU는 항공 유지. 다음 검토: 12개월차 게이트"]
  PT -->|"하나라도 아니오"| HOLD["REVIEW (재검토). 실패한 테스트 보완 후 30일 내 재검증. 확정된 PO 외 신규 자본 없음. 재검증은 1회, 이후 통과 또는 중단"]
  class K1,K2,K3,K4,K5,P1,P2,P3,P4,P5,P6 dec
  class KILL kill
  class SCALE good
  class HOLD ext
""", width=13, node=12, rank=50, curve="linear", padding=10)

dims["07"] = render("07_five_year_channel_map_deck_ko", """
flowchart LR
  subgraph Y1 ["1년차: 아마존 US만"]
    direction TB
    A1["아마존 US. 부모 ASIN 하나 아래 주력 5개 SKU. FBA. Brand Registry"]
    G1{{"90일 게이트"}}
    G2{{"12개월차 게이트"}}
    PB["아마존 실패 시 Plan B: B2B 설치업체 유통과 한국 도매"]
    A1 --> G1
    G1 -->|"통과"| G2
    G1 -->|"중단"| PB
    G2 -->|"중단"| PB
  end
  subgraph Y2 ["2년차: 오버플로와<br/>설치업체"]
    direction TB
    B1["아마존 US: 전체 길이 라인업, 검증된 SKU는 해상 운송"]
    B2["eBay와 Walmart 오버플로. 멀티채널 풀필먼트, 추가 노력 거의 없음"]
    B3["B2B 설치업체와 AV 통합업체 프로그램. 3PL 출고, 소매가의 60%"]
    B1 ~~~ B2 ~~~ B3
  end
  subgraph Y3 ["3년차: 자사 채널과<br/>신제품"]
    direction TB
    C1["아마존 외부 이메일 목록(Spec Match 파인더, 트레이드업 등록)으로 운영하는 DTC 사이트"]
    C2["로드맵 SKU: HDMI 2.2 16K, DP 2.1, USB4 AOC"]
    C1 ~~~ C2
  end
  subgraph Y4 ["4년차: 캐나다와<br/>프로 라인"]
    direction TB
    D1["아마존 캐나다, 미국 재고 이전"]
    D2["프로 라인: CL3, 플레넘, 배관 포설, 장거리 광 KVM. B2B 60%, 아마존 40%"]
    D1 ~~~ D2
  end
  subgraph Y5 ["5년차: 확대"]
    direction TB
    E1["모든 채널 확대. Medium 밴드 목표: 순이익률 10%에서 15% (모델 참조)"]
    E2["Temu 현지 셀러 프로그램: 선택 사항, 아마존 가격 하한이 유지된 뒤에만. 기준 케이스 미포함"]
    E1 ~~~ E2
  end
  Y1 --> Y2 --> Y3 --> Y4 --> Y5
  class A1,B1,B2,B3,C1,C2,D1,D2 step
  class G1,G2 dec
  class E1 good
  class E2 ext
  class PB kill
""", width=15, node=20, rank=50, padding=10)

(OUT / "dims_ko.json").write_text(json.dumps(dims, indent=1))
print("wrote", OUT / "dims_ko.json")

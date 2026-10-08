#!/usr/bin/env python3
"""Build output/10_One_Page_Summary_EN.pdf and output/10_One_Page_Summary_KO.pdf (one page each).

Reads build/figures.json (model numbers through fmt.py) and build/facts.json (static facts).
Korean text uses Noto Sans KR. The installed file is a variable font whose default instance is
Thin (wght 100), so this script instantiates static Regular (400) and Bold (700) copies with
fontTools into build/assets/fonts/ and registers those with reportlab.

Usage: source build/env.sh && python3 build/build_onepagers.py
QA renders go to build/qa/onepagers/. The script asserts page count == 1 with pypdf and shrinks the
type scale and rebuilds if a language ever spills to a second page.
"""
import json, os, sys, subprocess
from pathlib import Path

ROOT = Path(os.environ.get("ROOT", "/home/user/lakehouse-tracker/BRAND_US_Launch"))
BUILD = ROOT / "build"
OUT = ROOT / "output"
QA = BUILD / "qa" / "onepagers"
QA.mkdir(parents=True, exist_ok=True)
sys.path.insert(0, str(BUILD))
import fmt  # noqa: E402

from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_RIGHT
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Table, TableStyle, Spacer, Flowable)
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from pypdf import PdfReader

fig = json.load(open(BUILD / "figures.json"))
facts = json.load(open(BUILD / "facts.json"))

# ---------------------------------------------------------------- fonts
FONT_DIR = BUILD / "assets" / "fonts"
FONT_DIR.mkdir(parents=True, exist_ok=True)
VF = Path("/usr/share/fonts/truetype/notokr/NotoSansKR-VF.ttf")

def instantiate(weight, name):
    target = FONT_DIR / f"NotoSansKR-{name}.ttf"
    if target.exists():
        return target
    from fontTools.ttLib import TTFont as FTFont
    from fontTools.varLib import instancer
    f = FTFont(str(VF))
    inst = instancer.instantiateVariableFont(f, {"wght": weight})
    # give each static instance its own family, full and PostScript name so PDF viewers never merge them
    nt = inst["name"]
    for rec in nt.names:
        if rec.nameID in (1, 16):
            rec.string = f"Noto Sans KR {name}"
        elif rec.nameID == 2 or rec.nameID == 17:
            rec.string = name
        elif rec.nameID == 4:
            rec.string = f"Noto Sans KR {name}"
        elif rec.nameID == 6:
            rec.string = f"NotoSansKR-{name}"
    inst["OS/2"].usWeightClass = int(weight)
    inst.save(str(target))
    return target

KO_REG = instantiate(400, "Regular")
KO_BOLD = instantiate(700, "Bold")
pdfmetrics.registerFont(TTFont("NotoKR", str(KO_REG)))
pdfmetrics.registerFont(TTFont("NotoKR-Bold", str(KO_BOLD)))
pdfmetrics.registerFontFamily("NotoKR", normal="NotoKR", bold="NotoKR-Bold", italic="NotoKR", boldItalic="NotoKR-Bold")
pdfmetrics.registerFont(TTFont("Lib", "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf"))
pdfmetrics.registerFont(TTFont("Lib-Bold", "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf"))
pdfmetrics.registerFontFamily("Lib", normal="Lib", bold="Lib-Bold", italic="Lib", boldItalic="Lib-Bold")

# ---------------------------------------------------------------- palette (deck palette, print friendly)
NAVY = colors.HexColor("#0A0F1E")
CARD = colors.HexColor("#141C36")
CYAN = colors.HexColor("#1AA7C2")
VIOLET = colors.HexColor("#6D4ACB")
GREEN = colors.HexColor("#2E9E6E")
AMBER = colors.HexColor("#D98B1C")
MUTED = colors.HexColor("#5B6480")
LIGHT = colors.HexColor("#EEF2FA")
LINE = colors.HexColor("#C9D1E3")

# ---------------------------------------------------------------- numbers (figures.json through fmt)
M = fig["scenarios"]["Medium"]
L = fig["scenarios"]["Low"]
H = fig["scenarios"]["High"]
FX = fig["meta"]["fx"]
opening = M["cash"]["opening_capital"]
N = {
    "opening": fmt.money(opening),
    "opening_k": fmt.money_k(opening),
    "opening_krw": fmt.krw_m(opening * FX),
    "m_peak": fmt.money_k(M["cash"]["peak_need"]),
    "m_peak_full": fmt.money(M["cash"]["peak_need"]),
    "m_peak_month": fmt.month_label(M["cash"]["peak_need_month"]),
    "m_gap": fmt.money_k(M["cash"]["funding_gap"]),
    "m_gap_full": fmt.money(M["cash"]["funding_gap"]),
    "m_gap_krw": fmt.krw_m(M["cash"]["funding_gap"] * FX),
    "m_gap_month": fmt.month_label(M["cash"]["funding_needed_by_month"]),
    "m_runway": fmt.num(M["cash"]["runway_months_day90"], 1),
    "l_peak": fmt.money_k(L["cash"]["peak_need"]),
    "l_peak_month": fmt.month_label(L["cash"]["peak_need_month"]),
    "l_gap": fmt.money_k(L["cash"]["funding_gap"]),
    "l_gap_full": fmt.money(L["cash"]["funding_gap"]),
    "l_gap_krw": fmt.krw_m(L["cash"]["funding_gap"] * FX),
    "l_gap_month": fmt.month_label(L["cash"]["funding_needed_by_month"]),
    "h_gap": fmt.money_k(H["cash"]["funding_gap"]),
    "h_gap_month": fmt.month_label(H["cash"]["funding_needed_by_month"]),
    "m18_rev_m": fmt.money_m(M["m18"]["revenue"]),
    "m18_net_m": fmt.money_k(M["m18"]["net"]),
    "m18_margin_m": fmt.pct(M["m18"]["net_margin"]),
    "m18_be_m": fmt.month_label(M["m18"]["breakeven_month"]),
    "m18_rev_l": fmt.money_k(L["m18"]["revenue"]),
    "m18_net_l": fmt.money_k(L["m18"]["net"]),
    "m18_rev_h": fmt.money_m(H["m18"]["revenue"]),
    "m18_net_h": fmt.money_k(H["m18"]["net"]),
    "y5_rev_m": fmt.money_m(M["annual"]["Y5"]["revenue"]),
    "y5_net_m": fmt.money_k(M["annual"]["Y5"]["net"]),
    "y5_rev_l": fmt.money_m(L["annual"]["Y5"]["revenue"]),
    "y5_net_l": fmt.money_k(L["annual"]["Y5"]["net"]),
    "fx": fmt.num(FX),
    "wy_form": fmt.money(facts["entity"]["formation_cost_usd"]),
    "wy_annual": fmt.money(facts["entity"]["annual_cost_usd"]),
    "penalty": fmt.money(facts["entity"]["form_5472_penalty_usd"]),
    "bond": fmt.money(facts["logistics"]["broker_bond_annual_usd"]),
    "tm_fee": fmt.money(facts["trademark"]["uspto_fee_per_class_usd"]),
    "ofa_threshold": fmt.krw_m_short(facts["korea"]["overseas_financial_account_report_threshold_krw"]),
    "insurance": fmt.money_m_short(facts["amazon"]["product_liability_insurance_usd"]),
    "insurance_trigger": fmt.money_k(facts["amazon"]["insurance_trigger_monthly_sales_usd"]),
    "reg_no": facts["korea"]["company_registration_number"],
    "built": fig["meta"]["built"],
    "gate_l_m12": L["scenario_summary_gate_month12"] if "scenario_summary_gate_month12" in L else fig["scenario_summary"]["Low"]["gate_month12"],
    "gate_m_day90": fig["scenario_summary"]["Medium"]["gate_day90"],
}

# ---------------------------------------------------------------- text, both languages
# Form names, the LLC name, the brand placeholder and official numbers stay in English in Korean text.
T = {}
T["en"] = {
    "font": "Lib",
    "title": "BRAND LLC: US launch summary for the bank and the tax accountant",
    "subtitle": f"One page. Prepared {N['built']}. Every number comes from the financial model (output/03_Financial_Model_5Y.xlsx) in the Medium case unless marked Low. Basis labels: Evidence based or Management estimate.",
    "h_who": "1. Who and what",
    "who": [
        f"<b>Lucy (Hyejin Yoo)</b>, Seoul. Owner of <b>H2 Company</b>, a Korean sole proprietorship (registration number {N['reg_no']}, middle digits 64). H2 Company buys COMQ brand hybrid active optical HDMI cables from a factory in China and sells them wholesale in Korea.",
        f"<b>BRAND LLC</b>, a Wyoming single member LLC (limited liability company), owned 100% by Lucy personally. Formation {N['wy_form']}, about {N['wy_annual']} per year, no state income tax. BRAND is a placeholder until the US brand name is chosen. The LLC sells on Amazon US and is the importer of record through a customs broker.",
        "<b>Mark Taylor</b>, Memphis. Advisor and US operations support under a written consulting or commission agreement. No equity, no office, does not negotiate or sign contracts. If he became an owner, the LLC would become a partnership (Form 1065, 37% withholding), so the plan keeps him as a consultant.",
        "<b>Tax position (Evidence based):</b> under the 1976 US Korea tax treaty, business profits are taxed only where the permanent establishment (PE) sits. Goods stored and shipped from Amazon FBA warehouses do not create a PE (Article 9(3)(b)). No US office, no US employees, no US agent who signs contracts. Amazon collects marketplace sales tax. California income tax nexus from FBA inventory: on the CPA list.",
        "<b>Question for the tax accountant:</b> should H2 Company incorporate first and own the LLC, or should Lucy own it personally (the default used here)? Decide before the capital transfer.",
    ],
    "h_diag": "2. Structure, money and goods",
    "diag": {
        "h2": "H2 Company\nSeoul, Lucy 100%",
        "llc": "BRAND LLC\nWyoming, Lucy 100%",
        "factory": "Factory\nChina",
        "us": "US customs, 3PL, FBA\nBRAND LLC importer",
        "cust": "Amazon US\ncustomer",
        "mark": "Mark Taylor\nconsultant, no equity",
        "cap": "capital after ODI report",
        "po": "PO and payment",
        "goods": "goods",
        "sales": "sales",
        "payout": "Amazon payout",
        "fee": "consulting fee",
        "legend": "Dashed arrows are money, solid arrows are goods. 1 capital from Lucy after the ODI report at the foreign exchange bank. 2 consulting fee to Mark. 3 purchase orders and factory payments. 4 goods: formal customs entry, BRAND LLC as importer of record through a broker with a continuous bond. 5 sales on Amazon US. 6 Amazon payouts to the LLC account, then owner draw to Lucy.",
    },
    "h_filings": "3. Filings in order, relative to the capital transfer",
    "filings_head": ["When", "Filing", "Who", "Cost and note"],
    "filings": [
        ["Before any capital moves", "Overseas direct investment (ODI) report to Lucy's foreign exchange bank under the Foreign Exchange Transactions Act (FTA). Required for any stake of 10% or more, any amount.", "Lucy, bank, tax accountant", "Bank form. Enforcement against individual online sellers is rising. Evidence based."],
        ["Weeks 1 to 2", "Wyoming LLC formation with a registered agent. Exact legal name BRAND LLC used on every later form.", "Lucy, attorney", f"{N['wy_form']} state fee. Evidence based."],
        ["Weeks 1 to 2, after formation", "EIN (Employer Identification Number) from the IRS by phone or fax with Form SS-4. No SSN needed. Format 12-3456789.", "Lucy, Mark", "No fee. Evidence based."],
        ["Weeks 2 to 3", "USD bank account for BRAND LLC: Mercury or Relay first, Payoneer or WorldFirst as fallback. Then the capital transfer from Lucy to the LLC account.", "Lucy", f"Opening capital {N['opening']} (about {N['opening_krw']} at {N['fx']} KRW per USD). Evidence based (August ask)."],
        ["After the transfer", "Korean follow up reports to the same bank: remittance report, securities acquisition report, then the annual business performance report.", "Lucy, tax accountant", "Bank forms. Evidence based."],
        ["Weeks 3 to 4", "Amazon Professional seller account: tax interview as a foreign owned entity (Form W-8BEN-E), identity verification, Brand Registry with the pending trademark.", "Lucy, Mark", "Trademark through a US licensed attorney, about " + N["tm_fee"] + " per class plus attorney fees [PLACEHOLDER: attorney quote]."],
        ["Every year, US", f"Form 5472 with a pro forma Form 1120 to the IRS (foreign owned single member LLC). Penalty for missing it {N['penalty']}. Wyoming annual report, about {N['wy_annual']}. BOI report not required (FinCEN exemption, August 11, 2026).", "Lucy, US CPA [PLACEHOLDER]", "Form 5472 preparation [PLACEHOLDER: CPA quote]. Evidence based."],
        ["Every year, Korea", f"Overseas financial account report when the month end aggregate of overseas accounts reaches {N['ofa_threshold']}. Annual business performance report (ODI follow up). Korean income tax on the LLC profit through the tax accountant.", "Lucy, tax accountant", "Threshold and timing to be confirmed by the tax accountant. Evidence based threshold."],
    ],
    "h_capital": "4. Capital need and funding gap",
    "cap_head": ["Item", "Medium case", "Low case (downside)", "Basis"],
    "capital": [
        ["Opening capital, self funded by Lucy", N["opening"], N["opening"], "Evidence based, August ask"],
        ["Peak cash need (lowest cash balance before new funding)", f"{N['m_peak']} in {N['m_peak_month']}", f"{N['l_peak']} in {N['l_peak_month']}", "Model"],
        ["Funding gap (1.2 times peak need), line of credit or investor", f"{N['m_gap']} ({N['m_gap_full']}, about {N['m_gap_krw']})", f"{N['l_gap']} ({N['l_gap_full']}, about {N['l_gap_krw']})", "Model"],
        ["Date the funding must be in place", f"{N['m_gap_month']} (Month 1 = first month with inventory live on Amazon, assumed early 2027)", N["l_gap_month"], "Model"],
        ["18 month Amazon US revenue and net", f"{N['m18_rev_m']} revenue, {N['m18_net_m']} net ({N['m18_margin_m']}), cumulative breakeven {N['m18_be_m']}", f"{N['m18_rev_l']} revenue, {N['m18_net_l']} net, no breakeven in 18 months", "Model, August basis"],
        ["Year 5 revenue and net", f"{N['y5_rev_m']} revenue, {N['y5_net_m']} net", f"{N['y5_rev_l']} revenue, {N['y5_net_l']} net; Month 12 gate result {N['gate_l_m12']}", "Model, Management estimate"],
    ],
    "cap_note": f"Working capital is the risk, not margin: inventory is paid 30% at order and 70% at shipment, ships before it sells, and Amazon pays 14 days after the sale. High case gap for reference: {N['h_gap']} by {N['h_gap_month']}. The duty rate (copper 40.1% versus fiber 12.5%, pending a CBP binding ruling) is the single variable that separates Low from Medium.",
    "h_timeline": "5. Timeline, weeks 1 to 22, then the 90 day gate",
    "tl_head": ["Weeks", "Step", "Owner"],
    "timeline": [
        ["1", "Customs broker engaged, continuous bond (about " + N["bond"] + " per year), CBP binding ruling request filed (30 to 90 days)", "Broker, Mark"],
        ["1 to 2", "Brand name chosen, trademark filed; Wyoming LLC formed; EIN by Form SS-4", "Lucy, attorney"],
        ["Before transfer", "ODI report accepted by the foreign exchange bank", "Lucy, bank"],
        ["2 to 3", "USD bank account opened, capital transferred", "Lucy"],
        ["3 to 4", "Amazon account, tax interview, Brand Registry", "Lucy, Mark"],
        ["4 to 8", "Test inventory: 5 launch SKUs, 100 to 300 units each, air freight, factory QC and FNSKU labels", "Lucy, factory, 3PL"],
        ["6 to 9", "Listings and content live; product liability insurance " + N["insurance"] + " before sales pass " + N["insurance_trigger"] + " in a month", "Mark"],
        ["10 to 22", "90 day validation: conversion, TACOS, return rate, contribution per SKU, cash runway", "Lucy, Mark"],
        ["22 onward", f"Day 90 gate: PASS, REVIEW or KILL. Medium case result {N['gate_m_day90']}. Scale winners on ocean freight, delist losers. KILL means no further capital.", "Lucy"],
    ],
    "h_contacts": "6. Contacts (placeholders until confirmed)",
    "contacts": [
        ["Lucy (Hyejin Yoo), H2 Company", "[PLACEHOLDER: phone and email]"],
        ["Mark Taylor, consultant, Memphis", "[PLACEHOLDER: phone and email]"],
        ["Korean tax accountant", "[PLACEHOLDER: name, firm, phone]"],
        ["Foreign exchange bank, ODI desk", "[PLACEHOLDER: bank, branch, officer]"],
        ["US attorney (LLC, trademark)", "[PLACEHOLDER: name, firm]"],
        ["US customs broker and US CPA", "[PLACEHOLDER: names]"],
    ],
    "footer": "Glossary: ODI overseas direct investment report; FTA Foreign Exchange Transactions Act; EIN Employer Identification Number; LLC limited liability company; IRS Internal Revenue Service; PE permanent establishment; FBA Fulfillment by Amazon; 3PL third party logistics warehouse; PO purchase order; SKU stock keeping unit; CBP US Customs and Border Protection; TACOS total advertising cost of sales; FNSKU Fulfillment Network Stock Keeping Unit; QC quality control. Nothing here is legal or tax advice; the tax accountant, bank and attorney confirm each filing.",
}

T["ko"] = {
    "font": "NotoKR",
    "title": "BRAND LLC 미국 진출 요약: 은행 및 세무사 제출용",
    "subtitle": f"한 장 요약. 작성일 {N['built']}. 모든 숫자는 재무모델(output/03_Financial_Model_5Y.xlsx)의 Medium 시나리오 기준이며, Low로 표시한 항목만 Low 시나리오입니다. 근거 표시: Evidence based(검증된 근거) 또는 Management estimate(경영진 추정).",
    "h_who": "1. 당사자와 사업 내용",
    "who": [
        f"<b>유혜진(Lucy)</b>, 서울. <b>H2 Company</b>(사업자등록번호 {N['reg_no']}, 가운데 자리 64는 개인사업자)의 대표입니다. H2 Company는 중국 공장에서 COMQ 브랜드 hybrid AOC(active optical cable, 광섬유 HDMI 케이블)를 구매하여 한국에서 도매로 판매합니다.",
        f"<b>BRAND LLC</b>: 미국 Wyoming 주의 단일 구성원 LLC(limited liability company)이며 유혜진 대표가 개인 자격으로 지분 100%를 보유합니다. 설립비 {N['wy_form']}, 연간 약 {N['wy_annual']}, 주 소득세 없음. BRAND는 미국 브랜드명 확정 전까지 사용하는 임시 명칭입니다. LLC가 Amazon US에서 판매하며 관세사(customs broker)를 통한 수입자(importer of record)가 됩니다.",
        "<b>Mark Taylor</b>, 미국 Memphis. 서면 자문 계약 또는 수수료 계약에 따른 컨설턴트입니다. 지분 없음, 사무실 없음, 계약 협상이나 서명 권한 없음. Mark가 지분을 가지면 LLC는 파트너십이 되어 Form 1065 제출과 37% 원천징수가 발생하므로, 본 계획은 컨설턴트 관계를 유지합니다.",
        "<b>세무 포지션(Evidence based):</b> 1976년 한미 조세조약에 따라 사업소득은 고정사업장(PE, permanent establishment)이 있는 국가에서만 과세됩니다. Amazon FBA 창고에 보관 및 출고되는 재고는 고정사업장을 구성하지 않습니다(Article 9(3)(b)). 미국 내 사무실, 직원, 계약 체결 대리인을 두지 않습니다. 판매세(sales tax)는 Amazon이 징수합니다. California 주의 FBA 재고 기준 소득세 과세 가능성은 미국 CPA 확인 항목입니다.",
        "<b>세무사 확인 질문:</b> H2 Company를 먼저 법인으로 전환하여 LLC를 소유할지, 아니면 유혜진 대표가 개인으로 소유할지(본 문서의 기본 가정)를 자본 송금 전에 결정해야 합니다.",
    ],
    "h_diag": "2. 구조, 자금 및 물품 흐름",
    "diag": {
        "h2": "H2 Company\n서울, 유혜진 100%",
        "llc": "BRAND LLC\nWyoming, 유혜진 100%",
        "factory": "공장\n중국",
        "us": "미국 세관, 3PL, FBA\nBRAND LLC 수입자",
        "cust": "Amazon US\n소비자",
        "mark": "Mark Taylor\n컨설턴트, 지분 없음",
        "cap": "ODI 신고 후 자본금 송금",
        "po": "발주 및 대금",
        "goods": "물품",
        "sales": "판매",
        "payout": "Amazon 정산금",
        "fee": "자문료",
        "legend": "점선 화살표는 자금, 실선 화살표는 물품입니다. 1 외국환은행에 ODI 신고 후 유혜진 대표가 송금하는 자본금. 2 Mark에게 지급하는 자문료. 3 공장 발주 및 대금 지급. 4 물품: 정식 통관, 관세사와 연간 보증보험을 통해 BRAND LLC가 수입자. 5 Amazon US 판매. 6 Amazon 정산금은 LLC 계좌로 입금되고, 이후 유혜진 대표에게 배당(owner draw).",
    },
    "h_filings": "3. 신고 및 제출 순서 (자본 송금 시점 기준)",
    "filings_head": ["시점", "신고 및 제출 사항", "담당", "비용 및 비고"],
    "filings": [
        ["자본 송금 전", "외국환거래법(FTA, Foreign Exchange Transactions Act)에 따른 해외직접투자 신고(ODI report)를 유혜진 대표의 외국환은행에 제출합니다. 지분 10% 이상이면 금액과 무관하게 신고 대상입니다.", "유혜진, 은행, 세무사", "은행 양식. 개인 온라인 판매자에 대한 단속이 강화되는 추세입니다. Evidence based."],
        ["1~2주차", "Wyoming 주 LLC 설립(등록대리인 지정). 이후 모든 서류에 법인명 BRAND LLC를 동일하게 사용합니다.", "유혜진, 미국 변호사", f"주 수수료 {N['wy_form']}. Evidence based."],
        ["1~2주차, 설립 후", "IRS에 Form SS-4로 EIN(Employer Identification Number, 고용주 식별번호) 신청. 전화 또는 팩스, SSN 불필요. 형식 12-3456789.", "유혜진, Mark", "수수료 없음. Evidence based."],
        ["2~3주차", "BRAND LLC 명의 미국 달러 은행계좌 개설: Mercury 또는 Relay 우선, Payoneer 또는 WorldFirst 대안. 이후 유혜진 대표가 LLC 계좌로 자본금을 송금합니다.", "유혜진", f"초기 자본금 {N['opening']} (환율 {N['fx']}원 기준 약 {N['opening_krw']}). Evidence based (8월 제안서)."],
        ["송금 후", "같은 외국환은행에 사후보고: 송금보고서, 증권취득보고서, 이후 매년 연간 사업실적보고서.", "유혜진, 세무사", "은행 양식. Evidence based."],
        ["3~4주차", "Amazon Professional 판매자 계정: 외국 법인으로서 세금 인터뷰(Form W-8BEN-E), 신원 확인, 상표 출원 중 상태로 Brand Registry 등록.", "유혜진, Mark", "상표는 미국 변호사를 통해 출원, 분류당 약 " + N["tm_fee"] + " 및 변호사 비용 [PLACEHOLDER: 변호사 견적]."],
        ["매년, 미국", f"외국인 소유 단일 구성원 LLC로서 IRS에 Form 5472와 pro forma Form 1120 제출. 미제출 벌금 {N['penalty']}. Wyoming annual report 약 {N['wy_annual']}. BOI 보고는 불필요(FinCEN 면제 확정, 2026년 8월 11일).", "유혜진, 미국 CPA [PLACEHOLDER]", "Form 5472 작성 비용 [PLACEHOLDER: CPA 견적]. Evidence based."],
        ["매년, 한국", f"해외금융계좌 신고: 해외 계좌 잔액의 월말 합계가 {N['ofa_threshold']}에 이르면 신고 대상. 연간 사업실적보고서(ODI 사후보고). LLC 이익에 대한 한국 소득세 신고는 세무사를 통해 진행합니다.", "유혜진, 세무사", "기준 금액과 신고 시기는 세무사가 확인합니다. 기준 금액은 Evidence based."],
    ],
    "h_capital": "4. 필요 자본과 자금 부족분",
    "cap_head": ["항목", "Medium 시나리오", "Low 시나리오 (하방 위험)", "근거"],
    "capital": [
        ["초기 자본금 (유혜진 대표 자기자금)", N["opening"], N["opening"], "Evidence based, 8월 제안서"],
        ["최대 현금 소요 (추가 자금 조달 전 최저 현금 잔액)", f"{N['m_peak']} ({N['m_peak_month']})", f"{N['l_peak']} ({N['l_peak_month']})", "재무모델"],
        ["자금 부족분 (최대 소요의 1.2배), 신용한도 또는 투자금", f"{N['m_gap']} ({N['m_gap_full']}, 약 {N['m_gap_krw']})", f"{N['l_gap']} ({N['l_gap_full']}, 약 {N['l_gap_krw']})", "재무모델"],
        ["자금이 준비되어야 하는 시점", f"{N['m_gap_month']} (Month 1 = Amazon에 재고가 입고되는 첫 달, 2027년 초 예상)", N["l_gap_month"], "재무모델"],
        ["18개월 Amazon US 매출과 순이익", f"매출 {N['m18_rev_m']}, 순이익 {N['m18_net_m']} ({N['m18_margin_m']}), 누적 손익분기 {N['m18_be_m']}", f"매출 {N['m18_rev_l']}, 순이익 {N['m18_net_l']}, 18개월 내 손익분기 없음", "재무모델, 8월 기준"],
        ["5년차 매출과 순이익", f"매출 {N['y5_rev_m']}, 순이익 {N['y5_net_m']}", f"매출 {N['y5_rev_l']}, 순이익 {N['y5_net_l']}; Month 12 게이트 결과 {N['gate_l_m12']}", "재무모델, Management estimate"],
    ],
    "cap_note": f"위험은 마진이 아니라 운전자본입니다. 재고 대금은 발주 시 30%, 선적 시 70%를 지급하고, 재고는 판매 전에 먼저 출고되며, Amazon 정산은 판매 후 14일입니다. 참고로 High 시나리오의 부족분은 {N['h_gap']} ({N['h_gap_month']})입니다. 관세율(구리 케이블 분류 40.1% 대 광섬유 분류 12.5%, CBP 사전심사 결정 대기 중)이 Low와 Medium을 가르는 단일 변수입니다.",
    "h_timeline": "5. 일정: 1~22주차, 이후 90일 게이트",
    "tl_head": ["주차", "단계", "담당"],
    "timeline": [
        ["1", "관세사 선임, 연간 보증보험(continuous bond, 연 약 " + N["bond"] + "), CBP 사전심사(binding ruling) 신청 (30~90일 소요)", "관세사, Mark"],
        ["1~2", "브랜드명 확정 및 상표 출원; Wyoming LLC 설립; Form SS-4로 EIN 취득", "유혜진, 변호사"],
        ["송금 전", "외국환은행의 ODI 신고 수리", "유혜진, 은행"],
        ["2~3", "미국 달러 계좌 개설, 자본금 송금", "유혜진"],
        ["3~4", "Amazon 계정, 세금 인터뷰, Brand Registry", "유혜진, Mark"],
        ["4~8", "시험 재고: 출시 SKU 5종, 종당 100~300개, 항공 운송, 공장 품질검사 및 FNSKU 라벨", "유혜진, 공장, 3PL"],
        ["6~9", "상품 페이지 및 콘텐츠 게시; 월 매출 " + N["insurance_trigger"] + " 초과 전 제조물책임보험 " + N["insurance"] + " 가입", "Mark"],
        ["10~22", "90일 검증: 전환율, TACOS, 반품률, SKU별 공헌이익, 현금 여력", "유혜진, Mark"],
        ["22주차 이후", f"Day 90 게이트: PASS, REVIEW 또는 KILL. Medium 시나리오 결과 {N['gate_m_day90']}. 검증된 SKU는 해상 운송으로 확대, 부진 SKU는 판매 중단. KILL이면 추가 자본 투입 중단.", "유혜진"],
    ],
    "h_contacts": "6. 연락처 (확정 전까지 자리표시)",
    "contacts": [
        ["유혜진(Lucy), H2 Company", "[PLACEHOLDER: 전화 및 이메일]"],
        ["Mark Taylor, 컨설턴트, Memphis", "[PLACEHOLDER: 전화 및 이메일]"],
        ["세무사", "[PLACEHOLDER: 성명, 사무소, 전화]"],
        ["외국환은행 해외직접투자 담당", "[PLACEHOLDER: 은행, 지점, 담당자]"],
        ["미국 변호사 (LLC, 상표)", "[PLACEHOLDER: 성명, 사무소]"],
        ["미국 관세사 및 미국 CPA", "[PLACEHOLDER: 성명]"],
    ],
    "footer": "용어: ODI 해외직접투자 신고(overseas direct investment report); FTA 외국환거래법; EIN 고용주 식별번호; LLC 유한책임회사; IRS 미국 국세청; PE 고정사업장; FBA Fulfillment by Amazon(Amazon 물류 대행); 3PL 제3자 물류창고; PO 발주서; SKU 재고 관리 단위; CBP 미국 관세국경보호청; TACOS 총광고비 비율(광고비 나누기 총매출); FNSKU Amazon 물류 바코드; QC 품질검사. 본 문서는 법률 또는 세무 자문이 아니며, 각 신고 사항은 세무사, 은행, 변호사가 확인합니다.",
}


# ---------------------------------------------------------------- structure diagram flowable
class StructureDiagram(Flowable):
    def __init__(self, d, font, width, height, scale):
        super().__init__()
        self.d, self.font, self.w, self.h, self.s = d, font, width, height, scale

    def wrap(self, aw, ah):
        return self.w, self.h

    def _box(self, c, x, y, w, h, text, fill, tcolor=colors.white):
        c.setFillColor(fill); c.setStrokeColor(fill)
        c.roundRect(x, y, w, h, 3, fill=1, stroke=0)
        c.setFillColor(tcolor)
        lines = text.split("\n")
        fs = 6.4 * self.s
        # shrink until both lines fit inside the box with 3 pt side padding
        while fs > 4.0 and max(pdfmetrics.stringWidth(lines[0], self.font + "-Bold", fs),
                               pdfmetrics.stringWidth(lines[-1], self.font, fs - 0.6)) > w - 6:
            fs -= 0.2
        c.setFont(self.font + "-Bold", fs)
        c.drawCentredString(x + w / 2, y + h / 2 + fs * 0.35, lines[0])
        c.setFont(self.font, fs - 0.6)
        if len(lines) > 1:
            c.drawCentredString(x + w / 2, y + h / 2 - fs * 0.85, lines[1])

    def _arrow(self, c, x1, y1, x2, y2, num, dashed=False, color=NAVY, dx=0, dy=4):
        import math
        c.setStrokeColor(color); c.setFillColor(color); c.setLineWidth(0.8)
        c.setDash(2, 2) if dashed else c.setDash()
        c.line(x1, y1, x2, y2)
        c.setDash()
        ang = math.atan2(y2 - y1, x2 - x1)
        L = 3.5
        p = c.beginPath()
        p.moveTo(x2, y2)
        p.lineTo(x2 - L * math.cos(ang - 0.45), y2 - L * math.sin(ang - 0.45))
        p.lineTo(x2 - L * math.cos(ang + 0.45), y2 - L * math.sin(ang + 0.45))
        p.close()
        c.drawPath(p, fill=1, stroke=0)
        # numbered badge at the midpoint, offset off the line
        mx, my = (x1 + x2) / 2 + dx, (y1 + y2) / 2 + dy
        r = 3.6 * self.s
        c.setFillColor(colors.white); c.setStrokeColor(color); c.setLineWidth(0.6)
        c.circle(mx, my, r, fill=1, stroke=1)
        c.setFillColor(color); c.setFont(self.font + "-Bold", 5.2 * self.s)
        c.drawCentredString(mx, my - 1.8 * self.s, str(num))

    def draw(self):
        c = self.canv
        d = self.d
        W, Hh = self.w, self.h
        bw, bh = W * 0.27, Hh * 0.26
        gap = (W - 3 * bw) / 2
        ytop = Hh - bh - 1
        ymid = 1
        xs = [0, bw + gap, 2 * (bw + gap)]
        self._box(c, xs[0], ytop, bw, bh, d["h2"], CARD)
        self._box(c, xs[1], ytop, bw, bh, d["llc"], CYAN)
        self._box(c, xs[2], ytop, bw, bh, d["mark"], VIOLET)
        self._box(c, xs[0], ymid, bw, bh, d["factory"], MUTED)
        self._box(c, xs[1], ymid, bw, bh, d["us"], MUTED)
        self._box(c, xs[2], ymid, bw, bh, d["cust"], GREEN)
        ym_top = ytop + bh * 0.5
        ym_mid = ymid + bh * 0.5
        # 1 capital H2 -> LLC (money)
        self._arrow(c, xs[0] + bw, ym_top, xs[1], ym_top, 1, dashed=True, color=CYAN, dy=5.5)
        # 2 consulting fee LLC -> Mark (money)
        self._arrow(c, xs[1] + bw, ym_top, xs[2], ym_top, 2, dashed=True, color=VIOLET, dy=5.5)
        # 3 PO and payment LLC -> factory (money)
        self._arrow(c, xs[1] + bw * 0.2, ytop, xs[0] + bw * 0.7, ymid + bh, 3, dashed=True, color=CYAN, dx=-6, dy=0)
        # 4 goods factory -> US (goods)
        self._arrow(c, xs[0] + bw, ym_mid, xs[1], ym_mid, 4, color=NAVY, dy=5.5)
        # 5 sales US -> customer (goods)
        self._arrow(c, xs[1] + bw, ym_mid, xs[2], ym_mid, 5, color=NAVY, dy=5.5)
        # 6 Amazon payout customer -> LLC (money)
        self._arrow(c, xs[2] + bw * 0.5, ymid + bh, xs[1] + bw * 0.8, ytop, 6, dashed=True, color=GREEN, dx=6, dy=0)


# ---------------------------------------------------------------- page build
def build(lang, scale=1.0):
    t = T[lang]
    F = t["font"]
    FB = F + "-Bold"
    out = OUT / f"10_One_Page_Summary_{lang.upper()}.pdf"
    lead = 1.22 if lang == "en" else 1.32
    base = 7.0 * scale
    st_title = ParagraphStyle("t", fontName=FB, fontSize=12.5 * scale, leading=15 * scale, textColor=NAVY)
    st_sub = ParagraphStyle("s", fontName=F, fontSize=base - 0.6, leading=(base - 0.6) * lead, textColor=MUTED)
    st_h = ParagraphStyle("h", fontName=FB, fontSize=base + 1.4, leading=(base + 1.4) * 1.2, textColor=CYAN, spaceBefore=3 * scale, spaceAfter=1.5 * scale)
    st_b = ParagraphStyle("b", fontName=F, fontSize=base, leading=base * lead, textColor=NAVY)
    st_bul = ParagraphStyle("bul", parent=st_b, leftIndent=6, bulletIndent=0, spaceAfter=1.2 * scale)
    st_cell = ParagraphStyle("c", fontName=F, fontSize=base - 0.4, leading=(base - 0.4) * lead, textColor=NAVY)
    st_cellb = ParagraphStyle("cb", parent=st_cell, fontName=FB)
    st_head = ParagraphStyle("hd", fontName=FB, fontSize=base - 0.4, leading=(base - 0.4) * 1.2, textColor=colors.white)
    st_foot = ParagraphStyle("f", fontName=F, fontSize=base - 1.2, leading=(base - 1.2) * lead, textColor=MUTED)

    margin = 0.42 * inch
    page_w = letter[0] - 2 * margin
    doc = SimpleDocTemplate(str(out), pagesize=letter, leftMargin=margin, rightMargin=margin, topMargin=0.36 * inch, bottomMargin=0.32 * inch,
                            title=t["title"], author="BRAND LLC launch package", subject="One page summary for the bank and the tax accountant")

    def P(s, style=st_cell):
        return Paragraph(s, style)

    def grid(data, widths, head_style=True, zebra=True):
        tbl = Table(data, colWidths=widths, repeatRows=1 if head_style else 0)
        cmds = [
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("LEFTPADDING", (0, 0), (-1, -1), 3), ("RIGHTPADDING", (0, 0), (-1, -1), 3),
            ("TOPPADDING", (0, 0), (-1, -1), 1.6 * scale), ("BOTTOMPADDING", (0, 0), (-1, -1), 1.6 * scale),
            ("LINEBELOW", (0, 0), (-1, -1), 0.3, LINE),
        ]
        if head_style:
            cmds += [("BACKGROUND", (0, 0), (-1, 0), NAVY)]
        if zebra:
            for i in range(1 if head_style else 0, len(data)):
                if i % 2 == 0:
                    cmds.append(("BACKGROUND", (0, i), (-1, i), LIGHT))
        tbl.setStyle(TableStyle(cmds))
        return tbl

    story = []
    # title block
    story.append(Table([[P(t["title"], st_title), P("BRAND LLC<br/>Wyoming", ParagraphStyle("r", parent=st_sub, alignment=TA_RIGHT, fontName=FB, textColor=CYAN))]],
                       colWidths=[page_w * 0.82, page_w * 0.18], style=TableStyle([("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 0)])))
    story.append(P(t["subtitle"], st_sub))
    story.append(Spacer(1, 2))

    # row 1: who and what | diagram
    who = [P(t["h_who"], st_h)] + [Paragraph(s, st_bul, bulletText="•") for s in t["who"]]
    left_w = page_w * 0.60
    right_w = page_w - left_w
    diag_h = 1.25 * inch
    diag = [P(t["h_diag"], st_h), StructureDiagram(t["diag"], F, right_w - 6, diag_h, scale), Spacer(1, 3), P(t["diag"]["legend"], st_sub)]
    row1 = Table([[who, diag]], colWidths=[left_w, right_w], style=TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"), ("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (0, 0), 6), ("RIGHTPADDING", (1, 0), (1, 0), 0),
        ("TOPPADDING", (0, 0), (-1, -1), 0), ("BOTTOMPADDING", (0, 0), (-1, -1), 0)]))
    story.append(row1)

    # row 2: filings
    story.append(P(t["h_filings"], st_h))
    fw = [page_w * 0.13, page_w * 0.50, page_w * 0.13, page_w * 0.24]
    data = [[P(h, st_head) for h in t["filings_head"]]]
    for r in t["filings"]:
        data.append([P(r[0], st_cellb), P(r[1]), P(r[2]), P(r[3])])
    story.append(grid(data, fw))

    # row 3: capital | timeline
    cw = [left_w * 0.27, left_w * 0.30, left_w * 0.26, left_w * 0.17]
    cdata = [[P(h, st_head) for h in t["cap_head"]]]
    for r in t["capital"]:
        cdata.append([P(r[0], st_cellb), P(r[1]), P(r[2]), P(r[3])])
    cap_block = [P(t["h_capital"], st_h), grid(cdata, cw), Spacer(1, 1.5), P(t["cap_note"], st_sub)]
    tw = [right_w * 0.14, right_w * 0.64, right_w * 0.22]
    tdata = [[P(h, st_head) for h in t["tl_head"]]]
    for r in t["timeline"]:
        tdata.append([P(r[0], st_cellb), P(r[1]), P(r[2])])
    tl_block = [P(t["h_timeline"], st_h), grid(tdata, tw)]
    row3 = Table([[cap_block, tl_block]], colWidths=[left_w, right_w], style=TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"), ("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (0, 0), 6), ("RIGHTPADDING", (1, 0), (1, 0), 0),
        ("TOPPADDING", (0, 0), (-1, -1), 0), ("BOTTOMPADDING", (0, 0), (-1, -1), 0)]))
    story.append(row3)

    # row 4: contacts, three pairs per row
    story.append(P(t["h_contacts"], st_h))
    cts = t["contacts"]
    cdata2 = []
    for i in range(0, len(cts), 3):
        row = []
        for j in range(3):
            if i + j < len(cts):
                row += [P(cts[i + j][0], st_cellb), P(cts[i + j][1])]
            else:
                row += ["", ""]
        cdata2.append(row)
    cwid = [page_w / 3 * 0.46, page_w / 3 * 0.54] * 3
    story.append(grid(cdata2, cwid, head_style=False, zebra=False))
    story.append(Spacer(1, 2))
    story.append(P(t["footer"], st_foot))

    doc.build(story)
    pages = len(PdfReader(str(out)).pages)
    return out, pages


def main():
    results = {}
    for lang in ("en", "ko"):
        scale = 1.0
        for attempt in range(8):
            out, pages = build(lang, scale)
            print(f"{lang}: scale {scale:.2f} pages {pages}")
            if pages == 1:
                break
            scale -= 0.04
        assert pages == 1, f"{lang} still {pages} pages"
        results[lang] = {"file": str(out), "pages": pages, "scale": round(scale, 2)}
    json.dump({"results": results, "numbers": N}, open(QA / "onepagers_build.json", "w"), indent=1, ensure_ascii=False)
    print(json.dumps(results, indent=1))


if __name__ == "__main__":
    main()

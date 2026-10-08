#!/usr/bin/env python3
"""Build output/06_Compliance_Checklist.xlsx.
Every static fact comes from build/facts.json, every model number from build/figures.json (formatted through build/fmt.py),
working capital inputs from build/sops_data.json. Estimated cost = USD for the first 12 months after formation
(facts where known, otherwise a labeled management estimate). Basis is exactly "Evidence based" or "Management estimate".
No hyphens, en dashes or em dashes in any text (official form and registration numbers keep their format).
"""
import json, math, os, sys, datetime
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import fmt as F
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from openpyxl.utils import get_column_letter

def patch_fonts(path):
    """After $RECALC (LibreOffice store) the cells with Hangul carry the fallback font name. Rewrite xl/styles.xml so every font is Arial; cached values are untouched."""
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
    patch_fonts(os.path.join(os.environ.get("ROOT", "/home/user/lakehouse-tracker/BRAND_US_Launch"), "output", "06_Compliance_Checklist.xlsx"))
    sys.exit(0)

ROOT = os.environ.get("ROOT", "/home/user/lakehouse-tracker/BRAND_US_Launch")
BUILD = os.path.join(ROOT, "build"); OUT = os.path.join(ROOT, "output")
facts = json.load(open(os.path.join(BUILD, "facts.json")))
fig = json.load(open(os.path.join(BUILD, "figures.json")))
sd = json.load(open(os.path.join(BUILD, "sops_data.json")))

EB, ME = "Evidence based", "Management estimate"
FONT, SZ = "Arial", 10
FILL_IN = PatternFill("solid", fgColor="FFFF00"); BLUE = "0000FF"; GREEN = "008000"
FILL_HDR = PatternFill("solid", fgColor="D9E1F2"); FILL_SEC = PatternFill("solid", fgColor="EDEDED")
FILL_TOT = PatternFill("solid", fgColor="E2EFDA")
FMT_MONEY = '$#,##0;($#,##0);0'; FMT_PCT = '0.0%'
THIN = Side(style="thin", color="BFBFBF"); BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
BUILT = fig["meta"]["built"]

# ---- facts and model numbers used in the text ----
e, k, tm, am, lg, cp, ut = facts["entity"], facts["korea"], facts["trademark"], facts["amazon"], facts["logistics"], facts["compliance"], facts["us_tax"]
wc = sd["wc_inputs"]
FORM_FEE = e["formation_cost_usd"]; WY_ANNUAL = e["annual_cost_usd"]; PEN_5472 = e["form_5472_penalty_usd"]
TM_FEE = tm["uspto_fee_per_class_usd"]; BOND = lg["broker_bond_annual_usd"]; ENTRY = lg["broker_entry_fee_usd"]
SELLER_FEE = am["professional_fee_monthly_usd"]; INS_COVER = am["product_liability_insurance_usd"]; INS_TRIG = am["insurance_trigger_monthly_sales_usd"]
INS_YEAR = wc["insurance_per_year"]["Medium"]; PREP_5472 = wc["form_5472_prep"]["Medium"]
OFA_KRW = k["overseas_financial_account_report_threshold_krw"]
INS_MONTH = sd["insurance_trigger_month"]["Medium"]
M = fig["scenarios"]["Medium"]
COMP_Y2 = M["annual"]["Y2"]["compliance"]; OPEN_CAP = M["cash"]["opening_capital"]; GAP = M["cash"]["funding_gap"]; GAP_MONTH = M["cash"]["funding_needed_by_month"]
LAUNCH_ONE = facts["august"]["m18_assumptions"]["launch_one_time"]; FIXED_M = facts["august"]["m18_assumptions"]["fixed_monthly"]
QC_UNIT = lg["qc_label_per_unit_usd"]
ENTRIES_Y1 = 8          # management estimate: about 8 formal entries in the first 12 months (launch air lots, then monthly POs)
PH = lambda s: "[PLACEHOLDER: " + s + "]"

money = lambda x: F.money(x)

# ---- the checklist rows: Area, Item, Jurisdiction, Owner, Due, Cost (USD first 12 months), Evidence required, Notes, Basis ----
R = []
def add(area, item, jur, owner, due, cost, evidence, notes, basis):
    assert basis in (EB, ME)
    R.append([area, item, jur, owner, due, float(cost), evidence, notes, basis])

A = "US entity"
add(A, "Decide the US brand name so the legal name BRAND LLC is final before any filing", "US (Wyoming)", "Lucy, Mark", "Week 1", 0,
    "Written decision in 11_Decision_Log.md", "One time. The Korean deck used MBF; nothing is filed until the name is chosen. The same spelling is then used at the state, the IRS, the bank and Amazon.", EB)
add(A, "Check name availability with the Wyoming Secretary of State", "US (Wyoming)", "Lucy, attorney", "Week 1", 0,
    "Screenshot of the availability search", "One time. Free search. Do it the same day as the trademark clearance search so both names agree.", ME)
add(A, "File the articles of organization for a Wyoming single member LLC", "US (Wyoming)", "Lucy, attorney or formation service", "Weeks 1 to 2", FORM_FEE,
    "Certificate of organization from the Secretary of State", "One time. State fee " + money(FORM_FEE) + " (brief). No state income tax; strong privacy. Delaware costs about " + money(e["delaware_annual_cost_usd"]) + " per year and is only worth it for raising investment.", EB)
add(A, "Appoint a Wyoming registered agent", "US (Wyoming)", "Lucy", "Weeks 1 to 2", 125,
    "Registered agent agreement and the agent's address on the articles", "Per year. Fee is a management estimate " + PH("agent quote") + ". The agent receives state and legal mail; set the renewal reminder with the annual report.", ME)
add(A, "Sign a single member operating agreement naming Lucy (Hyejin Yoo) as 100% owner", "US (Wyoming)", "Lucy, attorney", "Weeks 1 to 2", 300,
    "Signed operating agreement in the entity folder", "One time. Attorney fee is a management estimate " + PH("attorney quote") + ". Banks and Amazon ask for it. Mark holds no equity (DECISIONS.md decision 2); if he did, the LLC becomes a partnership with Form 1065 and 37% withholding on Lucy's share.", EB)
add(A, "Obtain the EIN by phone or fax with Form SS-4 (no SSN needed)", "US (federal, IRS)", "Lucy, attorney", "Weeks 1 to 2", 0,
    "IRS EIN confirmation letter " + PH("EIN"), "One time. Free. The EIN is the importer of record number for CBP and the tax ID for the bank and Amazon.", EB)
add(A, "Exact legal name match across state, IRS, bank and Amazon", "US", "Lucy, Mark", "Weeks 1 to 4, checked at every registration", 0,
    "One page showing the name on the certificate, the EIN letter, the bank letter and Seller Central", "Ongoing. A single spelling difference (LLC versus L.L.C., a missing comma) blocks Amazon identity verification.", EB)
add(A, "Record that no BOI report is needed (FinCEN exemption for US formed companies, finalized August 11, 2026)", "US (federal, FinCEN)", "Lucy, attorney", "Week 2", 0,
    "Note in the entity folder citing the exemption", "One time. No filing. Revisit only if the rule changes or the LLC is reorganized under a foreign parent.", EB)
add(A, "Written consulting or commission agreement between BRAND LLC and Mark Taylor (no equity, no authority to negotiate or sign)", "US", "Lucy, Mark, attorney", "Before the Amazon account opens (week 3)", 500,
    "Signed agreement", "One time. Attorney fee is a management estimate " + PH("attorney quote") + ". The wording keeps the treaty position: Mark is not a dependent agent of the LLC.", EB)
add(A, "No US office, no US employees, no US agent who negotiates or signs contracts (treaty position policy)", "US", "Lucy, Mark", "Ongoing, reviewed every quarter", 0,
    "Quarterly confirmation line in the compliance review", "Ongoing. Goods stored and shipped from FBA warehouses do not create a permanent establishment (1976 US Korea tax treaty, Article 9(3)(b)). Any hire or lease in the US reopens the question with the CPA first.", EB)
add(A, "US business mailing address for the LLC (registered agent or virtual mailbox)", "US", "Lucy, Mark", "Week 2", 120,
    "Mailbox agreement; address used on the IRS, bank and Amazon forms", "Per year. Management estimate " + PH("mailbox quote") + ". A mailbox is not an office and does not create a permanent establishment; do not use Mark's home or business address as the LLC's place of business.", ME)
add(A, "Wyoming annual report and license tax", "US (Wyoming)", "Lucy, registered agent", "Annually, formation anniversary month " + PH("exact date from the registered agent"), WY_ANNUAL,
    "Filing receipt from the Secretary of State", "Per year, about " + money(WY_ANNUAL) + " (brief). Missing it costs good standing and can freeze the bank or Amazon account.", EB)
add(A, "Open the USD business bank account in the LLC's name (Mercury or Relay; Payoneer or WorldFirst as fallback)", "US", "Lucy", "Weeks 2 to 3", 0,
    "Account opening letter with the exact legal name and EIN", "One time. Bank fees " + PH("bank fee schedule") + ". Needs the certificate, EIN letter and operating agreement. Capital arrives only after the Korean ODI report is accepted.", EB)

A = "US tax"
add(A, "Engage a US CPA familiar with foreign owned single member LLCs", "US (federal)", "Lucy, CPA", "Week 4", PREP_5472,
    "Engagement letter", "Per year. " + money(PREP_5472) + " covers Form 5472 preparation (management estimate, model input). CPA is a US certified public accountant.", ME)
add(A, "File Form 5472 with a pro forma Form 1120 for each calendar year", "US (federal, IRS)", "CPA, Lucy", "Annually by April 15 (extension through the CPA " + PH("confirm") + ")", 0,
    "Filed copy and proof of transmission", "Per year; preparation cost sits in the CPA row. Penalty for missing it: " + money(PEN_5472) + " (brief). Required for every foreign owned single member LLC even with no income.", EB)
add(A, "Keep the reportable transaction log for Form 5472 (capital, owner draws, fees between Lucy, H2 Company and BRAND LLC)", "US (federal, IRS)", "Mark, CPA", "Every quarter", 0,
    "Quarter end list sent to the CPA", "Ongoing. Every amount that crosses between the owner, H2 Company and the LLC is reportable.", EB)
add(A, "Treaty position memo: business profits taxed only where the permanent establishment sits; FBA storage creates no PE", "US (federal)", "CPA, Lucy", "Month 1", 0,
    "CPA memo in the tax folder citing Article 9(3)(b)", "One time, refreshed if the facts change. Cost is inside the CPA engagement. The memo lists the three conditions: no US office, no US employees, no agent who negotiates or signs.", EB)
add(A, "CPA question: does Lucy need an ITIN or a US individual return (Form 1040-NR) with no permanent establishment", "US (federal, IRS)", "CPA, Lucy", "Month 1", 0,
    "CPA written answer " + PH("CPA answer"), "One time. Not asserted in this plan; the CPA decides. ITIN is the individual taxpayer identification number.", ME)
add(A, "California income tax nexus question from FBA inventory stored in California", "US (California)", "CPA", "Month 1, then annually", 0,
    "CPA written answer and any California filing " + PH("CPA answer"), "Per year. Amazon collects marketplace sales tax, but California may assert income tax nexus from FBA inventory (brief). Any California filing fee is " + PH("CPA quote") + ".", EB)
add(A, "Confirm that no state sales tax permit is needed when Amazon collects marketplace sales tax", "US (states)", "CPA", "Month 1", 0,
    "CPA written answer " + PH("CPA answer"), "One time. Revisit when the DTC site starts in Year 3, because direct sales are not marketplace sales.", ME)
add(A, "Collect Form W-9 from Mark and issue Form 1099-NEC for consulting fees paid in the year", "US (federal, IRS)", "CPA, Lucy", "Annually by January 31", 0,
    "Filed Form 1099-NEC copy " + PH("CPA confirms the threshold and form"), "Per year. Cost inside the CPA engagement. Applies to the written consulting agreement, not to equity.", ME)
add(A, "Owner draws: document each profit repatriation to Korea as a draw, not salary", "US (federal)", "Lucy, CPA", "Each draw", 0,
    "Bank transfer record tagged as owner draw", "Ongoing. The draw is a reportable transaction on Form 5472 and a remittance event on the Korean side.", ME)

A = "Korea foreign exchange"
add(A, "Report the overseas direct investment (ODI) to the designated foreign exchange bank before any capital moves (any stake of 10% or more, any amount)", "Korea (Foreign Exchange Transactions Act)", "Lucy, 세무사", "Before capital transfer", 0,
    "Bank acceptance of the ODI report", "One time. No bank filing fee; 세무사 fee is in its own row. Hard stop: no transfer to the US account before acceptance. Enforcement against individual online sellers is rising (brief).", EB)
add(A, "Designate the foreign exchange bank and open the ODI file in Lucy's name (sole proprietorship, registration 599-64-00496)", "Korea", "Lucy, 세무사", "Week 1", 0,
    "Bank confirmation of the designated bank", "One time. The registration number's middle digits 64 mean H2 Company is a sole proprietorship, so the investor is Lucy personally unless H2 is incorporated first (see Korea tax).", EB)
add(A, "Remittance report after the capital is sent", "Korea (Foreign Exchange Transactions Act)", "Lucy, 세무사", "At the capital transfer " + PH("bank deadline"), 0,
    "Bank acknowledgment and the transfer confirmation", "One time per transfer. Keep the model's opening capital (" + money(OPEN_CAP) + ") and any later funding (" + F.money_k(GAP) + " funding gap by " + F.month_label(GAP_MONTH) + " in the Medium case) each as its own reported transfer.", EB)
add(A, "Securities acquisition report after the capital is paid into the LLC", "Korea (Foreign Exchange Transactions Act)", "Lucy, 세무사", "After the capital lands " + PH("bank deadline"), 0,
    "Bank acknowledgment", "One time per capital event. The LLC membership interest is the security acquired.", EB)
add(A, "Annual business performance report on the overseas direct investment", "Korea (Foreign Exchange Transactions Act)", "Lucy, 세무사", "Annually " + PH("date set by the foreign exchange bank"), 0,
    "Bank acknowledgment with the LLC's financial statements attached", "Per year. Needs the LLC's year end numbers from the monthly close (05_Operating_Processes.docx section 7).", EB)
add(A, "Change report when ownership, capital amount or the LLC's business changes", "Korea (Foreign Exchange Transactions Act)", "Lucy, 세무사", "Before any change (for example a second capital transfer or an investor)", 0,
    "Bank acknowledgment", "As needed. The funding plan's line of credit or investor amount is a change event; report it before the money moves.", ME)
add(A, "세무사 engagement for the foreign exchange filings", "Korea", "Lucy, 세무사", "Week 1", 300,
    "Engagement confirmation", "First year. Fee is a management estimate " + PH("세무사 quote") + ". 세무사 is the Korean licensed tax accountant.", ME)
add(A, "Keep every ODI acceptance and follow up acknowledgment in one folder for five years", "Korea", "Lucy", "Ongoing", 0,
    "Folder index", "Ongoing. Retention period to be confirmed by the 세무사 " + PH("retention period") + ".", ME)

A = "Korea tax"
add(A, "Incorporate H2 Company first? Put the question to the 세무사 before the ODI report", "Korea", "Lucy, 세무사", "Week 1", 0,
    "세무사 written advice " + PH("advice"), "One time. The plan assumes Lucy invests personally. Incorporating first changes the investor, the reports and the Korean tax on the LLC's profit.", EB)
add(A, "Korean tax treatment of the LLC's profit and of owner draws (global income reporting, foreign tax credit)", "Korea", "세무사", "Before the first draw " + PH("advice"), 0,
    "세무사 written advice", "One time, then yearly. Not asserted in this plan; the 세무사 decides how the US LLC income is reported in Korea.", ME)
add(A, "Overseas financial account report if the month end aggregate of foreign accounts passes KRW " + F.num(OFA_KRW / 1e6) + "M in any month", "Korea", "Lucy, 세무사", "Each year, date per 세무사", 0,
    "Filing receipt or a note that the threshold was not met", "Per year. Threshold " + F.krw(OFA_KRW) + " month end aggregate (brief). Check the LLC bank balance plus Amazon reserves every month end.", EB)
add(A, "Monthly bookkeeping summary of the LLC for the 세무사", "Korea", "Mark, Lucy", "Monthly, week 4", 0,
    "Monthly pack sent", "Ongoing. Same pack as the US monthly close; the 세무사 reads it in Korean (10_One_Page_Summary_KO.pdf explains the structure).", ME)
add(A, "Korean VAT and invoicing treatment of any services or goods H2 Company supplies to BRAND LLC", "Korea", "세무사", "Before the first H2 to LLC invoice " + PH("advice"), 0,
    "세무사 written advice", "One time. The base plan has the LLC buying from the factory directly; this row only matters if H2 invoices the LLC.", ME)
add(A, "세무사 annual fee for the Korea side filings", "Korea", "Lucy", "Annually", 600,
    "Invoice", "Per year. Management estimate " + PH("세무사 quote") + ".", ME)

A = "Amazon account"
add(A, "Open the Professional seller account in the LLC's exact legal name", "Amazon US", "Lucy, Mark", "Weeks 3 to 4", round(SELLER_FEE * 12, 2),
    "Account approval email", "Per month " + F.money(SELLER_FEE, 2) + " (brief), shown here for 12 months. Lucy is the primary contact and beneficial owner on the account.", EB)
add(A, "Identity verification (owner ID, proof of address, business documents, video call)", "Amazon US", "Lucy", "Weeks 3 to 4", 0,
    "Verification completed notice", "One time. Documents must show the same legal name as the certificate and the bank letter.", EB)
add(A, "Tax interview as a foreign owned entity (Form W-8BEN or W-8BEN-E as the interview directs)", "Amazon US", "Lucy, CPA", "Weeks 3 to 4", 0,
    "Tax interview confirmation " + PH("CPA confirms which form"), "One time. A single member LLC owned by a nonresident is disregarded for US tax, so the CPA confirms the form before Lucy answers the interview.", EB)
add(A, "Link the LLC's USD bank account for payouts (payout lag " + str(am["payout_lag_days"]) + " days)", "Amazon US", "Lucy", "Week 4", 0,
    "Deposit method shown as verified in Seller Central", "One time. The " + str(am["payout_lag_days"]) + " day lag is a working capital input in the model.", EB)
add(A, "Two step verification and user permissions (Lucy owner, Mark admin, VA limited)", "Amazon US", "Lucy, Mark", "Week 4", 0,
    "User list screenshot", "One time. The VA gets only the permissions the daily checks need.", ME)
add(A, "Product identifiers: GS1 barcodes (GTIN) for every SKU or an approved exemption", "Amazon US", "Mark", "Week 5, before listing creation", 250,
    "GS1 certificate or Amazon exemption approval " + PH("GS1 prefix"), "First year. Management estimate " + PH("GS1 US quote") + ". Amazon checks the brand on the barcode record against Brand Registry.", ME)
add(A, "Category and listing approval for the electronics accessories category", "Amazon US", "Mark", "Week 5", 0,
    "Approval or confirmation that the category is open " + PH("category status"), "One time. If an invoice is requested, the factory commercial invoice in the LLC's name serves.", ME)
add(A, "Dangerous goods review: confirm the AOC has no battery and is not hazmat", "Amazon US", "Mark, factory", "Week 5", 0,
    "Exemption sheet or safety data from the factory " + PH("document"), "One time. The active heads draw power from the HDMI port; no battery is in the product.", ME)
add(A, "Product liability insurance certificate uploaded to Seller Central once monthly sales pass " + money(INS_TRIG), "Amazon US", "Lucy", "Month " + str(INS_MONTH) + " (first month sales pass " + money(INS_TRIG) + " in the Medium case)", 0,
    "Certificate of insurance showing Amazon as additional insured, upload confirmation", "One time, then at each renewal. Policy cost is in the insurance area.", EB)
add(A, "Account Health, Performance Notifications and Voice of the Customer checked every working day", "Amazon US", "VA", "Daily from launch", 0,
    "Daily log line; any warning escalated to Mark the same day", "Ongoing. A suppressed listing or an intellectual property complaint open more than " + str(fig["gate_thresholds"]["d90_ip_days"]["value"]) + " days is a Day 90 kill trigger.", EB)
add(A, "Listing text follows the HDMI naming rule and the marketing language rule", "Amazon US", "Mark", "Before each listing goes live", 0,
    "Listing review sign off against 09_Amazon_Listing_Kit.docx", "Ongoing. Category name Ultra High Speed HDMI Cable, no version numbers; 48Gbps and 8K allowed in titles; fiber optic HDMI or hybrid active optical, never pure fiber or all glass.", EB)
add(A, "Insert card and review requests comply with Amazon's communication and review policies", "Amazon US", "Mark", "Before the first inbound shipment", 0,
    "Insert card proof approved", "One time. No incentives for reviews, no steering to off Amazon support for Amazon orders beyond what policy allows " + PH("policy text check") + ".", ME)
add(A, "FBA inbound: FNSKU labels applied at the factory, carton labels, prep requirements", "Amazon US", "Factory, 3PL, Mark", "Before the first shipment", 0,
    "Shipment plan accepted; label proof", "Per shipment. QC and labeling at the factory costs about " + F.money(QC_UNIT, 2) + " per unit in the model (management estimate), shown in the model not here.", EB)
add(A, "Amazon Vine enrollment for the first parent ASIN (optional review velocity lever)", "Amazon US", "Mark", "Week 10, after Brand Registry", 200,
    "Enrollment confirmation", "One time per parent. Fee is a management estimate " + PH("current Vine fee") + ". Lever 5 in the model.", ME)

A = "Brand and trademark"
add(A, "Trademark clearance search on the chosen name", "US (federal, USPTO)", "Attorney", "Week 1", 500,
    "Attorney search report", "One time. Management estimate " + PH("attorney quote") + ". Foreign applicants must file through a US licensed attorney (brief).", ME)
add(A, "File the USPTO trademark application in the LLC's name (about " + money(TM_FEE) + " per class plus attorney fees)", "US (federal, USPTO)", "Attorney, Lucy", "Weeks 1 to 2", TM_FEE,
    "USPTO serial number and filing receipt " + PH("serial number"), "One time. Government fee " + money(TM_FEE) + " per class (brief); one class assumed " + PH("attorney confirms the class") + ". Attorney fee in its own row.", EB)
add(A, "Attorney fee for preparing and filing the application", "US (federal, USPTO)", "Attorney", "Weeks 1 to 2", 1000,
    "Attorney invoice", "One time. Management estimate " + PH("attorney quote") + ".", ME)
add(A, "Enroll in Amazon Brand Registry on the pending application", "Amazon US", "Mark, Lucy", "Weeks 3 to 4, after the serial number arrives", 0,
    "Brand Registry approval email", "One time. Brand Registry accepts a pending application (brief) and gates A+ content, the Brand Store, video ads and hijacker protection.", EB)
add(A, "Respond to USPTO office actions within their deadlines", "US (federal, USPTO)", "Attorney, Lucy", "As issued " + PH("attorney docket"), 0,
    "Office action and the filed response", "As needed. Attorney fee " + PH("attorney quote") + ". A missed response abandons the application and Brand Registry lapses with it.", EB)
add(A, "Registration certificate received; update Brand Registry from pending to registered", "US (federal, USPTO)", "Attorney, Mark", "When issued " + PH("expected timing from the attorney"), 0,
    "Registration certificate and registration number " + PH("number"), "One time. Then use the registered mark symbol on packaging; use TM until then.", EB)
add(A, "Trademark maintenance filings between years 5 and 6 after registration", "US (federal, USPTO)", "Attorney, Lucy", "Year 5 " + PH("attorney docket"), 0,
    "Filed maintenance documents", "One time in Year 5. The model carries a placeholder of " + money(0) + " until the attorney quotes.", ME)
add(A, "Register the brand domain name and lock the social handles", "US", "Mark", "Week 2", 20,
    "Registrar receipt", "Per year. Management estimate " + PH("registrar price") + ". Needed for the Spec Match finder and trade up registration in Year 3.", ME)
add(A, "Packaging and listing artwork carries the brand mark consistently (TM until registered)", "US", "Mark, factory", "Before the first PO", 0,
    "Approved artwork proofs", "One time per artwork version.", ME)

A = "HDMI licensing"
add(A, "Confirm the factory is a current HDMI Adopter and get the certificate copy", "Industry (HDMI LA)", "Lucy, factory", "Before the first PO", 0,
    "HDMI Adopter certificate " + PH("certificate number from the factory"), "One time, rechecked yearly. HDMI LA (HDMI Licensing Administrator) runs seizures of unlicensed product (brief).", EB)
add(A, "Collect the ATC test report for each cable family (8K line and 4K line)", "Industry (HDMI LA)", "Lucy, factory", "Before the first PO", 0,
    "ATC test reports " + PH("report numbers"), "One time per family. ATC is the HDMI Authorized Test Center. Each new length or head design needs the report that covers it.", EB)
add(A, "Packaging and promotion use the category name only: Ultra High Speed HDMI Cable for the 8K line", "US", "Mark, factory", "Before artwork approval", 0,
    "Artwork proof sign off", "One time per artwork version. No HDMI version numbers on the box or in promotion (brief).", EB)
add(A, "4K line category name: Premium High Speed HDMI Cable pending certification", "Industry (HDMI LA)", "Lucy, factory", "Before the 4K 30M listing goes live", 0,
    "Certification evidence from the factory " + PH("Premium High Speed certification status"), "One time. Placeholder until the factory shows the certification; until then the 4K listing uses no category claim it cannot support.", EB)
add(A, "HDMI logo and word mark use on packaging follows the adopter's license terms", "Industry (HDMI LA)", "Lucy, factory, attorney", "Before artwork approval", 0,
    "Written confirmation from the factory on which logos the artwork may carry " + PH("logo terms"), "One time. The brand is a reseller of the factory's licensed product; the factory's license terms decide the logo use.", ME)
add(A, "Annual recheck of the factory's adopter status and the test reports on file", "Industry (HDMI LA)", "Lucy", "Annually", 0,
    "Dated note in the compliance folder", "Per year. A lapsed adopter puts every unit in the FBA network at risk of seizure.", ME)

A = "FCC"
add(A, "Obtain the factory's FCC test report for the active heads (the AOC heads contain active electronics)", "US (federal, FCC)", "Lucy, factory", "Before the first PO", 0,
    "FCC test report " + PH("report number and test lab"), "One time per head design. Brief: obtain the factory's FCC test report. If the factory has none, testing is a management estimate " + PH("test lab quote") + ".", EB)
add(A, "FCC compliance statement and any required marking on the product, packaging or manual", "US (federal, FCC)", "Mark, factory", "Before artwork approval", 0,
    "Artwork proof with the statement " + PH("wording from the test lab"), "One time. Wording and placement come from the test lab, not from this plan.", ME)
add(A, "Retest when the head changes (direction proof connector, detachable head for conduit pulls)", "US (federal, FCC)", "Lucy, factory", "Before any redesigned head ships", 1500,
    "New test report " + PH("report number"), "As needed. Management estimate " + PH("test lab quote") + ". Budgeted once in the first year for the direction proof connector (lever 3).", ME)
add(A, "Keep the FCC report on file and attach it to the listing compliance record in Seller Central if requested", "Amazon US", "Mark", "Week 6", 0,
    "Report stored; Seller Central compliance request answered within its deadline", "Ongoing.", ME)

A = "Prop 65"
add(A, "Factory chemical declaration for the jacket, connectors and packaging (listed substances)", "US (California)", "Lucy, factory", "Before the first PO", 0,
    "Factory declaration or test report " + PH("document"), "One time per material change. If the factory cannot declare, testing is a management estimate " + PH("test lab quote") + ".", ME)
add(A, "Add the California Prop 65 warning label to the packaging", "US (California)", "Mark, factory", "Before artwork approval", 0,
    "Artwork proof with the label", "One time per artwork version. Brief: add a California Prop 65 label. Cost is inside the packaging print run.", EB)
add(A, "Set the Prop 65 warning attribute on every Amazon listing", "Amazon US", "Mark", "Before each listing goes live", 0,
    "Listing attribute screenshot", "Ongoing. The listing warning must match the label on the box.", ME)
add(A, "Insert the same warning in the DTC site product pages when the site launches", "US (California)", "Mark", "Year 3, before the DTC site goes live", 0,
    "Page screenshot", "One time. Same wording as the packaging.", ME)

A = "CL3"
add(A, "Get the CL3 documentation from the factory before any in wall claim", "US (fire rating, Class 3)", "Lucy, factory", "Before the in wall claim appears anywhere", 0,
    "CL3 listing or test documentation " + PH("document and listing number"), "One time per cable construction. Brief: get CL3 documentation before claiming in wall. CL3 is the in wall cable fire rating (Class 3).", EB)
add(A, "Hold the in wall claim out of titles, bullets, images and packaging until the documentation is on file", "Amazon US", "Mark", "Launch, until cleared", 0,
    "Listing review sign off", "Ongoing. The differentiator stays in the plan; the claim waits for the evidence.", EB)
add(A, "Plenum and conduit pull documentation for the pro line", "US (fire rating)", "Lucy, factory", "Year 4, before the pro line PO", 2000,
    "Documentation per pro SKU " + PH("documents"), "One time. Management estimate " + PH("test or listing quote") + " for the pro line (CL3, plenum, conduit pull, long run fiber KVM).", ME)

A = "Insurance"
add(A, "Product liability insurance quote (" + F.money_m(INS_COVER) + " cover) obtained before launch", "US", "Lucy, Mark", "Week 6", 0,
    "Written quote " + PH("insurer and premium"), "One time. The quote is in hand so the policy can bind the month sales pass " + money(INS_TRIG) + ".", EB)
add(A, "Bind the product liability policy in the first month sales pass " + money(INS_TRIG) + " (Medium case Month " + str(INS_MONTH) + ")", "US", "Lucy", "Month " + str(INS_MONTH), INS_YEAR,
    "Policy document and certificate of insurance", "Per year, " + money(INS_YEAR) + " (management estimate, model input). Required by Amazon once sales exceed " + money(INS_TRIG) + " in a month (brief).", ME)
add(A, "Renew the product liability policy 12 months after binding", "US", "Lucy", "Month " + str(INS_MONTH + 12), 0,
    "Renewed policy", "Per year from the second year; the first year premium is in the binding row.", ME)
add(A, "Cargo and transit insurance for each air and ocean shipment", "International", "Lucy, broker", "Each shipment", 300,
    "Certificate per shipment", "First year total is a management estimate " + PH("forwarder quote") + " across the launch lots and the first ocean shipments.", ME)
add(A, "Insurance renewal and certificate dates on the compliance calendar", "US", "Mark", "Monthly, week 4", 0,
    "Calendar entry", "Ongoing.", ME)

A = "Customs"
add(A, "Engage a licensed customs broker and sign the customs power of attorney", "US (federal, CBP)", "Lucy, broker", "Week 1", 0,
    "Signed power of attorney; broker engagement letter", "One time. Broker is paid per entry (row below). CBP is US Customs and Border Protection.", EB)
add(A, "Importer of record registration with CBP under the LLC's EIN (CBP Form 5106 through the broker)", "US (federal, CBP)", "Broker, Lucy", "Week 1", 0,
    "Broker confirmation of the importer number", "One time. The LLC is the importer of record; no parcel dropship model because de minimis is gone (brief).", EB)
add(A, "Post the continuous customs bond (about " + money(BOND) + " per year)", "US (federal, CBP)", "Broker, Lucy", "Week 1", BOND,
    "Bond number and surety confirmation " + PH("bond number"), "Per year, " + money(BOND) + " (brief). Renewal is on the compliance calendar.", EB)
add(A, "File the CBP binding ruling request on classification (8544.70.00 optical fiber cables versus 8544.42.90 electric conductors)", "US (federal, CBP)", "Broker, Mark", "Week 1 (decision in 30 to 90 days)", 250,
    "Ruling request receipt, then the ruling letter " + PH("ruling number"), "One time. Broker preparation fee is a management estimate " + PH("broker quote") + ". Copper rate " + F.pct(facts["duty"]["copper"]["total"]) + ", fiber rate " + F.pct(facts["duty"]["fiber"]["total"]) + "; the 12.5% forced labor Section 301 rate applies either way. Lever 1 in the model.", EB)
add(A, "Act on the ruling: update the HTS line, set the DutyScenario selector in 03_Financial_Model_5Y.xlsx, request refunds on earlier entries if fiber", "US (federal, CBP)", "Broker, Mark", "When the ruling arrives", 0,
    "Ruling letter filed; refund route confirmed " + PH("broker route for refunds"), "One time. HTS is the Harmonized Tariff Schedule.", EB)
add(A, "Formal entry for every China origin shipment (de minimis is gone)", "US (federal, CBP)", "Broker", "Each shipment", ENTRY * ENTRIES_Y1,
    "Entry summary per shipment", "Per entry " + money(ENTRY) + " (model input); about " + str(ENTRIES_Y1) + " entries in the first 12 months is a management estimate.", EB)
add(A, "Importer Security Filing for every ocean shipment before loading", "US (federal, CBP)", "Broker, 3PL", "Each ocean shipment (from Month 7 in the Medium case)", 0,
    "Filing confirmation per shipment " + PH("broker fee per filing"), "Per ocean shipment. Air launch lots do not need it. Fee " + PH("broker quote") + ".", ME)
add(A, "Country of origin marking (Made in China) on each unit and carton, confirmed with the broker", "US (federal, CBP)", "Broker, factory", "Before artwork approval", 0,
    "Artwork proof and broker confirmation " + PH("broker confirmation"), "One time per artwork version.", ME)
add(A, "Commercial invoice and packing list in the LLC's name with HTS line, unit values in USD and FOB terms", "US (federal, CBP)", "Factory, broker", "Each shipment", 0,
    "Document set per shipment handed to the broker the same day", "Per shipment. The customs value is FOB China; the model uses the H2 cost base times " + F.num(facts["fx"]["fob_factor_vs_korea_landed"], 2) + " (management estimate) until the factory quote in USD arrives.", EB)
add(A, "Section 301 status check with the broker (pending second investigation; stress case adds 10 points)", "US (federal, USTR and CBP)", "Broker, Mark", "Every quarter", 0,
    "Broker note in the quarterly review", "Ongoing. Stress rate " + F.pct(facts["duty"]["stress"]["total"]) + " in the model.", EB)
add(A, "Customs record keeping: entries, invoices, rulings and payments kept for the retention period", "US (federal, CBP)", "Mark, broker", "Ongoing", 0,
    "Folder per entry " + PH("broker confirms the retention period"), "Ongoing.", ME)
add(A, "Duty payment funding: duty and fees paid at entry are in the cash model (in transit inventory funded by the business)", "US (federal, CBP)", "Lucy, Mark", "Each shipment", 0,
    "Entry duty amount matched to the bank payment", "Per shipment. Duty amounts come from the SKU Unit Economics tab, not from this sheet.", EB)

A = "Product safety and labeling"
add(A, "Direction proof labeling: SOURCE and DISPLAY ends color coded with printed arrows and a QR setup video", "Product", "Lucy, factory, Mark", "Before the first PO", 0,
    "Artwork proof; QC check that every sampled unit carries both labels", "One time per design. Differentiator from the brief; the return rate lever (3% to 2%).", EB)
add(A, "Written terms for the 5 Year Signal Guarantee (replacement shipped before the return arrives)", "US", "Mark, attorney", "Before launch", 300,
    "Published warranty terms on the insert card and the listing " + PH("attorney review of US written warranty rules"), "One time. Attorney review is a management estimate " + PH("attorney quote") + ".", EB)
add(A, "Packaging marks: brand, category name, length, country of origin, Prop 65 warning, FCC statement, FNSKU, warranty line", "US", "Mark, factory", "Before artwork approval", 0,
    "Final artwork proof signed by Lucy", "One time per artwork version. One proof carries every mark in this sheet so nothing is printed twice.", EB)
add(A, "Marketing language check: fiber optic HDMI or hybrid active optical only; never pure fiber or all glass", "US", "Mark", "Before each listing or ad goes live", 0,
    "Listing and ad copy sign off", "Ongoing. Brief rule.", EB)
add(A, "Third party QC before departure with a written sample plan", "China (factory)", "Lucy, QC agency", "Each PO", 0,
    "QC report per PO " + PH("AQL level agreed with the agency"), "Per PO. Cost is in the model per unit (" + F.money(QC_UNIT, 2) + " QC and labeling, management estimate), not here.", EB)
add(A, "Product safety incident log and recall contact plan", "US", "Mark", "Before launch", 0,
    "One page plan in the compliance folder", "One time. Any injury or fire report goes to Lucy and the insurer the same day " + PH("insurer notice terms") + ".", ME)
add(A, "Packaging waste and recycling marks where the retailer or state requires them", "US (states)", "Mark, factory", "Year 2, before the Walmart listing", 0,
    "Artwork proof " + PH("state and retailer requirements check"), "One time. Not asserted here; checked before the second channel.", ME)

A = "Data and privacy (placeholder)"
add(A, "Privacy policy and terms for the Spec Match finder and trade up registration before any email is collected", "US (federal and states)", "Mark, attorney", "Year 3, before the DTC site and email list go live", 500,
    "Published policy " + PH("attorney review"), "One time. Management estimate " + PH("attorney quote") + ". Placeholder area: no customer data is collected off Amazon until this row is Done.", ME)
add(A, "Amazon customer data: no use of Amazon buyer data for off Amazon marketing", "Amazon US", "Mark, VA", "Ongoing from launch", 0,
    "Policy acknowledgment by the VA " + PH("current Amazon data protection policy text"), "Ongoing. The email list is built only from people who register on the brand site.", ME)
add(A, "Commercial email rules (consent, unsubscribe, sender address) for the off Amazon list", "US (federal)", "Mark", "Year 3, before the first campaign", 0,
    "Email platform settings screenshot " + PH("attorney confirms the rules"), "One time. Not asserted here.", ME)
add(A, "State privacy law applicability check (for example California) once the list exists", "US (states)", "Attorney", "Year 3", 0,
    "Attorney note " + PH("thresholds and applicability"), "One time. Not asserted here; most thresholds depend on revenue and record counts.", ME)
add(A, "Access control and password manager for Seller Central, bank, registrar and email platform", "US", "Mark", "Week 4", 60,
    "Password manager account; access list", "Per year. Management estimate " + PH("tool price") + ".", ME)

AREAS = ["US entity", "US tax", "Korea foreign exchange", "Korea tax", "Amazon account", "Brand and trademark", "HDMI licensing", "FCC", "Prop 65", "CL3", "Insurance", "Customs", "Product safety and labeling", "Data and privacy (placeholder)"]
STATUSES = ["Not started", "In progress", "Done", "Blocked"]
assert set(r[0] for r in R) == set(AREAS), set(r[0] for r in R) ^ set(AREAS)
assert len(R) >= 60, len(R)

# ---------- workbook ----------
wb = Workbook()
def W(ws, r, c, v, bold=False, color="000000", fill=None, fmt=None, wrap=False, size=None, italic=False, align=None, border=False):
    cell = ws.cell(row=r, column=c, value=v)
    cell.font = Font(name=FONT, size=size or SZ, bold=bold, italic=italic, color=color)
    if fill: cell.fill = fill
    if fmt: cell.number_format = fmt
    cell.alignment = Alignment(wrap_text=wrap, vertical="top", horizontal=align)
    if border: cell.border = BOX
    return cell

# ---- Cover ----
cv = wb.active; cv.title = "Cover"
cv.column_dimensions["A"].width = 30; cv.column_dimensions["B"].width = 110
r = 1
W(cv, r, 1, "06 Compliance Checklist: BRAND LLC US launch", bold=True, size=14); r += 1
W(cv, r, 1, "Built " + BUILT + ". Generator build/build_checklist.py. Facts from build/facts.json, model numbers from build/figures.json.", italic=True, color="595959"); r += 2
rows = [
    ("Purpose", "Every US, Korea, Amazon, HDMI, FCC, Prop 65, CL3, insurance, customs, product labeling and data item the launch needs, each with an owner, a due point, an estimated cost, the evidence that proves it is done, a status and the basis of the entry."),
    ("How to use", "Work in the Checklist tab. Change only the Status column (dropdown: Not started, In progress, Done, Blocked) and, when a quote arrives, the Estimated cost. The summary block at the top and the Summary tab recalculate. Mark reviews it in week 4 of every month (05_Operating_Processes.docx section 8); Lucy approves and pays; the owner files; the evidence goes into the compliance folder."),
    ("Readers", "Lucy (owner, signs everything), Mark (keeps the calendar), 세무사 (Korea foreign exchange and Korea tax rows), US CPA (US tax rows), attorney (brand and trademark, agreements), customs broker (customs rows), factory (HDMI, FCC, Prop 65, CL3 documents)."),
    ("Due", "Relative to the launch critical path: weeks count from the start of formation; Month 1 is the first month with inventory live on Amazon (assumed early 2027); fixed calendar deadlines (Form 5472 by April 15) are shown as dates."),
    ("Estimated cost", "USD for the first 12 months after formation, Medium case. Facts from the brief where known (state fee " + money(FORM_FEE) + ", annual report " + money(WY_ANNUAL) + ", trademark " + money(TM_FEE) + " per class, bond " + money(BOND) + ", entry fee " + money(ENTRY) + ", seller fee " + F.money(SELLER_FEE, 2) + " per month). Everything else is a management estimate marked [PLACEHOLDER: ...] in Notes until the named professional quotes. Costs carried per unit in the model (QC, broker per unit, duty) are not repeated here."),
    ("Tie to the model", "The August model carries Year 1 compliance inside the " + money(LAUNCH_ONE) + " one time launch costs and the " + money(FIXED_M) + " per month fixed costs; from Year 2 the model carries " + money(COMP_Y2) + " per year (Form 5472 preparation " + money(PREP_5472) + ", Wyoming report " + money(WY_ANNUAL) + ", bond " + money(BOND) + ", insurance " + money(INS_YEAR) + "). Opening capital " + money(OPEN_CAP) + "; Medium funding gap " + F.money_k(GAP) + " needed by " + F.month_label(GAP_MONTH) + " (03_Financial_Model_5Y.xlsx)."),
    ("Basis", "Evidence based = brief section 1, the August model, or a verified regulatory fact. Management estimate = the build team's assumption, to be confirmed by the named owner. Nothing in this file asserts a certification, license or legal status that the brief did not verify; those cells carry a labeled placeholder."),
    ("Legend", "Yellow fill with blue text = cell you edit (Status, Estimated cost). Black = fixed text or formula. Green = pulled from another sheet."),
    ("Hard stops", "No capital transfer before the ODI report is accepted by the foreign exchange bank. No PO before the HDMI Adopter certificate, ATC report and FCC report are on file. No in wall claim before the CL3 documentation. No HDMI version number on any box or promotion."),
]
for a, b in rows:
    W(cv, r, 1, a, bold=True); W(cv, r, 2, b, wrap=True)
    cv.row_dimensions[r].height = 15 * max(1, math.ceil(len(b) / 120)) + 2; r += 1
r += 1
W(cv, r, 1, "Glossary", bold=True); r += 1
GLOSS = [("AOC", "active optical cable"), ("HDMI", "High Definition Multimedia Interface"), ("HDMI LA", "HDMI Licensing Administrator"), ("ATC", "HDMI Authorized Test Center"), ("FCC", "Federal Communications Commission"), ("Prop 65", "California Proposition 65 warning label"), ("CL3", "in wall cable fire rating (Class 3)"), ("CBP", "US Customs and Border Protection"), ("HTS", "Harmonized Tariff Schedule"), ("MFN", "most favored nation duty rate"), ("Section 301", "US trade law tariffs on China origin goods"), ("EIN", "Employer Identification Number"), ("LLC", "limited liability company"), ("IRS", "Internal Revenue Service"), ("PE", "permanent establishment"), ("ODI", "overseas direct investment report (해외직접투자 신고)"), ("FTA", "Foreign Exchange Transactions Act (외국환거래법)"), ("세무사", "Korean licensed tax accountant"), ("CPA", "US certified public accountant"), ("USPTO", "United States Patent and Trademark Office"), ("FBA", "Fulfillment by Amazon"), ("FNSKU", "Fulfillment Network Stock Keeping Unit (Amazon's barcode)"), ("ASIN", "Amazon Standard Identification Number"), ("GTIN", "Global Trade Item Number (the barcode number from GS1)"), ("DTC", "direct to consumer website"), ("FOB", "free on board (the price at the factory's port)"), ("DDP", "delivered duty paid"), ("3PL", "third party logistics warehouse"), ("QC", "quality control"), ("PO", "purchase order"), ("SKU", "stock keeping unit"), ("VA", "virtual assistant"), ("KVM", "keyboard, video and mouse extension"), ("ITIN", "individual taxpayer identification number")]
for a, b in GLOSS:
    W(cv, r, 1, a); W(cv, r, 2, b); r += 1

# ---- Checklist ----
ws = wb.create_sheet("Checklist")
HDR = ["Area", "Item", "Jurisdiction", "Owner", "Due", "Estimated cost (USD, first 12 months)", "Evidence required", "Status", "Notes", "Basis"]
WIDTHS = [22, 58, 22, 20, 28, 16, 40, 13, 60, 21]
for i, w in enumerate(WIDTHS, 1): ws.column_dimensions[get_column_letter(i)].width = w
HROW = 8; FIRST = HROW + 1; LAST = FIRST + len(R) - 1; RANGE_END = LAST + 40   # 40 spare rows for items added by hand
SC = "$H$%d:$H$%d" % (FIRST, RANGE_END); CC = "$F$%d:$F$%d" % (FIRST, RANGE_END); AC = "$A$%d:$A$%d" % (FIRST, RANGE_END); NC = "$I$%d:$I$%d" % (FIRST, RANGE_END)
W(ws, 1, 1, "Compliance checklist: BRAND LLC US launch (Lucy, Hyejin Yoo, 100% owner; H2 Company, Seoul)", bold=True, size=12)
W(ws, 2, 1, "Edit the yellow cells only. Status dropdown: Not started, In progress, Done, Blocked. Spare rows below the last item take new entries and the counts include them. Built " + BUILT + ".", italic=True, color="595959")
# summary block (live COUNTIF)
W(ws, 4, 1, "Summary", bold=True, fill=FILL_SEC)
labels = ["Items", "Not started", "In progress", "Done", "Blocked", "Done %", "Cost (12 months)"]
for j, lab in enumerate(labels):
    W(ws, 4, 2 + j, lab, bold=True, fill=FILL_HDR, wrap=True, border=True)
W(ws, 5, 2, "=COUNTA(%s)" % AC, fmt="0", border=True)
for j, s in enumerate(STATUSES):
    W(ws, 5, 3 + j, '=COUNTIF(%s,"%s")' % (SC, s), fmt="0", border=True)
W(ws, 5, 7, "=IF(B5=0,0,E5/B5)", fmt=FMT_PCT, border=True)
W(ws, 5, 8, "=SUM(%s)" % CC, fmt=FMT_MONEY, border=True, fill=FILL_TOT, bold=True)
ws.row_dimensions[4].height = 30
W(ws, 6, 1, "Status check (0 = ok)", italic=True, color="595959"); W(ws, 6, 2, "=B5-SUM(C5:F5)", fmt="0", border=True)
for i, h in enumerate(HDR, 1):
    W(ws, HROW, i, h, bold=True, fill=FILL_HDR, wrap=True, border=True)
ws.row_dimensions[HROW].height = 44
CH = {2: 58, 7: 40, 9: 60, 5: 28, 3: 22, 4: 20, 1: 22, 10: 21}   # chars per line per wrapped column (about 1 char per width unit)
for n, row in enumerate(R):
    rr = FIRST + n
    area, item, jur, owner, due, cost, ev, notes, basis = row
    vals = {1: area, 2: item, 3: jur, 4: owner, 5: due, 6: cost, 7: ev, 8: "Not started", 9: notes, 10: basis}
    for c, v in vals.items():
        if c == 6:
            W(ws, rr, c, v, fmt=FMT_MONEY, color=BLUE, fill=FILL_IN, border=True)
        elif c == 8:
            W(ws, rr, c, v, color=BLUE, fill=FILL_IN, border=True)
        else:
            W(ws, rr, c, v, wrap=True, border=True)
    lines = max(math.ceil(len(str(vals[c])) / CH[c]) for c in CH)
    ws.row_dimensions[rr].height = max(15, 13.2 * lines + 3)
# spare rows: status cells styled and validated, cost cells styled
for rr in range(LAST + 1, RANGE_END + 1):
    W(ws, rr, 8, None, color=BLUE, fill=FILL_IN); W(ws, rr, 6, None, fmt=FMT_MONEY, color=BLUE, fill=FILL_IN)
W(ws, RANGE_END + 1, 1, "End of list. Add new items in the spare rows above this line so the summary counts them.", italic=True, color="595959")
dv = DataValidation(type="list", formula1='"' + ",".join(STATUSES) + '"', allow_blank=True, showErrorMessage=True, errorTitle="Status", error="Choose Not started, In progress, Done or Blocked")
ws.add_data_validation(dv); dv.add("H%d:H%d" % (FIRST, RANGE_END))
dvb = DataValidation(type="list", formula1='"' + EB + "," + ME + '"', allow_blank=True, showErrorMessage=True, errorTitle="Basis", error="Evidence based or Management estimate")
ws.add_data_validation(dvb); dvb.add("J%d:J%d" % (FIRST, RANGE_END))
ws.auto_filter.ref = "A%d:J%d" % (HROW, LAST)
ws.freeze_panes = "C%d" % FIRST
rng = "H%d:H%d" % (FIRST, RANGE_END)
ws.conditional_formatting.add(rng, FormulaRule(formula=['$H%d="Done"' % FIRST], fill=PatternFill("solid", fgColor="C6EFCE"), font=Font(name=FONT, size=SZ, color="006100")))
ws.conditional_formatting.add(rng, FormulaRule(formula=['$H%d="Blocked"' % FIRST], fill=PatternFill("solid", fgColor="FFC7CE"), font=Font(name=FONT, size=SZ, color="9C0006")))
ws.conditional_formatting.add(rng, FormulaRule(formula=['$H%d="In progress"' % FIRST], fill=PatternFill("solid", fgColor="FFEB9C"), font=Font(name=FONT, size=SZ, color="9C5700")))
ws.print_title_rows = "%d:%d" % (HROW, HROW)
ws.page_setup.orientation = "landscape"; ws.page_setup.paperSize = ws.PAPERSIZE_TABLOID
ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0; ws.sheet_properties.pageSetUpPr.fitToPage = True
ws.print_area = "A1:J%d" % (LAST + 1)
ws.page_margins.left = ws.page_margins.right = 0.3; ws.page_margins.top = ws.page_margins.bottom = 0.4

# ---- Summary ----
sm = wb.create_sheet("Summary")
for col, w in zip("ABCDEFGH", [56, 14, 14, 14, 14, 14, 18, 16]): sm.column_dimensions[col].width = w
W(sm, 1, 1, "Summary: counts by status and area, cost by area (live formulas reading the Checklist tab)", bold=True, size=12)
W(sm, 2, 1, "All counts and sums are COUNTIF, COUNTIFS and SUMIFS over the Checklist tab, including the spare rows.", italic=True, color="595959")
hdr2 = ["Area", "Items", "Not started", "In progress", "Done", "Blocked", "Estimated cost, first 12 months", "Done %"]
for i, h in enumerate(hdr2, 1): W(sm, 4, i, h, bold=True, fill=FILL_HDR, wrap=True, border=True)
sm.row_dimensions[4].height = 30
CL = "Checklist!"
for n, a in enumerate(AREAS):
    rr = 5 + n
    W(sm, rr, 1, a, border=True)
    W(sm, rr, 2, '=COUNTIF(%s%s,A%d)' % (CL, AC, rr), fmt="0", color=GREEN, border=True)
    for j, s in enumerate(STATUSES):
        W(sm, rr, 3 + j, '=COUNTIFS(%s%s,$A%d,%s%s,"%s")' % (CL, AC, rr, CL, SC, s), fmt="0", color=GREEN, border=True)
    W(sm, rr, 7, '=SUMIFS(%s%s,%s%s,$A%d)' % (CL, CC, CL, AC, rr), fmt=FMT_MONEY, color=GREEN, border=True)
    W(sm, rr, 8, '=IF(B%d=0,0,E%d/B%d)' % (rr, rr, rr), fmt=FMT_PCT, border=True)
TR = 5 + len(AREAS)
W(sm, TR, 1, "Total", bold=True, fill=FILL_TOT, border=True)
for c in range(2, 8):
    L = get_column_letter(c)
    W(sm, TR, c, "=SUM(%s5:%s%d)" % (L, L, TR - 1), fmt=(FMT_MONEY if c == 7 else "0"), bold=True, fill=FILL_TOT, border=True)
W(sm, TR, 8, "=IF(B%d=0,0,E%d/B%d)" % (TR, TR, TR), fmt=FMT_PCT, bold=True, fill=FILL_TOT, border=True)
W(sm, TR + 1, 1, "Check: area totals equal the Checklist item count (0 = ok)", italic=True, color="595959")
W(sm, TR + 1, 2, "=B%d-COUNTA(%s%s)" % (TR, CL, AC), fmt="0", color=GREEN, border=True)
r = TR + 3
W(sm, r, 1, "By status", bold=True, fill=FILL_SEC); W(sm, r, 2, "Items", bold=True, fill=FILL_HDR, border=True); W(sm, r, 3, "Share", bold=True, fill=FILL_HDR, border=True); r += 1
for s in STATUSES:
    W(sm, r, 1, s, border=True); W(sm, r, 2, '=COUNTIF(%s%s,A%d)' % (CL, SC, r), fmt="0", color=GREEN, border=True); W(sm, r, 3, "=IF($B$%d=0,0,B%d/$B$%d)" % (TR, r, TR), fmt=FMT_PCT, border=True); r += 1
r += 1
W(sm, r, 1, "By owner (an item with two owners counts under each)", bold=True, fill=FILL_SEC); W(sm, r, 2, "Items", bold=True, fill=FILL_HDR, border=True); W(sm, r, 3, "Open (not Done)", bold=True, fill=FILL_HDR, border=True); r += 1
OWNERS = ["Lucy", "Mark", "세무사", "CPA", "Attorney", "Broker", "Factory", "3PL", "VA", "Registered agent", "QC agency"]
OC = "$D$%d:$D$%d" % (FIRST, RANGE_END)
for o in OWNERS:
    W(sm, r, 1, o, border=True)
    W(sm, r, 2, '=COUNTIF(%s%s,"*%s*")' % (CL, OC, o.lower() if o in ("Attorney", "Broker", "Factory", "Registered agent") else o), fmt="0", color=GREEN, border=True)
    W(sm, r, 3, '=COUNTIFS(%s%s,"*%s*",%s%s,"<>Done")' % (CL, OC, o.lower() if o in ("Attorney", "Broker", "Factory", "Registered agent") else o, CL, SC), fmt="0", color=GREEN, border=True); r += 1
r += 1
W(sm, r, 1, "By basis", bold=True, fill=FILL_SEC); W(sm, r, 2, "Items", bold=True, fill=FILL_HDR, border=True); W(sm, r, 3, "Estimated cost", bold=True, fill=FILL_HDR, border=True); r += 1
BC = "$J$%d:$J$%d" % (FIRST, RANGE_END)
for b in (EB, ME):
    W(sm, r, 1, b, border=True); W(sm, r, 2, '=COUNTIF(%s%s,A%d)' % (CL, BC, r), fmt="0", color=GREEN, border=True); W(sm, r, 3, '=SUMIFS(%s%s,%s%s,A%d)' % (CL, CC, CL, BC, r), fmt=FMT_MONEY, color=GREEN, border=True); r += 1
r += 1
W(sm, r, 1, "Items with an open placeholder in Notes or Evidence", border=True); W(sm, r, 2, '=COUNTIF(%s%s,"*PLACEHOLDER*")+COUNTIFS(%s$G$%d:$G$%d,"*PLACEHOLDER*",%s%s,"<>*PLACEHOLDER*")' % (CL, NC, CL, FIRST, RANGE_END, CL, NC), fmt="0", color=GREEN, border=True); r += 1
W(sm, r, 1, "Hard stop items (ODI before capital, HDMI and FCC documents before the PO, CL3 before the in wall claim) not Done", border=True)
W(sm, r, 2, '=COUNTIFS(%s$B$%d:$B$%d,"*before any capital moves*",%s%s,"<>Done")+COUNTIFS(%s$B$%d:$B$%d,"*HDMI Adopter*",%s%s,"<>Done")+COUNTIFS(%s$B$%d:$B$%d,"*FCC test report for the active heads*",%s%s,"<>Done")+COUNTIFS(%s$B$%d:$B$%d,"*CL3 documentation from the factory*",%s%s,"<>Done")' % ((CL, FIRST, RANGE_END, CL, SC) * 4), fmt="0", color=GREEN, border=True)
sm.row_dimensions[r].height = 46; sm.cell(row=r, column=1).alignment = Alignment(wrap_text=True, vertical="top")
sm.freeze_panes = "A5"
sm.page_setup.orientation = "landscape"; sm.page_setup.fitToWidth = 1; sm.page_setup.fitToHeight = 0; sm.sheet_properties.pageSetUpPr.fitToPage = True
cv.page_setup.orientation = "landscape"; cv.page_setup.fitToWidth = 1; cv.page_setup.fitToHeight = 0; cv.sheet_properties.pageSetUpPr.fitToPage = True

os.makedirs(OUT, exist_ok=True)
path = os.path.join(OUT, "06_Compliance_Checklist.xlsx")
wb.save(path)
print(json.dumps({"file": path, "items": len(R), "areas": len(AREAS), "header_row": HROW, "first_row": FIRST, "last_row": LAST,
                  "total_cost_first_12_months": round(sum(r[5] for r in R), 2),
                  "evidence_based": sum(1 for r in R if r[8] == EB), "management_estimate": sum(1 for r in R if r[8] == ME)}))

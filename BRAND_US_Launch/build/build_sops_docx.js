// 05_Operating_Processes.docx generator. Run through build/build_sops.py (which writes build/sops_data.json first).
// Reads build/figures.json (model numbers), build/facts.json (static facts), build/fmt.js (formatter), build/sops_data.json.
'use strict';
const fs = require('fs');
const path = require('path');
const docx = require('docx');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, ShadingType, AlignmentType,
  HeadingLevel, LevelFormat, BorderStyle, PageNumber, Header, Footer, PageBreak, TableLayoutType, VerticalAlign,
} = docx;

const ROOT = process.env.ROOT || '/home/user/lakehouse-tracker/BRAND_US_Launch';
const BUILD = path.join(ROOT, 'build');
const OUT = path.join(ROOT, 'output');
const fig = JSON.parse(fs.readFileSync(path.join(BUILD, 'figures.json'), 'utf8'));
const facts = JSON.parse(fs.readFileSync(path.join(BUILD, 'facts.json'), 'utf8'));
const sd = JSON.parse(fs.readFileSync(path.join(BUILD, 'sops_data.json'), 'utf8'));
const F = require(path.join(BUILD, 'fmt.js'));

// ---------- shortcuts to the numbers ----------
const G = fig.gate_thresholds;
const gv = (k) => G[k].value;
const M = fig.scenarios.Medium;
const L = fig.scenarios.Low;
const H = fig.scenarios.High;
const hero = fig.unit_economics.hero;
const duty = fig.duty;
const wc = sd.wc_inputs;
const payoutLag = facts.amazon.payout_lag_days;
const coverM = wc.cover_months.Medium, coverL = wc.cover_months.Low, coverH = wc.cover_months.High;
const airLead = wc.air_lead_months.Medium, oceanLead = wc.ocean_lead_months.Medium, oceanFrom = wc.ocean_from_month.Medium;
const deposit = wc.po_deposit.Medium;
const insTriggerMonth = M.insurance_trigger_month != null ? M.insurance_trigger_month : sd.insurance_trigger_month.Medium; // figures.json, single source with the deck
const insMargin = M.insurance_trigger_margin_usd != null ? M.insurance_trigger_margin_usd : (M.monthly[insTriggerMonth - 1].amazon_revenue - facts.amazon.insurance_trigger_monthly_sales_usd);
const launchSkus = fig.unit_economics.per_sku.filter((s) => s.launch_sku === 'Yes');
const skuName = (s) => s.sku.replace('_', ' ');

// ---------- style helpers ----------
const FONT = 'Arial';
const PAGE_W = 12240, PAGE_H = 15840, MARGIN = 1080; // US Letter, 0.75 inch margins
const CONTENT_W = PAGE_W - 2 * MARGIN; // 10080 DXA
const TXT = 22; // Arial 11 pt for body, table cells and notes (BUILD_SPEC section 7)
const INK = '1F2937', MUTED = '4B5563', HEAD_FILL = 'E8EEF7', LABEL_FILL = 'F3F6FB', ACCENT = '0B4F8A', WARN_FILL = 'FFF4E0';

function run(text, opts = {}) { return new TextRun({ text, font: FONT, size: opts.size || 22, bold: opts.bold, italics: opts.italics, color: opts.color || INK }); }
function p(text, opts = {}) {
  const runs = Array.isArray(text) ? text : [run(text, opts)];
  return new Paragraph({ children: runs, spacing: { after: opts.after == null ? 120 : opts.after, before: opts.before || 0 }, alignment: opts.align, keepNext: opts.keepNext });
}
function small(text, opts = {}) { return p(text, { size: TXT, color: MUTED, ...opts }); }
function h1(text) { return new Paragraph({ text, heading: HeadingLevel.HEADING_1, spacing: { before: 360, after: 160 }, keepNext: true }); }
function h2(text) { return new Paragraph({ text, heading: HeadingLevel.HEADING_2, spacing: { before: 240, after: 120 }, keepNext: true }); }
function h3(text) { return new Paragraph({ text, heading: HeadingLevel.HEADING_3, spacing: { before: 160, after: 80 }, keepNext: true }); }
function bullet(text, ref = 'bullets') {
  const runs = Array.isArray(text) ? text : [run(text)];
  return new Paragraph({ children: runs, numbering: { reference: ref, level: 0 }, spacing: { after: 60 } });
}
function numbered(text, ref) {
  const runs = Array.isArray(text) ? text : [run(text)];
  return new Paragraph({ children: runs, numbering: { reference: ref, level: 0 }, spacing: { after: 60 } });
}
const thin = { style: BorderStyle.SINGLE, size: 4, color: 'B8C2D1' };
const borders = { top: thin, bottom: thin, left: thin, right: thin, insideHorizontal: thin, insideVertical: thin };
function cell(children, width, opts = {}) {
  const kids = (Array.isArray(children) ? children : [children]).map((c) => (typeof c === 'string' ? p(c, { size: opts.size || TXT, after: 40, bold: opts.bold, color: opts.color }) : c));
  return new TableCell({
    children: kids.length ? kids : [p('', { after: 0 })],
    width: { size: width, type: WidthType.DXA },
    shading: opts.fill ? { type: ShadingType.CLEAR, fill: opts.fill, color: 'auto' } : undefined,
    margins: { top: 60, bottom: 60, left: 100, right: 100 },
    verticalAlign: opts.valign || VerticalAlign.TOP,
    columnSpan: opts.span,
  });
}
function fit(widths) { // scale the column widths to the content width so every table fills the page
  const total = widths.reduce((a, b) => a + b, 0);
  const out = widths.map((w) => Math.round((w * CONTENT_W) / total));
  out[out.length - 1] += CONTENT_W - out.reduce((a, b) => a + b, 0);
  return out;
}
function table(widths0, rows, opts = {}) {
  const widths = fit(widths0);
  return new Table({
    width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA },
    columnWidths: widths,
    layout: TableLayoutType.FIXED,
    borders,
    rows: rows.map((r, i) => new TableRow({ children: r, tableHeader: i === 0 && opts.headerRow !== false, cantSplit: opts.cantSplit })),
  });
}
function headerRow(labels, widths) { return labels.map((t, i) => cell(p(t, { size: TXT, bold: true, after: 0, color: ACCENT }), widths[i], { fill: HEAD_FILL })); }
function gridTable(labels, widths, rows, opts = {}) {
  widths = fit(widths);
  const body = rows.map((r) => r.map((v, i) => cell(typeof v === 'string' ? p(v, { size: TXT, after: 0 }) : v, widths[i], opts.rowOpts || {})));
  return table(widths, [headerRow(labels, widths), ...body], opts);
}
function spacer(after = 120) { return new Paragraph({ children: [], spacing: { after } }); }

// ---------- SOP table ----------
let sopIndex = 0;
function sopTable(sop) {
  sopIndex += 1;
  const ref = `steps${sopIndex}`;
  const W = fit([1700, 7948]);
  const label = (t) => cell(p(t, { size: TXT, bold: true, after: 0 }), W[0], { fill: LABEL_FILL });
  const textCell = (items) => cell(items.map((t) => (typeof t === 'string' ? p(t, { size: TXT, after: 60 }) : t)), W[1]);
  const listCell = (items) => cell(items.map((t) => bullet(Array.isArray(t) ? t : [run(t, { size: TXT })])), W[1]);
  const stepCell = (items) => cell(items.map((t) => numbered(Array.isArray(t) ? t : [run(t, { size: TXT })], ref)), W[1]);
  const rows = [
    [label('Owner'), textCell(sop.owner)],
    [label('Trigger'), textCell(sop.trigger)],
    [label('Steps'), stepCell(sop.steps)],
    [label('System'), textCell(sop.system)],
    [label('KPI'), listCell(sop.kpi)],
    [label('Failure modes'), listCell(sop.failures)],
  ];
  if (sop.chart) rows.push([label('Flowchart'), textCell(sop.chart)]);
  return table(W, rows, { headerRow: false });
}

// ---------- numbering config ----------
const numbering = {
  config: [
    { reference: 'bullets', levels: [{ level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 360, hanging: 240 } } } }] },
    ...Array.from({ length: 12 }, (_, i) => ({
      reference: `steps${i + 1}`,
      levels: [{ level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 420, hanging: 320 } } } }],
    })),
  ],
};

// ---------- content ----------
const B = (t) => run(t, { bold: true });
const built = fig.meta.built;

const children = [];

// Title block
children.push(new Paragraph({ children: [run('05 Operating Processes', { size: 44, bold: true, color: ACCENT })], spacing: { before: 300, after: 120 } }));
children.push(p([run('BRAND LLC (a Wyoming limited liability company) with H2 Company (Seoul). Standard operating procedures for the US Amazon launch of BRAND fiber optic HDMI (High Definition Multimedia Interface) cables.', { size: 24 })], { after: 240 }));
children.push(p([B('Version: '), run(`built ${built}. Model basis: Medium scenario, Copper duty scenario, FX ${F.num(fig.meta.fx)} KRW per USD, from ${fig.meta.workbook}.`)]));
children.push(p([B('Readers: '), run('Lucy (owner and operator), Mark (US operations advisor), the virtual assistant (VA), the customs broker, the 3PL (third party logistics warehouse). The 세무사 (Korean licensed tax accountant) reads sections 7 and 8.')]));
children.push(p([B('How to read a number: '), run('every model figure in this document comes from build/figures.json, which is pulled from the recalculated workbook. Every assumption carries a basis label: Evidence based (the brief, the August 2026 deal, the August model, a verified regulatory fact) or Management estimate. A bracketed [PLACEHOLDER] is a fact that does not exist yet and must not be filled in by guesswork.')]));
children.push(p([B('Standing rule: '), run('no SOP (standard operating procedure) is complete until its owner has run it once and recorded the KPI (key performance indicator). Each SOP names the flowchart in output/04_Flowcharts that draws the same process.')]));

// Contents table
children.push(h2('Contents'));
const toc = [
  ['1', 'Purchasing and QC (quality control)', 'Lucy', '03_inbound_supply_chain'],
  ['2', 'Customs and duty management', 'Broker, Mark', '03_inbound_supply_chain, 08_entity_money_goods_flow'],
  ['3', 'FBA (Fulfillment by Amazon) and 3PL inbound', 'Lucy, 3PL', '03_inbound_supply_chain'],
  ['4', 'Listing creation and PPC (pay per click) management', 'Mark, VA', '02_order_to_cash_amazon'],
  ['5', 'Inventory planning and reorder points', 'Lucy, Mark', '05_monthly_operating_cadence'],
  ['6', 'Returns and warranty (5 Year Signal Guarantee)', 'Mark, VA', '04_returns_and_signal_guarantee'],
  ['7', 'Monthly close and KPI review', 'Lucy, Mark, 세무사', '05_monthly_operating_cadence, 06_day90_scale_or_kill'],
  ['8', 'Compliance calendar', 'Lucy, attorney, broker, 세무사', '08_entity_money_goods_flow'],
  ['A', 'Numbers used in this document and where they come from', '', ''],
  ['B', 'Glossary', '', ''],
];
children.push(gridTable(['No.', 'Process', 'Owner', 'Flowchart (output/04_Flowcharts)'], [600, 3500, 1750, 3798], toc));
children.push(spacer());
children.push(p([B('Owners in one line. '), run('Lucy decides and signs: purchase orders, factory payments, capital, filings. Mark prepares and runs the US side: listings, advertising, customer service, the monthly KPI pack. Mark does not negotiate or sign contracts for BRAND LLC; that keeps the US Korea treaty position (no permanent establishment, PE) clean. The VA executes daily checks under written instructions.')]));
// no forced page break here: the owners paragraph spills onto page 2 and a break would leave that page almost empty

// ---------- SOP 1: Purchasing and QC ----------
children.push(h1('1. Purchasing and QC'));
children.push(p('Buy the right units at the right time, labeled and tested so that nothing has to be touched again in the US. QC means quality control. SKU is a stock keeping unit (one length of one line). FNSKU is the Fulfillment Network Stock Keeping Unit, the barcode Amazon scans. PO is a purchase order. FOB is free on board, the price at the factory port. DDP is delivered duty paid (the freight price includes duty and delivery). MOQ is the minimum order quantity.'));
children.push(sopTable({
  owner: ['Lucy (places the PO, pays the factory, owns the factory relationship). Mark reviews the quantity against the reorder table (section 5) before Lucy signs. The third party QC agency reports to Lucy with a copy to Mark.'],
  trigger: [
    `Monthly, in the inventory review (section 5), when the reorder table shows a PO to place this month (an arrival the model needs lead time months out). Launch lots: ${facts.logistics.first_lots} (brief, evidence based).`,
    `Ad hoc: a stockout forecast inside the lead time (${airLead} month by air, ${oceanLead} months by ocean), or a factory price or MOQ change.`,
  ],
  steps: [
    `Pull the PO quantity from the reorder table (section 5): the arrival the model needs lead time months out, corrected for actuals (model target closing stock for the arrival month, minus actual on hand and in transit stock, plus the units forecast until that arrival). Split the line total by SKU share of trailing 30 day sales. Round up to the factory carton quantity. [PLACEHOLDER: carton quantity per SKU from the factory]`,
    `Confirm the FOB price per SKU in writing against the H2 cost base (${F.krw(launchSkus[1].h2_cost_krw)} landed Korea for the hero 8K 10M, evidence based) and the Alibaba benchmark of $${facts.august.alibaba_benchmark.fob_low_usd} to $${facts.august.alibaba_benchmark.fob_high_usd} FOB for a comparable 30 ft cable. If the quote is above the H2 base, stop and rerun the SKU Unit Economics tab before signing; at the Alibaba cost the hero SKU loses money at $${hero.price}.`,
    `Issue the PO with: SKU, length, quantity, FOB price, Incoterm (FOB China for ocean, DDP air for launch lots), FNSKU labels applied at the factory, carton marks, ship date, QC hold point, and the packaging rule (category name Ultra High Speed HDMI Cable on the box, never an HDMI version number).`,
    `Pay the deposit: ${F.pct(deposit, 0)} at order, ${F.pct(1 - deposit, 0)} after the QC report passes and before shipment (management estimate, model input). Pay from the BRAND LLC bank account, never from a personal account.`,
    `Collect the document set from the factory before the QC visit: HDMI Adopter certificate [PLACEHOLDER: number], ATC (HDMI Authorized Test Center) test report [PLACEHOLDER: report number], FCC (Federal Communications Commission) test report for the active heads [PLACEHOLDER], CL3 (in wall cable fire rating, Class 3) documentation [PLACEHOLDER] if in wall is claimed, Prop 65 (California Proposition 65) warning label artwork proof, packaging artwork proof.`,
    `Book the third party QC inspection before departure (budget $${F.num(facts.logistics.qc_label_per_unit_usd, 2)} per unit for QC and labeling, management estimate). Sample plan: [PLACEHOLDER: AQL (acceptable quality limit) level agreed with the agency]. Mandatory tests: signal lock at full length on an 8K 60Hz or 4K 120Hz source, both ends labeled SOURCE and DISPLAY with arrows, FNSKU scan on every unit sampled, carton drop test, label and insert check.`,
    `Decision on the QC report: pass, rework, or reject. Rework and reject hold the balance payment. Record the result in the PO log.`,
    `Release the balance payment, receive the commercial invoice, packing list and the bill of lading or air waybill, and hand the document set to the customs broker (section 2) the same day.`,
    `Log the PO: date, SKUs, units, FOB total in USD and KRW, deposit date, balance date, QC result, ship date, expected arrival. This log feeds the cash forecast (section 7).`,
  ],
  system: ['PO log (a tab in 07_90_Day_Launch_Plan.xlsx or a shared sheet), the factory quote sheet, the QC agency portal, the BRAND LLC bank account. Document folder per PO: quote, PO, deposit receipt, QC report, invoice, packing list, bill of lading, test reports.'],
  kpi: [
    'QC first pass rate: target 100% of lots pass without rework (management estimate).',
    `On time ship rate: 90% of POs ship within 7 days of the PO ship date (management estimate).`,
    `FOB price variance to the H2 cost base: 0% or better; any increase goes to the SKU Unit Economics tab the same week.`,
    `Defect rate found at QC: below 1% of units sampled (management estimate); above 2% triggers a factory corrective action request.`,
  ],
  failures: [
    'Factory ships without FNSKU labels: Amazon rejects or relabels at a fee and the lot sits at the 3PL. Fix: FNSKU on the PO, label check in the QC sample plan.',
    'Factory quote drifts toward the Alibaba benchmark: the hero SKU net falls from $' + F.num(hero.net, 2) + ' to a loss. Fix: price confirmed in writing before every PO, volume tier negotiated at 5,000 units per quarter (lever 7).',
    'HDMI version number printed on the box: HDMI LA (HDMI Licensing Administrator) can seize the lot. Fix: artwork proof signed off by Mark before production.',
    'Missing test reports (ATC, FCC, CL3): listing claims cannot be supported and an Amazon or regulator request cannot be answered. Fix: no PO without the document set, placeholders filled with real numbers.',
    'Balance paid before QC: no leverage on a failed lot. Fix: balance only after the QC pass in writing.',
  ],
  chart: ['03_inbound_supply_chain (factory PO, QC, freight, customs, 3PL, FBA). The money side is on 08_entity_money_goods_flow.'],
}));

// ---------- SOP 2: Customs and duty ----------
children.push(h1('2. Customs and duty management'));
const dc = facts.duty.copper, df = facts.duty.fiber, ds = facts.duty.stress;
children.push(p(`Every China origin shipment needs a formal customs entry; de minimis is gone (verified fact). BRAND LLC is the importer of record through a licensed customs broker with a continuous bond (about $${F.num(facts.logistics.broker_bond_annual_usd)} per year). CBP is US Customs and Border Protection. HTS is the Harmonized Tariff Schedule. MFN is the most favored nation duty rate. Section 301 is the US trade law tariff on China origin goods.`));
children.push(p('Duty scenarios the model carries (facts.json, duty table; the 12.5% forced labor action applies either way):'));
children.push(gridTable(
  ['Scenario', 'HTS', 'MFN', 'Section 301 list', 'Section 301 forced labor', 'Pending action', 'Total', 'Basis'],
  [1150, 1500, 700, 1150, 1250, 1150, 750, 1998],
  [
    [dc.name + ' (default)', dc.hts, F.pct(dc.mfn), F.pct(dc.section_301_list, 0) + ' ' + dc.section_301_list_name, F.pct(dc.section_301_forced_labor), F.pct(dc.pending_301_add, 0), F.pct(dc.total), dc.basis],
    [df.name, df.hts, F.pct(df.mfn, 0), df.section_301_list_name, F.pct(df.section_301_forced_labor), F.pct(df.pending_301_add, 0), F.pct(df.total), 'Evidence based rates; List 3 coverage is a broker check [PLACEHOLDER]'],
    [ds.name, '8544.42.90 plus the pending action', F.pct(ds.mfn), F.pct(ds.section_301_list, 0), F.pct(ds.section_301_forced_labor), F.pct(ds.pending_301_add, 0), F.pct(ds.total), ds.basis],
  ],
  { cantSplit: true },
));
children.push(spacer());
children.push(p(`What the ruling is worth (figures.json, full recalc, Medium): landed COGS (cost of goods sold) ${F.pct(duty.copper.landed_pct_medium)} at the copper rate versus ${F.pct(duty.fiber.landed_pct_medium)} at the fiber rate; hero net per unit $${F.num(duty.copper.hero_net, 2)} versus $${F.num(duty.fiber.hero_net, 2)}; 18 month net ${F.money_k(fig.full_recalc.duty.copper.m18_net)} versus ${F.money_k(fig.full_recalc.duty.fiber.m18_net)}. The stress case cuts 18 month net to ${F.money_k(fig.full_recalc.duty.stress.m18_net)}.`));
children.push(sopTable({
  owner: ['Customs broker (entries, bond, ruling request drafting). Mark (reviews every entry summary, tracks the ruling). Lucy (signs the power of attorney, pays duty and broker fees, decides the classification position with the broker).'],
  trigger: [
    `Week 1 of the launch plan: engage the broker, post the continuous bond, file the CBP binding ruling request (${facts.logistics.binding_ruling_timing}).`,
    'Every shipment: the factory document set arrives (section 1 step 8).',
    'Any tariff change notice from the broker, including a decision in the pending second Section 301 investigation.',
  ],
  steps: [
    `Sign the broker power of attorney in the exact legal name BRAND LLC, matching the Wyoming filing, the EIN (Employer Identification Number) and the bank account. Post the continuous bond (about $${F.num(facts.logistics.broker_bond_annual_usd)} per year, evidence based).`,
    `File the binding ruling request in week 1 with the product construction: fiber cores carry the video data, thin copper conductors carry power and control. Ask for 8544.70.00 (optical fiber cables, fiber rate ${F.pct(df.total)}) against 8544.42.90 (electric conductors fitted with connectors, copper rate ${F.pct(dc.total)}). Attach the factory construction drawing [PLACEHOLDER] and the ATC test report.`,
    `Until the ruling arrives, enter at the copper rate (the model default). Do not enter at the fiber rate on an unsupported position; a later CBP reclassification brings back duty plus interest.`,
    `Per shipment: broker receives the commercial invoice, packing list, bill of lading or air waybill, and the HTS line per SKU. Mark checks the declared value equals the FOB price on the PO (customs value = FOB, evidence based model convention) and that country of origin reads China on every carton.`,
    `Broker files the entry. Duty, MPF and HMF (merchandise processing fee and harbor maintenance fee, about ${F.pct(facts.logistics.mpf_hmf_ocean, 2)} of FOB on ocean entries, management estimate), and the broker entry fee (about $${F.num(facts.logistics.broker_entry_fee_usd)} per entry) are paid from the BRAND LLC account. Reconcile the duty bill to the model duty per unit ($${F.num(launchSkus[1].duty_copper, 2)} on the hero at the copper rate; $${F.num(launchSkus[1].duty_fiber, 2)} at the fiber rate).`,
    `File the entry summary, the duty receipt and the ruling correspondence in the customs folder. Keep every record five years (broker to confirm the retention rule [PLACEHOLDER]).`,
    `When the ruling arrives: if fiber, switch the HTS line on the next entry, ask the broker about refunds on earlier entries, and move the DutyScenario selector in the workbook to 2 (Fiber). If copper, keep the default and reprice the SKU Unit Economics tab if any SKU falls below the Day 90 contribution floor of $${F.num(gv('d90_cpu_pass'))}.`,
    `Quarterly: broker confirms the Section 301 status (List 3 coverage, the forced labor action, the pending second investigation). Mark updates the Landed Cost & Tariff Scenarios tab and reports the change in the monthly KPI pack.`,
  ],
  system: ['Broker portal and ACE entry summaries (ACE is the CBP Automated Commercial Environment), the customs folder, the workbook tab Landed Cost & Tariff Scenarios, the DutyScenario selector on Assumptions.'],
  kpi: [
    'Ruling filed in week 1; decision logged on receipt (30 to 90 days).',
    'Entry accuracy: 0 post entry corrections per quarter (management estimate).',
    `Duty paid per unit versus model: within 2% of $${F.num(launchSkus[1].duty_copper, 2)} on the hero at the copper rate.`,
    'Customs holds: 0 exams caused by paperwork; any exam logged with cause and days lost.',
  ],
  failures: [
    'Legal name mismatch between the bond, the EIN and the bank: entry rejected or payments bounce. Fix: one legal name string copied from the Wyoming certificate everywhere.',
    'Entering at the fiber rate without a ruling: back duty plus interest and a compliance flag. Fix: copper rate until the ruling, documented position.',
    `Stress case: the pending Section 301 action adds 10 points (total ${F.pct(ds.total)}). Fix: the model stress scenario is prepared; the response is a price ladder test (lever 10) and a volume price tier (lever 7), not a margin cut.`,
    'HDMI LA seizure at the border for unlicensed product: lot lost. Fix: HDMI Adopter status and ATC report on file before the first PO.',
    'Undervaluation on the invoice to save duty: a customs fraud exposure for the LLC and for Lucy. Never.',
  ],
  chart: ['03_inbound_supply_chain (customs entry node) and 08_entity_money_goods_flow (duty and filings edges).'],
}));

// ---------- SOP 3: FBA and 3PL inbound ----------
children.push(h1('3. FBA and 3PL inbound'));
children.push(p(`FBA is Fulfillment by Amazon. The 3PL holds bulk stock and feeds FBA in small replenishments, so Amazon storage fees and aged inventory surcharges stay low (lever 9: storage and inbound from ${F.pct(fig.levers[8].from, 1)} to ${F.pct(fig.levers[8].to, 1)} of revenue, management estimate). Launch lots go air DDP straight to FBA; from Month ${oceanFrom} proven SKUs move by ocean through the 3PL (model input, management estimate).`));
children.push(sopTable({
  owner: ['Lucy (3PL contract, freight booking, payment). 3PL (receiving, storage, FBA shipment prep). VA (creates the FBA shipment plans and tracks receipts). Mark (approves the replenishment quantity each week).'],
  trigger: [
    'Customs release of a shipment (section 2) or a 3PL stock arrival notice.',
    `Weekly: FBA available stock on any launch SKU falls below 4 weeks of forward sales (management estimate).`,
    'Amazon inbound capacity or placement fee changes.',
  ],
  steps: [
    'Air DDP launch lots: the forwarder delivers directly to the FBA receive center named in the shipment plan. VA creates the plan in Seller Central before departure so the carton labels print at the factory.',
    `Ocean lots (from Month ${oceanFrom}): drayage from the port to the 3PL. 3PL receives against the packing list, counts cartons, scans a sample of FNSKUs, reports damage within 24 hours, and books the stock as available.`,
    `VA builds the FBA replenishment plan weekly: quantity per SKU = 4 weeks of forward units from the Medium ramp (for example Month 12: ${F.units(sd.medium_month12_units)} units per month on the whole line, split by SKU share of trailing 30 day sales) minus FBA available and inbound.`,
    `3PL picks, labels the cartons with the Amazon shipment labels, and ships by the carrier the plan names. Inbound placement and prep budget: $${F.num(facts.logistics.inbound_per_unit_usd, 2)} per unit (management estimate).`,
    'VA tracks each shipment to Receiving complete, reconciles received units to shipped units, and opens a case for any shortfall above 1% within 30 days of receipt.',
    'Monthly: 3PL stock count reconciled to the inventory record; FBA stock reconciled to the Seller Central inventory ledger; aged inventory (over 180 days; management estimate, confirm against the current Amazon fee schedule) reviewed and either moved out with a coupon or removed before the surcharge.',
    'Record every inbound: shipment ID, SKU, units shipped, units received, days from ship to receive, discrepancies. The 3PL invoice is checked line by line against this record before payment.',
  ],
  system: ['Seller Central (Send to Amazon, inventory ledger, inbound performance), the 3PL warehouse portal, the inbound log, the freight forwarder booking.'],
  kpi: [
    'In stock rate on the launch SKUs: at or above 97% of days (management estimate).',
    'Receiving discrepancy: below 0.5% of units; every case above 1% opened within 30 days.',
    'Ship to FBA receive time: within 10 days for the 3PL feed; within 25 days door to door for air lots (management estimates).',
    `Storage and inbound cost: ${F.pct(fig.levers[8].from, 1)} of revenue at launch, ${F.pct(fig.levers[8].to, 1)} once the 3PL feed runs (model lever 9).`,
  ],
  failures: [
    'Stockout on the hero SKU: rank and reviews stall, and the ad spend buys nothing. Fix: the 4 week FBA cover rule, the section 5 reorder table, and air top ups when the ocean lot is late.',
    'Over sending to FBA: storage and aged inventory surcharges, capacity limits. Fix: 3PL holds bulk, FBA gets 4 weeks.',
    'Carton labels wrong or missing: Amazon relabel fee or refused receipt. Fix: shipment plan before the factory prints labels; 3PL label check.',
    '3PL fails (closes, loses stock, bills wrong): sales stop. Fix: monthly count reconciliation, two 3PL quotes on file, stock insured [PLACEHOLDER: policy].',
    'Receiving shortfalls never claimed: silent margin loss. Fix: reconciliation within 30 days, cases opened by the VA.',
  ],
  chart: ['03_inbound_supply_chain (3PL and FBA inbound nodes).'],
}));

// ---------- SOP 4: Listing and PPC ----------
children.push(h1('4. Listing creation and PPC management'));
children.push(p(`PPC is pay per click advertising. ASIN is the Amazon Standard Identification Number. ACOS is advertising cost of sales (ad spend divided by ad attributed sales). TACOS is total advertising cost of sales (ad spend divided by total sales); the gates use TACOS. The model ad spend by phase (Medium, evidence based from the August model): ${F.pct(facts.august.m18_assumptions.ads_phases[0][2], 0)} of revenue Months 1 to 3, ${F.pct(facts.august.m18_assumptions.ads_phases[1][2], 0)} Months 4 to 6, ${F.pct(facts.august.m18_assumptions.ads_phases[2][2], 0)} Months 7 to 12, ${F.pct(facts.august.m18_assumptions.ads_phases[3][2], 0)} Months 13 to 18, never below 8% in any year.`));
children.push(sopTable({
  owner: ['Mark (listing structure, copy approval, campaign structure, bid rules, weekly optimization). VA (daily checks, search term reports, negative keywords, coupon setup). Lucy (approves claims that need factory evidence: CL3, FCC, HDMI category name).'],
  trigger: [
    'Brand Registry approved on the pending trademark application (weeks 3 to 4), then listings built in weeks 6 to 9 before the first inventory receipt.',
    'Weekly: the PPC optimization slot in the monthly cadence (05_monthly_operating_cadence).',
    `Any listing suppression, policy warning or IP (intellectual property) complaint: same day. An item open more than ${gv('d90_ip_days')} days is a Day 90 kill trigger.`,
  ],
  steps: [
    `Build one parent ASIN per line with the lengths as child variations (lever 6): the 8K line (${launchSkus.filter((s) => s.sku.startsWith('8K')).map(skuName).join(', ')} at launch) and the 4K line (${launchSkus.filter((s) => s.sku.startsWith('4K')).map(skuName).join(', ')} at launch). Titles, bullets, A+ outline, backend keywords and the image brief are in 09_Amazon_Listing_Kit.docx.`,
    'Apply the naming rule to every field: the category name Ultra High Speed HDMI Cable (8K line) or Premium High Speed HDMI Cable (4K line, pending certification, placeholder), never HDMI 2.1 or HDMI 2.0. Titles may say 48Gbps and 8K. Say fiber optic HDMI or hybrid active optical; never pure fiber or all glass.',
    'Claims gate: in wall or CL3 only with the CL3 document on file; FCC statement only with the factory FCC report; the 5 Year Signal Guarantee wording exactly as in section 6. Add the Prop 65 statement. Lucy signs off the claims list.',
    `Set the price ladder: list price and sale price per SKU from the SKU table (hero list $${F.num(hero.list, 2)}, sale $${F.num(hero.price, 2)}). Coupons and the plus or minus 5% ASP (average selling price) test (lever 10) run one at a time, two weeks each, logged with the result.`,
    'Launch campaigns in this order: automatic campaign per parent for discovery, then exact match manual campaigns on the winning search terms, then a branded defense campaign. Daily budget per the phase rate on forecast revenue.',
    'Weekly: VA pulls the search term report, adds negatives for terms with 15 clicks and no order (lever 4, keyword pruning, management estimate rule), moves converting terms to exact match, and reports ACOS, TACOS, conversion and spend per SKU.',
    'Review velocity (lever 5): enroll in Vine where eligible, use the Brand Registry follow up email, and the insert card with the QR (quick response) code setup video. Never incentivize or gate reviews; never contact buyers outside Amazon messaging for reviews.',
    'Account health: VA checks Account Health, Voice of the Customer and the Performance Notifications page every working day and logs the status. Any suppression or warning goes to Mark the same day with a plan to clear it within 14 days.',
    'Month end: Mark writes the listing and ads section of the KPI pack: conversion by SKU, TACOS versus phase rate, review count and rating (reported, never estimated), suppressions, and the next month budget.',
  ],
  system: ['Seller Central (Listings, Brand Registry, A+, Advertising console, Account Health, Voice of the Customer), the keyword and negative list sheet, 09_Amazon_Listing_Kit.docx, the claims evidence folder.'],
  kpi: [
    `Unit session conversion: Day 90 pass at or above ${F.pct(gv('d90_conv_pass'), 0)} on at least 3 of the 5 launch SKUs; hero below ${F.pct(gv('d90_conv_kill'), 0)} is a kill trigger (observed at the gate).`,
    `TACOS: at or below ${F.pct(gv('d90_tacos_pass'), 0)} in Month 3 and falling month over month (kill above ${F.pct(gv('d90_tacos_kill'), 0)}); at or below ${F.pct(gv('m12_tacos_pass'), 0)} in Month 12 (kill above ${F.pct(gv('m12_tacos_kill'), 0)}).`,
    `Suppressed listings and IP complaints open more than ${gv('d90_ip_days')} days: 0.`,
    'Review velocity: reported monthly; no target is stated because no review count exists yet [PLACEHOLDER: first 90 day actuals].',
  ],
  failures: [
    'HDMI version number in a title or image: listing removal and licensing exposure. Fix: naming rule in the claims gate, VA checks every new field.',
    `Ad spend runs above the phase rate without conversion: TACOS above ${F.pct(gv('d90_tacos_kill'), 0)} in Month 3 is a kill trigger. Fix: weekly negatives, budget caps per campaign, exact match migration.`,
    'Listing suppression that stays open: a lost quarter is the Low scenario (Months 19 to 21 at 10% of trend). Fix: daily health check, documents ready to upload, 14 day clearance rule.',
    'Review policy breach: account suspension. Fix: Vine and insert card only, no incentives, no outside contact.',
    'Unsupported in wall or CL3 claim: liability and a false advertising exposure. Fix: claims gate with documents on file.',
  ],
  chart: ['02_order_to_cash_amazon (session to settlement). The gate numbers are drawn on 06_day90_scale_or_kill.'],
}));

// ---------- SOP 5: Inventory planning and reorder points ----------
children.push(h1('5. Inventory planning and reorder points'));
children.push(p(`Working capital is what kills cable businesses, not margin. The model holds ${coverM} months of forward units on hand (Medium; Low ${coverL}, High ${coverH}; management estimates) plus the lead time demand in transit, funded by the business. Lead time ${airLead} month by air (Months 1 to ${oceanFrom - 1}) and ${oceanLead} months by ocean from Month ${oceanFrom} on proven SKUs. Deposit ${F.pct(deposit, 0)} at order, ${F.pct(1 - deposit, 0)} at shipment. Amazon pays out with a ${payoutLag} day lag, so ${payoutLag}/30 of a month's net proceeds is still a receivable at month end (evidence based).`));
const RM = sd.reorder.Medium;
const fobUnit = RM.fob_per_unit_usd;
children.push(p([B('Stock rule (the workbook rule, Cash Flow and Working Capital tab). '), run(`At each month end, target closing stock = cover months (${coverM} in Medium) times the average of the next three months of units sold on all channels. Arrivals needed in a month = target closing stock minus (opening stock minus that month's sales), never below zero. The PO for an arrival is placed lead time months earlier: ${airLead} month by air for arrivals through Month ${oceanFrom - 1}, ${oceanLead} months by ocean for arrivals from Month ${oceanFrom}, so Month ${oceanFrom - oceanLead} places two POs (the last air lot and the first ocean lot) and Month 0, before launch, places the launch PO. Deposit ${F.pct(deposit, 0)} of FOB at order, balance ${F.pct(1 - deposit, 0)} at shipment (the month before arrival); freight, duty, broker, inbound and QC are paid on arrival. The table below is the model's own numbers (figures.json scenarios.Medium.monthly: units, closing_stock, arrivals, inventory_on_hand, inventory_in_transit), so the cash forecast and the PO plan agree.`)]));
const poText = (r) => {
  if (!r.po_for_months.length) return 'none';
  return r.po_for_months.map((a) => `${F.units(Math.round(M.monthly[a - 1].arrivals))} for Month ${a} by ${a >= oceanFrom ? 'ocean' : 'air'}`).join('; ');
};
children.push(gridTable(
  ['Month', 'Units sold', `Target closing stock (${coverM} x next 3 months average)`, 'Arrivals landing this month (units)', 'PO to place this month (units, arrival month, mode)', `Deposit due (${F.pct(deposit, 0)} of FOB)`, 'On hand at month end ($, landed cost)', 'In transit plus deposits ($)'],
  [1200, 1000, 1300, 1100, 1850, 1150, 1300, 1180],
  RM.rows.map((r) => [
    `Month ${r.month}`,
    r.month === 0 ? 'n/a' : F.units(r.units),
    r.month === 0 ? 'n/a' : F.units(Math.round(r.target_closing_stock)),
    r.month === 0 ? 'n/a' : F.units(Math.round(r.arrivals)),
    poText(r),
    r.po_units > 0 ? F.money(r.deposit_usd) : '$0',
    r.month === 0 ? 'n/a' : F.money(r.inventory_on_hand_usd),
    r.month === 0 ? 'n/a' : F.money(r.inventory_in_transit_usd),
  ]),
));
children.push(small(`Month 0 is the pre launch month: its PO is the launch lot. Units sold are all channels (Amazon US only in Year 1). Source: figures.json scenarios.Medium.monthly (units, closing_stock, arrivals, inventory_on_hand, inventory_in_transit); cover, lead times, ocean start and deposit from the Assumptions tab (Active column). Deposit base: the model keeps one inventory pool at the hero cost, so FOB is $${F.num(fobUnit, 2)} per unit (8K 10M at FX ${F.num(fig.meta.fx)} KRW per USD) for every unit and on hand stock is valued at the P&L landed cost ($${F.num(RM.landed_cost_per_unit_month1, 2)} per unit in Year 1); in transit includes the deposits on open orders. Low: cover ${coverL} months, air only; High: cover ${coverH} months. Split the line total by SKU share of trailing 30 day sales; at launch use the brief's ${facts.logistics.first_lots} (evidence based).`, { before: 80 }));
children.push(p(`Cash check on the hero SKU: a 1,000 unit PO at the August landed cost of $${F.num(hero.landed_cost, 2)} per unit is about ${F.money(1000 * hero.landed_cost)} of inventory, with ${F.money(1000 * hero.landed_cost * deposit)} due at order. The model's peak cash need in the Medium case is ${F.money_k(M.cash.peak_need)} in ${M.cash.peak_need_month_label} against opening capital of ${F.money(M.cash.opening_capital)}, and the funding line of ${F.money_k(M.cash.funding_gap)} must be in place by ${F.month_label(M.cash.funding_needed_by_month)}.`));
children.push(sopTable({
  owner: ['Lucy (decides and funds the PO). Mark (prepares the reorder table, the SKU split and the cash check). VA (updates sales and stock actuals weekly).'],
  trigger: [
    'Monthly, week 2 of the cadence: inventory and cash review (05_monthly_operating_cadence).',
    'Weekly: FBA available on any launch SKU below 4 weeks of forward sales (section 3).',
    `At the Day 90 gate: switch proven SKUs to ocean (Month ${oceanFrom} in the model); delist losers.`,
  ],
  steps: [
    'VA updates the inventory sheet: units sold last 30 days per SKU, FBA available, FBA inbound, 3PL on hand, in transit (by PO), next arrival dates.',
    'Mark refreshes the forecast: Medium ramp as the base, adjusted by the trailing 30 day run rate per SKU (if actuals are more than 20% off the ramp for two months, switch the active scenario in the workbook and tell Lucy).',
    `Compute the target closing stock for the arrival month (${coverM} times the average of the next three months of units) and the PO quantity: target minus the stock expected at the start of that month (actual on hand plus in transit, less the units forecast until then). Split by SKU, round to cartons. If actuals match the model, the quantity equals the arrivals column of the table.`,
    `Cash check: PO deposit (${F.pct(deposit, 0)}) and balance (${F.pct(1 - deposit, 0)}) placed on the cash forecast by month; confirm the balance after the PO stays above zero for the next three months after funding, and that the runway stays at or above ${gv('d90_run_pass')} months (Day 90 pass level; kill below ${gv('d90_run_kill')}).`,
    `Freight mode: air for launch lots and for any top up that would otherwise stock out; ocean (lead ${oceanLead} months) for proven SKUs after the Day 90 gate. Air costs $${F.num(facts.logistics.air_per_kg_usd, 2)} per kg packaged, ocean $${F.num(facts.logistics.ocean_per_kg_usd, 2)} per kg (management estimates): on the hero that is $${F.num(launchSkus[1].air_freight, 2)} versus $${F.num(launchSkus[1].ocean_freight, 2)} per unit.`,
    'Lucy approves or trims the PO. Trims are recorded with the reason (cash, demand, factory capacity).',
    'Hand the approved quantity to section 1 (purchasing). Record the decision in the inventory log and the KPI pack.',
    'Slow movers: any SKU below 50% of its forecast for two months is flagged; three months triggers a coupon sell down and no reorder until the gate review.',
  ],
  system: ['Inventory sheet (per SKU actuals), the workbook tab Cash Flow and Working Capital (monthly 1 to 24), the PO log, Seller Central inventory reports, 3PL portal.'],
  kpi: [
    `Months of cover on hand plus in transit: ${coverM} (Medium) with a floor of 6 weeks on the hero.`,
    'Stockout days on launch SKUs: 0 per month (management estimate).',
    'Inventory on hand within 15% of the model month (Medium Month 12 on hand ' + F.money_k(M.monthly[11].inventory_on_hand) + ', in transit ' + F.money_k(M.monthly[11].inventory_in_transit) + ').',
    `Cash runway: at or above ${gv('d90_run_pass')} months at every review.`,
  ],
  failures: [
    'Over ordering on an optimistic forecast: cash locked in stock, aged inventory, the funding line exhausted. Fix: the target stock rule, trailing 30 day adjustment, Lucy trims.',
    'Under ordering before the ocean switch: a stockout exactly when the ad spend has built rank. Fix: lead time plus cover in the rule, air top up option.',
    `Switching to ocean too early on an unproven SKU: ${oceanLead} months of stock on a loser. Fix: ocean only after the Day 90 gate per SKU.`,
    'Forecast not updated: the PO plan and the cash forecast drift apart. Fix: VA weekly actuals, Mark monthly refresh.',
    'Low case behavior: the model keeps buying through the lost quarter; an operator must pause POs while a listing is suppressed (model note low_peak_need_ordering).',
  ],
  chart: ['05_monthly_operating_cadence (inventory and cash week) and 03_inbound_supply_chain.'],
}));

// ---------- SOP 6: Returns and warranty ----------
children.push(h1('6. Returns and warranty (5 Year Signal Guarantee)'));
children.push(p(`The guarantee is the differentiator: if a BRAND cable loses signal within 5 years, the replacement ships before the return arrives. Direction proof design (color coded SOURCE and DISPLAY ends, printed arrows, QR setup video) is lever 3, which the model credits with returns from ${F.pct(fig.levers[2].from, 1)} to ${F.pct(fig.levers[2].to, 1)} (Medium, Year 2 onward). Returns and storage reserve on the hero: $${F.num(hero.returns_storage, 2)} per unit (August model).`));
children.push(sopTable({
  owner: ['Mark (policy, triage rules, factory claims with Lucy). VA (daily customer messages, replacement orders, return log). Lucy (factory credit negotiation for defects).'],
  trigger: [
    'A buyer message, a return request, a negative review or seller feedback mentioning signal loss, no picture, flicker or wrong direction.',
    'Weekly: return rate and reason code review. Monthly: root cause summary to the factory.',
  ],
  steps: [
    'Within 24 hours the VA answers with the triage script: confirm the SOURCE end is at the player or console and the DISPLAY end at the screen; confirm the source output setting (4K 120Hz or 8K 60Hz) and that the device port supports it; send the QR setup video link. [PLACEHOLDER: share of signal complaints that are direction or settings issues, from the first 90 day actuals].',
    'If the triage does not fix it: register the guarantee claim (order ID, SKU, purchase date, fault description, photos) in the return log.',
    'Ship the replacement the same working day from FBA (multi channel fulfillment order) or from the 3PL. Do not wait for the return. Send the buyer the prepaid return label for the faulty unit.',
    'Within the Amazon return window the buyer may also use the Amazon return path; the VA records the Amazon return reason code against the same claim so the rate is not double counted.',
    'Returned units go to the 3PL (not back to FBA) for a bench test: signal lock at full length, head inspection, connector damage. Result logged as: no fault found, user damage, manufacturing defect, shipping damage.',
    `Monthly: Mark sends the factory the defect summary and claims a credit for manufacturing defects against the next PO. Lucy agrees the credit in the PO (section 1). Target: factory credit covers 100% of confirmed manufacturing defects [PLACEHOLDER: warranty terms in the factory agreement].`,
    `Monthly KPI: return rate by SKU and by reason code against the gate levels (Day 90 pass at or below ${F.pct(gv('d90_ret_pass'), 0)}, kill above ${F.pct(gv('d90_ret_kill'), 0)}; Month 12 pass at or below ${F.pct(gv('m12_ret_pass'), 0)}, kill above ${F.pct(gv('m12_ret_kill'), 0)}). Product changes (lever 3) go to the factory with the data.`,
    'Guarantee registrations (the trade up registration and the Spec Match finder on the off Amazon site, lever 11) are stored in the email list with consent, never used for review requests.',
  ],
  system: ['Seller Central (Buyer messages, Manage Returns, Voice of the Customer, multi channel fulfillment), the return log, the 3PL bench test form, the factory claim file, the email list tool [PLACEHOLDER: tool choice].'],
  kpi: [
    `Return rate: ${F.pct(facts.august.m18_assumptions.returns, 0)} Year 1 in the Medium model, ${F.pct(fig.levers[2].to, 1)} from Year 2; never below 2% in any plan (brief realism rule).`,
    'First response time: under 24 hours on 95% of messages (management estimate).',
    'Replacement shipped the same working day on 100% of accepted claims.',
    'No fault found share of returns: reported monthly; a fall shows the direction proof design and the setup video are working.',
  ],
  failures: [
    `Return rate above ${F.pct(gv('d90_ret_kill'), 0)} at Day 90: a kill trigger. Fix: triage script, direction proof packaging, bench test data to the factory within the first 60 days.`,
    'Replacement held until the return arrives: the guarantee promise is broken and reviews say so. Fix: same day replacement rule, stock reserved for claims (1% of units, management estimate).',
    'Returned units sent back into FBA: a known faulty unit resold. Fix: returns to the 3PL only, bench test before any restock.',
    'Factory refuses credits: margin erodes by the defect rate. Fix: warranty clause in the PO terms before the first order.',
    'Review requests sent to guarantee registrants: an Amazon policy breach. Fix: consent and purpose limits in the email tool.',
  ],
  chart: ['04_returns_and_signal_guarantee.'],
}));

// ---------- SOP 7: Monthly close and KPI review ----------
children.push(h1('7. Monthly close and KPI review'));
children.push(p(`Close the books by working day 10, review the KPI pack against the model, and decide on cash and orders. The Day 90 gate (end of Month 3) and the Month 12 gate are run inside this process with the thresholds below. Gate result rule: KILL if any kill trigger fires; PASS if every modeled pass threshold is met; REVIEW otherwise, which means fix and retest within 30 days. Model results: Day 90 Medium ${fig.kill_criteria.results.Medium.day90}, Low ${fig.kill_criteria.results.Low.day90}, High ${fig.kill_criteria.results.High.day90}; Month 12 Medium ${fig.kill_criteria.results.Medium.month12}, Low ${fig.kill_criteria.results.Low.month12}, High ${fig.kill_criteria.results.High.month12}.`));
children.push(h3('Day 90 gate (end of Month 3)'));
children.push(gridTable(['Metric', 'Pass', 'Kill trigger', 'Source', 'Medium model value'], [2900, 1850, 1700, 1250, 2380],
  [
    ['Unit session conversion on at least 3 of the 5 launch SKUs', `at or above ${F.pct(gv('d90_conv_pass'), 0)}`, `hero below ${F.pct(gv('d90_conv_kill'), 0)}`, 'observed', 'observed at the gate'],
    ['TACOS in Month 3, falling month over month', `at or below ${F.pct(gv('d90_tacos_pass'), 0)}`, `above ${F.pct(gv('d90_tacos_kill'), 0)}`, 'model', F.pct(M.gates.day90.metrics.tacos.value, 0)],
    ['Return rate', `at or below ${F.pct(gv('d90_ret_pass'), 0)}`, `above ${F.pct(gv('d90_ret_kill'), 0)}`, 'model', F.pct(M.gates.day90.metrics.ret.value, 0)],
    ['Contribution per unit before advertising on at least 3 of 5 SKUs', `at or above $${F.num(gv('d90_cpu_pass'))}`, `hero below $${F.num(gv('d90_cpu_kill'))}`, 'model', '$' + F.num(M.gates.day90.metrics.cpu.value, 2) + ' (hero)'],
    ['Cash runway after funding the next purchase order', `at or above ${gv('d90_run_pass')} months`, `below ${gv('d90_run_kill')} months`, 'model', F.num(M.gates.day90.metrics.runway.value, 1) + ' months'],
    ['Suppressed listing or IP complaint', 'none open', `any open more than ${gv('d90_ip_days')} days`, 'observed', 'observed at the gate'],
  ], { cantSplit: true }));
children.push(spacer());
children.push(h3('Month 12 gate'));
children.push(gridTable(['Metric', 'Pass', 'Kill trigger', 'Source', 'Medium model value'], [2900, 1850, 1700, 1250, 2380],
  [
    ['Monthly revenue in Month 12', `at or above ${F.money(gv('m12_rev_pass'))}`, `below ${F.money(gv('m12_rev_kill'))}`, 'model', F.money(M.gates.month12.metrics.rev.value) + ' (Low ' + F.money(L.gates.month12.metrics.rev.value) + ')'],
    ['Cumulative net', `at or above ${F.money(gv('m12_cum_pass'))}`, `below ${F.money(gv('m12_cum_kill'))}`, 'model', F.money(M.gates.month12.metrics.cum.value)],
    ['TACOS', `at or below ${F.pct(gv('m12_tacos_pass'), 0)}`, `above ${F.pct(gv('m12_tacos_kill'), 0)}`, 'model', F.pct(M.gates.month12.metrics.tacos.value, 0)],
    ['Return rate', `at or below ${F.pct(gv('m12_ret_pass'), 0)}`, `above ${F.pct(gv('m12_ret_kill'), 0)}`, 'model', F.pct(M.gates.month12.metrics.ret.value, 0)],
    ['Operating cash flow', 'positive two consecutive months', `three consecutive negative months with a funding gap above ${F.money(gv('m12_gap_kill'))}`, 'model', M.gates.month12.metrics.opcf ? (M.gates.month12.metrics.opcf.passes ? 'passes' : 'does not pass') : 'see workbook'],
  ], { cantSplit: true }));
children.push(small(`Kill means: ${fig.kill_criteria.kill_means}. Plan B (figures.json plan_b, management estimate): B2B (business to business) installer distribution plus Korea wholesale, revenue ${F.money_k(fig.plan_b.Y1.revenue)} in Year 1 with net ${F.money_k(fig.plan_b.Y1.net)}, ${F.money_k(fig.plan_b.Y3.revenue)} in Year 3 with net ${F.money_k(fig.plan_b.Y3.net)}.`, { before: 80 }));
children.push(sopTable({
  owner: ['Lucy (owner of the close, decides on cash and the gates). Mark (prepares the KPI pack and the gate scorecard). VA (pulls the Amazon reports). 세무사 (Korean books, the overseas direct investment follow up reports, the Korean filings). US CPA (certified public accountant) [PLACEHOLDER: firm] (Form 5472 preparation, California nexus question).'],
  trigger: [
    'Calendar month end; pack due working day 10; review call in week 2 of the cadence.',
    'The Day 90 gate at the end of Month 3 and the Month 12 gate at the end of Month 12, run as part of that month\'s review.',
  ],
  steps: [
    `VA downloads from Seller Central: settlement reports for the month, the Amazon payouts received, the inventory ledger, the advertising invoices and the campaign report, returns and refunds, Voice of the Customer, Account Health. Note the ${payoutLag} day payout lag: the last settlement of the month lands in the next month, so the receivable is ${payoutLag}/30 of the month's net proceeds.`,
    'Mark books revenue, referral fees, FBA fees, ad spend, refunds, storage, 3PL invoices, freight, duty and broker fees, insurance, software, the seller fee and the consulting fee into the bookkeeping file by model line. Inventory on hand and in transit are valued at landed cost.',
    `Reconcile: bank statement to the books; Amazon payouts to settlements; FBA units to the inventory ledger; duty paid to entries. Unexplained differences above $100 are investigated before the pack is issued (management estimate).`,
    'Produce the KPI pack (one page plus the tables): units, revenue, net and cumulative net versus the model month; TACOS; contribution per unit before and after advertising; return rate; conversion by SKU; inventory on hand and in transit; cash balance versus the model, where Medium Month 6 model cash before funding is ' + F.money_k(M.monthly[5].cash_balance) + '; runway; open compliance items.',
    'Scenario check: if two consecutive months land within the Low band, set the workbook selector to Low and rerun the cash and funding plan; present the funding gap and date to Lucy with the line of credit or investor option.',
    `Gate months: Mark fills the gate scorecard with observed conversion and listing status plus the modeled metrics; the result (PASS, REVIEW, KILL) is written into the pack. KILL stops further capital: sell down inventory, delist losers, move to Plan B. REVIEW: fix and retest within 30 days.`,
    'Lucy and Mark review (one hour): decisions on the next PO (section 5), ad budget (section 4), price tests, compliance actions (section 8). Decisions and owners logged.',
    'Send the 세무사 the month\'s bank statements, the Amazon payout summary, the consulting fee invoice and the owner draw record, in Korean summary form. Send the US CPA the quarter end books for the Form 5472 reportable transactions list (capital contributions, owner draws, any payments between Lucy, H2 Company and BRAND LLC).',
  ],
  system: ['Bookkeeping file [PLACEHOLDER: tool, for example a cloud ledger connected to the Mercury or Relay account], Seller Central reports, the workbook (Monthly PL Y1 Y2, Cash Flow and Working Capital, Funding and Breakeven), the KPI pack template, the decision log.'],
  kpi: [
    'Close completed by working day 10: 12 of 12 months.',
    'Bank to books reconciliation difference: $0 after investigation.',
    `Net versus model: within 10% of the Medium month (Medium 18 month net ${F.money_k(M.m18.net)}, ${F.pct(M.m18.net_margin)}; cumulative breakeven ${F.month_label(M.m18.breakeven_month)}).`,
    'Gate scorecards delivered in Month 3 and Month 12 with a written result.',
  ],
  failures: [
    'Payout timing booked as revenue: revenue overstated and cash understated. Fix: settlement based revenue, payout lag receivable.',
    'Owner draws, consulting fees and capital mixed with trade payments: the Form 5472 reportable transaction list is wrong and the penalty is $' + F.num(facts.entity.form_5472_penalty_usd) + '. Fix: separate ledger accounts and a monthly list to the CPA.',
    'Low band ignored for months: the funding gap arrives unannounced. Fix: two month rule for switching the scenario.',
    'Gate skipped because the month was busy: capital keeps flowing into a losing plan. Fix: the gate is a line in the Month 3 and Month 12 close checklist with Lucy\'s signature.',
    'Mark negotiating or signing for the LLC in the US: dependent agent risk to the treaty position. Fix: Lucy signs; Mark advises under the consulting agreement.',
  ],
  chart: ['05_monthly_operating_cadence and 06_day90_scale_or_kill.'],
}));

// ---------- SOP 8: Compliance calendar ----------
children.push(h1('8. Compliance calendar'));
children.push(p(`Formation month F is the month the Wyoming articles are filed. Model Month 1 (first inventory live on Amazon) is assumed to be about F+2, because the build sequence needs 8 to 10 weeks before the first sale (DECISIONS.md; management estimate). Fixed calendar dates (for example the Form 5472 due date) are shown as calendar dates. ODI is the overseas direct investment report (해외직접투자 신고) under the Foreign Exchange Transactions Act (FTA, 외국환거래법). USPTO is the United States Patent and Trademark Office.`));
children.push(sopTable({
  owner: ['Lucy (every filing is in her name or the LLC\'s; she signs). Attorney (trademark). Customs broker (bond and entries). 세무사 (Korean filings). US CPA (Form 5472). Mark (keeps the calendar, chases evidence, files the evidence folder). VA (Amazon account health daily).'],
  trigger: ['The calendar below, run at the week 4 compliance slot of the monthly cadence; any notice from the IRS (Internal Revenue Service), the Wyoming Secretary of State, the registered agent, the foreign exchange bank, USPTO, CBP or Amazon triggers an entry the same day.'],
  steps: [
    'Mark reviews the calendar in week 4 of every month: items due in the next 60 days are listed with owner, cost and the evidence required.',
    'Lucy approves and pays; the owner files; the evidence (filing receipt, confirmation number, policy document, bank acknowledgment) goes into the compliance folder and the row is marked Done in 06_Compliance_Checklist.xlsx.',
    'Any missed date is escalated to Lucy within 24 hours with the penalty exposure and the cure (for example a late Form 5472 carries a $' + F.num(facts.entity.form_5472_penalty_usd) + ' penalty).',
    'Once a year (formation anniversary month) Mark and the 세무사 reconcile the calendar against the actual filings and update the next 12 months.',
  ],
  system: ['06_Compliance_Checklist.xlsx (status dropdown), the compliance folder, the registered agent portal, the foreign exchange bank, the USPTO docket from the attorney, Seller Central Account Health.'],
  kpi: ['Filings on time: 100%. Penalties paid: $0. Account Health: healthy every day; any warning cleared within 14 days.'],
  failures: [
    `Form 5472 with the pro forma Form 1120 missed: $${F.num(facts.entity.form_5472_penalty_usd)} penalty per year. Fix: CPA engaged in F+1, reportable transaction list every quarter (section 7).`,
    'Capital sent to the US before the ODI report is accepted: an FTA breach; enforcement against individual online sellers is rising (brief). Fix: the bank report is a hard stop before any transfer.',
    'Wyoming annual report missed: the LLC loses good standing and the bank or Amazon can freeze the account. Fix: registered agent reminder plus the calendar.',
    `Product liability insurance not bound in the month sales pass $${F.num(facts.amazon.insurance_trigger_monthly_sales_usd)} (Medium Month ${insTriggerMonth}, which passes the threshold by ${F.money(insMargin)}, so bind before Month ${insTriggerMonth} closes): Amazon can suspend the listing. Fix: quote in hand before launch, bind on the trigger.`,
    'Trademark office action missed: the application abandons and Brand Registry lapses with it. Fix: attorney docket mirrored in the calendar.',
  ],
  chart: ['08_entity_money_goods_flow (filings edges).'],
}));
children.push(spacer());
children.push(h3('Calendar by month, relative to formation month F'));
const CW = [1100, 3250, 1200, 2000, 2530]; // Item column wide enough that official form numbers (Form SS-4, W-8BEN-E, 599-64-00496) do not wrap mid number
const cal = [
  ['F+0', 'File the Wyoming single member LLC articles; appoint the registered agent. Exact legal name BRAND LLC recorded once and copied everywhere.', 'Lucy, attorney or formation service', `$${F.num(facts.entity.formation_cost_usd)} state fee (evidence based)`, 'Certificate of organization, registered agent agreement'],
  ['F+0', 'Form SS-4: EIN by phone or fax, no SSN (Social Security number) needed.', 'Lucy', '$0', 'IRS EIN confirmation letter [PLACEHOLDER: EIN; CPA to confirm the notice name, CP 575 expected]'],
  ['F+0', 'ODI report to the foreign exchange bank before any capital moves (any stake of 10% or more, any amount). Flag the incorporate H2 first question to the 세무사 in the same meeting: registration 599-64-00496 (H2 Company) is a sole proprietorship.', 'Lucy, 세무사', '$0 filing; 세무사 fee [PLACEHOLDER]', 'Bank acceptance of the ODI report'],
  ['F+0', `Trademark application through a US licensed attorney (foreign applicants must). About $${F.num(facts.trademark.uspto_fee_per_class_usd)} per class plus attorney fees ${facts.trademark.attorney_fees}.`, 'Lucy, attorney', `$${F.num(facts.trademark.uspto_fee_per_class_usd)} per class plus fees`, 'USPTO serial number and filing receipt [PLACEHOLDER]'],
  ['F+0', `Customs broker engaged, power of attorney signed, continuous bond posted (about $${F.num(facts.logistics.broker_bond_annual_usd)} per year). CBP binding ruling request filed in week 1 (${facts.logistics.binding_ruling_timing}).`, 'Lucy, broker, Mark', `$${F.num(facts.logistics.broker_bond_annual_usd)} bond; ruling fee per broker quote [PLACEHOLDER]`, 'Bond number, ruling request receipt'],
  ['F+1', 'USD bank account for the LLC (Mercury or Relay; Payoneer or WorldFirst as fallback). Capital transfer only after the ODI acceptance; then the Korean remittance follow up report.', 'Lucy, 세무사', '$0; bank fees [PLACEHOLDER]', 'Account opening letter, transfer confirmation, remittance report receipt'],
  ['F+1', 'Securities acquisition follow up report after the capital is paid into the LLC (timing per the foreign exchange bank) [PLACEHOLDER: bank deadline].', 'Lucy, 세무사', '$0', 'Bank acknowledgment'],
  ['F+1', 'Form W-8BEN-E tax interview as a foreign owned entity; Amazon Professional seller account ($' + F.num(facts.amazon.professional_fee_monthly_usd, 2) + ' per month); identity verification; bank linked. Brand Registry on the pending trademark application.', 'Lucy, Mark', `$${F.num(facts.amazon.professional_fee_monthly_usd, 2)} per month`, 'Account approval email, Brand Registry approval'],
  ['F+1', 'US CPA engaged for Form 5472 with the pro forma Form 1120, and the California income tax nexus question from FBA inventory put on the CPA list.', 'Lucy, CPA', `$${F.num(wc.form_5472_prep.Medium)} per year preparation (management estimate)`, 'Engagement letter'],
  ['F+2', 'Product compliance file before the first PO: HDMI Adopter certificate and ATC test report from the factory, FCC test report for the active heads, CL3 documentation (before any in wall claim), Prop 65 label on the packaging, HDMI naming rule on all artwork.', 'Lucy, factory, Mark', 'Factory document fees [PLACEHOLDER]', 'Certificates and reports with numbers [PLACEHOLDER]'],
  ['F+2', 'Written consulting or commission agreement between BRAND LLC and Mark (no equity, no authority to negotiate or sign). Keeps the treaty position: no US office, no US employees, no dependent agent.', 'Lucy, Mark, attorney', 'Attorney fee [PLACEHOLDER]', 'Signed agreement'],
  ['F+3 to F+5', `CBP binding ruling decision expected 30 to 90 days after filing. On receipt: HTS line updated, DutyScenario selector set, refunds requested if fiber.`, 'Broker, Mark', '$0', 'Ruling letter and number [PLACEHOLDER]'],
  ['About F+' + (insTriggerMonth + 2), `Product liability insurance, $${F.money_m(facts.amazon.product_liability_insurance_usd).replace('$', '')} cover, bound in the first month sales pass $${F.num(facts.amazon.insurance_trigger_monthly_sales_usd)} (Medium model Month ${insTriggerMonth}, Amazon revenue ${F.money(M.monthly[insTriggerMonth - 1].amazon_revenue)}; Low Month ${L.insurance_trigger_month}, High Month ${H.insurance_trigger_month}). Month ${insTriggerMonth} passes the threshold by ${F.money(insMargin)}, so bind before Month ${insTriggerMonth} closes. Quote obtained before launch.`, 'Lucy', `$${F.num(wc.insurance_per_year.Medium)} per year (management estimate)`, 'Policy document, certificate of insurance uploaded to Seller Central'],
  ['About F+5', 'Day 90 gate review (end of model Month 3): scorecard signed by Lucy (section 7).', 'Lucy, Mark', '$0', 'Signed scorecard'],
  ['F+12', `Wyoming annual report and fee, due in the formation anniversary month (registered agent to confirm the exact date [PLACEHOLDER]).`, 'Lucy, registered agent', `$${F.num(facts.entity.annual_cost_usd)} (evidence based)`, 'Filing receipt'],
  ['F+12', 'Customs continuous bond renewal.', 'Broker', `$${F.num(facts.logistics.broker_bond_annual_usd)}`, 'Renewed bond notice'],
  ['F+12', 'Korean annual business performance report on the overseas direct investment (date set by the foreign exchange bank) [PLACEHOLDER: deadline].', 'Lucy, 세무사', '세무사 fee [PLACEHOLDER]', 'Bank acknowledgment'],
  ['About F+14', 'Month 12 gate review (end of model Month 12): scorecard signed by Lucy (section 7).', 'Lucy, Mark', '$0', 'Signed scorecard'],
  ['About F+' + (insTriggerMonth + 14), 'Product liability insurance renewal, 12 months after binding.', 'Lucy', `$${F.num(wc.insurance_per_year.Medium)} per year`, 'Renewed policy'],
  ['April 15 each year', `Form 5472 with the pro forma Form 1120 for the prior calendar year (foreign owned single member LLC). Penalty for missing it: $${F.num(facts.entity.form_5472_penalty_usd)}. Extension available through the CPA [PLACEHOLDER: confirm].`, 'CPA, Lucy', `$${F.num(wc.form_5472_prep.Medium)} (management estimate)`, 'Filed copy and proof of transmission'],
  ['Each year, date per 세무사', `Korean overseas financial account report if the month end aggregate of foreign accounts passes KRW ${F.num(facts.korea.overseas_financial_account_report_threshold_krw / 1000000)}M in any month of the prior year.`, 'Lucy, 세무사', '세무사 fee [PLACEHOLDER]', 'Filing receipt'],
  ['As issued', 'USPTO office actions and, after registration, the maintenance filings [PLACEHOLDER: attorney to confirm the maintenance window; years 5 to 6 expected] (attorney docket; placeholder $0 in the model until Year 5).', 'Attorney, Lucy', 'Attorney quote [PLACEHOLDER]', 'USPTO notices and responses'],
  ['Every working day', 'Amazon Account Health, Performance Notifications and Voice of the Customer check; any warning logged and escalated to Mark.', 'VA', '$0', 'Daily log line'],
  ['Every month, week 4', 'Compliance slot: calendar review, evidence filed, 06_Compliance_Checklist.xlsx status updated; Korean bookkeeping summary to the 세무사.', 'Mark, Lucy, 세무사', '$0', 'Updated checklist'],
  ['Every quarter', 'Reportable transaction list for Form 5472 (capital, draws, fees between Lucy, H2 Company and BRAND LLC) to the CPA; Section 301 status check with the broker.', 'Mark, CPA, broker', '$0', 'Quarter end list and broker note'],
];
children.push(gridTable(['When', 'Item', 'Owner', 'Cost', 'Evidence'], CW, cal, { cantSplit: true }));
children.push(small('Basis: filing names, fees and penalties are facts from the brief (facts.json). Timing relative to formation is a management estimate from the launch critical path (01_launch_critical_path). Anything marked [PLACEHOLDER] is confirmed by the named professional before the date is relied on.', { before: 80 }));
children.push(spacer(240));

// ---------- Appendix A: numbers ----------
children.push(h1('Appendix A. Numbers used in this document'));
children.push(p('Every figure below is read by the generator from build/figures.json (model) or build/facts.json (static facts). Rerun build/build_sops.py after the workbook changes and the document updates itself.'));
const NW = [3300, 1850, 1450, 3048];
const nums = [
  ['Hero SKU 8K 10M sale price (list)', `$${F.num(hero.price, 2)} ($${F.num(hero.list, 2)})`, 'Evidence based', 'unit_economics.hero.price'],
  ['Hero landed cost, August basis (share of price)', `$${F.num(hero.landed_cost, 2)} (${F.pct(hero.landed_pct, 0)})`, 'Evidence based', 'unit_economics.hero.landed_cost'],
  ['Hero net per unit (margin)', `$${F.num(hero.net, 2)} (${F.pct(hero.net_pct)})`, 'Evidence based', 'unit_economics.hero.net'],
  ['Hero contribution before advertising at ASP $42', `$${F.num(hero.contribution_before_ads_at_asp42, 2)}`, 'Evidence based', 'unit_economics.hero.contribution_before_ads_at_asp42'],
  ['Hero duty per unit, copper versus fiber', `$${F.num(launchSkus[1].duty_copper, 2)} versus $${F.num(launchSkus[1].duty_fiber, 2)}`, 'Evidence based rates', 'unit_economics.per_sku (8K_10M)'],
  ['Hero freight per unit, air versus ocean', `$${F.num(launchSkus[1].air_freight, 2)} versus $${F.num(launchSkus[1].ocean_freight, 2)}`, 'Management estimate', 'unit_economics.per_sku (8K_10M)'],
  ['Landed COGS % Medium: copper, fiber, stress', `${F.pct(duty.copper.landed_pct_medium)}, ${F.pct(duty.fiber.landed_pct_medium)}, ${F.pct(duty.stress.landed_pct_medium)}`, 'Model', 'duty.*.landed_pct_medium'],
  ['18 month net Medium: copper, fiber, stress', `${F.money_k(fig.full_recalc.duty.copper.m18_net)}, ${F.money_k(fig.full_recalc.duty.fiber.m18_net)}, ${F.money_k(fig.full_recalc.duty.stress.m18_net)}`, 'Model', 'full_recalc.duty.*.m18_net'],
  ['Medium 18 month revenue, net, margin, breakeven', `${F.money_m(M.m18.revenue)}, ${F.money_k(M.m18.net)}, ${F.pct(M.m18.net_margin)}, ${F.month_label(M.m18.breakeven_month)}`, 'Model', 'scenarios.Medium.m18'],
  ['Medium peak cash need (month), funding gap (needed by)', `${F.money_k(M.cash.peak_need)} (${M.cash.peak_need_month_label}), ${F.money_k(M.cash.funding_gap)} (${F.month_label(M.cash.funding_needed_by_month)})`, 'Model', 'scenarios.Medium.cash'],
  ['Low peak cash need (month), funding gap', `${F.money_k(L.cash.peak_need)} (${L.cash.peak_need_month_label}), ${F.money_k(L.cash.funding_gap)}`, 'Model', 'scenarios.Low.cash'],
  ['High peak cash need (month), funding gap', `${F.money_k(H.cash.peak_need)} (${H.cash.peak_need_month_label}), ${F.money_k(H.cash.funding_gap)}`, 'Model', 'scenarios.High.cash'],
  ['Medium Month 12 units, revenue, cumulative net', `${F.units(M.monthly[11].units)}, ${F.money(M.monthly[11].revenue)}, ${F.money(M.monthly[11].cum_net)}`, 'Model', 'scenarios.Medium.monthly (Month 12)'],
  ['Inventory cover months Low, Medium, High', `${coverL}, ${coverM}, ${coverH}`, 'Management estimate', 'Assumptions tab (Active column)'],
  ['Lead time air, ocean; ocean from Month', `${airLead} month, ${oceanLead} months; Month ${oceanFrom}`, 'Management estimate', 'Assumptions tab'],
  ['PO deposit at order', F.pct(deposit, 0), 'Management estimate', 'Assumptions tab'],
  ['Amazon payout lag', `${payoutLag} days`, 'Evidence based', 'facts.amazon.payout_lag_days'],
  ['Product liability insurance trigger and cover', `$${F.num(facts.amazon.insurance_trigger_monthly_sales_usd)} monthly sales; ${F.money_m(facts.amazon.product_liability_insurance_usd)} cover; Medium Month ${insTriggerMonth} (passes by ${F.money(insMargin)}), Low Month ${L.insurance_trigger_month}, High Month ${H.insurance_trigger_month}`, 'Evidence based', 'facts.amazon; scenarios.*.insurance_trigger_month'],
  ['Hero FOB per unit at the model FX (PO deposit base)', `$${F.num(fobUnit, 2)}`, 'Evidence based (H2 cost base times 0.92, management estimate)', 'unit_economics.hero.fob_usd'],
  ['Medium Month 12 target closing stock, arrivals, on hand, in transit', `${F.units(Math.round(M.monthly[11].closing_stock))} units, ${F.units(Math.round(M.monthly[11].arrivals))} units, ${F.money_k(M.monthly[11].inventory_on_hand)}, ${F.money_k(M.monthly[11].inventory_in_transit)}`, 'Model', 'scenarios.Medium.monthly (Month 12)'],
  ['Form 5472 penalty; Wyoming formation and annual fee', `$${F.num(facts.entity.form_5472_penalty_usd)}; $${F.num(facts.entity.formation_cost_usd)} and $${F.num(facts.entity.annual_cost_usd)} per year`, 'Evidence based', 'facts.entity'],
  ['Continuous bond; broker entry fee', `$${F.num(facts.logistics.broker_bond_annual_usd)} per year; $${F.num(facts.logistics.broker_entry_fee_usd)} per entry`, 'Evidence based; management estimate', 'facts.logistics'],
  ['Trademark USPTO fee', `$${F.num(facts.trademark.uspto_fee_per_class_usd)} per class plus attorney fees`, 'Evidence based', 'facts.trademark'],
  ['Plan B revenue and net, Year 1 and Year 3', `${F.money_k(fig.plan_b.Y1.revenue)} / ${F.money_k(fig.plan_b.Y1.net)}; ${F.money_k(fig.plan_b.Y3.revenue)} / ${F.money_k(fig.plan_b.Y3.net)}`, 'Management estimate', 'plan_b'],
];
children.push(gridTable(['Figure', 'Value', 'Basis', 'Source key'], NW, nums, { cantSplit: true }));
children.push(spacer());

// ---------- Appendix B: glossary ----------
children.push(h1('Appendix B. Glossary'));
const gloss = [
  ['AOC', 'active optical cable'], ['HDMI', 'High Definition Multimedia Interface'], ['ASP', 'average selling price'], ['COGS', 'cost of goods sold'],
  ['FOB', 'free on board (the price at the factory\'s port)'], ['DDP', 'delivered duty paid'], ['FBA', 'Fulfillment by Amazon'],
  ['FNSKU', 'Fulfillment Network Stock Keeping Unit (Amazon\'s barcode)'], ['ASIN', 'Amazon Standard Identification Number'],
  ['ACOS', 'advertising cost of sales (ad spend divided by ad attributed sales)'], ['TACOS', 'total advertising cost of sales (ad spend divided by total sales)'],
  ['PPC', 'pay per click advertising'], ['3PL', 'third party logistics warehouse'], ['HTS', 'Harmonized Tariff Schedule'], ['CBP', 'US Customs and Border Protection'],
  ['MFN', 'most favored nation duty rate'], ['Section 301', 'US trade law tariffs on China origin goods'], ['IEEPA', 'International Emergency Economic Powers Act'],
  ['EIN', 'Employer Identification Number'], ['LLC', 'limited liability company'], ['IRS', 'Internal Revenue Service'], ['PE', 'permanent establishment'],
  ['ODI', 'overseas direct investment report (해외직접투자 신고)'], ['세무사', 'Korean licensed tax accountant'], ['FTA', 'Foreign Exchange Transactions Act (외국환거래법)'],
  ['USPTO', 'United States Patent and Trademark Office'], ['ATC', 'HDMI Authorized Test Center'], ['HDMI LA', 'HDMI Licensing Administrator'],
  ['FCC', 'Federal Communications Commission'], ['Prop 65', 'California Proposition 65 warning label'], ['CL3', 'in wall cable fire rating (Class 3)'],
  ['KVM', 'keyboard, video and mouse extension'], ['DTC', 'direct to consumer website'], ['LOC', 'line of credit'], ['QC', 'quality control'],
  ['MOQ', 'minimum order quantity'], ['PO', 'purchase order'], ['SKU', 'stock keeping unit'], ['UPH', 'units per hour'], ['KPI', 'key performance indicator'],
  ['SOP', 'standard operating procedure'], ['VA', 'virtual assistant'], ['eARC', 'enhanced audio return channel'], ['Gbps', 'gigabits per second'],
  ['MPF, HMF', 'merchandise processing fee and harbor maintenance fee (US customs fees on entries)'], ['ACE', 'Automated Commercial Environment (the CBP entry system)'],
  ['IP', 'intellectual property (a trademark, patent or copyright complaint on Amazon)'], ['CPA', 'certified public accountant (US licensed accountant)'], ['SSN', 'Social Security number (US personal tax number; not needed for the EIN)'],
  ['AQL', 'acceptable quality limit (the sampling standard for a QC inspection)'], ['B2B', 'business to business (sales to installers, integrators and resellers)'], ['QR', 'quick response code (the square barcode on the insert card that opens the setup video)'],
];
children.push(gridTable(['Term', 'Meaning'], [1800, CONTENT_W - 1800], gloss));

// ---------- document ----------
const doc = new Document({
  creator: 'BRAND US Launch build',
  title: '05 Operating Processes',
  description: 'SOPs for BRAND LLC: purchasing and QC, customs, inbound, listings and PPC, inventory, returns, monthly close, compliance calendar',
  styles: {
    default: { document: { run: { font: FONT, size: 22, color: INK } } },
    paragraphStyles: [
      { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { font: FONT, size: 32, bold: true, color: ACCENT }, paragraph: { spacing: { before: 360, after: 160 }, outlineLevel: 0 } },
      { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { font: FONT, size: 26, bold: true, color: INK }, paragraph: { spacing: { before: 240, after: 120 }, outlineLevel: 1 } },
      { id: 'Heading3', name: 'Heading 3', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { font: FONT, size: 22, bold: true, color: ACCENT }, paragraph: { spacing: { before: 160, after: 80 }, outlineLevel: 2 } },
    ],
  },
  numbering,
  sections: [{
    properties: { page: { size: { width: PAGE_W, height: PAGE_H }, margin: { top: MARGIN, bottom: MARGIN, left: MARGIN, right: MARGIN } } },
    headers: { default: new Header({ children: [new Paragraph({ children: [run('BRAND LLC. 05 Operating Processes. Built ' + built + '.', { size: 16, color: MUTED })], alignment: AlignmentType.RIGHT })] }) },
    footers: { default: new Footer({ children: [new Paragraph({ children: [run('Page ', { size: 16, color: MUTED }), new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 16, color: MUTED }), run(' of ', { size: 16, color: MUTED }), new TextRun({ children: [PageNumber.TOTAL_PAGES], font: FONT, size: 16, color: MUTED })], alignment: AlignmentType.CENTER })] }) },
    children,
  }],
});

Packer.toBuffer(doc).then((buf) => {
  const out = path.join(OUT, '05_Operating_Processes.docx');
  fs.writeFileSync(out, buf);
  console.log('wrote', out, buf.length, 'bytes');
});

#!/usr/bin/env node
/* BRAND US launch proposal deck generator.
   Usage: node build_deck.js --lang en|ko
   Reads build/figures.json, build/facts.json, build/fmt.js and build/strings/<lang>.json.
   Writes output/01_Business_Proposal_EN.pptx or output/02_Business_Proposal_KO.pptx.
   Every model number is formatted from figures.json through fmt.js at build time; strings carry tokens only. */
'use strict';
const fs = require('fs');
const path = require('path');
const { execFileSync } = require('child_process');
const pptxgen = require('pptxgenjs');
const React = require('react');
const ReactDOMServer = require('react-dom/server');
const Fi = require('react-icons/fi');
const sharp = require('sharp');
const fmt = require('./fmt.js');
const { applyTheme } = require('/mnt/skills/public/pptx/scripts/apply_theme.js');

const ROOT = path.resolve(__dirname, '..');
const BUILD = __dirname;
const argLang = (process.argv.find((a) => a.startsWith('--lang')) || '--lang=en').replace(/^--lang=?/, '') || process.argv[process.argv.indexOf('--lang') + 1] || 'en';
const LANG = argLang === 'ko' ? 'ko' : 'en';
// --out=<path> writes somewhere else (QA builds); the default is the delivered file.
const argOut = (process.argv.find((a) => a.startsWith('--out=')) || '').replace(/^--out=/, '');
const OUT_FILE = argOut || path.join(ROOT, 'output', LANG === 'ko' ? '02_Business_Proposal_KO.pptx' : '01_Business_Proposal_EN.pptx');

const F = JSON.parse(fs.readFileSync(path.join(BUILD, 'figures.json'), 'utf8'));
const K = JSON.parse(fs.readFileSync(path.join(BUILD, 'facts.json'), 'utf8'));
const ST = JSON.parse(fs.readFileSync(path.join(BUILD, 'strings', LANG + '.json'), 'utf8'));

if (!fs.existsSync(path.join(BUILD, 'assets', 'fiber_bg.png'))) execFileSync('python3', [path.join(BUILD, 'make_assets.py')]);
// Deck specific flowchart renders (terse labels, slide shaped layout) so labels read at 9 pt or more on the page.
const DECK_FLOW = path.join(BUILD, 'assets', 'deck_flow');
if (!fs.existsSync(path.join(DECK_FLOW, 'dims.json'))) execFileSync('python3', [path.join(BUILD, 'deck_flowcharts.py')]);
// The KO deck uses Korean renders of the same three charts (deck_flowcharts_ko.py): same steps, owners, weeks and thresholds.
if (LANG === 'ko' && !fs.existsSync(path.join(DECK_FLOW, 'dims_ko.json'))) execFileSync('python3', [path.join(BUILD, 'deck_flowcharts_ko.py')]);
const DF = JSON.parse(fs.readFileSync(path.join(DECK_FLOW, LANG === 'ko' ? 'dims_ko.json' : 'dims.json'), 'utf8'));
const flowFile = (f) => (LANG === 'ko' ? f.replace(/_deck\.png$/, '_deck_ko.png') : f);

// ---------------------------------------------------------------- palette and theme
const HEX = { bg: '0A0F1E', card: '141C36', card2: '10172D', cyan: '2DE2FF', violet: '8B5CF6', magenta: 'FF4FD8', text: 'F3F6FF', muted: '9AA5C4', green: '3DDC97', amber: 'FFB547', blue: '4F7CFF', grid: '26304F', head: '1B2447', border: '2A3558' };
const FONT = LANG === 'ko' ? 'Noto Sans KR' : 'Calibri';
const THEME = {
  name: 'BRAND Fiber',
  headFontFace: FONT,
  bodyFontFace: FONT,
  colors: { dk1: HEX.text, lt1: HEX.bg, dk2: HEX.muted, lt2: HEX.card, accent1: HEX.cyan, accent2: HEX.violet, accent3: HEX.magenta, accent4: HEX.green, accent5: HEX.amber, accent6: HEX.blue, hlink: HEX.cyan, folHlink: HEX.violet },
};

// ---------------------------------------------------------------- formatted values (tokens)
const { money, money_k, money_m, money_auto, pct, pts, num, units } = fmt;
// Month labels come from the strings file ("Month {n}" in EN, "{n}개월차" in KO) so the number never changes, only the wording.
const month_label = (m) => {
  if (m === null || m === undefined || m === '' || m === 'none') return ST.common.none;
  const n = Number(m);
  if (Number.isNaN(n)) return String(m);
  return T(ST.common.month, { n: String(Math.trunc(n)) });
};
const pts_label = (x, d) => (ST.common.points ? T(ST.common.points, { n: num(x * 100, d) }) : pts(x, d));
// Dates in facts.json are English ("February 20, 2026"); the KO deck shows them in Korean form.
const MONTHS = { January: 1, February: 2, March: 3, April: 4, May: 5, June: 6, July: 7, August: 8, September: 9, October: 10, November: 11, December: 12 };
function dateLabel(s) {
  if (LANG !== 'ko' || typeof s !== 'string') return s;
  let m;
  if ((m = s.match(/^(\w+) (\d+) to (\w+) (\d+), (\d{4})$/))) return `${m[5]}년 ${MONTHS[m[1]]}월 ${m[2]}일부터 ${MONTHS[m[3]]}월 ${m[4]}일까지`;
  if ((m = s.match(/^(\w+) (\d+), (\d{4})$/))) return `${m[3]}년 ${MONTHS[m[1]]}월 ${m[2]}일`;
  if ((m = s.match(/^late (\w+) (\d{4})$/))) return `${m[2]}년 ${MONTHS[m[1]]}월 말`;
  if ((m = s.match(/^(\w+) (\d{4})$/))) return `${m[2]}년 ${MONTHS[m[1]]}월`;
  return s;
}
const money_km = (x) => (Math.abs(x) >= 1000000 ? money_m(x) : money_k(x)); // $K below one million, $M above (spec 1.9)
const pctAuto = (x) => (Math.abs(x * 100 - Math.round(x * 100)) < 1e-9 ? pct(x, 0) : pct(x, 1));
const moneyAuto2 = (x) => (Math.abs(x - Math.round(x)) < 1e-9 ? money(x, 0) : money(x, 2));
const krwRound = (x) => (x % 1000000 === 0 ? 'KRW ' + num(x / 1000000, 0) + 'M' : fmt.krw_m(x));
const signedPct = (x) => (x >= 0 ? '+' : '') + pctAuto(x);
const V = {};
const SC = { L: 'Low', M: 'Medium', H: 'High' };
V.fx = num(F.meta.fx); V.fx_lo = num(F.meta.fx_scenarios[0]); V.fx_hi = num(F.meta.fx_scenarios[2]);
V.built = dateLabel(F.meta.built);
for (const [k, name] of Object.entries(SC)) {
  const s = F.scenarios[name];
  V['m18_rev_' + k] = money_auto(s.m18.revenue);
  V['m18_net_' + k] = money_k(s.m18.net);
  V['m18_margin_' + k] = pct(s.m18.net_margin);
  V['be_' + k] = s.m18.breakeven_month == null ? ST.common.none : month_label(s.m18.breakeven_month);
  V['be_' + k + '_en'] = fmt.month_label(s.m18.breakeven_month); // workbook wording (Month 8), quoted once in the KO deck
  V['rr_' + k] = money_auto(s.m18.run_rate);
  V['exit_' + k] = pct(s.m18.exit_margin);
  V['m18_units_' + k] = units(s.m18.units);
  V['m24_rev_' + k] = money_auto(s.m24.revenue);
  V['m24_net_' + k] = money_k(s.m24.net);
  for (let y = 1; y <= 5; y++) {
    const a = s.annual['Y' + y];
    V[`y${y}_rev_${k}`] = money_auto(a.revenue);
    V[`y${y}_net_${k}`] = money_km(a.net);
    V[`y${y}_margin_${k}`] = pct(a.net_margin);
    V[`asp_y${y}_${k}`] = moneyAuto2(a.asp);
    V[`landed_y${y}_${k}`] = pctAuto(a.landed_pct);
    V[`ads_y${y}_${k}`] = pctAuto(a.ads_pct);
    V[`fixed_y${y}_${k}`] = money_k(a.fixed + a.seller_fee + a.compliance);
    V[`y${y}_units_${k}`] = units(a.amazon_units);
    for (const [ch, short] of Object.entries({ amazon_us: 'amz', ebay_walmart: 'ebay', b2b: 'b2b', dtc: 'dtc', amazon_canada: 'ca', pro_line: 'pro' })) V[`y${y}_${short}_${k}`] = money_auto(a.revenue_by_channel[ch]);
  }
  for (let y = 3; y <= 5; y++) V[`g${y}_${k}`] = signedPct(s.annual['Y' + y].amazon_units / s.annual['Y' + (y - 1)].amazon_units - 1);
  V['peak_' + k] = money_k(s.cash.peak_need);
  V['peak_month_' + k] = month_label(s.cash.peak_need_month);
  V['gap_' + k] = money_k(s.cash.funding_gap);
  V['gap_month_' + k] = month_label(s.cash.funding_needed_by_month);
  V['runway_' + k] = typeof s.cash.runway_months_day90 === 'number' ? num(s.cash.runway_months_day90, 1) : String(s.cash.runway_months_day90);
  V['cash24_' + k] = money_k(s.cash.cash_balance_month24_after_funding);
  V['cash_y5_' + k] = money_k(s.cash.year_end_balance_after_funding.Y5);
  V['m12_rev_' + k] = money(s.gates.month12.metrics.rev.value);
  V['m12_cum_' + k] = money(s.gates.month12.metrics.cum.value);
  V['m12_tacos_' + k] = pctAuto(s.gates.month12.metrics.tacos.value);
  V['m12_ret_' + k] = pctAuto(s.gates.month12.metrics.ret.value);
  V['m12_opcf_' + k] = money(s.gates.month12.metrics.opcf.value);
  V['d90_tacos_' + k] = pctAuto(s.gates.day90.metrics.tacos.value);
  V['d90_ret_' + k] = pctAuto(s.gates.day90.metrics.ret.value);
  V['d90_cpu_' + k] = money(s.gates.day90.metrics.cpu.value, 2);
  V['d90_runway_' + k] = num(s.gates.day90.metrics.runway.value, 1);
  V['gate_d90_' + k] = s.gates.day90.result;
  V['gate_m12_' + k] = s.gates.month12.result;
}
V.opening_capital = money_k(F.scenarios.Medium.cash.opening_capital);
V.cum_net_5y_M = money_m(F.full_recalc.medium_base.cum_net_5y);
const H = F.unit_economics.hero;
V.hero_price = money(H.price, 2); V.hero_list = money(H.list, 2); V.hero_landed = money(H.landed_cost, 2); V.hero_landed_pct = pctAuto(H.landed_pct);
V.hero_ref = money(H.referral, 2); V.hero_fba = money(H.fba, 2); V.hero_ads = money(H.ads, 2); V.hero_ret = money(H.returns_storage, 2);
V.hero_net = money(H.net, 2); V.hero_net_pct = pct(H.net_pct); V.hero_korea = money(H.korea_wholesale_gross, 2);
const HS = F.unit_economics.per_sku.find((p) => p.sku === '8K_10M');
V.hero_fba_sku = money(HS.fba_fee, 2); V.hero_net_sku_air = money(HS.net_air, 2);
V.hero_cba = money(H.contribution_before_ads, 2); V.hero_cba42 = money(H.contribution_before_ads_at_asp42, 2); V.hero_fob = money(H.fob_usd, 2); V.h2_fob = V.hero_fob;
const CB = F.unit_economics.cost_bases;
V.cb_h2_landed = money(CB.h2.landed, 2); V.cb_h2_net = money(CB.h2.net, 2); V.cb_aug_landed = money(CB.august.landed, 2); V.cb_aug_net = money(CB.august.net, 2);
V.cb_ali_lo_landed = money(CB.alibaba_low.landed, 2); V.cb_ali_lo_net = money(CB.alibaba_low.net, 2); V.cb_ali_hi_landed = money(CB.alibaba_high.landed, 2); V.cb_ali_hi_net = money(CB.alibaba_high.net, 2);
V.ali_lo = money(K.august.alibaba_benchmark.fob_low_usd); V.ali_hi = money(K.august.alibaba_benchmark.fob_high_usd); V.fob_factor = String(K.fx.fob_factor_vs_korea_landed);
const D = F.duty;
V.duty_copper = pct(D.copper.rate); V.duty_fiber = pct(D.fiber.rate); V.duty_stress = pct(D.stress.rate); V.duty_check = pct(D.fiber_list3_check.rate);
V.landed_copper = pctAuto(D.copper.landed_pct_medium); V.landed_fiber = pct(D.fiber.landed_pct_medium); V.landed_stress = pct(D.stress.landed_pct_medium);
V.hero_net_copper = money(D.copper.hero_net, 2); V.hero_net_fiber = money(D.fiber.hero_net, 2); V.hero_net_stress = money(D.stress.hero_net, 2);
V.fiber_uplift = money(D.fiber.hero_net - D.copper.hero_net, 2); V.stress_penalty = money(D.stress.hero_net - D.copper.hero_net, 2);
const FR = F.full_recalc;
V.fiber_m18_net = money_k(FR.duty.fiber.m18_net); V.fiber_be = month_label(FR.duty.fiber.m18_breakeven_month); V.fiber_y5_net = money_k(FR.duty.fiber.y5_net);
V.stress_m18_net = money_k(FR.duty.stress.m18_net); V.stress_y5_net = money_k(FR.duty.stress.y5_net);
V.fiber_m18_delta = money_k(FR.duty.fiber.m18_net_vs_base); V.stress_m18_delta_abs = money_k(Math.abs(FR.duty.stress.m18_net_vs_base));
V.fiber_gap = money_k(FR.duty.fiber.funding_gap); V.stress_gap = money_k(FR.duty.stress.funding_gap);
V.fx_lo_m18_net = money_k(FR.fx['1300'].m18_net); V.fx_hi_m18_net = money_k(FR.fx['1550'].m18_net);
V.mfn_copper = pct(K.duty.copper.mfn); V.list_copper = pctAuto(K.duty.copper.section_301_list); V.fl_rate = pct(K.duty.copper.section_301_forced_labor); V.pending_add = pts_label(K.duty.stress.pending_301_add, 0);
V.y5_net_base = money_k(F.sensitivity.base_y5_net);
const tornado = F.sensitivity.tornado;
V.ali_y5_net = money_m(tornado.find((t) => t.driver.startsWith('Factory')).low_value);
V.sens_mid_swing = money_k(tornado.filter((t) => /Duty|ASP|Ad spend/.test(t.driver)).reduce((a, t) => a + t.swing, 0) / 3);
for (const [k, t] of Object.entries(F.gate_thresholds)) V['gt_' + k] = t.unit === '%' ? pctAuto(t.value) : t.unit === '$' ? money(t.value) : num(t.value);
for (const y of ['Y1', 'Y2', 'Y3']) { const p = F.plan_b[y]; const yy = y.toLowerCase(); V[`pb_${yy}_rev`] = money_k(p.revenue); V[`pb_${yy}_net`] = money_k(p.net); V[`pb_${yy}_b2b`] = money_k(p.b2b_revenue); V[`pb_${yy}_korea`] = money_k(p.korea_gross_profit); }
// facts
V.wy_form = money(K.entity.formation_cost_usd); V.wy_annual = money(K.entity.annual_cost_usd); V.de_annual = money(K.entity.delaware_annual_cost_usd); V.penalty_5472 = money(K.entity.form_5472_penalty_usd);
V.tm_fee = money(K.trademark.uspto_fee_per_class_usd); V.bond = money(K.logistics.broker_bond_annual_usd); V.entry_fee = money(K.logistics.broker_entry_fee_usd);
V.ins_amount = '$' + num(K.amazon.product_liability_insurance_usd / 1000000, 0) + 'M'; V.ins_trigger = money_k(K.amazon.insurance_trigger_monthly_sales_usd);
V.ins_month_M = month_label(F.scenarios.Medium.insurance_trigger_month); V.ins_month_L = month_label(F.scenarios.Low.insurance_trigger_month); V.ins_month_H = month_label(F.scenarios.High.insurance_trigger_month);
V.kr_ofar = krwRound(K.korea.overseas_financial_account_report_threshold_krw); V.reg_no = K.korea.company_registration_number;
V.seller_fee = money(K.amazon.professional_fee_monthly_usd, 2); V.payout_lag = num(K.amazon.payout_lag_days); V.referral = pctAuto(K.amazon.referral_rate);
const A = K.august.m18_assumptions;
V.fixed_monthly = money(A.fixed_monthly); V.launch_one_time = money(A.launch_one_time); V.asp42 = money(A.asp);
A.ads_phases.forEach((p, i) => (V['ads_p' + (i + 1)] = pctAuto(p[2])));
V.deal_units = units(K.august.deal.units); V.deal_cost = fmt.krw_m(K.august.deal.cost_krw); V.deal_rev = fmt.krw_m(K.august.deal.revenue_ex_vat_krw); V.deal_gp = fmt.krw_m(K.august.deal.gross_profit_krw);
V.deal_margin = pct(K.august.deal.gross_margin); V.deal_gp_usd = money(K.august.deal.gross_profit_usd_approx); V.share_units = pctAuto(K.august.deal.share_units_8k_short_runs); V.share_gp = pctAuto(K.august.deal.share_gp_8k_short_runs);
V.air_kg = money(K.logistics.air_per_kg_usd, 2); V.ocean_kg = money(K.logistics.ocean_per_kg_usd, 2); V.fund_ask = money_k(K.august.funding_ask_usd);
V.ieepa_date = dateLabel(K.tariff_history.ieepa_struck_down); V.s122_period = dateLabel(K.tariff_history.section_122_period); V.fl_date = dateLabel(K.tariff_history.section_301_forced_labor_effective); V.s301_2_opened = dateLabel(K.tariff_history.second_301_investigation_opened);
V.fba_from = money(F.levers[7].from, 2); V.fba_to = money(F.levers[7].to, 2);
V.file = '';

const USED_KEYS = new Set();
const SLIDE_KEYS = {};
let CUR = 0;
function T(str, extra) {
  if (str == null) return '';
  return String(str).replace(/\{([a-zA-Z0-9_]+)\}/g, (m, k) => {
    if (extra && k in extra) return extra[k];
    if (k === 'full_size' && extra && extra.file) return T(ST.common.full_size, extra);
    if (!(k in V)) throw new Error('Missing token ' + k + ' in: ' + str);
    USED_KEYS.add(k);
    (SLIDE_KEYS[CUR] = SLIDE_KEYS[CUR] || new Set()).add(k);
    return V[k];
  });
}

// ---------------------------------------------------------------- icons
const iconCache = {};
async function icon(name, colorHex, px = 256) {
  const key = name + colorHex + px;
  if (iconCache[key]) return iconCache[key];
  const Comp = Fi[name] || Fi.FiCircle;
  const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(Comp, { color: '#' + colorHex, size: px, strokeWidth: 1.75 }));
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  iconCache[key] = 'image/png;base64,' + buf.toString('base64');
  return iconCache[key];
}

// ---------------------------------------------------------------- deck
const pres = new pptxgen();
pres.layout = 'LAYOUT_WIDE';
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.author = 'H2 Company and BRAND LLC';
pres.title = ST.s01.title;
pres.subject = 'BRAND US launch proposal';
const C = pres.SchemeColor;
const W = 13.333, M = 0.5, CW = 12.333, GAP = 0.3; // GAP: minimum gutter between sibling cards
const FOOTER = T(ST.common.footer);

pres.defineSlideMaster({
  title: 'TITLE', background: { path: path.join(BUILD, 'assets', 'fiber_bg.png') },
  objects: [
    { placeholder: { options: { name: 'title', type: 'title', x: 0.75, y: 2.15, w: 9.2, h: 1.6, fontSize: 40, bold: true, color: HEX.text, valign: 'bottom', align: 'left', margin: 0 }, text: '' } },
    { placeholder: { options: { name: 'body', type: 'body', x: 0.75, y: 3.85, w: 8.6, h: 1.1, fontSize: 18, color: HEX.muted, valign: 'top', margin: 0 }, text: '' } },
  ],
});
pres.defineSlideMaster({
  title: 'SECTION', background: { path: path.join(BUILD, 'assets', 'fiber_section_bg.png') },
  objects: [
    { placeholder: { options: { name: 'title', type: 'title', x: 0.75, y: 4.1, w: 10, h: 1.1, fontSize: 36, bold: true, color: HEX.text, valign: 'bottom', align: 'left', margin: 0 }, text: '' } },
    { placeholder: { options: { name: 'body', type: 'body', x: 0.75, y: 5.25, w: 10, h: 0.5, fontSize: 16, color: HEX.cyan, valign: 'top', margin: 0 }, text: '' } },
  ],
  slideNumber: { x: 12.3, y: 7.02, w: 0.6, h: 0.3, fontSize: 9, color: HEX.muted, align: 'right' },
});
pres.defineSlideMaster({
  title: 'CONTENT', background: { color: HEX.bg },
  objects: [
    { placeholder: { options: { name: 'title', type: 'title', x: M, y: 0.32, w: CW, h: 0.8, fontSize: 32, bold: true, color: HEX.text, valign: 'middle', margin: 0 }, text: '' } },
    { text: { text: FOOTER, options: { x: M, y: 7.0, w: 11.2, h: 0.3, fontSize: 9, color: HEX.muted, margin: 0, valign: 'middle' } } },
  ],
  slideNumber: { x: 12.3, y: 7.0, w: 0.6, h: 0.3, fontSize: 9, color: HEX.muted, align: 'right' },
});

// helpers
let slideNo = 0;
function content(title, sectionTitle, notes) {
  const s = pres.addSlide({ masterName: 'CONTENT', sectionTitle });
  slideNo++; CUR = slideNo;
  s.addText(T(title), { placeholder: 'title' });
  if (notes) s.addNotes(T(notes));
  return s;
}
function card(s, x, y, w, h, opts = {}) {
  s.addShape(pres.ShapeType.roundRect, { x, y, w, h, fill: { color: opts.fill || C.background2 }, line: opts.line ? { color: opts.line, width: 0.75 } : { color: opts.fill || HEX.card, width: 0 }, rectRadius: opts.r || 0.12, objectName: opts.name || 'card' });
}
// Hangul inside an EN Calibri run falls back badly in LibreOffice (stray spaces in the PDF), so Hangul substrings get
// their own run in Noto Sans KR. The KO deck already uses that font for everything.
const HANGUL = /[\u1100-\u11FF\u3130-\u318F\uAC00-\uD7A3]+/;
function hangulRuns(str, base) {
  if (LANG === 'ko' || !HANGUL.test(str)) return null;
  const out = [];
  let rest = str;
  while (rest.length) {
    const m = HANGUL.exec(rest);
    if (!m) { out.push({ text: rest, options: Object.assign({}, base) }); break; }
    if (m.index > 0) out.push({ text: rest.slice(0, m.index), options: Object.assign({}, base) });
    out.push({ text: m[0], options: Object.assign({}, base, { fontFace: 'Noto Sans KR' }) });
    rest = rest.slice(m.index + m[0].length);
  }
  return out;
}
function text(s, str, o, extra) {
  const opts = Object.assign({ fontSize: 14, color: C.text1, margin: 0, isTextBox: true, valign: 'top', align: 'left', paraSpaceAfter: 0 }, o);
  if (typeof str === 'string') str = T(str, extra);
  if (typeof str === 'string') { const hr = hangulRuns(str, { fontSize: opts.fontSize, color: opts.color, bold: opts.bold, italic: opts.italic }); if (hr) str = hr; }
  s.addText(str, opts);
}
function bullets(s, items, o) {
  const size = o.fontSize || 14;
  const runs = items.map((it, i) => ({ text: T(it), options: { bullet: { indent: Math.round(size * 0.9) }, breakLine: i < items.length - 1, paraSpaceAfter: o.gap != null ? o.gap : Math.round(size * 0.45) } }));
  s.addText(runs, Object.assign({ fontSize: size, color: C.text1, margin: 0, isTextBox: true, valign: 'top', align: 'left' }, o));
}
function runs(s, parts, o) {
  // parts: [{t, b, c, size, br}]
  const arr = [];
  parts.forEach((p) => {
    const base = { bold: !!p.b, color: p.c || C.text1, fontSize: p.size || o.fontSize || 14, breakLine: false, paraSpaceAfter: p.gap || 0 };
    const t = T(p.t);
    const hr = (p.font ? null : hangulRuns(t, base)) || [{ text: t, options: p.font ? Object.assign({}, base, { fontFace: p.font }) : base }];
    hr[hr.length - 1].options.breakLine = !!p.br;
    arr.push(...hr);
  });
  s.addText(arr, Object.assign({ fontSize: o.fontSize || 14, color: C.text1, margin: 0, isTextBox: true, valign: 'top', align: 'left' }, o));
}
function table(s, head, rows, o) {
  const size = o.fontSize || 11;
  const border = { type: 'solid', pt: 0.5, color: HEX.border };
  const mk = (txt, opt) => ({ text: T(txt), options: Object.assign({ fontSize: size, color: HEX.text, border, margin: [0.04, 0.07, 0.04, 0.07], valign: 'middle', fontFace: FONT }, opt) });
  const data = [];
  if (head) data.push(head.map((h, i) => mk(h, { bold: true, color: HEX.cyan, fill: { color: HEX.head }, align: (o.align && o.align[i]) || 'left' })));
  rows.forEach((r, ri) => data.push(r.map((c, i) => {
    const cell = typeof c === 'object' && c !== null && 'text' in c ? c : { text: c };
    return mk(cell.text, Object.assign({ fill: { color: ri % 2 ? HEX.card2 : HEX.card }, align: (o.align && o.align[i]) || 'left' }, cell.options || {}));
  })));
  s.addTable(data, Object.assign({ x: o.x, y: o.y, w: o.w, colW: o.colW, autoPage: false, rowH: o.rowH }, o.extra || {}));
}
function stat(s, x, y, w, h, value, label, sub, color) {
  card(s, x, y, w, h);
  text(s, value, { x: x + 0.2, y: y + 0.15, w: w - 0.4, h: 0.75, fontSize: 30, bold: true, color: color || C.accent1, valign: 'middle' });
  text(s, label, { x: x + 0.2, y: y + 0.95, w: w - 0.4, h: 0.6, fontSize: 13, color: C.text1 });
  if (sub) text(s, sub, { x: x + 0.2, y: y + h - 0.6, w: w - 0.4, h: 0.5, fontSize: 11, color: C.text2 });
}
async function iconCircle(s, name, x, y, d, colorHex, bgHex) {
  s.addShape(pres.ShapeType.ellipse, { x, y, w: d, h: d, fill: { color: bgHex || HEX.head }, line: { color: bgHex || HEX.head, width: 0 } });
  s.addImage({ data: await icon(name, colorHex || HEX.cyan), x: x + d * 0.22, y: y + d * 0.22, w: d * 0.56, h: d * 0.56 });
}
// Place a deck flowchart render inside a white rounded frame, fitted to maxW x maxH, centered horizontally in [x, x + maxW].
function flowImage(s, key, x, y, maxW, maxH, file) {
  const d = DF[key]; const pad = 0.1;
  let w = maxW - 2 * pad, h = w * d.h / d.w;
  if (h > maxH - 2 * pad) { h = maxH - 2 * pad; w = h * d.w / d.h; }
  const fx = x + (maxW - w - 2 * pad) / 2;
  s.addShape(pres.ShapeType.roundRect, { x: fx, y, w: w + 2 * pad, h: h + 2 * pad, fill: { color: 'FFFFFF' }, line: { color: 'FFFFFF', width: 0 }, rectRadius: 0.08 });
  s.addImage({ path: path.join(DECK_FLOW, file), x: fx + pad, y: y + pad, w, h });
  const labelPt = d.font_px * (w / d.natural_w) * 72;
  return { x: fx, y, w: w + 2 * pad, h: h + 2 * pad, bottom: y + h + 2 * pad, labelPt };
}
// The theme carries only a Latin typeface, so Korean chart text names Noto Sans KR explicitly instead of the theme font.
const CHART_FONT = LANG === 'ko' ? FONT : '+mn-lt';
const CHART_BASE = () => ({
  chartArea: { fill: { color: HEX.card }, roundedCorners: false },
  plotArea: { fill: { color: HEX.card } },
  catAxisLabelColor: HEX.muted, valAxisLabelColor: HEX.muted, catAxisLabelFontSize: 11, valAxisLabelFontSize: 11,
  catAxisLabelFontFace: CHART_FONT, valAxisLabelFontFace: CHART_FONT, dataLabelFontFace: CHART_FONT, legendFontFace: CHART_FONT, catAxisTitleFontFace: CHART_FONT, valAxisTitleFontFace: CHART_FONT,
  valGridLine: { color: HEX.grid, size: 0.5 }, catGridLine: { style: 'none' },
  catAxisLineShow: false, valAxisLineShow: false,
  legendColor: HEX.muted, legendFontSize: 11, legendPos: 'b',
  dataLabelColor: HEX.text, dataLabelFontSize: 11,
  showTitle: false,
});
const NEG_FMT = '$#,##0;($#,##0)';
const NEG_FMT_K = '$#,##0,"K";($#,##0,"K")';

// ---------------------------------------------------------------- slides
(async () => {
  // 1 Title
  pres.addSection({ title: 'Opening' });
  {
    const s = pres.addSlide({ masterName: 'TITLE', sectionTitle: 'Opening' }); slideNo++; CUR = slideNo;
    text(s, ST.s01.kicker, { x: 0.75, y: 1.45, w: 6, h: 0.6, fontSize: 24, bold: true, color: C.accent1, charSpacing: 6 });
    s.addText(T(ST.s01.title), { placeholder: 'title' });
    s.addText(T(ST.s01.subtitle), { placeholder: 'body' });
    s.addShape(pres.ShapeType.roundRect, { x: 0.6, y: 6.05, w: 9.9, h: 1.0, fill: { color: HEX.bg, transparency: 30 }, line: { color: HEX.bg, width: 0, transparency: 30 }, rectRadius: 0.08 });
    text(s, ST.s01.meta_line1, { x: 0.75, y: 6.15, w: 9.5, h: 0.35, fontSize: 12, color: C.text2 });
    text(s, ST.s01.meta_line2, { x: 0.75, y: 6.5, w: 9.5, h: 0.5, fontSize: 12, color: C.text2 });
    s.addNotes(T(ST.s01.notes));
  }
  // 2 Executive summary
  {
    const s = content(ST.s02.title, 'Opening', ST.s02.notes);
    const n = ST.s02.stats.length, gap = GAP, w = (CW - gap * (n - 1)) / n;
    const colors = [C.accent1, C.accent4, C.accent1, C.accent4, C.accent5];
    ST.s02.stats.forEach((st, i) => stat(s, M + i * (w + gap), 1.3, w, 2.1, st.value, st.label, st.sub, colors[i]));
    card(s, M, 3.7, CW, 3.1);
    const icons = ['FiCpu', 'FiShoppingCart', 'FiCheckCircle'];
    for (let i = 0; i < ST.s02.points.length; i++) {
      const y = 3.9 + i * 0.95;
      await iconCircle(s, icons[i], M + 0.25, y + 0.05, 0.6);
      text(s, ST.s02.points[i], { x: M + 1.1, y, w: CW - 1.4, h: 0.85, fontSize: 14, valign: 'middle' });
    }
  }
  // 3 What we are deciding
  {
    const s = content(ST.s03.title, 'Opening', ST.s03.notes);
    const gap = GAP, w = (CW - 2 * gap) / 3;
    ST.s03.cards.forEach((c, i) => {
      const x = M + i * (w + gap);
      card(s, x, 1.3, w, 3.2);
      text(s, c.n, { x: x + 0.25, y: 1.45, w: 0.8, h: 0.8, fontSize: 40, bold: true, color: C.accent1 });
      text(s, c.head, { x: x + 1.05, y: 1.55, w: w - 1.3, h: 0.6, fontSize: 18, bold: true, valign: 'middle' });
      text(s, c.body, { x: x + 0.25, y: 2.35, w: w - 0.5, h: 2.05, fontSize: 15 });
    });
    card(s, M, 4.8, CW, 2.0, { fill: HEX.card2 });
    text(s, ST.s03.open_head, { x: M + 0.3, y: 4.92, w: 4, h: 0.4, fontSize: 15, bold: true, color: C.accent5 });
    bullets(s, ST.s03.open_items, { x: M + 0.3, y: 5.4, w: CW - 0.6, h: 1.3, fontSize: 14, gap: 5 });
  }
  const section = (key, name) => {
    pres.addSection({ title: name });
    const s = pres.addSlide({ masterName: 'SECTION', sectionTitle: name }); slideNo++; CUR = slideNo;
    s.addText(T(ST[key].title), { placeholder: 'title' });
    s.addText(T(ST[key].subtitle), { placeholder: 'body' });
    s.addNotes(T(ST[key].notes));
  };
  // 4 Section
  section('s04', 'Market and why fiber');
  // 5 Why copper fails
  {
    const s = content(ST.s05.title, 'Market and why fiber', ST.s05.notes);
    card(s, M, 1.3, 5.6, 5.5);
    text(s, ST.s05.left_head, { x: M + 0.3, y: 1.45, w: 5.0, h: 0.5, fontSize: 18, bold: true, color: C.accent1 });
    text(s, ST.s05.left_body, { x: M + 0.3, y: 2.05, w: 5.0, h: 3.3, fontSize: 15.5, lineSpacingMultiple: 1.1 });
    text(s, ST.s05.left_rule, { x: M + 0.3, y: 5.5, w: 5.0, h: 1.1, fontSize: 14, italic: true, color: C.accent5 });
    table(s, ST.s05.table_head, ST.s05.table_rows, { x: 6.4, y: 1.3, w: 6.43, colW: [1.75, 2.0, 2.68], fontSize: 12, rowH: 0.62 });
    text(s, ST.s05.footnote, { x: 6.4, y: 5.6, w: 6.43, h: 0.9, fontSize: 11, color: C.text2 });
  }
  // 6 Who buys long HDMI
  {
    const s = content(ST.s06.title, 'Market and why fiber', ST.s06.notes);
    const icons = ['FiMonitor', 'FiPlay', 'FiTool', 'FiBriefcase'];
    const w = 3.9, h = 2.55;
    for (let i = 0; i < 4; i++) {
      const x = M + (i % 2) * (w + GAP), y = 1.3 + Math.floor(i / 2) * (h + GAP);
      const seg = ST.s06.segments[i];
      card(s, x, y, w, h);
      await iconCircle(s, icons[i], x + 0.2, y + 0.2, 0.55);
      text(s, seg.head, { x: x + 0.9, y: y + 0.2, w: w - 1.1, h: 0.55, fontSize: 15, bold: true, valign: 'middle' });
      text(s, seg.body, { x: x + 0.2, y: y + 0.9, w: w - 0.4, h: 1.1, fontSize: 13 });
      text(s, seg.size, { x: x + 0.2, y: y + h - 0.6, w: w - 0.4, h: 0.5, fontSize: 11, color: C.accent5 });
    }
    const x = M + 2 * (w + GAP);
    card(s, x, 1.3, CW - 2 * (w + GAP), 5.4, { fill: HEX.card2 });
    text(s, ST.s06.evidence_head, { x: x + 0.25, y: 1.45, w: 3.5, h: 0.8, fontSize: 15, bold: true, color: C.accent4 });
    text(s, '{share_gp}', { x: x + 0.25, y: 2.35, w: 3.5, h: 1.0, fontSize: 44, bold: true, color: C.accent1 });
    text(s, ST.s06.evidence_body, { x: x + 0.25, y: 3.45, w: 3.5, h: 3.0, fontSize: 13 });
  }
  // 7 Price ladder
  {
    const s = content(ST.s07.title, 'Market and why fiber', ST.s07.notes);
    text(s, ST.s07.line, { x: M, y: 1.2, w: 8.3, h: 0.6, fontSize: 26, bold: true, italic: true, color: C.accent1 });
    text(s, ST.s07.line_sub, { x: M, y: 1.8, w: 8.3, h: 0.4, fontSize: 13, color: C.text2 });
    const rows = F.unit_economics.per_sku.filter((p) => p.launch_sku === 'Yes').map((p) => [p.sku.replace('_', ' '), p.sku.startsWith('8K') ? (ST.s07.cat_8k || 'Ultra High Speed HDMI Cable') : (ST.s07.cat_4k || 'Premium High Speed HDMI Cable (placeholder)'), num(p.length_m) + ' m', money(p.us_list, 2), money(p.us_sale_price, 2), money(p.net_air, 2), pct(p.net_air_pct)]);
    table(s, ST.s07.table_head, rows, { x: M, y: 2.35, w: 8.3, colW: [0.9, 2.5, 0.8, 1.0, 1.25, 1.0, 0.85], fontSize: 11, align: [null, null, 'right', 'right', 'right', 'right', 'right'], rowH: 0.45 });
    text(s, ST.s07.table_note, { x: M, y: 5.3, w: 8.3, h: 1.3, fontSize: 11, color: C.text2 });
    card(s, 9.1, 1.2, 3.73, 5.5);
    text(s, ST.s07.ladder_head, { x: 9.35, y: 1.35, w: 3.3, h: 0.45, fontSize: 16, bold: true, color: C.accent1 });
    bullets(s, ST.s07.ladder_items, { x: 9.35, y: 1.9, w: 3.3, h: 4.6, fontSize: 13, gap: 8 });
  }
  // 8 Section Product
  section('s08', 'Product');
  // 9 Line and roadmap
  {
    const s = content(ST.s09.title, 'Product', ST.s09.notes);
    const rows = F.unit_economics.per_sku.map((p) => [p.sku.replace('_', ' '), p.sku.startsWith('8K') ? ST.s09.line_8k : ST.s09.line_4k, num(p.length_m) + ' m', money(p.us_list, 2), money(p.us_sale_price, 2), num(p.h2_cost_krw), p.launch_sku === 'Yes' ? ST.s09.yes : ST.s09.no]);
    table(s, ST.s09.table_head, rows, { x: M, y: 1.3, w: 7.7, colW: [0.85, 3.05, 0.7, 0.8, 0.8, 0.8, 0.7], fontSize: 10.5, align: [null, null, 'right', 'right', 'right', 'right', 'center'], rowH: 0.36 });
    text(s, ST.s09.note, { x: M, y: 5.75, w: 7.7, h: 1.1, fontSize: 11, color: C.text2 });
    card(s, 8.5, 1.3, 4.33, 5.5);
    text(s, ST.s09.roadmap_head, { x: 8.75, y: 1.45, w: 3.8, h: 0.45, fontSize: 16, bold: true, color: C.accent1 });
    ST.s09.roadmap.forEach((r, i) => {
      const y = 2.0 + i * 1.55;
      text(s, r.when, { x: 8.75, y, w: 3.8, h: 0.35, fontSize: 13, bold: true, color: C.accent5 });
      text(s, r.what, { x: 8.75, y: y + 0.38, w: 3.8, h: 1.1, fontSize: 12.5 });
    });
  }
  // 10 Differentiators
  {
    const s = content(ST.s10.title, 'Product', ST.s10.notes);
    const n = 5, gap = GAP, w = (CW - gap * (n - 1)) / n;
    for (let i = 0; i < n; i++) {
      const c = ST.s10.cards[i], x = M + i * (w + gap);
      card(s, x, 1.5, w, 4.0);
      await iconCircle(s, c.icon, x + (w - 0.9) / 2, 1.75, 0.9);
      text(s, c.head, { x: x + 0.15, y: 2.8, w: w - 0.3, h: 0.9, fontSize: 15, bold: true, align: 'center', valign: 'middle' });
      text(s, c.body, { x: x + 0.2, y: 3.75, w: w - 0.4, h: 1.65, fontSize: 14, align: 'left' });
    }
  }
  // 11 Brand architecture
  {
    const s = content(ST.s11.title, 'Product', ST.s11.notes);
    s.addShape(pres.ShapeType.roundRect, { x: M, y: 1.3, w: 5.8, h: 0.7, fill: { color: HEX.head }, line: { color: HEX.cyan, width: 1 }, rectRadius: 0.1 });
    text(s, ST.s11.brand_root, { x: M + 0.2, y: 1.3, w: 5.4, h: 0.7, fontSize: 15, bold: true, color: C.accent1, valign: 'middle', align: 'center' });
    ST.s11.lines.forEach((l, i) => {
      const y = 2.25 + i * 1.55;
      s.addShape(pres.ShapeType.line, { x: M + 0.5, y: y - 0.25, w: 0, h: 0.25, line: { color: HEX.cyan, width: 1 } });
      card(s, M, y, 5.8, 1.35);
      text(s, l.head, { x: M + 0.25, y: y + 0.12, w: 5.3, h: 0.4, fontSize: 15, bold: true });
      text(s, l.body, { x: M + 0.25, y: y + 0.55, w: 5.3, h: 0.75, fontSize: 12.5, color: C.text2 });
    });
    card(s, 6.6, 1.3, 6.23, 2.95);
    text(s, ST.s11.rule_head, { x: 6.85, y: 1.42, w: 5.8, h: 0.4, fontSize: 16, bold: true, color: C.accent5 });
    bullets(s, ST.s11.rule_items, { x: 6.85, y: 1.9, w: 5.75, h: 2.3, fontSize: 12.5, gap: 5 });
    card(s, 6.6, 4.45, 6.23, 2.35);
    text(s, ST.s11.tm_head, { x: 6.85, y: 4.57, w: 5.8, h: 0.4, fontSize: 16, bold: true, color: C.accent1 });
    bullets(s, ST.s11.tm_items, { x: 6.85, y: 5.05, w: 5.75, h: 1.7, fontSize: 12.5, gap: 5 });
  }
  // 12 Section Structure
  section('s12', 'Structure and compliance');
  // 13 US entity and tax
  {
    const s = content(ST.s13.title, 'Structure and compliance', ST.s13.notes);
    const w = (CW - GAP) / 2, h = 2.6;
    ST.s13.cards.forEach((c, i) => {
      const x = M + (i % 2) * (w + GAP), y = 1.3 + Math.floor(i / 2) * (h + GAP);
      card(s, x, y, w, h);
      text(s, c.head, { x: x + 0.25, y: y + 0.12, w: w - 0.5, h: 0.4, fontSize: 15, bold: true, color: C.accent1 });
      bullets(s, c.items, { x: x + 0.25, y: y + 0.58, w: w - 0.5, h: h - 0.7, fontSize: 12, gap: 3 });
    });
  }
  // 14 Korea filings
  {
    const s = content(ST.s14.title, 'Structure and compliance', ST.s14.notes);
    ST.s14.steps.forEach((st, i) => {
      const y = 1.3 + i * 1.38, h = 1.26;
      card(s, M, y, 8.2, h);
      s.addShape(pres.ShapeType.roundRect, { x: M + 0.15, y: y + 0.15, w: 1.7, h: h - 0.3, fill: { color: HEX.head }, line: { color: HEX.head, width: 0 }, rectRadius: 0.08 });
      text(s, st.when, { x: M + 0.2, y: y + 0.15, w: 1.6, h: h - 0.3, fontSize: 11.5, bold: true, color: C.accent1, align: 'center', valign: 'middle' });
      text(s, st.head, { x: M + 2.05, y: y + 0.1, w: 6.0, h: 0.35, fontSize: 14, bold: true });
      text(s, st.body, { x: M + 2.05, y: y + 0.45, w: 6.0, h: h - 0.5, fontSize: 11.5, color: C.text2 });
    });
    card(s, 9.0, 1.3, 3.83, 5.4, { fill: HEX.card2 });
    text(s, ST.s14.side_head, { x: 9.2, y: 1.45, w: 3.4, h: 0.4, fontSize: 16, bold: true, color: C.accent1 });
    ST.s14.side_items.forEach((it, i) => text(s, it, { x: 9.2, y: 2.0 + i * 1.1, w: 3.4, h: 1.05, fontSize: 13 }));
  }
  // 15 Product compliance
  {
    const s = content(ST.s15.title, 'Structure and compliance', ST.s15.notes);
    const w = (CW - 2 * GAP) / 3, h = 2.6;
    for (let i = 0; i < 6; i++) {
      const c = ST.s15.cards[i], x = M + (i % 3) * (w + GAP), y = 1.3 + Math.floor(i / 3) * (h + GAP);
      card(s, x, y, w, h);
      await iconCircle(s, c.icon, x + 0.2, y + 0.18, 0.55);
      text(s, c.head, { x: x + 0.9, y: y + 0.18, w: w - 1.1, h: 0.55, fontSize: 15, bold: true, valign: 'middle' });
      text(s, c.body, { x: x + 0.2, y: y + 0.9, w: w - 0.4, h: 1.0, fontSize: 12.5 });
      text(s, c.status, { x: x + 0.2, y: y + h - 0.7, w: w - 0.4, h: 0.6, fontSize: 11, color: C.accent5 });
    }
  }
  // 16 Tariff stack
  {
    const s = content(ST.s16.title, 'Structure and compliance', ST.s16.notes);
    const ts = F.duty.tariff_stack;
    const rows = ST.s16.rows.map((r, i) => {
      const d = i < 3 ? ts[i] : { mfn: K.duty.fiber_list3_check.mfn, section_301_list: K.duty.fiber_list3_check.section_301_list, section_301_forced_labor: K.duty.fiber_list3_check.section_301_forced_labor, pending_second_action: 0, total: K.duty.fiber_list3_check.total };
      return [r.name, r.hts, pctAuto(d.mfn), pctAuto(d.section_301_list), pct(d.section_301_forced_labor), pctAuto(d.pending_second_action), { text: pct(d.total), options: { bold: true, color: i === 0 ? HEX.cyan : i === 1 ? HEX.green : i === 2 ? HEX.amber : HEX.muted } }];
    });
    table(s, ST.s16.table_head, rows, { x: M, y: 1.3, w: 7.35, colW: [1.45, 1.75, 0.65, 1.0, 1.05, 0.75, 0.7], fontSize: 10.5, align: [null, null, 'right', 'right', 'right', 'right', 'right'], rowH: 0.42 });
    card(s, M, 3.75, 7.35, 3.05);
    text(s, ST.s16.history_head, { x: M + 0.25, y: 3.85, w: 6.8, h: 0.4, fontSize: 15, bold: true, color: C.accent1 });
    bullets(s, ST.s16.history, { x: M + 0.25, y: 4.3, w: 6.9, h: 2.45, fontSize: 11.5, gap: 3 });
    // waterfall chart
    const wf = F.duty.waterfall;
    const cats = ST.s16.chart_cats;
    const base = wf.map((w) => w.base);
    const netv = wf.map((w, i) => (i % 2 === 0 ? w.value : 0));
    const up = wf.map((w, i) => (i === 1 ? w.value : 0));
    const down = wf.map((w, i) => (i === 3 ? w.value : 0));
    text(s, ST.s16.chart_title, { x: 8.1, y: 1.3, w: 4.73, h: 0.35, fontSize: 13, bold: true, color: C.text1 });
    s.addChart(pres.ChartType.bar, [
      { name: 'base', labels: cats, values: base },
      { name: 'net', labels: cats, values: netv },
      { name: 'uplift', labels: cats, values: up },
      { name: 'penalty', labels: cats, values: down },
    ], Object.assign(CHART_BASE(), { x: 8.1, y: 1.65, w: 4.73, h: 3.45, barDir: 'col', barGrouping: 'stacked', barGapWidthPct: 45, chartColors: ['transparent', HEX.cyan, HEX.green, HEX.amber], showLegend: false, showValue: true, dataLabelPosition: 'ctr', dataLabelFormatCode: '$0.00;($0.00);', dataLabelFontSize: 10, dataLabelColor: HEX.bg, dataLabelFontBold: true, valAxisLabelFormatCode: '$0', valAxisMinVal: 0, valAxisMaxVal: 12, valAxisMajorUnit: 2, catAxisLabelFontSize: 9 }));
    card(s, 8.1, 5.3, 4.73, 1.5);
    text(s, ST.s16.effect_head, { x: 8.3, y: 5.38, w: 4.4, h: 0.35, fontSize: 13, bold: true, color: C.accent4 });
    text(s, ST.s16.effect, { x: 8.3, y: 5.72, w: 4.4, h: 1.05, fontSize: 11.5 });
  }
  // 17 Section Launch
  section('s17', 'Launch');
  // 18 Critical path
  {
    const s = content(ST.s18.title, 'Launch', ST.s18.notes);
    const g = flowImage(s, '01', M, 1.2, CW, 4.1, flowFile('01_launch_critical_path_deck.png'));
    text(s, ST.s18.caption, { x: M, y: g.bottom + 0.05, w: CW, h: 0.3, fontSize: 11, color: C.text2 }, { file: '01_launch_critical_path.png', full_size: T(ST.common.full_size, { file: '01_launch_critical_path.png' }) });
    const gy = g.bottom + 0.4, gh = 6.9 - gy, gw = (CW - 2 * GAP) / 3;
    text(s, ST.s18.gates_head, { x: M, y: gy, w: CW, h: 0.3, fontSize: 12.5, bold: true, color: C.accent1 });
    ST.s18.gates.forEach((t, i) => {
      const x = M + i * (gw + GAP);
      card(s, x, gy + 0.35, gw, gh - 0.35, { fill: HEX.card2 });
      text(s, String(i + 1), { x: x + 0.2, y: gy + 0.42, w: 0.5, h: 0.6, fontSize: 24, bold: true, color: C.accent1 });
      text(s, t, { x: x + 0.7, y: gy + 0.45, w: gw - 0.9, h: gh - 0.55, fontSize: 11.5 });
    });
  }
  // 19 Day 90 gate
  {
    const s = content(ST.s19.title, 'Launch', ST.s19.notes);
    text(s, ST.s19.d90_head, { x: M, y: 1.2, w: 7.8, h: 0.3, fontSize: 13, bold: true, color: C.accent1 });
    table(s, ST.s19.table_head, ST.s19.d90_rows.map((r) => [r.metric, r.pass, r.kill, r.model]), { x: M, y: 1.5, w: 7.8, colW: [3.5, 1.45, 1.45, 1.4], fontSize: 10, rowH: 0.33 });
    text(s, ST.s19.m12_head, { x: M, y: 4.3, w: 7.8, h: 0.3, fontSize: 13, bold: true, color: C.accent1 });
    table(s, ST.s19.table_head, ST.s19.m12_rows.map((r) => [r.metric, r.pass, r.kill, r.model]), { x: M, y: 4.6, w: 7.8, colW: [2.1, 1.7, 2.6, 1.4], fontSize: 10, rowH: 0.33 });
    card(s, 8.6, 1.22, 4.23, 2.25);
    text(s, ST.s19.results_head, { x: 8.85, y: 1.32, w: 3.8, h: 0.35, fontSize: 13, bold: true, color: C.accent1 });
    const colorOf = (r) => (r === 'PASS' ? HEX.green : r === 'KILL' ? HEX.magenta : HEX.amber);
    ST.s19.results.forEach((r, i) => {
      const y = 1.75 + i * 0.52;
      text(s, r.scenario, { x: 8.85, y, w: 1.2, h: 0.42, fontSize: 12.5, bold: true, valign: 'middle' });
      [[r.d90, ST.s19.chip_d90 || 'Day 90'], [r.m12, ST.s19.chip_m12 || 'Month 12']].forEach(([val, lab], j) => {
        const v = T(val); // PASS, REVIEW or KILL from the model; the KO strings map the word shown, the color keys on the model value
        const shown = (ST.s19.result_labels && ST.s19.result_labels[v]) || v;
        const x = 10.05 + j * 1.4;
        s.addShape(pres.ShapeType.roundRect, { x, y, w: 1.3, h: 0.42, fill: { color: colorOf(v) }, line: { color: colorOf(v), width: 0 }, rectRadius: 0.08 });
        text(s, lab + ': ' + shown, { x, y, w: 1.3, h: 0.42, fontSize: 9.5, bold: true, color: HEX.bg, align: 'center', valign: 'middle' });
      });
    });
    card(s, 8.6, 3.65, 4.23, 3.15, { fill: HEX.card2 });
    await iconCircle(s, 'FiAlertOctagon', 8.85, 3.85, 0.55, HEX.magenta);
    text(s, ST.s19.rule, { x: 8.85, y: 4.5, w: 3.8, h: 2.2, fontSize: 12 });
  }
  // 20 Section Economics
  section('s20', 'Economics');
  // 21 Hero unit economics
  {
    const s = content(ST.s21.title, 'Economics', ST.s21.notes);
    const h = F.unit_economics.hero;
    const costs = [h.landed_cost, h.referral, h.fba, h.ads, h.returns_storage];
    const base = [0]; let run = h.price; const costv = [0]; const price = [h.price, 0, 0, 0, 0, 0, 0]; const net = [0, 0, 0, 0, 0, 0, h.net];
    costs.forEach((c) => { run -= c; base.push(run); costv.push(c); });
    base.push(0); costv.push(0);
    s.addChart(pres.ChartType.bar, [
      { name: 'base', labels: ST.s21.chart_cats, values: base },
      { name: 'price', labels: ST.s21.chart_cats, values: price },
      { name: 'cost', labels: ST.s21.chart_cats, values: costv },
      { name: 'net', labels: ST.s21.chart_cats, values: net },
    ], Object.assign(CHART_BASE(), { x: M, y: 1.3, w: 8.0, h: 4.55, barDir: 'col', barGrouping: 'stacked', barGapWidthPct: 40, chartColors: ['transparent', HEX.cyan, HEX.amber, HEX.green], showLegend: false, showValue: true, dataLabelPosition: 'ctr', dataLabelFormatCode: '$0.00;($0.00);', dataLabelFontSize: 10.5, dataLabelColor: HEX.bg, dataLabelFontBold: true, valAxisLabelFormatCode: '$0', valAxisMinVal: 0, valAxisMaxVal: 50, valAxisMajorUnit: 10, catAxisLabelFontSize: 10 }));
    text(s, ST.s21.note, { x: M, y: 5.95, w: 8.0, h: 0.95, fontSize: 11, color: C.text2 });
    ST.s21.side.forEach((c, i) => {
      const y = 1.3 + i * 1.6;
      card(s, 8.8, y, 4.03, 1.45);
      text(s, c.label, { x: 9.0, y: y + 0.1, w: 3.6, h: 0.3, fontSize: 12, color: C.text2 });
      text(s, c.value, { x: 9.0, y: y + 0.4, w: 3.6, h: 0.6, fontSize: 28, bold: true, color: i === 1 ? C.accent5 : C.accent4, valign: 'middle' });
      text(s, c.sub, { x: 9.0, y: y + 1.0, w: 3.6, h: 0.42, fontSize: 11, color: C.text2 });
    });
    card(s, 8.8, 6.1, 4.03, 0.75, { fill: HEX.card2 });
    text(s, ST.s21.compare || '{hero_net} vs {hero_korea}: Amazon US vs Korea wholesale, same unit', { x: 9.0, y: 6.1, w: 3.7, h: 0.75, fontSize: 11.5, bold: true, color: C.accent1, valign: 'middle' });
  }
  // 22 Factory cost cross check
  {
    const s = content(ST.s22.title, 'Economics', ST.s22.notes);
    table(s, ST.s22.table_head, ST.s22.rows.map((r) => [r.name, { text: r.landed, options: { align: 'right' } }, { text: r.net, options: { align: 'right', bold: true } }, r.basis]), { x: M, y: 1.3, w: 7.9, colW: [3.1, 1.3, 1.3, 2.2], fontSize: 11, rowH: 0.56 });
    text(s, ST.s22.table_note, { x: M, y: 4.12, w: 7.9, h: 0.4, fontSize: 10, color: C.text2 });
    card(s, M, 4.55, 7.9, 1.45);
    text(s, ST.s22.callout_head, { x: M + 0.25, y: 4.6, w: 7.4, h: 0.32, fontSize: 14, bold: true, color: C.accent3 });
    text(s, ST.s22.callout, { x: M + 0.25, y: 4.92, w: 7.4, h: 1.05, fontSize: 12 });
    text(s, ST.s22.action, { x: M, y: 6.05, w: 7.9, h: 0.85, fontSize: 12, bold: true, color: C.accent5 });
    text(s, ST.s22.chart_title, { x: 8.65, y: 1.3, w: 4.18, h: 0.35, fontSize: 13, bold: true });
    const cb = [CB.h2.net, CB.august.net, CB.alibaba_low.net, CB.alibaba_high.net];
    s.addChart(pres.ChartType.bar, [{ name: 'net', labels: ST.s22.chart_cats, values: cb }], Object.assign(CHART_BASE(), { x: 8.65, y: 1.65, w: 4.18, h: 4.3, barDir: 'col', barGapWidthPct: 50, chartColors: [HEX.green, HEX.cyan, HEX.magenta, HEX.magenta], showLegend: false, showValue: true, dataLabelPosition: 'outEnd', dataLabelFormatCode: '$0.00;($0.00)', dataLabelFontSize: 10.5, valAxisLabelFormatCode: '$0;($0)', valAxisMinVal: -10, valAxisMaxVal: 15, valAxisMajorUnit: 5, catAxisLabelFontSize: 10, catAxisLabelPos: 'low' }));
  }
  // 23 Eighteen month P&L
  {
    const s = content(ST.s23.title, 'Economics', ST.s23.notes);
    const mon = (sc, key) => F.scenarios[sc].monthly.slice(0, 18).map((m) => m[key]);
    const labels = F.scenarios.Medium.monthly.slice(0, 18).map((m) => String(m.month));
    const cc = ST.common;
    s.addChart([
      { type: pres.ChartType.line, data: [{ name: cc.medium_revenue, labels, values: mon('Medium', 'amazon_revenue') }, { name: cc.medium_net, labels, values: mon('Medium', 'amazon_net') }], options: { chartColors: [HEX.cyan, HEX.green], lineSize: 3, lineDataSymbol: 'none' } },
      { type: pres.ChartType.line, data: [{ name: cc.low_revenue, labels, values: mon('Low', 'amazon_revenue') }, { name: cc.high_revenue, labels, values: mon('High', 'amazon_revenue') }], options: { chartColors: [HEX.amber, HEX.violet], lineSize: 1.25, lineDataSymbol: 'none' } },
      { type: pres.ChartType.line, data: [{ name: cc.low_net, labels, values: mon('Low', 'amazon_net') }, { name: cc.high_net, labels, values: mon('High', 'amazon_net') }], options: { chartColors: [HEX.amber, HEX.violet], lineSize: 1.25, lineDash: 'dash', lineDataSymbol: 'none' } },
    ], Object.assign(CHART_BASE(), { x: M, y: 1.3, w: 7.95, h: 4.5, valAxisLabelFormatCode: NEG_FMT_K, showCatAxisTitle: true, catAxisTitle: cc.chart_month_axis, catAxisTitleColor: HEX.muted, catAxisTitleFontSize: 10, showLegend: true, legendPos: 'b', valAxisMinVal: -20000, valAxisMaxVal: 240000, valAxisMajorUnit: 40000 }));
    text(s, ST.s23.note, { x: M, y: 5.9, w: CW, h: 1.0, fontSize: 11, color: C.text2 });
    table(s, ST.s23.table_head, ST.s23.rows.map((r) => [r.label, r.L, { text: r.M, options: { bold: true, color: HEX.cyan } }, r.H]), { x: 8.7, y: 1.3, w: 4.13, colW: [1.63, 0.8, 0.9, 0.8], fontSize: 10.5, align: [null, 'right', 'right', 'right'], rowH: 0.45 });
  }
  // 24 Five year trajectory
  {
    const s = content(ST.s24.title, 'Economics', ST.s24.notes);
    const years = [1, 2, 3, 4, 5];
    const labels = years.map((y) => T(ST.common.year, { n: String(y) }));
    const chs = ['amazon_us', 'ebay_walmart', 'b2b', 'dtc', 'amazon_canada', 'pro_line'];
    const data = chs.map((ch, i) => ({ name: ST.s24.chart_series[i], labels, values: years.map((y) => Math.round(F.scenarios.Medium.annual['Y' + y].revenue_by_channel[ch])) }));
    s.addChart(pres.ChartType.bar, data, Object.assign(CHART_BASE(), { x: M, y: 1.3, w: 7.3, h: 5.4, barDir: 'col', barGrouping: 'stacked', barGapWidthPct: 55, chartColors: [HEX.cyan, HEX.violet, HEX.green, HEX.amber, HEX.magenta, HEX.blue], showLegend: true, legendPos: 'b', showValue: false, valAxisLabelFormatCode: '$#,##0.0,,"M"', valAxisMinVal: 0, valAxisMaxVal: 5000000, valAxisMajorUnit: 1000000 }));
    const rows = ST.s24.rows.map((r) => [{ text: r.label, options: { bold: true } }, ...years.map((y) => V[`y${y}_rev_${r.key}`]), { text: V[`y5_net_${r.key}`] + ', ' + V[`y5_margin_${r.key}`], options: { color: r.key === 'L' ? HEX.amber : HEX.green } }]);
    for (const r of ST.s24.rows) for (const y of years) { USED_KEYS.add(`y${y}_rev_${r.key}`); (SLIDE_KEYS[CUR] = SLIDE_KEYS[CUR] || new Set()).add(`y${y}_rev_${r.key}`); }
    table(s, ST.s24.table_head, rows, { x: 8.05, y: 1.3, w: 4.78, colW: [0.78, 0.6, 0.62, 0.62, 0.62, 0.62, 0.92], fontSize: 9.5, align: [null, 'right', 'right', 'right', 'right', 'right', 'right'], rowH: 0.45 });
    card(s, 8.05, 3.35, 4.78, 3.35, { fill: HEX.card2 });
    text(s, ST.s24.note, { x: 8.3, y: 3.5, w: 4.3, h: 3.1, fontSize: 14 });
  }
  // 25 What must be true
  {
    const s = content(ST.s25.title, 'Economics', ST.s25.notes);
    table(s, ST.s25.col_head, ST.s25.rows.map((r) => [{ text: r.driver, options: { bold: true } }, r.L, { text: r.M, options: { color: HEX.cyan } }, r.H]), { x: M, y: 1.3, w: CW, colW: [2.9, 3.0, 3.3, 3.13], fontSize: 10.5, rowH: 0.46 });
    text(s, ST.s25.note, { x: M, y: 6.15, w: CW, h: 0.7, fontSize: 11, color: C.text2 });
  }
  // 26 Channel expansion roadmap
  {
    const s = content(ST.s26.title, 'Economics', ST.s26.notes);
    const g = flowImage(s, '07', M, 1.2, CW, 4.55, flowFile('07_five_year_channel_map_deck.png'));
    text(s, ST.s26.caption, { x: M, y: g.bottom + 0.05, w: CW, h: 0.45, fontSize: 11, color: C.text2 }, { file: '07_five_year_channel_map.png', full_size: T(ST.common.full_size, { file: '07_five_year_channel_map.png' }) });
    const cy = 6.3, n = ST.s26.rows.length, cw = (CW - 2.0 - GAP * (n - 1)) / n;
    text(s, ST.s26.chips_head, { x: M, y: cy, w: 2.0, h: 0.6, fontSize: 12.5, bold: true, color: C.accent1, valign: 'middle' });
    ST.s26.rows.forEach((r, i) => {
      const x = M + 2.0 + i * (cw + GAP);
      card(s, x, cy, cw, 0.6, { fill: HEX.card2 });
      const yw = LANG === 'ko' ? 0.65 : 0.9; // "1년차" is shorter than "Year 1"; the value needs the room (Noto digits are wider)
      text(s, r.year, { x: x + 0.15, y: cy, w: yw, h: 0.6, fontSize: 12, color: C.text2, valign: 'middle' });
      text(s, r.rev, { x: x + yw, y: cy, w: cw - yw - 0.15, h: 0.6, fontSize: 16, bold: true, color: C.accent4, valign: 'middle', align: 'right' });
      T(r.what);
    });
  }
  // 27 Cash by month
  {
    const s = content(ST.s27.title, 'Economics', ST.s27.notes);
    const labels = ['0', ...F.scenarios.Medium.monthly.map((m) => String(m.month)), ''];
    const series = (sc) => [F.scenarios[sc].cash_balance_month0, ...F.scenarios[sc].monthly.map((m) => m.cash_balance)].map((v) => Math.round(v)).concat([null]);
    const mins = labels.map(() => null);
    for (const sc of ['Low', 'Medium', 'High']) { const c = F.scenarios[sc].cash; if (c.peak_need_month <= 24) mins[c.peak_need_month] = Math.round(-c.peak_need); }
    const names = ST.s27.chart_series;
    s.addChart([
      { type: pres.ChartType.line, data: [{ name: names[1], labels, values: series('Medium') }], options: { chartColors: [HEX.cyan], lineSize: 3, lineDataSymbol: 'none' } },
      { type: pres.ChartType.line, data: [{ name: names[0], labels, values: series('Low') }, { name: names[2], labels, values: series('High') }], options: { chartColors: [HEX.amber, HEX.green], lineSize: 1.25, lineDataSymbol: 'none' } },
      { type: pres.ChartType.line, data: [{ name: names[3], labels, values: mins }], options: { chartColors: [HEX.magenta], lineSize: 1, lineDataSymbol: 'circle', lineDataSymbolSize: 10, showValue: true, dataLabelPosition: 'b', dataLabelFormatCode: NEG_FMT, dataLabelColor: HEX.magenta, dataLabelFontSize: 10, dataLabelFontBold: true } },
    ], Object.assign(CHART_BASE(), { x: M, y: 1.3, w: 8.0, h: 5.5, displayBlanksAs: 'gap', valAxisLabelFormatCode: NEG_FMT_K, showCatAxisTitle: true, catAxisTitle: ST.common.chart_month_axis, catAxisTitleColor: HEX.muted, catAxisTitleFontSize: 10, showLegend: true, legendPos: 'b', valAxisMinVal: -150000, valAxisMaxVal: 600000, valAxisMajorUnit: 150000 }));
    ST.s27.cards.forEach((c, i) => {
      const y = 1.3 + i * 1.2;
      card(s, 8.8, y, 4.03, 1.08);
      text(s, c.label, { x: 9.0, y: y + 0.08, w: 3.6, h: 0.28, fontSize: 11, color: C.text2 });
      text(s, c.value, { x: 9.0, y: y + 0.33, w: 1.6, h: 0.45, fontSize: 22, bold: true, color: i === 0 ? C.accent1 : i === 1 ? C.accent5 : C.accent4, valign: 'middle' });
      text(s, c.sub, { x: 10.6, y: y + 0.3, w: 2.1, h: 0.72, fontSize: 10, color: C.text2, valign: 'middle' });
    });
    card(s, 8.8, 4.95, 4.03, 1.85, { fill: HEX.card2 });
    text(s, ST.s27.assump_head, { x: 9.0, y: 5.0, w: 3.6, h: 0.3, fontSize: 12, bold: true, color: C.accent1 });
    bullets(s, ST.s27.assump, { x: 9.0, y: 5.3, w: 3.65, h: 1.5, fontSize: 9.5, gap: 1 });
  }
  // 28 Funding plan
  {
    const s = content(ST.s28.title, 'Economics', ST.s28.notes);
    table(s, ST.s28.table_head, ST.s28.rows.map((r) => [r.label, r.L, { text: r.M, options: { bold: true, color: HEX.cyan } }, r.H]), { x: M, y: 1.3, w: 7.9, colW: [3.5, 1.4, 1.5, 1.5], fontSize: 11.5, align: [null, 'right', 'right', 'right'], rowH: 0.5 });
    card(s, 8.7, 1.3, 4.13, 5.5);
    text(s, ST.s28.order_head, { x: 8.95, y: 1.45, w: 3.7, h: 0.4, fontSize: 16, bold: true, color: C.accent1 });
    ST.s28.order.forEach((it, i) => text(s, it, { x: 8.95, y: 2.0 + i * 1.15, w: 3.65, h: 1.1, fontSize: 12.5 }));
  }
  // 29 Sensitivity tornado
  {
    const s = content(ST.s29.title, 'Economics', ST.s29.notes);
    const rows = [...tornado].reverse();
    const labels = rows.map((t) => ST.s29.drivers[t.driver] || t.driver);
    s.addChart(pres.ChartType.bar, [
      { name: ST.s29.chart_series[0], labels, values: rows.map((t) => Math.round(t.low_minus_base)) },
      { name: ST.s29.chart_series[1], labels, values: rows.map((t) => Math.round(t.high_minus_base)) },
    ], Object.assign(CHART_BASE(), { x: M, y: 1.3, w: 6.5, h: 4.9, barDir: 'bar', barGrouping: 'stacked', barGapWidthPct: 40, chartColors: [HEX.magenta, HEX.green], showLegend: true, legendPos: 'b', showValue: false, valAxisLabelFormatCode: NEG_FMT_K, valAxisMinVal: -400000, valAxisMaxVal: 300000, valAxisMajorUnit: 100000, catAxisLabelFontSize: 10, catAxisLabelPos: 'low', valGridLine: { color: HEX.grid, size: 0.5 } }));
    text(s, ST.s29.base_label, { x: M, y: 6.25, w: 6.5, h: 0.3, fontSize: 11.5, bold: true, color: C.accent1 });
    text(s, ST.s29.note, { x: M, y: 6.55, w: 6.5, h: 0.45, fontSize: 10, color: C.text2 });
    const trows = tornado.map((t) => [ST.s29.drivers[t.driver] || t.driver, ST.s29.labels[t.low_label] || t.low_label, { text: money_k(t.low_value), options: { align: 'right', color: t.low_value < 0 ? HEX.magenta : HEX.text } }, ST.s29.labels[t.high_label] || t.high_label, { text: money_k(t.high_value), options: { align: 'right', color: HEX.green } }]);
    table(s, ST.s29.table_head, trows, { x: 7.3, y: 1.3, w: 5.53, colW: [1.55, 1.2, 0.8, 1.2, 0.78], fontSize: 10, rowH: 0.52 });
  }
  // 30 Twelve levers
  {
    const s = content(ST.s30.title, 'Economics', ST.s30.notes);
    const rows = F.levers.map((l, i) => {
      const st = ST.s30.levers[i]; const ft = ST.s30.from_to[String(l.n)];
      const eff = typeof l.model_effect === 'number' ? money_k(l.model_effect) : (ST.s30.effect_notes[String(l.n)] || String(l.model_effect));
      return [String(l.n), st.lever, st.moves, ft[0], ft[1], st.owner, { text: eff, options: { align: 'right', color: typeof l.model_effect === 'number' ? HEX.green : HEX.muted } }, st.basis];
    });
    table(s, ST.s30.table_head, rows, { x: M, y: 1.25, w: CW, colW: [0.35, 3.1, 2.1, 1.3, 2.0, 1.15, 1.2, 1.13], fontSize: 9.5, rowH: 0.4 });
  }
  // 31 Marketing playbook
  {
    const s = content(ST.s31.title, 'Economics', ST.s31.notes);
    const w = (CW - 2 * GAP) / 3, h = 2.6;
    for (let i = 0; i < 6; i++) {
      const c = ST.s31.cards[i], x = M + (i % 3) * (w + GAP), y = 1.3 + Math.floor(i / 3) * (h + GAP);
      card(s, x, y, w, h);
      await iconCircle(s, c.icon, x + 0.2, y + 0.18, 0.55);
      text(s, c.head, { x: x + 0.9, y: y + 0.18, w: w - 1.1, h: 0.55, fontSize: 15, bold: true, valign: 'middle' });
      text(s, c.body, { x: x + 0.2, y: y + 0.9, w: w - 0.4, h: h - 1.0, fontSize: 14 });
    }
  }
  // 32 Operating model and team
  {
    const s = content(ST.s32.title, 'Economics', ST.s32.notes);
    const w = 3.9, h = 2.55;
    for (let i = 0; i < 4; i++) {
      const p = ST.s32.people[i], x = M + (i % 2) * (w + GAP), y = 1.3 + Math.floor(i / 2) * (h + GAP);
      card(s, x, y, w, h);
      await iconCircle(s, p.icon, x + 0.2, y + 0.18, 0.55);
      text(s, p.head, { x: x + 0.9, y: y + 0.18, w: w - 1.1, h: 0.55, fontSize: 14, bold: true, valign: 'middle' });
      text(s, p.body, { x: x + 0.2, y: y + 0.9, w: w - 0.4, h: h - 1.0, fontSize: 12 });
    }
    const x = M + 2 * (w + GAP), tw = CW - 2 * (w + GAP);
    text(s, ST.s32.cost_head, { x, y: 1.3, w: tw, h: 0.35, fontSize: 12.5, bold: true, color: C.accent1 });
    table(s, null, ST.s32.cost_rows.map((r) => [{ text: r.year, options: { bold: true } }, { text: r.amount, options: { bold: true, color: HEX.green, align: 'right' } }, r.what]), { x, y: 1.7, w: tw, colW: [0.7, 0.7, tw - 1.4], fontSize: 9.5, rowH: 0.98 });
  }
  // 33 Risk register summary
  {
    const s = content(ST.s33.title, 'Economics', ST.s33.notes);
    table(s, ST.s33.table_head, ST.s33.rows.map((r) => [{ text: r.risk, options: { bold: true } }, r.signal, r.mitigation, r.owner]), { x: M, y: 1.25, w: CW, colW: [3.3, 2.9, 4.83, 1.3], fontSize: 10.5, rowH: 0.56 });
    text(s, ST.s33.note, { x: M, y: 6.45, w: CW, h: 0.45, fontSize: 10.5, color: C.text2 });
  }
  // 34 Kill criteria and Plan B
  {
    const s = content(ST.s34.title, 'Economics', ST.s34.notes);
    const g = flowImage(s, '06', M, 1.2, CW, 3.55, flowFile('06_day90_scale_or_kill_deck.png'));
    text(s, ST.s34.caption, { x: M, y: g.bottom + 0.05, w: CW, h: 0.3, fontSize: 11, color: C.text2 }, { file: '06_day90_scale_or_kill.png', full_size: T(ST.common.full_size, { file: '06_day90_scale_or_kill.png' }) });
    const ry = g.bottom + 0.42, tw = 7.4, kx = M + tw + GAP, kw = CW - tw - GAP;
    const head = [ST.s34.planb_head, ...ST.s34.planb_table_head.slice(1)];
    table(s, head, ST.s34.planb_rows.map((r) => [r.label, ...r.v.map((v) => ({ text: v, options: { align: 'right' } }))]), { x: M, y: ry, w: tw, colW: [tw - 2.7, 0.9, 0.9, 0.9], fontSize: 10.5, rowH: 0.34, align: [null, 'right', 'right', 'right'] });
    card(s, kx, ry, kw, 6.9 - ry, { fill: HEX.card2 });
    text(s, ST.s34.kill_head, { x: kx + 0.2, y: ry + 0.08, w: kw - 0.4, h: 0.32, fontSize: 12.5, bold: true, color: C.accent3 });
    text(s, ST.s34.kill_body, { x: kx + 0.2, y: ry + 0.42, w: kw - 0.4, h: 0.85, fontSize: 11.5 });
    text(s, ST.s34.planb_note, { x: kx + 0.2, y: ry + 1.25, w: kw - 0.4, h: 6.9 - ry - 1.3, fontSize: 10, color: C.text2 });
  }
  // 35 The ask
  {
    const s = content(ST.s35.title, 'Economics', ST.s35.notes);
    const gap = GAP, w = (CW - 3 * gap) / 4;
    ST.s35.cards.forEach((c, i) => {
      const x = M + i * (w + gap);
      card(s, x, 1.3, w, 3.2);
      text(s, c.n, { x: x + 0.2, y: 1.4, w: 0.7, h: 0.7, fontSize: 34, bold: true, color: C.accent1 });
      text(s, c.head, { x: x + 0.9, y: 1.5, w: w - 1.1, h: 0.5, fontSize: 17, bold: true, valign: 'middle' });
      text(s, c.body, { x: x + 0.2, y: 2.2, w: w - 0.4, h: 2.2, fontSize: 14 });
    });
    card(s, M, 4.75, CW, 2.05, { fill: HEX.card2 });
    text(s, ST.s35.upside_head, { x: M + 0.3, y: 4.9, w: 6, h: 0.4, fontSize: 15, bold: true, color: C.accent4 });
    text(s, ST.s35.upside, { x: M + 0.3, y: 5.35, w: CW - 0.6, h: 1.3, fontSize: 15 });
  }
  // 36 First 30 days
  {
    const s = content(ST.s36.title, 'Economics', ST.s36.notes);
    const gap = GAP, w = (CW - 3 * gap) / 4;
    ST.s36.weeks.forEach((wk, i) => {
      const x = M + i * (w + gap);
      card(s, x, 1.3, w, 3.55);
      text(s, wk.head, { x: x + 0.2, y: 1.4, w: w - 0.4, h: 0.45, fontSize: 17, bold: true, color: C.accent1 });
      bullets(s, wk.items, { x: x + 0.2, y: 1.9, w: w - 0.4, h: 2.9, fontSize: 12, gap: 4 });
    });
    card(s, M, 5.05, CW, 1.8, { fill: HEX.card2 });
    text(s, ST.s36.monday_head, { x: M + 0.3, y: 5.15, w: 6, h: 0.35, fontSize: 15, bold: true, color: C.accent4 });
    ST.s36.monday.forEach((m, i) => text(s, m, { x: M + 0.3 + i * (CW / 3), y: 5.6, w: CW / 3 - 0.4, h: 1.1, fontSize: 13.5 }));
  }
  // 37 Glossary
  {
    const s = content(ST.s37.title, 'Economics', ST.s37.notes);
    const terms = ST.s37.terms, w = (CW - 2 * GAP) / 3;
    // split into three columns with about equal rendered lines (a term wraps at about 50 characters at 12 pt; Hangul glyphs are wider, about 30)
    // KO: a Hangul glyph is about two Latin glyphs wide (EN keeps the plain count so its column split is unchanged)
    const weight = (str) => [...str].reduce((a, ch) => a + (LANG === 'ko' && HANGUL.test(ch) ? 2 : 1), 0);
    const lines = terms.map(([t, d]) => Math.ceil((weight(t) + weight(d) + 2) / 50));
    const total = lines.reduce((a, b) => a + b, 0); const cols = [[], [], []]; let acc = 0, ci = 0;
    const slack = LANG === 'ko' ? 0.5 : 0; // KO moves to the next column once the current one is within half an entry of its share
    terms.forEach((tm, i) => { if (ci < 2 && acc + lines[i] * slack >= (total / 3) * (ci + 1)) ci++; cols[ci].push(tm); acc += lines[i]; });
    for (let c = 0; c < 3; c++) {
      const x = M + c * (w + GAP);
      card(s, x, 1.25, w, 5.6);
      const parts = [];
      cols[c].forEach(([t, d], i, arr) => {
        const ko = HANGUL.test(t); // a full width colon inside the Noto run avoids LibreOffice's Asian text autospacing gap before a Latin colon
        parts.push({ t: t + (ko ? '\uFF1A' : ': '), b: true, c: C.accent1, size: 12, font: ko ? 'Noto Sans KR' : undefined });
        parts.push({ t: d, size: 12, br: i < arr.length - 1, gap: 4 });
      });
      runs(s, parts, { x: x + 0.2, y: 1.4, w: w - 0.4, h: 5.3, fontSize: 12 });
    }
  }

  if (slideNo !== 37) throw new Error('Expected 37 slides, built ' + slideNo);
  await pres.writeFile({ fileName: OUT_FILE });
  await applyTheme(OUT_FILE, THEME);
  await postProcess(OUT_FILE);
  const CHART_BLOCKS = { 7: ['unit_economics.per_sku'], 9: ['unit_economics.per_sku'], 16: ['duty.tariff_stack', 'duty.waterfall'], 21: ['unit_economics.hero'], 22: ['unit_economics.cost_bases'], 23: ['scenarios.{Low,Medium,High}.monthly[1..18].amazon_revenue and amazon_net'], 24: ['scenarios.Medium.annual.Y1..Y5.revenue_by_channel'], 27: ['scenarios.{Low,Medium,High}.cash_balance_month0, monthly[].cash_balance, cash.peak_need, cash.peak_need_month'], 29: ['sensitivity.tornado'], 30: ['levers'] };
  const perSlide = {};
  for (let i = 1; i <= slideNo; i++) perSlide[i] = { tokens: [...(SLIDE_KEYS[i] || [])].sort(), figure_blocks: CHART_BLOCKS[i] || [] };
  fs.writeFileSync(path.join(BUILD, 'deck_tokens_used_' + LANG + '.json'), JSON.stringify({ tokens: [...USED_KEYS].sort(), per_slide: perSlide }, null, 1));
  console.log('wrote', OUT_FILE, 'slides', slideNo, 'tokens', USED_KEYS.size);
})().catch((e) => { console.error(e); process.exit(1); });

// Post processing: hide data labels on the invisible "base" series of the waterfall charts and drop blank points so the
// marker series shows only its real points.
// For the KO deck it also tags every run ko-KR (word level line breaking for Hangul in PowerPoint and LibreOffice) and
// writes Noto Sans KR into the East Asian font slots, which pptxgenjs leaves empty.
async function postProcess(file) {
  const JSZip = require(require.resolve('jszip', { paths: [require.resolve('pptxgenjs')] }));
  const zip = await JSZip.loadAsync(fs.readFileSync(file));
  for (const name of Object.keys(zip.files)) {
    if (/^ppt\/charts\/chart\d+\.xml$/.test(name)) {
      let xml = await zip.file(name).async('string');
      xml = xml.replace(/<c:ser>([\s\S]*?)<\/c:ser>/g, (m, inner) => {
        if (/<c:tx>[\s\S]*?<c:v>base<\/c:v>/.test(inner)) inner = inner.replace(/<c:dLbls>[\s\S]*?<\/c:dLbls>/, '<c:dLbls><c:delete val="1"/></c:dLbls>');
        inner = inner.replace(/<c:pt idx="\d+"><c:v><\/c:v><\/c:pt>/g, '');
        return '<c:ser>' + inner + '</c:ser>';
      });
      if (LANG === 'ko') xml = xml.replace(/<a:latin typeface="Noto Sans KR"\/>/g, '<a:latin typeface="Noto Sans KR"/><a:ea typeface="Noto Sans KR"/>');
      zip.file(name, xml);
    } else if (LANG === 'ko' && /^ppt\/(slides|notesSlides|slideLayouts|slideMasters)\/[^/]+\.xml$/.test(name)) {
      let xml = await zip.file(name).async('string');
      xml = xml.replace(/ lang="en-US"( altLang="en-US")?/g, ' lang="ko-KR" altLang="en-US"');
      xml = xml.replace(/<a:latin typeface="Noto Sans KR"\/>(?!<a:ea)/g, '<a:latin typeface="Noto Sans KR"/><a:ea typeface="Noto Sans KR"/>');
      zip.file(name, xml);
    } else if (LANG === 'ko' && name === 'ppt/theme/theme1.xml') {
      let xml = await zip.file(name).async('string');
      xml = xml.replace(/<a:ea typeface=""\/>/g, '<a:ea typeface="Noto Sans KR"/>').replace(/<a:font script="Hang" typeface="[^"]*"\/>/g, '<a:font script="Hang" typeface="Noto Sans KR"/>');
      zip.file(name, xml);
    }
  }
  fs.writeFileSync(file, await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }));
}

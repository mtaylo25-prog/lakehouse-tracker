// 09_Amazon_Listing_Kit.docx generator (node docx).
// Reads build/figures.json (SKU prices, gate thresholds), build/facts.json (naming rule, compliance facts), build/fmt.js (formatter).
// Usage: source build/env.sh && node build/build_listing.js
// Every model number comes from figures.json through fmt.js. No hyphens, en dashes or em dashes in any text.
'use strict';
const fs = require('fs');
const path = require('path');
const docx = require('docx');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, ShadingType, AlignmentType,
  HeadingLevel, LevelFormat, BorderStyle, PageNumber, Header, Footer, TableLayoutType, VerticalAlign,
} = docx;

const ROOT = process.env.ROOT || '/home/user/lakehouse-tracker/BRAND_US_Launch';
const BUILD = path.join(ROOT, 'build');
const OUT = path.join(ROOT, 'output');
const fig = JSON.parse(fs.readFileSync(path.join(BUILD, 'figures.json'), 'utf8'));
const facts = JSON.parse(fs.readFileSync(path.join(BUILD, 'facts.json'), 'utf8'));
const F = require(path.join(BUILD, 'fmt.js'));

// ---------- numbers and facts ----------
const G = fig.gate_thresholds;
const gv = (k) => G[k].value;
const hero = fig.unit_economics.hero;
const L3 = fig.levers.find((l) => l.n === 3) || fig.levers[2];
const INS = fig.insurance_trigger_month; // first month monthly sales pass the insurance trigger, by scenario
const ADS_M1_3 = facts.august.m18_assumptions.ads_phases[0][2]; // ad spend share of revenue, Months 1 to 3 (August model)
const skus = fig.unit_economics.per_sku;
const sku8k = skus.filter((s) => s.sku.startsWith('8K'));
const sku4k = skus.filter((s) => s.sku.startsWith('4K'));
const built = fig.meta.built;
const CAT8K = facts.compliance.category_8k; // Ultra High Speed HDMI Cable
const CAT4K_PLACEHOLDER = 'Premium High Speed HDMI Cable';
const built_fx = F.num(fig.meta.fx);
const ft = (m) => Math.round(m * 3.28084);
const skuName = (s) => s.sku.replace('_', ' ');
const lenLabel = (s) => `${ft(s.length_m)} ft (${s.length_m} m)`;
const lc = (t) => t.charAt(0).toLowerCase() + t.slice(1);

// ---------- guards: title length, keyword bytes, hyphen free ----------
const DASH = /[-–—−]/;
function guardText(t, where) {
  if (DASH.test(t)) throw new Error(`dash found in ${where}: ${t}`);
  return t;
}

// ---------- listing copy ----------
// Title template (8K line). Category name, 48Gbps and 8K allowed; never a version number.
function title8k(s) {
  return guardText(`BRAND Fiber Optic HDMI Cable ${lenLabel(s)}, ${CAT8K}, 48Gbps 8K 60Hz 4K 120Hz, Direction Proof Hybrid Active Optical, eARC, for PS5 Xbox Projector AV Receiver`, 'title ' + s.sku);
}
// 4K line: the category name is a placeholder until the certification is in hand, so the live title carries no category name.
function title4k(s) {
  return guardText(`BRAND Fiber Optic HDMI Cable ${lenLabel(s)}, 4K 60Hz 18Gbps Hybrid Active Optical Long Run HDMI Cord, Direction Proof, for Projector AV Receiver Conference Room Signage`, 'title ' + s.sku);
}
const useCase = {
  '8K_7M': 'a wall mounted TV to a media cabinet across the room, or a short ceiling projector run',
  '8K_10M': 'a ceiling projector in a mid size living room or home theater',
  '8K_15M': 'a large home theater, a conference room, or a projector mounted at the back of a long room',
  '8K_20M': 'a media room where the rack sits in a closet, or a gaming PC in another room',
  '8K_25M': 'a large conference room, classroom, or house of worship',
  '8K_50M': 'auditoriums, digital signage, trade show floors and commercial installs',
  '4K_15M': 'a 4K projector in a mid size room',
  '4K_20M': 'a 4K projector or display across a large room',
  '4K_30M': 'a 4K conference room or classroom with the source in an equipment closet',
  '4K_40M': 'a 4K digital signage or auditorium run',
};
function bullets8k(s) {
  return [
    `THE SIGNAL ARRIVES: fiber cores carry the full 48Gbps video signal the whole ${lenLabel(s)}, while thin copper conductors carry power and control. Copper HDMI cables physically fail past 10 ft (3 m) at 48Gbps. This hybrid active optical cable does not.`,
    `DIRECTION PROOF: the SOURCE end and the DISPLAY end are color coded, printed with arrows, and shown in the QR setup video on the box. Plug the SOURCE end into the player, console or PC and the DISPLAY end into the TV or projector. No guessing, no black screen.`,
    `5 YEAR SIGNAL GUARANTEE: if the signal drops at any time in 5 years, BRAND ships a replacement before your return arrives. Register with the card in the box to activate the guarantee and get the Spec Match finder for your setup.`,
    `MADE FOR THE LONG RUN: the right length for ${useCase[s.sku]}. Thin, light fiber construction pulls through tight spaces. [PLACEHOLDER: detachable head for conduit pulls, confirm on the production unit]. [PLACEHOLDER: CL3 in wall rating, add only after the CL3 documentation is on file].`,
    `WORKS WITH YOUR GEAR: 48Gbps bandwidth supports 8K 60Hz and 4K 120Hz signal modes, eARC audio, and HDR passthrough as defined for the ${CAT8K} category. Fits PS5, Xbox Series X, gaming PCs, 4K and 8K TVs, projectors and AV receivers. Powered by the HDMI port, no adapter needed.`,
  ].map((b) => guardText(b, 'bullet ' + s.sku));
}
function bullets4k(s) {
  return [
    `THE SIGNAL ARRIVES: fiber cores carry the full 4K 60Hz 18Gbps video signal the whole ${lenLabel(s)}, while thin copper conductors carry power and control. Long copper HDMI cables lose the signal at this length. This hybrid active optical cable does not.`,
    `DIRECTION PROOF: the SOURCE end and the DISPLAY end are color coded, printed with arrows, and shown in the QR setup video on the box. Plug the SOURCE end into the player or PC and the DISPLAY end into the TV or projector.`,
    `5 YEAR SIGNAL GUARANTEE: if the signal drops at any time in 5 years, BRAND ships a replacement before your return arrives. Register with the card in the box to activate the guarantee.`,
    `MADE FOR THE LONG RUN: the right length for ${useCase[s.sku]}. Thin, light fiber construction pulls through tight spaces. [PLACEHOLDER: detachable head for conduit pulls, confirm on the production unit]. [PLACEHOLDER: CL3 in wall rating, add only after the CL3 documentation is on file].`,
    `WORKS WITH YOUR GEAR: 4K 60Hz, HDR passthrough and audio return channel for 4K TVs, projectors, AV receivers, media players and PCs. Powered by the HDMI port, no adapter needed. [PLACEHOLDER: ${CAT4K_PLACEHOLDER} category name, add after certification].`,
  ].map((b) => guardText(b, 'bullet ' + s.sku));
}
const titles = {};
for (const s of sku8k) { titles[s.sku] = title8k(s); if (titles[s.sku].length >= 200) throw new Error(`title ${s.sku} is ${titles[s.sku].length} chars`); }
for (const s of sku4k) { titles[s.sku] = title4k(s); if (titles[s.sku].length >= 200) throw new Error(`title ${s.sku} is ${titles[s.sku].length} chars`); }
// 4K title once the certificate is filed: category name after the length, "Conference Room Signage" dropped. Checked here so the note in section 2 is true.
function title4kCert(s) { return guardText(`BRAND Fiber Optic HDMI Cable ${lenLabel(s)}, ${CAT4K_PLACEHOLDER}, 4K 60Hz 18Gbps Hybrid Active Optical Long Run HDMI Cord, Direction Proof, for Projector AV Receiver`, 'certified title ' + s.sku); }
const maxTitle4kCert = Math.max(...sku4k.map((s) => title4kCert(s).length));
if (maxTitle4kCert >= 200) throw new Error(`certified 4K title reaches ${maxTitle4kCert} chars`);
// Backend keywords: lowercase, space separated, no commas, no competitor brand names, no version numbers, no claims the listing cannot prove,
// and no word that already appears in any title of that parent (Amazon indexes title words already; the generator checks every word).
const KW8K = 'cord long extension ceiling mount gaming home theater soundbar monitor laptop pc media player streaming box uhd 2160p 4k60 25ft 30ft 40ft 50ft 65ft 75ft 100ft 150ft fibre hdr thin lightweight conduit rack closet';
const KW4K = 'extension ceiling mount classroom church auditorium office boardroom display tv monitor laptop pc media player uhd 2160p 50ft 65ft 75ft 100ft 130ft fibre hdr thin lightweight conduit rack closet commercial install';
function kwBytes(s) { return Buffer.byteLength(s, 'utf8'); }
function titleWords(list) { return new Set(list.map((s) => titles[s.sku]).join(' ').toLowerCase().replace(/[(),]/g, ' ').split(/\s+/).filter(Boolean)); }
for (const [k, v, list] of [['8K', KW8K, sku8k], ['4K', KW4K, sku4k]]) {
  guardText(v, 'keywords ' + k);
  if (kwBytes(v) >= 250) throw new Error(`backend keywords ${k} are ${kwBytes(v)} bytes, must be under 250`);
  const tw = titleWords(list);
  const dup = v.split(' ').filter((w) => tw.has(w));
  if (dup.length) throw new Error(`backend keywords ${k} repeat title words: ${dup.join(' ')}`);
}

// ---------- style helpers ----------
const FONT = 'Arial';
const PAGE_W = 12240, PAGE_H = 15840, MARGIN = 1296; // US Letter, 0.9 inch margins
const CONTENT_W = PAGE_W - 2 * MARGIN; // 9648 DXA
const INK = '1F2937', MUTED = '4B5563', HEAD_FILL = 'E8EEF7', LABEL_FILL = 'F3F6FB', ACCENT = '0B4F8A', WARN_FILL = 'FFF4E0', HERO_FILL = 'E6F7F1';

function run(text, opts = {}) { return new TextRun({ text, font: FONT, size: opts.size || 22, bold: opts.bold, italics: opts.italics, color: opts.color || INK }); }
function p(text, opts = {}) {
  const runs = Array.isArray(text) ? text : [run(text, opts)];
  return new Paragraph({ children: runs, spacing: { after: opts.after == null ? 120 : opts.after, before: opts.before || 0 }, alignment: opts.align, keepNext: opts.keepNext, shading: opts.fill ? { type: ShadingType.CLEAR, fill: opts.fill, color: 'auto' } : undefined });
}
function small(text, opts = {}) { return p(text, { size: 18, color: MUTED, ...opts }); }
function h1(text) { return new Paragraph({ text, heading: HeadingLevel.HEADING_1, spacing: { before: 360, after: 160 }, keepNext: true }); }
function h2(text) { return new Paragraph({ text, heading: HeadingLevel.HEADING_2, spacing: { before: 240, after: 120 }, keepNext: true }); }
function h3(text) { return new Paragraph({ text, heading: HeadingLevel.HEADING_3, spacing: { before: 160, after: 80 }, keepNext: true }); }
function bullet(text, opts = {}) {
  const runs = Array.isArray(text) ? text : [run(text, opts)];
  return new Paragraph({ children: runs, numbering: { reference: 'bullets', level: 0 }, spacing: { after: 60 } });
}
let numRef = 0;
function numberedList(items) {
  numRef += 1;
  const ref = `num${numRef}`;
  return items.map((t) => new Paragraph({ children: Array.isArray(t) ? t : [run(t)], numbering: { reference: ref, level: 0 }, spacing: { after: 60 } }));
}
const thin = { style: BorderStyle.SINGLE, size: 4, color: 'B8C2D1' };
const borders = { top: thin, bottom: thin, left: thin, right: thin, insideHorizontal: thin, insideVertical: thin };
function cell(children, width, opts = {}) {
  const kids = (Array.isArray(children) ? children : [children]).map((c) => (typeof c === 'string' ? p(c, { size: opts.size || 19, after: 40, bold: opts.bold, color: opts.color }) : c));
  return new TableCell({
    children: kids.length ? kids : [p('', { after: 0 })],
    width: { size: width, type: WidthType.DXA },
    shading: opts.fill ? { type: ShadingType.CLEAR, fill: opts.fill, color: 'auto' } : undefined,
    margins: { top: 60, bottom: 60, left: opts.hmargin || 100, right: opts.hmargin || 100 },
    verticalAlign: opts.valign || VerticalAlign.TOP,
    columnSpan: opts.span,
  });
}
function table(widths, rows, opts = {}) {
  return new Table({
    width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA },
    columnWidths: widths,
    layout: TableLayoutType.FIXED,
    borders,
    rows: rows.map((r, i) => new TableRow({ children: r, tableHeader: i === 0 && opts.headerRow !== false, cantSplit: true })),
  });
}
// Header row paragraphs carry keepNext so a header never sits alone at the foot of a page. Body cell paragraphs never carry keepNext; rows carry cantSplit (table()).
function headerRow(labels, widths, opts = {}) { return labels.map((t, i) => cell(p(t, { size: opts.size || 19, bold: true, after: 0, color: ACCENT, keepNext: true }), widths[i], { fill: HEAD_FILL, hmargin: opts.hmargin })); }
// opts.rowOpts and opts.headerOpts: { size, hmargin }.
function gridTable(labels, widths, rows, opts = {}) {
  const ro = { ...(opts.rowOpts || {}) };
  const body = rows.map((r) => r.map((v, i) => (v && v.constructor && v.constructor.name === 'TableCell' ? v : cell(typeof v === 'string' ? p(v, { size: ro.size || 19, after: 0 }) : v, widths[i], ro))));
  return table(widths, [headerRow(labels, widths, opts.headerOpts || {}), ...body], opts);
}
function kvTable(rows, w0 = 2000) {
  const W = [w0, CONTENT_W - w0];
  return table(W, rows.map(([k, v]) => [cell(p(k, { size: 19, bold: true, after: 0 }), W[0], { fill: LABEL_FILL }), cell(Array.isArray(v) ? v : [typeof v === 'string' ? p(v, { size: 19, after: 0 }) : v], W[1])]), { headerRow: false });
}
function note(text) { return p([run(text, { size: 19 })], { fill: WARN_FILL, after: 120 }); }
function spacer(after = 120) { return new Paragraph({ children: [], spacing: { after } }); }
const B = (t) => run(t, { bold: true });

// ---------- numbering config ----------
const numbering = {
  config: [
    { reference: 'bullets', levels: [{ level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 360, hanging: 240 } } } }] },
    ...Array.from({ length: 30 }, (_, i) => ({
      reference: `num${i + 1}`,
      levels: [{ level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 420, hanging: 320 } } } }],
    })),
  ],
};

// ---------- content ----------
const children = [];

// Title block
children.push(new Paragraph({ children: [run('09 Amazon Listing Kit', { size: 44, bold: true, color: ACCENT })], spacing: { before: 300, after: 120 } }));
children.push(p([run('BRAND fiber optic HDMI (High Definition Multimedia Interface) cables on Amazon US. Parent and child structure, titles, bullets, A+ outline, backend keywords, image and video brief, FAQ (frequently asked questions), compliance notes, guarantee copy and launch prices.', { size: 24 })], { after: 240 }));
children.push(p([B('Version: '), run(`built ${built}. Prices from ${fig.meta.workbook} (the SKU Unit Economics tab; SKU is stock keeping unit) through build/figures.json at the FX (foreign exchange) rate of ${built_fx} KRW per USD. All listing content is English and stays English in every language version of the package.`)]));
children.push(p([B('Readers: '), run('Mark (listing owner), the virtual assistant (VA) who loads the listing, Lucy (approves prices and claims), the factory (image and spec confirmations), the photographer.')]));
children.push(p([B('Standing rules: '), run(`never write an HDMI version number on the box, in the listing or in promotion. Use the cable category name: ${CAT8K} for the 8K line. Titles may say 48Gbps and 8K. Say "fiber optic HDMI" or "hybrid active optical"; never "pure fiber" or "all glass". Every bracketed [PLACEHOLDER] is a fact that does not exist yet and must be filled from a document, never from memory: the ASIN (Amazon Standard Identification Number), the FNSKU (Fulfillment Network Stock Keeping Unit, Amazon's barcode), the UPC (Universal Product Code) or its GTIN (Global Trade Item Number) exemption, certificate numbers and URLs. Prices are launch estimates (management estimate) until Lucy signs off.`)]));
children.push(p([B('Why the listing matters to the plan: '), run(`the Day 90 gate needs unit session conversion at or above ${F.pct(gv('d90_conv_pass'), 0)} on at least 3 of the 5 launch SKUs (kill if the hero is below ${F.pct(gv('d90_conv_kill'), 0)}), TACOS (total advertising cost of sales) at or below ${F.pct(gv('d90_tacos_pass'), 0)} in Month 3, and a return rate at or below ${F.pct(gv('d90_ret_pass'), 0)}. The direction proof copy and images exist to cut returns; the length guide exists to lift conversion.`)]));

// Contents
children.push(h2('Contents'));
children.push(gridTable(['No.', 'Section', 'Who uses it'], [600, 5400, 3648], [
  ['1', 'Parent and child ASIN structure', 'Mark, VA'],
  ['2', 'Titles (under 200 characters)', 'Mark, VA'],
  ['3', 'Bullets (five per child)', 'Mark, VA'],
  ['4', 'A+ content outline (7 modules)', 'Mark, designer'],
  ['5', 'Backend keywords (under 250 bytes)', 'Mark, VA'],
  ['6', 'Image brief (7 slots) and the video slot', 'Photographer, factory, Mark'],
  ['7', 'FAQ', 'VA (customer questions), Mark'],
  ['8', 'Direction proof copy and 5 Year Signal Guarantee copy', 'Mark, factory (insert card)'],
  ['9', 'Compliance notes (HDMI naming, HDMI Adopter, FCC (Federal Communications Commission), Prop 65 (California Proposition 65 warning label), CL3 (Class 3 in wall cable fire rating))', 'Lucy, factory, Mark'],
  ['10', 'Launch prices (estimates) and price ladder', 'Lucy, Mark'],
  ['11', 'Pre publish checklist', 'Mark, VA'],
  ['12', 'Glossary', 'Everyone'],
]));

// 1. Parent and child structure
children.push(h1('1. Parent and child ASIN structure'));
children.push(p('One parent per line, children by length. The variation theme is Size (cable length). All lengths of one line share reviews, questions and the A+ page under the parent, which is lever 6 in the model (one parent instead of separate ASINs: conversion, review pooling, ad efficiency). The parent itself is never sold; each child carries its own ASIN, FNSKU and price.'));
children.push(p([B('Launch order: '), run(`the 5 launch SKUs (${skus.filter((s) => s.launch_sku === 'Yes').map(skuName).join(', ')}) go live first, air freighted at 100 to 300 units each. The other children are created in the catalog at launch but stay out of stock until the Day 90 gate passes, so the length ladder is visible to shoppers from day one.`)]));

children.push(h2(`Parent A: BRAND ${CAT8K}, fiber optic (8K line)`));
const W1 = [1250, 1700, 1050, 1050, 1900, 2698];
const IDS = '[PLACEHOLDER: ASIN], [PLACEHOLDER: FNSKU], [PLACEHOLDER: UPC or GTIN exemption]';
children.push(gridTable(['Child SKU', 'Variation value (Size)', 'List price', 'Sale price', 'Launch SKU', 'Identifiers'], W1,
  sku8k.map((s) => {
    const heroRow = s.sku === '8K_10M';
    const vals = [skuName(s), lenLabel(s), F.money(s.us_list, 2), F.money(s.us_sale_price, 2), s.launch_sku === 'Yes' ? 'Yes' : 'No (after Day 90 gate)', IDS];
    return vals.map((v, i) => cell(p(v, { size: 19, bold: heroRow && i === 0, after: 0 }), W1[i], { fill: heroRow ? HERO_FILL : undefined }));
  })));
children.push(small(`Hero SKU 8K 10M shaded. Spec basis: the internal spec sheet and the factory ATC (HDMI Authorized Test Center) test report; only the category name ${CAT8K} appears on the listing. Prices are launch estimates (management estimate) from the SKU table in the model; see section 10. Parent ASIN: [PLACEHOLDER: parent ASIN after listing creation]. Brand name field: BRAND (placeholder until the trademark name is chosen).`));

children.push(h2('Parent B: BRAND fiber optic HDMI cable, 4K line'));
children.push(gridTable(['Child SKU', 'Variation value (Size)', 'List price', 'Sale price', 'Launch SKU', 'Identifiers'], W1,
  sku4k.map((s) => [skuName(s), lenLabel(s), F.money(s.us_list, 2), F.money(s.us_sale_price, 2), s.launch_sku === 'Yes' ? 'Yes' : 'No (after Day 90 gate)', IDS])));
children.push(note(`Spec basis: [PLACEHOLDER: ${CAT4K_PLACEHOLDER} certification]. The 4K line category name is a placeholder pending certification. Until the certificate is on file, the 4K titles and bullets carry no category name and say only 4K 60Hz 18Gbps. The 4K 30M child is a launch SKU; the other three wait for the Day 90 gate.`));

children.push(h2('Listing attributes shared by every child'));
children.push(kvTable([
  ['Brand', 'BRAND [PLACEHOLDER: final brand name, must match the trademark application and Brand Registry]'],
  ['Manufacturer', 'BRAND LLC (limited liability company, Wyoming) for the listing; the factory name stays on the compliance file, not the listing'],
  ['Product type', 'HDMI cable (Electronics, Accessories and Supplies, Audio and Video Accessories, Cables and Interconnects, HDMI Cables)'],
  ['Variation theme', 'Size (cable length in feet and meters, written "33 ft (10 m)")'],
  ['Connector type', 'HDMI Type A male to HDMI Type A male'],
  ['Cable construction', 'Hybrid active optical: fiber cores for video data, thin copper conductors for power and control'],
  ['Color', 'Black [PLACEHOLDER: confirm jacket color on the production unit]'],
  ['Country of origin', 'China (required at entry and on the packaging)'],
  ['Item package quantity', '1'],
  ['Unit count', '1 Count'],
  ['Included components', 'Cable, QR (quick response code) setup card, guarantee registration card'],
  ['Warranty description', '5 Year Signal Guarantee (section 8)'],
  ['Compliance fields', 'Prop 65 warning: Yes with the statement in section 9. FCC: [PLACEHOLDER: factory FCC test report number]. Battery: none.'],
]));

// 2. Titles
children.push(h1('2. Titles'));
children.push(p(`Rules: under 200 characters, category name ${CAT8K} on the 8K line, never an HDMI version number, 48Gbps and 8K allowed, length first in feet with meters in brackets so the Size variation reads the same way, no promotional words (free, best, sale), no symbols other than commas and brackets. eARC is enhanced audio return channel; AV is audio video. The generator checks every title length before writing this file.`));
const W2 = [1300, 6850, 1498];
children.push(gridTable(['Child', 'Title', 'Characters'], W2, [...sku8k, ...sku4k].map((s) => [skuName(s), titles[s.sku], String(titles[s.sku].length)])));
children.push(small(`When the 4K certificate is filed, insert "${CAT4K_PLACEHOLDER}" after the length and drop "Conference Room Signage" so every 4K title stays under 200 characters (longest becomes ${maxTitle4kCert} characters, 4K 40M).`));
children.push(small(`Parent title (not shown to shoppers once children exist): "BRAND Fiber Optic HDMI Cable, ${CAT8K}, 48Gbps 8K, Direction Proof Hybrid Active Optical" and "BRAND Fiber Optic HDMI Cable, 4K 60Hz Hybrid Active Optical, Long Run".`));

// 3. Bullets
children.push(h1('3. Bullets (five per child)'));
children.push(p('Each bullet opens with a capitalized benefit label, then the proof. Bullets 1, 2, 3 and 5 are shared across the line with the length filled in; bullet 4 carries the length specific use case. Keep every bullet under 500 characters (Amazon trims longer bullets on mobile). HDR is high dynamic range video. No bullet carries a URL (web address) or any other off Amazon link; the registration address lives on the insert card in the box (section 8). The two placeholders in bullet 4 (detachable head, CL3) stay out of the live listing until the factory confirms the feature and the CL3 documentation is on file.'));
for (const s of [...sku8k, ...sku4k]) {
  children.push(h3(`${skuName(s)}: ${lenLabel(s)}${s.sku === '8K_10M' ? ' (hero)' : ''}`));
  const bl = s.sku.startsWith('8K') ? bullets8k(s) : bullets4k(s);
  for (const b of bl) children.push(bullet(b, { size: 20 }));
}

// 4. A+ outline
children.push(h1('4. A+ content outline (7 modules)'));
children.push(p(`A+ content needs Brand Registry, which accepts a pending trademark application (${lc(facts.trademark.rule)}, about ${F.money(facts.trademark.uspto_fee_per_class_usd)} per class plus attorney fees). Build one A+ page and attach it to every child of both parents. Image sizes are the Amazon standard module sizes at the time of upload; the designer checks the current size sheet.`));
const W4 = [600, 2100, 3500, 3448];
children.push(gridTable(['No.', 'Module', 'Content', 'Image and copy'], W4, [
  ['1', 'Brand banner (Standard Image Header with Text)', 'Positioning line: Sell the signal, not the cable. One sentence: BRAND builds long run fiber optic HDMI cables for the installs where copper fails.', 'Wide hero image of a cable run across a room to a ceiling projector. Logo [PLACEHOLDER: BRAND logo]. 40 words max.'],
  ['2', 'Why copper fails past 3 m at 48Gbps (Standard Image and Text Overlay)', 'Plain explanation: at 48Gbps a copper cable loses the signal after about 10 ft (3 m). Fiber cores carry light, so the length does not matter to the signal. Hybrid design: fiber for video, thin copper for power and control.', 'Diagram image: copper run with a fading signal versus fiber run with a solid signal. No competitor names. Accurate wording only: never "pure fiber" or "all glass".'],
  ['3', 'Direction proof setup in three steps (Standard Three Images and Text)', 'Step 1: find the SOURCE end (color and arrow). Step 2: plug SOURCE into the player, console or PC. Step 3: plug DISPLAY into the TV or projector. QR code to the setup video.', 'Three close up photos of the labeled heads. Text under each photo of 25 words max. The QR code image is a placeholder until the video URL exists.'],
  ['4', 'Length guide (Standard Comparison Chart)', `Columns: every child of the 8K line (${sku8k.map((s) => lenLabel(s)).join(', ')}). Rows: typical room, 8K 60Hz, 4K 120Hz, eARC, in wall [PLACEHOLDER: CL3], weight, guarantee. Checkmarks link each column to its child ASIN.`, 'One small product image per column (same photo, length label changed). The chart is the biggest conversion lever on the page: a shopper who picks the right length does not return the cable.'],
  ['5', 'Install (Standard Four Image and Text)', 'Four panels: pulling through a wall plate, pulling through conduit with the detachable head [PLACEHOLDER: confirm feature], ceiling projector drop, rack to display. Bend radius and maximum pull force: [PLACEHOLDER: factory spec sheet values].', 'Four photos or renders of real installs. In wall wording only after the CL3 documentation is on file.'],
  ['6', '5 Year Signal Guarantee and registration (Standard Single Image and Sidebar)', 'The guarantee copy from section 8. Replacement shipped before the return arrives. Registration builds the Spec Match finder and the trade up offer (lever 11, off Amazon email list).', 'Image of the registration card and the QR code. Sidebar: three bullets, 20 words each.'],
  ['7', 'Tech specs and compliance (Standard Tech Specs)', `Category: ${CAT8K}. Bandwidth 48Gbps. Signal modes 8K 60Hz and 4K 120Hz. eARC. Connector HDMI Type A to Type A. Construction hybrid active optical. Power from the HDMI port. Compliance line: [PLACEHOLDER: HDMI Adopter certificate number], [PLACEHOLDER: FCC test report number], Prop 65 statement (section 9).`, 'No image needed. Every value comes from the factory ATC (HDMI Authorized Test Center) test report; nothing is typed from memory.'],
]));
children.push(small('Brand Store (also gated by Brand Registry): one page per line plus a length guide page. Reuse modules 2, 4 and 6. Video ads reuse the setup video from section 6.'));

// 5. Backend keywords
children.push(h1('5. Backend keywords'));
children.push(p('Rules: under 250 bytes, lowercase, separated by single spaces, no commas, no repeats of words that already appear in the titles of that parent (Amazon indexes the title words already; the generator checks every keyword against every title word before it writes this file), no competitor brand names, no HDMI version numbers (the naming rule applies to hidden fields as well), no claims the listing cannot prove. Feet values are the rounded lengths shoppers search for near the lengths the line sells.'));
children.push(kvTable([
  ['Parent A (8K line)', [p(KW8K, { size: 19, after: 40 }), small(`${kwBytes(KW8K)} bytes`, { after: 0 })]],
  ['Parent B (4K line)', [p(KW4K, { size: 19, after: 40 }), small(`${kwBytes(KW4K)} bytes`, { after: 0 })]],
  ['Pruning rule (lever 4)', 'Every two weeks the VA exports the search term report. Terms with 20 clicks and no order go to the negative list of every campaign. Terms with orders move from automatic to exact match. Keep the backend field under 250 bytes after each change.'],
]));

// 6. Image brief
children.push(h1('6. Image brief (7 slots) and the video slot'));
children.push(p('All images 2000 by 2000 pixels minimum, square, sRGB, JPEG. Slot 1 follows the Amazon main image rules: pure white background, product only, no text, no badges, product fills at least 85% of the frame. Slots 2 to 7 are lifestyle and infographic images with short text (20 words max per image, 60 pt or larger so it reads on a phone). The four 1600 by 1600 listing images from the August work were not present in ./inputs for this build; if they are found, map them to slots 1, 3, 4 and 7 and reshoot only the gaps. [PLACEHOLDER: confirm with Amazon\'s current image and A+ guidelines whether a QR code or URL may appear in listing images; if not, the QR code stays on the box and insert card only].'));
const W6 = [700, 1900, 4000, 3048];
children.push(gridTable(['Slot', 'Name', 'What the image shows', 'Text on image and notes'], W6, [
  ['1', 'Main image', 'The coiled cable with both heads visible in front, SOURCE and DISPLAY colors readable. White background, no props.', 'No text (Amazon rule). Shoot once per length with the same coil size; only the length label on the box changes in slot 2.'],
  ['2', 'Length in a room', 'A top down room diagram with the cable drawn from a media cabinet to a ceiling projector, the length written in feet and meters.', `"${lenLabel(sku8k[1])}" style label per child. One master layout, the length text swaps per child. Shows the shopper what the number means.`],
  ['3', 'Direction proof', 'Close up of both heads side by side. SOURCE end and DISPLAY end color coded with printed arrows.', '"SOURCE end to your player. DISPLAY end to your screen." QR code to the setup video.'],
  ['4', 'Why fiber', 'Infographic: a copper cable fading after 10 ft (3 m) at 48Gbps next to a fiber cable solid to the end of the run.', '"48Gbps the whole way." Hybrid note in small text: fiber for video, copper for power. No competitor names, no "pure fiber".'],
  ['5', 'Works with your gear', 'Generic device silhouettes (console, gaming PC, TV, projector, AV receiver) connected by the cable. No third party logos.', '"8K 60Hz. 4K 120Hz. eARC." Powered by the HDMI port, no adapter.'],
  ['6', 'Install', 'The cable pulled through a wall plate and a conduit; the detachable head shown apart from the cable.', '"Thin. Light. Pulls through conduit." Detachable head and in wall text are placeholders until the factory confirms the feature and CL3 documentation is on file.'],
  ['7', 'Guarantee', 'The registration card, the QR code and a calendar graphic for 5 years.', '"5 Year Signal Guarantee. Replacement ships before your return arrives." Registration URL placeholder.'],
  ['Video', 'Setup video (under 45 seconds)', 'Unbox, find the SOURCE end by color and arrow, plug into the console, plug DISPLAY into the TV, picture appears, scan the QR for the length guide. Shot on a phone is acceptable for launch; captions on, no music licensing needed (silent or royalty free track with the license on file).', 'The same file serves the listing video slot, the A+ module 3 QR code and the Brand Registry video ads. [PLACEHOLDER: video URL].'],
]));

// 7. FAQ
children.push(h1('7. FAQ'));
children.push(p('Use these answers word for word in the customer questions section, in the insert card and in the VA reply templates. Every answer stays inside the naming rule and the placeholder rule.'));
const faq = [
  ['Does it matter which end goes where?', 'Yes. This is an active cable with electronics in each head. The SOURCE end (color coded, arrow printed) goes into the device that sends the picture: console, PC, player, receiver output. The DISPLAY end goes into the TV, monitor or projector. Plug it backwards and you get no picture; swap the ends and it works. The QR code on the box opens a setup video under 45 seconds. [PLACEHOLDER: factory confirmation that a reversed connection causes no damage; only then add "nothing is damaged" here and on the insert card].'],
  ['Can I run it inside a wall or ceiling?', `An in wall run needs a cable with an in wall fire rating (CL3, Class 3). [PLACEHOLDER: CL3 documentation]. Until that documentation is on file the answer is: "We will confirm the in wall rating for this cable; for now run it through conduit or a raceway." Never claim CL3 before the certificate exists.`],
  ['Will it work with my projector and AV receiver?', 'Yes for any projector or receiver with an HDMI input or output. Typical setup: receiver HDMI output to the SOURCE end, projector to the DISPLAY end. If the receiver sits between the console and the projector, the long fiber run goes from the receiver to the projector; a short cable connects the console to the receiver.'],
  ['Does it support 4K 120Hz on PS5 and Xbox Series X?', `Yes. The 8K line carries 48Gbps, which is the bandwidth the ${CAT8K} category defines for 4K 120Hz and 8K 60Hz modes. Turn on the 120Hz mode in the console video settings and make sure the TV input is set to its high bandwidth mode. Values confirmed against the factory ATC test report [PLACEHOLDER: report number].`],
  ['Does eARC work over this cable?', 'Yes. eARC (enhanced audio return channel) sends audio from the TV back to a soundbar or receiver. Plug the DISPLAY end into the TV eARC port and the SOURCE end into the receiver or soundbar eARC input. Confirmed on the ATC test report [PLACEHOLDER: report number].'],
  ['Which length should I buy?', `Measure the real path (along the wall, up to the ceiling, across, down), not the straight line, then add 6 ft (2 m) for connections and slack. Lengths: ${sku8k.map((s) => lenLabel(s)).join(', ')}. A cable that is too long is fine; one that is too short is a return.`],
  ['Does it need external power?', 'No. The heads draw power from the HDMI port of the source device. If a very old source device cannot supply the power, [PLACEHOLDER: factory guidance on a USB power option].'],
  ['Is this a pure fiber cable?', 'No. It is a hybrid active optical cable: fiber cores carry the video data and thin copper conductors carry power and control signals. That hybrid design is what lets the cable power itself from the HDMI port.'],
  ['What HDMI version is it?', `Cables are certified by category, not by version. This cable is sold as an ${CAT8K} (48Gbps, 8K). The HDMI Adopter certificate and ATC test report are on file: [PLACEHOLDER: certificate number].`],
  ['How tight can I bend it, and can I pull it through conduit?', 'Minimum bend radius: [PLACEHOLDER: factory spec sheet value]. Maximum pull force: [PLACEHOLDER: factory spec sheet value]. [PLACEHOLDER: detachable head for conduit pulls, confirm on the production unit]. Never pull by the head.'],
  ['What happens if the signal fails?', 'Register the cable. If the signal drops at any time within 5 years, contact BRAND with the order number. A replacement ships before the return arrives (section 8). Amazon return rules apply for the first 30 days regardless.'],
  ['Is it compatible with 8K TVs and HDR?', 'Yes. 48Gbps carries 8K 60Hz and HDR formats passed through by the source and display. The cable does not create HDR; it carries what the devices send.'],
];
children.push(gridTable(['Question', 'Answer'], [3000, 6648], faq));

// 8. Direction proof and guarantee copy
children.push(h1('8. Direction proof copy and 5 Year Signal Guarantee copy'));
children.push(h2('Direction proof: box, insert card and listing wording'));
children.push(kvTable([
  ['Box panel', 'SOURCE end: plug into your player, console or PC. DISPLAY end: plug into your TV or projector. Scan the QR code for the setup video (under 45 seconds).'],
  ['Head labels', 'SOURCE printed on one head with an arrow pointing away from the device; DISPLAY printed on the other head with an arrow pointing into the screen. Two colors: [PLACEHOLDER: color pair chosen with the factory; must read for color blind users, so the words and arrows stay].'],
  ['Insert card (front)', 'Three steps with icons: 1 Find the SOURCE end. 2 Plug SOURCE into the device that sends the picture. 3 Plug DISPLAY into the screen. If there is no picture, swap the ends. [PLACEHOLDER: add "Nothing is damaged" only after the factory confirms that a reversed connection causes no damage].'],
  ['Insert card (back)', 'Register your 5 Year Signal Guarantee: [PLACEHOLDER: registration URL or QR]. Registration unlocks the Spec Match finder (tells you the right length and mode for your devices) and the trade up offer when the next generation ships. We never ask for a review on this card.'],
  ['Listing wording', `Direction proof design: color coded SOURCE and DISPLAY ends, printed arrows, QR setup video. Reversed ends are the most common reason an AOC (active optical cable) comes back (management estimate; confirm against the return reasons of the first 90 days). The model credits this design with a return rate move from ${F.pct(L3.from)} to ${F.pct(L3.to)} from Year 2 (lever 3, Medium case).`],
]));
children.push(h2('5 Year Signal Guarantee'));
children.push(p('Listing and A+ wording (consumer facing, no URL):', { keepNext: true }));
children.push(p([run('"BRAND 5 Year Signal Guarantee. If this cable stops carrying a signal at any time in the 5 years after purchase, we ship a replacement before your return arrives. Register with the card in the box, with your order number. The guarantee covers signal failure of the cable in normal use; it does not cover cut, crushed or burned cables or damage to other devices. Amazon return rules apply for the first 30 days regardless."', { italics: true })], { fill: LABEL_FILL }));
children.push(p('Amazon warranty field: the same wording plus "Register at [PLACEHOLDER: registration URL]". [PLACEHOLDER: confirm against current Amazon policy that the warranty field may carry the registration address; if not, the address stays on the box and the insert card only].'));
children.push(p('Operating terms (internal, from the returns SOP (standard operating procedure) in 05_Operating_Processes.docx, section 6):', { keepNext: true }));
children.push(...numberedList([
  'Customer contacts BRAND through Amazon messaging or the registration site with the order number and a description of the fault.',
  'VA replies within 24 hours with the two step check (swap the ends, test another input) and the prepaid return label.',
  'Replacement ships the same day from FBA (Fulfillment by Amazon, through multi channel fulfillment) or the 3PL (third party logistics warehouse); the return is not a condition of the replacement.',
  'Returned cable is tested by the 3PL; failures are logged by SKU and lot and reported to the factory monthly for credit.',
  `Guarantee cost is carried in the returns reserve of the model (Medium ${F.pct(L3.from)} of revenue in Year 1, ${F.pct(L3.to)} from Year 2; figures.json lever 3). [PLACEHOLDER: legal review of the guarantee wording for US consumer warranty law before the listing goes live].`,
]));

// 9. Compliance notes
children.push(h1('9. Compliance notes'));
children.push(kvTable([
  ['HDMI naming rule', `${facts.compliance.naming_rule} On this listing: the 8K line is sold as ${CAT8K}; the 4K line waits for its certification before any category name appears. The internal spec column in the model (which carries the version numbers) never leaves the model and the factory file.`],
  ['HDMI Adopter and ATC', `${facts.compliance.hdmi_adopter.replace('HDMI LA runs', 'HDMI LA (HDMI Licensing Administrator) runs')}. Before the listing goes live: [PLACEHOLDER: HDMI Adopter certificate number from factory], [PLACEHOLDER: ATC test report number for the 8K line], [PLACEHOLDER: ATC test report number for the 4K line]. The Spec values in the A+ module 7 and the FAQ (48Gbps, 8K 60Hz, 4K 120Hz, eARC) are confirmed against the report, not typed from memory. Evidence based rule (brief); report numbers are placeholders.`],
  ['FCC', `${facts.compliance.fcc}. Listing compliance field and packaging statement: [PLACEHOLDER: FCC statement wording and report number supplied by the factory test lab; the wording depends on the device class the lab assigns]. Do not write an FCC statement before the report exists.`],
  ['California Prop 65', `${facts.compliance.prop65}. Packaging label and Amazon Prop 65 compliance field: [PLACEHOLDER: Prop 65 warning statement; the exact wording and the listed chemical depend on the factory material disclosure for the jacket and heads]. Amazon shows the warning on the listing when the field is set; the box label is printed by the factory.`],
  ['CL3 in wall', `${facts.compliance.cl3}. No title, bullet, image or answer says "in wall" or "CL3" until the certificate is in the compliance file. Section 3 bullet 4, A+ modules 4 and 5, image slot 6 and the FAQ carry the placeholder.`],
  ['Marketing language', `${facts.compliance.marketing_language}.`],
  ['Country of origin', 'China on the packaging and the entry documents. De minimis is gone: every shipment enters formally through the broker with the continuous bond; no parcel dropship.'],
  ['Brand Registry', `Needs a filed trademark application (${facts.trademark.brand_registry}). Gates: ${facts.trademark.brand_registry_gates.join(', ')}. Brand name on the listing must match the application exactly.`],
  ['Reviews', 'No incentives, no review requests on the insert card, no family or friend reviews. Only the Amazon Request a Review button and the Vine program (Amazon sends free units to selected reviewers; lever 5). Review counts and star ratings are never quoted in the package before they exist.'],
  ['Product liability insurance', `${F.money(facts.amazon.product_liability_insurance_usd)} cover required once sales pass ${F.money(facts.amazon.insurance_trigger_monthly_sales_usd)} in a month. [PLACEHOLDER: policy number], uploaded to Seller Central in ${F.month_label(INS.Medium)}, the first month sales pass ${F.money(facts.amazon.insurance_trigger_monthly_sales_usd)} in the Medium case (${F.month_label(INS.Low)} in the Low case, ${F.month_label(INS.High)} in the High case; figures.json insurance_trigger_month, the same month as 06_Compliance_Checklist.xlsx).`],
]));

// 10. Prices
children.push(h1('10. Launch prices (estimates) and price ladder'));
children.push(p(`Every price below is a launch estimate (management estimate pending the Amazon rate card check), pulled from the SKU Unit Economics tab of the model through figures.json at FX ${built_fx} KRW per USD and the Copper duty scenario (${F.pct(fig.duty.copper.rate ?? facts.duty.copper.total)}). The list price is the strike through reference; the sale price is the everyday price. Net per unit is after landed cost, referral (${F.pct(facts.amazon.referral_rate, 0)}), FBA fee, advertising (${F.pct(hero.ads / hero.price, 0)} of price) and the returns and storage reserve (${F.pct(hero.returns_storage / hero.price, 0)} of price), on air freight at launch. The hero SKU in the August model earned ${F.money(hero.net, 2)} per unit (${F.pct(hero.net_pct)}) at ${F.money(hero.price, 2)} on a ${F.pct(hero.landed_pct, 0)} landed cost; the model's air freight build gives a lower landed cost on the H2 cost base. At the Alibaba benchmark cost (FOB, free on board, ${F.money(facts.august.alibaba_benchmark.fob_low_usd)} to ${F.money(facts.august.alibaba_benchmark.fob_high_usd)} for a comparable 30 ft cable, landed ${F.money(fig.unit_economics.cost_bases.alibaba_low.landed, 2)} to ${F.money(fig.unit_economics.cost_bases.alibaba_high.landed, 2)}) the hero loses money at ${F.money(hero.price, 2)}: net ${F.money(fig.unit_economics.cost_bases.alibaba_low.net, 2)} to ${F.money(fig.unit_economics.cost_bases.alibaba_high.net, 2)} per unit. That gap is the largest unverified assumption in the plan.`));
// Contribution per unit before advertising (the Day 90 gate measure) = net per unit plus advertising, from figures.json.
const cbAir = (s) => s.net_air + s.ads;
const cbOcean = (s) => s.net_ocean + s.ads;
const cpuFloor = gv('d90_cpu_pass');
const belowAir = skus.filter((s) => cbAir(s) < cpuFloor);
const belowOcean = skus.filter((s) => cbOcean(s) < cpuFloor);
const low8kCb = sku8k.reduce((a, s) => (cbAir(s) < cbAir(a) ? s : a), sku8k[0]);
const s4k30 = skus.find((s) => s.sku === '4K_30M');
const keepShare = 1 - s4k30.referral / s4k30.us_sale_price - s4k30.returns_storage / s4k30.us_sale_price; // price share left after referral and the returns and storage reserve
const priceFloorExact = (cpuFloor + s4k30.landed_air + s4k30.fba_fee) / keepShare;
const priceFloor99 = Math.ceil(priceFloorExact - 0.99) + 0.99;
const W10 = [850, 1000, 870, 870, 900, 700, 900, 750, 1404, 1404];
const tightCell = { size: 18, hmargin: 70 };
const tightHead = { size: 17, hmargin: 70 };
children.push(gridTable(['Child', 'Length', 'List', 'Sale', 'Landed cost, air', 'FBA fee', 'Net per unit, air', 'Net %, air', 'Contribution before ads, air', 'Contribution before ads, ocean'], W10,
  skus.map((s) => [skuName(s), lenLabel(s), F.money(s.us_list, 2), F.money(s.us_sale_price, 2), F.money(s.landed_air, 2), F.money(s.fba_fee, 2), F.money(s.net_air, 2), F.pct(s.net_air_pct), F.money(cbAir(s), 2), F.money(cbOcean(s), 2)]),
  { rowOpts: tightCell, headerOpts: tightHead }));
const min4k = sku4k.reduce((a, s) => (s.net_air < a.net_air ? s : a), sku4k[0]);
const min8k = sku8k.reduce((a, s) => (s.net_air < a.net_air ? s : a), sku8k[0]);
children.push(small(`Basis: list and sale prices, FBA fees and weights are management estimates; H2 factory cost is evidence based (August 2026 deal). Source: build/figures.json, unit_economics.per_sku. Contribution before ads = net per unit plus advertising, the measure the Day 90 gate reads. Reading the table: the 8K line nets ${F.money(min8k.net_air, 2)} per unit or more on air freight (lowest: ${skuName(min8k)}); the 4K line nets as little as ${F.money(min4k.net_air, 2)} (${skuName(min4k)}, ${F.pct(min4k.net_air_pct)}), so the 4K children beyond the 30M launch SKU wait for ocean freight and the Day 90 gate before they scale.`));
children.push(note(`Floor check against rule 5 below: ${belowAir.map((s) => `${skuName(s)} ${F.money(cbAir(s), 2)}`).join(', ')} ${belowAir.length === 1 ? 'sits' : 'sit'} below the ${F.money(cpuFloor)} Day 90 floor on air freight (every 8K child clears it; lowest ${skuName(low8kCb)} ${F.money(cbAir(low8kCb), 2)}). The launch SKU 4K 30M reaches ${F.money(cbAir(s4k30), 2)} on air and ${F.money(cbOcean(s4k30), 2)} on ocean${belowOcean.length ? `; on ocean only ${belowOcean.map((s) => `${skuName(s)} (${F.money(cbOcean(s), 2)})`).join(' and ')} still ${belowOcean.length === 1 ? 'misses' : 'miss'} the floor` : ''}. Lucy decides before she signs the price sheet: either raise 4K 30M to at least ${F.money(priceFloor99, 2)} (${F.pct(keepShare, 0)} of price minus landed cost ${F.money(s4k30.landed_air, 2)} minus FBA fee ${F.money(s4k30.fba_fee, 2)} reaches ${F.money(cpuFloor)} at ${F.money(priceFloorExact, 2)}), or apply rule 5 to the 4K line only after the move to ocean freight. This kit uses the second option as its default (management estimate), so the model price of ${F.money(s4k30.us_sale_price, 2)} stands until she decides.`));
children.push(h2('Price ladder and coupon rules (lever 10)'));
children.push(...numberedList([
  'Launch at the sale price with the list price shown as the reference. Never set the list price above a price the product has sold at elsewhere; Amazon suppresses fake reference prices.',
  `Test plus or minus 5% on the hero (${F.money(hero.price * 0.95, 2)} to ${F.money(hero.price * 1.05, 2)}) in two week blocks after Day 30; keep the price that gives the higher contribution per unit, not the higher conversion.`,
  `Coupons of 5% to 10% in Months 1 to 3 only, funded from the launch advertising budget (${F.pct(ADS_M1_3, 0)} of revenue in Months 1 to 3 in the Medium case).`,
  'Each longer length prices above the next shorter one so the ladder reads cleanly on the comparison chart. Never price a longer cable below a shorter one.',
  `Floor: no child sells below the price that leaves contribution per unit before advertising under ${F.money(cpuFloor)} (Day 90 pass threshold; kill if the hero falls below ${F.money(gv('d90_cpu_kill'))}). The rule binds the 8K line from launch and the 4K line from the move to ocean freight, unless Lucy raises the 4K 30M price to ${F.money(priceFloor99, 2)} or more (floor check above); on air freight the 4K line misses the floor.`,
  'The Amazon price is the floor for every other channel (eBay, Walmart, the DTC (direct to consumer) site); B2B (business to business) installer pricing at 60% of retail sits in a separate, non public price list.',
]));

// 11. Pre publish checklist
children.push(h1('11. Pre publish checklist'));
children.push(gridTable(['Check', 'Owner', 'Done when'], [4900, 1750, 2998], [
  ['Every title under 200 characters, no version number, category name on the 8K line', 'VA', 'Checked in Seller Central preview'],
  ['Five bullets per child, no placeholder text left in a live bullet', 'Mark', 'Placeholders resolved or the sentence removed'],
  ['Backend keywords under 250 bytes, no competitor names, no version numbers', 'VA', 'Byte count recorded'],
  ['HDMI Adopter certificate and ATC report numbers in the compliance file', 'Lucy, factory', 'Scanned copies filed'],
  ['FCC test report filed; FCC statement added to packaging and listing', 'Lucy, factory', 'Report number recorded'],
  ['Prop 65 statement on the box and in the listing field', 'Lucy, factory', 'Label proof approved'],
  ['CL3 claim removed everywhere unless the certificate is filed', 'Mark', 'Search of the listing text for "CL3" and "in wall"'],
  ['Brand name matches the trademark application and Brand Registry', 'Mark, attorney', 'Brand Registry approved'],
  ['Prices approved by Lucy as launch estimates; list price not above any prior selling price', 'Lucy', 'Signed price sheet'],
  ['Seven images and the video uploaded; main image passes the white background rule', 'Photographer, VA', 'All slots show in preview'],
  ['FNSKU labels match the child ASINs; factory label proof approved', 'Lucy, factory', 'Label proof filed'],
  ['Guarantee wording reviewed for US consumer warranty law', 'Attorney', '[PLACEHOLDER: review date]'],
  ['Insert card contains no review request and no off Amazon purchase link', 'Mark', 'Card proof approved'],
  ['No URL or external link in any title, bullet or description', 'Mark, VA', 'Search of the live listing text for "http", "www" and ".com" finds nothing'],
]));

// 12. Glossary
children.push(h1('12. Glossary'));
const gl = [
  ['AOC', 'active optical cable'], ['HDMI', 'High Definition Multimedia Interface'], ['ASP', 'average selling price'], ['ASIN', 'Amazon Standard Identification Number'],
  ['FNSKU', 'Fulfillment Network Stock Keeping Unit (Amazon\'s barcode)'], ['FBA', 'Fulfillment by Amazon'], ['SKU', 'stock keeping unit'], ['UPC', 'Universal Product Code (the retail barcode; a GTIN exemption replaces it for a brand without barcodes)'],
  ['ACOS', 'advertising cost of sales (ad spend divided by ad attributed sales)'], ['TACOS', 'total advertising cost of sales (ad spend divided by total sales)'], ['PPC', 'pay per click advertising'],
  ['ATC', 'HDMI Authorized Test Center'], ['HDMI LA', 'HDMI Licensing Administrator'], ['FCC', 'Federal Communications Commission'], ['Prop 65', 'California Proposition 65 warning label'],
  ['CL3', 'in wall cable fire rating (Class 3)'], ['eARC', 'enhanced audio return channel'], ['Gbps', 'gigabits per second'], ['HDR', 'high dynamic range video'],
  ['3PL', 'third party logistics warehouse'], ['DTC', 'direct to consumer website'], ['QC', 'quality control'], ['VA', 'virtual assistant'], ['FOB', 'free on board (the price at the factory\'s port)'],
  ['COGS', 'cost of goods sold'], ['KPI', 'key performance indicator'], ['SOP', 'standard operating procedure'], ['Vine', 'Amazon\'s program that sends free units to selected reviewers'],
  ['FX', 'foreign exchange rate (KRW per USD; an input cell in the model)'], ['GTIN', 'Global Trade Item Number (the family of retail barcodes that includes UPC; Amazon grants a GTIN exemption to a brand without barcodes)'],
  ['QR', 'quick response code (the square barcode a phone camera reads)'], ['LLC', 'limited liability company'], ['FAQ', 'frequently asked questions'], ['B2B', 'business to business (installer and AV integrator sales)'],
];
children.push(gridTable(['Term', 'Meaning'], [1800, 7848], gl));

// ---------- document ----------
const doc = new Document({
  creator: 'BRAND US Launch build',
  title: '09 Amazon Listing Kit',
  description: 'Amazon US listing kit for BRAND fiber optic HDMI cables: ASIN structure, titles, bullets, A+ outline, keywords, image brief, FAQ, compliance, guarantee, prices',
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
    headers: { default: new Header({ children: [new Paragraph({ children: [run('BRAND LLC. 09 Amazon Listing Kit. Built ' + built + '.', { size: 16, color: MUTED })], alignment: AlignmentType.RIGHT })] }) },
    footers: { default: new Footer({ children: [new Paragraph({ children: [run('Page ', { size: 16, color: MUTED }), new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 16, color: MUTED }), run(' of ', { size: 16, color: MUTED }), new TextRun({ children: [PageNumber.TOTAL_PAGES], font: FONT, size: 16, color: MUTED })], alignment: AlignmentType.CENTER })] }) },
    children,
  }],
});

Packer.toBuffer(doc).then((buf) => {
  const out = path.join(OUT, '09_Amazon_Listing_Kit.docx');
  fs.writeFileSync(out, buf);
  const summary = {
    file: out, bytes: buf.length,
    titles: Object.fromEntries(Object.entries(titles).map(([k, v]) => [k, v.length])),
    keyword_bytes: { '8K': kwBytes(KW8K), '4K': kwBytes(KW4K) },
    children: { parentA: sku8k.length, parentB: sku4k.length },
  };
  fs.writeFileSync(path.join(BUILD, 'qa', 'listing', 'listing_summary.json'), JSON.stringify(summary, null, 1));
  console.log(JSON.stringify(summary));
});

#!/usr/bin/env node
// Diagnostic spike: isolate WHY Word kashida stretches some cells but not others.
// Four one-row, two-cell RTL tables (both cells identical per table):
//   (A) default font + mediumKashida + break        (reproduces marsiya "no stretch"?)
//   (B) default font + mediumKashida + break + w:lang bidi ar-SA  (does language enable kashida?)
//   (C) Scheherazade New + mediumKashida + break     (control that DID stretch)
//   (D) default font + mediumKashida + break + SHRUNK paragraph mark (does the empty line shrink?)
// Open test-documents/kashida-spike.docx in Word; report which stretch + which empty line is small.
// Run: node scripts/make-kashida-spike-doc.mjs

import { createRequire } from "module";
import { fileURLToPath } from "url";
import path from "path";
import fs from "fs";

const require = createRequire(import.meta.url);
const AdmZip = require("adm-zip");

const CELL_W = 4680;

function escXml(s) {
  return String(s || "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

// opts: { jc, withBreak, font (cs name|null), lang (bool), shrinkMark (bool) }
function cellPara(text, o) {
  var runFont = o.font ? '<w:rFonts w:cs="' + o.font + '" w:ascii="' + o.font + '" w:hAnsi="' + o.font + '"/>' : '';
  var lang = o.lang ? '<w:lang w:bidi="ar-SA"/>' : '';
  var brk = o.withBreak
    ? '<w:r><w:rPr><w:sz w:val="4"/><w:szCs w:val="4"/></w:rPr><w:br/></w:r>'
    : '';
  // Optionally shrink the paragraph mark so the trailing empty line is tiny.
  var markPr = o.shrinkMark ? '<w:rPr><w:sz w:val="4"/><w:szCs w:val="4"/></w:rPr>' : '';
  return '<w:p><w:pPr><w:bidi/>' + markPr + '<w:jc w:val="' + o.jc + '"/></w:pPr>' +
    '<w:r><w:rPr><w:rtl/>' + runFont + lang +
    '<w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr>' +
    '<w:t xml:space="preserve">' + escXml(text) + '</w:t></w:r>' + brk + '</w:p>';
}

function cell(paraXml, w) {
  return '<w:tc><w:tcPr><w:tcW w:w="' + (w || CELL_W) + '" w:type="dxa"/></w:tcPr>' + paraXml + '</w:tc>';
}
function thinBorder() {
  var b = 'w:val="single" w:sz="4" w:space="0" w:color="auto"';
  return '<w:tblBorders><w:top ' + b + '/><w:left ' + b + '/><w:bottom ' + b + '/>' +
    '<w:right ' + b + '/><w:insideH ' + b + '/><w:insideV ' + b + '/></w:tblBorders>';
}
function baytTable(o) {
  return '<w:tbl><w:tblPr><w:bidiVisual/><w:tblW w:w="' + (CELL_W * 2) + '" w:type="dxa"/>' +
    thinBorder() + '</w:tblPr>' +
    '<w:tblGrid><w:gridCol w:w="' + CELL_W + '"/><w:gridCol w:w="' + CELL_W + '"/></w:tblGrid>' +
    '<w:tr>' + cell(cellPara(misra, o)) + cell(cellPara(misra, o)) + '</w:tr></w:tbl>';
}
function heading(text, level) {
  var sz = level === 1 ? "28" : "22";
  return '<w:p><w:pPr><w:spacing w:before="240" w:after="80"/></w:pPr>' +
    '<w:r><w:rPr><w:b/><w:sz w:val="' + sz + '"/><w:szCs w:val="' + sz + '"/></w:rPr>' +
    '<w:t>' + escXml(text) + '</w:t></w:r></w:p>';
}
function note(text) {
  return '<w:p><w:pPr><w:ind w:left="360"/><w:spacing w:after="120"/></w:pPr>' +
    '<w:r><w:rPr><w:i/><w:color w:val="5F666D"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr>' +
    '<w:t>' + escXml(text) + '</w:t></w:r></w:p>';
}
function blankPara() { return '<w:p/>'; }

const misra = "العلم نور والجهل ظلام";
// Realistic marsiya 3-misra row (from the corpus): each goes in its own narrow column.
const m1 = "شاه كے اصحاب تھے";
const m2 = "خلق ميں الباب تھے";
const m3 = "صدق كے ارباب تھے";
const COL3 = 3120; // ~1/3 of a 9360-twip text width — a real marsiya column width

// Three-column RTL row, each cell independently justified with opts o.
function threeColTable(o) {
  return '<w:tbl><w:tblPr><w:bidiVisual/><w:tblW w:w="' + (COL3 * 3) + '" w:type="dxa"/>' +
    thinBorder() + '</w:tblPr>' +
    '<w:tblGrid><w:gridCol w:w="' + COL3 + '"/><w:gridCol w:w="' + COL3 + '"/><w:gridCol w:w="' + COL3 + '"/></w:tblGrid>' +
    '<w:tr>' + cell(cellPara(m1, o), COL3) + cell(cellPara(m2, o), COL3) + cell(cellPara(m3, o), COL3) + '</w:tr></w:tbl>';
}

// Bayt with a small middle gap column: sadr (wide) | gap (narrow) | ajuz (wide).
// Mixed cell widths — this is the asymmetric case (5+2+5-style).
const SADR_W = 4000, GAP_W = 1360, AJUZ_W = 4000; // total ~9360
const sadr = "إذا غامرت في شرف مروم";
const ajuz = "فلا تقنع بما دون النجوم";
function baytGapTable(o) {
  var gapCell = '<w:tc><w:tcPr><w:tcW w:w="' + GAP_W + '" w:type="dxa"/></w:tcPr><w:p><w:pPr><w:bidi/></w:pPr></w:p></w:tc>';
  return '<w:tbl><w:tblPr><w:bidiVisual/><w:tblW w:w="' + (SADR_W + GAP_W + AJUZ_W) + '" w:type="dxa"/>' +
    thinBorder() + '</w:tblPr>' +
    '<w:tblGrid><w:gridCol w:w="' + SADR_W + '"/><w:gridCol w:w="' + GAP_W + '"/><w:gridCol w:w="' + AJUZ_W + '"/></w:tblGrid>' +
    '<w:tr>' + cell(cellPara(sadr, o), SADR_W) + gapCell + cell(cellPara(ajuz, o), AJUZ_W) + '</w:tr></w:tbl>';
}

// Full marsiya grid: 1-2-3 / 4-5-6 / -7- / -8- (3 equal cols; solo lines centered in the middle col).
const g4 = "اهلِ بيت تھے عجب", g5 = "اٗفقِ دعوت كے شٗہٗب", g6 = "هو گئے قربان سب";
const g7 = "هو گئے شہ پر فدا", g8 = "هائے كربلاء والو";
function gridCell(text, o) { return cell(cellPara(text, o), COL3); }
function emptyCell() {
  return '<w:tc><w:tcPr><w:tcW w:w="' + COL3 + '" w:type="dxa"/></w:tcPr><w:p><w:pPr><w:bidi/></w:pPr></w:p></w:tc>';
}
function marsiyaGrid(o) {
  var gc = '<w:gridCol w:w="' + COL3 + '"/>';
  return '<w:tbl><w:tblPr><w:bidiVisual/><w:tblW w:w="' + (COL3 * 3) + '" w:type="dxa"/>' +
    thinBorder() + '</w:tblPr>' +
    '<w:tblGrid>' + gc + gc + gc + '</w:tblGrid>' +
    '<w:tr>' + gridCell(m1, o) + gridCell(m2, o) + gridCell(m3, o) + '</w:tr>' +
    '<w:tr>' + gridCell(g4, o) + gridCell(g5, o) + gridCell(g6, o) + '</w:tr>' +
    '<w:tr>' + emptyCell() + gridCell(g7, o) + emptyCell() + '</w:tr>' +
    '<w:tr>' + emptyCell() + gridCell(g8, o) + emptyCell() + '</w:tr>' +
    '</w:tbl>';
}

const parts = [];
const add = (...xs) => parts.push(...xs);

add(heading("ROOT-CAUSE TEST — forcing a Latin w:cs font on Arabic", 1));
add(note("Hypothesis: the runtime bakes the doc's default LATIN font (Calibri) as w:cs on Arabic, which Word won't kashida. (X) forces w:cs=Calibri, (Y) forces w:cs=Arial — both should FAIL to stretch. (Z) no forced font (control) — should stretch."));
add(heading("(X) w:cs=\"Calibri\" + mediumKashida + break", 2));
add(baytTable({ jc: "mediumKashida", withBreak: true, font: "Calibri", lang: false, shrinkMark: true }));
add(blankPara());
add(heading("(Y) w:cs=\"Arial\" + mediumKashida + break", 2));
add(baytTable({ jc: "mediumKashida", withBreak: true, font: "Arial", lang: false, shrinkMark: true }));
add(blankPara());
add(heading("(Z) no forced font (control) — should stretch", 2));
add(baytTable({ jc: "mediumKashida", withBreak: true, font: null, lang: false, shrinkMark: true }));
add(blankPara());

add(heading("Kashida diagnostic — marsiya grid (1-2-3 / 4-5-6 / -7- / -8-)", 1));
add(note("Full marsiya-style grid: two 3-misra rows, then two centered solo lines. All cells mediumKashida + break + shrunk mark, default font. Do all filled cells stretch, incl. the solo lines in the middle column?"));
add(marsiyaGrid({ jc: "mediumKashida", withBreak: true, font: null, lang: false, shrinkMark: true }));
add(blankPara());

add(heading("Kashida diagnostic — bayt with small middle gap (5+2+5)", 1));
add(note("Asymmetric mixed-width row: wide sadr | narrow gap | wide ajuz. Do the wide cells kashida-stretch with a small empty middle column between them?"));

add(heading("(H) bayt · sadr | small gap | ajuz · mediumKashida + break + shrunk mark", 2));
add(note("Sadr and ajuz justified; middle gap column empty. Do both outer cells stretch to their edges?"));
add(baytGapTable({ jc: "mediumKashida", withBreak: true, font: null, lang: false, shrinkMark: true }));
add(blankPara());

add(heading("Reference — equal narrow 3-column (marsiya) layout", 1));
add(note("Real marsiya rows are 3 narrow columns (~1/3 page each). Earlier the wide 2-col cells all stretched; does kashida still stretch in NARROW columns where the misra nearly fills the cell?"));

add(heading("(E) 3 columns · default font · mediumKashida + break + shrunk mark", 2));
add(note("Matches a normal marsiya 3-misra row. Does each narrow cell kashida-stretch? Is the empty line small?"));
add(threeColTable({ jc: "mediumKashida", withBreak: true, font: null, lang: false, shrinkMark: true }));
add(blankPara());

add(heading("(F) 3 columns · default font · highKashida + break + shrunk mark", 2));
add(note("Same but highest kashida intensity — does more intensity help in narrow cells?"));
add(threeColTable({ jc: "highKashida", withBreak: true, font: null, lang: false, shrinkMark: true }));
add(blankPara());

add(heading("(G) 2-column control (wide cells) — mediumKashida + break + shrunk mark", 2));
add(note("Wide-cell control (like earlier A–D) for comparison."));
add(baytTable({ jc: "mediumKashida", withBreak: true, font: null, lang: false, shrinkMark: true }));
add(blankPara());

const wns = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const documentXml =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<w:document ' + wns + '><w:body>' + parts.join("") + '<w:sectPr/></w:body></w:document>';
const contentTypes =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
  '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
  '<Default Extension="xml" ContentType="application/xml"/>' +
  '<Override PartName="/word/document.xml"' +
  ' ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
  '</Types>';
const relsRoot =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '<Relationship Id="rId1"' +
  ' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"' +
  ' Target="word/document.xml"/></Relationships>';
const relsWord =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>';

const zip = new AdmZip();
zip.addFile("[Content_Types].xml", Buffer.from(contentTypes, "utf8"));
zip.addFile("_rels/.rels", Buffer.from(relsRoot, "utf8"));
zip.addFile("word/document.xml", Buffer.from(documentXml, "utf8"));
zip.addFile("word/_rels/document.xml.rels", Buffer.from(relsWord, "utf8"));

const dir = path.join(path.dirname(fileURLToPath(import.meta.url)), "..", "test-documents");
fs.mkdirSync(dir, { recursive: true });
const out = path.join(dir, "kashida-spike.docx");
zip.writeZip(out);
console.log("Created", out);

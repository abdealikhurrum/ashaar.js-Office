#!/usr/bin/env node
// Generates test-documents/adopt-test.docx — a document that looks like a user's
// own file: plain hand-built poetry tables (NOT Ashaar content controls) plus raw
// marked-up text in several separator styles. Use it to exercise:
//   • Adopt Existing Table (Phase 1)        → the tables below
//   • Separator flexibility / conversion (Phase 2) → the raw text blocks below
// Run: node scripts/make-adopt-test-doc.mjs

import { createRequire } from "module";
import { fileURLToPath } from "url";
import path from "path";
import fs from "fs";

const require = createRequire(import.meta.url);
const AdmZip = require("adm-zip");

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const OUT = path.join(__dirname, "..", "test-documents", "adopt-test.docx");
const TATWEEL = "ـ";

function esc(s) {
  return String(s == null ? "" : s)
    .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

// A run of Arabic-script text (rtl) or plain LTR text.
function run(text, rtl) {
  var rpr = rtl ? "<w:rPr><w:rtl/></w:rPr>" : "";
  return "<w:r>" + rpr + '<w:t xml:space="preserve">' + esc(text) + "</w:t></w:r>";
}

// Paragraph. align: left|right|center; rtl toggles bidi.
function para(text, opts) {
  opts = opts || {};
  var jc = opts.align || (opts.rtl ? "right" : "left");
  var ppr = "<w:pPr>" + (opts.rtl ? "<w:bidi/>" : "") +
    '<w:jc w:val="' + jc + '"/>' +
    (opts.spaceAfter != null ? '<w:spacing w:after="' + opts.spaceAfter + '"/>' : "") +
    "</w:pPr>";
  return "<w:p>" + ppr + run(text, !!opts.rtl) + "</w:p>";
}

// An instruction label (LTR, bold-ish via larger gap) for the tester.
function label(text) {
  return "<w:p><w:pPr><w:spacing w:before=\"200\" w:after=\"60\"/></w:pPr>" +
    '<w:r><w:rPr><w:b/><w:color w:val="1F6F68"/></w:rPr><w:t xml:space="preserve">' + esc(text) + "</w:t></w:r></w:p>";
}

// A paragraph whose hemistichs are separated by a real tab character.
function tabPara(left, right) {
  return "<w:p><w:pPr><w:bidi/></w:pPr>" +
    run(left, true) + "<w:r><w:tab/></w:r>" + run(right, true) + "</w:p>";
}

function cellPara(text, align) {
  return "<w:p><w:pPr><w:bidi/><w:jc w:val=\"" + align + "\"/></w:pPr>" + run(text, true) + "</w:p>";
}

// cell: { text, span?, align? }
function tc(cell, colWidthTwips) {
  var span = cell.span || 1;
  var align = cell.align || "center";
  return "<w:tc><w:tcPr>" +
    '<w:gridSpan w:val="' + span + '"/>' +
    '<w:tcW w:w="' + (span * colWidthTwips) + '" w:type="dxa"/>' +
    "</w:tcPr>" + cellPara(cell.text || "", align) + "</w:tc>";
}

function tr(cells, colWidthTwips) {
  return "<w:tr>" + cells.map(function (c) { return tc(c, colWidthTwips); }).join("") + "</w:tr>";
}

// table: rows (array of cell-arrays), gridCols count, colWidthTwips per grid column.
function table(rows, gridCols, colWidthTwips) {
  var thin = 'w:val="single" w:sz="4" w:space="0" w:color="808080"';
  var borders = "<w:tblBorders>" +
    "<w:top " + thin + "/><w:left " + thin + "/><w:bottom " + thin + "/>" +
    "<w:right " + thin + "/><w:insideH " + thin + "/><w:insideV " + thin + "/></w:tblBorders>";
  var tblPr = "<w:tblPr>" +
    '<w:tblW w:w="' + (gridCols * colWidthTwips) + '" w:type="dxa"/>' +
    '<w:jc w:val="center"/>' + borders + "<w:bidiVisual/></w:tblPr>";
  var grid = "<w:tblGrid>";
  for (var i = 0; i < gridCols; i++) grid += '<w:gridCol w:w="' + colWidthTwips + '"/>';
  grid += "</w:tblGrid>";
  var body = rows.map(function (cells) { return tr(cells, colWidthTwips); }).join("");
  return "<w:tbl>" + tblPr + grid + body + "</w:tbl>";
}

function spacer() { return "<w:p/>"; }

// ── Document body ───────────────────────────────────────────────────────────

var parts = [];

parts.push(para("Ashaar — Adopt & Convert test document", { align: "left", spaceAfter: 120 }));
parts.push(para("Tables below are plain Word tables (not Ashaar blocks). Click inside one and use “Adopt Existing Table”. The raw text blocks at the bottom are for “Load Selection” / paste + separator conversion.", { align: "left", spaceAfter: 200 }));

// TABLE 1 — plain couplet (2 cols)
parts.push(label("TABLE 1 — Plain couplet (2 columns). Adopt → balanced Ashaar block."));
parts.push(table([
  [{ text: "دل ناداں تجھے ہوا کیا ہے", align: "right" }, { text: "آخر اس درد کی دوا کیا ہے", align: "left" }],
  [{ text: "ہم ہیں مشتاق اور وہ بیزار", align: "right" }, { text: "یا الٰہی یہ ماجرا کیا ہے", align: "left" }]
], 2, 4680));
parts.push(spacer());

// TABLE 2 — mixed solo + couplet (2 grid cols; solos span both)
parts.push(label("TABLE 2 — Mixed solos + couplet. Solo rows span the full width; couplet splits."));
parts.push(table([
  [{ text: "یہ ایک تنہا مصرع ہے", span: 2, align: "center" }],
  [{ text: "میں بھی منہ میں زبان رکھتا ہوں", align: "right" }, { text: "کاش پوچھو کہ مدعا کیا ہے", align: "left" }],
  [{ text: "اور یہ دوسرا تنہا مصرع", span: 2, align: "center" }]
], 2, 4680));
parts.push(spacer());

// TABLE 3 — multi-misra (3 cols)
parts.push(label("TABLE 3 — Three-misra row (3 columns) → multi-misra line."));
parts.push(table([
  [{ text: "شاه كے اصحاب تھے", align: "right" }, { text: "خلق ميں الباب تھے", align: "center" }, { text: "صدق كے ارباب تھے", align: "left" }]
], 3, 3120));
parts.push(spacer());

// TABLE 4 — already justified (tatweels present → must be stripped on adopt)
parts.push(label("TABLE 4 — Already kashida-justified (has tatweels). Adopt should clean them."));
parts.push(table([
  [{ text: "دل نـ" + TATWEEL + "اداں تجھے ہـ" + TATWEEL + "وا کیا ہے", align: "right" },
   { text: "آخر اس درد کی دوا کیا ہے", align: "left" }]
], 2, 4680));
parts.push(spacer());

// TABLE 5 — explicit gap column (sadr / empty gap / ajuz)
parts.push(label("TABLE 5 — Couplet with an empty middle gap column (3 cols). Gap must be ignored."));
parts.push(table([
  [{ text: "خودی کو کر بلند اتنا", align: "right" }, { text: "", align: "center" }, { text: "کہ ہر تقدیر سے پہلے", align: "left" }]
], 3, 4380)); // note: gap col rendered same width here; engine ignores empty cells regardless
parts.push(spacer());

// ── Raw text blocks for separator conversion ───────────────────────────────

parts.push(label("TEXT A — Dash-separated. Select these two lines → Load Selection."));
parts.push(para("بیا کہ قصر امل سخت سست بنیاد است - بیار بادہ کہ بنیاد عمر بر باد است", { rtl: true }));
parts.push(para("غلام ہمت آنم کہ زیر چرخ کبود - ز ہر چہ رنگ تعلق پذیرد آزاد است", { rtl: true }));
parts.push(spacer());

parts.push(label("TEXT B — Tab-separated."));
parts.push(tabPara("على قدرِ أهلِ العزمِ تأتي العزائمُ", "وتأتي على قدرِ الكرامِ المكارمُ"));
parts.push(spacer());

parts.push(label("TEXT C — Asterisk-separated (m1 * m2)."));
parts.push(para("إذا غامَرتَ في شَرَفٍ مَرومِ * فَلا تَقنَع بما دونَ النُّجومِ", { rtl: true }));
parts.push(spacer());

parts.push(label("TEXT D — Wide gap (multiple spaces between hemistichs)."));
parts.push(para("بشنو این نی چون شکایت می‌کند     از جدایی‌ها حکایت می‌کند", { rtl: true }));
parts.push(spacer());

parts.push(label("TEXT E — One hemistich per line (use “Pair every 2 lines”)."));
parts.push(para("الٹی ہو گئیں سب تدبیریں کچھ نہ دوا نے کام کیا", { rtl: true }));
parts.push(para("دیکھا اس بیماریِ دل نے آخر کام تمام کیا", { rtl: true }));
parts.push(para("مت سہل ہمیں جانو پھرتا ہے فلک برسوں", { rtl: true }));
parts.push(para("تب خاک کے پردے سے انسان نکلتے ہیں", { rtl: true }));
parts.push(spacer());

parts.push(label("TEXT F — Hyphenated word (must NOT be split): far-flung, well-known."));
parts.push(para("یہ far-flung اور well-known مصرع ہے", { rtl: true }));

var bodyXml = parts.join("") +
  '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>';

var documentXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
  "<w:body>" + bodyXml + "</w:body></w:document>";

var contentTypes = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
  '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
  '<Default Extension="xml" ContentType="application/xml"/>' +
  '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
  "</Types>";

var rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
  "</Relationships>";

var docRels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';

var zip = new AdmZip();
zip.addFile("[Content_Types].xml", Buffer.from(contentTypes, "utf8"));
zip.addFile("_rels/.rels", Buffer.from(rels, "utf8"));
zip.addFile("word/document.xml", Buffer.from(documentXml, "utf8"));
zip.addFile("word/_rels/document.xml.rels", Buffer.from(docRels, "utf8"));
fs.mkdirSync(path.dirname(OUT), { recursive: true });
zip.writeZip(OUT);
console.log("Wrote " + OUT + " (" + fs.statSync(OUT).size + " bytes)");

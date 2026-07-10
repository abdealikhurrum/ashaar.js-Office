#!/usr/bin/env node
// Verification spike: proves the "trailing soft break + mediumKashida" trick works in Word.
// Generates three one-row, two-cell RTL tables demonstrating:
//   (a) mediumKashida + trailing <w:br/> (expected: kashida stretch + hidden break)
//   (b) mediumKashida without break (control: no stretch)
//   (c) distribute (control: spacing fill)
// Open test-documents/kashida-spike.docx in Word and observe which cells stretch.
// Run: node scripts/make-kashida-spike-doc.mjs

import { createRequire } from "module";
import { fileURLToPath } from "url";
import path from "path";
import fs from "fs";

const require = createRequire(import.meta.url);
const AdmZip = require("adm-zip");

const CELL_W = 4680; // twips (~2.5-3.25")

function escXml(s) {
  return String(s || "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

// Cell paragraph: jc="mediumKashida"|"distribute"; withBreak=boolean.
// Trailing <w:br/> is shrunk to 4pt so the empty line is barely visible.
function cellPara(text, jc, withBreak) {
  var brk = withBreak
    ? '<w:r><w:rPr><w:sz w:val="4"/><w:szCs w:val="4"/></w:rPr><w:br/></w:r>'
    : '';
  return '<w:p><w:pPr><w:bidi/><w:jc w:val="' + jc + '"/></w:pPr>' +
    '<w:r><w:rPr><w:rtl/><w:rFonts w:cs="Scheherazade New"/>' +
    '<w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr>' +
    '<w:t xml:space="preserve">' + escXml(text) + '</w:t></w:r>' + brk + '</w:p>';
}

// A cell with a single paragraph.
function cell(paraXml) {
  return "<w:tc><w:tcPr><w:tcW w:w=\"" + CELL_W + "\" w:type=\"dxa\"/></w:tcPr>" +
    paraXml + "</w:tc>";
}

// Cell paragraph with no explicit jc (default).
function cellParaDefault(text) {
  return '<w:p><w:pPr><w:bidi/></w:pPr>' +
    '<w:r><w:rPr><w:rtl/><w:rFonts w:cs="Scheherazade New"/>' +
    '<w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr>' +
    '<w:t xml:space="preserve">' + escXml(text) + '</w:t></w:r></w:p>';
}

function thinBorder() {
  var b = 'w:val="single" w:sz="4" w:space="0" w:color="auto"';
  return "<w:tblBorders><w:top " + b + "/><w:left " + b + "/><w:bottom " + b + "/>" +
    "<w:right " + b + "/><w:insideH " + b + "/><w:insideV " + b + "/></w:tblBorders>";
}

// Single-row, two-cell RTL table.
function baytTable(para1, para2) {
  return "<w:tbl><w:tblPr><w:bidiVisual/>" +
    "<w:tblW w:w=\"" + (CELL_W * 2) + "\" w:type=\"dxa\"/>" + thinBorder() + "</w:tblPr>" +
    "<w:tblGrid><w:gridCol w:w=\"" + CELL_W + "\"/><w:gridCol w:w=\"" + CELL_W + "\"/></w:tblGrid>" +
    "<w:tr>" + cell(para1) + cell(para2) + "</w:tr></w:tbl>";
}

function heading(text, level) {
  var sz = level === 1 ? "28" : "22";
  return "<w:p><w:pPr><w:spacing w:before=\"240\" w:after=\"80\"/></w:pPr>" +
    "<w:r><w:rPr><w:b/><w:sz w:val=\"" + sz + "\"/><w:szCs w:val=\"" + sz + "\"/></w:rPr>" +
    "<w:t>" + escXml(text) + "</w:t></w:r></w:p>";
}

function note(text) {
  return "<w:p><w:pPr><w:ind w:left=\"360\"/><w:spacing w:after=\"120\"/></w:pPr>" +
    "<w:r><w:rPr><w:i/><w:color w:val=\"5F666D\"/><w:sz w:val=\"18\"/><w:szCs w:val=\"18\"/></w:rPr>" +
    "<w:t>" + escXml(text) + "</w:t></w:r></w:p>";
}

function blankPara() { return "<w:p/>"; }

const parts = [];
const add = (...xs) => parts.push(...xs);

add(heading("Kashida + Trailing Break Verification Spike", 1));
add(note("Three tables below demonstrate the kashida + trailing-break trick. " +
  "Open in Word and observe which cells stretch. " +
  "(a) should kashida-stretch; (b) should NOT; (c) should space-fill."));

const misra = "العلم نور والجهل ظلام";

// Table (a): mediumKashida + trailing break — BOTH cells identical.
add(heading("(a) mediumKashida + trailing break — both cells", 2));
add(note("Expected: BOTH cells' letters stretch uniformly to fill; trailing empty line barely visible (2pt)."));
add(baytTable(
  cellPara(misra, "mediumKashida", true),
  cellPara(misra, "mediumKashida", true)
));

add(blankPara());

// Table (b): mediumKashida without break (control) — BOTH cells identical.
add(heading("(b) mediumKashida (no break, control) — both cells", 2));
add(note("Expected: NEITHER cell stretches (confirms the break is what enables (a))."));
add(baytTable(
  cellPara(misra, "mediumKashida", false),
  cellPara(misra, "mediumKashida", false)
));

add(blankPara());

// Table (c): distribute (control) — BOTH cells identical.
add(heading("(c) distribute (spacing fill, control) — both cells", 2));
add(note("Expected: BOTH cells fill via word/letter spacing, not kashida."));
add(baytTable(
  cellPara(misra, "distribute", false),
  cellPara(misra, "distribute", false)
));

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

#!/usr/bin/env node
// Generates test-documents/mixed-style-test.docx: RTL poetry tables whose cells
// contain MIXED style runs, to exercise run-aware justification (justifyRuns /
// computeRunSpacing). Two motivating cases:
//   1. Larger first word   — first word at a bigger font size, rest normal.
//   2. Bold refrain word    — one word bold, rest normal.
// Select a whole table in Word, then click "Justify Selected Text" (Kashida or
// Spacing). Expect: the bold / larger word keeps its styling, tatweels/spacing
// fill the cell, and re-justify does not compound.
// Run: node scripts/make-mixed-style-doc.mjs

import { createRequire } from "module";
import { fileURLToPath } from "url";
import path from "path";
import fs from "fs";

const require = createRequire(import.meta.url);
const AdmZip = require("adm-zip");

const CELL_W = 4680; // twips (~3.25")

function escXml(s) {
  return String(s || "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

// One styled run. size in half-points (sz), bold optional. cs font keeps Arabic
// shaping consistent across runs so the only differences are size / weight.
function run(text, size, bold) {
  return "<w:r><w:rPr><w:rtl/>" +
    "<w:rFonts w:cs=\"Scheherazade New\" w:ascii=\"Scheherazade New\" w:hAnsi=\"Scheherazade New\"/>" +
    (bold ? "<w:b/><w:bCs/>" : "") +
    "<w:sz w:val=\"" + size + "\"/><w:szCs w:val=\"" + size + "\"/>" +
    "</w:rPr><w:t xml:space=\"preserve\">" + escXml(text) + "</w:t></w:r>";
}

// A cell paragraph: RTL, justified both edges (so spacing/kashida have a target).
function cell(runsXml) {
  return "<w:tc><w:tcPr><w:tcW w:w=\"" + CELL_W + "\" w:type=\"dxa\"/></w:tcPr>" +
    "<w:p><w:pPr><w:bidi/><w:jc w:val=\"both\"/></w:pPr>" + runsXml + "</w:p></w:tc>";
}

function thinBorder() {
  var b = 'w:val="single" w:sz="4" w:space="0" w:color="auto"';
  return "<w:tblBorders><w:top " + b + "/><w:left " + b + "/><w:bottom " + b + "/>" +
    "<w:right " + b + "/><w:insideH " + b + "/><w:insideV " + b + "/></w:tblBorders>";
}

// A single-row, two-cell (sadr | ajuz) RTL table.
function baytTable(sadrRuns, ajuzRuns) {
  return "<w:tbl><w:tblPr><w:bidiVisual/>" +
    "<w:tblW w:w=\"" + (CELL_W * 2) + "\" w:type=\"dxa\"/>" + thinBorder() + "</w:tblPr>" +
    "<w:tblGrid><w:gridCol w:w=\"" + CELL_W + "\"/><w:gridCol w:w=\"" + CELL_W + "\"/></w:tblGrid>" +
    "<w:tr>" + cell(sadrRuns) + cell(ajuzRuns) + "</w:tr></w:tbl>";
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

add(heading("Run-Aware Justification — Mixed-Style Test", 1));
add(note("Select a whole table (drag across it or use the table move-handle), then click " +
  "“Justify Selected Text” in the pane. Try both Kashida and Spacing modes. " +
  "The bold / larger word must keep its styling; the line should fill the cell; " +
  "re-justifying must not compound."));

// Case 1 — larger first word (size varies run-to-run; same family/weight).
add(heading("Case 1 — larger first word (size varies)", 2));
add(baytTable(
  run("العلمُ", 36) + run(" نورٌ مبينٌ", 24),   // "العلمُ" large + " نورٌ مبينٌ"
  run("والجهلُ", 36) + run(" ظلامٌ مهينٌ", 24) // "والجهلُ" large + " ظلامٌ مهينٌ"
));

add(blankPara());

// Case 2 — bold refrain word (weight varies; same family/size).
add(heading("Case 2 — bold word (weight varies)", 2));
add(baytTable(
  run("هائٔے ", 28) + run("كربلاء", 28, true) + run(" والو", 28), // "هائے" + bold "كربلاء" + "والو"
  run("شاه كے ", 28) + run("اصحاب", 28, true) + run(" تھے", 28)             // "شاہ کے" + bold "اصحاب" + "تھے"
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
const out = path.join(dir, "mixed-style-test.docx");
zip.writeZip(out);
console.log("Created", out);

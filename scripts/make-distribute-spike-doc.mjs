#!/usr/bin/env node
// Spike: does Word's Distributed justification (w:jc="distribute") realize the
// spacing residual we currently inject as hair-spaces? For an RTL Arabic/Urdu
// misra in a cell WIDER than its natural width, does distribute stretch the
// WORD GAPS (good — equals CSS word-spacing) or add INTER-LETTER spacing (bad)?
// And does it compose with concentrated tatweels (elongation:spacing ratio)?
//
// Open test-documents/distribute-spike.docx in Word and eyeball each labelled cell.
// Run: node scripts/make-distribute-spike-doc.mjs

import { createRequire } from "module";
import { fileURLToPath } from "url";
import path from "path";
import fs from "fs";

const require = createRequire(import.meta.url);
const AdmZip = require("adm-zip");
const AshaarJustify = require("../src/vendor/ashaar-justify.js");

// A wide cell so there's clear room to stretch (~4.2in), and a large font.
const CELL_W = 6000;          // twips
const SZ = 32;                // half-points = 16pt

function escXml(s) {
  return String(s || "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

// Faithful concentrated tatweels via the real engine. Mock measure = 10px/char so
// 1 tatweel ≈ 10px; perPositionEm 1.5 → cap ≈ 32px ≈ 3 tatweels/join, so a budget
// spreads across a few best joins (multi per join) exactly like the real path.
function tatweel(text, budgetPx, maxPositions) {
  var runs = [{ text: text, fontSize: 16, fontProfile: null, measure: function (s) { return s.length * 10; } }];
  var natural = text.length * 10;
  var out = AshaarJustify.justifyRunsConcentrated(runs, natural + budgetPx, {
    perPositionEm: 1.5, maxPositions: maxPositions || 0
  });
  return out.runs[0].text;
}

// One paragraph: RTL Arabic run, given jc.
function cellPara(text, jc) {
  return '<w:p><w:pPr><w:bidi/><w:jc w:val="' + jc + '"/></w:pPr>' +
    '<w:r><w:rPr><w:rtl/><w:sz w:val="' + SZ + '"/><w:szCs w:val="' + SZ + '"/></w:rPr>' +
    '<w:t xml:space="preserve">' + escXml(text) + '</w:t></w:r></w:p>';
}

function cell(paraXml, w) {
  return '<w:tc><w:tcPr><w:tcW w:w="' + (w || CELL_W) + '" w:type="dxa"/></w:tcPr>' + paraXml + '</w:tc>';
}
function thinBorder() {
  var b = 'w:val="single" w:sz="4" w:space="0" w:color="auto"';
  return '<w:tblBorders><w:top ' + b + '/><w:left ' + b + '/><w:bottom ' + b + '/>' +
    '<w:right ' + b + '/><w:insideH ' + b + '/><w:insideV ' + b + '/></w:tblBorders>';
}
// One-cell, fixed-width RTL table so the cell edges (fill target) are visible.
function oneCell(text, jc) {
  return '<w:tbl><w:tblPr><w:bidiVisual/><w:tblW w:w="' + CELL_W + '" w:type="dxa"/>' +
    thinBorder() + '</w:tblPr><w:tblGrid><w:gridCol w:w="' + CELL_W + '"/></w:tblGrid>' +
    '<w:tr>' + cell(cellPara(text, jc)) + '</w:tr></w:tbl>';
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
function blank() { return '<w:p/>'; }

// Ghalib misra — long connected words (تحریر / فریادی) make inter-letter stretch obvious.
const MISRA = "نقش فریادی ہے کس کی شوخئ تحریر کا";
const HAIR = " ";
// Current approach: hair-spaces at word gaps (word-spacing equivalent) — reference.
function hairFill(text, n) {
  var parts = text.split(" ");
  var gaps = parts.length - 1;
  if (gaps <= 0 || n <= 0) return text;
  var per = Math.floor(n / gaps), rem = n % gaps, out = parts[0];
  for (var i = 1; i < parts.length; i++) {
    var add = per + (i <= rem ? 1 : 0);
    out += " " + new Array(add + 1).join(HAIR) + parts[i];
  }
  return out;
}

const parts = [];
const add = (...xs) => parts.push(...xs);

add(heading("SPIKE — Word Distributed justification (w:jc=\"distribute\") on RTL Arabic", 1));
add(note("Every cell below is the SAME " + CELL_W + "-twip width. The question: with jc=distribute, does the misra fill the cell by widening the WORD GAPS (what we want — same as CSS word-spacing) or by adding space BETWEEN LETTERS within words (bad — words look gappy)? Compare the connected words نقش / فریادی / تحریر across the cells."));

add(heading("A) jc=distribute — NO tatweels (this is the strength-1 / all-spacing case)", 2));
add(note("KEY CELL. Does the line reach both edges by spreading the word gaps, with letters INSIDE each word untouched? If yes, distribute == word-spacing and the whole approach works."));
add(oneCell(MISRA, "distribute"));
add(blank());

add(heading("B) jc=both (Justify) — NO tatweels — comparison", 2));
add(note("Word 'Justify'. A single/last line often does NOT stretch under 'both'. If this cell sits at natural width (short of the left edge) while (A) fills, that's exactly why we need 'distribute'."));
add(oneCell(MISRA, "both"));
add(blank());

add(heading("C) jc=right — NO tatweels — natural-width baseline", 2));
add(note("No fill at all — the misra's natural width. Use it to see how far (A) and (B) stretched."));
add(oneCell(MISRA, "right"));
add(blank());

add(heading("Distribute composing with concentrated tatweels (elongation:spacing ratio)", 1));
add(note("Same cell width; the engine pre-inserts concentrated tatweels (the elongation portion) and jc=distribute fills the rest. As tatweels increase, Word should add LESS distributed spacing — the ratio lever."));

add(heading("D) distribute + FEW tatweels (low strength) — mostly spacing", 2));
add(oneCell(tatweel(MISRA, 40, 2), "distribute"));
add(blank());
add(heading("E) distribute + MORE tatweels (mid strength) — mix", 2));
add(oneCell(tatweel(MISRA, 100, 0), "distribute"));
add(blank());
add(heading("F) distribute + MANY tatweels (high strength) — mostly elongation", 2));
add(oneCell(tatweel(MISRA, 200, 0), "distribute"));
add(blank());

add(heading("Current vs proposed — same fill target", 1));
add(heading("G) CURRENT: jc=right + injected hair-spaces at word gaps", 2));
add(note("Today's mechanism (hair-space injection). Compare its word gaps to (A)'s distribute gaps — they should look the same if distribute == word-spacing."));
add(oneCell(hairFill(MISRA, 24), "right"));
add(blank());
add(heading("H) PROPOSED: jc=distribute + few tatweels", 2));
add(oneCell(tatweel(MISRA, 60, 3), "distribute"));
add(blank());

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
const out = path.join(dir, "distribute-spike.docx");
zip.writeZip(out);
console.log("Created", out);
console.log("A(no-tat):", MISRA);
console.log("D(few):  ", tatweel(MISRA, 40, 2));
console.log("F(many): ", tatweel(MISRA, 200, 0));

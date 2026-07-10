#!/usr/bin/env node
// Generates test-documents/nastaliq-sample.docx: one poetry block per font so you
// can exercise each kashida mechanism. Select a block's table, pick the matching
// font + Kashida in the pane, and Justify:
//   Mehr    -> tatweel  (discrete trailing tatweels at whitelisted joins)
//   Jameel  -> font-swap (swapped fasls render in the wider Kasheeda face)
//   Gulzar / Noto / Arabic-serif -> whitespace (spacing fill)
// Word justify ("css") re-renders any block with native Word kashida.
// Run: node scripts/make-nastaliq-sample-doc.mjs

import { createRequire } from "module";
import { fileURLToPath } from "url";
import path from "path";
import fs from "fs";

const require = createRequire(import.meta.url);
const AdmZip = require("adm-zip");
const Ashaar = require("../src/vendor/ashaar.js");
const AshaarWord = require("../src/taskpane/word-html.js");

const TEXT_WIDTH_TWIPS = 9360; // 6.5"

// ── Poetry ────────────────────────────────────────────────────────────────
const GHALIB = "دل ناداں تجھے ہوا کیا ہے \\ آخر اس درد کی دوا کیا ہے";
const IQBAL  = "خودی کو کر بلند اتنا کہ ہر تقدیر سے پہلے \\ خدا بندے سے خود پوچھے بتا تیری رضا کیا ہے";
const MARSIYA = [
  "شاه كے اصحاب تھے \\ خلق ميں الباب تھے \\ صدق كے ارباب تھے \\",
  "هو گئے شہ پر فدا \\",
  "هائے كربلاء والو \\ هائے كربلاء والو"
].join("\n");

// ── OOXML helpers ───────────────────────────────────────────────────────────
function escXml(s) {
  return String(s || "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}
function heading(text, level) {
  var sz = level === 1 ? "30" : "24";
  var color = level === 1 ? '<w:color w:val="1F6F68"/>' : "";
  return "<w:p><w:pPr><w:spacing w:before=\"280\" w:after=\"100\"/></w:pPr>" +
    "<w:r><w:rPr><w:b/>" + color + "<w:sz w:val=\"" + sz + "\"/><w:szCs w:val=\"" + sz + "\"/></w:rPr>" +
    "<w:t>" + escXml(text) + "</w:t></w:r></w:p>";
}
function note(text) {
  return "<w:p><w:pPr><w:ind w:left=\"360\"/><w:spacing w:after=\"120\"/></w:pPr>" +
    "<w:r><w:rPr><w:i/><w:color w:val=\"5F666D\"/><w:sz w:val=\"20\"/><w:szCs w:val=\"20\"/></w:rPr>" +
    "<w:t>" + escXml(text) + "</w:t></w:r></w:p>";
}
function blankPara() { return "<w:p><w:pPr><w:spacing w:after=\"80\"/></w:pPr></w:p>"; }

// A poem block rendered unjustified in a given font, ready for the user to
// select + Justify. Larger size so the kashida is easy to see.
function poem(source, fontMode) {
  return AshaarWord.renderForWordOoxml(source, {
    justifyMode: "none", gapWidth: 1, fontMode: fontMode
  }, Ashaar, TEXT_WIDTH_TWIPS);
}

const parts = [];
const add = (...xs) => parts.push(...xs);

add(heading("Nastaliq Fonts & Kashida — Sample", 1));
add(note("Select a block's table, choose the matching Font + Kashida in the pane, and click " +
  "Justify Selected Text. Each font uses a different kashida mechanism (see each heading). " +
  "Or pick “Word justify” on any block to re-render it with native Word kashida."));

add(heading("Mehr Nastaliq — tatweel mechanism", 2));
add(note("Font → Mehr Nastaliq, Kashida. Expect discrete trailing tatweels at whitelisted joins."));
add(poem(GHALIB, "mehr"));
add(blankPara());

add(heading("Jameel Noori Kasheeda — font-swap mechanism", 2));
add(note("Font → Jameel Noori Kasheeda, Kashida. Expect swapped fasls in the wider Kasheeda face."));
add(poem(GHALIB, "jameel"));
add(blankPara());

add(heading("Gulzar — whitespace (spacing) mechanism", 2));
add(note("Font → Gulzar, Kashida. No designed tatweel, so it fills with spacing."));
add(poem(GHALIB, "gulzar"));
add(blankPara());

add(heading("Noto Nastaliq Urdu — whitespace", 2));
add(note("Font → Noto Nastaliq Urdu. Longer Iqbal couplet — good for testing no-wrap auto-widen."));
add(poem(IQBAL, "noto"));
add(blankPara());

add(heading("Arabic serif (Scheherazade) — Ashaar.js engine kashida", 2));
add(note("Font → Arabic serif, Kashida + strength. The tatweel engine with the expressive 15–24 regime."));
add(poem(GHALIB, "arabic-serif"));
add(blankPara());

add(heading("Marsiya grid (Mehr) — 3-misra + solo + refrain", 2));
add(note("Multi-row band. Justify with Mehr (tatweel) or Word justify; solo/refrain sit at misra width."));
add(poem(MARSIYA, "mehr"));
add(blankPara());

// ── Assemble docx ─────────────────────────────────────────────────────────
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
const out = path.join(dir, "nastaliq-sample.docx");
zip.writeZip(out);
console.log("Created", out);

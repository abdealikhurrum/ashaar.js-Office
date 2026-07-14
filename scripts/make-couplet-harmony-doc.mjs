#!/usr/bin/env node
// Couplet-equalization (harmony) verification doc — Ghalib couplets rendered
// UN-justified inside real v3-tagged Ashaar Poem content controls, so the
// settings panel manages them and Apply exercises the PRODUCTION fill-mode
// paths (Natural-fit / Cell-fit / Adaptive) in Word.
//
// Repro target: the open report "equalization did not work in couplet mode"
// (single-table multi-bayt block) — check whether the harmony matrix spans
// all rows of a single-stanza table.
//
// Open test-documents/couplet-harmony-test.docx in desktop Word, then per block:
//   1. Click inside → panel header shows "Poem".
//   2. Justification: Kashida, Fill mode: Natural-fit → Apply.
//      PASS: every first hemistich ends flush down column 1, every second
//      hemistich flush down column 2 (each column justified to its own
//      longest natural line). FAIL: ragged column edges after Apply.
//   3. Fill mode: Cell-fit → Apply. PASS: every cell filled to its edge.
//   4. If ragged: enable "Capture justification metrics" (Debug), Apply
//      again, and save the dump — it discriminates whether the harmony
//      matrix covered all rows.
// Block A = one table (the reported failure shape). Block C = same couplets
// split across two tables — if A fails and C passes, the bug is confirmed
// to be single-table row coverage.
// Run: node scripts/make-couplet-harmony-doc.mjs

import { createRequire } from "module";
import { fileURLToPath } from "url";
import path from "path";
import fs from "fs";

const require = createRequire(import.meta.url);
const AdmZip     = require("adm-zip");
const Ashaar     = require("../src/vendor/ashaar.js");
const AshaarWord = require("../src/taskpane/word-html.js");

const TEXT_WIDTH_TWIPS = 9360; // 6.5" — standard text column

// ── Ghalib — دل ناداں (uniform short meter; realistic variance) ─────────────
const GHAZAL = [
  "دل ناداں تجھے ہوا کیا ہے \\ آخر اس درد کی دوا کیا ہے",
  "ہم ہیں مشتاق اور وہ بیزار \\ یا الٰہی یہ ماجرا کیا ہے",
  "میں بھی منہ میں زبان رکھتا ہوں \\ کاش پوچھو کہ مدعا کیا ہے",
  "جان تم پر نثار کرتا ہوں \\ میں نہیں جانتا دعا کیا ہے",
  "ہم کو ان سے وفا کی ہے امید \\ جو نہیں جانتے وفا کیا ہے",
].join("\n");

// ── Stress mix — couplets from different meters (exaggerated variance) ──────
// Long/short alternation makes a harmony failure impossible to miss: without
// column equalization the short couplets sit far from the long ones' edge.
const STRESS = [
  "ہزاروں خواہشیں ایسی کہ ہر خواہش پہ دم نکلے \\ بہت نکلے مرے ارمان لیکن پھر بھی کم نکلے",
  "دل ناداں تجھے ہوا کیا ہے \\ آخر اس درد کی دوا کیا ہے",
  "یہ نہ تھی ہماری قسمت کہ وصال یار ہوتا \\ اگر اور جیتے رہتے یہی انتظار ہوتا",
  "کوئی امید بر نہیں آتی \\ کوئی صورت نظر نہیں آتی",
].join("\n");

// Same couplets as GHAZAL but split into two stanzas → two tables.
const GHAZAL_SPLIT = GHAZAL.split("\n").slice(0, 2).join("\n") + "\n\n" +
                     GHAZAL.split("\n").slice(2).join("\n");

// ── OOXML helpers (packaging mirrors scripts/make-qasida-wordfill-doc.mjs) ──
function escXml(s) {
  return String(s || "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}
function escAttr(s) { return escXml(s).replace(/"/g, "&quot;"); }

function heading(text, level) {
  var before = level === 1 ? "360" : "220";
  var after  = level === 1 ? "120" : "60";
  var sz     = level === 1 ? "30"  : "24";
  var color  = level === 1 ? "<w:color w:val=\"1F6F68\"/>" : "<w:color w:val=\"A7352A\"/>";
  return "<w:p><w:pPr><w:spacing w:before=\"" + before + "\" w:after=\"" + after + "\"/></w:pPr>" +
    "<w:r><w:rPr><w:b/>" + color + "<w:sz w:val=\"" + sz + "\"/><w:szCs w:val=\"" + sz + "\"/></w:rPr>" +
    "<w:t>" + escXml(text) + "</w:t></w:r></w:p>";
}

function note(text) {
  return "<w:p><w:pPr><w:spacing w:before=\"40\" w:after=\"80\"/><w:ind w:left=\"200\"/></w:pPr>" +
    "<w:r><w:rPr><w:i/><w:color w:val=\"5F666D\"/><w:sz w:val=\"18\"/><w:szCs w:val=\"18\"/></w:rPr>" +
    "<w:t>" + escXml(text) + "</w:t></w:r></w:p>";
}

function blankPara() { return "<w:p><w:pPr><w:spacing w:after=\"120\"/></w:pPr></w:p>"; }

// Managed poem block: render un-justified through the production emitter and
// wrap in the SAME <w:sdt> shape the add-in's insert paths embed
// (word-html.js wrapOoxmlControl), with a real v3 tag so the panel manages it.
function managedPoem(source) {
  var opts = { justifyMode: "none", gapWidth: 4 };
  var pats = AshaarWord.poemCellPatterns(source, opts, Ashaar);
  var tag  = AshaarWord.contentControlTag(source, opts, pats);
  var body = AshaarWord.renderForWordOoxml(source, opts, Ashaar, TEXT_WIDTH_TWIPS);
  var idNum = 0;
  for (var i = 0; i < tag.length; i++) idNum = ((idNum << 5) - idNum + tag.charCodeAt(i)) | 0;
  idNum = (Math.abs(idNum) % 2000000000) || 1;
  return "<w:sdt><w:sdtPr>" +
    '<w:alias w:val="Ashaar Poem"/>' +
    '<w:tag w:val="' + escAttr(tag) + '"/>' +
    '<w:id w:val="' + idNum + '"/>' +
    "</w:sdtPr><w:sdtContent>" + body + "</w:sdtContent></w:sdt>";
}

// ── Build body ──────────────────────────────────────────────────────────────
const parts = [];
function add() { for (var i = 0; i < arguments.length; i++) parts.push(arguments[i]); }

add(heading("Couplet equalization (harmony) — manual verification", 1));
add(note("Each block below is a managed Ashaar Poem (un-justified). Click in → Kashida + Natural-fit → Apply: both columns must end flush (each column justified to its longest natural line). Then Cell-fit → Apply: every cell filled to its edge. If a column stays ragged, enable Debug metrics, Apply again, and save the dump."));

add(heading("A · Ghazal — ONE table, five couplets (the reported failure shape)", 2));
add(note("Ghalib, dil-e-nādāṉ. Uniform meter, realistic width variance. Natural-fit must equalize all five rows per column."));
add(managedPoem(GHAZAL));
add(blankPara());

add(heading("B · Stress mix — ONE table, alternating long/short couplets", 2));
add(note("Couplets from different meters. Exaggerated variance: a harmony failure is unmissable here — short couplets will sit far short of the long couplets' edge."));
add(managedPoem(STRESS));
add(blankPara());

add(heading("C · Same ghazal — split into TWO tables (2 + 3 couplets)", 2));
add(note("Control for block A: if A stays ragged but C equalizes, the harmony matrix is not spanning all rows of a single-stanza table — capture the debug dump on A."));
add(managedPoem(GHAZAL_SPLIT));
add(blankPara());

// ── Assemble docx ────────────────────────────────────────────────────────────
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
  ' Target="word/document.xml"/>' +
  '</Relationships>';

const relsWord =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>';

const zip = new AdmZip();
zip.addFile("[Content_Types].xml",          Buffer.from(contentTypes, "utf8"));
zip.addFile("_rels/.rels",                  Buffer.from(relsRoot,     "utf8"));
zip.addFile("word/document.xml",            Buffer.from(documentXml,  "utf8"));
zip.addFile("word/_rels/document.xml.rels", Buffer.from(relsWord,     "utf8"));

const dir = path.join(path.dirname(fileURLToPath(import.meta.url)), "..", "test-documents");
fs.mkdirSync(dir, { recursive: true });
const out = path.join(dir, "couplet-harmony-test.docx");
zip.writeZip(out);
console.log("Created", out);

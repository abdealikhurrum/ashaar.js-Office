#!/usr/bin/env node
// Generates test-documents/corpus-test.docx with the full poetry corpus:
// Arabic, Farsi, and Urdu across all structures and layout modes.
// Run: node scripts/make-test-doc.mjs  (or: npm run make-test-doc)

import { createRequire } from "module";
import { fileURLToPath } from "url";
import path from "path";
import fs from "fs";

const require = createRequire(import.meta.url);
const AdmZip   = require("adm-zip");
const Ashaar   = require("../src/vendor/ashaar.js");
const AshaarWord     = require("../src/taskpane/word-html.js");
const AshaarTabStop  = require("../src/taskpane/word-tabstop.js");

const TEXT_WIDTH_TWIPS = 9360; // 6.5" at 1440 twips/inch

// ── Poetry corpus ─────────────────────────────────────────────────────────────

const ARABIC_BAYT =
  "إذا غامَرتَ في شَرَفٍ مَرومِ \\ فَلا تَقنَع بما دونَ النُّجومِ";

const ARABIC_QASIDA = [
  "على قدرِ أهلِ العزمِ تأتي العزائمُ \\ وتأتي على قدرِ الكرامِ المكارمُ",
  "وتعظمُ في عينِ الصغيرِ صغارُها \\ وتصغرُ في عينِ العظيمِ العظائمُ",
  "يُكلَّفُ سيفُ الدولةِ الجيشَ همَّهُ \\ وقد عجزَت عنهُ الجيوشُ الخضارمُ",
].join("\n\n");

const ARABIC_TRIPLE = "العلمُ نورٌ \\ والجهلُ ظلامٌ \\ والعقلُ ميزانٌ";

const FARSI_GHAZAL = [
  "بیا که قصر امل سخت سست بنیاد است \\ بیار باده که بنیاد عمر بر باد است",
  "غلام همت آنم که زیر چرخ کبود \\ ز هر چه رنگ تعلق پذیرد آزاد است",
  "چه باشد ار شود افشای راز دلدارم \\ چو عاشقان همه عالم بدین گرفتارند",
].join("\n\n");

const FARSI_MASNAVI =
  "بشنو این نی چون شکایت می‌کند \\ از جدایی‌ها حکایت می‌کند";

const FARSI_SAADI = "| بنی آدم اعضای یکدیگرند \\ که در آفرینش ز یک گوهرند";

const URDU_GHALIB = [
  "دل ناداں تجھے ہوا کیا ہے \\ آخر اس درد کی دوا کیا ہے",
  "ہم ہیں مشتاق اور وہ بیزار \\ یا الہی یہ ماجرا کیا ہے",
  "میں بھی منہ میں زبان رکھتا ہوں \\ کاش پوچھو کہ مدعا کیا ہے",
].join("\n\n");

const URDU_IQBAL =
  "خودی کو کر بلند اتنا کہ ہر تقدیر سے پہلے \\ خدا بندے سے خود پوچھے بتا تیری رضا کیا ہے";

const URDU_MIR = [
  "الٹی ہو گئیں سب تدبیریں کچھ نہ دوا نے کام کیا \\ دیکھا اس بیماریِ دل نے آخر کام تمام کیا",
  "مت سہل ہمیں جانو پھرتا ہے فلک برسوں \\ تب خاک کے پردے سے انسان نکلتے ہیں",
].join("\n\n");

const URDU_MARSIYA = [
  "شاه كے اصحاب تھے \\ خلق ميں الباب تھے \\ صدق كے ارباب تھے \\",
  "هو گئے شہ پر فدا \\",
  "هائے كربلاء والو \\ هائے كربلاء والو",
].join("\n");

const URDU_MARSIYA_2 = [
  "اهلِ بيت تھے عجب \\ اٗفقِ دعوت كے شٗہٗب \\ هو گئے قربان سب \\",
  "خالي لشكر هو گيا \\",
  "هائے كربلاء والو \\ هائے كربلاء والو",
].join("\n");

const URDU_RUBAI = "مست چلو \\ شاد چلو \\ آزاد چلو \\ خوش چلو";

const URDU_REFRAIN = [
  "دل ناداں تجھے ہوا کیا ہے \\ آخر اس درد کی دوا کیا ہے %",
  "ہم ہیں مشتاق اور وہ بیزار \\ یا الہی یہ ماجرا کیا ہے",
  "میں بھی منہ میں زبان رکھتا ہوں \\ کاش پوچھو کہ مدعا کیا ہے",
  "دل ناداں تجھے ہوا کیا ہے \\ آخر اس درد کی دوا کیا ہے %",
].join("\n\n");

const MULTI_POEM_MULTI_SCRIPT = [
  ARABIC_BAYT,
  "---",
  FARSI_MASNAVI,
  "---",
  "دل ناداں تجھے ہوا کیا ہے \\ آخر اس درد کی دوا کیا ہے",
].join("\n");

// ── OOXML helpers ─────────────────────────────────────────────────────────────

function escXml(s) {
  return String(s || "")
    .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

function heading(text, level) {
  // level 1 = section title, level 2 = subsection label
  var before = level === 1 ? "360" : "200";
  var after  = level === 1 ? "120" :  "60";
  var sz     = level === 1 ? "28"  :  "22";
  var bold   = level === 1 ? "<w:b/>" : "<w:b/><w:i/>";
  var color  = level === 1 ? "<w:color w:val=\"1F6F68\"/>" : "";
  return "<w:p>" +
    "<w:pPr><w:spacing w:before=\"" + before + "\" w:after=\"" + after + "\"/></w:pPr>" +
    "<w:r><w:rPr>" + bold + color +
    "<w:sz w:val=\"" + sz + "\"/><w:szCs w:val=\"" + sz + "\"/></w:rPr>" +
    "<w:t>" + escXml(text) + "</w:t></w:r>" +
    "</w:p>";
}

function blankPara() {
  return "<w:p><w:pPr><w:spacing w:after=\"80\"/></w:pPr></w:p>";
}

function poem(source, opts) {
  var o = Object.assign({ justifyMode: "none", gapWidth: 1 }, opts || {});
  return AshaarWord.renderForWordOoxml(source, o, Ashaar, TEXT_WIDTH_TWIPS);
}

function tabstopPoem(source, opts) {
  var o = Object.assign({ justifyMode: "none" }, opts || {});
  return AshaarTabStop.poemToOoxml(source, o, Ashaar, TEXT_WIDTH_TWIPS);
}

function note(text) {
  return "<w:p>" +
    "<w:pPr><w:spacing w:before=\"40\" w:after=\"40\"/>" +
    "<w:ind w:left=\"360\"/></w:pPr>" +
    "<w:r><w:rPr><w:i/><w:color w:val=\"5F666D\"/>" +
    "<w:sz w:val=\"18\"/><w:szCs w:val=\"18\"/></w:rPr>" +
    "<w:t>" + escXml(text) + "</w:t></w:r></w:p>";
}

// ── Build document body ───────────────────────────────────────────────────────

const parts = [];

function add(...items) { parts.push(...items); }

// ═══════════════════════════════════════════════════════════════
//  ARABIC
// ═══════════════════════════════════════════════════════════════
add(heading("Arabic Poetry", 1));

add(heading("Al-Mutanabbi — single couplet, balanced", 2));
add(note("إذا غامرت في شرف مروم / فلا تقنع بما دون النجوم"));
add(poem(ARABIC_BAYT));

add(heading("Al-Mutanabbi — qasida (3 stanzas), balanced", 2));
add(poem(ARABIC_QASIDA));

add(heading("3-misra triple row (علم / جهل / عقل)", 2));
add(poem(ARABIC_TRIPLE));

add(heading("Single couplet — stacked layout", 2));
add(poem(ARABIC_BAYT, { layoutMode: "stacked" }));

add(heading("Single couplet — compact layout", 2));
add(poem(ARABIC_BAYT, { layoutMode: "compact" }));

add(heading("Single couplet — kashida (tatweelCount 6)", 2));
add(poem(ARABIC_BAYT, { justifyMode: "kashida", tatweelCount: 6 }));

add(heading("Single couplet — tab-stop paragraphs", 2));
add(tabstopPoem(ARABIC_BAYT));

add(blankPara());

// ═══════════════════════════════════════════════════════════════
//  FARSI
// ═══════════════════════════════════════════════════════════════
add(heading("Farsi / Persian Poetry", 1));

add(heading("Hafiz — ghazal (3 bayts)", 2));
add(poem(FARSI_GHAZAL));

add(heading("Rumi — Masnavi opening couplet", 2));
add(note("بشنو این نی چون شکایت می‌کند / از جدایی‌ها حکایت می‌کند"));
add(poem(FARSI_MASNAVI));

add(heading("Saadi — Gulistan solo line (| syntax)", 2));
add(poem(FARSI_SAADI));

add(heading("Hafiz ghazal — Arabic serif font", 2));
add(poem(FARSI_GHAZAL, { fontMode: "arabic-serif" }));

add(heading("Rumi couplet — kashida (tatweelCount 5)", 2));
add(poem(FARSI_MASNAVI, { justifyMode: "kashida", tatweelCount: 5 }));

add(heading("Rumi couplet — tab-stop paragraphs", 2));
add(tabstopPoem(FARSI_MASNAVI));

add(blankPara());

// ═══════════════════════════════════════════════════════════════
//  URDU
// ═══════════════════════════════════════════════════════════════
add(heading("Urdu Poetry", 1));

add(heading("Ghalib — ghazal (3 bayts)", 2));
add(poem(URDU_GHALIB));

add(heading("Iqbal — single couplet, Nastaliq (no justification)", 2));
add(note("خودی کو کر بلند اتنا کہ ہر تقدیر سے پہلے / خدا بندے سے خود پوچھے بتا تیری رضا کیا ہے"));
add(poem(URDU_IQBAL, { fontMode: "nastaliq" }));

add(heading("Iqbal — Nastaliq, spacing mode (w:jc both — Word word-spacing)", 2));
add(poem(URDU_IQBAL, { fontMode: "nastaliq", justifyMode: "spacing" }));

add(heading("Ghalib ghazal — Nastaliq, spacing mode (3 bayts)", 2));
add(poem(URDU_GHALIB, { fontMode: "nastaliq", justifyMode: "spacing" }));

add(heading("Farsi Hafiz — Arabic serif, spacing mode", 2));
add(poem(FARSI_GHAZAL, { fontMode: "arabic-serif", justifyMode: "spacing" }));

add(heading("Mir Taqi Mir — 2 bayts", 2));
add(poem(URDU_MIR));

add(heading("Ghalib ghazal — with refrain bayts (% marker)", 2));
add(poem(URDU_REFRAIN));

add(heading("Marsiya stanza 1 — 3-misra + solo + paired refrain", 2));
add(poem(URDU_MARSIYA));

add(heading("Marsiya stanza 2 — 3-misra + solo + paired refrain", 2));
add(poem(URDU_MARSIYA_2));

add(heading("Multi-stanza marsiya (--- separator)", 2));
add(poem([URDU_MARSIYA, "---", URDU_MARSIYA_2].join("\n")));

add(heading("Rubaʿi — 4-misra row", 2));
add(poem(URDU_RUBAI));

add(heading("Ghalib ghazal — stacked layout", 2));
add(poem(URDU_GHALIB, { layoutMode: "stacked" }));

add(heading("Ghalib ghazal — compact layout", 2));
add(poem(URDU_GHALIB, { layoutMode: "compact" }));

add(heading("Iqbal — tab-stop paragraphs", 2));
add(tabstopPoem(URDU_IQBAL));

add(heading("Marsiya — tab-stop paragraphs (3-misra)", 2));
add(tabstopPoem(URDU_MARSIYA));

add(heading("Ghalib — kashida (tatweelCount 6)", 2));
add(poem(URDU_GHALIB, { justifyMode: "kashida", tatweelCount: 6 }));

add(heading("Ghalib — Nastaliq + kashida", 2));
add(poem(URDU_GHALIB, { fontMode: "nastaliq", justifyMode: "kashida", tatweelCount: 4 }));

add(blankPara());

// ═══════════════════════════════════════════════════════════════
//  MULTI-SCRIPT
// ═══════════════════════════════════════════════════════════════
add(heading("Multi-Script — Arabic + Farsi + Urdu (--- separator)", 1));
add(poem(MULTI_POEM_MULTI_SCRIPT));

add(blankPara());

// ═══════════════════════════════════════════════════════════════
//  GRID + TEMPLATE OOXML
// ═══════════════════════════════════════════════════════════════
add(heading("Grid & Template", 1));

add(heading("Blank 12-column grid (Drop Grid)", 2));
add(note("Merge cells in Word, then use Capture from Word to save as a template."));
add(AshaarWord.generateBareGrid12Ooxml(TEXT_WIDTH_TWIPS));

add(heading("Captured template — bayt layout (5 + 2 + 5)", 2));
add(AshaarWord.templateToOoxml(
  { columnCount: 12, rows: [[{ span: 5 }, { span: 2 }, { span: 5 }]] },
  TEXT_WIDTH_TWIPS, {}
));

add(heading("Captured template — marsiya layout (4 + 4 + 4 / 12)", 2));
add(AshaarWord.templateToOoxml(
  { columnCount: 12, rows: [
    [{ span: 4 }, { span: 4 }, { span: 4 }],
    [{ span: 12 }],
  ]},
  TEXT_WIDTH_TWIPS, {}
));

add(blankPara());

// ── Assemble docx ─────────────────────────────────────────────────────────────

const wns = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const documentXml =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<w:document ' + wns + '><w:body>' +
  parts.join("") +
  '<w:sectPr/></w:body></w:document>';

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

const out = path.join(dir, "corpus-test.docx");
zip.writeZip(out);
console.log("Created", out);
console.log("Sections:", [
  "Arabic: single couplet, qasida (3 stanzas), triple, stacked, compact, kashida, tab-stop",
  "Farsi:  Hafiz ghazal, Rumi Masnavi, Saadi solo, Arabic serif font, kashida, tab-stop",
  "Urdu:   Ghalib ghazal, Iqbal Nastaliq, Mir, refrain bayts, marsiya (3-misra × 2),",
  "        multi-stanza marsiya, rubaʿi (4-misra), stacked, compact, tab-stop, kashida",
  "Multi-script: Arabic + Farsi + Urdu (--- separator)",
  "Grid:   blank 12-col grid, bayt template (5+2+5), marsiya template (4+4+4 / 12)",
].map(s => "  " + s).join("\n"));

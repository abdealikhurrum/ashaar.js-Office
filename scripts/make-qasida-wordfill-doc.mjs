#!/usr/bin/env node
// §3 "Let Word fill it" verification doc — al-Mutanabbi's Sayfiyya qasida
// rendered through the PRODUCTION path (renderForWordOoxml, justifyMode:"css")
// at escalating strengths, next to an un-justified baseline.
//
// Each panel is what the add-in's INSERT path emits for word-fill mode:
//   misraParaXml → native Word kashida jc (low/medium/high) + a shrunk trailing
//   <w:br/> so the single (last) line of each misra cell actually stretches.
// Width per panel mirrors the app's §4 expansion (kashidaExpansionFraction,
// 0..~15% by strength), capped at the page width.
//
// Open test-documents/qasida-wordfill-test.docx in desktop Word and check:
//   1. Arabic letters KASHIDA-STRETCH (stroke elongation), not word-spacing.
//   2. Strength escalates: low (str 4) < medium (str 12) < high (str 22).
//   3. The trailing empty line under each misra is negligible (~2pt).
//   4. Cells widen with strength (up to ~15% at full).
//   5. The Latin control line falls back to spacing (distribute), no stretch.
// Run: node scripts/make-qasida-wordfill-doc.mjs

import { createRequire } from "module";
import { fileURLToPath } from "url";
import path from "path";
import fs from "fs";

const require = createRequire(import.meta.url);
const AdmZip     = require("adm-zip");
const Ashaar     = require("../src/vendor/ashaar.js");
const AshaarWord = require("../src/taskpane/word-html.js");

const PAGE_TWIPS = 9360; // 6.5" at 1440 twips/inch — the hard page ceiling
const BASE_TWIPS = 6120; // ~4.25": narrower than the page so kashida has slack
                         // to fill AND so §4 width-expansion is visible.

// ── al-Mutanabbi — al-Sayfiyya (على قدر أهل العزم) ─────────────────────────
// Best-known verses of the qasida to Sayf al-Dawla on the reconquest of
// al-Hadath (954 CE). Meter: al-Tawil; rhyme: -mu. `\` = the two hemistichs
// of a bayt; blank line = next bayt. Diacritics per common printed diwans.
const SAYFIYYA = [
  "على قَدرِ أهلِ العَزمِ تأتي العَزائمُ \\ وتأتي على قَدرِ الكِرامِ المَكارمُ",
  "وتَعظُمُ في عَينِ الصَغيرِ صِغارُها \\ وتَصغُرُ في عَينِ العَظيمِ العَظائمُ",
  "يُكَلِّفُ سَيفُ الدَولةِ الجَيشَ هَمَّهُ \\ وقد عَجِزَت عنه الجُيوشُ الخَضارمُ",
  "ويَطلُبُ عندَ الناسِ ما عندَ نفسِهِ \\ وذلك ما لا تَدَّعيهِ الضَراغِمُ",
  "هلِ الحَدَثُ الحَمراءُ تَعرِفُ لَونَها \\ وتَعلَمُ أيُّ الساقِيَينِ الغَمائمُ",
  "سَقَتها الغَمامُ الغُرُّ قبلَ نُزولِهِ \\ فلمّا دَنا منها سَقَتها الجَماجمُ",
  "وقَفتَ وما في المَوتِ شَكٌّ لواقفٍ \\ كأنّكَ في جَفنِ الرَّدى وهو نائمُ",
  "تَمُرُّ بكَ الأبطالُ كَلمى هَزيمةً \\ ووَجهُكَ وَضّاحٌ وثَغرُكَ باسمُ",
  "تَجاوَزتَ مِقدارَ الشَجاعةِ والنُّهى \\ إلى قَولِ قَومٍ أنتَ بالغَيبِ عالمُ",
].join("\n\n");

// Diacritic-free stress version — cleanest kashida stretch (isolates the
// justification behaviour from mark-stacking).
const SAYFIYYA_BARE = [
  "على قدر أهل العزم تأتي العزائم \\ وتأتي على قدر الكرام المكارم",
  "وتعظم في عين الصغير صغارها \\ وتصغر في عين العظيم العظائم",
  "يكلف سيف الدولة الجيش همه \\ وقد عجزت عنه الجيوش الخضارم",
].join("\n\n");

// Latin control — must NOT kashida; falls back to distribute (spacing).
const LATIN_CONTROL = "On the measure of the resolute the resolves arrive \\ and on the measure of the noble come the noble deeds";

// ── OOXML helpers (packaging mirrors scripts/make-test-doc.mjs) ─────────────
function escXml(s) {
  return String(s || "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

function heading(text, level) {
  var before = level === 1 ? "360" : "220";
  var after  = level === 1 ? "120" : "60";
  var sz     = level === 1 ? "30"  : "24";
  var bold   = level === 1 ? "<w:b/>" : "<w:b/>";
  var color  = level === 1 ? "<w:color w:val=\"1F6F68\"/>" : "<w:color w:val=\"A7352A\"/>";
  return "<w:p><w:pPr><w:spacing w:before=\"" + before + "\" w:after=\"" + after + "\"/></w:pPr>" +
    "<w:r><w:rPr>" + bold + color + "<w:sz w:val=\"" + sz + "\"/><w:szCs w:val=\"" + sz + "\"/></w:rPr>" +
    "<w:t>" + escXml(text) + "</w:t></w:r></w:p>";
}

function note(text) {
  return "<w:p><w:pPr><w:spacing w:before=\"40\" w:after=\"80\"/><w:ind w:left=\"200\"/></w:pPr>" +
    "<w:r><w:rPr><w:i/><w:color w:val=\"5F666D\"/><w:sz w:val=\"18\"/><w:szCs w:val=\"18\"/></w:rPr>" +
    "<w:t>" + escXml(text) + "</w:t></w:r></w:p>";
}

function blankPara() { return "<w:p><w:pPr><w:spacing w:after=\"120\"/></w:pPr></w:p>"; }

// Render a poem exactly as the add-in's word-fill INSERT path does, including
// the §4 width expansion (kashidaExpansionFraction), page-capped.
function wordFill(source, strength) {
  var frac = AshaarWord.kashidaExpansionFraction(strength);
  var width = Math.min(PAGE_TWIPS, Math.round(BASE_TWIPS * (1 + frac)));
  var opts = { justifyMode: "css", tatweelCount: strength, gapWidth: 1 };
  return AshaarWord.renderForWordOoxml(source, opts, Ashaar, width);
}

function plain(source) {
  var opts = { justifyMode: "none", gapWidth: 1 };
  return AshaarWord.renderForWordOoxml(source, opts, Ashaar, BASE_TWIPS);
}

// ── Build body ──────────────────────────────────────────────────────────────
const parts = [];
function add() { for (var i = 0; i < arguments.length; i++) parts.push(arguments[i]); }

add(heading("§3 — “Let Word fill it” verification", 1));
add(note("al-Mutanabbi, al-Sayfiyya (على قدر أهل العزم). Each section below is the SAME qasida rendered through the add-in's word-fill path. Compare top-to-bottom: the baseline sits at natural width; each kashida panel should fill wider and stretch harder as strength climbs. Kashida quality depends on the Word font in use (this mode needs no font loaded into the pane)."));

add(heading("0 · Baseline — Leave as typed (justifyMode: none)", 2));
add(note("Reference: lines sit at their natural width, aligned by column. No stretch, no trailing break."));
add(plain(SAYFIYYA));
add(blankPara());

add(heading("1 · Word justify — LOW strength (4 → lowKashida)", 2));
add(note("Expect: gentle letter elongation; cell width ≈ baseline (little §4 expansion)."));
add(wordFill(SAYFIYYA, 4));
add(blankPara());

add(heading("2 · Word justify — MEDIUM strength (12 → mediumKashida)", 2));
add(note("Expect: noticeably more elongation than LOW; cell ≈ 7% wider."));
add(wordFill(SAYFIYYA, 12));
add(blankPara());

add(heading("3 · Word justify — HIGH strength (22 → highKashida)", 2));
add(note("Expect: strongest elongation; cell ≈ 14% wider (capped at page)."));
add(wordFill(SAYFIYYA, 22));
add(blankPara());

add(heading("4 · Stress test — HIGH strength, diacritic-free", 2));
add(note("Cleanest kashida: bare consonantal skeleton, no mark-stacking. This is the clearest read on stroke elongation."));
add(wordFill(SAYFIYYA_BARE, 22));
add(blankPara());

add(heading("5 · Control — Latin text, HIGH strength", 2));
add(note("Expect: NO kashida. Non-Arabic falls back to distribute (word-spacing), with no trailing break."));
add(wordFill(LATIN_CONTROL, 22));
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
const out = path.join(dir, "qasida-wordfill-test.docx");
zip.writeZip(out);
console.log("Created", out);

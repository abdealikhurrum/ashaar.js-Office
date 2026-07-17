#!/usr/bin/env node
// Generates test-documents/styles-test.docx — a plain prose "bayan"-style
// document (mostly Urdu, with a vocalized Arabic Quran verse) whose content is
// UNSTYLED on purpose. Use it to exercise the Styles tab (Task 10 manual
// verification):
//   • Select a heading line          → apply Ashaar Heading 1 / 2 / 3
//   • Select a phrase mid-sentence    → apply Ashaar Emphasis (red + size bump)
//   • Select a paragraph              → apply Ashaar Quote (left/right borders)
//   • Select the vocalized verse      → apply Ashaar Quran Quote (font + line height)
//   • RTL document setup              → run once, then check the body font and
//                                       the pre-inserted footnote's numbering
//   • Switch style groups             → re-check the applied styles reflow
// The document deliberately does NOT pre-apply any Ashaar style — the whole
// point is to select this content and apply styles from the pane.
// Run: node scripts/make-styles-test-doc.mjs   (or: npm run make-styles-doc)

import { createRequire } from "module";
import { fileURLToPath } from "url";
import path from "path";
import fs from "fs";

const require = createRequire(import.meta.url);
const AdmZip = require("adm-zip");

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const OUT = path.join(__dirname, "..", "test-documents", "styles-test.docx");

function esc(s) {
  return String(s == null ? "" : s)
    .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

// A run of Arabic-script text (rtl) or plain LTR text.
function run(text, rtl) {
  var rpr = rtl ? "<w:rPr><w:rtl/></w:rPr>" : "";
  return "<w:r>" + rpr + '<w:t xml:space="preserve">' + esc(text) + "</w:t></w:r>";
}

// Paragraph. align: left|right|center; rtl toggles bidi. runsXml overrides the
// default single-run body when a paragraph needs mixed runs (emphasis span,
// footnote reference).
function para(text, opts) {
  opts = opts || {};
  var jc = opts.align || (opts.rtl ? "right" : "left");
  var ppr = "<w:pPr>" + (opts.rtl ? "<w:bidi/>" : "") +
    '<w:jc w:val="' + jc + '"/>' +
    (opts.spaceAfter != null ? '<w:spacing w:after="' + opts.spaceAfter + '"/>' : "") +
    "</w:pPr>";
  var body = opts.runsXml != null ? opts.runsXml : run(text, !!opts.rtl);
  return "<w:p>" + ppr + body + "</w:p>";
}

// An instruction label for the tester (LTR, bold, teal — matches adopt-test).
function label(text) {
  return "<w:p><w:pPr><w:spacing w:before=\"200\" w:after=\"60\"/></w:pPr>" +
    '<w:r><w:rPr><w:b/><w:color w:val="1F6F68"/></w:rPr><w:t xml:space="preserve">' + esc(text) + "</w:t></w:r></w:p>";
}

function spacer() { return "<w:p/>"; }

// ── Document body ───────────────────────────────────────────────────────────

var parts = [];

parts.push(para("Ashaar Styles — test document", { align: "left", spaceAfter: 120 }));
parts.push(para("Every block below is plain, UNSTYLED text. Open the Styles tab, then select a block and apply the matching style. Nothing here is pre-styled — that is the point.", { align: "left", spaceAfter: 200 }));

// HEADINGS — three levels
parts.push(label("HEADINGS — select each line, apply Ashaar Heading 1 / 2 / 3 in turn."));
parts.push(para("پہلا باب — نصیحت", { rtl: true }));
parts.push(para("پہلی فصل — آغازِ کلام", { rtl: true }));
parts.push(para("ذیلی عنوان — ایک نکتہ", { rtl: true }));
parts.push(spacer());

// BODY — plain Normal paragraphs (for RTL document setup + general reading)
parts.push(label("BODY — plain paragraphs (leave as Normal). Run “Set up RTL document”, then confirm these pick up the complex-script font/size."));
parts.push(para("یہ ایک عام پیراگراف ہے جو معمول کے انداز میں لکھا گیا ہے۔ اس میں کوئی خاص طرزِ تحریر لاگو نہیں کی گئی۔", { rtl: true, spaceAfter: 80 }));
parts.push(para("دوسرا پیراگراف بھی سادہ ہے تاکہ آپ دیکھ سکیں کہ آر ٹی ایل سیٹ اپ کے بعد پورے متن پر فونٹ اور سائز کیسے لاگو ہوتے ہیں۔", { rtl: true }));
parts.push(spacer());

// EMPHASIS — a phrase mid-sentence
parts.push(label("EMPHASIS — select ONLY the middle phrase between the brackets, then apply Ashaar Emphasis (should turn red and grow a few points relative to the surrounding text)."));
parts.push(para(null, {
  rtl: true,
  runsXml:
    run("اس جملے میں ", true) +
    run("«یہ الفاظ نمایاں کریں»", true) +
    run(" اور باقی متن عام رہے۔", true)
}));
parts.push(spacer());

// QUOTE — a block-quote candidate
parts.push(label("QUOTE — select this whole paragraph, apply Ashaar Quote (left + right borders). Try the per-instance indent override on it too."));
parts.push(para("جو قومیں اپنی تاریخ سے سبق نہیں سیکھتیں، وہ اسے دہرانے پر مجبور ہو جاتی ہیں۔", { rtl: true }));
parts.push(spacer());

// QURAN QUOTE — a fully vocalized Arabic verse (harakat/tashkeel present)
parts.push(label("QURAN QUOTE — select the vocalized verse below, apply Ashaar Quran Quote. Confirm the harakat (tashkeel) render correctly and the line height is comfortable. Try the per-instance line-height override."));
parts.push(para("بِسْمِ ٱللَّهِ ٱلرَّحْمَٰنِ ٱلرَّحِيمِ ۝ ٱلْحَمْدُ لِلَّهِ رَبِّ ٱلْعَٰلَمِينَ ۝ ٱلرَّحْمَٰنِ ٱلرَّحِيمِ ۝ مَٰلِكِ يَوْمِ ٱلدِّينِ", { rtl: true }));
parts.push(spacer());

// FOOTNOTE — pre-inserted, for the RTL footnote-numbering check (Task 9/10)
parts.push(label("FOOTNOTE — this paragraph carries a real footnote (see the marker). After “Set up RTL document”, confirm the footnote marker and the footnote-pane text follow right-to-left layout/numbering."));
parts.push(para(null, {
  rtl: true,
  runsXml:
    run("اس جملے کے آخر میں ایک حاشیہ موجود ہے", true) +
    '<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/><w:rtl/></w:rPr><w:footnoteReference w:id="1"/></w:r>' +
    run("۔", true)
}));

var bodyXml = parts.join("") +
  '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>';

var documentXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
  "<w:body>" + bodyXml + "</w:body></w:document>";

// Footnotes part: the two required boilerplate footnotes (separator id -1,
// continuation separator id 0) plus the single real footnote (id 1) referenced
// from the body above. The separator paragraphs are bidi + right-aligned so the
// separator line sits on the RIGHT in an RTL document — Office.js cannot reach
// these special footnotes, so authoring them here (and in the future template)
// is the only programmatic way to move the separator; otherwise it's a manual
// Draft-view step (References → Show Notes → Footnote Separator).
var footnotesXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
  '<w:footnote w:type="separator" w:id="-1"><w:p><w:pPr><w:bidi/><w:jc w:val="right"/><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:separator/></w:r></w:p></w:footnote>' +
  '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:pPr><w:bidi/><w:jc w:val="right"/><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>' +
  '<w:footnote w:id="1"><w:p><w:pPr><w:bidi/><w:jc w:val="right"/><w:pStyle w:val="FootnoteText"/></w:pPr>' +
  '<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/><w:rtl/></w:rPr><w:footnoteRef/></w:r>' +
  '<w:r><w:rPr><w:rtl/></w:rPr><w:t xml:space="preserve"> یہ حاشیے کا متن ہے — آر ٹی ایل ترتیب کی جانچ کے لیے۔</w:t></w:r>' +
  '</w:p></w:footnote>' +
  '</w:footnotes>';

var contentTypes = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
  '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
  '<Default Extension="xml" ContentType="application/xml"/>' +
  '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
  '<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>' +
  "</Types>";

var rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
  "</Relationships>";

var docRels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>' +
  "</Relationships>";

var zip = new AdmZip();
zip.addFile("[Content_Types].xml", Buffer.from(contentTypes, "utf8"));
zip.addFile("_rels/.rels", Buffer.from(rels, "utf8"));
zip.addFile("word/document.xml", Buffer.from(documentXml, "utf8"));
zip.addFile("word/footnotes.xml", Buffer.from(footnotesXml, "utf8"));
zip.addFile("word/_rels/document.xml.rels", Buffer.from(docRels, "utf8"));
fs.mkdirSync(path.dirname(OUT), { recursive: true });
zip.writeZip(OUT);
console.log("Wrote " + OUT + " (" + fs.statSync(OUT).size + " bytes)");

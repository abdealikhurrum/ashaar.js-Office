/**
 * poetry-corpus.test.js
 *
 * Comprehensive corpus tests covering Arabic, Farsi, and Urdu poetry across
 * all supported structures: couplets, multi-misra rows, solo lines, refrains,
 * multi-stanza, multi-poem, and all layout/OOXML paths.
 * Also covers new features: bare grid, template OOXML, font-profile justification.
 */

const assert = require("assert");
const Ashaar = require("../src/vendor/ashaar");
const AshaarWord = require("../src/taskpane/word-html");
const AshaarTabStop = require("../src/taskpane/word-tabstop");
const AshaarJustify = require("../src/vendor/ashaar-justify");

// ── Poetry corpus ─────────────────────────────────────────────────────────────

// Arabic — Al-Mutanabbi (classical qasida couplets)
const ARABIC_QASIDA = [
  "على قدرِ أهلِ العزمِ تأتي العزائمُ \\ وتأتي على قدرِ الكرامِ المكارمُ",
  "وتعظمُ في عينِ الصغيرِ صغارُها \\ وتصغرُ في عينِ العظيمِ العظائمُ",
].join("\n\n");  // two stanzas

// Arabic — single couplet with diacritics
const ARABIC_BAYT = "إذا غامَرتَ في شَرَفٍ مَرومِ \\ فَلا تَقنَع بما دونَ النُّجومِ";

// Arabic — 3-misra sajʿ-style row (single stanza)
const ARABIC_TRIPLE = "العلمُ نورٌ \\ والجهلُ ظلامٌ \\ والعقلُ ميزانٌ";

// Farsi — Hafiz, Divan (ghazal, 2 bayts)
const FARSI_GHAZAL = [
  "بیا که قصر امل سخت سست بنیاد است \\ بیار باده که بنیاد عمر بر باد است",
  "غلام همت آنم که زیر چرخ کبود \\ ز هر چه رنگ تعلق پذیرد آزاد است",
].join("\n\n");

// Farsi — Rumi, Masnavi opening (single couplet)
const FARSI_MASNAVI = "بشنو این نی چون شکایت می‌کند \\ از جدایی‌ها حکایت می‌کند";

// Farsi — Saadi, Gulistan (solo line with | syntax)
const FARSI_SOLO = "| بنی آدم اعضای یکدیگرند";

// Urdu — Ghalib ghazal (2 bayts)
const URDU_GHAZAL = [
  "دل ناداں تجھے ہوا کیا ہے \\ آخر اس درد کی دوا کیا ہے",
  "ہم ہیں مشتاق اور وہ بیزار \\ یا الہی یہ ماجرا کیا ہے",
].join("\n\n");

// Urdu — Iqbal (single long couplet)
const URDU_IQBAL = "خودی کو کر بلند اتنا کہ ہر تقدیر سے پہلے \\ خدا بندے سے خود پوچھے بتا تیری رضا کیا ہے";

// Urdu — Marsiya (3-misra + solo + paired refrain)
const URDU_MARSIYA = [
  "شاه كے اصحاب تھے \\ خلق ميں الباب تھے \\ صدق كے ارباب تھے \\",
  "هو گئے شہ پر فدا \\",
  "هائے كربلاء والو \\ هائے كربلاء والو",
].join("\n");

// Multi-poem: Arabic + Farsi separated by ---
const MULTI_POEM = [
  ARABIC_BAYT,
  "---",
  FARSI_MASNAVI,
].join("\n");

// Mixed ghazal with refrain bayt marker (% at end of bayt line)
// Format: "sadr \\ ajuz %" marks the whole bayt as refrain
const URDU_WITH_REFRAIN = [
  "دل ناداں تجھے ہوا کیا ہے \\ آخر اس درد کی دوا کیا ہے %",
  "ہم ہیں مشتاق اور وہ بیزار \\ یا الہی یہ ماجرا کیا ہے",
].join("\n\n");

// Urdu — 4-misra rubaʿi
const URDU_RUBAI = "مست چلو \\ شاد چلو \\ آزاد چلو \\ خوش چلو";

// ── Helpers ───────────────────────────────────────────────────────────────────

function countMatches(str, re) {
  return (str.match(re) || []).length;
}

// ── 1. Arabic couplet — renderForWord ─────────────────────────────────────────

console.log("  Arabic: single couplet renderForWord");
{
  const html = AshaarWord.renderForWord(ARABIC_BAYT, { justifyMode: "none" }, Ashaar);
  assert.match(html, /<table dir="rtl"/);
  assert.match(html, /إذا غامَرتَ في شَرَفٍ مَرومِ/);
  assert.match(html, /فَلا تَقنَع بما دونَ النُّجومِ/);
  // 12-column grid
  assert.equal(countMatches(html, /<col style=/g), 12);
  // 2 cells (sadr + ajuz)
  const row = html.match(/<tr>([\s\S]*?)<\/tr>/)[1];
  assert.equal(countMatches(row, /<td/g), 2);
}

// ── 2. Arabic qasida — multi-stanza renderForWord ────────────────────────────

console.log("  Arabic: multi-stanza qasida renderForWord");
{
  const html = AshaarWord.renderForWord(ARABIC_QASIDA, { justifyMode: "none" }, Ashaar);
  assert.match(html, /على قدرِ أهلِ العزمِ/);
  assert.match(html, /وتعظمُ في عينِ الصغيرِ/);
  // Two separate tables (two stanzas)
  assert.equal(countMatches(html, /<table dir="rtl"/g), 2);
}

// ── 3. Arabic 3-misra triple row ──────────────────────────────────────────────

console.log("  Arabic: 3-misra row (triple) renderForWord");
{
  const html = AshaarWord.renderForWord(ARABIC_TRIPLE, { justifyMode: "none" }, Ashaar);
  assert.match(html, /العلمُ نورٌ/);
  assert.match(html, /والجهلُ ظلامٌ/);
  assert.match(html, /والعقلُ ميزانٌ/);
  assert.equal(countMatches(html, /<col style=/g), 12);
  const row = html.match(/<tr>([\s\S]*?)<\/tr>/)[1];
  assert.equal(countMatches(row, /<td/g), 3);
}

// ── 4. Farsi ghazal — renderForWord ──────────────────────────────────────────

console.log("  Farsi: ghazal renderForWord");
{
  const html = AshaarWord.renderForWord(FARSI_GHAZAL, { justifyMode: "none" }, Ashaar);
  assert.match(html, /بیا که قصر امل/);
  assert.match(html, /غلام همت آنم/);
  assert.equal(countMatches(html, /<table dir="rtl"/g), 2);
}

// ── 5. Farsi Masnavi — renderForWordOoxml ────────────────────────────────────

console.log("  Farsi: Masnavi couplet OOXML");
{
  const ooxml = AshaarWord.renderForWordOoxml(FARSI_MASNAVI,
    { justifyMode: "none", gapWidth: 1 }, Ashaar, 9360);
  assert.match(ooxml, /<w:tbl>/);
  assert.match(ooxml, /بشنو این نی چون شکایت می‌کند/);
  assert.match(ooxml, /از جدایی‌ها حکایت می‌کند/);
  // N=2, gapCols=1 → GRID=13 (2*6+1)
  assert.equal(countMatches(ooxml, /<w:gridCol /g), 13);
  // 2 misra cells + 1 gap = 3 tc elements
  assert.equal(countMatches(ooxml, /<w:tc>/g), 3);
  assert.match(ooxml, /<w:bidiVisual\/>/);
}

// ── 6. Farsi solo line ────────────────────────────────────────────────────────

console.log("  Farsi: solo line (| syntax) renderForWord");
{
  const html = AshaarWord.renderForWord(FARSI_SOLO, { justifyMode: "none" }, Ashaar);
  assert.match(html, /بنی آدم اعضای یکدیگرند/);
  assert.match(html, /colspan="12"/);  // solo spans full 12-col grid
}

// ── 7. Urdu ghazal — renderForWord ───────────────────────────────────────────

console.log("  Urdu: ghazal renderForWord");
{
  const html = AshaarWord.renderForWord(URDU_GHAZAL, { justifyMode: "none" }, Ashaar);
  assert.match(html, /دل ناداں تجھے ہوا کیا ہے/);
  assert.match(html, /ہم ہیں مشتاق اور وہ بیزار/);
  assert.equal(countMatches(html, /<table dir="rtl"/g), 2);
}

// ── 8. Urdu Iqbal — Nastaliq font ────────────────────────────────────────────

console.log("  Urdu: Iqbal, Nastaliq font");
{
  const html = AshaarWord.renderForWord(URDU_IQBAL, { justifyMode: "none", fontMode: "nastaliq" }, Ashaar);
  assert.match(html, /خودی کو کر بلند اتنا/);
  assert.match(html, /Noto Nastaliq Urdu/);
}

// ── 9. Urdu Marsiya — 3-misra + solo + pair OOXML ────────────────────────────

console.log("  Urdu: Marsiya (3-misra + solo + pair) OOXML");
{
  const ooxml = AshaarWord.renderForWordOoxml(URDU_MARSIYA,
    { justifyMode: "none", gapWidth: 1 }, Ashaar, 9360);
  // N=3, gapCols=1 → GRID=20 (3*6+2)
  assert.equal(countMatches(ooxml, /<w:gridCol /g), 20);
  assert.match(ooxml, /شاه كے اصحاب تھے/);
  assert.match(ooxml, /هو گئے شہ پر فدا/);
  assert.match(ooxml, /هائے كربلاء والو/);
  // Solo row is centered
  assert.match(ooxml, /<w:jc w:val="center"/);
  // 3-misra row: 3 misra cells + 2 gap cells = 5 tc in first row
  const firstRow = ooxml.match(/<w:tr>([\s\S]*?)<\/w:tr>/)[1];
  assert.equal(countMatches(firstRow, /<w:tc>/g), 5);
}

// ── 10. Urdu 4-misra rubaʿi OOXML ────────────────────────────────────────────

console.log("  Urdu: rubaʿi (4-misra) OOXML");
{
  const ooxml = AshaarWord.renderForWordOoxml(URDU_RUBAI,
    { justifyMode: "none", gapWidth: 1 }, Ashaar, 9360);
  // N=4, gapCols=1 → GRID=4*6 + 3*1 = 27
  assert.equal(countMatches(ooxml, /<w:gridCol /g), 27);
  assert.match(ooxml, /مست چلو/);
  assert.match(ooxml, /شاد چلو/);
  assert.match(ooxml, /آزاد چلو/);
  assert.match(ooxml, /خوش چلو/);
  // 4 misra cells + 3 gap cells = 7 tc in the row
  const row = ooxml.match(/<w:tr>([\s\S]*?)<\/w:tr>/)[1];
  assert.equal(countMatches(row, /<w:tc>/g), 7);
}

// ── 11. Multi-poem (--- separator) ───────────────────────────────────────────

console.log("  Multi-poem: Arabic + Farsi separated by ---");
{
  const html = AshaarWord.renderForWord(MULTI_POEM, { justifyMode: "none" }, Ashaar);
  assert.match(html, /إذا غامَرتَ في شَرَفٍ مَرومِ/);
  assert.match(html, /بشنو این نی چون شکایت می‌کند/);
  // Two separate table blocks (two poems)
  assert.equal(countMatches(html, /<table dir="rtl"/g), 2);
}

// ── 12. Refrain bayt (% prefix) ──────────────────────────────────────────────

console.log("  Urdu: refrain bayt (% prefix) colour in HTML");
{
  const html = AshaarWord.renderForWord(URDU_WITH_REFRAIN, { justifyMode: "none" }, Ashaar);
  assert.match(html, /دل ناداں تجھے ہوا کیا ہے/);
  assert.match(html, /ہم ہیں مشتاق اور وہ بیزار/);
  // Refrain bayt should have refrain colour applied
  assert.match(html, /color:#a7352a/);
}

// ── 13. Refrain colour in OOXML ───────────────────────────────────────────────

console.log("  Urdu: refrain colour in OOXML");
{
  const ooxml = AshaarWord.renderForWordOoxml(URDU_WITH_REFRAIN,
    { justifyMode: "none", gapWidth: 1 }, Ashaar, 9360);
  assert.match(ooxml, /<w:color w:val="A7352A"\/>/);
  assert.match(ooxml, /دل ناداں تجھے ہوا کیا ہے/);
}

// ── 14. Stacked layout — all scripts ─────────────────────────────────────────

console.log("  Stacked layout: Arabic, Farsi, Urdu");
{
  for (const [label, src] of [["Arabic", ARABIC_BAYT], ["Farsi", FARSI_MASNAVI], ["Urdu", URDU_IQBAL]]) {
    const html = AshaarWord.renderForWord(src, { justifyMode: "none", layoutMode: "stacked" }, Ashaar);
    assert.match(html, /colspan="12"/, label + ": stacked should span 12 cols");
    assert.match(html, /<br>/, label + ": stacked should have <br> between misras");
  }
}

// ── 15. Compact layout — Arabic ──────────────────────────────────────────────

console.log("  Compact layout: Arabic couplet");
{
  const html = AshaarWord.renderForWord(ARABIC_BAYT,
    { justifyMode: "none", layoutMode: "compact", gapWidth: 4 }, Ashaar);
  assert.match(html, /إذا غامَرتَ في شَرَفٍ مَرومِ/);
  assert.match(html, /فَلا تَقنَع بما دونَ النُّجومِ/);
}

// ── 16. Arabic kashida justification (tatweelCount path) ─────────────────────

console.log("  Kashida: Arabic with tatweelCount");
{
  const html = AshaarWord.renderForWord(ARABIC_BAYT,
    { justifyMode: "kashida", tatweelCount: 4 }, Ashaar);
  assert.match(html, /ـ/);  // tatweel present
}

// ── 17. Farsi kashida justification ──────────────────────────────────────────

console.log("  Kashida: Farsi with tatweelCount");
{
  const plain = AshaarWord.justifyPlainTextBlock("بشنو این نی چون شکایت می‌کند",
    { justifyMode: "kashida", tatweelCount: 3 });
  assert.match(plain, /ـ/);
}

// ── 18. Urdu kashida justification with font profile mock ────────────────────

console.log("  Kashida: Urdu with font profile mock");
{
  // Mock a minimal FontProfile (as probeFont would return)
  const mockFontProfile = {
    getQuality: function (prev, next) { return 0.8; }
  };
  const text = "دل ناداں تجھے ہوا کیا ہے";
  // spreadTatweels path (no canvas ctx)
  const basic = AshaarJustify.spreadTatweels(text, 3);
  assert.match(basic, /ـ/);
  // buildSlots + applySlots path with font profile
  const slots = AshaarJustify.buildSlots(text, { fontQualityBoost: 1.8 }, mockFontProfile);
  assert.ok(Array.isArray(slots), "buildSlots should return array");
  if (slots.length > 0) {
    const applied = AshaarJustify.applySlots(text, slots, 2);
    assert.match(applied, /ـ/);
  }
}

// ── 19. justifyPlainTextBlock — multi-line ────────────────────────────────────

console.log("  justifyPlainTextBlock: multi-line Arabic");
{
  const block = "على قدرِ أهلِ العزمِ تأتي العزائمُ\nوتأتي على قدرِ الكرامِ المكارمُ";
  const result = AshaarWord.justifyPlainTextBlock(block,
    { justifyMode: "kashida", tatweelCount: 2 });
  assert.match(result, /ـ/);
  assert.ok(result.includes("\n"), "newline preserved");
}

// ── 20. Content control tags — all scripts ───────────────────────────────────

console.log("  Content control tags: Arabic, Farsi, Urdu");
{
  for (const [label, src] of [
    ["Arabic", ARABIC_BAYT],
    ["Farsi", FARSI_MASNAVI],
    ["Urdu", URDU_IQBAL],
  ]) {
    const tag = AshaarWord.contentControlTag(src, {
      layoutMode: "balanced", justifyMode: "kashida", tatweelCount: 4, gapWidth: 4
    });
    assert.match(tag, /^ashaar:/, label + ": tag should start with ashaar:");
    const payload = JSON.parse(decodeURIComponent(tag.replace(/^ashaar:/, "")));
    assert.equal(payload.k, "ashaar-poem", label + ": payload.k");
    assert.ok(payload.sourceHash, label + ": sourceHash present");
  }
  // Different scripts should yield different hashes
  const tagArabic = AshaarWord.contentControlTag(ARABIC_BAYT, {});
  const tagFarsi  = AshaarWord.contentControlTag(FARSI_MASNAVI, {});
  const tagUrdu   = AshaarWord.contentControlTag(URDU_IQBAL, {});
  const hashOf = (t) => JSON.parse(decodeURIComponent(t.replace(/^ashaar:/, ""))).sourceHash;
  assert.notEqual(hashOf(tagArabic), hashOf(tagFarsi), "Arabic ≠ Farsi hash");
  assert.notEqual(hashOf(tagFarsi),  hashOf(tagUrdu),  "Farsi ≠ Urdu hash");
}

// ── 21. extractMisras — Arabic, Farsi, Urdu ──────────────────────────────────

console.log("  extractMisras: multi-script");
{
  const arabicMisras = AshaarWord.extractMisras(ARABIC_BAYT);
  assert.equal(arabicMisras.length, 2);
  assert.equal(arabicMisras[0], "إذا غامَرتَ في شَرَفٍ مَرومِ");
  assert.equal(arabicMisras[1], "فَلا تَقنَع بما دونَ النُّجومِ");

  const farsiMisras = AshaarWord.extractMisras(FARSI_MASNAVI);
  assert.equal(farsiMisras.length, 2);
  assert.equal(farsiMisras[0], "بشنو این نی چون شکایت می‌کند");
  assert.equal(farsiMisras[1], "از جدایی‌ها حکایت می‌کند");

  const marsiyaMisras = AshaarWord.extractMisras(URDU_MARSIYA);
  assert.equal(marsiyaMisras.length, 6);
  assert.equal(marsiyaMisras[0], "شاه كے اصحاب تھے");
  assert.equal(marsiyaMisras[3], "هو گئے شہ پر فدا");
}

// ── 22. generateBareGrid12Ooxml ───────────────────────────────────────────────

console.log("  generateBareGrid12Ooxml: structure");
{
  const ooxml = AshaarWord.generateBareGrid12Ooxml(9360);
  assert.match(ooxml, /<w:tbl>/);
  assert.match(ooxml, /<w:bidiVisual\/>/);
  // Exactly 12 equal grid columns
  assert.equal(countMatches(ooxml, /<w:gridCol /g), 12);
  // 12 cells in the single row
  assert.equal(countMatches(ooxml, /<w:tc>/g), 12);
  // Each cell spans exactly 1 grid column
  const spans = [...ooxml.matchAll(/w:gridSpan w:val="(\d+)"/g)].map(m => Number(m[1]));
  assert.equal(spans.length, 12);
  assert.ok(spans.every(s => s === 1), "All 12 cells should span 1 col each");
  // Thin visible borders
  assert.match(ooxml, /w:val="single"/);
  // RTL bidi
  assert.match(ooxml, /<w:bidi\/>/);

  // Column width: 9360 twips / 12 = 780 twips each
  assert.match(ooxml, /w:tcW w:w="780"/);

  // Falls back to 9360 when no width given
  const defaultOoxml = AshaarWord.generateBareGrid12Ooxml(0);
  assert.equal(countMatches(defaultOoxml, /<w:gridCol /g), 12);
}

// ── 23. templateToOoxml — various row configs ─────────────────────────────────

console.log("  templateToOoxml: row configurations");
{
  // Standard bayt layout: sadr(5) + gap(2) + ajuz(5)
  const baytTemplate = {
    columnCount: 12,
    rows: [
      [{ span: 5 }, { span: 2 }, { span: 5 }],
    ]
  };
  const baytOoxml = AshaarWord.templateToOoxml(baytTemplate, 9360, {});
  assert.match(baytOoxml, /<w:tbl>/);
  assert.equal(countMatches(baytOoxml, /<w:gridCol /g), 12);
  assert.equal(countMatches(baytOoxml, /<w:tc>/g), 3);
  // Check spans: 5, 2, 5
  const spans = [...baytOoxml.matchAll(/w:gridSpan w:val="(\d+)"/g)].map(m => Number(m[1]));
  assert.deepEqual(spans, [5, 2, 5]);
  // Widths: 5*780, 2*780, 5*780
  assert.match(baytOoxml, /w:tcW w:w="3900"/); // 5 * 780
  assert.match(baytOoxml, /w:tcW w:w="1560"/); // 2 * 780

  // Multi-row: bayt row + full-width refrain
  const marsiyaTemplate = {
    columnCount: 12,
    rows: [
      [{ span: 4 }, { span: 4 }, { span: 4 }],  // triple row
      [{ span: 12 }],                              // full-width solo
    ]
  };
  const marsiyaOoxml = AshaarWord.templateToOoxml(marsiyaTemplate, 9360, {});
  assert.equal(countMatches(marsiyaOoxml, /<w:tr>/g), 2);
  assert.equal(countMatches(marsiyaOoxml, /<w:tc>/g), 4); // 3 + 1
  const marsiyaSpans = [...marsiyaOoxml.matchAll(/w:gridSpan w:val="(\d+)"/g)].map(m => Number(m[1]));
  assert.deepEqual(marsiyaSpans, [4, 4, 4, 12]);

  // Nastaliq font applied to runs
  const nastaliqOoxml = AshaarWord.templateToOoxml(baytTemplate, 9360, { fontMode: "nastaliq" });
  assert.match(nastaliqOoxml, /Noto Nastaliq Urdu/);

  // Arabic serif font
  const arabicOoxml = AshaarWord.templateToOoxml(baytTemplate, 9360, { fontMode: "arabic-serif" });
  assert.match(arabicOoxml, /Scheherazade New/);
}

// ── 24. wrapOoxml — both html and ooxml wrappers ─────────────────────────────

console.log("  wrapOoxml: package structure");
{
  const wrapped = AshaarWord.wrapOoxml("<w:p/>");
  assert.match(wrapped, /<?xml version/);
  assert.match(wrapped, /pkg:package/);
  assert.match(wrapped, /word\/document\.xml/);
  assert.match(wrapped, /<w:body>/);
  assert.match(wrapped, /<w:sectPr\/>/);
  assert.match(wrapped, /<w:p\/>/);
}

// ── 25. Tab-stop path — Arabic, Farsi, Urdu ──────────────────────────────────

console.log("  Tab-stop path: Arabic, Farsi, Urdu couplets");
{
  for (const [label, src] of [
    ["Arabic",  ARABIC_BAYT],
    ["Farsi",   FARSI_MASNAVI],
    ["Urdu",    "دل ناداں تجھے ہوا کیا ہے \\ آخر اس درد کی دوا کیا ہے"],
  ]) {
    const ooxml = AshaarTabStop.poemToOoxml(src, { justifyMode: "none" }, Ashaar);
    assert.match(ooxml, /<w:tab w:val="left" w:pos="0"/, label + ": left tab stop");
    assert.match(ooxml, /<w:bidi\/>/, label + ": bidi paragraph");
    const misras = AshaarWord.extractMisras(src);
    assert.match(ooxml, new RegExp(misras[0].replace(/[.*+?^${}()|[\]\\]/g, "\\$&")),
      label + ": sadr text present");
    assert.match(ooxml, new RegExp(misras[1].replace(/[.*+?^${}()|[\]\\]/g, "\\$&")),
      label + ": ajuz text present");
  }
}

// ── 26. Tab-stop — Urdu Marsiya (3-misra) ────────────────────────────────────

console.log("  Tab-stop path: Urdu Marsiya 3-misra");
{
  const ooxml = AshaarTabStop.poemToOoxml(URDU_MARSIYA, { justifyMode: "none" }, Ashaar);
  assert.match(ooxml, /<w:tab w:val="left" w:pos="0"/);
  assert.match(ooxml, /<w:tab w:val="center" w:pos="4680"/); // N=3 center stop
  assert.match(ooxml, /شاه كے اصحاب تھے/);
  assert.match(ooxml, /هو گئے شہ پر فدا/);
  assert.match(ooxml, /هائے كربلاء والو/);
}

// ── 27. Tab-stop — multi-poem kashida ────────────────────────────────────────

console.log("  Tab-stop path: kashida across Arabic + Farsi");
{
  const ooxml = AshaarTabStop.poemToOoxml(MULTI_POEM,
    { justifyMode: "kashida", tatweelCount: 2 }, Ashaar);
  assert.match(ooxml, /ـ/);  // tatweels present
}

// ── 28. layoutTablesForPoem — Farsi multi-misra ──────────────────────────────

console.log("  layoutTablesForPoem: Farsi 3-misra stanza");
{
  const farsiTriple = [
    "بیا که قصر امل سخت سست بنیاد است \\ بیار باده که بنیاد عمر بر باد است \\ غلام همت آنم \\",
    "ز هر چه رنگ تعلق پذیرد آزاد است \\",
  ].join("\n");

  const tables = AshaarWord.layoutTablesForPoem(farsiTriple, {}, Ashaar);
  assert.ok(tables, "should return tables for 3-misra stanza");
  assert.equal(tables.length, 1);
  assert.equal(tables[0].columnCount, 5); // M = 2*3-1 = 5
  assert.equal(tables[0].rows[0][0].text, "بیا که قصر امل سخت سست بنیاد است");
  assert.equal(tables[0].rows[0][2].text, "بیار باده که بنیاد عمر بر باد است");
  assert.equal(tables[0].rows[0][4].text, "غلام همت آنم");
}

// ── 29. misraSpans — Arabic long-short mix ───────────────────────────────────

console.log("  misraSpans: proportional allocation");
{
  // Without canvas: falls back to visibleWeight
  const shortArabic = "نورٌ";
  const longArabic  = "على قدرِ أهلِ العزمِ تأتي العزائمُ";
  const spans = AshaarWord.misraSpans([shortArabic, longArabic], 6, null);
  assert.equal(spans[0] + spans[1], 6, "spans sum to contentCols");
  assert.ok(spans[1] > spans[0], "longer Arabic text gets more columns");
}

// ── 30. Edge cases ────────────────────────────────────────────────────────────

console.log("  Edge cases: empty input, single misra, whitespace");
{
  // Empty → empty string
  assert.equal(AshaarWord.renderForWord("", {}, Ashaar), "");
  assert.equal(AshaarWord.renderForWordOoxml("", {}, Ashaar, 9360), "");

  // Single misra only → centered solo row
  const singleHtml = AshaarWord.renderForWord("| تنها", { justifyMode: "none" }, Ashaar);
  assert.match(singleHtml, /تنها/);
  assert.match(singleHtml, /colspan="12"/);

  // Whitespace-only lines are skipped
  const wsSource = "  \n  \n  ";
  assert.equal(AshaarWord.extractMisras(wsSource).length, 0);
}

console.log("poetry-corpus tests passed");

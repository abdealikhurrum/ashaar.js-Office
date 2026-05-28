const assert = require("assert");
const Ashaar = require("../src/vendor/ashaar");
const AshaarTabStop = require("../src/taskpane/word-tabstop");

// ── Tab stop geometry ───────────────────────────────────────────────────────

// N=2: one RIGHT stop only (standard maqta: ajuz at left, sadr right-aligns)
const stops2 = AshaarTabStop.tabStopsForN(2);
assert.equal(stops2.length, 1);
assert.equal(stops2[0].val, "right");
assert.equal(stops2[0].pos, 9360);

// N=3: CENTER at page midpoint + RIGHT at right margin
const stops3 = AshaarTabStop.tabStopsForN(3);
assert.equal(stops3.length, 2);
assert.equal(stops3[0].val, "center");
assert.equal(stops3[0].pos, 4680); // (2×1+1) × 9360/(2×3) = 3×3120/2 = 4680
assert.equal(stops3[1].val, "right");
assert.equal(stops3[1].pos, 9360);

// tabStopsForN respects a custom textWidth (e.g. A4 with narrow margins)
const stopsCustom = AshaarTabStop.tabStopsForN(3, 8640); // hypothetical 6-inch text width
assert.equal(stopsCustom[0].pos, 4320); // midpoint of 8640 for N=3
assert.equal(stopsCustom[1].pos, 8640); // right margin

// N=5: three CENTER stops + one RIGHT stop
const stops5 = AshaarTabStop.tabStopsForN(5);
assert.equal(stops5.length, 4);
assert.equal(stops5[0].val, "center");
assert.equal(stops5[0].pos, 2808); // (2×1+1) × 9360/(2×5) = 3×936 = 2808
assert.equal(stops5[1].val, "center");
assert.equal(stops5[1].pos, 4680); // (2×2+1) × 9360/(2×5) = 5×936 = 4680
assert.equal(stops5[2].val, "center");
assert.equal(stops5[2].pos, 6552); // (2×3+1) × 9360/(2×5) = 7×936 = 6552
assert.equal(stops5[3].val, "right");
assert.equal(stops5[3].pos, 9360);

// ── Regular 2-misra poem ───────────────────────────────────────────────────

const source2 = "دل ناداں تجھے ہوا کیا ہے \\ آخر اس درد کی دوا کیا ہے";
const ooxml2 = AshaarTabStop.poemToOoxml(source2, { justifyMode: "none" }, Ashaar);

assert.match(ooxml2, /<w:tab w:val="right" w:pos="9360"/);
assert.doesNotMatch(ooxml2, /<w:tab w:val="center"/); // N=2: no interior center stop
assert.match(ooxml2, /دل ناداں تجھے ہوا کیا ہے/);    // ajuz (leftmost in XML)
assert.match(ooxml2, /آخر اس درد کی دوا کیا ہے/);    // sadr (after TAB)
// LTR column order in XML: ajuz (col 0 = left = second source part) before sadr (col 1 = right = first source part)
// source: "sadr \\ ajuz" → sadr="دل ناداں...", ajuz="آخر اس درد..."
assert.ok(
  ooxml2.indexOf("آخر اس درد کی دوا کیا ہے") < ooxml2.indexOf("دل ناداں تجھے ہوا کیا ہے"),
  "ajuz (left col) before sadr (right col) in XML"
);

// ── 3-misra marsiya stanza ─────────────────────────────────────────────────

const marsiyaSource = [
  "شاه كے اصحاب تھے \\ خلق ميں الباب تھے \\ صدق كے ارباب تھے \\",
  "هو گئے شہ پر فدا \\",
  "هائے كربلاء والو \\ هائے كربلاء والو"
].join("\n");

const ooxml3 = AshaarTabStop.poemToOoxml(marsiyaSource, { justifyMode: "none" }, Ashaar);

// Tab stops for N=3
assert.match(ooxml3, /<w:tab w:val="center" w:pos="4680"/);
assert.match(ooxml3, /<w:tab w:val="right" w:pos="9360"/);

// All three misras present
assert.match(ooxml3, /شاه كے اصحاب تھے/); // misras[0] = sadr (rightmost col)
assert.match(ooxml3, /خلق ميں الباب تھے/); // misras[1] = middle col
assert.match(ooxml3, /صدق كے ارباب تھے/); // misras[2] = ajuz (leftmost col)

// Triple row: LTR column order in XML — ajuz (صدق) first, middle (خلق) second, sadr (شاه) last
const ajuzIdx  = ooxml3.indexOf("صدق كے ارباب تھے");
const midIdx   = ooxml3.indexOf("خلق ميں الباب تھے");
const sadrIdx  = ooxml3.indexOf("شاه كے اصحاب تھے");
assert.ok(ajuzIdx < midIdx,  "ajuz before middle in XML");
assert.ok(midIdx  < sadrIdx, "middle before sadr in XML");

// Solo line: uses centered paragraph, not tab stops
assert.match(ooxml3, /<w:jc w:val="center"/);
assert.match(ooxml3, /هو گئے شہ پر فدا/);

// Maqta (regular 2-misra bayt): uses tab-stop paragraph (not centered)
assert.match(ooxml3, /هائے كربلاء والو/);

// ── 5-misra poem ──────────────────────────────────────────────────────────

const fiveMisraSource = "م1 \\ م2 \\ م3 \\ م4 \\ م5\nنعرہ \\";
const ooxml5 = AshaarTabStop.poemToOoxml(fiveMisraSource, { justifyMode: "none" }, Ashaar);

// Tab stops for N=5
assert.match(ooxml5, /<w:tab w:val="center" w:pos="2808"/);
assert.match(ooxml5, /<w:tab w:val="center" w:pos="4680"/);
assert.match(ooxml5, /<w:tab w:val="center" w:pos="6552"/);
assert.match(ooxml5, /<w:tab w:val="right" w:pos="9360"/);

// All 5 misras present
for (var i = 1; i <= 5; i++) {
  assert.match(ooxml5, new RegExp("م" + i), "misra م" + i + " missing");
}

// LTR column order: م5 (ajuz/leftmost) … م1 (sadr/rightmost)
var m5Idx = ooxml5.indexOf("م5");
var m1Idx = ooxml5.lastIndexOf("م1"); // lastIndexOf avoids hitting "م10" etc.
assert.ok(m5Idx < m1Idx, "م5 (ajuz) before م1 (sadr) in XML");

// Solo refrain: centered paragraph
assert.match(ooxml5, /<w:jc w:val="center"/);
assert.match(ooxml5, /نعرہ/);

// ── OOXML wrapper ──────────────────────────────────────────────────────────

const wrapped = AshaarTabStop.wrapOoxml("<w:p/>");
assert.match(wrapped, /<?xml version/);
assert.match(wrapped, /<w:document/);
assert.match(wrapped, /<w:body>/);
assert.match(wrapped, /<w:sectPr\/>/);
assert.match(wrapped, /<w:p\/>/);

// ── Kashida justification ─────────────────────────────────────────────────

const justified = AshaarTabStop.poemToOoxml(
  "سلام \\ دنیا",
  { justifyMode: "kashida", tatweelCount: 3 },
  Ashaar
);
assert.match(justified, /ـ/, "kashida tatweel should appear");

console.log("word-tabstop tests passed");

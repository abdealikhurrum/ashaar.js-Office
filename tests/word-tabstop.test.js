const assert = require("assert");
const Ashaar = require("../src/vendor/ashaar");
const AshaarTabStop = require("../src/taskpane/word-tabstop");

// ── Tab stop geometry ───────────────────────────────────────────────────────

// RTL paragraph tab stops: sadr at right (no stop), ajuz anchors to left via LEFT stop at 0.
// Intermediate columns use CENTER stops symmetric around W/2.

// N=2: one LEFT stop at 0 (ajuz at left margin)
const stops2 = AshaarTabStop.tabStopsForN(2);
assert.equal(stops2.length, 1);
assert.equal(stops2[0].val, "left");
assert.equal(stops2[0].pos, 0);

// N=3: LEFT at 0 + CENTER at W/2 (page midpoint for middle column)
const stops3 = AshaarTabStop.tabStopsForN(3);
assert.equal(stops3.length, 2);
assert.equal(stops3[0].val, "left");
assert.equal(stops3[0].pos, 0);
assert.equal(stops3[1].val, "center");
assert.equal(stops3[1].pos, 4680); // W - 1.5×colW = 9360 - 4680 = 4680

// tabStopsForN respects a custom textWidth (e.g. A4 with narrow margins)
const stopsCustom = AshaarTabStop.tabStopsForN(3, 8640); // hypothetical 6-inch text width
assert.equal(stopsCustom[0].pos, 0);    // left stop for ajuz
assert.equal(stopsCustom[1].pos, 4320); // center of middle column: 8640 - 1.5×2880 = 4320

// N=5: LEFT at 0 + three CENTER stops going rightward (ascending order)
const stops5 = AshaarTabStop.tabStopsForN(5);
assert.equal(stops5.length, 4);
assert.equal(stops5[0].val, "left");
assert.equal(stops5[0].pos, 0);
assert.equal(stops5[1].val, "center");
assert.equal(stops5[1].pos, 2808); // W - 3.5×colW = 9360 - 6552 = 2808
assert.equal(stops5[2].val, "center");
assert.equal(stops5[2].pos, 4680); // W - 2.5×colW = 9360 - 4680 = 4680
assert.equal(stops5[3].val, "center");
assert.equal(stops5[3].pos, 6552); // W - 1.5×colW = 9360 - 2808 = 6552

// ── Regular 2-misra poem ───────────────────────────────────────────────────

const source2 = "دل ناداں تجھے ہوا کیا ہے \\ آخر اس درد کی دوا کیا ہے";
const ooxml2 = AshaarTabStop.poemToOoxml(source2, { justifyMode: "none" }, Ashaar);

assert.match(ooxml2, /<w:tab w:val="left" w:pos="0"/);
assert.doesNotMatch(ooxml2, /<w:tab w:val="center"/); // N=2: no interior center stop
assert.match(ooxml2, /دل ناداں تجھے ہوا کیا ہے/);    // sadr (first in RTL XML = right margin)
assert.match(ooxml2, /آخر اس درد کی دوا کیا ہے/);    // ajuz (after TAB = left margin)
// RTL column order in XML: sadr (col 0 = right) before ajuz (col 1 = left)
// source: "sadr \\ ajuz" → sadr="دل ناداں...", ajuz="آخر اس درد..."
assert.ok(
  ooxml2.indexOf("دل ناداں تجھے ہوا کیا ہے") < ooxml2.indexOf("آخر اس درد کی دوا کیا ہے"),
  "sadr (right col) before ajuz (left col) in RTL XML"
);

// ── 3-misra marsiya stanza ─────────────────────────────────────────────────

const marsiyaSource = [
  "شاه كے اصحاب تھے \\ خلق ميں الباب تھے \\ صدق كے ارباب تھے \\",
  "هو گئے شہ پر فدا \\",
  "هائے كربلاء والو \\ هائے كربلاء والو"
].join("\n");

const ooxml3 = AshaarTabStop.poemToOoxml(marsiyaSource, { justifyMode: "none" }, Ashaar);

// Tab stops for N=3 (RTL: LEFT at 0, CENTER at W/2)
assert.match(ooxml3, /<w:tab w:val="left" w:pos="0"/);
assert.match(ooxml3, /<w:tab w:val="center" w:pos="4680"/);

// All three misras present
assert.match(ooxml3, /شاه كے اصحاب تھے/); // misras[0] = sadr (col 0 = right in RTL)
assert.match(ooxml3, /خلق ميں الباب تھے/); // misras[1] = middle col
assert.match(ooxml3, /صدق كے ارباب تھے/); // misras[2] = ajuz (col N-1 = left in RTL)

// Triple row: RTL column order in XML — sadr (شاه) first, middle (خلق) second, ajuz (صدق) last
const sadrIdx  = ooxml3.indexOf("شاه كے اصحاب تھے");
const midIdx   = ooxml3.indexOf("خلق ميں الباب تھے");
const ajuzIdx  = ooxml3.indexOf("صدق كے ارباب تھے");
assert.ok(sadrIdx < midIdx,  "sadr before middle in RTL XML");
assert.ok(midIdx  < ajuzIdx, "middle before ajuz in RTL XML");

// Solo line: uses centered RTL paragraph
assert.match(ooxml3, /<w:bidi\/>/);
assert.match(ooxml3, /<w:jc w:val="center"/);
assert.match(ooxml3, /هو گئے شہ پر فدا/);

// Maqta (regular 2-misra bayt): uses tab-stop paragraph (not centered)
assert.match(ooxml3, /هائے كربلاء والو/);

// ── 5-misra poem ──────────────────────────────────────────────────────────

const fiveMisraSource = "م1 \\ م2 \\ م3 \\ م4 \\ م5\nنعرہ \\";
const ooxml5 = AshaarTabStop.poemToOoxml(fiveMisraSource, { justifyMode: "none" }, Ashaar);

// Tab stops for N=5 (RTL: LEFT at 0 + three CENTER stops)
assert.match(ooxml5, /<w:tab w:val="left" w:pos="0"/);
assert.match(ooxml5, /<w:tab w:val="center" w:pos="2808"/);
assert.match(ooxml5, /<w:tab w:val="center" w:pos="4680"/);
assert.match(ooxml5, /<w:tab w:val="center" w:pos="6552"/);

// All 5 misras present
for (var i = 1; i <= 5; i++) {
  assert.match(ooxml5, new RegExp("م" + i), "misra م" + i + " missing");
}

// RTL column order: م1 (sadr/rightmost) first … م5 (ajuz/leftmost) last
var m1Idx = ooxml5.indexOf("م1");
var m5Idx = ooxml5.lastIndexOf("م5");
assert.ok(m1Idx < m5Idx, "م1 (sadr) before م5 (ajuz) in RTL XML");

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

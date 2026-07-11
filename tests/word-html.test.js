const assert = require("assert");
const Ashaar = require("../src/vendor/ashaar");
const AshaarWord = require("../src/taskpane/word-html");

const source = "دل ناداں تجھے ہوا کیا ہے \\ آخر اس درد کی دوا کیا ہے";

const html = AshaarWord.renderForWord(source, {
  justifyMode: "kashida",
  tatweelCount: 4,
  gapWidth: 4,
  layoutMode: "balanced",
  widthMode: "optimized"
}, Ashaar);

assert.match(html, /<table dir="rtl"/);
assert.match(html, /<colgroup>/);
assert.match(html, /data-ashaar-layout="balanced"/);
// 12-column grid: <td> now has colspan before style
assert.match(html, /text-align:left/);
assert.match(html, /text-align:right/);
assert.match(html, /ـ/);

// N=2 → 2 cells (sadr colspan=6, ajuz colspan=6); no gap cell
const firstRow = html.match(/<tr>(.*?)<\/tr>/)[1];
const cells = firstRow.match(/<td[^>]*style="[^"]*"/g);
assert.equal(cells.length, 2);
assert.match(cells[0], /text-align:right/); // sadr (visual right in RTL)
assert.match(cells[1], /text-align:left/);  // ajuz (visual left in RTL)

const stacked = AshaarWord.renderForWord(source, {
  justifyMode: "none",
  layoutMode: "stacked"
}, Ashaar);

assert.match(stacked, /colspan="12"/); // stacked mode: single cell spans full 12-col grid
assert.match(stacked, /<br>/);

const compact = AshaarWord.renderForWord(source, {
  justifyMode: "none",
  layoutMode: "compact",
  gapWidth: 4
}, Ashaar);

assert.match(compact, /colspan="5"|<col style="width:8%"/);

const tag = AshaarWord.contentControlTag(source, {
  layoutMode: "balanced",
  widthMode: "optimized",
  justifyMode: "kashida",
  tatweelCount: 6,
  gapWidth: 4
});

assert.match(tag, /^ashaar:/);
const payload = JSON.parse(decodeURIComponent(tag.replace(/^ashaar:/, "")));
assert.equal(payload.k, "ashaar-poem");
assert.equal(payload.layoutMode, "balanced");
assert.equal(payload.fontMode, "document");
assert.ok(payload.sourceHash);

const template = AshaarWord.renderTemplateForWord({
  bandhCount: 2,
  misraCount: 4,
  misraPattern: "paired",
  layoutMode: "equal",
  widthMode: "fixed",
  fontMode: "document"
});

assert.equal((template.match(/data-ashaar-template="true"/g) || []).length, 2);
assert.equal((template.match(/<tr>/g) || []).length, 4);
assert.doesNotMatch(template, /font-family/);
const templateCells = template.match(/<td style="[^"]*"/g);
assert.match(templateCells[0], /text-align:right/);
assert.match(templateCells[1], /text-align:center/);
assert.match(templateCells[2], /text-align:left/);
assert.match(template, />1<\/td>/);
assert.match(template, />2<\/td>/);

assert.deepEqual(AshaarWord.templateGrid({ misraCount: 6, misraPattern: "four-plus-centered" }), [
  { type: "pair", right: 1, left: 2 },
  { type: "pair", right: 3, left: 4 },
  { type: "center", misra: 5, align: "center" },
  { type: "center", misra: 6, align: "center" }
]);

assert.deepEqual(AshaarWord.templateGrid({ misraCount: 4, misraPattern: "alternate-right" }), [
  { type: "right", misra: 1 },
  { type: "left", misra: 2 },
  { type: "right", misra: 3 },
  { type: "left", misra: 4 }
]);

assert.deepEqual(AshaarWord.templateGrid({ misraCount: 6, misraPattern: "three-plus-center-refrain" }), [
  { type: "triple", right: 1, middle: 2, left: 3 },
  { type: "center", misra: 4, align: "center", colspan: 4 },
  { type: "refrain", right: 5, left: 6 }
]);

assert.deepEqual(AshaarWord.templateGrid({ misraCount: 6, misraPattern: "multi-misra-row" }), [
  { type: "triple", right: 1, middle: 2, left: 3 },
  { type: "center", misra: 4, align: "center", colspan: 4 },
  { type: "refrain", right: 5, left: 6 }
]);

assert.deepEqual(AshaarWord.parseLayoutSpec("3 | 2 | 1\n<4>\n6 - 5"), [
  { type: "multi", cells: ["3", "2", "1"] },
  { type: "center", misra: "4", align: "center", colspan: 4 },
  { type: "pair", left: "6", right: "5" }
]);

const karbalaTemplate = AshaarWord.renderTemplateForWord({
  bandhCount: 3,
  misraCount: 6,
  layoutSpec: "3 | 2 | 1\n<4>\n6 - 5",
  fontMode: "document"
});

assert.equal((karbalaTemplate.match(/data-ashaar-template="true"/g) || []).length, 3);
assert.equal((karbalaTemplate.match(/<tr>/g) || []).length, 9);
assert.match(karbalaTemplate, /<colgroup><col style="width:48%">/);
assert.match(karbalaTemplate, /<td style="[^"]*text-align:right[^"]*">3<\/td><td style="[^"]*text-align:center[^"]*">2<\/td><td style="[^"]*text-align:left[^"]*">1<\/td>/);
assert.match(karbalaTemplate, /<td style="[^"]*text-align:right[^"]*"> <\/td><td style="[^"]*text-align:center[^"]*">4<\/td><td style="[^"]*text-align:left[^"]*"> <\/td>/);
assert.match(karbalaTemplate, />6<\/td><td style="[^"]*text-align:center[^"]*"> <\/td><td style="[^"]*text-align:left[^"]*">5<\/td>/);

const nastaliq = AshaarWord.renderTemplateForWord({
  bandhCount: 1,
  misraCount: 1,
  misraPattern: "paired",
  fontMode: "nastaliq"
});

assert.match(nastaliq, /Noto Nastaliq Urdu/);

const plain = AshaarWord.justifyPlainTextBlock("سلام دنیا", {
  justifyMode: "kashida",
  tatweelCount: 3
});

assert.match(plain, /ـ/);

const marsiya = [
  "شاه كے اصحاب تھے \\ خلق ميں الباب تھے \\ صدق كے ارباب تھے \\",
  "هو گئے شہ پر فدا \\",
  "هائے كربلاء والو \\ هائے كربلاء والو",
  "—",
  "اهلِ بيت تھے عجب \\ اٗفقِ دعوت كے شٗہٗب \\ هو گئے قربان سب \\",
  "خالي لشكر هو گيا \\",
  "هائے كربلاء والو \\ هائے كربلاء والو"
].join("\n");

assert.deepEqual(AshaarWord.extractMisras(marsiya.split("\n—\n")[0]), [
  "شاه كے اصحاب تھے",
  "خلق ميں الباب تھے",
  "صدق كے ارباب تھے",
  "هو گئے شہ پر فدا",
  "هائے كربلاء والو",
  "هائے كربلاء والو"
]);

const marsiyaHtml = AshaarWord.renderTextWithLayoutForWord(marsiya, {
  layoutSpec: "3 | 2 | 1\n<4>\n6 - 5",
  fontMode: "document"
});

assert.equal((marsiyaHtml.match(/<table dir="rtl"/g) || []).length, 2);
assert.match(marsiyaHtml, /صدق كے ارباب تھے/);
assert.match(marsiyaHtml, /خلق ميں الباب تھے/);
assert.match(marsiyaHtml, /شاه كے اصحاب تھے/);
assert.match(marsiyaHtml, /هو گئے شہ پر فدا/);
assert.match(marsiyaHtml, /هائے كربلاء والو/);
assert.doesNotMatch(marsiyaHtml, /colspan=/);

const marsiyaTables = AshaarWord.layoutTablesForText(marsiya, {
  layoutSpec: "3 | 2 | 1\n<4>\n6 - 5",
  fontMode: "document"
});

assert.equal(marsiyaTables.length, 2);
assert.equal(marsiyaTables[0].columnCount, 3);
assert.deepEqual(marsiyaTables[0].rows.map((row) => row.map((cell) => cell.text)), [
  ["صدق كے ارباب تھے", "خلق ميں الباب تھے", "شاه كے اصحاب تھے"],
  ["", "هو گئے شہ پر فدا", ""],
  ["هائے كربلاء والو", "", "هائے كربلاء والو"]
]);
assert.deepEqual(marsiyaTables[0].rows[0].map((cell) => cell.align), ["right", "center", "left"]);

// Test: multi-misra rows are auto-detected in renderForWord with dynamic column count

function multiMisraTest(lineCount, misraCount) {
  var line = [];
  for (var i = 0; i < misraCount; i++) line.push("م" + (i + 1));
  var source = line.join(" \\ ");
  // Stanza: one multi-misra line + one solo refrain
  var fullSource = source + "\nنعرہ \\";

  var html = AshaarWord.renderForWord(fullSource, { justifyMode: "none", layoutMode: "balanced" }, Ashaar);

  // All misras must appear
  for (var j = 0; j < misraCount; j++) {
    assert.match(html, new RegExp("م" + (j + 1)), "misra " + (j + 1) + " missing for " + misraCount + "-misra line");
  }

  // 12-column grid: always 12 <col> elements
  var cols = (html.match(/<col style=/g) || []).length;
  assert.equal(cols, 12, "Expected 12 grid columns for " + misraCount + "-misra stanza");

  // Multi-misra row has exactly misraCount cells (each spanning colsPerMisra cols)
  var rowCells = html.match(/<tr>(<td[^>]*>[\s\S]*?<\/td>){1,}<\/tr>/g) || [];
  var firstRowCellCount = (rowCells[0].match(/<td/g) || []).length;
  assert.equal(firstRowCellCount, misraCount, "Expected " + misraCount + " cells in full row");

  // Solo misra row spans the full 12-column grid
  assert.match(html, /colspan="12"/, "Solo row should span full 12-col grid");
}

multiMisraTest(1, 3);
multiMisraTest(1, 4);
multiMisraTest(1, 5);

// Marsiya stanza with 3-misra lines + solo + maqta pair
const marsiyaSource = [
  "شاه كے اصحاب تھے \\ خلق ميں الباب تھے \\ صدق كے ارباب تھے \\",
  "هو گئے شہ پر فدا \\",
  "هائے كربلاء والو \\ هائے كربلاء والو"
].join("\n");

const marsiyaWordHtml = AshaarWord.renderForWord(marsiyaSource, {
  justifyMode: "none",
  layoutMode: "balanced"
}, Ashaar);

assert.match(marsiyaWordHtml, /شاه كے اصحاب تھے/);
assert.match(marsiyaWordHtml, /خلق ميں الباب تھے/);
assert.match(marsiyaWordHtml, /صدق كے ارباب تھے/);
assert.match(marsiyaWordHtml, /هائے كربلاء والو/);
assert.equal((marsiyaWordHtml.match(/<col style=/g) || []).length, 12); // 12-column grid
assert.match(marsiyaWordHtml, /colspan="12"/); // solo misra spans full grid

// Test: layoutTablesForPoem — native table path for multi-misra poems
// Uses 2N-1 interleaved columns: content at even indices (0,2,4), gap at odd indices (1,3)
const poemTables3 = AshaarWord.layoutTablesForPoem(marsiyaSource, { justifyMode: "none" }, Ashaar);
assert.ok(poemTables3, "Should return tables for multi-misra poem");
assert.equal(poemTables3.length, 1);
assert.equal(poemTables3[0].columnCount, 5); // M = 2*3-1 = 5 cols (3 content + 2 gap)
assert.equal(poemTables3[0].rows.length, 3);
// Triple row: col 0 = misras[0] (sadr, visual right in RTL), col 2 = misras[1], col 4 = misras[2] (ajuz, visual left)
assert.equal(poemTables3[0].rows[0][0].text, "شاه كے اصحاب تھے"); // col 0 = misras[0] = sadr
assert.equal(poemTables3[0].rows[0][1].text, "");                   // col 1 = gap
assert.equal(poemTables3[0].rows[0][2].text, "خلق ميں الباب تھے"); // col 2 = misras[1]
assert.equal(poemTables3[0].rows[0][3].text, "");                   // col 3 = gap
assert.equal(poemTables3[0].rows[0][4].text, "صدق كے ارباب تھے"); // col 4 = misras[2] = ajuz
assert.equal(poemTables3[0].rows[0][0].align, "right");
assert.equal(poemTables3[0].rows[0][2].align, "center");
assert.equal(poemTables3[0].rows[0][4].align, "left");
// Solo row: text in center content col (col 2 for N=3), others empty
assert.equal(poemTables3[0].rows[1][2].text, "هو گئے شہ پر فدا");
assert.equal(poemTables3[0].rows[1][0].text, "");
// Maqta row: sadr in col 0 (visual right in RTL), ajuz in col 4 — both are same refrain text here
assert.equal(poemTables3[0].rows[2][0].text, "هائے كربلاء والو");
assert.equal(poemTables3[0].rows[2][4].text, "هائے كربلاء والو");

// Regular 2-misra poem returns null (falls through to renderForWord)
const regularPoem = "دل ناداں تجھے ہوا کیا ہے \\ آخر اس درد کی دوا کیا ہے";
const poemTablesRegular = AshaarWord.layoutTablesForPoem(regularPoem, {}, Ashaar);
assert.equal(poemTablesRegular, null, "Regular couplet poem should return null");

// 5-misra poem: M = 2*5-1 = 9 cols (5 content + 4 gap)
const fiveMisraSource = "م1 \\ م2 \\ م3 \\ م4 \\ م5\nنعرہ \\";
const poemTables5 = AshaarWord.layoutTablesForPoem(fiveMisraSource, {}, Ashaar);
assert.ok(poemTables5);
assert.equal(poemTables5[0].columnCount, 9); // M = 2*5-1 = 9
assert.equal(poemTables5[0].rows[0].length, 9);
assert.equal(poemTables5[0].rows[0][0].text, "م1"); // misras[0] (sadr) → col 0 = visual right in RTL

// ── renderForWordOoxml ─────────────────────────────────────────────────────

// 2-misra poem: N=2, gapCols=1 → GRID=13 (2*6+1*1=13)
const ooxml2misra = AshaarWord.renderForWordOoxml(
  "دل ناداں تجھے ہوا کیا ہے \\ آخر اس درد کی دوا کیا ہے",
  { justifyMode: "none", gapWidth: 1 }, Ashaar, 9360
);
assert.match(ooxml2misra, /<w:tbl>/);
assert.match(ooxml2misra, /<w:bidiVisual\/>/);
assert.match(ooxml2misra, /<w:gridSpan w:val="/);
// GRID=13 columns: 13 gridCol elements
assert.equal((ooxml2misra.match(/<w:gridCol /g) || []).length, 13);
// Two misra cells + one gap cell = 3 tc elements in the row (gridSpan sums to 7)
const tc2 = (ooxml2misra.match(/<w:tc>/g) || []).length;
assert.equal(tc2, 3, "N=2: 2 misra cells + 1 gap cell");
assert.match(ooxml2misra, /دل ناداں تجھے ہوا کیا ہے/);
assert.match(ooxml2misra, /آخر اس درد کی دوا کیا ہے/);

// 3-misra marsiya stanza: one table with GRID=20 for all rows
// Row 1 (3-misra): spans=[6,1,6,1,6]=20; Row 2 (solo): padded; Row 3 (2-misra): spans=[10,1,9]=20
const ooxml3misra = AshaarWord.renderForWordOoxml(marsiyaSource,
  { justifyMode: "none", gapWidth: 1 }, Ashaar, 9360
);
assert.equal((ooxml3misra.match(/<w:gridCol /g) || []).length, 20); // single GRID=20 table
assert.match(ooxml3misra, /<w:bidiVisual\/>/);
assert.match(ooxml3misra, /شاه كے اصحاب تھے/);
assert.match(ooxml3misra, /هو گئے شہ پر فدا/); // solo row (padded, centered)
assert.match(ooxml3misra, /<w:jc w:val="center"\/>/); // solo paragraph centered

// ── layoutTableToOoxml: non-span (Numbers-view) tables must be RTL ─────────
// The "Numbers" layout used to be inserted via the native Word.insertTable API,
// which yields an LTR table (cell/tab order runs left-to-right) even when the
// content is arranged to look right. Routing it through OOXML with a
// <w:bidiVisual/> flag makes it a genuine RTL table, matching the Grid path.
const layoutOoxml = AshaarWord.layoutTableToOoxml(
  { columnCount: 2, widths: [50, 50], rows: [[{ text: "١", align: "right" }, { text: "٢", align: "left" }]] },
  9360, {}
);
assert.match(layoutOoxml, /<w:tbl>/);
assert.match(layoutOoxml, /<w:bidiVisual\/>/, "non-span layout table must be a visually RTL table");
assert.equal((layoutOoxml.match(/<w:gridCol /g) || []).length, 2, "one gridCol per column");
assert.match(layoutOoxml, /<w:gridCol w:w="4680"\/>/, "50% column of a 9360-twip width = 4680");
// Cells stay in LOGICAL order (col 0 first); bidiVisual flips them visually, not the source.
const firstMisraAt = layoutOoxml.indexOf("١");
const secondMisraAt = layoutOoxml.indexOf("٢");
assert.ok(firstMisraAt !== -1 && secondMisraAt !== -1 && firstMisraAt < secondMisraAt,
  "logical-first cell (visual right in RTL) is emitted first");

// misraSpans: proportional allocation
const spans = AshaarWord.misraSpans(["abc", "abcdef"], 6); // weights 3:6 → 2:4
assert.equal(spans[0] + spans[1], 6, "spans must sum to contentCols");
assert.ok(spans[1] > spans[0], "longer text gets more columns");

// ── Stanza-internal independence (couplet must not depend on solo widths) ───
// A stanza is one Word table with a shared grid. Word's default auto-fit resizes
// shared grid columns to content across ALL rows, so a wide solo misra would drag
// the couplet's column split with it. The table must declare fixed layout with a
// definite width so the grid is rigid and each row lays out independently.

function ooxmlStanza(soloText) {
  // 4 solos + 1 couplet + 1 solo, all in one stanza
  const src = [soloText, soloText, soloText, soloText,
    "سادر بيت \\ عجز بيت", soloText].join("\n");
  return AshaarWord.renderForWordOoxml(src, { justifyMode: "none", gapWidth: 1 }, Ashaar, 9360);
}

const stanzaShort = ooxmlStanza("الف");
const stanzaLong = ooxmlStanza("الف الف الف الف الف الف الف الف الف الف");

// Fixed table layout + definite width (the rigidity that removes cross-row coupling)
assert.match(stanzaShort, /<w:tblLayout w:type="fixed"\/>/, "stanza table must use fixed layout");
assert.match(stanzaShort, /<w:tblW w:w="\d+" w:type="dxa"\/>/, "stanza table must have a definite dxa width");
assert.doesNotMatch(stanzaShort, /<w:tblW w:w="0" w:type="auto"\/>/, "stanza table must not be auto-width");

// The couplet row (the one containing both hemistichs) must be byte-identical
// regardless of how wide the sibling solo misras are.
function coupletRow(ooxml) {
  const rows = ooxml.match(/<w:tr>[\s\S]*?<\/w:tr>/g) || [];
  return rows.filter((r) => /سادر بيت/.test(r) && /عجز بيت/.test(r))[0];
}
const rowShort = coupletRow(stanzaShort);
const rowLong = coupletRow(stanzaLong);
assert.ok(rowShort && rowLong, "couplet row must be found in both renders");
assert.equal(rowShort, rowLong, "couplet row must be independent of sibling solo widths");

// ── Justify engine idempotency (vendored ashaar-justify) ────────────────────
// Re-justification must re-derive from the bare line: reducible and idempotent,
// never compounding on previously inserted tatweels.
const AshaarJustify = require("../src/vendor/ashaar-justify");
const _ctx = { measureText: (s) => ({ width: s.replace(/\s/g, "").length }) };
const _base = "قفا نبك من ذكرى";
const _wide = AshaarJustify.justifyLine(_base, 30, _ctx, { targetFill: 1 });
const _wideN = (_wide.match(/ـ/g) || []).length;
assert.ok(_wideN > 0, "justifyLine adds tatweels for a wide target");
assert.ok((AshaarJustify.justifyLine(_wide, 18, _ctx, { targetFill: 1 }).match(/ـ/g) || []).length < _wideN,
  "re-justify an already-stretched line reduces toward a narrower target");
assert.doesNotMatch(AshaarJustify.justifyLine(_wide, _base.replace(/\s/g, "").length, _ctx, { targetFill: 1 }), /ـ/,
  "re-justify strips back to bare when the line already fits");
assert.equal(
  AshaarJustify.spreadTatweels(AshaarJustify.spreadTatweels("ليلي", 3), 3),
  AshaarJustify.spreadTatweels("ليلي", 3),
  "spreadTatweels is idempotent (no compounding)"
);

// ── Column consistency across rows of a stanza ──────────────────────────────
// Every sadr column must be the same width, and every ajuz column the same width,
// even when rows have different-length hemistichs (the sadr>ajuz asymmetry stays).
{
  const src =
    "هل مظهر ذي العرش سوى صنو الرسول \\ في کل ظہور\n" +
    "أم هل للورى معط سواه \\ في کل ظہور";
  const o = AshaarWord.renderForWordOoxml(src, { justifyMode: "none", gapWidth: 1 }, Ashaar, 9360);
  const rows = o.match(/<w:tr>[\s\S]*?<\/w:tr>/g);
  const spansOf = (r) => [...r.matchAll(/<w:gridSpan w:val="(\d+)"\/>/g)].map((m) => Number(m[1]));
  const r1 = spansOf(rows[0]);
  const r2 = spansOf(rows[1]);
  assert.equal(r1[0], r2[0], "sadr column width is consistent across rows");
  assert.equal(r1[r1.length - 1], r2[r2.length - 1], "ajuz column width is consistent across rows");
  assert.ok(r1[0] > r1[r1.length - 1], "sadr column stays wider than ajuz (asymmetry preserved)");
}

// ── Table width scaling (taskpane scales textWidthTwips; generators must honor it) ──
const wFull = AshaarWord.renderForWordOoxml("a \\ b", { justifyMode: "none", gapWidth: 1 }, Ashaar, 9360);
const wHalf = AshaarWord.renderForWordOoxml("a \\ b", { justifyMode: "none", gapWidth: 1 }, Ashaar, 4680);
const colFull = Number(wFull.match(/<w:gridCol w:w="(\d+)"/)[1]);
const colHalf = Number(wHalf.match(/<w:gridCol w:w="(\d+)"/)[1]);
assert.ok(Math.abs(colHalf * 2 - colFull) <= 2, "gridCol width scales with textWidthTwips");
const tblFull = Number(wFull.match(/<w:tblW w:w="(\d+)"/)[1]);
const tblHalf = Number(wHalf.match(/<w:tblW w:w="(\d+)"/)[1]);
assert.ok(tblHalf < tblFull && Math.abs(tblHalf * 2 - tblFull) <= 14, "tblW scales with textWidthTwips");

// Content-control tag carries the chosen table width
const tagW = AshaarWord.contentControlTag("x", { tableWidthPct: 50 });
assert.equal(
  JSON.parse(decodeURIComponent(tagW.replace(/^ashaar:/, ""))).tableWidthPct, 50,
  "tag carries tableWidthPct"
);

// Content-control tag carries the qaseeda name and round-trips through the parser
const tagQ = AshaarWord.contentControlTag("x", { qaseeda: "Karbala" });
assert.equal(
  AshaarWord.parseContentControlTag(tagQ).qaseeda, "Karbala",
  "tag carries and parses the qaseeda name"
);
const tagNoQ = "ashaar:" + encodeURIComponent(JSON.stringify({ k: "ashaar-poem", v: 1 }));
assert.equal(
  AshaarWord.parseContentControlTag(tagNoQ).qaseeda, "",
  "qaseeda defaults to empty when absent from the payload"
);
// parseContentControlTag tolerates non-ashaar / malformed tags
assert.equal(AshaarWord.parseContentControlTag(""), null, "empty tag => null");
assert.equal(AshaarWord.parseContentControlTag("not-ashaar"), null, "non-ashaar tag => null");
// Round-trips a real payload
const tagFull = AshaarWord.contentControlTag("poem", { layoutMode: "balanced", tableWidthPct: 75, qaseeda: "Q1" });
const parsedFull = AshaarWord.parseContentControlTag(tagFull);
assert.equal(parsedFull.k, "ashaar-poem", "parses kind");
assert.equal(parsedFull.tableWidthPct, 75, "parses tableWidthPct");
assert.equal(parsedFull.qaseeda, "Q1", "parses qaseeda");

// setTagQaseeda rewrites only the qaseeda field, leaving other payload intact
const tagBase = AshaarWord.contentControlTag("poem", { tableWidthPct: 60, qaseeda: "Old" });
const tagSet = AshaarWord.setTagQaseeda(tagBase, "New");
const parsedSet = AshaarWord.parseContentControlTag(tagSet);
assert.equal(parsedSet.qaseeda, "New", "setTagQaseeda updates the qaseeda name");
assert.equal(parsedSet.tableWidthPct, 60, "setTagQaseeda preserves other payload fields");
assert.equal(AshaarWord.setTagQaseeda("not-ashaar", "X"), "not-ashaar", "non-ashaar tag returned unchanged");
assert.equal(AshaarWord.parseContentControlTag(AshaarWord.setTagQaseeda(tagBase, "")).qaseeda, "", "clearing the name yields empty string");

// ── coalesceRuns ────────────────────────────────────────────────────────────
{
  // Two words, same style → one run; refs keep both source words in order.
  const w = [
    { text: "درد", name: "Amiri", size: 16, bold: false, italic: false },
    { text: "دل",  name: "Amiri", size: 16, bold: false, italic: false },
  ];
  const r = AshaarWord.coalesceRuns(w);
  assert.equal(r.length, 1);
  assert.equal(r[0].text, "درد دل");
  assert.equal(r[0].name, "Amiri");
  assert.deepEqual(r[0].refs, w);
}
{
  // Style change (bold) splits into two runs, order preserved, one ref each.
  const r = AshaarWord.coalesceRuns([
    { text: "درد", name: "Amiri", size: 16, bold: false, italic: false },
    { text: "دل",  name: "Amiri", size: 16, bold: true,  italic: false },
  ]);
  assert.equal(r.length, 2);
  assert.equal(r[0].text, "درد");
  assert.equal(r[1].text, "دل");
  assert.equal(r[1].bold, true);
  assert.equal(r[0].refs.length, 1);
  assert.equal(r[1].refs.length, 1);
}
{
  // Size change splits; the second run coalesces its two same-size words.
  const r = AshaarWord.coalesceRuns([
    { text: "الف", name: "Amiri", size: 24, bold: false, italic: false },
    { text: "ب",   name: "Amiri", size: 16, bold: false, italic: false },
    { text: "ج",   name: "Amiri", size: 16, bold: false, italic: false },
  ]);
  assert.deepEqual(r.map(x => x.text), ["الف", "ب ج"]);
  assert.equal(r[1].refs.length, 2);
  assert.deepEqual(AshaarWord.coalesceRuns([]), []);
}

// ── distributeMicroSpaces ───────────────────────────────────────────────────
{
  const HAIR = " ";
  // Two runs; gaps: run0 "a b" (1), run1 "c d" (1) → 2 gaps. n=2 → 1 each.
  const out = AshaarWord.distributeMicroSpaces(["a b", "c d"], 2, HAIR);
  assert.deepEqual(out, ["a " + HAIR + "b", "c " + HAIR + "d"]);
}
{
  const HAIR = " ";
  // n=3 over 2 gaps → first gap 2, second gap 1 (round-robin).
  const out = AshaarWord.distributeMicroSpaces(["a b", "c d"], 3, HAIR);
  assert.deepEqual(out, ["a " + HAIR + HAIR + "b", "c " + HAIR + "d"]);
}
{
  const HAIR = " ";
  // No gaps (single words) or n<=0 → unchanged.
  assert.deepEqual(AshaarWord.distributeMicroSpaces(["a", "b"], 5, HAIR), ["a", "b"]);
  assert.deepEqual(AshaarWord.distributeMicroSpaces(["a b"], 0, HAIR), ["a b"]);
}

// ── strengthToKashidaLevel — thirds of 1–10 ─────────────────────────────────
assert.equal(AshaarWord.strengthToKashidaLevel(1),  "lowKashida");
assert.equal(AshaarWord.strengthToKashidaLevel(3),  "lowKashida");
assert.equal(AshaarWord.strengthToKashidaLevel(4),  "mediumKashida");
assert.equal(AshaarWord.strengthToKashidaLevel(6),  "mediumKashida");
assert.equal(AshaarWord.strengthToKashidaLevel(7),  "highKashida");
assert.equal(AshaarWord.strengthToKashidaLevel(10), "highKashida");
assert.equal(AshaarWord.strengthToKashidaLevel(undefined), "mediumKashida");

// ── strengthToElongationShare: 1–10 → φ elongation share ─────────────────────
assert.strictEqual(AshaarWord.strengthToElongationShare(1), 0);
assert.strictEqual(AshaarWord.strengthToElongationShare(10), 1);
assert.ok(Math.abs(AshaarWord.strengthToElongationShare(5) - (4/9)) < 1e-9, "s5 → 4/9");
assert.strictEqual(AshaarWord.strengthToElongationShare(0), 0);   // clamps low
assert.strictEqual(AshaarWord.strengthToElongationShare(24), 1);  // clamps high
assert.strictEqual(AshaarWord.strengthToElongationShare(undefined), 0);

// ── strengthToMaxPositions: low-strength position cap K(s) ───────────────────
assert.strictEqual(AshaarWord.strengthToMaxPositions(1), 1);
assert.strictEqual(AshaarWord.strengthToMaxPositions(2), 2);
assert.strictEqual(AshaarWord.strengthToMaxPositions(3), 3);
assert.strictEqual(AshaarWord.strengthToMaxPositions(4), 0);   // unbounded
assert.strictEqual(AshaarWord.strengthToMaxPositions(10), 0);
assert.strictEqual(AshaarWord.strengthToMaxPositions(0), 1);   // clamps low → 1
assert.strictEqual(AshaarWord.strengthToMaxPositions(undefined), 1);

// ── containsArabic / wordFillJc ─────────────────────────────────────────────
assert.equal(AshaarWord.containsArabic("العلم"), true);
assert.equal(AshaarWord.containsArabic("hello"), false);
assert.equal(AshaarWord.containsArabic("۱۲۳"), true);   // Urdu digits are in-range
assert.equal(AshaarWord.wordFillJc("العلم نور", 2),  "lowKashida");
assert.equal(AshaarWord.wordFillJc("العلم نور", 7), "highKashida");
assert.equal(AshaarWord.wordFillJc("hello world", 7), "distribute");

// ── kashidaExpansionFraction — 0 at 1, ~0.15 at 10 ───────────────────────
assert.equal(AshaarWord.kashidaExpansionFraction(1), 0);
assert.equal(AshaarWord.kashidaExpansionFraction(10), 0.15);
assert.equal(AshaarWord.kashidaExpansionFraction(999), 0.15); // clamp

// ── misraParaXml: word-fill mode ────────────────────────────────────────────
{
  const opts = { justifyMode: "css", tatweelCount: 10 };
  const xml = AshaarWord.misraParaXml("العلم نور", "center", false, opts, 0);
  assert.ok(xml.indexOf('w:jc w:val="highKashida"') !== -1, "Arabic → highKashida jc");
  assert.ok(xml.indexOf("<w:br/>") !== -1, "Arabic word-fill appends a trailing break");
  assert.ok(xml.indexOf('w:sz w:val="4"') !== -1, "trailing break run is shrunk");
}
{
  const opts = { justifyMode: "css", tatweelCount: 10 };
  const xml = AshaarWord.misraParaXml("hello world", "center", false, opts, 0);
  assert.ok(xml.indexOf('w:jc w:val="distribute"') !== -1, "non-Arabic → distribute");
  assert.ok(xml.indexOf("<w:br/>") === -1, "distribute needs no trailing break");
}
{
  // Non-word-fill modes unchanged: jc still by position, no break.
  const xml = AshaarWord.misraParaXml("العلم", "right", false, { justifyMode: "kashida" }, 0);
  assert.ok(xml.indexOf('w:jc w:val="right"') !== -1, "kashida mode keeps positional jc");
  assert.ok(xml.indexOf("<w:br/>") === -1, "no break outside word-fill");
}

// ── misraParaXml: word-fill emits one trailing break, tiny break run ────────
{
  const opts = { justifyMode: "css", tatweelCount: 5 };
  const xml = AshaarWord.misraParaXml("العلم نور", "center", false, opts, 0);
  assert.ok(xml.indexOf('w:jc w:val="mediumKashida"') !== -1, "tatweelCount 5 -> mediumKashida");
  assert.equal((xml.match(/<w:br\/>/g) || []).length, 1, "exactly one trailing break");
  assert.ok(xml.indexOf('<w:rPr><w:sz w:val="4"/><w:szCs w:val="4"/></w:rPr><w:br/>') !== -1,
    "the trailing break run keeps its own tiny size");
  // document-default fontMode injects no rFonts (font comes from the render, not per-cell).
  assert.ok(xml.indexOf("w:rFonts") === -1, "no rFonts for document-default fontMode");
}

// ── misraParaXml: word-fill shrinks the paragraph mark (empty-line fix) ──────
{
  const xml = AshaarWord.misraParaXml("العلم نور", "center", false, { justifyMode: "css", tatweelCount: 5 }, 0);
  // paragraph-mark rPr with sz 4 present inside pPr, before the first text run
  const pPrEnd = xml.indexOf("</w:pPr>");
  assert.ok(pPrEnd !== -1, "has pPr");
  assert.ok(xml.slice(0, pPrEnd).indexOf('w:sz w:val="4"') !== -1, "paragraph mark shrunk to sz 4 inside pPr");
}
{
  // non-word-fill unchanged: no paragraph-mark shrink
  const xml = AshaarWord.misraParaXml("العلم", "right", false, { justifyMode: "kashida" }, 0);
  const pPrEnd = xml.indexOf("</w:pPr>");
  assert.ok(xml.slice(0, pPrEnd).indexOf('w:sz w:val="4"') === -1, "no paragraph-mark shrink outside word-fill");
}

// ── misraParaXml: opts.fontSizePt preserves size on re-render ────────────────
// A size-preserving re-render (gap/width apply that rebuilds) must carry the
// cell's font size, or the rebuilt table reverts to Word's 12pt default.
{
  const xml = AshaarWord.misraParaXml("العلم", "right", false, { fontSizePt: 16 }, 0);
  assert.ok(xml.indexOf('w:sz w:val="32"') !== -1, "fontSizePt 16 → body run w:sz 32 (half-points)");
  assert.ok(xml.indexOf('w:szCs w:val="32"') !== -1, "fontSizePt 16 → body run w:szCs 32");
}
{
  // no fontSizePt (e.g. plain insert) → no body-run size, Word default preserved
  const xml = AshaarWord.misraParaXml("العلم", "right", false, {}, 0);
  assert.ok(xml.indexOf('w:sz w:val=') === -1, "no fontSizePt → no w:sz (insert default unchanged)");
}

// ── misraParaXml: opts.fontCsName overrides the cs font (Re-render font pin) ──
// Re-render preserves an arbitrary existing font (e.g. Fatemi Maqala, not in the
// registry) by pinning the exact cs name, overriding fontMode's registry lookup.
{
  const xml = AshaarWord.misraParaXml("العلم", "right", false, { fontMode: "document", fontCsName: "Fatemi Maqala" }, 0);
  assert.ok(xml.indexOf('w:cs="Fatemi Maqala"') !== -1, "fontCsName pins the cs font even under Document default");
}
{
  // fontCsName wins over the fontMode registry font
  const xml = AshaarWord.misraParaXml("العلم", "right", false, { fontMode: "mehr", fontCsName: "Fatemi Maqala" }, 0);
  assert.ok(xml.indexOf('w:cs="Fatemi Maqala"') !== -1, "fontCsName overrides fontMode's wordName");
  assert.ok(xml.indexOf('w:cs="Mehr Nastaliq Web"') === -1, "fontMode font not emitted when fontCsName set");
}

// ── soloRow (OOXML): solo/refrain lines are misra-width, not full-grid ──────
// A marsiya's solo line and paired refrain lines used to stretch a single cell
// across the FULL grid (gridSpan = si.GRID). That looks wrong: they should be
// the SAME width as an ordinary misra (BASE_CPM = 6 columns), centered, with
// empty pad cells flanking them so the row's spans still sum to si.GRID.
{
  // marsiyaSource: N=3, gapCols=1 (default) → GRID = 3*6 + 2*1 = 20
  const ooxml = AshaarWord.renderForWordOoxml(marsiyaSource,
    { justifyMode: "none", gapWidth: 1 }, Ashaar, 9360);
  const rows = ooxml.match(/<w:tr>[\s\S]*?<\/w:tr>/g);
  const soloRowXml = rows.filter((r) => /هو گئے شہ پر فدا/.test(r))[0];
  assert.ok(soloRowXml, "solo row found");
  const spans = [...soloRowXml.matchAll(/<w:gridSpan w:val="(\d+)"\/>/g)].map((m) => Number(m[1]));
  // Must NOT be a single full-grid cell (gridSpan === GRID === 20).
  assert.ok(!(spans.length === 1 && spans[0] === 20),
    "solo row must not be a single full-grid (20) cell");
  // The text-bearing cell must span BASE_CPM = 6 columns, matching ordinary misra width.
  assert.match(soloRowXml, /<w:gridSpan w:val="6"\/>/,
    "solo misra cell must span BASE_CPM=6 columns, same as an ordinary misra");
  // Spans must still sum to GRID=20 (pad + solo + pad), so the fixed-layout table stays valid.
  assert.equal(spans.reduce((a, b) => a + b, 0), 20, "row spans must sum to GRID=20");
  // The refrain pair row (paired \"هائے كربلاء والو\" maqta) is a misraRow, not a
  // solo — untouched by this fix; still spans the full 2-misra split.
}

// Registry-driven fonts: Mehr/Gulzar reach the OOXML cs name and preview stack.
{
  const AshaarFonts = require("../src/taskpane/fonts");
  assert.strictEqual(AshaarFonts.wordNameOf("mehr"), "Mehr Nastaliq Web");
  // fontFamilyStyle now delegates to the registry
  const mehrCss = AshaarWord.fontFamilyStyle({ fontMode: "mehr" });
  assert.ok(/Mehr Nastaliq Web/.test(mehrCss), "mehr css from registry");
  const gulzarCss = AshaarWord.fontFamilyStyle({ fontMode: "gulzar" });
  assert.ok(/Gulzar/.test(gulzarCss), "gulzar css from registry");
  // legacy "nastaliq" alias still resolves to a Nastaliq face (Noto)
  const notoCss = AshaarWord.fontFamilyStyle({ fontMode: "nastaliq" });
  assert.ok(/Noto Nastaliq Urdu/.test(notoCss), "nastaliq alias preserved");
  console.log("word-html registry-font tests passed");
}

// ── runsToMisraXml (Jameel font-swap kashida — Task 8) ─────────────────────
{
  const xml = AshaarWord.runsToMisraXml(
    [{text:"كہہ", swap:true},{text:" ", swap:false},{text:"تھے", swap:false}],
    "right", { fontMode: "jameel" });
  assert.ok(xml.indexOf('w:cs="Jameel Noori Nastaleeq Kasheeda"') !== -1, "wider face on swapped fasl");
  assert.ok(xml.indexOf('w:cs="Jameel Noori Nastaleeq"') !== -1, "base face on non-swapped fasl");
  console.log("word-html font-swap emitter tests passed");
}

// ── mehrElongate (form-aware Mehr tatweel) ──────────────────────────────────
{
  const ISO = {}; "بپتٹثسشفکگ".split("").forEach((c) => { ISO[c] = true; }); // isolatedInto
  const FIN = {}; "بپتٹثفکگ".split("").forEach((c) => { FIN[c] = true; });   // finalInto (no س ش)
  const T = "ـ";
  // final form, allowed: "شب" — beh joined from sheen (dual) → final; ب ∈ FIN → tatweel after beh
  assert.equal(AshaarWord.mehrElongate("شب", ISO, FIN), "شب" + T);
  // seen in FINAL form NOT allowed: "نس" — noon joins → seen final; س ∉ FIN → unchanged
  assert.equal(AshaarWord.mehrElongate("نس", ISO, FIN), "نس");
  // seen in ISOLATED form allowed: "آس" — alef is right-only joiner → seen isolated; س ∈ ISO → tatweel
  assert.equal(AshaarWord.mehrElongate("آس", ISO, FIN), "آس" + T);
  // lone isolated allowed letter
  assert.equal(AshaarWord.mehrElongate("ب", ISO, FIN), "ب" + T);
  // trailing diacritic skipped; tatweel inserted after the base letter, before the mark
  assert.equal(AshaarWord.mehrElongate("کتب" + "ٌ", ISO, FIN), "کتب" + T + "ٌ");
  // not an allowed letter (lam) → unchanged
  assert.equal(AshaarWord.mehrElongate("دل", ISO, FIN), "دل");
}

// ── misraDistributeXml: Cell-fit distribute paragraph ────────────────────────
{
  const xml = AshaarWord.misraDistributeXml(
    [{ text: "دل", csName: "Fatemi Maqala" }, { text: " ", csName: "Fatemi Maqala" }, { text: "ناداں", csName: "Fatemi Maqala" }],
    16
  );
  assert.match(xml, /<w:jc w:val="distribute"\/>/, "distribute jc");
  assert.match(xml, /<w:bidi\/>/, "rtl paragraph");
  assert.match(xml, /<w:rtl\/>/, "rtl runs");
  assert.match(xml, /<w:rFonts w:cs="Fatemi Maqala"\/>/, "per-run cs font");
  assert.match(xml, /<w:sz w:val="32"\/>/, "16pt -> 32 half-points");
  assert.ok(xml.indexOf("دل") !== -1 && xml.indexOf("ناداں") !== -1, "carries text");
  // No injected hair/thin spaces (distribute is the residual, not micro-spaces).
  assert.ok(xml.indexOf(" ") === -1 && xml.indexOf(" ") === -1, "no micro-spaces injected");
}
{
  // Per-run size override wins over the fallback.
  const xml = AshaarWord.misraDistributeXml([{ text: "x", csName: "A", sizePt: 20 }], 16);
  assert.match(xml, /<w:sz w:val="40"\/>/, "per-run 20pt -> 40 half-points");
}

// ── cell patterns: mirror the OOXML generator's content/gap cell order ───────
{
  const Ashaar2 = require("../src/vendor/ashaar");
  // A simple couplet (sadr \ ajuz) → one table, one row [c,g,c].
  const pats = AshaarWord.poemCellPatterns("دل ناداں \\ آخر اس درد", { layoutMode: "balanced" }, Ashaar2);
  assert.equal(pats.length, 1, "one stanza → one pattern");
  assert.deepStrictEqual(pats[0], [["c", "g", "c"]], "couplet row = content,gap,content");
}
{
  const Ashaar2 = require("../src/vendor/ashaar");
  // Solo single misra (|) → one table, one row [g,c,g].
  const pats = AshaarWord.poemCellPatterns("تنہا مصرعہ |", { layoutMode: "balanced" }, Ashaar2);
  assert.deepStrictEqual(pats[0], [["g", "c", "g"]], "solo row = gap,content,gap");
}
{
  const Ashaar2 = require("../src/vendor/ashaar");
  // Stacked couplet → two solo rows.
  const pats = AshaarWord.poemCellPatterns("دل ناداں \\ آخر اس درد", { layoutMode: "stacked" }, Ashaar2);
  assert.deepStrictEqual(pats[0], [["g", "c", "g"], ["g", "c", "g"]], "stacked = two solo rows");
}

// ── cross-check: pattern shape/kind == the generator's actual <w:tc> cells ───
{
  const Ashaar2 = require("../src/vendor/ashaar");
  function tablesOf(xml) { return xml.match(/<w:tbl>[\s\S]*?<\/w:tbl>/g) || []; }
  function rowsOf(tblXml) { return tblXml.match(/<w:tr>[\s\S]*?<\/w:tr>/g) || []; }
  function cellsOf(trXml) { return trXml.match(/<w:tc>[\s\S]*?<\/w:tc>/g) || []; }
  function kindOf(tcXml) { return /<w:r[ >]/.test(tcXml) ? "c" : "g"; }

  const cases = [
    { src: "دل ناداں \\ آخر اس درد", opts: { layoutMode: "balanced" } },
    { src: "تنہا مصرعہ |", opts: { layoutMode: "balanced" } },
    { src: "دل ناداں \\ آخر اس درد", opts: { layoutMode: "stacked" } },
    { src: "الف \\ ب\n\nج \\ د", opts: { layoutMode: "balanced" } }, // two stanzas
  ];
  cases.forEach(function (c, ci) {
    const body = AshaarWord.renderForWordOoxml(c.src, c.opts, Ashaar2, 9360);
    const pats = AshaarWord.poemCellPatterns(c.src, c.opts, Ashaar2);
    const tbls = tablesOf(body);
    assert.equal(tbls.length, pats.length, "case " + ci + ": table count == pattern count");
    tbls.forEach(function (tbl, ti) {
      const rows = rowsOf(tbl);
      assert.equal(rows.length, pats[ti].length, "case " + ci + " tbl " + ti + ": row count");
      rows.forEach(function (tr, ri) {
        const kinds = cellsOf(tr).map(kindOf);
        assert.deepStrictEqual(kinds, pats[ti][ri], "case " + ci + " tbl " + ti + " row " + ri + ": cell kinds");
      });
    });
  });
}

console.log("word-html tests passed");

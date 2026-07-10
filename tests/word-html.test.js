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

// 2-misra poem: N=2, gapCols=1 → GRID=7 (2*3+1*1=7)
const ooxml2misra = AshaarWord.renderForWordOoxml(
  "دل ناداں تجھے ہوا کیا ہے \\ آخر اس درد کی دوا کیا ہے",
  { justifyMode: "none", gapWidth: 1 }, Ashaar, 9360
);
assert.match(ooxml2misra, /<w:tbl>/);
assert.match(ooxml2misra, /<w:bidiVisual\/>/);
assert.match(ooxml2misra, /<w:gridSpan w:val="/);
// GRID=7 columns: 7 gridCol elements
assert.equal((ooxml2misra.match(/<w:gridCol /g) || []).length, 7);
// Two misra cells + one gap cell = 3 tc elements in the row (gridSpan sums to 7)
const tc2 = (ooxml2misra.match(/<w:tc>/g) || []).length;
assert.equal(tc2, 3, "N=2: 2 misra cells + 1 gap cell");
assert.match(ooxml2misra, /دل ناداں تجھے ہوا کیا ہے/);
assert.match(ooxml2misra, /آخر اس درد کی دوا کیا ہے/);

// 3-misra marsiya stanza: one table with GRID=11 for all rows
// Row 1 (3-misra): spans=[3,1,3,1,3]=11; Row 2 (solo): padded; Row 3 (2-misra): spans=[5,1,5]=11
const ooxml3misra = AshaarWord.renderForWordOoxml(marsiyaSource,
  { justifyMode: "none", gapWidth: 1 }, Ashaar, 9360
);
assert.equal((ooxml3misra.match(/<w:gridCol /g) || []).length, 11); // single GRID=11 table
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
  // Two words, same style → one run.
  const r = AshaarWord.coalesceRuns([
    { text: "درد", name: "Amiri", size: 16, bold: false, italic: false },
    { text: "دل",  name: "Amiri", size: 16, bold: false, italic: false },
  ]);
  assert.deepEqual(r, [{ text: "درد دل", name: "Amiri", size: 16, bold: false, italic: false }]);
}
{
  // Style change (bold) splits into two runs, order preserved.
  const r = AshaarWord.coalesceRuns([
    { text: "درد", name: "Amiri", size: 16, bold: false, italic: false },
    { text: "دل",  name: "Amiri", size: 16, bold: true,  italic: false },
  ]);
  assert.equal(r.length, 2);
  assert.equal(r[0].text, "درد");
  assert.equal(r[1].text, "دل");
  assert.equal(r[1].bold, true);
}
{
  // Size change splits; empty input → [].
  const r = AshaarWord.coalesceRuns([
    { text: "الف", name: "Amiri", size: 24, bold: false, italic: false },
    { text: "ب",   name: "Amiri", size: 16, bold: false, italic: false },
    { text: "ج",   name: "Amiri", size: 16, bold: false, italic: false },
  ]);
  assert.deepEqual(r.map(x => x.text), ["الف", "ب ج"]);
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

console.log("word-html tests passed");

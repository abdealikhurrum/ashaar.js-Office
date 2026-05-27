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
assert.match(html, /<td style="[^"]*text-align:left/);
assert.match(html, /<td style="[^"]*text-align:right/);
assert.match(html, /ـ/);

const firstRow = html.match(/<tr>(.*?)<\/tr>/)[1];
const cells = firstRow.match(/<td style="[^"]*"/g);
assert.match(cells[0], /text-align:right/);
assert.match(cells[1], /text-align:center/);
assert.match(cells[2], /text-align:left/);

const stacked = AshaarWord.renderForWord(source, {
  justifyMode: "none",
  layoutMode: "stacked"
}, Ashaar);

assert.match(stacked, /colspan="3"/);
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

  // Colgroup has exactly misraCount columns
  var cols = (html.match(/<col style=/g) || []).length;
  assert.equal(cols, misraCount, "Expected " + misraCount + " columns for " + misraCount + "-misra stanza");

  // Multi-misra row has exactly misraCount cells
  var rowCells = html.match(/<tr>(<td[^>]*>[\s\S]*?<\/td>){1,}<\/tr>/g) || [];
  var firstRowCellCount = (rowCells[0].match(/<td/g) || []).length;
  assert.equal(firstRowCellCount, misraCount, "Expected " + misraCount + " cells in full row");

  // Solo misra row uses colspan=misraCount
  assert.match(html, new RegExp('colspan="' + misraCount + '"'), "Solo row should span all " + misraCount + " columns");
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
assert.equal((marsiyaWordHtml.match(/<col style=/g) || []).length, 3);
assert.match(marsiyaWordHtml, /colspan="3"/);

// Test: layoutTablesForPoem — native table path for multi-misra poems
const poemTables3 = AshaarWord.layoutTablesForPoem(marsiyaSource, { justifyMode: "none" }, Ashaar);
assert.ok(poemTables3, "Should return tables for multi-misra poem");
assert.equal(poemTables3.length, 1);
assert.equal(poemTables3[0].columnCount, 3);
assert.equal(poemTables3[0].rows.length, 3);
// Triple row: [ajuz-side, middle, sadr-side] — misras reversed into LTR column order
assert.equal(poemTables3[0].rows[0][0].text, "صدق كے ارباب تھے"); // col 0 = misras[2]
assert.equal(poemTables3[0].rows[0][1].text, "خلق ميں الباب تھے"); // col 1 = misras[1]
assert.equal(poemTables3[0].rows[0][2].text, "شاه كے اصحاب تھے"); // col 2 = misras[0]
assert.equal(poemTables3[0].rows[0][0].align, "right");
assert.equal(poemTables3[0].rows[0][2].align, "left");
// Solo row: text in middle cell, others empty
assert.equal(poemTables3[0].rows[1][1].text, "هو گئے شہ پر فدا");
assert.equal(poemTables3[0].rows[1][0].text, "");
// Maqta row: ajuz in col 0, sadr in col 2
assert.equal(poemTables3[0].rows[2][0].text, "هائے كربلاء والو");
assert.equal(poemTables3[0].rows[2][2].text, "هائے كربلاء والو");

// Regular 2-misra poem returns null (falls through to renderForWord)
const regularPoem = "دل ناداں تجھے ہوا کیا ہے \\ آخر اس درد کی دوا کیا ہے";
const poemTablesRegular = AshaarWord.layoutTablesForPoem(regularPoem, {}, Ashaar);
assert.equal(poemTablesRegular, null, "Regular couplet poem should return null");

// 5-misra poem
const fiveMisraSource = "م1 \\ م2 \\ م3 \\ م4 \\ م5\nنعرہ \\";
const poemTables5 = AshaarWord.layoutTablesForPoem(fiveMisraSource, {}, Ashaar);
assert.ok(poemTables5);
assert.equal(poemTables5[0].columnCount, 5);
assert.equal(poemTables5[0].rows[0].length, 5);
assert.equal(poemTables5[0].rows[0][4].text, "م1"); // misras[0] → rightmost col

console.log("word-html tests passed");

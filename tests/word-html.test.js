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
assert.match(cells[0], /text-align:left/);
assert.match(cells[1], /text-align:center/);
assert.match(cells[2], /text-align:right/);

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
assert.match(templateCells[0], /text-align:left/);
assert.match(templateCells[1], /text-align:center/);
assert.match(templateCells[2], /text-align:right/);
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

console.log("word-html tests passed");

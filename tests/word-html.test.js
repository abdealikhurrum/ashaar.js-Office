const assert = require("assert");
const Ashaar = require("../src/vendor/ashaar");
const AshaarWord = require("../src/taskpane/word-html");

const source = "دل ناداں تجھے ہوا کیا ہے \\ آخر اس درد کی دوا کیا ہے";

const html = AshaarWord.renderForWord(source, {
  justifyMode: "kashida",
  tatweelCount: 4,
  gapWidth: 4,
  layoutMode: "columns"
}, Ashaar);

assert.match(html, /<table dir="rtl"/);
assert.match(html, /<td style="[^"]*text-align:left/);
assert.match(html, /<td style="[^"]*text-align:right/);
assert.match(html, /ـ/);

const stacked = AshaarWord.renderForWord(source, {
  justifyMode: "none",
  layoutMode: "stacked"
}, Ashaar);

assert.match(stacked, /colspan="3"/);
assert.match(stacked, /<br>/);

const plain = AshaarWord.justifyPlainTextBlock("سلام دنیا", {
  justifyMode: "kashida",
  tatweelCount: 3
});

assert.match(plain, /ـ/);

console.log("word-html tests passed");

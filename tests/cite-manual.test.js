"use strict";
var assert = require("assert");
var M = require("../src/taskpane/cite-manual");

// parseNames
assert.deepStrictEqual(M.parseNames("al-Nuʿmān, al-Qāḍī"),
  [{ family: "al-Nuʿmān", given: "al-Qāḍī" }]);
assert.deepStrictEqual(M.parseNames("UNESCO"), [{ literal: "UNESCO" }]);
assert.deepStrictEqual(M.parseNames("Smith,"), [{ family: "Smith" }]); // empty given omitted
assert.deepStrictEqual(M.parseNames("Halm, Heinz\n\nDaftary, Farhad"),
  [{ family: "Halm", given: "Heinz" }, { family: "Daftary", given: "Farhad" }]); // blank line skipped
assert.deepStrictEqual(M.parseNames(""), []);
assert.deepStrictEqual(M.parseNames("   "), []);

// parseDateParts
assert.deepStrictEqual(M.parseDateParts("1951"), [[1951]]);
assert.deepStrictEqual(M.parseDateParts("2026-07-18"), [[2026, 7, 18]]);
assert.deepStrictEqual(M.parseDateParts("2026-07"), [[2026, 7]]);
assert.strictEqual(M.parseDateParts(""), null);
assert.strictEqual(M.parseDateParts("n.d."), null);

// buildManualItem — book
var book = M.buildManualItem({
  id: "manual-1", type: "book", title: "The Fatimid Empire",
  authors: "Daftary, Farhad", year: "2018",
  publisher: "Edinburgh University Press", place: "Edinburgh"
});
assert.strictEqual(book.id, "manual-1");
assert.strictEqual(book.type, "book");
assert.strictEqual(book.title, "The Fatimid Empire");
assert.deepStrictEqual(book.author, [{ family: "Daftary", given: "Farhad" }]);
assert.deepStrictEqual(book.issued, { "date-parts": [[2018]] });
assert.strictEqual(book.publisher, "Edinburgh University Press");
assert.strictEqual(book["publisher-place"], "Edinburgh");
assert.ok(!("container-title" in book)); // empty fields omitted

// buildManualItem — journal article
var art = M.buildManualItem({
  id: "manual-2", type: "article", title: "Isma'ili History",
  authors: "Halm, Heinz", year: "2001",
  containerTitle: "Journal of Islamic Studies", volume: "12", issue: "2", pages: "145-170"
});
assert.strictEqual(art.type, "article-journal");
assert.strictEqual(art["container-title"], "Journal of Islamic Studies");
assert.strictEqual(art.volume, "12");
assert.strictEqual(art.issue, "2");
assert.strictEqual(art.page, "145-170");

// buildManualItem — chapter with editor
var chap = M.buildManualItem({
  id: "manual-3", type: "chapter", title: "A Chapter",
  authors: "Author, An", editors: "Editor, Ed", containerTitle: "The Book",
  publisher: "Pub", pages: "10-20"
});
assert.strictEqual(chap.type, "chapter");
assert.deepStrictEqual(chap.editor, [{ family: "Editor", given: "Ed" }]);
assert.strictEqual(chap["container-title"], "The Book");

// buildManualItem — webpage
var web = M.buildManualItem({
  id: "manual-4", type: "webpage", title: "A Page",
  url: "https://example.org", accessed: "2026-07-18"
});
assert.strictEqual(web.type, "webpage");
assert.strictEqual(web.URL, "https://example.org");
assert.deepStrictEqual(web.accessed, { "date-parts": [[2026, 7, 18]] });

console.log("cite-manual.test.js passed");

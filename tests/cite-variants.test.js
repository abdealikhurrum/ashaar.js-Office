"use strict";
const assert = require("assert");
const CV = require("../src/taskpane/cite-variants");

// --- parseCne ---

// no cne-* lines -> null
assert.strictEqual(CV.parseCne("plain note text"), null, "no cne-* => null");
assert.strictEqual(CV.parseCne(""), null, "empty => null");

// simple field: title romanized
const r1 = CV.parseCne("cne-title-romanized: Uyun al-Akhbar Vol. 4");
assert.deepStrictEqual(r1.fields.title, { "ar-Latn": "Uyun al-Akhbar Vol. 4" }, "title romanized -> ar-Latn");

// hyphenated field name: container-title
const r2 = CV.parseCne("cne-container-title-romanized: al-Majalla");
assert.deepStrictEqual(r2.fields["container-title"], { "ar-Latn": "al-Majalla" }, "container-title parsed");

// creator: author 0 last romanized (literal-style, only last present)
const r3 = CV.parseCne("cne-author-0-last-romanized: al-Dai al-Ajal Syedna Idris Imaduddin RA");
assert.deepStrictEqual(
  r3.creators.author["0"],
  { "ar-Latn": { family: "al-Dai al-Ajal Syedna Idris Imaduddin RA" } },
  "author 0 last -> family"
);

// creator with both parts
const r4 = CV.parseCne("cne-author-0-last-romanized: al-Nuʿmān\ncne-author-0-first-romanized: al-Qāḍī");
assert.deepStrictEqual(
  r4.creators.author["0"]["ar-Latn"],
  { family: "al-Nuʿmān", given: "al-Qāḍī" },
  "last+first -> family+given"
);

// translated variant -> en tag
const r5 = CV.parseCne("cne-title-translated: The Sources of History");
assert.deepStrictEqual(r5.fields.title, { en: "The Sources of History" }, "translated -> en");

// unknown variant suffix -> line ignored (but still null if it's the only line)
assert.strictEqual(CV.parseCne("cne-title-banana: x"), null, "unknown variant => ignored");

// bidi control chars stripped
const r6 = CV.parseCne("cne-title-romanized: ‫Uyun‬");
assert.strictEqual(r6.fields.title["ar-Latn"], "Uyun", "bidi controls stripped");

console.log("cite-variants parseCne test passed");

// --- applyVariantsToItem ---
const baseItem = {
  id: "x", type: "book", language: "ar",
  title: "عيون الأخبار ج/4",
  author: [{ literal: "الداعي الأجل سيدنا إدريس عماد الدينؓ" }],
  note: "cne-title-romanized: Uyun al-Akhbar Vol. 4\ncne-author-0-last-romanized: al-Dai al-Ajal Syedna Idris Imaduddin RA"
};
const enriched = CV.applyVariantsToItem(baseItem);
assert.notStrictEqual(enriched, baseItem, "returns a new object");
assert.strictEqual(baseItem.multi, undefined, "input not mutated");
assert.deepStrictEqual(enriched.multi._keys.title, { "ar-Latn": "Uyun al-Akhbar Vol. 4" }, "title variant baked");
assert.deepStrictEqual(
  enriched.author[0].multi._key["ar-Latn"],
  { family: "al-Dai al-Ajal Syedna Idris Imaduddin RA" },
  "author variant baked"
);
// real fields preserved
assert.strictEqual(enriched.title, "عيون الأخبار ج/4", "real title preserved");

// no cne-* -> unchanged (same reference)
const plain = { id: "p", type: "book", title: "T", note: "just a note" };
assert.strictEqual(CV.applyVariantsToItem(plain), plain, "no variants => passthrough");

// creator index mismatch -> skipped, no throw
const mismatch = { id: "m", type: "book", author: [], note: "cne-author-3-last-romanized: Z" };
const em = CV.applyVariantsToItem(mismatch);
assert.ok(!(em.author && em.author[3]), "missing creator index skipped");

// enrichItemMap
const map = { x: baseItem, p: plain };
const em2 = CV.enrichItemMap(map);
assert.deepStrictEqual(em2.x.multi._keys.title, { "ar-Latn": "Uyun al-Akhbar Vol. 4" });
assert.strictEqual(em2.p, plain, "passthrough item shared by reference");
console.log("cite-variants applyVariantsToItem test passed");

// --- variantToLangPrefs ---
assert.strictEqual(CV.variantToLangPrefs("orig"), null, "orig => no override");

const lpT = CV.variantToLangPrefs("translit");
assert.deepStrictEqual(lpT.persons, ["translit"], "translit persons");
assert.deepStrictEqual(lpT.titles, ["translit"], "translit titles");
assert.deepStrictEqual(lpT.translit, ["ar-Latn"], "translit tag registered");
assert.deepStrictEqual(lpT.translat, ["en"], "translat tag registered");

const lpB = CV.variantToLangPrefs("both");
assert.deepStrictEqual(lpB.persons, ["orig", "translit"], "both persons");
assert.deepStrictEqual(lpB.titles, ["orig", "translit"], "both titles");

// unknown => treated as orig
assert.strictEqual(CV.variantToLangPrefs("nonsense"), null, "unknown => null");
console.log("cite-variants variantToLangPrefs test passed");

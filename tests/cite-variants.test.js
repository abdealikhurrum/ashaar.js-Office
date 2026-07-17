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

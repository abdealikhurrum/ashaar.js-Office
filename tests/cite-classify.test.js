"use strict";
var assert = require("assert");
var C = require("../src/taskpane/cite-classify");

// bucketForTags — defaults
assert.deepStrictEqual(C.bucketForTags([]), { corpus: "other", cls: "primary", key: "primary.other" });
assert.strictEqual(C.bucketForTags(["corpus:fatemi"]).key, "primary.fatemi");
assert.strictEqual(C.bucketForTags(["class:secondary"]).key, "secondary.other");
assert.strictEqual(C.bucketForTags(["corpus:fatemi", "class:secondary"]).key, "secondary.fatemi");
// class:primary is equivalent to default
assert.strictEqual(C.bucketForTags(["class:primary"]).key, "primary.other");
// malformed axis value falls back to default; unrelated tags ignored
assert.strictEqual(C.bucketForTags(["class:tertiary", "keyword:foo"]).key, "primary.other");
assert.strictEqual(C.bucketForTags(["corpus:other-thing"]).key, "primary.other");

// orderedBuckets — only non-empty, fixed order, order preserved within bucket
var tags = {
  a: ["corpus:fatemi"],                    // primary.fatemi
  b: ["class:secondary"],                  // secondary.other
  c: [],                                   // primary.other
  d: ["corpus:fatemi", "class:secondary"], // secondary.fatemi
  e: ["corpus:fatemi"]                     // primary.fatemi
};
var ob = C.orderedBuckets(["a", "b", "c", "d", "e"], tags);
assert.deepStrictEqual(ob.map(function (x) { return x.key; }),
  ["primary.fatemi", "primary.other", "secondary.fatemi", "secondary.other"]);
assert.deepStrictEqual(ob[0].citekeys, ["a", "e"]); // input order preserved

// headingFor — en + ar
assert.strictEqual(C.headingFor("primary.fatemi", "en-US"), "Primary Sources — Fatemi");
assert.strictEqual(C.headingFor("secondary.other", "en-US"), "Secondary Sources — Other");
assert.strictEqual(C.headingFor("primary.fatemi", "ar"), "المصادر الأساسية — الفاطمية");

// planBibliographySections — not sectioned => single flat section, no heading
var flat = C.planBibliographySections(["a", "b"], tags, { sectioned: false, lang: "en-US" });
assert.strictEqual(flat.length, 1);
assert.strictEqual(flat[0].heading, null);
assert.deepStrictEqual(flat[0].citekeys, ["a", "b"]);

// planBibliographySections — sectioned, >=2 buckets => one section per bucket, headings set
var sec = C.planBibliographySections(["a", "b", "c", "d"], tags, { sectioned: true, lang: "en-US" });
assert.strictEqual(sec.length, 4);
assert.strictEqual(sec[0].heading, "Primary Sources — Fatemi");
assert.deepStrictEqual(sec[0].citekeys, ["a"]);

// planBibliographySections — sectioned but only ONE non-empty bucket => collapse, no heading
var one = C.planBibliographySections(["a", "e"], { a: ["corpus:fatemi"], e: ["corpus:fatemi"] },
  { sectioned: true, lang: "en-US" });
assert.strictEqual(one.length, 1);
assert.strictEqual(one[0].heading, null);
assert.deepStrictEqual(one[0].citekeys, ["a", "e"]);

console.log("cite-classify.test.js passed");

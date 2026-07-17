"use strict";
const assert = require("assert");
const CiteStore = require("../src/taskpane/cite-store.js");

// fake settings bag (mirrors Office.context.document.settings)
function fakeSettings(saveStatus) {
  var bag = {};
  return {
    _bag: bag,
    set: function (k, v) { bag[k] = v; },
    get: function (k) { return Object.prototype.hasOwnProperty.call(bag, k) ? bag[k] : null; },
    remove: function (k) { delete bag[k]; },
    saveAsync: function (cb) { cb({ status: saveStatus || "succeeded" }); }
  };
}

// --- serialize/parse round-trip ---
var items = { A: { id: "A", title: "Alpha", type: "book" }, B: { id: "B", title: "بيتا", type: "document" } };
var str = CiteStore.serializeRefs(items);
assert.deepStrictEqual(CiteStore.parseRefs(str), items, "round-trips the items map (incl. Arabic)");
assert.deepStrictEqual(CiteStore.parseRefs(null), {}, "null → {}");
assert.deepStrictEqual(CiteStore.parseRefs(""), {}, "empty → {}");
assert.deepStrictEqual(CiteStore.parseRefs("{not json"), {}, "malformed → {}");
assert.deepStrictEqual(CiteStore.parseRefs(JSON.stringify({ v: 2, items: items })), {}, "wrong version → {}");
assert.deepStrictEqual(CiteStore.parseRefs(JSON.stringify({ v: 1 })), {}, "missing items → {}");
console.log("serialize/parse test passed");

// --- saveRefs writes REFS_KEY as a string + resolves on success ---
(async () => {
  var s = fakeSettings("succeeded");
  await CiteStore.saveRefs(items, s);
  assert.strictEqual(typeof s._bag[CiteStore.REFS_KEY], "string", "saved value is a JSON string");
  assert.deepStrictEqual(CiteStore.parseRefs(s._bag[CiteStore.REFS_KEY]), items);
  console.log("saveRefs test passed");
})();

// --- saveRefs rejects on a Failed status ---
(async () => {
  var s = fakeSettings("failed");
  var rejected = false;
  try { await CiteStore.saveRefs(items, s); } catch (e) { rejected = true; }
  assert.ok(rejected, "saveRefs rejects when saveAsync reports failed");
  console.log("saveRefs (failed) test passed");
})();

// --- loadRefs returns the stored map, {} when absent ---
(async () => {
  var s = fakeSettings("succeeded");
  await CiteStore.saveRefs(items, s);
  assert.deepStrictEqual(await CiteStore.loadRefs(s), items, "loadRefs returns the saved map");
  assert.deepStrictEqual(await CiteStore.loadRefs(fakeSettings()), {}, "absent key → {}");
  console.log("loadRefs test passed");
})();

// --- no settings (browser): saveRefs no-op resolves, loadRefs → {} ---
(async () => {
  await CiteStore.saveRefs(items, null); // must resolve, not throw
  assert.deepStrictEqual(await CiteStore.loadRefs(null), {}, "no settings → {}");
  console.log("no-settings test passed");
})();

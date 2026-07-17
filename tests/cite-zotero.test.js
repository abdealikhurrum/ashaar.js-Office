"use strict";
const assert = require("assert");
const CiteZotero = require("../src/taskpane/cite-zotero.js");

// --- buildExportRequest: pure JSON-RPC payload builder ---
assert.deepStrictEqual(
  CiteZotero.buildExportRequest(["YaumulMabasUyun"]),
  { jsonrpc: "2.0", method: "item.export", params: [["YaumulMabasUyun"], "Better CSL JSON"], id: 1 },
  "buildExportRequest matches the live-verified RPC payload shape"
);
console.log("buildExportRequest test passed");

// --- parseExportResult: success (result is a JSON string to be parsed) ---
const liveShaped = {
  jsonrpc: "2.0",
  result: "[\n  {\"id\":\"YaumulMabasUyun\",\"citation-key\":\"YaumulMabasUyun\",\"title\":\"Yaumul Mabas - Uyun 1\",\"type\":\"document\"}\n]\n",
  id: 1
};
const parsed = CiteZotero.parseExportResult(liveShaped, ["YaumulMabasUyun"]);
assert.deepStrictEqual(parsed, {
  YaumulMabasUyun: { id: "YaumulMabasUyun", "citation-key": "YaumulMabasUyun", title: "Yaumul Mabas - Uyun 1", type: "document" }
}, "parseExportResult keys the parsed CSL-JSON array by item.id");
console.log("parseExportResult (success) test passed");

// --- parseExportResult: defensive fallback keying by requested citekey when item.id absent ---
const noIdShaped = {
  jsonrpc: "2.0",
  result: JSON.stringify([{ title: "No Id Item", type: "document" }]),
  id: 1
};
const parsedNoId = CiteZotero.parseExportResult(noIdShaped, ["SomeKey"]);
assert.ok(parsedNoId.SomeKey, "falls back to keying by requested citekey when item.id is missing");
assert.strictEqual(parsedNoId.SomeKey.title, "No Id Item");
console.log("parseExportResult (defensive fallback) test passed");

// --- parseExportResult: fallback must NOT fabricate when lengths differ ---
const shortArrayShaped = {
  jsonrpc: "2.0",
  result: JSON.stringify([{ id: "B", title: "Item B", type: "document" }]),
  id: 1
};
const parsedShort = CiteZotero.parseExportResult(shortArrayShaped, ["A", "B"]);
assert.ok(
  !Object.prototype.hasOwnProperty.call(parsedShort, "A"),
  "must NOT fabricate an 'A' entry via positional fallback when array length != citekeys length"
);
assert.strictEqual(parsedShort.B.title, "Item B", "correctly-id'd 'B' entry is still present");
console.log("parseExportResult (no-fabrication on length mismatch) test passed");

// --- parseExportResult: error ---
assert.throws(
  () => CiteZotero.parseExportResult({ jsonrpc: "2.0", error: { code: -32603, message: "boom" }, id: null }, ["x"]),
  /boom/,
  "parseExportResult throws with the RPC error message"
);
console.log("parseExportResult (error) test passed");

// --- parseExportResult: error is a plain string (e.g. proxy's 502 body) ---
assert.throws(
  () => CiteZotero.parseExportResult({ error: "boom" }, ["x"]),
  /boom/,
  "parseExportResult throws with the string error itself, not 'undefined'"
);
console.log("parseExportResult (string error) test passed");

// --- parseCaywResult → [{citekey, locator, label}] (SP-A) ---
assert.deepStrictEqual(CiteZotero.parseCaywResult(""), []);
assert.deepStrictEqual(CiteZotero.parseCaywResult("   "), []);
assert.deepStrictEqual(CiteZotero.parseCaywResult(null), []);
assert.deepStrictEqual(CiteZotero.parseCaywResult("@YaumulMabasUyun"), [{ citekey: "YaumulMabasUyun" }]);
assert.deepStrictEqual(CiteZotero.parseCaywResult("[@YaumulMabasUyun]"), [{ citekey: "YaumulMabasUyun" }]);
assert.deepStrictEqual(CiteZotero.parseCaywResult("[@Key, p. 42]"), [{ citekey: "Key", locator: "42", label: "page" }]);
assert.deepStrictEqual(CiteZotero.parseCaywResult("[@Key, pp. 42-45]"), [{ citekey: "Key", locator: "42-45", label: "page" }]);
assert.deepStrictEqual(CiteZotero.parseCaywResult("[@Key, chap. 3]"), [{ citekey: "Key", locator: "3", label: "chapter" }]);
assert.deepStrictEqual(CiteZotero.parseCaywResult("[@Key, sec. 2]"), [{ citekey: "Key", locator: "2", label: "section" }]);
assert.deepStrictEqual(CiteZotero.parseCaywResult("[@Key, v. 7]"), [{ citekey: "Key", locator: "7", label: "verse" }]);
assert.deepStrictEqual(CiteZotero.parseCaywResult("[@Key, 99]"), [{ citekey: "Key", locator: "99", label: "page" }], "bare number ⇒ page");
assert.deepStrictEqual(CiteZotero.parseCaywResult("[@Key, mumble]"), [{ citekey: "Key" }], "unrecognized suffix ⇒ no locator");
assert.deepStrictEqual(
  CiteZotero.parseCaywResult("[@K1, p. 1; @K2]"),
  [{ citekey: "K1", locator: "1", label: "page" }, { citekey: "K2" }]
);
console.log("parseCaywResult test passed");

// --- ping: resolves true/false, never rejects ---
(async () => {
  const okTrue = await CiteZotero.ping(async () => ({ ok: true }));
  assert.strictEqual(okTrue, true, "ping resolves true when res.ok");

  const okFalse = await CiteZotero.ping(async () => ({ ok: false }));
  assert.strictEqual(okFalse, false, "ping resolves false when res.ok is false");

  const thrown = await CiteZotero.ping(async () => { throw new Error("network down"); });
  assert.strictEqual(thrown, false, "ping resolves false (never rejects) on thrown error");
  console.log("ping test passed");
})();

// --- caywPick: GET /zotero/cayw?format=pandoc, parses pandoc text via parseCaywResult ---
(async () => {
  let calledUrl = null;
  const fake = async (url) => {
    calledUrl = url;
    return { ok: true, text: async () => "[@YaumulMabasUyun; @IsraaWalMiraaj]" };
  };
  const keys = await CiteZotero.caywPick(fake);
  assert.deepStrictEqual(keys, [{ citekey: "YaumulMabasUyun" }, { citekey: "IsraaWalMiraaj" }]);
  assert.strictEqual(calledUrl, "/zotero/cayw?format=pandoc", "caywPick fetches the same-origin cayw route with the pandoc format");
  console.log("caywPick test passed");
})();

// --- caywPick: a genuine user cancel is an empty citekeys array from a 200 ---
(async () => {
  const fakeCancel = async () => ({ ok: true, text: async () => "" });
  const keys = await CiteZotero.caywPick(fakeCancel);
  assert.deepStrictEqual(keys, [], "caywPick resolves to [] on a genuine (200, empty-body) cancel");
  console.log("caywPick (cancel) test passed");
})();

// --- caywPick: !res.ok (Zotero offline → proxy 502) must reject, not return garbage citekeys ---
(async () => {
  const fakeOffline = async () => ({
    ok: false,
    status: 502,
    text: async () => JSON.stringify({ error: "zotero-unreachable" })
  });
  let rejected = false;
  try {
    await CiteZotero.caywPick(fakeOffline);
  } catch (e) {
    rejected = true;
    assert.ok(/502/.test(e.message), "caywPick error message includes the HTTP status");
  }
  assert.ok(rejected, "caywPick must reject when res.ok is false");
  console.log("caywPick (!res.ok) test passed");
})();

// --- fetchCslJson: same-origin RPC route, no :23119, correct shape for CiteEngine.build({items}) ---
(async () => {
  CiteZotero.clearCache();
  let callCount = 0;
  let lastBody = null;
  let lastUrl = null;
  let lastOptions = null;
  const fake = async (url, options) => {
    callCount++;
    lastUrl = url;
    lastOptions = options;
    lastBody = JSON.parse(options.body);
    const requested = lastBody.params[0];
    const result = requested.map((id) => ({ id, title: "Title for " + id, type: "document" }));
    return { ok: true, json: async () => ({ jsonrpc: "2.0", result: JSON.stringify(result), id: 1 }) };
  };

  const map1 = await CiteZotero.fetchCslJson(["YaumulMabasUyun"], fake);
  assert.strictEqual(callCount, 1, "first fetch issues exactly one network call");
  assert.strictEqual(lastUrl, "/zotero/json-rpc", "fetchCslJson posts to the same-origin json-rpc route");
  assert.strictEqual(lastOptions.method, "POST");
  assert.strictEqual(lastOptions.headers["Content-Type"], "application/json");
  assert.deepStrictEqual(lastBody, CiteZotero.buildExportRequest(["YaumulMabasUyun"]));
  assert.deepStrictEqual(map1, { YaumulMabasUyun: { id: "YaumulMabasUyun", title: "Title for YaumulMabasUyun", type: "document" } });

  // second call with the SAME key must be served from cache: no new network call
  const map2 = await CiteZotero.fetchCslJson(["YaumulMabasUyun"], fake);
  assert.strictEqual(callCount, 1, "cached key does not trigger a second network call");
  assert.deepStrictEqual(map2, map1);

  // mixed call: one cached key + one new key -> only the new key is fetched
  const map3 = await CiteZotero.fetchCslJson(["YaumulMabasUyun", "IsraaWalMiraaj"], fake);
  assert.strictEqual(callCount, 2, "mixed call fetches only the new key");
  assert.deepStrictEqual(lastBody.params[0], ["IsraaWalMiraaj"], "only the uncached citekey is requested");
  assert.deepStrictEqual(map3, {
    YaumulMabasUyun: { id: "YaumulMabasUyun", title: "Title for YaumulMabasUyun", type: "document" },
    IsraaWalMiraaj: { id: "IsraaWalMiraaj", title: "Title for IsraaWalMiraaj", type: "document" }
  });

  CiteZotero.clearCache();
  console.log("fetchCslJson (cache) test passed");
})();

// --- fetchCslJson([]): resolves to {} and issues zero network calls ---
(async () => {
  CiteZotero.clearCache();
  let callCount = 0;
  const fake = async (url, options) => {
    callCount++;
    return { ok: true, json: async () => ({ jsonrpc: "2.0", result: "[]", id: 1 }) };
  };
  const map = await CiteZotero.fetchCslJson([], fake);
  assert.deepStrictEqual(map, {}, "fetchCslJson([]) resolves to {}");
  assert.strictEqual(callCount, 0, "fetchCslJson([]) issues zero network calls");
  console.log("fetchCslJson (empty citekeys) test passed");
})();

// --- fetchCslJson: !res.ok (Zotero offline → proxy 502) must reject, not parse the error body ---
// NOTE: deliberately does NOT call clearCache() and uses a citekey unique to
// this test — other IIFEs in this file run concurrently (none are awaited at
// the top level) and also touch the shared module-level cache; a clearCache()
// here could race between another test's cache write and its subsequent read.
(async () => {
  const fakeOffline = async () => ({
    ok: false,
    status: 502,
    json: async () => ({ error: "zotero-unreachable" })
  });
  let rejected = false;
  try {
    await CiteZotero.fetchCslJson(["OfflineProbeKey"], fakeOffline);
  } catch (e) {
    rejected = true;
    assert.ok(/502/.test(e.message), "fetchCslJson error message includes the HTTP status");
  }
  assert.ok(rejected, "fetchCslJson must reject when res.ok is false");
  console.log("fetchCslJson (!res.ok) test passed");
})();

// --- no direct :23119 reference anywhere in the module source ---
const fs = require("fs");
const path = require("path");
const src = fs.readFileSync(path.join(__dirname, "..", "src", "taskpane", "cite-zotero.js"), "utf8");
assert.ok(!/23119/.test(src), "module must never reference the Zotero port directly");
assert.ok(
  !/document\./.test(src) && !/window\./.test(src) && !/Office\./.test(src) && !/navigator\./.test(src),
  "module must have no DOM/Office.js usage"
);
console.log("no-:23119/no-DOM guard test passed");

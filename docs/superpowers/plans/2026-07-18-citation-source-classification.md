# SP-4 Source Classification & Nested-Section Bibliography — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Classify references by two independent Zotero-tag axes (corpus: Fatemi/non-Fatemi; class: primary/secondary) and render bibliographies segmented into nested headed subsections when a `-fatemi` CSL style is selected.

**Architecture:** A new pure module `cite-classify.js` turns per-citekey Zotero tags into ordered non-empty buckets and a section plan (with the collapse-to-flat rule baked in). `cite-zotero.js` gains a `fetchTags()` I/O helper (a second `item.export` over the same citekeys via a tag-carrying translator). `cite-word.js` gains two pure assemblers that turn `[{heading, html}]` sections into the final HTML / OOXML. `cite-pane.js` wires it together: when the selected style ends in `-fatemi`, it fetches tags, plans sections, renders each bucket over a filtered item map, and inserts. Stock styles and single-bucket cases render one flat bibliography byte-identical to today.

**Tech Stack:** Vanilla ES5/UMD modules (no build step), citeproc-js (CSL-M), Office.js Word API, Node `assert` tests. Better BibTeX json-rpc over the existing `/zotero/json-rpc` reverse-proxy route.

## Global Constraints

- **No build step / no transpilation.** All modules are ES5 UMD (`var`, no arrow functions, no `const`/`let`) matching the existing `cite-*.js` files.
- **UMD module pattern** exactly as the existing cite modules (see `cite-word.js` lines 1–7): `module.exports` for Node, `root.<Name>` for browser.
- **Pure modules take no I/O.** `cite-classify.js` and the `cite-word.js` assemblers must be Node-testable with no DOM/Word/fetch.
- **Classification is bibliography-only.** Do not touch inline/footnote/endnote citation rendering.
- **Collapse rule:** if fewer than 2 buckets are non-empty, render flat with NO heading — byte-identical to today's output.
- **Degrade to flat on tag-fetch failure.** A Zotero/BBT error must never block bibliography insertion.
- **Fixed bucket order:** `primary.fatemi` → `primary.other` → `secondary.fatemi` → `secondary.other`.
- **Tag prefixes:** `corpus:fatemi` (absent = non-Fatemi/`other`); `class:secondary` (absent or `class:primary` = `primary`). Only these two prefixes are consulted; unknown tags ignored; malformed axis values fall back to default.
- **Every new test file must be registered** in the `package.json` `test` script chain.
- **Bump `ASHAAR_ASSET_VERSION`** in `taskpane.html` when pane code changes (installed users cache JS by this version).

---

### Task 1: `cite-classify.js` — pure classification & section planning

**Files:**
- Create: `src/taskpane/cite-classify.js`
- Test: `tests/cite-classify.test.js`
- Modify: `package.json:26` (register the test)

**Interfaces:**
- Consumes: nothing (pure).
- Produces (global `CiteClassify` / `module.exports`):
  - `bucketForTags(tags)` → `{ corpus: "fatemi"|"other", cls: "primary"|"secondary", key: string }` where `key = cls + "." + corpus`. `tags` is an array of strings.
  - `BUCKET_ORDER` → `["primary.fatemi","primary.other","secondary.fatemi","secondary.other"]`.
  - `orderedBuckets(citekeys, tagsByCitekey)` → array of `{ key, citekeys: [...] }` for NON-EMPTY buckets only, in `BUCKET_ORDER`, preserving input order of citekeys within each bucket. `tagsByCitekey` is `{ citekey: [tagStrings] }`; a citekey absent from the map is treated as no tags (default bucket).
  - `headingFor(bucketKey, lang)` → localized heading string (en default; ar when `lang` starts with `ar`).
  - `planBibliographySections(citekeys, tagsByCitekey, opts)` → array of `{ key: string|null, heading: string|null, citekeys: [...] }`. `opts = { sectioned: bool, lang: string }`. When `!sectioned` OR fewer than 2 non-empty buckets: returns a single section `{ key, heading: null, citekeys }` (collapse rule). Otherwise one section per non-empty bucket with `heading` resolved via `headingFor`.

- [ ] **Step 1: Write the failing test**

Create `tests/cite-classify.test.js`:

```js
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
```

- [ ] **Step 2: Run test to verify it fails**

Run: `node tests/cite-classify.test.js`
Expected: FAIL — `Cannot find module '../src/taskpane/cite-classify'`.

- [ ] **Step 3: Write minimal implementation**

Create `src/taskpane/cite-classify.js`:

```js
(function (root, factory) {
  if (typeof module !== "undefined" && module.exports) {
    module.exports = factory();
  } else {
    root.CiteClassify = factory();
  }
}(typeof globalThis !== "undefined" ? globalThis : this, function () {
  "use strict";

  var BUCKET_ORDER = ["primary.fatemi", "primary.other", "secondary.fatemi", "secondary.other"];

  var HEADINGS = {
    en: {
      "primary.fatemi": "Primary Sources — Fatemi",
      "primary.other": "Primary Sources — Other",
      "secondary.fatemi": "Secondary Sources — Fatemi",
      "secondary.other": "Secondary Sources — Other"
    },
    ar: {
      "primary.fatemi": "المصادر الأساسية — الفاطمية",
      "primary.other": "المصادر الأساسية — أخرى",
      "secondary.fatemi": "المصادر الثانوية — الفاطمية",
      "secondary.other": "المصادر الثانوية — أخرى"
    }
  };

  function bucketForTags(tags) {
    var list = tags || [];
    var corpus = "other";
    var cls = "primary";
    for (var i = 0; i < list.length; i++) {
      if (list[i] === "corpus:fatemi") { corpus = "fatemi"; }
      else if (list[i] === "class:secondary") { cls = "secondary"; }
      // class:primary and everything else leave the defaults in place
    }
    return { corpus: corpus, cls: cls, key: cls + "." + corpus };
  }

  function orderedBuckets(citekeys, tagsByCitekey) {
    var map = tagsByCitekey || {};
    var groups = {};
    (citekeys || []).forEach(function (ck) {
      var tags = Object.prototype.hasOwnProperty.call(map, ck) ? map[ck] : [];
      var key = bucketForTags(tags).key;
      if (!groups[key]) { groups[key] = []; }
      groups[key].push(ck);
    });
    var out = [];
    BUCKET_ORDER.forEach(function (key) {
      if (groups[key] && groups[key].length) { out.push({ key: key, citekeys: groups[key] }); }
    });
    return out;
  }

  function headingFor(bucketKey, lang) {
    var table = (/^ar/i.test(lang || "")) ? HEADINGS.ar : HEADINGS.en;
    return table[bucketKey] || bucketKey;
  }

  function planBibliographySections(citekeys, tagsByCitekey, opts) {
    var o = opts || {};
    var keys = (citekeys || []).slice();
    if (o.sectioned) {
      var buckets = orderedBuckets(keys, tagsByCitekey);
      if (buckets.length >= 2) {
        return buckets.map(function (b) {
          return { key: b.key, heading: headingFor(b.key, o.lang), citekeys: b.citekeys };
        });
      }
      // exactly one (or zero) non-empty bucket => collapse to flat
      if (buckets.length === 1) {
        return [{ key: buckets[0].key, heading: null, citekeys: buckets[0].citekeys }];
      }
    }
    return [{ key: null, heading: null, citekeys: keys }];
  }

  return {
    BUCKET_ORDER: BUCKET_ORDER,
    bucketForTags: bucketForTags,
    orderedBuckets: orderedBuckets,
    headingFor: headingFor,
    planBibliographySections: planBibliographySections
  };
}));
```

- [ ] **Step 4: Run test to verify it passes**

Run: `node tests/cite-classify.test.js`
Expected: PASS — `cite-classify.test.js passed`.

- [ ] **Step 5: Register the test in the npm chain**

In `package.json`, the `test` script (line 26) ends with `&& node tests/cite-store.test.js`. Append:

```
 && node tests/cite-classify.test.js
```

Run: `npm test`
Expected: full suite PASS, including `cite-classify.test.js passed`.

- [ ] **Step 6: Commit**

```bash
git add src/taskpane/cite-classify.js tests/cite-classify.test.js package.json
git commit -m "feat(cite): cite-classify — tags -> buckets + section plan (SP-4)

Co-Authored-By: Claude Opus 4.8 <noreply@anthropic.com>
Claude-Session: https://claude.ai/code/session_01Vy1kwq3DuKd6bpgQGdN6Yw"
```

---

### Task 2: `cite-zotero.js` — `fetchTags()` + pure request/response helpers

**Files:**
- Modify: `src/taskpane/cite-zotero.js` (add helpers + export; the file's structure is at lines 1–186)
- Test: `tests/cite-zotero.test.js` (append cases)

**Interfaces:**
- Consumes: the existing `/zotero/json-rpc` proxy route (POST, `Content-Type: application/json`).
- Produces (added to the `CiteZotero` exports):
  - `buildTagsRequest(citekeys)` → BBT json-rpc body object: `{ jsonrpc:"2.0", method:"item.export", params:[citekeys, "BetterBibTeX JSON"], id:1 }`.
  - `parseTagsResult(rpcResponse, citekeys)` → `{ citekey: [tagStrings] }`. Throws on `rpcResponse.error` (same shape as `parseExportResult`). Reads BBT-JSON shape `{ items: [ { citationKey|citekey, tags: [{tag}]|[string] } ] }`, tolerating `tags` entries that are either `{tag:"..."}` objects or plain strings. Citekeys with no export entry are omitted (caller defaults them).
  - `fetchTags(citekeys, fetchImpl)` → Promise of `{ citekey: [tagStrings] }`, caching per citekey (parallel to `fetchCslJson`'s `cache`).
- `clearCache()` must also clear the new tag cache.

> **Live-confirm note (do during Task 5 manual verification, not now):** the translator name `"BetterBibTeX JSON"` and the exact `tags` shape are confirmed against the running Zotero+BBT. If `item.export` rejects that translator name, the documented fallback is the Zotero local API (`localhost:23119/api/users/0/items`, already used read-only by `scripts/migrate-mlzsync-to-cne.mjs`), correlating item keys to citekeys. Keep `parseTagsResult` pure and adjust its shape-reading to whatever the live translator emits; the tests below pin the contract, so update them alongside any shape change.

- [ ] **Step 1: Write the failing test**

Append to `tests/cite-zotero.test.js` (before any final `console.log`; if the file ends with a summary log, add these above it):

```js
// --- SP-4: tag fetch helpers ---
(function () {
  var Z = require("../src/taskpane/cite-zotero");

  // buildTagsRequest
  var req = Z.buildTagsRequest(["k1", "k2"]);
  assert.strictEqual(req.method, "item.export");
  assert.deepStrictEqual(req.params, [["k1", "k2"], "BetterBibTeX JSON"]);

  // parseTagsResult — object-shaped tags
  var rpc = { result: JSON.stringify({ items: [
    { citationKey: "k1", tags: [{ tag: "corpus:fatemi" }, { tag: "class:secondary" }] },
    { citationKey: "k2", tags: [] }
  ] }) };
  var parsed = Z.parseTagsResult(rpc, ["k1", "k2"]);
  assert.deepStrictEqual(parsed.k1, ["corpus:fatemi", "class:secondary"]);
  assert.deepStrictEqual(parsed.k2, []);

  // parseTagsResult — string-shaped tags + alternate key field
  var rpc2 = { result: JSON.stringify({ items: [
    { citekey: "k3", tags: ["corpus:fatemi"] }
  ] }) };
  assert.deepStrictEqual(Z.parseTagsResult(rpc2, ["k3"]).k3, ["corpus:fatemi"]);

  // parseTagsResult — error surfaces
  assert.throws(function () { Z.parseTagsResult({ error: { message: "boom" } }, ["k1"]); }, /boom/);

  // fetchTags — uses a fake fetch, caches, returns per-citekey tags
  Z.clearCache();
  var calls = 0;
  function fakeFetch(url, init) {
    calls++;
    assert.strictEqual(url, "/zotero/json-rpc");
    assert.strictEqual(init.method, "POST");
    return Promise.resolve({
      ok: true,
      json: function () {
        return Promise.resolve({ result: JSON.stringify({ items: [
          { citationKey: "k1", tags: [{ tag: "corpus:fatemi" }] }
        ] }) });
      }
    });
  }
  return Z.fetchTags(["k1"], fakeFetch).then(function (map) {
    assert.deepStrictEqual(map.k1, ["corpus:fatemi"]);
    // second call for the same key hits the cache (no extra fetch)
    return Z.fetchTags(["k1"], fakeFetch).then(function () {
      assert.strictEqual(calls, 1);
      console.log("cite-zotero SP-4 tag tests passed");
    });
  });
})();
```

> If `tests/cite-zotero.test.js` currently runs sequentially with no async harness, wrap the async block so it actually executes — mirror how the file's existing `fetchCslJson` test is structured (it already awaits a fake-fetch promise; follow that exact pattern rather than the IIFE above if the file uses a different runner).

- [ ] **Step 2: Run test to verify it fails**

Run: `node tests/cite-zotero.test.js`
Expected: FAIL — `Z.buildTagsRequest is not a function`.

- [ ] **Step 3: Write minimal implementation**

In `src/taskpane/cite-zotero.js`, add a `tagCache` alongside the existing `cache` (near line 12):

```js
  // In-memory citekey -> [tagStrings] cache (parallel to `cache`).
  var tagCache = {};
```

Update `clearCache` (line 14) to also reset it:

```js
  function clearCache() {
    cache = {};
    tagCache = {};
  }
```

Add the pure helpers (place them next to `buildExportRequest`/`parseExportResult`, ~line 60):

```js
  function buildTagsRequest(citekeys) {
    return {
      jsonrpc: "2.0",
      method: "item.export",
      params: [citekeys, "BetterBibTeX JSON"],
      id: 1
    };
  }

  function normalizeTag(t) {
    if (t && typeof t === "object") { return t.tag; }
    return t;
  }

  function parseTagsResult(rpcResponse, citekeys) {
    if (rpcResponse && rpcResponse.error) {
      var err = rpcResponse.error;
      var message = (err && err.message) ? err.message
        : (typeof err === "string" ? err : JSON.stringify(err));
      throw new Error(message);
    }
    var payload = JSON.parse(rpcResponse.result) || {};
    var items = payload.items || [];
    var map = {};
    items.forEach(function (item) {
      if (!item) { return; }
      var key = item.citationKey || item.citekey || item.id;
      if (key === undefined || key === null) { return; }
      var tags = (item.tags || []).map(normalizeTag).filter(function (t) {
        return typeof t === "string" && t.length > 0;
      });
      map[key] = tags;
    });
    return map;
  }
```

Add the I/O `fetchTags` (place next to `fetchCslJson`, ~line 176):

```js
  function fetchTags(citekeys, fetchImpl) {
    var f = fetchImpl || (typeof fetch !== "undefined" ? fetch : undefined);
    var missing = (citekeys || []).filter(function (key) {
      return !Object.prototype.hasOwnProperty.call(tagCache, key);
    });
    var fetchStep = missing.length === 0
      ? Promise.resolve(null)
      : Promise.resolve()
        .then(function () {
          return f("/zotero/json-rpc", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify(buildTagsRequest(missing))
          });
        })
        .then(function (res) {
          if (!res.ok) { throw new Error("json-rpc HTTP " + res.status); }
          return res.json();
        })
        .then(function (rpcResponse) {
          var parsed = parseTagsResult(rpcResponse, missing);
          // Cache every REQUESTED key so an item with no tags (absent from the
          // export) is remembered as [] rather than re-fetched every time.
          missing.forEach(function (key) {
            tagCache[key] = Object.prototype.hasOwnProperty.call(parsed, key) ? parsed[key] : [];
          });
        });
    return fetchStep.then(function () {
      var result = {};
      (citekeys || []).forEach(function (key) {
        result[key] = Object.prototype.hasOwnProperty.call(tagCache, key) ? tagCache[key] : [];
      });
      return result;
    });
  }
```

Add all four to the exports object (the `return { ... }` at ~line 179):

```js
    buildTagsRequest: buildTagsRequest,
    parseTagsResult: parseTagsResult,
    fetchTags: fetchTags,
```

- [ ] **Step 4: Run test to verify it passes**

Run: `node tests/cite-zotero.test.js`
Expected: PASS — includes `cite-zotero SP-4 tag tests passed`.

- [ ] **Step 5: Run the full suite**

Run: `npm test`
Expected: full suite PASS (no regression in the existing cite-zotero cases).

- [ ] **Step 6: Commit**

```bash
git add src/taskpane/cite-zotero.js tests/cite-zotero.test.js
git commit -m "feat(cite): cite-zotero.fetchTags — per-citekey Zotero tags via BBT export (SP-4)

Co-Authored-By: Claude Opus 4.8 <noreply@anthropic.com>
Claude-Session: https://claude.ai/code/session_01Vy1kwq3DuKd6bpgQGdN6Yw"
```

---

### Task 3: `cite-word.js` — pure sectioned-bibliography assemblers

**Files:**
- Modify: `src/taskpane/cite-word.js` (add two functions + exports; existing helpers `sanitize`, `wrapRtlRuns`, `buildCitationParagraphOoxml`, `xmlEsc` are at lines 13, 40, 231, 149)
- Test: `tests/cite-word.test.js` (append cases)

**Interfaces:**
- Consumes: existing `sanitize(html)`, `wrapRtlRuns(html)`, `buildCitationParagraphOoxml(html, opts)`, `xmlEsc(s)` (all already in this file).
- Produces (added to `CiteWord` exports):
  - `buildSectionedBibliographyHtml(sections)` → concatenated HTML string. `sections` is `[{ heading: string|null, html: string }]`. For each section: body = `wrapRtlRuns(sanitize(html))`; if `heading` truthy, prefix `"<p><b>" + htmlEsc(heading) + "</b></p>"`. Join with `""`. A single heading-null section returns exactly `wrapRtlRuns(sanitize(html))` — byte-identical to today's non-RTL output.
  - `buildSectionedBibliographyOoxml(sections, opts)` → concatenated OOXML string (no `<w:body>` wrapper — the caller passes it to `AshaarTabStop.wrapOoxml`). For each section: body para = `buildCitationParagraphOoxml(sanitize(html), opts)`; if `heading` truthy, prepend a heading para = `buildCitationParagraphOoxml("<b>" + heading + "</b>", opts)`. `opts` carries `{ csFont }`. A single heading-null section returns exactly `buildCitationParagraphOoxml(sanitize(html), opts)` — byte-identical to today's RTL output.

- [ ] **Step 1: Write the failing test**

Append to `tests/cite-word.test.js` (above any final summary log):

```js
// --- SP-4: sectioned bibliography assemblers ---
(function () {
  // Single heading-null section == today's flat output (HTML)
  var flatHtml = CiteWord.buildSectionedBibliographyHtml([{ heading: null, html: "<div>Entry one.</div>" }]);
  assert.strictEqual(flatHtml, CiteWord.wrapRtlRuns(CiteWord.sanitize("<div>Entry one.</div>")));

  // Multi-section HTML: each heading becomes <p><b>…</b></p> then its body
  var secHtml = CiteWord.buildSectionedBibliographyHtml([
    { heading: "Primary Sources — Fatemi", html: "<div>A.</div>" },
    { heading: "Secondary Sources — Other", html: "<div>B.</div>" }
  ]);
  assert.ok(secHtml.indexOf("<p><b>Primary Sources — Fatemi</b></p>") !== -1);
  assert.ok(secHtml.indexOf("<p><b>Secondary Sources — Other</b></p>") !== -1);
  assert.ok(secHtml.indexOf("Primary Sources") < secHtml.indexOf("Secondary Sources")); // order

  // heading HTML-escaping (defensive)
  var escd = CiteWord.buildSectionedBibliographyHtml([{ heading: "A & <B>", html: "<div>x</div>" }]);
  assert.ok(escd.indexOf("A &amp; &lt;B&gt;") !== -1);

  // Single heading-null section == today's flat output (OOXML)
  var flatOoxml = CiteWord.buildSectionedBibliographyOoxml([{ heading: null, html: "<div>Entry.</div>" }], { csFont: "Scheherazade" });
  assert.strictEqual(flatOoxml, CiteWord.buildCitationParagraphOoxml(CiteWord.sanitize("<div>Entry.</div>"), { csFont: "Scheherazade" }));

  // Multi-section OOXML: heading para (bold) precedes each body para
  var secOoxml = CiteWord.buildSectionedBibliographyOoxml([
    { heading: "المصادر الأساسية — الفاطمية", html: "<div>A.</div>" },
    { heading: "المصادر الثانوية — أخرى", html: "<div>B.</div>" }
  ], { csFont: "Scheherazade" });
  assert.ok(secOoxml.indexOf("<w:b/>") !== -1);               // heading is bold
  assert.ok(secOoxml.indexOf("المصادر الأساسية") !== -1);
  assert.ok(secOoxml.indexOf("المصادر الأساسية") < secOoxml.indexOf("المصادر الثانوية")); // order

  console.log("cite-word SP-4 section tests passed");
})();
```

- [ ] **Step 2: Run test to verify it fails**

Run: `node tests/cite-word.test.js`
Expected: FAIL — `CiteWord.buildSectionedBibliographyHtml is not a function`.

- [ ] **Step 3: Write minimal implementation**

In `src/taskpane/cite-word.js`, add these two functions just above the `return { ... }` exports (after `buildCitationParagraphOoxml`, ~line 234):

```js
  function htmlEsc(s) {
    return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  }

  function buildSectionedBibliographyHtml(sections) {
    return (sections || []).map(function (s) {
      var body = wrapRtlRuns(sanitize(s.html));
      return s.heading ? ("<p><b>" + htmlEsc(s.heading) + "</b></p>" + body) : body;
    }).join("");
  }

  function buildSectionedBibliographyOoxml(sections, opts) {
    return (sections || []).map(function (s) {
      var body = buildCitationParagraphOoxml(sanitize(s.html), opts);
      if (!s.heading) { return body; }
      var head = buildCitationParagraphOoxml("<b>" + s.heading + "</b>", opts);
      return head + body;
    }).join("");
  }
```

Add to the exports object:

```js
    buildSectionedBibliographyHtml: buildSectionedBibliographyHtml,
    buildSectionedBibliographyOoxml: buildSectionedBibliographyOoxml,
```

> Note: the OOXML heading passes `"<b>" + s.heading + "</b>"` straight through `htmlToOoxmlRuns`, which xml-escapes the text content itself (`xmlEsc` inside `emitRun`) and auto-detects RTL per character — so Arabic headings render as bold RTL runs and en headings as bold LTR runs with no extra handling.

- [ ] **Step 4: Run test to verify it passes**

Run: `node tests/cite-word.test.js`
Expected: PASS — includes `cite-word SP-4 section tests passed`.

- [ ] **Step 5: Run the full suite**

Run: `npm test`
Expected: full suite PASS.

- [ ] **Step 6: Commit**

```bash
git add src/taskpane/cite-word.js tests/cite-word.test.js
git commit -m "feat(cite): sectioned bibliography assemblers (HTML+OOXML) (SP-4)

Co-Authored-By: Claude Opus 4.8 <noreply@anthropic.com>
Claude-Session: https://claude.ai/code/session_01Vy1kwq3DuKd6bpgQGdN6Yw"
```

---

### Task 4: Remove the inert `[genre]` branch from the Fatemi styles + re-sync

**Files:**
- Modify: `src/styles/chicago-notes-fatemi.csl` (remove inserted branch, ~lines 5974–5978)
- Modify: `src/styles/apa-fatemi.csl` (remove inserted branch, ~lines 2260–2264)
- Modify (via sync or direct copy): `src/vendor/csl-styles/chicago-notes-fatemi.csl`, `src/vendor/csl-styles/apa-fatemi.csl`

**Interfaces:**
- Consumes: nothing.
- Produces: `-fatemi` CSL files byte-identical to their stock parents except `<title>`/`<id>`. Classification no longer flows through any CSL variable.

The inserted branch in **both** files is exactly:

```xml
        <if variable="genre" match="any">
          <text variable="genre" prefix="[" suffix="] "/>
        </if>
      </choose>
      <choose>
```

This was spliced into the bibliography `<layout>`, splitting one `<choose>` into two around the inert `<if>`. Removing it must **rejoin** the split: delete the `<if>…</if>` block AND the extra `</choose>` + `<choose>` pair it introduced, leaving the original single `<choose>` intact.

- [ ] **Step 1: Confirm the exact splice in each file**

Run: `diff <(grep -v '<id>\|<title>' src/vendor/csl-styles/chicago-notes-bibliography.csl) <(grep -v '<id>\|<title>' src/styles/chicago-notes-fatemi.csl)`
Expected: shows ONLY the 5 added lines above (a `5973a5974,5978` hunk). Repeat for apa:
Run: `diff <(grep -v '<id>\|<title>' src/vendor/csl-styles/apa.csl) <(grep -v '<id>\|<title>' src/styles/apa-fatemi.csl)`
Expected: the same 5-line addition (a `2259a2260,2264` hunk).

- [ ] **Step 2: Remove the branch from `src/styles/chicago-notes-fatemi.csl`**

The stock parent has a single `<choose>` where the fatemi file has `…<choose> [inert if] </choose> <choose>…`. Delete the inserted lines so the two `<choose>` blocks rejoin into one. Concretely, find this fragment:

```xml
        </if>
        <if variable="genre" match="any">
          <text variable="genre" prefix="[" suffix="] "/>
        </if>
      </choose>
      <choose>
        <if ...(next real branch)...>
```

and remove the four inserted lines plus the `</choose>\n      <choose>` rejoin so it reads:

```xml
        </if>
        <if ...(next real branch)...>
```

Use the confirming diff from Step 1 to see the precise surrounding lines in this specific file; the `<if variable="genre" match="any">` occurrence is unique (verify: `grep -n 'variable="genre" match="any"' src/styles/chicago-notes-fatemi.csl` returns exactly one line).

- [ ] **Step 3: Remove the branch from `src/styles/apa-fatemi.csl`**

Same operation. Verify uniqueness: `grep -n 'variable="genre" match="any"' src/styles/apa-fatemi.csl` returns exactly one line; remove that inserted block and rejoin the split `<choose>`.

- [ ] **Step 4: Verify each fatemi file now differs from its stock parent ONLY by title/id**

Run: `diff <(grep -v '<id>\|<title>' src/vendor/csl-styles/chicago-notes-bibliography.csl) <(grep -v '<id>\|<title>' src/styles/chicago-notes-fatemi.csl)`
Expected: **no output** (identical apart from the filtered title/id).
Run: `diff <(grep -v '<id>\|<title>' src/vendor/csl-styles/apa.csl) <(grep -v '<id>\|<title>' src/styles/apa-fatemi.csl)`
Expected: **no output**.

- [ ] **Step 5: Propagate to the vendored copies**

Preferred (if the citeproc submodules are initialized): `npm run sync:citeproc`
Expected: `citeproc vendor sync complete`, and `src/vendor/csl-styles/*-fatemi.csl` now match `src/styles/*-fatemi.csl`.

If the submodules are NOT initialized (sync errors on a git call), copy the two files directly instead:

```bash
cp src/styles/chicago-notes-fatemi.csl src/vendor/csl-styles/chicago-notes-fatemi.csl
cp src/styles/apa-fatemi.csl src/vendor/csl-styles/apa-fatemi.csl
```

Then verify: `grep -c 'variable="genre" match="any"' src/vendor/csl-styles/chicago-notes-fatemi.csl src/vendor/csl-styles/apa-fatemi.csl`
Expected: `0` for both.

- [ ] **Step 6: Full suite still green (styles load + engine tests)**

Run: `npm test`
Expected: PASS (the `cite-engine` / `citeproc-vendor` tests still load and render these styles).

- [ ] **Step 7: Commit**

```bash
git add src/styles/chicago-notes-fatemi.csl src/styles/apa-fatemi.csl src/vendor/csl-styles/chicago-notes-fatemi.csl src/vendor/csl-styles/apa-fatemi.csl
git commit -m "refactor(cite): drop inert [genre] branch from Fatemi styles (SP-4)

Classification now flows through app-level sectioning, not a CSL variable;
the two -fatemi styles are now byte-identical to their stock parents except
title/id and serve purely as the nested-section toggle.

Co-Authored-By: Claude Opus 4.8 <noreply@anthropic.com>
Claude-Session: https://claude.ai/code/session_01Vy1kwq3DuKd6bpgQGdN6Yw"
```

---

### Task 5: `cite-pane.js` — wire sectioning into insert + refresh

**Files:**
- Modify: `src/taskpane/cite-pane.js` (`buildEngine` line 108; `renderBibliographyInto` line 327; `insertBibliography` line 415; `refreshCitations` line 474; add helpers)
- Modify: `src/taskpane/taskpane.html` (add `./cite-classify.js` to the script list at line 676; bump `ASHAAR_ASSET_VERSION` at line 15)

**Interfaces:**
- Consumes: `CiteClassify.planBibliographySections` (Task 1); `CiteZotero.fetchTags` (Task 2); `CiteWord.buildSectionedBibliographyHtml` / `buildSectionedBibliographyOoxml` (Task 3).
- Produces: no new public surface — internal wiring only. `buildEngine` gains an optional third param `itemsOverride`.

> This module is DOM/Word-bound and has no unit-test harness (there is no `cite-pane.test.js`); its pure logic already lives in Tasks 1 & 3. Verification here is `npm test` (regression — the module must still parse and not break the shared suite via any accidental cross-file change) plus the manual Word checklist in Task 6. Follow the existing async/`Word.run` patterns in the file exactly.

- [ ] **Step 1: Add an item-subset param to `buildEngine`**

Replace `buildEngine` (line 108) so it can render over a filtered map:

```js
  function buildEngine(styleFile, lang, itemsOverride) {
    return CiteEngine.build({
      styleXml: cache.styles[styleFile],
      locales: cache.locales,
      items: itemsOverride || cache.items,
      lang: lang,
      langPrefs: (typeof CiteVariants !== "undefined")
        ? CiteVariants.variantToLangPrefs(currentVariant())
        : null
    });
  }
```

- [ ] **Step 2: Add sectioning helpers (place after `buildEngine`, ~line 119)**

```js
  function isFatemiStyle(styleFile) { return /-fatemi$/.test(styleFile || ""); }

  // Fetch tags for the reference set only when the selected style opts in
  // (a -fatemi style). Degrades to an empty tag map on any Zotero/BBT error —
  // an empty map => every item in the default bucket => single bucket =>
  // collapse rule => flat bibliography.
  function fetchTagsIfSectioned(styleFile) {
    if (!isFatemiStyle(styleFile) || !cache.items || typeof CiteZotero === "undefined" || !CiteZotero.fetchTags) {
      return Promise.resolve({});
    }
    return CiteZotero.fetchTags(Object.keys(cache.items)).catch(function () { return {}; });
  }

  // Turn a section plan into insertion-ready body (HTML for LTR, OOXML for RTL).
  // Each section is rendered by a fresh engine over ONLY that section's items.
  function renderBibliographyBody(styleFile, lang, sections, csFont) {
    var rendered = (sections || []).map(function (s) {
      var subset = {};
      (s.citekeys || []).forEach(function (k) {
        if (cache.items && Object.prototype.hasOwnProperty.call(cache.items, k)) { subset[k] = cache.items[k]; }
      });
      var engine = buildEngine(styleFile, lang, subset);
      return { heading: s.heading, html: engine.bibliography() };
    });
    if (isRtlLang(lang)) {
      return { rtl: true, ooxml: AshaarTabStop.wrapOoxml(CiteWord.buildSectionedBibliographyOoxml(rendered, { csFont: csFont })) };
    }
    return { rtl: false, html: CiteWord.buildSectionedBibliographyHtml(rendered) };
  }
```

- [ ] **Step 3: Route `renderBibliographyInto` (used by refresh) through the section plan**

Replace `renderBibliographyInto` (line 327). It gains a `sections` param; when omitted it defaults to a single whole-library flat section, preserving the old behavior:

```js
  function renderBibliographyInto(ctx, range, styleFile, lang, csFont, sections) {
    var secs = sections || [{ key: null, heading: null, citekeys: cache.items ? Object.keys(cache.items) : [] }];
    if (isRtlLang(lang)) {
      var csFontPromise = csFont ? Promise.resolve(csFont) : readDocCsFont(ctx, range);
      return csFontPromise.then(function (resolvedCsFont) {
        var body = renderBibliographyBody(styleFile, lang, secs, resolvedCsFont);
        return range.insertOoxml(body.ooxml, Word.InsertLocation.replace);
      });
    }
    var body = renderBibliographyBody(styleFile, lang, secs);
    return Promise.resolve(range.insertHtml(body.html, Word.InsertLocation.replace));
  }
```

- [ ] **Step 4: Route `insertBibliography` through tag-fetch + section plan**

Replace the body of `insertBibliography` (lines 415–457). Fetch tags first, plan sections, then insert using `renderBibliographyBody`. The RTL branch inserts OOXML; the LTR branch inserts HTML (wrapped in `<p dir="rtl">` only when RTL — which never applies on the LTR branch, so no wrapper needed there):

```js
  function insertBibliography() {
    var styleFile = currentStyleFile();
    var lang = currentLang();
    var rtl = isRtlLang(lang);
    ensureAssets(styleFile).then(function () {
      return fetchTagsIfSectioned(styleFile).then(function (tagsByCitekey) {
        var allKeys = cache.items ? Object.keys(cache.items) : [];
        var sections = CiteClassify.planBibliographySections(allKeys, tagsByCitekey,
          { sectioned: isFatemiStyle(styleFile), lang: lang });
        var bibTag = CiteWord.buildBibliographyTag({ style: styleFile, locale: lang, variant: currentVariant() });
        if (typeof Word === "undefined" || !Word.run) {
          setStatus("Word isn't available — this is preview-only in a browser.", true);
          return;
        }
        return Word.run(function (ctx) {
          var selRange = ctx.document.getSelection().getRange();
          if (rtl) {
            return readDocCsFont(ctx, selRange).then(function (csFont) {
              var body = renderBibliographyBody(styleFile, lang, sections, csFont);
              var oRange = selRange.insertOoxml(body.ooxml, Word.InsertLocation.after);
              var occ = oRange.insertContentControl();
              occ.tag = bibTag;
              occ.title = "Ashaar Bibliography";
              return ctx.sync();
            });
          }
          var body = renderBibliographyBody(styleFile, lang, sections);
          var range = selRange.insertHtml(body.html, Word.InsertLocation.after);
          var cc = range.insertContentControl();
          cc.tag = bibTag;
          cc.title = "Ashaar Bibliography";
          return ctx.sync();
        }).then(function () {
          setStatus("Inserted bibliography.");
        }).catch(function (e) {
          setStatus("Insert failed: " + (e && e.message ? e.message : String(e)), true);
        });
      });
    }).catch(function (e) {
      setStatus("Couldn't load citation assets: " + (e && e.message ? e.message : String(e)), true);
    });
  }
```

> This drops the old `buildBibliographyPayload`/`bibHtml` locals — `renderBibliographyBody` now owns sanitize + RTL-run wrapping for both paths. Confirm `buildBibliographyPayload` has no other caller before assuming it's dead: `grep -n buildBibliographyPayload src/taskpane/*.js` — it is still exported from `cite-word.js`; leave the export (harmless), just stop calling it here.

- [ ] **Step 5: Fetch tags + plan sections once in `refreshCitations`, pass to each bib CC**

In `refreshCitations` (line 474), after `ensureAssets(styleFile).then(function () {` (line 482), fetch tags and build the section plan BEFORE `Word.run`, then thread `sections` into the `renderBibliographyInto` call. Change the start of the `.then` to:

```js
    ensureAssets(styleFile).then(function () {
      return fetchTagsIfSectioned(styleFile).then(function (tagsByCitekey) {
        var allKeys = cache.items ? Object.keys(cache.items) : [];
        var bibSections = CiteClassify.planBibliographySections(allKeys, tagsByCitekey,
          { sectioned: isFatemiStyle(styleFile), lang: lang });
        return Word.run(function (ctx) {
```

Then update the bib branch (line 561) call to pass `bibSections`:

```js
                    ops.push(renderBibliographyInto(ctx, bibRange, styleFile, lang, csFont, bibSections).then(function () {
                      counts.bibs++;
                    }).catch(function () { counts.failed++; }));
```

Close the extra `fetchTagsIfSectioned` `.then(...)` you opened: the existing `Word.run(...).then(function () { …status… })` block must now be nested inside it. Match braces carefully — the `Word.run` promise chain that ends at the status `setStatus(msg, …)` (line 587) becomes the return value of the `fetchTagsIfSectioned().then` callback, and the outer `.catch` at line 589 stays at the `ensureAssets` level.

- [ ] **Step 6: Register `cite-classify.js` in the pane script list + bump asset version**

In `src/taskpane/taskpane.html`, add `"./cite-classify.js",` to the `srcs` array immediately after `"./cite-word.js",` (line 674) and before `"./cite-zotero.js"` (CiteClassify is consumed by cite-pane; it only needs to load before cite-pane.js, but placing it next to the other pure cite modules keeps the group tidy):

```js
          "./cite-word.js",
          "./cite-classify.js",
          "./cite-zotero.js",
```

Bump the asset version (line 15) so installed users re-fetch the changed pane JS:

```js
      window.ASHAAR_ASSET_VERSION = "20260718-cite-classify";
```

- [ ] **Step 7: Run the full suite (regression)**

Run: `npm test`
Expected: full suite PASS. (This proves the shared modules still load and no cross-file breakage was introduced; cite-pane itself is exercised manually in Task 6.)

- [ ] **Step 8: Sanity-check the pane parses in a browser dev server (optional but recommended)**

Run: `npm run dev-server` and load `https://localhost:3000/taskpane.html` in a browser. Open the console; confirm no `CiteClassify is not defined` / syntax errors and the Cite tab renders. Stop the server.

- [ ] **Step 9: Commit**

```bash
git add src/taskpane/cite-pane.js src/taskpane/taskpane.html
git commit -m "feat(cite): nested-section bibliography wiring — insert + refresh (SP-4)

-fatemi style => fetch Zotero tags, plan buckets, render per-bucket over a
filtered item map, insert with localized RTL-aware headings. Stock styles and
single-bucket cases render flat, byte-identical to before. Tag-fetch failure
degrades to flat.

Co-Authored-By: Claude Opus 4.8 <noreply@anthropic.com>
Claude-Session: https://claude.ai/code/session_01Vy1kwq3DuKd6bpgQGdN6Yw"
```

---

### Task 6: Manual Word verification checklist

**Files:**
- Create: `docs/superpowers/specs/2026-07-18-citation-source-classification-manual-checklist.md`

**Interfaces:**
- Consumes: the running add-in + a Zotero library with Better BibTeX.
- Produces: a checklist doc (matches the SP-3 checklist style at `docs/superpowers/specs/2026-07-17-citation-multilingual-variants-manual-checklist.md`).

- [ ] **Step 1: Write the checklist**

Create the file with these checks (fill in as literal checklist items):

```markdown
# SP-4 Source Classification — Manual Word Checklist

Prereqs: Zotero running with Better BibTeX; a few library items tagged with
`corpus:fatemi` and/or `class:secondary`; the rest untagged (default primary/non-Fatemi).

1. [ ] **Live tag fetch.** With a `-fatemi` style selected, add several items from Zotero
   spanning ≥2 buckets. Insert bibliography. Confirm it splits into headed subsections in the
   fixed order Primary·Fatemi → Primary·Other → Secondary·Fatemi → Secondary·Other, and empty
   buckets are skipped.
   - If headings do NOT appear: check the console — the BBT translator name may differ from
     `"BetterBibTeX JSON"`. Confirm the live name (Zotero → File → Export → translator list) and,
     if needed, update `buildTagsRequest` + `parseTagsResult` (+ their tests) or fall back to the
     Zotero local API path (see the design's feasibility note).
2. [ ] **Collapse rule.** Select items that all fall in ONE bucket (e.g. all untagged). Insert
   bibliography under a `-fatemi` style → confirm a single flat list with NO heading (identical
   to a stock-style bibliography).
3. [ ] **Stock style = flat.** Switch to Chicago (notes & bibliography) or APA. Insert
   bibliography → confirm one flat list, no headings, no tag fetch behavior change.
4. [ ] **Arabic RTL headings.** Set locale = ar, select a `-fatemi` style, insert a multi-bucket
   bibliography → confirm Arabic headings render bold, right-to-left, no tofu, and entries follow
   under each (Arabic titles upright, not italic).
5. [ ] **Refresh reproduces sections.** With a sectioned bibliography inserted, change nothing and
   click "Refresh citations" → confirm the bibliography CC re-renders with the same sections.
6. [ ] **Refresh collapses on style switch.** Switch from a `-fatemi` style to a stock style, click
   "Refresh citations" → confirm the bibliography collapses to one flat list in place.
7. [ ] **Tag-fetch failure degrades to flat.** Quit Zotero (or block the proxy), select a
   `-fatemi` style, insert bibliography → confirm it still inserts as a flat list (no crash, a
   status message is acceptable).
8. [ ] **Save/close/reopen.** Save the doc, close, reopen → confirm the sectioned bibliography
   text persists (it is static once inserted) and a subsequent Refresh still works.
```

- [ ] **Step 2: Commit**

```bash
git add docs/superpowers/specs/2026-07-18-citation-source-classification-manual-checklist.md
git commit -m "docs(cite): SP-4 source classification — manual Word checklist

Co-Authored-By: Claude Opus 4.8 <noreply@anthropic.com>
Claude-Session: https://claude.ai/code/session_01Vy1kwq3DuKd6bpgQGdN6Yw"
```

---

## Self-Review

**Spec coverage:**
- Classification model (two axes, prefixed tags, defaults, malformed handling) → Task 1 (`bucketForTags` + tests). ✅
- Tag retrieval via second `item.export` → Task 2 (`fetchTags`), with the live-confirm note carried from the spec. ✅
- Nested sections / fixed order / collapse rule → Task 1 (`planBibliographySections`) + Task 3 (assemblers). ✅
- Localized RTL-aware headings → Task 1 (`headingFor` en/ar) + Task 3 (OOXML bold-run path) + Task 5 (`renderBibliographyBody`). ✅
- Fatemi-style-as-toggle → Task 5 (`isFatemiStyle`). ✅
- Remove inert `[genre]` branch → Task 4. ✅
- SP-C refresh parity → Task 5 Step 5. ✅
- Degrade-to-flat on failure → Task 5 (`fetchTagsIfSectioned` `.catch`). ✅
- Bibliography-only (citations untouched) → no task modifies `insertCitation`/`renderCitationInto` logic. ✅
- Tests registered / asset version bumped → Task 1 Step 5, Task 5 Step 6. ✅
- Manual verification → Task 6. ✅

**Placeholder scan:** No TBD/TODO; every code step shows complete code. The one deferred decision (exact BBT translator name) is explicitly a live-confirm with a documented fallback and pinned by tests — not a plan placeholder.

**Type consistency:** `planBibliographySections` returns `{key, heading, citekeys}`; consumed as such in `renderBibliographyBody` and `renderBibliographyInto`'s default. `fetchTags` returns `{citekey:[tags]}`; consumed by `planBibliographySections`'s `tagsByCitekey`. `buildSectionedBibliography{Html,Ooxml}` consume `[{heading, html}]`, produced by `renderBibliographyBody`'s `rendered`. `bucketForTags` returns `{corpus, cls, key}` consistently (note: `cls`, not `class`, to avoid the reserved word — used consistently in Task 1). `buildEngine(styleFile, lang, itemsOverride)` — the new 3rd param is optional and all existing 2-arg calls stay valid. Consistent. ✅

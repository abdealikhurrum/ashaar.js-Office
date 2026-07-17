# SP-3 Multilingual Variant Model — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Render Arabic-original / romanized / both citation variants, sourced from CNE `cne-*` Extra data, with a one-time utility that migrates legacy Juris-M mlzsync data into CNE.

**Architecture:** A new pure UMD module `src/taskpane/cite-variants.js` parses `cne-*` variant data and normalizes it into citeproc's `multi` model (which `cite-engine.js` already consumes via `langPrefs`). The Cite pane gains a Variant selector (Original/Romanized/Both) threaded through `buildEngine`, persisted in a v2 citation tag, re-applied on Refresh. A separate dev-run script `scripts/migrate-mlzsync-to-cne.mjs` converts mlzsync→`cne-*` and PATCHes the Zotero library via the local API.

**Tech Stack:** Vanilla ES5/UMD JS (no build step), citeproc-js (CSL-M), Node `assert` tests, Office.js (Word), Zotero local API (`localhost:23119/api`).

## Global Constraints

- **No build step / no transpilation.** ES5-compatible syntax in `src/` (UMD modules, `var`, `function`). The migration script under `scripts/` is an `.mjs` ESM module (Node-only, like the other `scripts/*.mjs`) and may use modern JS.
- **UMD module pattern** for `src/taskpane/cite-variants.js` — exact factory shape from `cite-word.js` (Node `module.exports` + browser `root.CiteVariants`).
- **Pure functions only** in `cite-variants.js` — no Office.js, no DOM, no `fetch`; Node-`assert` testable.
- **Never edit `src/vendor/`** directly.
- **Every new test file** must be appended to the `test` script chain in `package.json`.
- **Bump `window.ASHAAR_ASSET_VERSION`** in `taskpane.html` when pane code ships (currently `"20260717-cite-refresh"`).
- **Internal transliteration tag = `"ar-Latn"`; translation tag = `"en"`.** All parsers normalize romanized content under `ar-Latn` and translated content under `en`, regardless of the source's label. This matches the existing `tests/fixtures/cite-multi.json` (which uses `ar-Latn`).
- **Citation tag versions:** current `AshaarCite:`/`AshaarBib:` payloads are `v:1`. This plan introduces `v:2` adding `variant`. v1 reads migrate to `variant:"orig"`.
- **Commit trailer** on every commit:
  ```
  Co-Authored-By: Claude Opus 4.8 <noreply@anthropic.com>
  Claude-Session: https://claude.ai/code/session_01Vy1kwq3DuKd6bpgQGdN6Yw
  ```

---

## Task 1: `cite-variants.js` — `parseCne(text)`

Parse `cne-*` lines out of a CSL-JSON `note` string (or a raw Extra string) into a normalized intermediate.

**Files:**
- Create: `src/taskpane/cite-variants.js`
- Test: `tests/cite-variants.test.js`

**Interfaces:**
- Produces: `CiteVariants.parseCne(text) -> { fields, creators } | null`
  - `fields`: `{ [cslField]: { [tag]: string } }` — e.g. `{ "title": { "ar-Latn": "Uyun al-Akhbar Vol. 4" } }`
  - `creators`: `{ [cslCreatorVar]: { [index]: { [tag]: { family?, given?, literal? } } } }` — e.g. `{ "author": { "0": { "ar-Latn": { literal: "al-Dai…" } } } }`
  - Returns `null` when no `cne-*` line is present.
- Tag mapping: variant suffix `romanized`→`ar-Latn`, `translated`→`en`, `original`→ skipped (real field holds it) for v1.
- Line grammar: `cne-<field>-<variant>: value` (simple field) OR `cne-<creatorType>-<index>-<part>-<variant>: value` (creator, `index` numeric, `part`∈{`last`,`first`}). `<field>`/`<creatorType>` may themselves contain hyphens (e.g. `container-title`, `container-author`); the numeric segment disambiguates a creator line.
- Bidi control chars (U+202B, U+202C, U+200E, U+200F, U+202A, U+202D) stripped from values.

- [ ] **Step 1: Write the failing test**

Append to `tests/cite-variants.test.js`:

```javascript
"use strict";
const assert = require("assert");
const CV = require("../src/taskpane/cite-variants");

// --- parseCne ---
const VARIANT_TAG = "ar-Latn";

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
```

- [ ] **Step 2: Run test to verify it fails**

Run: `node tests/cite-variants.test.js`
Expected: FAIL — `Cannot find module '../src/taskpane/cite-variants'`.

- [ ] **Step 3: Write minimal implementation**

Create `src/taskpane/cite-variants.js`:

```javascript
(function (root, factory) {
  if (typeof module !== "undefined" && module.exports) {
    module.exports = factory();
  } else {
    root.CiteVariants = factory();
  }
}(typeof globalThis !== "undefined" ? globalThis : this, function () {
  "use strict";

  var BIDI = /[‎‏‪‫‬‭‮]/g;
  var VARIANT_TAGS = { romanized: "ar-Latn", translated: "en" }; // 'original' -> skip (real field)

  function stripBidi(s) { return String(s == null ? "" : s).replace(BIDI, "").trim(); }

  // Parse one "cne-<rest>" key (prefix already removed) + value into the
  // normalized intermediate, mutating `out`. Ignores unknown-variant lines.
  function addCneKey(out, rest, value) {
    var segs = rest.split("-");
    if (segs.length < 2) { return; }
    var variant = segs[segs.length - 1];
    var tag = VARIANT_TAGS[variant];
    if (!tag) { return; } // unknown/original variant -> ignore
    var body = segs.slice(0, segs.length - 1); // field or creator body

    // creator line: contains a numeric segment followed by last|first
    var numIdx = -1;
    for (var i = 0; i < body.length; i++) { if (/^\d+$/.test(body[i])) { numIdx = i; break; } }
    if (numIdx !== -1 && numIdx + 1 < body.length &&
        (body[numIdx + 1] === "last" || body[numIdx + 1] === "first")) {
      var creatorType = body.slice(0, numIdx).join("-");
      var index = body[numIdx];
      var part = body[numIdx + 1];
      out.creators[creatorType] = out.creators[creatorType] || {};
      out.creators[creatorType][index] = out.creators[creatorType][index] || {};
      out.creators[creatorType][index][tag] = out.creators[creatorType][index][tag] || {};
      out.creators[creatorType][index][tag][part === "last" ? "family" : "given"] = stripBidi(value);
      return;
    }

    // simple field
    var field = body.join("-");
    out.fields[field] = out.fields[field] || {};
    out.fields[field][tag] = stripBidi(value);
  }

  function parseCne(text) {
    var s = String(text || "");
    if (s.indexOf("cne-") === -1) { return null; }
    var out = { fields: {}, creators: {} };
    var lines = s.split(/\r?\n/);
    var seen = false;
    for (var i = 0; i < lines.length; i++) {
      var m = /^\s*cne-([^:]+):\s*([\s\S]*)$/.exec(lines[i]);
      if (!m) { continue; }
      var before = JSON.stringify(out);
      addCneKey(out, m[1].trim(), m[2]);
      if (JSON.stringify(out) !== before) { seen = true; }
    }
    return seen ? out : null;
  }

  return { parseCne: parseCne, stripBidi: stripBidi };
}));
```

- [ ] **Step 4: Run test to verify it passes**

Run: `node tests/cite-variants.test.js`
Expected: PASS — `cite-variants parseCne test passed`.

- [ ] **Step 5: Commit**

```bash
git add src/taskpane/cite-variants.js tests/cite-variants.test.js
git commit -m "feat(cite): cite-variants parseCne — cne-* Extra -> normalized variants

Co-Authored-By: Claude Opus 4.8 <noreply@anthropic.com>
Claude-Session: https://claude.ai/code/session_01Vy1kwq3DuKd6bpgQGdN6Yw"
```

---

## Task 2: `cite-variants.js` — `applyVariantsToItem` + `enrichItemMap`

Bake the normalized `cne-*` variants into a CSL-JSON item's `multi` model.

**Files:**
- Modify: `src/taskpane/cite-variants.js`
- Test: `tests/cite-variants.test.js`

**Interfaces:**
- Consumes: `parseCne` (Task 1). Reads variants from `item.note` (CSL-JSON surfaces Extra as `note`).
- Produces:
  - `CiteVariants.applyVariantsToItem(item) -> item2` — returns a **new** object (shallow-cloned enough to attach `multi`); when no `cne-*` present, returns the input unchanged. Sets `item2.multi = { main, _keys }` and, per creator variable/index found, `item2[creatorVar][index].multi = { _key: {...} }`.
  - `CiteVariants.enrichItemMap(items) -> items2` — maps `applyVariantsToItem` over an `{id:item}` map into a new map.
- Creator matching: `creators.author["0"]` maps to `item.author[0]`. If the creator array/index is absent, that variant is skipped (no throw).

- [ ] **Step 1: Write the failing test**

Append to `tests/cite-variants.test.js`:

```javascript
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
```

- [ ] **Step 2: Run test to verify it fails**

Run: `node tests/cite-variants.test.js`
Expected: FAIL — `CV.applyVariantsToItem is not a function`.

- [ ] **Step 3: Write minimal implementation**

In `src/taskpane/cite-variants.js`, add before the `return`:

```javascript
  // Shallow clone + attach multi models from parsed cne-* variants.
  function applyVariantsToItem(item) {
    if (!item || typeof item !== "object") { return item; }
    var parsed = parseCne(item.note);
    if (!parsed) { return item; }

    var out = {};
    var k;
    for (k in item) { if (Object.prototype.hasOwnProperty.call(item, k)) { out[k] = item[k]; } }

    // fields
    var keys = {};
    var main = {};
    for (var f in parsed.fields) {
      if (Object.prototype.hasOwnProperty.call(parsed.fields, f)) {
        keys[f] = parsed.fields[f];
        main[f] = item.language || "ar";
      }
    }
    if (Object.keys(keys).length) { out.multi = { main: main, _keys: keys }; }

    // creators — clone the target creator array + entry before attaching multi
    for (var cv in parsed.creators) {
      if (!Object.prototype.hasOwnProperty.call(parsed.creators, cv)) { continue; }
      if (!Array.isArray(out[cv])) { continue; }
      out[cv] = out[cv].slice();
      var byIdx = parsed.creators[cv];
      for (var idx in byIdx) {
        if (!Object.prototype.hasOwnProperty.call(byIdx, idx)) { continue; }
        var i = parseInt(idx, 10);
        if (!out[cv][i]) { continue; }
        var c = {};
        for (var ck in out[cv][i]) { if (Object.prototype.hasOwnProperty.call(out[cv][i], ck)) { c[ck] = out[cv][i][ck]; } }
        c.multi = { main: item.language || "ar", _key: byIdx[idx] };
        out[cv][i] = c;
      }
    }
    return out;
  }

  function enrichItemMap(items) {
    var out = {};
    for (var id in items) {
      if (Object.prototype.hasOwnProperty.call(items, id)) { out[id] = applyVariantsToItem(items[id]); }
    }
    return out;
  }
```

And extend the returned object:

```javascript
  return {
    parseCne: parseCne,
    stripBidi: stripBidi,
    applyVariantsToItem: applyVariantsToItem,
    enrichItemMap: enrichItemMap
  };
```

- [ ] **Step 4: Run test to verify it passes**

Run: `node tests/cite-variants.test.js`
Expected: PASS — `cite-variants applyVariantsToItem test passed`.

- [ ] **Step 5: Commit**

```bash
git add src/taskpane/cite-variants.js tests/cite-variants.test.js
git commit -m "feat(cite): applyVariantsToItem + enrichItemMap — bake cne-* into multi model

Co-Authored-By: Claude Opus 4.8 <noreply@anthropic.com>
Claude-Session: https://claude.ai/code/session_01Vy1kwq3DuKd6bpgQGdN6Yw"
```

---

## Task 3: `cite-variants.js` — `variantToLangPrefs`

Map the pane's variant choice to a citeproc `langPrefs` object.

**Files:**
- Modify: `src/taskpane/cite-variants.js`
- Test: `tests/cite-variants.test.js`

**Interfaces:**
- Produces: `CiteVariants.variantToLangPrefs(variant) -> langPrefs | null`
  - `"orig"` → `null` (no override; real fields render — current behavior).
  - `"translit"` → all segment groups `["translit"]` + `translit:["ar-Latn"]` + `translat:["en"]`.
  - `"both"` → all segment groups `["orig","translit"]` + same tag registration.
- Segment groups set: `persons, institutions, titles, journals, publishers, places, number, "title-short"` (citeproc-js CSL-M `langPrefs` groups).
- The shape is exactly what `cite-engine.js` `build({langPrefs})` already consumes (see `src/taskpane/cite-engine.js:39-50`).

- [ ] **Step 1: Write the failing test**

Append to `tests/cite-variants.test.js`:

```javascript
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
```

- [ ] **Step 2: Run test to verify it fails**

Run: `node tests/cite-variants.test.js`
Expected: FAIL — `CV.variantToLangPrefs is not a function`.

- [ ] **Step 3: Write minimal implementation**

In `src/taskpane/cite-variants.js`, add before the `return`:

```javascript
  var SEGMENTS = ["persons", "institutions", "titles", "journals", "publishers", "places", "number", "title-short"];

  function variantToLangPrefs(variant) {
    var slots;
    if (variant === "translit") { slots = ["translit"]; }
    else if (variant === "both") { slots = ["orig", "translit"]; }
    else { return null; } // orig / unknown -> no override
    var lp = { translit: ["ar-Latn"], translat: ["en"] };
    for (var i = 0; i < SEGMENTS.length; i++) { lp[SEGMENTS[i]] = slots.slice(); }
    return lp;
  }
```

Add `variantToLangPrefs: variantToLangPrefs` to the returned object.

- [ ] **Step 4: Run test to verify it passes**

Run: `node tests/cite-variants.test.js`
Expected: PASS — `cite-variants variantToLangPrefs test passed`.

- [ ] **Step 5: Commit**

```bash
git add src/taskpane/cite-variants.js tests/cite-variants.test.js
git commit -m "feat(cite): variantToLangPrefs — variant choice -> citeproc langPrefs

Co-Authored-By: Claude Opus 4.8 <noreply@anthropic.com>
Claude-Session: https://claude.ai/code/session_01Vy1kwq3DuKd6bpgQGdN6Yw"
```

---

## Task 4: `cite-variants.js` — `parseMlzsync` + `mlzsyncToCneLines` (migration converter)

Pure converter used **only** by the migration utility (Task 8). Inflates the legacy mlzsync blob and emits `cne-*` lines.

**Files:**
- Modify: `src/taskpane/cite-variants.js`
- Test: `tests/cite-variants.test.js`

**Interfaces:**
- Produces:
  - `CiteVariants.parseMlzsync(text) -> { fields, creators } | null` — inflates `mlzsync1:NNNN{…}`. `fields`: `{ [cslField]: { [srcTag]: string } }`. `creators`: `{ [flatIndex]: { [srcTag]: { family?, given?, literal? } } }` (keyed by the flat Zotero creator index, since mlzsync has no creator type). `null` if no `mlzsync1:` prefix or on parse error.
  - `CiteVariants.mlzsyncToCneLines(parsed, creators) -> string[]` — `creators` is the ordered native Zotero `data.creators` array (`[{creatorType, ...}]`) used to resolve the flat index to a `creatorType` + within-type index. Emits `cne-<field>-romanized: v` and `cne-<creatorType>-<i>-<part>-romanized: v`. mlzsync `firstName`/`lastName` map to `first`/`last`; a literal (fieldMode 1, only `lastName`) emits just the `-last-` line. Idempotent (stable output).
- mlzsync source tag `en` holds transliteration → emitted as `romanized` (`ar-Latn` at read time; `romanized` at write time).

- [ ] **Step 1: Write the failing test**

Append to `tests/cite-variants.test.js`:

```javascript
// --- parseMlzsync + mlzsyncToCneLines (migration) ---
const fs = require("fs");
const path = require("path");
const realArr = JSON.parse(fs.readFileSync(path.join(__dirname, "fixtures", "cite-mlzsync-real.json"), "utf8"));
const real = realArr[0];

const pm = CV.parseMlzsync(real.note);
assert.deepStrictEqual(pm.fields.title, { en: "Uyun al-Akhbar Vol. 4" }, "mlzsync title (bidi stripped)");
assert.deepStrictEqual(
  pm.creators["0"].en,
  { family: "al-Dai al-Ajal Syedna Idris Imaduddin RA" },
  "mlzsync creator literal -> family (lastName)"
);

// no prefix -> null
assert.strictEqual(CV.parseMlzsync("cne-title-romanized: x"), null, "no mlzsync prefix => null");
assert.strictEqual(CV.parseMlzsync("mlzsync1:9999{bad json"), null, "malformed => null (no throw)");

// converter: needs native creators for type resolution
const nativeCreators = [{ creatorType: "author", name: "…" }];
const lines = CV.mlzsyncToCneLines(pm, nativeCreators);
assert.ok(lines.indexOf("cne-title-romanized: Uyun al-Akhbar Vol. 4") !== -1, "emits title line");
assert.ok(lines.indexOf("cne-author-0-last-romanized: al-Dai al-Ajal Syedna Idris Imaduddin RA") !== -1, "emits author line");
// idempotent
assert.deepStrictEqual(CV.mlzsyncToCneLines(pm, nativeCreators), lines, "stable output");
console.log("cite-variants mlzsync test passed");
```

- [ ] **Step 2: Create the real-item fixture**

Create `tests/fixtures/cite-mlzsync-real.json` by copying the user's exported item verbatim:

```bash
cp "/Users/abdealikhurrum/Documents/Exported Items.json" tests/fixtures/cite-mlzsync-real.json
```

Verify it is a JSON array whose `[0].note` starts with `mlzsync1:` (Read the file).

- [ ] **Step 3: Run test to verify it fails**

Run: `node tests/cite-variants.test.js`
Expected: FAIL — `CV.parseMlzsync is not a function`.

- [ ] **Step 4: Write minimal implementation**

In `src/taskpane/cite-variants.js`, add before the `return`:

```javascript
  function parseMlzsync(text) {
    var s = String(text || "");
    var at = s.indexOf("mlzsync1:");
    if (at === -1) { return null; }
    var rest = s.slice(at + "mlzsync1:".length);
    // 4-digit zero-padded length prefix, then JSON of that length.
    var m = /^(\d{4})/.exec(rest);
    if (!m) { return null; }
    var len = parseInt(m[1], 10);
    var json = rest.slice(4, 4 + len);
    var blob;
    try { blob = JSON.parse(json); } catch (e) {
      try { blob = JSON.parse(rest.slice(4)); } catch (e2) { return null; }
    }
    var out = { fields: {}, creators: {} };
    var mf = (blob && blob.multifields && blob.multifields._keys) || {};
    for (var f in mf) {
      if (!Object.prototype.hasOwnProperty.call(mf, f)) { continue; }
      out.fields[f] = {};
      for (var tag in mf[f]) {
        if (Object.prototype.hasOwnProperty.call(mf[f], tag)) { out.fields[f][tag] = stripBidi(mf[f][tag]); }
      }
    }
    var mc = (blob && blob.multicreators) || {};
    for (var idx in mc) {
      if (!Object.prototype.hasOwnProperty.call(mc, idx)) { continue; }
      var keyObj = mc[idx]._key || {};
      out.creators[idx] = {};
      for (var t in keyObj) {
        if (!Object.prototype.hasOwnProperty.call(keyObj, t)) { continue; }
        var nm = keyObj[t];
        var v = {};
        if (nm.lastName) { v.family = stripBidi(nm.lastName); }
        if (nm.firstName) { v.given = stripBidi(nm.firstName); }
        out.creators[idx][t] = v;
      }
    }
    return out;
  }

  function mlzsyncToCneLines(parsed, creators) {
    var lines = [];
    if (!parsed) { return lines; }
    // fields: any source tag -> romanized (mlzsync 'en' holds transliteration)
    var fnames = Object.keys(parsed.fields).sort();
    fnames.forEach(function (f) {
      var byTag = parsed.fields[f];
      var tags = Object.keys(byTag).sort();
      if (tags.length) { lines.push("cne-" + f + "-romanized: " + byTag[tags[0]]); }
    });
    // creators: resolve flat index -> creatorType + within-type index
    var typeCount = {};
    (creators || []).forEach(function (c, flat) {
      var type = (c && c.creatorType) || "author";
      var within = typeCount[type] || 0;
      typeCount[type] = within + 1;
      var byTag = parsed.creators[String(flat)];
      if (!byTag) { return; }
      var tags = Object.keys(byTag).sort();
      if (!tags.length) { return; }
      var nm = byTag[tags[0]];
      if (nm.family) { lines.push("cne-" + type + "-" + within + "-last-romanized: " + nm.family); }
      if (nm.given) { lines.push("cne-" + type + "-" + within + "-first-romanized: " + nm.given); }
    });
    return lines;
  }
```

Add `parseMlzsync: parseMlzsync, mlzsyncToCneLines: mlzsyncToCneLines` to the returned object.

- [ ] **Step 5: Run test to verify it passes**

Run: `node tests/cite-variants.test.js`
Expected: PASS — `cite-variants mlzsync test passed`.

- [ ] **Step 6: Commit**

```bash
git add src/taskpane/cite-variants.js tests/cite-variants.test.js tests/fixtures/cite-mlzsync-real.json
git commit -m "feat(cite): parseMlzsync + mlzsyncToCneLines — migration converter (pure)

Co-Authored-By: Claude Opus 4.8 <noreply@anthropic.com>
Claude-Session: https://claude.ai/code/session_01Vy1kwq3DuKd6bpgQGdN6Yw"
```

---

## Task 5: Engine test — variant policy round-trip (ar + en)

Prove the enriched item + `variantToLangPrefs` render the selected variant through the real engine.

**Files:**
- Modify: `tests/cite-engine.test.js`

**Interfaces:**
- Consumes: `CiteVariants.applyVariantsToItem`, `CiteVariants.variantToLangPrefs` (Tasks 2–3); `CiteEngine.build` (existing).

- [ ] **Step 1: Write the failing test**

Append to `tests/cite-engine.test.js`:

```javascript
// --- SP-3: variant policy end-to-end ---
const CV = require("../src/taskpane/cite-variants");
const cneItem = CV.applyVariantsToItem({
  id: "cne-1", type: "book", language: "ar",
  title: "دعائم الإسلام",
  author: [{ family: "النعمان", given: "القاضي" }],
  issued: { "date-parts": [[1951]] },
  note: "cne-title-romanized: Daʿāʾim al-Islām\ncne-author-0-last-romanized: al-Nuʿmān\ncne-author-0-first-romanized: al-Qāḍī"
});
const cneMap = { "cne-1": cneItem };

// translit policy -> romanized renders
const tEngine = CiteEngine.build({
  styleXml: read("csl-styles/chicago-notes-bibliography.csl"),
  locales, items: cneMap, lang: "en-US",
  langPrefs: CV.variantToLangPrefs("translit")
});
const tBib = tEngine.bibliography();
assert.match(tBib, /Nuʿm/, "cne translit: author romanization renders");
assert.match(tBib, /Islām/, "cne translit: title romanization renders");

// orig policy -> Arabic renders (no romanization)
const oEngine = CiteEngine.build({
  styleXml: read("csl-styles/chicago-notes-bibliography.csl"),
  locales, items: cneMap, lang: "ar",
  langPrefs: CV.variantToLangPrefs("orig") // null
});
const oBib = oEngine.bibliography();
assert.match(oBib, /دعائم الإسلام/, "cne orig: Arabic title renders");
assert.ok(oBib.indexOf("Islām") === -1, "cne orig: no romanization");
console.log("cite-engine (cne variant policy) test passed");
```

- [ ] **Step 2: Run test to verify it fails or passes**

Run: `node tests/cite-engine.test.js`
Expected: PASS (engine already consumes langPrefs; this is the integration guard). If it FAILS on `Nuʿm`, inspect the `multi` shape produced by `applyVariantsToItem` against `tests/fixtures/cite-multi.json` (the known-good shape) and reconcile.

- [ ] **Step 3: Commit**

```bash
git add tests/cite-engine.test.js
git commit -m "test(cite): variant policy end-to-end (cne item -> translit/orig render)

Co-Authored-By: Claude Opus 4.8 <noreply@anthropic.com>
Claude-Session: https://claude.ai/code/session_01Vy1kwq3DuKd6bpgQGdN6Yw"
```

---

## Task 6: Citation tag v2 — `variant` field + migration

Add `variant` to the citation/bibliography tag payloads, migrating v1 reads to `"orig"`.

**Files:**
- Modify: `src/taskpane/cite-word.js:108-132`
- Test: `tests/cite-word.test.js`

**Interfaces:**
- Modifies: `buildCitationTag(o)` — payload gains `v:2` + `variant: o.variant || "orig"`.
  `parseCitationTag(tag)` — returns parsed obj with `variant` (missing → `"orig"`).
  `buildBibliographyTag(o)` — gains `v:2` + `variant`.
- Back-compat: existing v1 tags (no `variant`) parse with `variant:"orig"`.

- [ ] **Step 1: Write the failing test**

Append to `tests/cite-word.test.js`:

```javascript
// --- SP-3: tag v2 variant field ---
const v2 = CiteWord.buildCitationTag({ style: "s", locale: "en-US", variant: "translit", items: [{ id: "a" }] });
const pv2 = CiteWord.parseCitationTag(v2);
assert.strictEqual(pv2.v, 2, "new tags are v2");
assert.strictEqual(pv2.variant, "translit", "variant round-trips");

// default when omitted
const vDef = CiteWord.parseCitationTag(CiteWord.buildCitationTag({ style: "s", locale: "en-US", items: [{ id: "a" }] }));
assert.strictEqual(vDef.variant, "orig", "omitted variant defaults to orig");

// v1 back-compat: hand-build a v1 payload (no variant) and confirm it reads as orig
const b64 = (typeof Buffer !== "undefined")
  ? Buffer.from(JSON.stringify({ v: 1, style: "s", locale: "en-US", keys: [{ id: "a", locator: null, label: null }] })).toString("base64")
  : null;
const v1parsed = CiteWord.parseCitationTag("AshaarCite:" + b64);
assert.strictEqual(v1parsed.variant, "orig", "v1 tag migrates to orig");
assert.strictEqual(v1parsed.keys[0].id, "a", "v1 keys still parse");

// bibliography tag carries variant too
const bibTag = CiteWord.buildBibliographyTag({ style: "s", locale: "ar", variant: "both" });
assert.strictEqual(bibTag.indexOf("AshaarBib:"), 0, "bib tag prefix");
console.log("cite-word tag v2 test passed");
```

- [ ] **Step 2: Run test to verify it fails**

Run: `node tests/cite-word.test.js`
Expected: FAIL — `pv2.v` is `1`, `pv2.variant` is `undefined`.

- [ ] **Step 3: Write minimal implementation**

In `src/taskpane/cite-word.js`, replace `buildCitationTag` (lines ~108-118):

```javascript
  function buildCitationTag(o) {
    var payload = {
      v: 2,
      style: o.style,
      locale: o.locale,
      variant: o.variant || "orig",
      keys: (o.items || []).map(function (i) {
        return { id: i.id, locator: i.locator || null, label: i.label || null };
      })
    };
    return "AshaarCite:" + b64encode(JSON.stringify(payload));
  }
```

Replace `parseCitationTag` to default `variant`:

```javascript
  function parseCitationTag(tag) {
    var s = String(tag || "");
    if (s.indexOf("AshaarCite:") !== 0) { return null; }
    try {
      var obj = JSON.parse(b64decode(s.slice("AshaarCite:".length)));
      if (!obj || !Array.isArray(obj.keys)) { return null; }
      if (!obj.variant) { obj.variant = "orig"; } // v1 migration
      return obj;
    } catch (e) { return null; }
  }
```

Replace `buildBibliographyTag`:

```javascript
  function buildBibliographyTag(o) {
    return "AshaarBib:" + b64encode(JSON.stringify({ v: 2, style: o.style, locale: o.locale, variant: o.variant || "orig" }));
  }
```

- [ ] **Step 4: Run test to verify it passes**

Run: `node tests/cite-word.test.js`
Expected: PASS — `cite-word tag v2 test passed`.

- [ ] **Step 5: Commit**

```bash
git add src/taskpane/cite-word.js tests/cite-word.test.js
git commit -m "feat(cite): citation tag v2 — variant field + v1 read migration

Co-Authored-By: Claude Opus 4.8 <noreply@anthropic.com>
Claude-Session: https://claude.ai/code/session_01Vy1kwq3DuKd6bpgQGdN6Yw"
```

---

## Task 7: Pane wiring — Variant selector + enrich item map + thread langPrefs + tag it

Wire the feature into the Cite pane (DOM + non-unit-tested integration; verified via the manual checklist).

**Files:**
- Modify: `src/taskpane/taskpane.html` (add `#cite-variant` select; load `cite-variants.js`; bump asset version)
- Modify: `src/taskpane/cite-pane.js` (`currentVariant`, `buildEngine` langPrefs, `enrichItemMap` at load + Zotero, pass `variant` into tag builders, refresh)

**Interfaces:**
- Consumes: `CiteVariants.enrichItemMap`, `CiteVariants.variantToLangPrefs` (window global `CiteVariants`).
- `buildEngine(styleFile, lang)` internally adds `langPrefs: CiteVariants.variantToLangPrefs(currentVariant())`.
- All `CiteWord.buildCitationTag(...)` / `buildBibliographyTag(...)` call sites pass `variant: currentVariant()`.

- [ ] **Step 1: Add the HTML select + script include + asset bump**

In `src/taskpane/taskpane.html`, after the Locale `.field` block (line ~580), insert:

```html
          <div class="field">
            <label for="cite-variant">Variant</label>
            <select id="cite-variant">
              <option value="orig" selected>Original (ar)</option>
              <option value="translit">Romanized</option>
              <option value="both">Both (orig + romanized)</option>
            </select>
          </div>
```

Add the module script include alongside the other cite module includes (search for `cite-word.js` `<script>` and add after it):

```html
    <script src="cite-variants.js"></script>
```

Bump the asset version:

```html
      window.ASHAAR_ASSET_VERSION = "20260717-cite-variants";
```

- [ ] **Step 2: Add `currentVariant()` + thread into `buildEngine`**

In `src/taskpane/cite-pane.js`, after `currentForm()` (line ~97):

```javascript
  function currentVariant() { return (byId("cite-variant") || {}).value || "orig"; }
```

Modify `buildEngine` (lines ~101-108) to include langPrefs:

```javascript
  function buildEngine(styleFile, lang) {
    return CiteEngine.build({
      styleXml: cache.styles[styleFile],
      locales: cache.locales,
      items: cache.items,
      lang: lang,
      langPrefs: (typeof CiteVariants !== "undefined")
        ? CiteVariants.variantToLangPrefs(currentVariant())
        : null
    });
  }
```

- [ ] **Step 3: Enrich the item map at both load sites**

In `ensureAssets` fixture load (line ~84), wrap the parsed items:

```javascript
        return fetchText("fixtures/cite-sample.json").then(function (txt) {
          var raw = JSON.parse(txt);
          cache.items = (typeof CiteVariants !== "undefined") ? CiteVariants.enrichItemMap(raw) : raw;
        });
```

In the CiteStore restore branch (line ~83), enrich the saved set too:

```javascript
        if (saved && Object.keys(saved).length) {
          cache.items = (typeof CiteVariants !== "undefined") ? CiteVariants.enrichItemMap(saved) : saved;
          return;
        }
```

In `addFromZotero` (lines ~613-620), enrich each fetched item as it is merged:

```javascript
      return CiteZotero.fetchCslJson(citekeys).then(function (items) {
        if (!cache.items) { cache.items = {}; }
        var enriched = (typeof CiteVariants !== "undefined") ? CiteVariants.enrichItemMap(items) : items;
        Object.keys(enriched).forEach(function (id) {
          cache.items[id] = enriched[id];
        });
```

(Adjust the existing loop body to use `enriched` and keep the rest of the checkbox/selection logic unchanged.)

- [ ] **Step 4: Pass `variant` into every tag builder + refresh**

Find each `CiteWord.buildCitationTag({` and `CiteWord.buildBibliographyTag({` call in `cite-pane.js` and add `variant: currentVariant(),` to the object literal (search: `grep -n "buildCitationTag\|buildBibliographyTag" src/taskpane/cite-pane.js`). The `refreshCitations()` path already reads `currentStyleFile()`/`currentLang()`; ensure its re-render builds tags with `variant: currentVariant()` too (the engine picks up the variant automatically via `buildEngine`).

- [ ] **Step 5: Verify unit tests still pass (no regressions)**

Run: `npm test`
Expected: all suites PASS (pane changes are DOM-only; unit suites unaffected).

- [ ] **Step 6: Commit**

```bash
git add src/taskpane/taskpane.html src/taskpane/cite-pane.js
git commit -m "feat(cite): Variant selector wired — enrich item map, langPrefs, v2 tags

Co-Authored-By: Claude Opus 4.8 <noreply@anthropic.com>
Claude-Session: https://claude.ai/code/session_01Vy1kwq3DuKd6bpgQGdN6Yw"
```

---

## Task 8: Migration utility — `scripts/migrate-mlzsync-to-cne.mjs`

One-time dev-run converter: read Zotero items via the local API, convert mlzsync→`cne-*`, PATCH `extra`.

**Files:**
- Create: `scripts/migrate-mlzsync-to-cne.mjs`
- Modify: `package.json` (add `migrate:cne` script)

**Interfaces:**
- Consumes: `CiteVariants.parseMlzsync`, `CiteVariants.mlzsyncToCneLines` (require the UMD module from Node).
- CLI: `node scripts/migrate-mlzsync-to-cne.mjs [--write] [--force] [--strip-mlzsync] [--base=http://localhost:23119]`
  - default (no `--write`): dry-run, prints per-item diff, writes nothing.
  - `--write`: PATCH items; backs up affected items first.
  - `--force`: overwrite existing `cne-*` lines (default skips already-migrated items).
  - `--strip-mlzsync`: remove the `mlzsync1:…` block from `extra` after adding `cne-*` (default keeps it).

- [ ] **Step 1: First gate — verify local-API read + write against the running Zotero**

Run (with Zotero open):

```bash
curl -s "http://localhost:23119/api/users/0/items?limit=1&format=json" | head -c 400
```

Expected: a JSON array of item objects, each with `key`, `version`, and `data.extra`. If this 404s or returns HTML, the local API is not exposed — STOP and report; the migration falls back to the export/paste path noted in the spec. Record the exact base URL and library path (`users/0` for the local default library) that works.

- [ ] **Step 2: Write the script**

Create `scripts/migrate-mlzsync-to-cne.mjs`:

```javascript
#!/usr/bin/env node
// One-time migration: Juris-M mlzsync (Extra) -> CNE cne-* (Extra) via the Zotero local API.
import { createRequire } from "module";
import { writeFileSync, mkdirSync } from "fs";
import { fileURLToPath } from "url";
import { dirname, join } from "path";

const require = createRequire(import.meta.url);
const CV = require("../src/taskpane/cite-variants.js");

const args = process.argv.slice(2);
const has = (f) => args.includes(f);
const WRITE = has("--write");
const FORCE = has("--force");
const STRIP = has("--strip-mlzsync");
const BASE = (args.find((a) => a.startsWith("--base=")) || "--base=http://localhost:23119").split("=")[1];
const LIB = "users/0";

async function getAllItems() {
  const out = [];
  let start = 0;
  for (;;) {
    const res = await fetch(`${BASE}/api/${LIB}/items?limit=100&start=${start}&format=json`);
    if (!res.ok) { throw new Error(`GET items failed: ${res.status}`); }
    const batch = await res.json();
    if (!batch.length) { break; }
    out.push(...batch);
    start += batch.length;
    if (batch.length < 100) { break; }
  }
  return out;
}

// Merge cne-* lines into an extra string. Skips lines whose key already exists
// unless FORCE. Optionally strips the mlzsync block.
function mergeExtra(extra, cneLines) {
  let lines = String(extra || "").split(/\r?\n/);
  if (STRIP) { lines = lines.filter((l) => l.indexOf("mlzsync1:") === -1); }
  const existingKeys = new Set(
    lines.map((l) => (/^\s*(cne-[^:]+):/.exec(l) || [])[1]).filter(Boolean)
  );
  const additions = cneLines.filter((l) => {
    const k = (/^\s*(cne-[^:]+):/.exec(l) || [])[1];
    return FORCE || !existingKeys.has(k);
  });
  if (!additions.length && !STRIP) { return null; } // nothing to do
  return lines.concat(additions).filter((l) => l.length).join("\n");
}

async function patchExtra(item, extra) {
  const res = await fetch(`${BASE}/api/${LIB}/items/${item.key}`, {
    method: "PATCH",
    headers: { "Content-Type": "application/json", "If-Unmodified-Since-Version": String(item.version) },
    body: JSON.stringify({ extra })
  });
  if (res.status === 412) { throw new Error("412 version conflict"); }
  if (!res.ok && res.status !== 204) { throw new Error(`PATCH failed: ${res.status}`); }
}

(async () => {
  const items = await getAllItems();
  const affected = [];
  let converted = 0, skipped = 0, failed = 0;

  for (const it of items) {
    const extra = (it.data && it.data.extra) || "";
    const parsed = CV.parseMlzsync(extra);
    if (!parsed) { continue; }
    const creators = (it.data && it.data.creators) || [];
    const cneLines = CV.mlzsyncToCneLines(parsed, creators);
    if (!cneLines.length) { continue; }
    const newExtra = mergeExtra(extra, cneLines);
    if (newExtra === null) { skipped++; continue; }

    console.log(`\n# ${it.key}  ${(it.data && it.data.title) || ""}`);
    cneLines.forEach((l) => console.log("  + " + l));
    affected.push({ key: it.key, before: extra, after: newExtra, item: it });
  }

  console.log(`\n${affected.length} item(s) to convert, ${skipped} already migrated.`);

  if (!WRITE) {
    console.log("\nDRY RUN — no changes written. Re-run with --write to apply.");
    return;
  }

  // Backup before any write.
  const __dirname = dirname(fileURLToPath(import.meta.url));
  const backupDir = join(__dirname, "..", "scratch");
  try { mkdirSync(backupDir, { recursive: true }); } catch (e) {}
  const backupPath = join(backupDir, "mlzsync-backup.json");
  writeFileSync(backupPath, JSON.stringify(affected.map((a) => a.item), null, 2));
  console.log(`Backup written: ${backupPath}`);

  for (const a of affected) {
    try {
      await patchExtra(a.item, a.after);
      converted++;
      console.log(`  ✓ ${a.key}`);
    } catch (e) {
      failed++;
      console.error(`  ✗ ${a.key}: ${e.message}`);
    }
  }
  console.log(`\nDone. converted=${converted} skipped=${skipped} failed=${failed}`);
})().catch((e) => { console.error(e); process.exit(1); });
```

- [ ] **Step 3: Add the npm script**

In `package.json` `scripts`, after `"update:citeproc"`:

```json
    "migrate:cne": "node scripts/migrate-mlzsync-to-cne.mjs",
```

- [ ] **Step 4: Dry-run against the live library**

Run (Zotero open): `npm run migrate:cne`
Expected: prints a per-item `+ cne-*` diff for each mlzsync item and `N item(s) to convert`, writing nothing. Inspect the diff for the known book item (`Uyun al-Akhbar`): it must show `cne-title-romanized: Uyun al-Akhbar Vol. 4` and `cne-author-0-last-romanized: al-Dai al-Ajal Syedna Idris Imaduddin RA`.

- [ ] **Step 5: Commit**

```bash
git add scripts/migrate-mlzsync-to-cne.mjs package.json
git commit -m "feat(cite): migrate-mlzsync-to-cne — one-time mlzsync -> cne-* via Zotero local API

Co-Authored-By: Claude Opus 4.8 <noreply@anthropic.com>
Claude-Session: https://claude.ai/code/session_01Vy1kwq3DuKd6bpgQGdN6Yw"
```

---

## Task 9: Register test suite + manual checklist

**Files:**
- Modify: `package.json` (add `cite-variants.test.js` to the `test` chain)
- Create: `docs/superpowers/specs/2026-07-17-citation-multilingual-variants-manual-checklist.md`

- [ ] **Step 1: Add the new test to the chain**

In `package.json`, append to the `test` script after `node tests/cite-word.test.js`:

```
 && node tests/cite-variants.test.js
```

- [ ] **Step 2: Run the full suite**

Run: `npm test`
Expected: all suites PASS, including `cite-variants …` lines.

- [ ] **Step 3: Write the manual checklist**

Create `docs/superpowers/specs/2026-07-17-citation-multilingual-variants-manual-checklist.md` with these live-Word / live-Zotero checks:

```markdown
# SP-3 Multilingual Variants — Manual Verification Checklist

## Component B — migration (live Zotero)
- [ ] Zotero open; `curl http://localhost:23119/api/users/0/items?limit=1&format=json` returns item JSON with `data.extra`.
- [ ] `npm run migrate:cne` (dry-run) prints correct `cne-*` diffs; `Uyun al-Akhbar` shows title + author romanized lines.
- [ ] `npm run migrate:cne -- --write` writes; `scratch/mlzsync-backup.json` exists; Zotero item's Extra now shows the `cne-*` lines (mlzsync block still present).
- [ ] Re-run `--write`: reports the items as already migrated (idempotent, 0 converted).

## Component A — feature (live Word)
- [ ] Cite tab shows the Variant dropdown (Original / Romanized / Both).
- [ ] "Add from Zotero" a migrated Arabic item; preview updates.
- [ ] Variant = Romanized → footnote/bibliography renders the romanized title + author.
- [ ] Variant = Original → renders Arabic; Variant = Both → Arabic + romanized.
- [ ] Insert a footnote at each variant; the AshaarCite tag stores the variant (re-open confirms).
- [ ] Change the Variant dropdown, click "Refresh citations" → inserted notes re-format to the new variant in place.
- [ ] An item with NO variants still cites correctly (Arabic real fields) under every variant setting.
```

- [ ] **Step 4: Commit**

```bash
git add package.json docs/superpowers/specs/2026-07-17-citation-multilingual-variants-manual-checklist.md
git commit -m "test(cite): register cite-variants suite + SP-3 manual checklist

Co-Authored-By: Claude Opus 4.8 <noreply@anthropic.com>
Claude-Session: https://claude.ai/code/session_01Vy1kwq3DuKd6bpgQGdN6Yw"
```

---

## Self-Review notes (author)

- **Spec coverage:** Component A parser (T1–2), langPrefs (T3), engine proof (T5), tag v2 (T6), pane wiring (T7); Component B converter (T4) + script (T8); tests+checklist (T9). All spec sections mapped.
- **Load-bearing gates:** local-API write (T8 Step 1), BBT `cne-*` passthrough — the runtime reads `item.note`; T7 Step 5 + the manual checklist confirm live passthrough. If BBT strips `cne-*`, add a fallback in `addFromZotero` to read `data.extra` from the local API (documented in spec Risks).
- **Type consistency:** normalized intermediate `{fields, creators}` shared; `parseCne` keys creators by `creatorType`, `parseMlzsync` by flat index (consumed only by `mlzsyncToCneLines` with native creators) — divergence intentional and documented.
- **Tags:** transliteration `ar-Latn`, translation `en` used consistently in parsers, `variantToLangPrefs`, and matches `cite-multi.json`.
```

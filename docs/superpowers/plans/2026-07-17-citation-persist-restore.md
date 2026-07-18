# Persist & Restore Citation Reference Set (SP-B) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development. Steps use checkbox (`- [ ]`) syntax.

**Goal:** Persist the Cite tab's reference set (`cache.items`) in the document via `Office.context.document.settings` and restore it on pane load, so it survives save/close/reopen.

**Architecture:** New dependency-free UMD `cite-store.js` (pure serialize/parse + save/load over an injectable `settings` object) mirrors `cite-zotero.js`'s I/O-seam pattern. `cite-pane.js` restores on load (falls back to the fixture) and saves after the two `cache.items` mutations.

**Tech Stack:** Vanilla ES5/UMD JS, Office Common API `Office.context.document.settings`, node `assert`.

**Design spec:** `docs/superpowers/specs/2026-07-17-citation-persist-restore-design.md`

## Global Constraints

- **No new deps; no build step.** `cite-store.js` is dependency-free UMD (`module.exports` in node, `root.CiteStore` in browser), same shape as `cite-engine.js`/`cite-zotero.js`.
- **Storage:** one setting, key exactly `"AshaarCiteRefs"`, value = JSON string `{"v":1,"items":{…}}`.
- **`parseRefs` never throws** — `null`/`""`/malformed/`v!==1`/missing `items` → `{}`. `loadRefs` never rejects (→ `{}`). `saveRefs` resolves no-op when settings unavailable; rejects only on an async Failed status.
- **Restore-vs-fixture:** saved refs win; fixture is the fallback when there are none (fresh doc / bare browser). The fixture must NOT overwrite a restored set.
- **Save triggers:** only after `addFromZotero`'s merge and in `removeItem`. Fire-and-forget; a save failure is a non-fatal status hint, never blocks the UI.
- **Guard `typeof CiteStore`/`typeof Office`** so the pane + a bare browser still work without them.
- **Tests in the `npm test` chain.** OUT of scope: reading `AshaarCite:` tags / refresh (SP-C).

---

### Task 1: `cite-store.js` — persist/restore over document settings

**Files:**
- Create: `src/taskpane/cite-store.js`
- Test: `tests/cite-store.test.js`
- Modify: `package.json` (add the test to the chain)

**Interfaces:**
- Produces: `CiteStore.REFS_KEY`, `CiteStore.serializeRefs(items)`, `CiteStore.parseRefs(str)`, `CiteStore.saveRefs(items, settingsImpl?)` → Promise, `CiteStore.loadRefs(settingsImpl?)` → Promise<map>, `CiteStore.resolveSettings()`.

- [ ] **Step 1: Write the failing tests** — create `tests/cite-store.test.js`:

```js
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
```

- [ ] **Step 2: Run to verify it fails** — `node tests/cite-store.test.js` → FAIL (module missing).

- [ ] **Step 3: Implement** — create `src/taskpane/cite-store.js`:

```js
(function (root, factory) {
  if (typeof module !== "undefined" && module.exports) { module.exports = factory(); }
  else { root.CiteStore = factory(); }
}(typeof globalThis !== "undefined" ? globalThis : this, function () {
  "use strict";

  var REFS_KEY = "AshaarCiteRefs";

  function serializeRefs(items) { return JSON.stringify({ v: 1, items: items || {} }); }

  function parseRefs(str) {
    if (!str || typeof str !== "string") { return {}; }
    try {
      var obj = JSON.parse(str);
      if (!obj || obj.v !== 1 || !obj.items || typeof obj.items !== "object") { return {}; }
      return obj.items;
    } catch (e) { return {}; }
  }

  // The document settings bag, or null when Office isn't present (bare browser).
  function resolveSettings() {
    try {
      if (typeof Office !== "undefined" && Office.context && Office.context.document &&
          Office.context.document.settings) { return Office.context.document.settings; }
    } catch (e) { /* fall through */ }
    return null;
  }

  function saveRefs(items, settingsImpl) {
    var s = settingsImpl || resolveSettings();
    if (!s) { return Promise.resolve(); } // browser preview: no-op
    return new Promise(function (resolve, reject) {
      s.set(REFS_KEY, serializeRefs(items));
      s.saveAsync(function (res) {
        var status = res && res.status;
        // Office.AsyncResultStatus.Succeeded === "succeeded" (compare by value).
        if (status === "succeeded" || status === 0) { resolve(); }
        else { reject(new Error((res && res.error && res.error.message) || "settings saveAsync failed")); }
      });
    });
  }

  function loadRefs(settingsImpl) {
    var s = settingsImpl || resolveSettings();
    if (!s) { return Promise.resolve({}); }
    return Promise.resolve().then(function () { return parseRefs(s.get(REFS_KEY)); })
      .catch(function () { return {}; });
  }

  return {
    REFS_KEY: REFS_KEY,
    serializeRefs: serializeRefs,
    parseRefs: parseRefs,
    resolveSettings: resolveSettings,
    saveRefs: saveRefs,
    loadRefs: loadRefs
  };
}));
```

- [ ] **Step 4: Run to verify it passes** — `node tests/cite-store.test.js` → PASS. Add `&& node tests/cite-store.test.js` to the `test` script in `package.json`.

- [ ] **Step 5: Full suite** — `npm test` → all pass.

- [ ] **Step 6: Commit**

```bash
git add src/taskpane/cite-store.js tests/cite-store.test.js package.json
git commit -m "feat(cite): cite-store.js — persist/restore reference set via document settings"
```
(End with the two session trailer lines.)

---

### Task 2: `cite-pane.js` — restore on load, save on mutate

**Files:**
- Modify: `src/taskpane/cite-pane.js`
- Modify: `src/taskpane/taskpane.html` (`srcs` + `ASHAAR_ASSET_VERSION`)
- No new node test (DOM/Office wiring; the store is covered by Task 1). Verify via `npm test` regression + the manual Word checklist.

**Interfaces (consumed):** `CiteStore.loadRefs()` → `Promise<map>`, `CiteStore.saveRefs(items)` → `Promise`.

- [ ] **Step 1: Restore on load.** In `cite-pane.js`, `ensureAssets` currently does:

```js
    if (!cache.items) {
      jobs.push(fetchText("fixtures/cite-sample.json").then(function (txt) {
        cache.items = JSON.parse(txt);
      }));
    }
```

Replace that block with a saved-refs-first load (fixture fallback):

```js
    if (!cache.items) {
      var loadRefs = (typeof CiteStore !== "undefined")
        ? CiteStore.loadRefs()
        : Promise.resolve({});
      jobs.push(loadRefs.then(function (saved) {
        if (saved && Object.keys(saved).length) { cache.items = saved; return; }
        return fetchText("fixtures/cite-sample.json").then(function (txt) { cache.items = JSON.parse(txt); });
      }));
    }
```

- [ ] **Step 2: Save after `addFromZotero`'s merge.** In `addFromZotero`, right after the
`Object.keys(fetched).forEach(function (id) { cache.items[id] = fetched[id]; });` merge loop, persist:

```js
        persistRefs();
```

- [ ] **Step 3: Save in `removeItem`.** In `removeItem`, after `delete cache.items[id]` (before or after the rebuild — the map is already mutated), persist:

```js
    persistRefs();
```

- [ ] **Step 4: Add the `persistRefs` helper.** Add near `setStatus` in `cite-pane.js`:

```js
  // Fire-and-forget persistence of the reference set into the document; a save
  // failure is a non-fatal hint (never blocks the UI). No-op without CiteStore/Office.
  function persistRefs() {
    if (typeof CiteStore === "undefined") { return; }
    CiteStore.saveRefs(cache.items).catch(function () {
      setStatus("Couldn't save your reference list to the document.", true);
    });
  }
```

- [ ] **Step 5: `taskpane.html`.** Add `"./cite-store.js"` to the `srcs` array **before** `"./cite-pane.js"` (and before `"./cite-zotero.js"` is fine too — no ordering dependency between store and zotero). Bump `ASHAAR_ASSET_VERSION` to `"20260717-cite-persist"`.

- [ ] **Step 6: Regression + browser smoke.** `npm test` → all pass. Start `node server.mjs` (reuse if :3000 bound), load the pane via Playwright MCP: 0 console errors, the Cite tab still populates from the fixture (bare browser has no Office settings → loadRefs `{}` → fixture) and the preview renders. Do NOT click Add-from-Zotero. Stop any server you started.

- [ ] **Step 7: Commit**

```bash
git add src/taskpane/cite-pane.js src/taskpane/taskpane.html
git commit -m "feat(cite): restore reference set on load + persist on add/remove"
```

---

## Manual checklist (needs Word)
1. Add Zotero items + `×` remove one → save the document → close → reopen: the Items list is the saved set (not the fixture).
2. Brand-new document → shows the fixture (no saved refs).
3. Bare browser (dev server) → still shows the fixture, no console errors (no Office settings).

## Self-review (author)
- Spec coverage: store module (T1), restore-on-load + save-on-mutate + fixture fallback + guards (T2). Mapped.
- Type consistency: `loadRefs()→Promise<map>`, `saveRefs(items)→Promise`, `REFS_KEY` string — consistent across tasks.
- Save status compared by value (`"succeeded"`/`0`) so tests need no Office global; matches the spec's defensive note.

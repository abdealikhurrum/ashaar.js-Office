# Citation Locators + Reference-List Editing + Tagging (SP-A) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add per-citation locators (page/chapter/section/verse), a `×` to remove works from the reference list, and hardened `{style,locale,keys,locators}` content-control tags on inserted citations.

**Architecture:** Extend the SP-1/SP-2 client modules. Three pure modules gain features (`cite-engine.js` locators, `cite-word.js` base64-JSON tags, `cite-zotero.js` locator capture); `cite-pane.js` + `taskpane.html` wire them into the Cite tab. Insertion stays static — tags are written but not yet read (SP-C reads them).

**Tech Stack:** Vanilla ES5/UMD JS, citeproc-js (vendored), Office.js, node `assert` tests. No build step.

**Design spec:** `docs/superpowers/specs/2026-07-17-citation-locators-editing-design.md`

## Global Constraints

Every task must honor these (copied from the spec):
- **No new runtime dependencies; no build step.** UMD modules match existing style (`cite-engine.js`/`cite-word.js`/`cite-zotero.js`): `module.exports` in node, `root.X` in browser.
- **Locator label set is exactly `page`, `chapter`, `section`, `verse`** (CSL-valid locator terms). Value is free text (`42`, `42-45`).
- **Citation tag format:** `"AshaarCite:" + base64(JSON.stringify({v:1, style, locale, keys:[{id, locator, label}]}))`. Bibliography tag: `"AshaarBib:" + base64(JSON.stringify({v:1, style, locale}))`. base64 of JSON — never delimiter-joined (fixes SP-1's `:`/`,` collision). base64 must be UTF-8-safe (keys/locators may be Arabic).
- **`bibliography()` stays unchanged** — it renders every item in the working set; `×` (deleting from `cache.items`) is the entire "remove from bibliography" mechanism.
- **`caywPick` resolves to `[{citekey, locator, label}]`;** `fetchCslJson` still takes **bare citekeys** (callers map `.citekey` out).
- **Locators are per-insertion:** cleared after a successful insert; never stored on the item.
- **The pane never talks to `:23119` directly** (SP-2 constraint, unchanged).
- **New/changed tests run in the `npm test` chain.**
- **OUT of scope (do NOT build):** SP-B (persist/restore reference set in the document), SP-C (refresh/re-format, reading tags back), multilingual langPrefs, Fatemi grouping, Hijri.

---

### Task 1: `cite-engine.js` — locator-aware `cite()`

**Files:**
- Modify: `src/taskpane/cite-engine.js` (the `cite` method of the object returned by `build`)
- Test: `tests/cite-engine.test.js` (append)

**Interfaces:**
- Consumes: existing `CiteEngine.build({styleXml, locales, items, lang})` (unchanged).
- Produces: `engine.cite(citationItems)` where `citationItems` is `Array<{id, locator?, label?}>`
  (also accepts bare-string ids for back-compat). Returns the citation-cluster HTML string.

- [ ] **Step 1: Write the failing test** — append to `tests/cite-engine.test.js`. Reuse the existing engine setup in that file (it already builds an engine from the fixture + Chicago style + locales for `en-US` and `ar`; match how the existing tests obtain `styleXml`/`locales`/`items`). Add:

```js
// --- locators (SP-A) ---
// Pick a real fixture id present in the test's items map (reuse the one the
// existing tests use — assign it to `someId` from that setup).
const withPage = enEngine.cite([{ id: someId, locator: "42", label: "page" }]);
const noLoc = enEngine.cite([{ id: someId }]);
assert.ok(withPage.indexOf("42") !== -1, "page locator value appears in the citation");
assert.ok(noLoc.indexOf("42") === -1, "no locator ⇒ value absent (locator plumbing is real)");
// label-term plumbing: a chapter locator renders the localized 'chap.' term (en)
const withChap = enEngine.cite([{ id: someId, locator: "3", label: "chapter" }]);
assert.ok(/chap/i.test(withChap), "chapter label renders the 'chap.' term (en)");
// bare-string back-compat still works
assert.ok(typeof enEngine.cite([someId]) === "string", "cite() still accepts bare id strings");
// locale-independent plumbing: value also present under the ar engine
const arWithPage = arEngine.cite([{ id: someId, locator: "42", label: "page" }]);
assert.ok(arWithPage.indexOf("42") !== -1, "locator value appears under the ar locale too");
console.log("cite locators test passed");
```

(If the existing test file exposes only one engine, build a second with `lang:"ar"` following the same `build({...})` call, or reuse an existing ar engine variable. Name them to match the file.)

- [ ] **Step 2: Run test to verify it fails**

Run: `node tests/cite-engine.test.js`
Expected: FAIL — current `cite()` ignores locator, so `noLoc` and `withPage` are identical (the `withPage.indexOf("42")` or the `/chap/` assertion fails).

- [ ] **Step 3: Implement** — in `src/taskpane/cite-engine.js`, replace the `cite` method:

```js
      cite: function (citationItems) {
        var items = (citationItems || []).map(function (c) {
          if (typeof c === "string") { return { id: c }; }
          var out = { id: c.id };
          if (c.locator !== undefined && c.locator !== null && String(c.locator) !== "") {
            out.locator = String(c.locator);
          }
          if (c.label) { out.label = c.label; }
          return out;
        });
        return engine.makeCitationCluster(items);
      },
```

- [ ] **Step 4: Run test to verify it passes** — Run: `node tests/cite-engine.test.js` → PASS.

- [ ] **Step 5: Run full suite** — Run: `npm test` → all pass (no regression; existing `cite([id])` callers still work via back-compat).

- [ ] **Step 6: Commit**

```bash
git add src/taskpane/cite-engine.js tests/cite-engine.test.js
git commit -m "feat(cite): locator-aware cite() (page/chapter/section/verse)"
```
(End every commit message with the two trailer lines from the session guidance.)

---

### Task 2: `cite-word.js` — hardened base64-JSON citation/bibliography tags

**Files:**
- Modify: `src/taskpane/cite-word.js` (replace `citationTag`/`parseCitationTag`; add `buildCitationTag`, `buildBibliographyTag`; add UTF-8 base64 helpers; update exports)
- Test: `tests/cite-word.test.js` (replace the old `citationTag`/`parseCitationTag` assertions)

**Interfaces:**
- Produces:
  - `CiteWord.buildCitationTag({style, locale, items})` → `string` tag. `items` = `[{id, locator?, label?}]`.
  - `CiteWord.parseCitationTag(tag)` → `{v, style, locale, keys:[{id, locator, label}]}` or `null`.
  - `CiteWord.buildBibliographyTag({style, locale})` → `string` tag.
  - Existing `sanitize`, `wrapRtlRuns`, `buildNotePayload`, `buildBibliographyPayload` unchanged.

- [ ] **Step 1: Write the failing tests** — in `tests/cite-word.test.js`, remove the old `citationTag`/`parseCitationTag` test block and add:

```js
// --- hardened tags (SP-A) ---
const tag = CiteWord.buildCitationTag({
  style: "chicago-notes-bibliography",
  locale: "ar",
  items: [{ id: "Key:With:Colons", locator: "42", label: "page" }, { id: "Second" }]
});
assert.ok(tag.indexOf("AshaarCite:") === 0, "citation tag is namespaced");
const parsed = CiteWord.parseCitationTag(tag);
assert.strictEqual(parsed.style, "chicago-notes-bibliography");
assert.strictEqual(parsed.locale, "ar");
assert.strictEqual(parsed.keys.length, 2);
assert.strictEqual(parsed.keys[0].id, "Key:With:Colons", "colons in id survive (no delimiter collision)");
assert.strictEqual(parsed.keys[0].locator, "42");
assert.strictEqual(parsed.keys[0].label, "page");
assert.strictEqual(parsed.keys[1].id, "Second");
// non-ASCII (Arabic) values survive the base64 round-trip
const arTag = CiteWord.buildCitationTag({ style: "s", locale: "ar", items: [{ id: "كتاب", locator: "٤٢", label: "page" }] });
assert.strictEqual(CiteWord.parseCitationTag(arTag).keys[0].id, "كتاب");
assert.strictEqual(CiteWord.parseCitationTag(arTag).keys[0].locator, "٤٢");
// non-Ashaar / corrupt tags → null
assert.strictEqual(CiteWord.parseCitationTag("AshaarBibliography"), null);
assert.strictEqual(CiteWord.parseCitationTag("AshaarCite:@@@not-base64@@@"), null);
assert.strictEqual(CiteWord.parseCitationTag(""), null);
// bibliography tag round-trips {style, locale}
const bibTag = CiteWord.buildBibliographyTag({ style: "apa", locale: "en-US" });
assert.ok(bibTag.indexOf("AshaarBib:") === 0);
console.log("hardened tags test passed");
```

- [ ] **Step 2: Run to verify it fails** — Run: `node tests/cite-word.test.js` → FAIL (`buildCitationTag` undefined).

- [ ] **Step 3: Implement** — in `src/taskpane/cite-word.js`, replace `citationTag`/`parseCitationTag` (lines ~98–106) with:

```js
  // UTF-8-safe base64 that works in Node (Buffer) and the Word WebView (btoa).
  function b64encode(str) {
    if (typeof Buffer !== "undefined") { return Buffer.from(str, "utf8").toString("base64"); }
    return btoa(unescape(encodeURIComponent(str)));
  }
  function b64decode(b64) {
    if (typeof Buffer !== "undefined") { return Buffer.from(b64, "base64").toString("utf8"); }
    return decodeURIComponent(escape(atob(b64)));
  }

  function buildCitationTag(o) {
    var payload = {
      v: 1,
      style: o.style,
      locale: o.locale,
      keys: (o.items || []).map(function (i) {
        return { id: i.id, locator: i.locator || null, label: i.label || null };
      })
    };
    return "AshaarCite:" + b64encode(JSON.stringify(payload));
  }

  function parseCitationTag(tag) {
    var s = String(tag || "");
    if (s.indexOf("AshaarCite:") !== 0) { return null; }
    try {
      var obj = JSON.parse(b64decode(s.slice("AshaarCite:".length)));
      if (!obj || !Array.isArray(obj.keys)) { return null; }
      return obj;
    } catch (e) { return null; }
  }

  function buildBibliographyTag(o) {
    return "AshaarBib:" + b64encode(JSON.stringify({ v: 1, style: o.style, locale: o.locale }));
  }
```

Update the returned object: remove `citationTag`, add `buildCitationTag`, `parseCitationTag`, `buildBibliographyTag`:

```js
  return {
    sanitize: sanitize,
    wrapRtlRuns: wrapRtlRuns,
    buildNotePayload: buildNotePayload,
    buildBibliographyPayload: buildBibliographyPayload,
    buildCitationTag: buildCitationTag,
    parseCitationTag: parseCitationTag,
    buildBibliographyTag: buildBibliographyTag
  };
```

- [ ] **Step 4: Run to verify it passes** — Run: `node tests/cite-word.test.js` → PASS.

- [ ] **Step 5: Run full suite** — Run: `npm test` → all pass.

- [ ] **Step 6: Commit**

```bash
git add src/taskpane/cite-word.js tests/cite-word.test.js
git commit -m "feat(cite): hardened base64-JSON citation + bibliography tags"
```

---

### Task 3: `cite-zotero.js` — capture locators from the CAYW pick

**Files:**
- Modify: `src/taskpane/cite-zotero.js` (`parseCaywResult`, `caywPick`)
- Test: `tests/cite-zotero.test.js` (replace the `parseCaywResult`/`caywPick` assertions)

**Interfaces:**
- Produces: `parseCaywResult(text)` → `Array<{citekey, locator, label}>` (locator/label `undefined` when none). `caywPick(fetchImpl?)` → `Promise<Array<{citekey, locator, label}>>`.
- `fetchCslJson(citekeys, fetchImpl?)` is UNCHANGED (still takes bare citekey strings).

- [ ] **Step 1: Write the failing tests** — in `tests/cite-zotero.test.js`, replace the current `parseCaywResult` assertions (they expect `string[]`) and the `caywPick` return assertion with:

```js
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
```

And update the caywPick test: fake fetch returns `"[@YaumulMabasUyun; @IsraaWalMiraaj]"`, assert `keys` deep-equals `[{citekey:"YaumulMabasUyun"},{citekey:"IsraaWalMiraaj"}]` and the URL is still `"/zotero/cayw?format=pandoc"`. Keep the `!res.ok` and cancel (`""` → `[]`) cases.

- [ ] **Step 2: Run to verify it fails** — Run: `node tests/cite-zotero.test.js` → FAIL (returns strings / wrong shape).

- [ ] **Step 3: Implement** — replace `parseCaywResult` in `src/taskpane/cite-zotero.js`:

```js
  var LOC_LABELS = [
    { re: /^(pp?\.?|pages?)$/i, label: "page" },
    { re: /^(chap\.?|chapters?)$/i, label: "chapter" },
    { re: /^(sec\.?|section|§)$/i, label: "section" },
    { re: /^(vv?\.?|verses?)$/i, label: "verse" }
  ];

  // Parse one pandoc item body ("@key" or "@key, p. 42") → {citekey, locator?, label?}.
  function parseCaywItem(body) {
    var trimmed = body.trim();
    if (trimmed.charAt(0) === "@") { trimmed = trimmed.slice(1); }
    var comma = trimmed.indexOf(",");
    if (comma === -1) {
      var bareKey = trimmed.trim();
      return bareKey ? { citekey: bareKey } : null;
    }
    var citekey = trimmed.slice(0, comma).trim();
    if (!citekey) { return null; }
    var suffix = trimmed.slice(comma + 1).trim();
    // suffix forms: "p. 42" | "pp. 42-45" | "chap. 3" | "42"
    var mLabel = /^([A-Za-z.§]+)\s*(.+)$/.exec(suffix);
    if (mLabel) {
      for (var i = 0; i < LOC_LABELS.length; i++) {
        if (LOC_LABELS[i].re.test(mLabel[1])) {
          return { citekey: citekey, locator: mLabel[2].trim(), label: LOC_LABELS[i].label };
        }
      }
      return { citekey: citekey }; // unrecognized label ⇒ no locator
    }
    if (/^[0-9]/.test(suffix)) { return { citekey: citekey, locator: suffix, label: "page" }; }
    return { citekey: citekey };
  }

  function parseCaywResult(text) {
    if (text === null || text === undefined) { return []; }
    var trimmed = String(text).trim();
    if (trimmed === "") { return []; }
    if ((trimmed[0] === "{" && trimmed[trimmed.length - 1] === "}") ||
        (trimmed[0] === "[" && trimmed[trimmed.length - 1] === "]")) {
      trimmed = trimmed.slice(1, -1).trim();
    }
    if (trimmed === "") { return []; }
    return trimmed.split(";")
      .map(function (part) { return parseCaywItem(part); })
      .filter(function (it) { return it && it.citekey; });
  }
```

Update `caywPick`'s final `.then` to just `return parseCaywResult(text);` (already does). No signature change needed — it now resolves to objects.

- [ ] **Step 4: Run to verify it passes** — Run: `node tests/cite-zotero.test.js` → PASS.

- [ ] **Step 5: Run full suite** — Run: `npm test` → all pass.

- [ ] **Step 6: Commit**

```bash
git add src/taskpane/cite-zotero.js tests/cite-zotero.test.js
git commit -m "feat(cite): capture locators from the CAYW pandoc pick"
```

---

### Task 4: `cite-pane.js` + `taskpane.html` — locator UI, `×` remove, tagging, Zotero pre-fill

**Files:**
- Modify: `src/taskpane/cite-pane.js`
- Modify: `src/taskpane/taskpane.html` (item-row is built in JS, so mainly CSS + asset version; add nothing to the static `#cite-items` list markup)
- Modify: `src/taskpane/taskpane.css` (locator row + `×` button styles)
- No new node test (DOM/Office wiring; pure logic is covered by Tasks 1–3). Verify via `npm test` (regression) + browser smoke + manual Word checklist.

**Interfaces (consumed, from Tasks 1–3):**
- `engine.cite([{id, locator, label}])`, `CiteWord.buildCitationTag({style, locale, items})`,
  `CiteWord.buildBibliographyTag({style, locale})`, `CiteZotero.caywPick()` → `[{citekey, locator, label}]`,
  `CiteZotero.fetchCslJson(citekeys)` (bare citekeys).

- [ ] **Step 1: Rebuild `populateItems` to render `×` + a locator row per checked item.** In `src/taskpane/cite-pane.js`, in `populateItems(skipSeed)`, for each item append (after the label): a `×` remove button and a hidden locator row that shows when the checkbox is checked. Example row construction to add inside the existing `forEach`:

```js
      // remove (×) button
      var rm = document.createElement("button");
      rm.type = "button";
      rm.className = "cite-item-remove";
      rm.setAttribute("aria-label", "Remove " + title);
      rm.textContent = "×";
      rm.addEventListener("click", function () { removeItem(id); });
      li.appendChild(rm);

      // locator row (type + value), shown only when the item is checked
      var loc = document.createElement("div");
      loc.className = "cite-locator-row";
      loc.hidden = !cb.checked;
      var sel = document.createElement("select");
      sel.className = "cite-locator-type";
      sel.setAttribute("data-cite-loc-type", id);
      ["page", "chapter", "section", "verse"].forEach(function (t) {
        var opt = document.createElement("option");
        opt.value = t; opt.textContent = t;
        sel.appendChild(opt);
      });
      var val = document.createElement("input");
      val.type = "text";
      val.className = "cite-locator-value";
      val.setAttribute("data-cite-loc-value", id);
      val.placeholder = "e.g. 42";
      val.addEventListener("input", renderPreview);
      sel.addEventListener("change", renderPreview);
      var prefix = document.createElement("span");
      prefix.className = "cite-locator-prefix";
      prefix.textContent = "cite at:";
      loc.appendChild(prefix); loc.appendChild(sel); loc.appendChild(val);
      li.appendChild(loc);

      // toggle the locator row + refresh preview when checked state changes
      cb.addEventListener("change", function () { loc.hidden = !cb.checked; });
```

Keep the existing `cb.addEventListener("change", renderPreview)` too (both change handlers fire).

- [ ] **Step 2: Add `removeItem` and `selectedCitationItems`.** Add near `selectedIds`:

```js
  function removeItem(id) {
    if (cache.items && cache.items[id]) { delete cache.items[id]; }
    itemsPopulated = false;
    populateItems(true);
    renderPreview();
  }

  // Checked items + their locator inputs → [{id, locator, label}] for cite().
  function selectedCitationItems() {
    var out = [];
    var boxes = document.querySelectorAll("#cite-items input[data-cite-id]");
    for (var i = 0; i < boxes.length; i++) {
      if (!boxes[i].checked) { continue; }
      var id = boxes[i].getAttribute("data-cite-id");
      var vEl = document.querySelector('[data-cite-loc-value="' + (window.CSS && CSS.escape ? CSS.escape(id) : id) + '"]');
      var tEl = document.querySelector('[data-cite-loc-type="' + (window.CSS && CSS.escape ? CSS.escape(id) : id) + '"]');
      var item = { id: id };
      if (vEl && vEl.value.trim()) { item.locator = vEl.value.trim(); item.label = (tEl && tEl.value) || "page"; }
      out.push(item);
    }
    return out;
  }
```

(Note: ids are BBT citekeys — alphanumeric — so the attribute-selector lookup is safe; `CSS.escape` guards defensively.)

- [ ] **Step 3: Route preview + insert through the new APIs.** In `renderPreview`, replace the `selectedIds()`/`engine.cite(ids)` usage with `selectedCitationItems()`:

```js
      var items = selectedCitationItems();
      var citeHtml = items.length
        ? CiteWord.wrapRtlRuns(CiteWord.sanitize(engine.cite(items)))
        : "<em>Select one or more items to preview a citation.</em>";
```

In `insertCitation`, build from `selectedCitationItems()`, tag the inserted content control, and clear locators after success:

```js
      var items = selectedCitationItems();
      if (!items.length) { setStatus("Select at least one item to cite.", true); return; }
      var engine = buildEngine(styleFile, lang);
      var html = CiteWord.wrapRtlRuns(CiteWord.sanitize(engine.cite(items)));
      var citeTag = CiteWord.buildCitationTag({ style: styleFile, locale: lang, items: items });
      // ... inside Word.run, after inserting `html` into the note/inline range,
      // wrap that range in a content control and set cc.tag = citeTag; cc.title = "Ashaar Citation";
      // (mirror insertBibliography's insertContentControl() pattern).
```

After the `Word.run(...).then(...)` success path, clear locator inputs:

```js
        var vals = document.querySelectorAll("#cite-items .cite-locator-value");
        for (var i = 0; i < vals.length; i++) { vals[i].value = ""; }
        renderPreview();
```

In `insertBibliography`, set the tag via the new helper: `cc.tag = CiteWord.buildBibliographyTag({ style: styleFile, locale: lang });` (replace the `payload.tag`/`buildBibliographyPayload` tag use; keep the RTL/right-align logic).

- [ ] **Step 4: Pre-fill locators from the Zotero pick in `addFromZotero`.** `caywPick()` now returns `[{citekey, locator, label}]`. Update:

```js
    CiteZotero.caywPick().then(function (picks) {
      if (!picks || !picks.length) { setStatus(""); return; }
      var citekeys = picks.map(function (p) { return p.citekey; });
      return CiteZotero.fetchCslJson(citekeys).then(function (fetched) {
        if (!cache.items) { cache.items = {}; }
        var previouslySelected = selectedIds();
        Object.keys(fetched).forEach(function (id) { cache.items[id] = fetched[id]; });
        itemsPopulated = false;
        populateItems(true);
        // re-check prior selections + the new picks; pre-fill located picks
        previouslySelected.concat(citekeys).forEach(function (id) {
          var cb = byId("cite-item-" + id);
          if (cb) { cb.checked = true; }
        });
        picks.forEach(function (p) {
          var cb = byId("cite-item-" + p.citekey);
          if (cb) { cb.checked = true; }
          var row = cb && cb.parentNode ? cb.parentNode.querySelector(".cite-locator-row") : null;
          if (row) { row.hidden = false; }
          if (p.locator) {
            var vEl = document.querySelector('[data-cite-loc-value="' + p.citekey + '"]');
            var tEl = document.querySelector('[data-cite-loc-type="' + p.citekey + '"]');
            if (vEl) { vEl.value = p.locator; }
            if (tEl && p.label) { tEl.value = p.label; }
          }
        });
        return renderPreview().then(function () {
          setStatus("Added " + citekeys.length + " item(s) from Zotero.");
        });
      });
    }).catch(...)  // unchanged
```

Keep `selectedIds()` (still used above for preserving prior selections).

- [ ] **Step 5: CSS + asset version.** In `taskpane.css` add styles for `.cite-item-remove` (small, muted `×`, right-aligned), `.cite-locator-row` (indented flex row), `.cite-locator-type`/`.cite-locator-value` (compact). In `taskpane.html` bump `window.ASHAAR_ASSET_VERSION` to `"20260717-cite-locators"`.

- [ ] **Step 6: Regression + browser smoke.** Run: `npm test` → all pass. Then start the dev server (`node server.mjs`) and load `https://localhost:3000/src/taskpane/taskpane.html` in a browser (or Playwright): confirm zero console errors, the item list shows `×` buttons, checking an item reveals the locator row, typing a value updates the preview (e.g. adds the page number), and `×` removes an item from the list + preview. (Do NOT auto-click "Add from Zotero" in a headless run — it pops the real picker.)

- [ ] **Step 7: Commit**

```bash
git add src/taskpane/cite-pane.js src/taskpane/taskpane.html src/taskpane/taskpane.css
git commit -m "feat(cite): locator inputs + remove-from-list + citation tagging in the Cite tab"
```

---

## Manual checklist (needs Word + live Zotero — user-run)

1. Cite tab: check an item → locator row appears; set page/chapter/section/verse + a value → the preview citation shows the locator; insert as footnote → the footnote text carries the locator (Arabic term under the `ar` locale).
2. Cite the same reference again with a different page → a second footnote with the new locator (locator inputs were cleared after the first insert).
3. `×` on an item → it leaves the list and the bibliography preview.
4. Add from Zotero with a page set in the CAYW popup → item lands checked with the locator pre-filled.
5. Inspect an inserted citation's content control tag → decodes to `{style, locale, keys:[{id, locator, label}]}` (developer check; nothing consumes it yet — SP-C will).

## Self-review notes (author)
- Spec coverage: locators (T1 + T4), × remove (T4), tagging (T2 + wired in T4), Zotero capture (T3 + T4), bibliography-unchanged (relies on × only) — all mapped.
- Type consistency: `cite([{id,locator,label}])`, `caywPick()→[{citekey,locator,label}]`, `fetchCslJson(bare citekeys)`, tag `{v,style,locale,keys}` — consistent across tasks.
- Locator-term assertion: tests assert on the locator VALUE presence (robust across styles) plus one label-term ("chap.") check, rather than assuming Chicago prints "p." for page (it prints a bare number) — deliberate.

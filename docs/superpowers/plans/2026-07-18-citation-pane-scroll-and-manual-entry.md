# Citation Pane Scrolling + Manual Entry — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax.

**Goal:** Make the task pane scroll vertically, and let users add a citation manually (no Zotero).

**Architecture:** Feature 1 is CSS-only (make `.pane` the scroll container). Feature 2 adds a pure UMD module `cite-manual.js` (name parsing + CSL-JSON build), an inline form in the Cite tab, and wiring in `cite-pane.js` that reuses the existing add/persist/populate/preview path.

**Tech Stack:** Vanilla ES5/UMD, Office.js, Node `assert` tests, no build step.

## Global Constraints

- No build step / no transpilation. ES5 UMD only: `var`, no arrow functions / `const`/`let`. Match existing `cite-*.js`.
- Pure modules take no I/O (no DOM/Word/fetch). `cite-manual.js` must be Node-testable.
- New test files registered in the `package.json` `test` script.
- Bump `ASHAAR_ASSET_VERSION` in `taskpane.html` when pane code changes.
- Manually-added items must be indistinguishable downstream from Zotero items (citable, refreshable, persisted via the existing `CiteStore` path).
- CSL type map: Book→`book`, Book chapter→`chapter`, Journal article→`article-journal`, Webpage→`webpage`.
- Author/Editor entry: one name per line, `Family, Given`; a line with no comma → `{literal: line}`; blank lines ignored.

---

### Task 1: Pane vertical scrolling (CSS only)

**Files:**
- Modify: `src/taskpane/taskpane.css` (the `body` rule ~line 85 and `.pane` rule ~line 99)

**Interfaces:** none (CSS).

- [ ] **Step 1: Make `.pane` the scroll container**

In `src/taskpane/taskpane.css`, add an `html, body` height rule and change `.pane` from
`min-height: 100vh` to a full-height scroll container. Replace the `.pane` block:

```css
html, body {
  height: 100%;
}

.pane {
  display: flex;
  flex-direction: column;
  gap: 14px;
  height: 100%;
  overflow-y: auto;
  overflow-x: hidden;
  padding: 16px;
}
```

(`* { box-sizing: border-box }` already applies, so padding is included in the height.)

- [ ] **Step 2: Verify in a browser**

Run: `npm run dev-server`, open `https://localhost:3000/taskpane.html`, narrow the window height,
switch to the Cite tab. Expected: a vertical scrollbar appears and every Cite control (down to the
preview) is reachable; no horizontal scrollbar; other tabs still look correct. Stop the server.

- [ ] **Step 3: Commit**

```bash
git add src/taskpane/taskpane.css
git commit -m "fix(cite): make the task pane scroll vertically

Co-Authored-By: Claude Opus 4.8 <noreply@anthropic.com>
Claude-Session: https://claude.ai/code/session_01Vy1kwq3DuKd6bpgQGdN6Yw"
```

---

### Task 2: `cite-manual.js` — pure name parsing + CSL-JSON build

**Files:**
- Create: `src/taskpane/cite-manual.js`
- Test: `tests/cite-manual.test.js`
- Modify: `package.json` (register the test)

**Interfaces:**
- Produces (global `CiteManual` / `module.exports`):
  - `parseNames(text)` → array of `{family, given}` / `{family}` / `{literal}`. Splits on newlines;
    trims; skips blank lines. First comma splits family (before) from given (after); empty given
    omitted; no comma → `{literal: trimmedLine}`.
  - `parseDateParts(str)` → `[[y]]` / `[[y,m]]` / `[[y,m,d]]` from a `yyyy`, `yyyy-mm`, or
    `yyyy-mm-dd` string; returns `null` if no leading numeric year.
  - `buildManualItem(values)` → CSL-JSON item. `values`: `{ id, type, title, authors, editors, year,
    publisher, place, containerTitle, volume, issue, pages, url, accessed }` (strings; `type` is one
    of `book|chapter|article|webpage`). Maps to CSL, omits empty fields, parses names/dates.

- [ ] **Step 1: Write the failing test**

Create `tests/cite-manual.test.js`:

```js
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
```

- [ ] **Step 2: Run test to verify it fails**

Run: `node tests/cite-manual.test.js`
Expected: FAIL — `Cannot find module '../src/taskpane/cite-manual'`.

- [ ] **Step 3: Write the implementation**

Create `src/taskpane/cite-manual.js`:

```js
(function (root, factory) {
  if (typeof module !== "undefined" && module.exports) {
    module.exports = factory();
  } else {
    root.CiteManual = factory();
  }
}(typeof globalThis !== "undefined" ? globalThis : this, function () {
  "use strict";

  var TYPE_MAP = { book: "book", chapter: "chapter", article: "article-journal", webpage: "webpage" };

  function parseNames(text) {
    var out = [];
    var lines = String(text || "").split(/\r?\n/);
    for (var i = 0; i < lines.length; i++) {
      var line = lines[i].trim();
      if (!line) { continue; }
      var comma = line.indexOf(",");
      if (comma === -1) { out.push({ literal: line }); continue; }
      var family = line.slice(0, comma).trim();
      var given = line.slice(comma + 1).trim();
      var name = { family: family };
      if (given) { name.given = given; }
      out.push(name);
    }
    return out;
  }

  function parseDateParts(str) {
    var s = String(str || "").trim();
    if (!/^\d{1,4}/.test(s)) { return null; }
    var parts = s.split("-");
    var nums = [];
    for (var i = 0; i < parts.length; i++) {
      var n = parseInt(parts[i], 10);
      if (isNaN(n)) { break; }
      nums.push(n);
    }
    return nums.length ? [nums] : null;
  }

  function setStr(item, key, val) {
    var v = (val === undefined || val === null) ? "" : String(val).trim();
    if (v) { item[key] = v; }
  }

  function buildManualItem(values) {
    var v = values || {};
    var item = { id: v.id, type: TYPE_MAP[v.type] || "document" };
    setStr(item, "title", v.title);
    setStr(item, "container-title", v.containerTitle);
    setStr(item, "publisher", v.publisher);
    setStr(item, "publisher-place", v.place);
    setStr(item, "volume", v.volume);
    setStr(item, "issue", v.issue);
    setStr(item, "page", v.pages);
    setStr(item, "URL", v.url);
    var issued = parseDateParts(v.year);
    if (issued) { item.issued = { "date-parts": issued }; }
    var accessed = parseDateParts(v.accessed);
    if (accessed) { item.accessed = { "date-parts": accessed }; }
    var authors = parseNames(v.authors);
    if (authors.length) { item.author = authors; }
    var editors = parseNames(v.editors);
    if (editors.length) { item.editor = editors; }
    return item;
  }

  return { parseNames: parseNames, parseDateParts: parseDateParts, buildManualItem: buildManualItem };
}));
```

- [ ] **Step 4: Run test to verify it passes**

Run: `node tests/cite-manual.test.js`
Expected: PASS — `cite-manual.test.js passed`.

- [ ] **Step 5: Register the test**

In `package.json`, append to the `test` script chain (after `node tests/cite-classify.test.js`):

```
 && node tests/cite-manual.test.js
```

Run: `npm test` → full suite PASS.

- [ ] **Step 6: Commit**

```bash
git add src/taskpane/cite-manual.js tests/cite-manual.test.js package.json
git commit -m "feat(cite): cite-manual — name parsing + CSL-JSON item build (pure)

Co-Authored-By: Claude Opus 4.8 <noreply@anthropic.com>
Claude-Session: https://claude.ai/code/session_01Vy1kwq3DuKd6bpgQGdN6Yw"
```

---

### Task 3: Manual-entry form (HTML) + wiring (cite-pane.js) + registration

**Files:**
- Modify: `src/taskpane/taskpane.html` (add the form after the Zotero actions section ~line 606; add `./cite-manual.js` to the `srcs` list after `./cite-classify.js`; bump `ASHAAR_ASSET_VERSION`)
- Modify: `src/taskpane/taskpane.css` (styles for the form field rows — reuse existing `.field`/`.controls` classes; add a `[hidden]` respect and a `.cite-manual-form` container)
- Modify: `src/taskpane/cite-pane.js` (toggle, per-type field visibility, submit handler)

**Interfaces:**
- Consumes: `CiteManual.buildManualItem` (Task 2); existing `cache.items`, `persistRefs`, `populateItems`, `renderPreview`, `selectedIds`, `byId`.
- Produces: no new public surface — internal wiring.

- [ ] **Step 1: Add the form markup**

In `src/taskpane/taskpane.html`, immediately after the `Zotero actions` section (the
`</section>` closing the block that contains `#cite-add-zotero`, ~line 606), insert:

```html
        <section class="actions single-action" aria-label="Manual entry">
          <button id="cite-manual-toggle" type="button" class="button--secondary">Add manually</button>
        </section>

        <section id="cite-manual-form" class="cite-manual-form" aria-label="Manual citation" hidden>
          <div class="field">
            <label for="cite-manual-type">Type</label>
            <select id="cite-manual-type">
              <option value="book" selected>Book</option>
              <option value="chapter">Book chapter</option>
              <option value="article">Journal article</option>
              <option value="webpage">Webpage</option>
            </select>
          </div>
          <div class="field" data-types="book chapter article webpage">
            <label for="cite-manual-title">Title</label>
            <input id="cite-manual-title" type="text" />
          </div>
          <div class="field" data-types="book chapter article webpage">
            <label for="cite-manual-authors">Author(s) — one per line, “Family, Given”</label>
            <textarea id="cite-manual-authors" rows="2"></textarea>
          </div>
          <div class="field" data-types="chapter article">
            <label for="cite-manual-container">Container title (book / journal)</label>
            <input id="cite-manual-container" type="text" />
          </div>
          <div class="field" data-types="chapter">
            <label for="cite-manual-editors">Editor(s) — one per line, “Family, Given”</label>
            <textarea id="cite-manual-editors" rows="2"></textarea>
          </div>
          <div class="field" data-types="book chapter article webpage">
            <label for="cite-manual-year">Year</label>
            <input id="cite-manual-year" type="text" inputmode="numeric" />
          </div>
          <div class="field" data-types="book chapter">
            <label for="cite-manual-publisher">Publisher</label>
            <input id="cite-manual-publisher" type="text" />
          </div>
          <div class="field" data-types="book chapter">
            <label for="cite-manual-place">Place</label>
            <input id="cite-manual-place" type="text" />
          </div>
          <div class="field" data-types="article">
            <label for="cite-manual-volume">Volume</label>
            <input id="cite-manual-volume" type="text" />
          </div>
          <div class="field" data-types="article">
            <label for="cite-manual-issue">Issue</label>
            <input id="cite-manual-issue" type="text" />
          </div>
          <div class="field" data-types="chapter article">
            <label for="cite-manual-pages">Pages</label>
            <input id="cite-manual-pages" type="text" />
          </div>
          <div class="field" data-types="webpage">
            <label for="cite-manual-url">URL</label>
            <input id="cite-manual-url" type="text" />
          </div>
          <div class="field" data-types="webpage">
            <label for="cite-manual-accessed">Accessed (YYYY-MM-DD)</label>
            <input id="cite-manual-accessed" type="text" />
          </div>
          <section class="actions two-col">
            <button id="cite-manual-add" type="button">Add to list</button>
            <button id="cite-manual-cancel" type="button" class="button--secondary">Cancel</button>
          </section>
        </section>
```

- [ ] **Step 2: Add form container CSS**

In `src/taskpane/taskpane.css`, append:

```css
.cite-manual-form {
  display: flex;
  flex-direction: column;
  gap: 10px;
  padding: 12px;
  border: 1px solid var(--line);
  border-radius: 6px;
}
.cite-manual-form .field[hidden] { display: none; }
```

- [ ] **Step 3: Register the module + bump asset version**

In `src/taskpane/taskpane.html`, add `"./cite-manual.js",` to the `srcs` array right after
`"./cite-classify.js",`. Bump the version:

```js
      window.ASHAAR_ASSET_VERSION = "20260718-cite-manual";
```

- [ ] **Step 4: Wire the form in cite-pane.js**

In `src/taskpane/cite-pane.js`, add these functions (near `addFromZotero`, before `bind`):

```js
  // Show only the field rows whose data-types include the selected type.
  function syncManualFields() {
    var type = (byId("cite-manual-type") || {}).value || "book";
    var form = byId("cite-manual-form");
    if (!form) { return; }
    var rows = form.querySelectorAll(".field[data-types]");
    for (var i = 0; i < rows.length; i++) {
      var types = (rows[i].getAttribute("data-types") || "").split(/\s+/);
      rows[i].hidden = types.indexOf(type) === -1;
    }
  }

  function toggleManualForm(show) {
    var form = byId("cite-manual-form");
    if (!form) { return; }
    form.hidden = (show === undefined) ? !form.hidden : !show;
    if (!form.hidden) { syncManualFields(); }
  }

  // Generate a stable id that doesn't collide with existing items.
  function nextManualId() {
    var n = 1;
    while (cache.items && Object.prototype.hasOwnProperty.call(cache.items, "manual-" + n)) { n++; }
    return "manual-" + n;
  }

  function mval(id) { var el = byId(id); return el ? el.value : ""; }

  function addManualItem() {
    if (typeof CiteManual === "undefined") { return; }
    var title = mval("cite-manual-title").trim();
    if (!title) { setStatus("Enter a title for the manual citation.", true); return; }
    if (!cache.items) { cache.items = {}; }
    var id = nextManualId();
    var item = CiteManual.buildManualItem({
      id: id, type: mval("cite-manual-type"),
      title: title, authors: mval("cite-manual-authors"), editors: mval("cite-manual-editors"),
      year: mval("cite-manual-year"), publisher: mval("cite-manual-publisher"),
      place: mval("cite-manual-place"), containerTitle: mval("cite-manual-container"),
      volume: mval("cite-manual-volume"), issue: mval("cite-manual-issue"),
      pages: mval("cite-manual-pages"), url: mval("cite-manual-url"),
      accessed: mval("cite-manual-accessed")
    });
    var previouslySelected = selectedIds();
    cache.items[id] = item;
    persistRefs();
    itemsPopulated = false;
    populateItems(true);
    previouslySelected.concat([id]).forEach(function (sid) {
      var cb = byId("cite-item-" + sid);
      if (cb) { cb.checked = true; }
    });
    // Clear the form inputs for the next entry.
    var ids = ["title", "authors", "editors", "year", "publisher", "place", "container",
      "volume", "issue", "pages", "url", "accessed"];
    ids.forEach(function (k) { var el = byId("cite-manual-" + k); if (el) { el.value = ""; } });
    toggleManualForm(false);
    return renderPreview().then(function () { setStatus("Added manual citation."); });
  }
```

Then in `bind()` (after the `cite-add-zotero` wiring), add listeners:

```js
    var manualToggle = byId("cite-manual-toggle");
    if (manualToggle) { manualToggle.addEventListener("click", function () { toggleManualForm(); }); }
    var manualType = byId("cite-manual-type");
    if (manualType) { manualType.addEventListener("change", syncManualFields); }
    var manualAdd = byId("cite-manual-add");
    if (manualAdd) { manualAdd.addEventListener("click", addManualItem); }
    var manualCancel = byId("cite-manual-cancel");
    if (manualCancel) { manualCancel.addEventListener("click", function () { toggleManualForm(false); }); }
```

- [ ] **Step 5: Parse-check + full suite**

Run: `node --check src/taskpane/cite-pane.js` → clean.
Run: `npm test` → full suite PASS (proves no cross-file breakage; the pure module is covered by Task 2).

- [ ] **Step 6: Browser smoke test**

Run: `npm run dev-server`, open the pane, Cite tab → click **Add manually** → form appears →
switch Type and confirm fields show/hide → fill a book (title + `Daftary, Farhad` + year) → **Add to
list** → item appears checked in Items, preview renders it, form clears/hides. No console errors. Stop.

- [ ] **Step 7: Commit**

```bash
git add src/taskpane/taskpane.html src/taskpane/taskpane.css src/taskpane/cite-pane.js
git commit -m "feat(cite): manual citation entry form (no Zotero)

Co-Authored-By: Claude Opus 4.8 <noreply@anthropic.com>
Claude-Session: https://claude.ai/code/session_01Vy1kwq3DuKd6bpgQGdN6Yw"
```

---

## Self-Review

**Spec coverage:** Pane scrolling → Task 1. Manual form UI (button + inline form, per-type fields) →
Task 3 Steps 1–2. Author/editor one-per-line parsing + CSL build → Task 2 (`parseNames`,
`buildManualItem`). Type map, empty-field omission, date-parts → Task 2. Integration (cache.items +
persist + populate + preview, generated id) → Task 3 Step 4. Module registration + asset bump → Task 3
Step 3. Tests registered → Task 2 Step 5. All covered.

**Placeholder scan:** none — every step has complete code/commands.

**Type consistency:** `buildManualItem(values)` value keys (`id, type, title, authors, editors, year,
publisher, place, containerTitle, volume, issue, pages, url, accessed`) match exactly what
`addManualItem` passes and the form field ids (`cite-manual-<k>`, with `container` mapped to
`containerTitle`). `parseNames`/`parseDateParts` return shapes match the tests and `buildManualItem`
usage. `nextManualId` returns `manual-<n>` consistent with the spec's id scheme.

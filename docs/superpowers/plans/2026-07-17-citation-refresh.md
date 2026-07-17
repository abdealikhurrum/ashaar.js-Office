# Refresh & Re-format Inserted Citations (SP-C) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development. Steps use checkbox (`- [ ]`) syntax.

**Goal:** A "Refresh citations" button that scans the document's `AshaarCite:`/`AshaarBib:` content controls, re-runs the engine at the pane's current style/locale, and rewrites each in place (preserving keys+locators). Also retires SP-A's vestigial `buildBibliographyPayload.tag`.

**Architecture:** A tiny pure `citationItemsFromTag` helper in `cite-word.js` (Node-testable). The bulk is `cite-pane.js` Office.js orchestration: enumerate content controls (main body + footnote/endnote bodies), dispatch by parsed tag, re-render via the existing insert pipeline (Arabic→OOXML, LTR→HTML), rewrite body + tag. Footnote reachability + per-CC rewrite are confirmed live in Word (defensive: refresh what's reachable, report counts).

**Tech Stack:** Vanilla ES5/UMD JS, Office.js `contentControls` / footnote-body traversal / `insertHtml`/`insertOoxml`, node `assert`.

**Design spec:** `docs/superpowers/specs/2026-07-17-citation-refresh-design.md`

## Global Constraints

- **No new deps; no build step.** Reuse existing modules (`cite-engine`, `cite-word`, `cite-store`, `AshaarTabStop.wrapOoxml`) and the pane's existing render helpers.
- **Target = pane's CURRENT style/locale.** Refresh rewrites each `AshaarCite:` tag's `style`/`locale` to the current selection; **keys + locators are preserved** from the old tag.
- **Reuse the SP-A / Arabic-OOXML render pipeline** — do NOT duplicate it. Factor a shared helper both `insertCitation` and `refreshCitations` call (Arabic → `buildCitationParagraphOoxml`+`wrapOoxml`+`insertOoxml`; LTR/inline → `insertHtml(wrapRtlRuns(sanitize(...)))`).
- **Defensive + graceful:** enumerate main-body `document.contentControls` PLUS footnote/endnote bodies' `contentControls`; guard each footnote step so an unsupported Word build just yields the main-body set. A per-CC failure is caught + counted, never aborts the whole run. A citation id absent from `cache.items` is counted "unresolved" + skipped (not errored).
- **Only touch our controls:** `parseCitationTag(tag)` → null means "not ours" → skip.
- **Guard `typeof Word`/`CiteWord`.** Browser → "Refresh needs Word."
- **Tests in the `npm test` chain.** OUT of scope: auto-refresh on style change; union-of-cited-keys bibliography.

---

### Task 1: `cite-word.js` — `citationItemsFromTag` + retire vestigial bib tag field

**Files:** Modify `src/taskpane/cite-word.js`; Test `tests/cite-word.test.js` (append + adjust).

**Interfaces:** Produces `CiteWord.citationItemsFromTag(parsed)` → `[{id, locator?, label?}]`.

- [ ] **Step 1: Write the failing tests** — append to `tests/cite-word.test.js`:

```js
// --- citationItemsFromTag (SP-C) ---
var parsed = CiteWord.parseCitationTag(CiteWord.buildCitationTag({
  style: "s", locale: "ar",
  items: [{ id: "A", locator: "42", label: "page" }, { id: "B" }]
}));
assert.deepStrictEqual(CiteWord.citationItemsFromTag(parsed),
  [{ id: "A", locator: "42", label: "page" }, { id: "B" }],
  "maps tag keys to cite items, dropping null locator/label");
assert.deepStrictEqual(CiteWord.citationItemsFromTag({ keys: [] }), []);
assert.deepStrictEqual(CiteWord.citationItemsFromTag({}), [], "missing keys → []");
assert.deepStrictEqual(CiteWord.citationItemsFromTag(null), [], "null → []");
console.log("citationItemsFromTag test passed");
```

Also **adjust the existing bibliography-payload test**: remove any assertion that `buildBibliographyPayload(...)` returns a `tag` field (it no longer will). Keep assertions on its `html`/`direction`.

- [ ] **Step 2: Run to verify it fails** — `node tests/cite-word.test.js` → FAIL (`citationItemsFromTag` undefined).

- [ ] **Step 3: Implement** — in `src/taskpane/cite-word.js`:

```js
  function citationItemsFromTag(parsed) {
    var keys = (parsed && parsed.keys) || [];
    return keys.map(function (k) {
      var it = { id: k.id };
      if (k.locator) { it.locator = k.locator; }
      if (k.label) { it.label = k.label; }
      return it;
    });
  }
```

Add `citationItemsFromTag` to the returned object. In `buildBibliographyPayload`, remove the `tag: o.tag || "AshaarBibliography"` property (return only `{ html, direction }`).

- [ ] **Step 4: Run to verify it passes** — `node tests/cite-word.test.js` → PASS.

- [ ] **Step 5: Full suite** — `npm test` → all pass (confirm no other consumer read `buildBibliographyPayload().tag` — grep; insertBibliography uses `buildBibliographyTag`).

- [ ] **Step 6: Commit**

```bash
git add src/taskpane/cite-word.js tests/cite-word.test.js
git commit -m "feat(cite): citationItemsFromTag helper; drop vestigial bib payload tag"
```
(End with the two session trailer lines.)

---

### Task 2: `cite-pane.js` — `refreshCitations()` orchestration + shared render helper + button

**Files:** Modify `src/taskpane/cite-pane.js`, `src/taskpane/taskpane.html` (button + `ASHAAR_ASSET_VERSION`). No new node test (Office.js orchestration; `citationItemsFromTag` is covered by Task 1). Verify via `npm test` regression + browser-load smoke + the manual Word checklist.

**Interfaces (consumed):** `CiteWord.parseCitationTag`, `CiteWord.citationItemsFromTag`, `CiteWord.buildCitationTag`, `CiteWord.buildBibliographyTag`, `CiteWord.sanitize`, `CiteWord.wrapRtlRuns`, `CiteWord.buildCitationParagraphOoxml`, `AshaarTabStop.wrapOoxml`, existing `buildEngine`, `ensureAssets`, `readDocCsFont`, `isRtlLang`, `currentStyleFile`/`currentLang`.

- [ ] **Step 1: Factor a shared citation-render helper.** In `cite-pane.js`, extract the body-building logic currently inside `insertCitation` into a reusable function so refresh reuses it verbatim:

```js
  // Build the citation body for a range and insert it, honoring RTL (OOXML) vs
  // LTR (HTML). Returns a Promise. `ctx` + `range` are live Word objects.
  // For rtl it reads the doc cs font and inserts OOXML; else inserts sanitized HTML.
  function renderCitationInto(ctx, range, items, styleFile, lang) {
    var engine = buildEngine(styleFile, lang);
    if (isRtlLang(lang)) {
      return readDocCsFont(ctx, range).then(function (csFont) {
        var pkg = AshaarTabStop.wrapOoxml(
          CiteWord.buildCitationParagraphOoxml(CiteWord.sanitize(engine.cite(items)), { csFont: csFont }));
        return range.insertOoxml(pkg, Word.InsertLocation.replace);
      });
    }
    var html = CiteWord.wrapRtlRuns(CiteWord.sanitize(engine.cite(items)));
    return Promise.resolve(range.insertHtml(html, Word.InsertLocation.replace));
  }
```

Refactor `insertCitation`'s footnote/endnote branch to call `renderCitationInto(ctx, noteRange, items, styleFile, lang)` (keeping its own CC-tag write on the returned range) so both paths share one renderer. Keep inline + en-US behavior identical.

- [ ] **Step 2: Implement `refreshCitations()`.** Add:

```js
  function collectTaggedControls(ctx) {
    // Main-body content controls + footnote/endnote body content controls.
    // Footnote enumeration is confirmed against the live WordApi during impl;
    // guard so an unsupported build still returns the main-body set.
    var groups = [];
    var main = ctx.document.contentControls;
    main.load("items/tag");
    groups.push(main);
    try {
      var notes = ctx.document.body.footnotes; // NoteItemCollection (confirm exact member live)
      notes.load("items");
      // note.body.contentControls loaded after first sync (see refresh flow)
      groups.push({ __notes: notes });
    } catch (e) { /* footnotes API unavailable on this build */ }
    return groups;
  }

  function refreshCitations() {
    if (typeof Word === "undefined" || !Word.run) { setStatus("Refresh needs Word.", true); return; }
    var styleFile = currentStyleFile(), lang = currentLang();
    setStatus("Refreshing citations…");
    ensureAssets(styleFile).then(function () {
      return Word.run(function (ctx) {
        // 1) gather all candidate CCs (main body + footnote/endnote bodies)
        //    — collect into a flat array `ccs` after the necessary sync(s).
        // 2) for each cc: parsed = CiteWord.parseCitationTag(cc.tag); if null skip.
        //    AshaarCite: items = citationItemsFromTag(parsed);
        //       if any it.id not in cache.items → unresolved++, skip;
        //       else renderCitationInto(ctx, cc.getRange(), items, styleFile, lang)
        //            then cc.tag = buildCitationTag({style:styleFile, locale:lang, items});  refreshed++
        //    AshaarBib: rebuild bibliography from cache.items at styleFile/lang →
        //       rtl? OOXML(buildCitationParagraphOoxml(sanitize(engine.bibliography()),{csFont}))
        //           : insertHtml(wrapRtlRuns(sanitize(engine.bibliography())));
        //       cc.tag = buildBibliographyTag({style:styleFile, locale:lang}); bibs++
        //    wrap each per-cc op in try/catch → failed++ (continue).
        // 3) return ctx.sync().
      }).then(function () {
        // setStatus with counts: refreshed / bibs / unresolved / failed / footnote note.
      });
    }).catch(function (e) {
      setStatus("Refresh failed: " + (e && e.message ? e.message : String(e)), true);
    });
  }
```

Implement the body per the comments. Concrete requirements:
- Load `cc.tag` for every candidate before parsing (one `ctx.sync()` after loading main + note CCs).
- Enumerate footnote/endnote CCs: for each `note` in the notes collection, `note.body.contentControls.load("items/tag")` then sync; flatten into the candidate list. **Confirm the exact footnote-collection member (`body.footnotes` vs `getFootnoteBody` vs `range.footnotes`) against MS Learn / the live WordApi before finalizing** (mirrors SP-2's translator check). If none is available, proceed with main-body only.
- `cache.items` membership check: `Object.prototype.hasOwnProperty.call(cache.items, it.id)`.
- Batch: prefer few `ctx.sync()` calls (load-all → process → one final sync); the RTL path's `readDocCsFont` adds its own sync per Arabic CC (acceptable).
- Final status, e.g.: `"Refreshed 3 citation(s), 1 bibliography."` + `" 1 unresolved (re-add from Zotero)."` if any + `" No footnote citations reached."` if the footnote group was empty/unsupported.

- [ ] **Step 3: Wire the button.** In `cite-pane.js` `bind()`, add `byId("cite-refresh").addEventListener("click", refreshCitations);` (guard for missing element like the others). In `taskpane.html`, add near the insert actions: `<button id="cite-refresh" type="button" class="button--secondary">Refresh citations</button>`. Bump `ASHAAR_ASSET_VERSION` to `"20260717-cite-refresh"`.

- [ ] **Step 4: Regression + browser smoke.** `npm test` → all pass. Start `node server.mjs` (reuse if bound), load the pane via Playwright MCP: 0 console errors, the Refresh button is present, and clicking it in a bare browser shows "Refresh needs Word." (Word isn't available in a browser.) Stop any server you started.

- [ ] **Step 5: Commit**

```bash
git add src/taskpane/cite-pane.js src/taskpane/taskpane.html
git commit -m "feat(cite): Refresh citations — re-format tagged citations + bibliography in place"
```

---

## Manual checklist (needs Word — the live-verify gate)
1. Insert a footnote citation (Arabic), an inline citation (en), and a bibliography. Change the pane style (Chicago → APA) and/or locale. Click **Refresh citations**.
2. Confirm: every citation + the bibliography re-render to the new style IN PLACE, footnotes included; locators are preserved; the status counts are accurate.
3. Confirm footnote coverage specifically (the load-bearing unknown): the footnote citation updated. If the status says "No footnote citations reached," report it — footnote enumeration needs a different member on this Word build.
4. A citation whose reference was `×`-removed from the list → reported "unresolved", not crashed; other citations still refresh.
5. Non-Ashaar content (plain text, other content controls) untouched.

## Self-review (author)
- Spec coverage: citationItemsFromTag + bib-tag cleanup (T1); enumerate+dispatch+rewrite+report, shared renderer, button (T2). Mapped.
- Type consistency: `citationItemsFromTag(parsed)→[{id,locator?,label?}]` feeds `engine.cite(items)` (same shape SP-A cite() takes) and `buildCitationTag({items})`; consistent.
- Risk: footnote-CC enumeration is the live-verify gate (T2 Step 2 confirms the exact member; graceful main-body-only fallback). Refresh orchestration is Office.js → manual checklist is its gate, `citationItemsFromTag` carries the node coverage.

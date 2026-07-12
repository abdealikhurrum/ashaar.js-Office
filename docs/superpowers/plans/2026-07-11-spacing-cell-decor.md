# Spacing-Cell Decorations (SP3) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Put a symbol (hemistich / decorative glyph) and a fill color (+ symbol text color) into gap cells — configured by a profile default per slot-position plus a per-slot override — applied via a decorate pass that fills the (SP1-tagged) spacing cells.

**Architecture:** `AshaarOverrides.resolveSlotDecor` (pure) merges a per-slot override onto the profile default. `word-html.js` gains `setTagSlotDecor` (write/remove a per-slot decoration on the block tag, mirroring `setTagOverride`) and normalizes `slotDecor` in `parseContentControlTag`. `profiles.js` gains a `spacingDecor` bucket (superseding the unused flat `misraSymbol`/`symbolColor`). `applyProfileToQaseeda` gains a decorate pass that, for each `kind:"spacing"` cell, writes the resolved symbol (+ text color) and sets `TableCell.shadingColor`. The SP2 active-cell editor is extended so a spacing cell shows a decoration editor. **Insert bake-in is out of MVP scope** — decorations apply via Save & Apply / per-slot edit (no generator changes).

**Tech Stack:** Vanilla JS (ES5/UMD, `var`/`function`, no build step), Office.js v1 (`TableCell.shadingColor`, `Body.font.color`), Node `assert` tests. Never edit `src/vendor/`.

**Spec:** `docs/superpowers/specs/2026-07-11-spacing-cell-decor-design.md`
**Builds on:** SP1 (`AshaarCellMap`, tag `cells`, spacing `slot` + `kind`), SP2 (`AshaarOverrides`, `setTagOverride`, active-cell editor, `applyProfileToQaseeda` block-walk).

## Global Constraints

- ES5/UMD only; never edit `src/vendor/`.
- Pure logic (`resolveSlotDecor`, `setTagSlotDecor`, profile bucket) is **node-tested**; the decorate pass + editor UI are **manual-verify** (final task).
- Decoration payload: `{ symbol?: string, fill?: "RRGGBB", color?: "RRGGBB" }`. In an override, an **empty string = explicit none** (suppress profile default); an absent field = inherit.
- Profile default keyed by slot-position (`"A#1"`); per-slot override on the block tag keyed `"<tableIndex>:<slot>"` via `AshaarOverrides.overrideKey`.
- Decorated gaps stay `kind:"spacing"` — justification still skips them.
- Colors are hex without `#`.
- `npm test` green after every task.

---

### Task 1: Pure decor helpers + profile bucket

**Files:**
- Modify: `src/taskpane/cell-overrides.js` (`resolveSlotDecor` + export)
- Modify: `src/taskpane/word-html.js` (`setTagSlotDecor` + export; `parseContentControlTag` normalizes `slotDecor`)
- Modify: `src/taskpane/profiles.js` (`spacingDecor` default + `mergeProfile` nested bucket)
- Test: `tests/cell-overrides.test.js`, `tests/profiles.test.js`

**Interfaces:**
- `AshaarOverrides.resolveSlotDecor(profileDecor, override) → { symbol, fill, color }` — per field: `override` has the key → use it (empty string kept as `""` = none); else inherit `profileDecor`; else `""`.
- `AshaarWord.setTagSlotDecor(tag, key, decor) → tag` — set `payload.slotDecor[key]` (any non-empty field) or delete it (all-empty/null). Non-ashaar tags unchanged.
- `parseContentControlTag(tag).slotDecor` → object (`{}` when absent).
- `defaultProfile().spacingDecor === {}`; carried through `normalizeProfile`.

- [ ] **Step 1: Failing tests**

Append to `tests/cell-overrides.test.js`:
```js
// ── resolveSlotDecor: override wins per field; "" = explicit none; else inherit
{
  const prof = { symbol: "؎", fill: "f5f0e0", color: "a7352a" };
  assert.deepStrictEqual(AshaarOverrides.resolveSlotDecor(prof, null),
    { symbol: "؎", fill: "f5f0e0", color: "a7352a" }, "no override → profile");
  assert.deepStrictEqual(AshaarOverrides.resolveSlotDecor(prof, { symbol: "*" }),
    { symbol: "*", fill: "f5f0e0", color: "a7352a" }, "symbol overridden");
  assert.deepStrictEqual(AshaarOverrides.resolveSlotDecor(prof, { symbol: "" }),
    { symbol: "", fill: "f5f0e0", color: "a7352a" }, "empty string suppresses");
  assert.deepStrictEqual(AshaarOverrides.resolveSlotDecor(null, null),
    { symbol: "", fill: "", color: "" }, "nothing → all none");
}

// ── setTagSlotDecor: add / remove, round-trip, cells/overrides intact ────────
{
  const t0 = AshaarWord.contentControlTag("poem", { qaseeda: "Q" }, [[["c","g","c"]]]);
  const t1 = AshaarWord.setTagSlotDecor(t0, "0:A#1", { symbol: "؎", fill: "eeeeee" });
  const p1 = AshaarWord.parseContentControlTag(t1);
  assert.deepStrictEqual(p1.slotDecor, { "0:A#1": { symbol: "؎", fill: "eeeeee" } });
  assert.deepStrictEqual(p1.cells, [[["c","g","c"]]], "cells intact");
  assert.equal(p1.qaseeda, "Q");
  const t2 = AshaarWord.setTagSlotDecor(t1, "0:A#1", null);
  assert.deepStrictEqual(AshaarWord.parseContentControlTag(t2).slotDecor, {}, "removed");
  const t3 = AshaarWord.setTagSlotDecor(t1, "0:A#1", { symbol: "", fill: "", color: "" });
  assert.deepStrictEqual(AshaarWord.parseContentControlTag(t3).slotDecor, {}, "all-empty removes");
  assert.strictEqual(AshaarWord.setTagSlotDecor("nope", "0:A#1", { symbol: "x" }), "nope");
}
```
Append to `tests/profiles.test.js`:
```js
// ── spacingDecor bucket ──────────────────────────────────────────────────────
{
  assert.deepStrictEqual(defaultProfile("Q").spacingDecor, {}, "defaults to empty");
  const n = normalizeProfile({ name: "Q", spacingDecor: { "A#1": { symbol: "؎" } } });
  assert.deepStrictEqual(n.spacingDecor, { "A#1": { symbol: "؎" } }, "carried through merge");
}
```

- [ ] **Step 2: Run to verify they fail**

Run: `node tests/cell-overrides.test.js`; `node tests/profiles.test.js`
Expected: FAIL (`resolveSlotDecor` / `setTagSlotDecor` not functions; `spacingDecor` undefined).

- [ ] **Step 3: Implement `resolveSlotDecor`**

In `src/taskpane/cell-overrides.js`, add and export:
```js
  // Merge a per-slot decoration override onto the profile default. Per field:
  // an override key present wins (empty string = explicit none); else inherit
  // the profile; else "". Fields: symbol / fill / color.
  function resolveSlotDecor(profileDecor, override) {
    profileDecor = profileDecor || {};
    override = override || {};
    function pick(k) { return (k in override) ? (override[k] || "") : (profileDecor[k] || ""); }
    return { symbol: pick("symbol"), fill: pick("fill"), color: pick("color") };
  }
```
Export: add `resolveSlotDecor: resolveSlotDecor,` to the returned object.

- [ ] **Step 4: Implement `setTagSlotDecor` + parse normalization**

In `src/taskpane/word-html.js`, after `setTagOverride`, add:
```js
  // Return a copy of an "ashaar:" tag with one per-slot decoration set or removed.
  // A null/all-empty decor deletes the key. Non-ashaar tags returned unchanged.
  function setTagSlotDecor(tag, key, decor) {
    var payload = parseContentControlTag(tag);
    if (!payload) return tag;
    var sd = payload.slotDecor && typeof payload.slotDecor === "object" ? payload.slotDecor : {};
    var clean = {};
    if (decor && typeof decor === "object") {
      if (decor.symbol) clean.symbol = decor.symbol;
      if (decor.fill) clean.fill = decor.fill;
      if (decor.color) clean.color = decor.color;
    }
    if (clean.symbol || clean.fill || clean.color) sd[key] = clean; else delete sd[key];
    payload.slotDecor = sd;
    return "ashaar:" + encodeURIComponent(JSON.stringify(payload));
  }
```
In `parseContentControlTag`, after the `overrides` normalization line add:
```js
      payload.slotDecor = (payload.slotDecor && typeof payload.slotDecor === "object") ? payload.slotDecor : {};
```
Export: add `setTagSlotDecor: setTagSlotDecor,`.

- [ ] **Step 5: Profile bucket**

In `src/taskpane/profiles.js` `defaultProfile`, add after `debugColors`:
```js
      spacingDecor: {},                            // { "<slot>": { symbol, fill, color } }
```
In `mergeProfile`, add `"spacingDecor"` to the `nested` array:
```js
    var nested = ["width", "justify", "debugColors", "fontCorrections", "derived", "spacingDecor"];
```

- [ ] **Step 6: Run to verify pass + full suite**

Run: `node tests/cell-overrides.test.js`; `node tests/profiles.test.js`; then `npm test`.
Expected: all green.

- [ ] **Step 7: Commit**

```bash
git add src/taskpane/cell-overrides.js src/taskpane/word-html.js src/taskpane/profiles.js tests/cell-overrides.test.js tests/profiles.test.js
git commit --no-gpg-sign -m "feat(decor): resolveSlotDecor + setTagSlotDecor + profile spacingDecor bucket"
```

---

### Task 2: Decorate pass in `applyProfileToQaseeda`

**Files:** Modify `src/taskpane/taskpane.js` (`applyProfileToQaseeda`).

**Interfaces:** Consumes `AshaarOverrides.resolveSlotDecor/overrideKey`, `profile.spacingDecor`, `parseContentControlTag(...).slotDecor`. Manual-verify.

**Design:** Spacing cells are currently skipped in the re-justify loop (`if (c.kind === "spacing") return;`). Instead, for a spacing cell resolve its decoration (profile default for its slot-position + the block's per-slot override) and write it: symbol text (+ font color) into the cell, and `shadingColor` for fill. Content cells are unchanged.

- [ ] **Step 1: Capture per-block slotDecor + each cell's slot**

Where `blockOverrides` is built (Task-2/SP2), add alongside:
```js
        var blockSlotDecor = blocks.map(function (cc) {
          var p = AshaarWord.parseContentControlTag(cc.tag);
          return (p && p.slotDecor) || {};
        });
        var allTableSlotDecor = [];
        blockTables.forEach(function (t, bi) {
          t.items.forEach(function () { allTableSlotDecor.push(blockSlotDecor[bi] || {}); });
        });
```
In the `tableInfos` cell capture, also record the spacing slot + its per-slot decor key. In the `cells.push({...})`, add:
```js
                slot: (mapped && mapped.kind === "spacing") ? mapped.slot : null,
                decorKey: (mapped && mapped.kind === "spacing" && mapped.slot)
                  ? AshaarOverrides.overrideKey(allTableBlockIdx[ai], mapped.slot) : null,
```
and attach the table's slotDecor to the returned info: `return { tbl: tbl, cells: cells, overrides: allTableOverrides[ai], slotDecor: allTableSlotDecor[ai] };`

- [ ] **Step 2: Decorate spacing cells instead of skipping**

In the re-justify loop, replace:
```js
            if (c.kind === "spacing") return;              // structural gap — never justified
            if (!c.base && c.kind !== "content") return;   // empty & not known-content → skip
            if (!c.base) return;
```
with:
```js
            if (c.kind === "spacing") {
              // Decorate (not justify) a structural gap: profile default for its
              // slot-position + the block's per-slot override.
              var pDecor = c.slot ? (profile.spacingDecor || {})[c.slot] : null;
              var oDecor = c.decorKey ? info.slotDecor[c.decorKey] : null;
              var decor = AshaarOverrides.resolveSlotDecor(pDecor, oDecor);
              c.cell.body.clear();
              if (decor.symbol) {
                c.cell.body.insertText(decor.symbol, Word.InsertLocation.replace);
                c.cell.body.font.color = decor.color || "black";
              }
              try { c.cell.shadingColor = decor.fill || "No color"; } catch (e) {}
              c.cell.body.paragraphs.getFirst().alignment = Word.Alignment.centered;
              changed++;
              return;
            }
            if (!c.base && c.kind !== "content") return;   // empty & not known-content → skip
            if (!c.base) return;
```
(`shadingColor = "No color"` clears any prior fill; wrap in try in case a host rejects the sentinel.)

- [ ] **Step 3: Syntax + suite + commit**

Run: `node --check src/taskpane/taskpane.js`; `npm test`; Expected: green.
```bash
git add src/taskpane/taskpane.js
git commit --no-gpg-sign -m "feat(decor): decorate pass fills spacing cells (symbol + fill) on apply"
```

---

### Task 3: Per-slot decoration editor + profile decor UI

**Files:** Modify `src/taskpane/taskpane.html`, `src/taskpane/taskpane.css`, `src/taskpane/taskpane.js`.

**Interfaces:** Consumes `AshaarCellMap.buildBandhCellMap`, `AshaarOverrides.overrideKey`, `AshaarWord.setTagSlotDecor`. Manual-verify.

- [ ] **Step 1: Editor markup**

In `taskpane.html`, after the `cell-override` panel, add:
```html
        <div id="slot-decor" class="cell-override" hidden>
          <div class="cell-override-head">Gap <span id="slot-decor-label"></span></div>
          <label>Symbol <input id="slot-decor-symbol" type="text" placeholder="e.g. ؎ (blank = none)"></label>
          <label>Fill color <input id="slot-decor-fill" type="color" value="#f5f0e0"><label class="adopt-check"><input type="checkbox" id="slot-decor-fill-on"> on</label></label>
          <label>Symbol color <input id="slot-decor-color" type="color" value="#a7352a"></label>
          <button id="slot-decor-clear" type="button" class="button--secondary">Clear gap decoration</button>
        </div>
```

- [ ] **Step 2: Show the decor editor for spacing cells**

In `reflectActiveCell`, the current code hides the editor for non-content cells. Extend the map-entry branch: when `entry.kind === "spacing"`, populate + show the decor editor and hide the justify editor; when content, the reverse. Replace the tail of `reflectActiveCell` (the `if (!entry || entry.kind !== "content")` guard and what follows) with:
```js
    var decorEl = document.getElementById("slot-decor");
    if (!entry) { editor.hidden = true; if (decorEl) decorEl.hidden = true; _activeOvKey = null; _activeDecorKey = null; return; }
    if (entry.kind === "content") {
      if (decorEl) decorEl.hidden = true;
      _activeDecorKey = null;
      _activeOvKey = AshaarOverrides.overrideKey(tIdx, entry.label);
      populateCellEditor(entry.label, (payload.overrides || {})[_activeOvKey]);
      editor.hidden = false;
    } else { // spacing
      editor.hidden = true;
      _activeOvKey = null;
      _activeDecorKey = AshaarOverrides.overrideKey(tIdx, entry.slot);
      populateDecorEditor(entry.slot, (payload.slotDecor || {})[_activeDecorKey]);
      if (decorEl) decorEl.hidden = false;
    }
```
Add module var `var _activeDecorKey = null;` near `_activeOvKey`.

- [ ] **Step 3: Decor editor populate/read/apply**

Add:
```js
  function hexToWord(v) { return (v || "").replace(/^#/, ""); }
  function populateDecorEditor(slot, d) {
    d = d || {};
    var lbl = document.getElementById("slot-decor-label"); if (lbl) lbl.textContent = slot || "";
    document.getElementById("slot-decor-symbol").value = d.symbol || "";
    document.getElementById("slot-decor-fill-on").checked = !!d.fill;
    if (d.fill) document.getElementById("slot-decor-fill").value = "#" + d.fill;
    if (d.color) document.getElementById("slot-decor-color").value = "#" + d.color;
  }
  function readDecorEditor() {
    var d = {};
    var sym = document.getElementById("slot-decor-symbol").value;
    if (sym) d.symbol = sym;
    if (document.getElementById("slot-decor-fill-on").checked) d.fill = hexToWord(document.getElementById("slot-decor-fill").value);
    if (sym) d.color = hexToWord(document.getElementById("slot-decor-color").value);
    return d;
  }
  async function applySlotDecor(clear) {
    if (!_activeDecorKey || typeof Word === "undefined") return;
    var d = clear ? null : readDecorEditor();
    _reflectBusy = true;
    var qname = "";
    try {
      await Word.run(async function (context) {
        var cc = context.document.getSelection().parentContentControlOrNullObject;
        cc.load("title,tag");
        await context.sync();
        if (cc.isNullObject || cc.title !== "Ashaar Poem") return;
        cc.tag = AshaarWord.setTagSlotDecor(cc.tag, _activeDecorKey, d);
        await context.sync();
        _lastBlockTag = cc.tag;
        qname = (AshaarWord.parseContentControlTag(cc.tag) || {}).qaseeda || "";
      });
    } catch (e) { /* ignore */ } finally { _reflectBusy = false; }
    if (clear) populateDecorEditor(document.getElementById("slot-decor-label").textContent, null);
    if (qname && loadProfileStore()[qname]) await applyProfileToQaseeda(qname); // re-decorate via the apply pass
  }
```
(Re-decorating routes through `applyProfileToQaseeda`, which now writes the gap decorations. It requires the block to belong to a saved qaseeda; if untagged, the tag is stored and takes effect on the next apply — acceptable for MVP.)

- [ ] **Step 4: Wire in `bind()`**

```js
    ["slot-decor-symbol", "slot-decor-fill", "slot-decor-color", "slot-decor-fill-on"].forEach(function (id) {
      var el = document.getElementById(id);
      if (el) el.addEventListener("change", function () { applySlotDecor(false); });
    });
    var decorClear = document.getElementById("slot-decor-clear");
    if (decorClear) decorClear.addEventListener("click", function () { applySlotDecor(true); });
```

- [ ] **Step 5: Profile decor in the cell-map view (read-back visibility)**

In `showCellMap`, when rendering a gap, show its resolved profile symbol if any: change the gap span to include `(payload.spacingDecor||{})`? The tag has no profile decor — profile lives in localStorage. Instead, annotate gaps with the block's per-slot decor from the tag: in the `map.forEach`, for a spacing entry with `payload.slotDecor["0:"+e.slot]` (table 0) show the symbol. Keep minimal — append the symbol when present:
```js
        rowHtml += e.kind === "content"
          ? "<span class=\"cell-map-cell\">" + e.label + "</span>"
          : "<span class=\"cell-map-gap\">(gap" + (((patterns._decor && patterns._decor[bi + ":" + e.slot]) ) ? " " + patterns._decor[bi + ":" + e.slot] : "") + ")</span>";
```
(Optional polish — if it complicates, skip; the editor already shows the active gap's decor.)

- [ ] **Step 6: Styles**

Reuse `.cell-override` styles (already present). No new CSS required.

- [ ] **Step 7: Syntax + suite + commit**

Run: `node --check src/taskpane/taskpane.js`; `npm test`; Expected: green.
```bash
git add src/taskpane/taskpane.html src/taskpane/taskpane.css src/taskpane/taskpane.js
git commit --no-gpg-sign -m "feat(decor): per-slot gap decoration editor (symbol + fill + color)"
```

---

### Task 4: Manual Word verification

**Files:** none. `npm start`.

- [ ] **Step 1 — profile default.** Save a qaseeda with a `spacingDecor` for a slot-position (e.g. `A#1` → symbol `؎`, fill). Apply → every bandh's `A#1` gap shows the symbol + shade.
- [ ] **Step 2 — per-slot override.** Cursor in a gap → the decoration editor shows its slot; set a different symbol/fill → only that gap changes after re-apply; **Clear** reverts to the profile default.
- [ ] **Step 3 — fill only / symbol only.** A gap with fill but no symbol shades with an empty paragraph; symbol but no fill shows the glyph, no shade.
- [ ] **Step 4 — justify still skips gaps.** Justify the block → decorated gaps keep their symbol/shade and receive NO tatweels/spaces.
- [ ] **Step 5 — content cells unaffected.** Content cells still justify normally; the decor editor never shows for a content cell (the justify-override editor does).
- [ ] **Step 6 — no-map block.** An adopted block (no `cells`) shows neither editor for gaps.

## Self-Review notes

- **Spec coverage:** §1 payload → Task 1; §2 config model (`resolveSlotDecor`, profile `spacingDecor`, tag `slotDecor`, key) → Task 1; §3 rendering — **realized via Office.js `shadingColor` + `insertText`/`font.color` in the decorate pass, not OOXML emitters** (insert bake-in deferred) → Task 2; §4 apply — decorate pass on Save & Apply / per-slot edit → Tasks 2/3 (insert bake-in explicitly deferred); §5 UI → Task 3; §6 testing → node Task 1 + manual Task 4.
- **Scope deviation (documented):** insert bake-in and the OOXML emitters (`spacingDecorParaXml`/`shdXml`/`tcXml`-shd) are **deferred** — the Office.js decorate pass covers the MVP without touching the generator, which is lower-risk. Freshly inserted poems show decorations after the first Save & Apply.
- **Type consistency:** decor `{symbol,fill,color}` and `resolveSlotDecor` return shape identical across Tasks 1–3; `slotDecor` key `"<ti>:<slot>"` via `overrideKey` in Tasks 2/3; `setTagSlotDecor(tag,key,decor)` in Tasks 1/3.
- **Reused unchanged:** SP1 `cells` map + `slot`/`kind`; SP2 `overrideKey`, active-cell detection, `applyProfileToQaseeda` block-walk; `Word.TableCell.shadingColor`.
- **Deferred:** auto-numbering, content-cell shading, presets, multi-glyph rules.
```

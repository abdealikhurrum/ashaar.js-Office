# Render Workflow & Batch Formatting Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Block-scoped Apply with visible refresh cost, a Re-render button, batch cell/gap formatting with capture-from-Word, artifact-safe font detection, probe/calibration caching, and vertical-rhythm settings (line height + table separator).

**Architecture:** Pure logic first, per module: canonical keys in `profiles.js`, OOXML emission in `word-html.js`, cache keys in a new `tune-cache.js`, cost labels in `settings-panel.js`, phase timing in a new `metrics.js`, batch key enumeration in `bandh-cell-map.js` — each with Node tests. Office.js glue in `taskpane.js` last per feature, browser-verified headlessly, Word-verified by the user.

**Tech Stack:** Vanilla ES5, UMD modules, Office.js v1, Node `assert` tests, no build step.

**Spec:** `docs/superpowers/specs/2026-07-13-render-workflow-and-batch-formatting-design.md` — read it before starting any task.

## Global Constraints

- ES5 only (`var`, `function`) in `src/taskpane/` (async functions are established in taskpane.js and OK there). UMD pattern for new modules. No new dependencies.
- Never edit `src/vendor/`.
- Canonical settings keys grow by exactly two: `lineHeightPt` (default `null`) and `separatorPt` (default `1`). All other keys unchanged: `justifyMode, fillMode, strength, gap, widthMode, widthPct, misraWidthPt, layoutMode, colWidthMode, capEm, fontCorrections, debugColors`.
- All tag payloads via `AshaarWord.parseContentControlTag` / setters — no ad-hoc JSON parsing of tags.
- New tag writes in Apply/action paths use `withWordStrict` + try/catch; failures keep pending; success/interim messages must never overwrite a later pipeline error (message-last pattern).
- `window.prompt`/`alert`/`confirm` are DISALLOWED (Office webview blocks modals) — inline UI only.
- Tests are plain Node scripts in the `npm test` chain (`package.json`); new test files must be appended to the chain.
- Commits: stage files EXPLICITLY by name. NEVER `git add -A`.
- **Working-tree caution:** the tree may hold user-owned uncommitted edits (MarkSafe font wiring) in `src/taskpane/fonts.js`, `src/taskpane/taskpane.css`, `tests/fonts.test.js`, `tests/word-html.test.js`. Before staging any of those four files, run `git diff <file>` and if it contains changes you did not make, STOP and report to the controller.
- Headless browser verification recipe (for taskpane.js tasks): `python3 -m http.server 3005` from repo root; Playwright MCP tools via ToolSearch; load `http://localhost:3005/src/taskpane/taskpane.html?v=<ts>`; cache-bust css/js by resetting href/src via browser_evaluate; mock `Office.context.document.settings` with an in-memory map when profile store access is needed; kill server + close browser after.
- Line numbers cited below WILL have drifted — locate by content; if the code doesn't match the description, report NEEDS_CONTEXT instead of guessing.

---

### Task 1: §9 canonical keys — `lineHeightPt` + `separatorPt` in profiles

**Files:**
- Modify: `src/taskpane/profiles.js`
- Modify: `tests/profiles-resolve.test.js` (append)

**Interfaces:**
- Produces: `defaultSettings()` gains `lineHeightPt: null` and `separatorPt: 1`; `settingsFromProfile(profile)` emits both when the profile carries them (profile schema fields `profile.lineHeightPt`, `profile.separatorPt`); `profileFromSettings(name, values)` maps them back. Resolver layering needs NO change (generic key loop).

- [ ] **Step 1: Write the failing test** — append to `tests/profiles-resolve.test.js` before the final `console.log`:

```js
// ── §9 vertical rhythm keys ──────────────────────────────────────────────────
{
  const d = defaultSettings();
  assert.strictEqual(d.lineHeightPt, null, "line height defaults to Word auto");
  assert.strictEqual(d.separatorPt, 1, "separator defaults to 1pt");

  const values = Object.assign(defaultSettings(), { lineHeightPt: 24, separatorPt: 6 });
  const p = profileFromSettings("V", values);
  assert.strictEqual(p.lineHeightPt, 24);
  assert.strictEqual(p.separatorPt, 6);
  const back = settingsFromProfile(p);
  assert.strictEqual(back.lineHeightPt, 24);
  assert.strictEqual(back.separatorPt, 6);

  // Layer through the resolver like any canonical key.
  const store = { V: p };
  const r = resolveSettings({ payload: { profile: "V", local: { separatorPt: 2 } }, profileStore: store, scope: { level: "poem" } });
  assert.strictEqual(r.values.lineHeightPt, 24, "profile layer");
  assert.strictEqual(r.values.separatorPt, 2, "local wins");
  assert.strictEqual(r.inherited.separatorPt, 6, "inherited = profile layer");
}
```

- [ ] **Step 2: Run to verify failure** — `node tests/profiles-resolve.test.js` — Expected: FAIL (`lineHeightPt` undefined vs null is OK for strictEqual? undefined !== null → fails at the first assert).

- [ ] **Step 3: Implement in `profiles.js`**

(a) In `defaultSettings()` add after `capEm: 0.28,`:

```js
      lineHeightPt: null,         // min line box (atLeast) in pt; null = Word auto
      separatorPt: 1,             // inter-table separator paragraph height (pt)
```

(b) In `settingsFromProfile` add before `return out;`:

```js
    if (p.lineHeightPt !== undefined) out.lineHeightPt = p.lineHeightPt;
    if (p.separatorPt != null) out.separatorPt = Number(p.separatorPt);
```

and in `normalizeProfile` ensure the two fields pass through unchanged if present (if `normalizeProfile` whitelists keys, add both; if it passes unknown keys through, no change — check and note which).

(c) In `profileFromSettings` add before `return p;`:

```js
    if (v.lineHeightPt !== undefined) p.lineHeightPt = v.lineHeightPt;
    if (v.separatorPt != null) p.separatorPt = Number(v.separatorPt);
```

- [ ] **Step 4: Run tests** — `node tests/profiles-resolve.test.js && node tests/profiles.test.js` then full `npm test`. Expected: green.

- [ ] **Step 5: Commit**

```bash
git add src/taskpane/profiles.js tests/profiles-resolve.test.js
git commit -m "feat(profiles): lineHeightPt + separatorPt canonical keys"
```

---

### Task 2: §9 OOXML emission — atLeast line spacing + exact-height separators

**Files:**
- Modify: `src/taskpane/word-html.js`
- Test: `tests/word-html.test.js` (append) — **check the working-tree caution in Global Constraints before staging this file**

**Interfaces:**
- Consumes: `opts.lineHeightPt` / `opts.separatorPt` (numbers or null/undefined) on the render opts already threaded into `renderForWordOoxml`/`misraParaXml`/`runsToMisraXml`.
- Produces: `AshaarWord.misraSpacingXml(opts)` → the `<w:spacing …/>` fragment used by ALL misra paragraph emitters; `AshaarWord.separatorParaXml(pt)` → the exact-height separator paragraph. Both exported for tests.

- [ ] **Step 1: Write the failing tests** — append to `tests/word-html.test.js`:

```js
// ── §9 vertical rhythm emission ──────────────────────────────────────────────
{
  assert.strictEqual(AshaarWord.misraSpacingXml({}), '<w:spacing w:after="80"/>', "auto when unset");
  assert.strictEqual(AshaarWord.misraSpacingXml({ lineHeightPt: 24 }),
    '<w:spacing w:after="80" w:line="480" w:lineRule="atLeast"/>', "atLeast, pt*20, never exact");

  const sep = AshaarWord.separatorParaXml(1);
  assert.ok(sep.indexOf('w:line="20"') !== -1 && sep.indexOf('w:lineRule="exact"') !== -1, "1pt exact");
  assert.ok(sep.indexOf('w:sz w:val="2"') !== -1, "1pt paragraph-mark font");
  assert.ok(sep.indexOf('w:before="0"') !== -1 && sep.indexOf('w:after="0"') !== -1, "no added spacing");

  // Emission goes through the shared helper in every misra paragraph.
  const xml = AshaarWord.renderForWordOoxml("الف \\ ب\n\nج \\ د", { lineHeightPt: 24, separatorPt: 2 }, Ashaar, 9360);
  assert.ok(xml.indexOf('w:line="480" w:lineRule="atLeast"') !== -1, "misra paragraphs carry line height");
  assert.ok(xml.indexOf('w:line="40" w:lineRule="exact"') !== -1, "tables joined by 2pt separator");
  assert.strictEqual(xml.indexOf("<w:p/>") === -1, true, "no bare separator paragraphs remain");
}
```

- [ ] **Step 2: Run to verify failure** — `node tests/word-html.test.js` — Expected: FAIL, `misraSpacingXml is not a function`.

- [ ] **Step 3: Implement**

(a) Add near `misraParaXml` and export both:

```js
  // Shared paragraph spacing for every misra paragraph. lineHeightPt (pt) emits
  // a minimum line box (atLeast) so tall nastaliq ink grows the line instead of
  // clipping — NEVER lineRule="exact". Also shields against host-doc Normal
  // styles with exact spacing.
  function misraSpacingXml(opts) {
    var pt = opts && Number(opts.lineHeightPt) > 0 ? Number(opts.lineHeightPt) : 0;
    return pt > 0
      ? '<w:spacing w:after="80" w:line="' + Math.round(pt * 20) + '" w:lineRule="atLeast"/>'
      : '<w:spacing w:after="80"/>';
  }

  // Inter-table separator: keeps adjacent tables from merging (the paragraph is
  // load-bearing) at a settable height. exact + tiny font ⇒ ~pt tall.
  function separatorParaXml(pt) {
    var p = Number(pt) > 0 ? Number(pt) : 1;
    return '<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="' + Math.round(p * 20) +
      '" w:lineRule="exact"/><w:rPr><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr></w:pPr></w:p>';
  }
```

(b) Replace the hardcoded `<w:spacing w:after="80"/>` inside the misra paragraph builders (`misraParaXml` ~line 1404, the two run-emitters ~1427/1443/1494 — every site that emits a MISRA paragraph `w:pPr`) with `misraSpacingXml(opts)`. The run-emitters must receive `opts` if they don't already — thread it, don't global it.

(c) `renderForWordOoxml`: change `return tables.join("<w:p/>");` to `return tables.join(separatorParaXml(opts.separatorPt));`.

(d) The rebuild-path empty paragraphs (~lines 1732 `w:after="200"` and 1750 `w:after="80"` — the `emptyPara` locals used when rebuilding blocks): replace their values with `separatorParaXml(opts.separatorPt)` where the paragraph separates TABLES; if one of them is a trailing/leading paragraph with a different purpose (read the surrounding code), leave it and note why in your report.

(e) Export `misraSpacingXml` and `separatorParaXml` from the return block.

- [ ] **Step 4: Run tests** — `node tests/word-html.test.js` then `npm test`. Some existing assertions may match on the old `<w:p/>` join — update them to the new separator (do not weaken what they assert about tables/rows).

- [ ] **Step 5: Commit**

```bash
git add src/taskpane/word-html.js tests/word-html.test.js
git commit -m "feat(word-html): atLeast misra line spacing + exact-height table separators"
```

---

### Task 3: §9 pane wiring — Advanced controls + options() mapping

**Files:**
- Modify: `src/taskpane/taskpane.html` (Advanced section), `src/taskpane/settings-panel.js` (SCOPE_FIELDS), `src/taskpane/taskpane.js` (`options()`)
- Test: `tests/settings-panel.test.js` (append)

**Interfaces:**
- Consumes: Task 1 keys; Task 2 emission (render opts `lineHeightPt`/`separatorPt`).
- Produces: `SCOPE_FIELDS.poem` includes `"lineHeightPt"` and `"separatorPt"`; `options()` output carries both from `panelValues()`.

- [ ] **Step 1: Failing test** — append to `tests/settings-panel.test.js`:

```js
// ── §9 keys are poem-scope panel fields ──────────────────────────────────────
{
  assert.ok(AshaarPanel.SCOPE_FIELDS.poem.indexOf("lineHeightPt") !== -1);
  assert.ok(AshaarPanel.SCOPE_FIELDS.poem.indexOf("separatorPt") !== -1);
}
```

- [ ] **Step 2: Run to verify failure** — `node tests/settings-panel.test.js`.

- [ ] **Step 3: Implement**
  - `settings-panel.js` `SCOPE_FIELDS.poem`: append `"lineHeightPt", "separatorPt"` (before `"fontCorrections"`).
  - `taskpane.html` Advanced `<details>` (after the misra-width field):

```html
            <div class="field"><label for="sp-line-height">Min line height (pt) <span class="sp-src" data-key="lineHeightPt"></span></label>
              <input id="sp-line-height" data-key="lineHeightPt" type="number" min="1" step="1" placeholder="auto"></div>
            <div class="field"><label for="sp-separator">Table separator (pt) <span class="sp-src" data-key="separatorPt"></span></label>
              <input id="sp-separator" data-key="separatorPt" type="number" min="1" step="1" value="1"></div>
```

  (Plain scalar `data-key` inputs — the generic change listener and provenance dots pick them up with zero JS.)
  - `taskpane.js` `options()`: add to the returned object:

```js
      lineHeightPt: v.lineHeightPt,
      separatorPt: v.separatorPt,
```

- [ ] **Step 4: Verify** — `npm test`; browser boot check (Global Constraints recipe): both inputs render in Advanced, changing them lights the dirty dot, ⟲ clears.

- [ ] **Step 5: Commit**

```bash
git add src/taskpane/settings-panel.js src/taskpane/taskpane.html src/taskpane/taskpane.js tests/settings-panel.test.js
git commit -m "feat(panel): line-height + separator controls wired through options()"
```

---

### Task 4: §7 strip before font determination

**Files:**
- Modify: `src/taskpane/fonts.js` (pure helper) — **working-tree caution applies: check `git diff src/taskpane/fonts.js` first; the MarkSafe edits may be present. If so STOP and report — the controller will decide (likely: proceed but stage carefully after user commits MarkSafe).**
- Modify: `src/taskpane/taskpane.js` (font-read sites)
- Test: `tests/fonts.test.js` (append) — same caution

**Interfaces:**
- Produces: `AshaarFonts.isArtifactRun(text)` → true when the text consists ONLY of justification artifacts (tatweel U+0640, hair space U+200A, thin space U+2009, ASCII space); `AshaarFonts.dominantRunFont(runs)` → the `font` (string) of the first non-artifact run, else the first run's font, else null. `runs` = `[{text, font}]`.

- [ ] **Step 1: Failing test** — append to `tests/fonts.test.js`:

```js
// ── §7 artifact-safe font determination ──────────────────────────────────────
assert.strictEqual(AshaarFonts.isArtifactRun("ــ"), true);
assert.strictEqual(AshaarFonts.isArtifactRun("   "), true);
assert.strictEqual(AshaarFonts.isArtifactRun("كلمة"), false);
assert.strictEqual(AshaarFonts.isArtifactRun(""), true);
// Word run in Jameel + tatweel run left in Arial (the observed regression):
assert.strictEqual(AshaarFonts.dominantRunFont([
  { text: "ــ", font: "Arial" },
  { text: "كلمة", font: "Jameel Noori Nastaleeq" },
]), "Jameel Noori Nastaleeq");
assert.strictEqual(AshaarFonts.dominantRunFont([{ text: "ـ", font: "Arial" }]), "Arial", "all-artifact falls back to first");
assert.strictEqual(AshaarFonts.dominantRunFont([]), null);
```

- [ ] **Step 2: Run to verify failure** — `node tests/fonts.test.js`.

- [ ] **Step 3: Implement in `fonts.js`** (export both):

```js
  var ARTIFACT_ONLY = /^[ـ   ]*$/;
  // True when a run is purely justification artifacts (tatweels / micro-spaces)
  // — such runs carry stale fonts after native font changes and must never
  // decide a word's font (observed: whole words falling back to Arial).
  function isArtifactRun(text) { return ARTIFACT_ONLY.test(String(text || "")); }
  function dominantRunFont(runs) {
    runs = runs || [];
    for (var i = 0; i < runs.length; i++) {
      if (!isArtifactRun(runs[i].text)) return runs[i].font || null;
    }
    return runs.length ? (runs[0].font || null) : null;
  }
```

- [ ] **Step 4: Use it in `taskpane.js`** — find every site that decides a cell's/word's font from run fonts: (a) the representative-font capture in `captureQaseedaTables` (per-word fonts of content cells, comment "Per-word fonts of every content cell"); (b) the generic per-run dispatch in `justifySelectionInner` where a run's font routes the mechanism (`descriptorForFontName(...)` on run font names); (c) `reRender`'s `existingFont` pick (first cell reporting one). At each: filter artifact runs out of the decision using `isArtifactRun`, or pick via `dominantRunFont`, WITHOUT changing what gets re-emitted (artifact runs are still stripped/regenerated downstream). Cite each changed site in your report. Additionally, in pipelines that re-render (Apply/Re-render), confirm `stripJustification` runs before run fonts are read for detection — if a path reads fonts pre-strip, reorder or filter and note it.

- [ ] **Step 5: Run everything** — `npm test` green; commit:

```bash
git add src/taskpane/fonts.js src/taskpane/taskpane.js tests/fonts.test.js
git commit -m "feat(fonts): justification artifacts never decide a word's font"
```

---

### Task 5: §8 probe/calibration caching

**Files:**
- Create: `src/taskpane/tune-cache.js`
- Create: `tests/tune-cache.test.js`; Modify: `package.json` (append to chain)
- Modify: `src/taskpane/taskpane.js` (probe/calibrate block ~2552-2581), `src/taskpane/taskpane.html` (script include before taskpane.js), `src/taskpane/font-store.js` or the font-registration handler in taskpane.js (cache bust hook)

**Interfaces:**
- Produces (UMD global `AshaarTuneCache`):
  - `probeKey(family, engineVersion)` → `"probe|<family>|<version>"`
  - `calibKey(family, sizePt, containerPx, texts)` → `"calib|<family>|<size>|<bucket>|<hash>"` where bucket = `Math.round(containerPx / 25) * 25` and hash = the same 32-bit string hash used elsewhere (`((h<<5)-h+c)|0` over the joined texts).
  - `makeCache(storage)` → `{ getProbe(family), putProbe(family, profile), getCalib(key), putCalib(key, params), bustAll() }`; `storage` is optional `{getItem,setItem,removeItem}` (localStorage-shaped) used ONLY for probes; calibration entries are in-memory only.
- Consumes: `ASHAAR_UPSTREAM_VERSION` (exists in the repo — grep for it; it identifies the vendored engine build) as `engineVersion`.

- [ ] **Step 1: Failing test** — create `tests/tune-cache.test.js`:

```js
const assert = require("assert");
const TC = require("../src/taskpane/tune-cache");

assert.strictEqual(TC.probeKey("Jameel", "v7"), "probe|Jameel|v7");
assert.strictEqual(TC.probeKey("Jameel", "v8") === TC.probeKey("Jameel", "v7"), false, "engine bump busts");

const k1 = TC.calibKey("Jameel", 14, 412, ["الف", "ب"]);
assert.strictEqual(TC.calibKey("Jameel", 14, 420, ["الف", "ب"]), k1, "same 25px bucket");
assert.notStrictEqual(TC.calibKey("Jameel", 14, 460, ["الف", "ب"]), k1, "different bucket");
assert.notStrictEqual(TC.calibKey("Jameel", 14, 412, ["الف", "ج"]), k1, "texts hash differs");

// storage-backed probes, in-memory calib
const mem = {};
const storage = { getItem: k => (k in mem ? mem[k] : null), setItem: (k, v) => { mem[k] = v; }, removeItem: k => { delete mem[k]; } };
const c = TC.makeCache(storage);
assert.strictEqual(c.getProbe("probe|J|v7"), null);
c.putProbe("probe|J|v7", { pairs: { AB: 1 } });
assert.deepStrictEqual(c.getProbe("probe|J|v7"), { pairs: { AB: 1 } });
assert.deepStrictEqual(TC.makeCache(storage).getProbe("probe|J|v7"), { pairs: { AB: 1 } }, "survives new instance via storage");
c.putCalib("k", { targetFill: 0.9 });
assert.deepStrictEqual(c.getCalib("k"), { targetFill: 0.9 });
assert.strictEqual(TC.makeCache(storage).getCalib("k"), null, "calib is in-memory only");
c.bustAll();
assert.strictEqual(c.getProbe("probe|J|v7"), null, "bust clears storage too");

console.log("tune-cache tests passed");
```

- [ ] **Step 2: Run to verify failure** — `node tests/tune-cache.test.js` — FAIL: cannot find module.

- [ ] **Step 3: Implement `src/taskpane/tune-cache.js`** (UMD like settings-panel.js; ES5):

```js
/**
 * tune-cache.js — memoization for AshaarTune probe/calibrate results.
 * probeFont depends only on (fontFamily, engine build) → persisted to storage.
 * calibrate depends on poem texts/width/font/size → in-memory only (texts churn).
 * Pure: no DOM, no Office.js. See spec §8.
 */
(function (root, factory) {
  if (typeof module !== "undefined" && module.exports) module.exports = factory();
  else root.AshaarTuneCache = factory();
}(typeof globalThis !== "undefined" ? globalThis : this, function () {
  "use strict";
  var PREFIX = "ashaar:fontProbe:";
  function hash32(s) {
    var h = 0; s = String(s || "");
    for (var i = 0; i < s.length; i++) h = ((h << 5) - h + s.charCodeAt(i)) | 0;
    return (h >>> 0).toString(16);
  }
  function probeKey(family, engineVersion) { return "probe|" + family + "|" + engineVersion; }
  function calibKey(family, sizePt, containerPx, texts) {
    var bucket = Math.round(Number(containerPx || 0) / 25) * 25;
    return "calib|" + family + "|" + Number(sizePt || 0) + "|" + bucket + "|" + hash32((texts || []).join(""));
  }
  function makeCache(storage) {
    var calib = {};
    function sKey(k) { return PREFIX + k; }
    return {
      getProbe: function (k) {
        if (!storage) return null;
        try { var raw = storage.getItem(sKey(k)); return raw ? JSON.parse(raw) : null; }
        catch (e) { return null; }
      },
      putProbe: function (k, profile) {
        if (!storage) return;
        try { storage.setItem(sKey(k), JSON.stringify(profile)); } catch (e) { /* quota — skip */ }
      },
      getCalib: function (k) { return (k in calib) ? calib[k] : null; },
      putCalib: function (k, params) { calib[k] = params; },
      bustAll: function () {
        calib = {};
        if (!storage) return;
        try {
          var kill = [], i;
          for (i = 0; storage.length != null && i < storage.length; i++) {
            var key = storage.key ? storage.key(i) : null;
            if (key && key.indexOf(PREFIX) === 0) kill.push(key);
          }
          if (!kill.length) { // storage without length/key (test shim): track nothing, best-effort
            kill = [];
          }
          kill.forEach(function (k) { storage.removeItem(k); });
        } catch (e) { /* ignore */ }
      },
    };
  }
  return { probeKey: probeKey, calibKey: calibKey, makeCache: makeCache, _hash32: hash32 };
}));
```

NOTE: the test's `bustAll` expectation uses the shim WITHOUT `length`/`key` — make `makeCache` track keys it wrote (`var written = {}` set in `putProbe`, cleared in `bustAll` via `removeItem`) instead of enumerating storage; adjust the implementation accordingly (this is the correct design — implement key-tracking, not enumeration).

- [ ] **Step 4: Integrate in `taskpane.js`** — module-level: `var _tuneCache = AshaarTuneCache.makeCache(typeof localStorage !== "undefined" ? localStorage : null);` In the probe/calibrate block of `justifySelectionInner` (~2552):
  - probe: `var pk = AshaarTuneCache.probeKey(repName, ASHAAR_UPSTREAM_VERSION); fontProfile = _tuneCache.getProbe(pk); if (!fontProfile) { …existing probeFont…; if (fontProfile) _tuneCache.putProbe(pk, fontProfile); }` (record hit/miss in the debug diags).
  - calibrate: build `var ck = AshaarTuneCache.calibKey(repName, repSize, avgPx, lineTexts); var cached = _tuneCache.getCalib(ck); if (cached) { calibParams = Object.assign({}, cached); } else { …existing calibrate…; _tuneCache.putCalib(ck, calibParams); }`.
  - Bust on font registration/replacement: in the fonts-strip "Add font" success handler, call `_tuneCache.bustAll();` (locate the `font-upload-add` handler).
  - Add `<script src="tune-cache.js"></script>` before `taskpane.js` in `taskpane.html`, and append `&& node tests/tune-cache.test.js` to the `test` script.

- [ ] **Step 5: Verify + commit** — `npm test` green; browser boot check (no console errors; `typeof AshaarTuneCache === "object"`).

```bash
git add src/taskpane/tune-cache.js tests/tune-cache.test.js package.json src/taskpane/taskpane.js src/taskpane/taskpane.html
git commit -m "feat(perf): probe results persisted per font+engine; calibration memoized per poem"
```

---

### Task 6: §6 refresh-cost labels

**Files:**
- Modify: `src/taskpane/settings-panel.js`, `src/taskpane/taskpane.html` (caption element), `src/taskpane/taskpane.js` (renderPanel)
- Test: `tests/settings-panel.test.js` (append)

**Interfaces:**
- Produces: `panelStateFor(...)` footer gains `costLabel` (string). Inputs it uses: `pending`, `target.scope.level`, `resolved.values.justifyMode`. Structural keys constant exported as `STRUCTURAL_KEYS = ["gap", "widthMode", "widthPct", "layoutMode", "colWidthMode", "separatorPt"]` (separator changes require a rebuild — it lives between tables).

- [ ] **Step 1: Failing test** — append to `tests/settings-panel.test.js`:

```js
// ── §6 refresh-cost labels ───────────────────────────────────────────────────
{
  const resolved = resolveSettings({ payload: { profile: "", local: {} }, profileStore: {}, scope: { level: "poem" } });
  const t = { kind: "block", scope: { level: "poem" }, cellEnabled: false, gapEnabled: false };
  const structural = AshaarPanel.panelStateFor({ resolved, pending: { set: { gap: 8 }, clear: [] }, target: t });
  assert.strictEqual(structural.footer.costLabel, "Apply — rebuilds poem tables");
  const light = AshaarPanel.panelStateFor({ resolved, pending: { set: { strength: 9 }, clear: [] }, target: t });
  assert.strictEqual(light.footer.costLabel, "Apply — re-justifies poem");
  const cellT = { kind: "block", scope: { level: "cell", key: "A2:3" }, cellEnabled: true, gapEnabled: false, cellLabel: "A2:3" };
  const cellResolved = resolveSettings({ payload: { profile: "", local: {} }, profileStore: {}, scope: { level: "cell", key: "A2:3" } });
  assert.strictEqual(AshaarPanel.panelStateFor({ resolved: cellResolved, pending: { set: {}, clear: [] }, target: cellT }).footer.costLabel,
    "Apply — re-justifies poem");
  // justifyMode none → unjustified suffix
  const noneResolved = resolveSettings({ payload: { profile: "", local: { justifyMode: "none" } }, profileStore: {}, scope: { level: "poem" } });
  assert.strictEqual(AshaarPanel.panelStateFor({ resolved: noneResolved, pending: { set: { gap: 8 }, clear: [] }, target: t }).footer.costLabel,
    "Apply — rebuilds poem tables (unjustified: Justification is None)");
  assert.deepStrictEqual(AshaarPanel.STRUCTURAL_KEYS, ["gap", "widthMode", "widthPct", "layoutMode", "colWidthMode", "separatorPt"]);
}
```

- [ ] **Step 2: Run to verify failure**, then **Step 3: implement** in `settings-panel.js`:

```js
  var STRUCTURAL_KEYS = ["gap", "widthMode", "widthPct", "layoutMode", "colWidthMode", "separatorPt"];

  function costLabelFor(pending, level, justifyMode) {
    var dirtyKeys = Object.keys(pending.set).concat(pending.clear);
    var structural = level === "poem" && dirtyKeys.some(function (k) { return STRUCTURAL_KEYS.indexOf(k) !== -1; });
    var label = structural ? "Apply — rebuilds poem tables" : "Apply — re-justifies poem";
    if (justifyMode === "none") label += " (unjustified: Justification is None)";
    return label;
  }
```

Call it in `panelStateFor` (`footer.costLabel = costLabelFor(pending, level, resolved.values.justifyMode)`), export `STRUCTURAL_KEYS`.

- [ ] **Step 4: Render it** — `taskpane.html`: add `<div id="sp-cost" class="sp-cost"></div>` inside the `sp-footer` div (before the buttons); CSS append `.sp-cost { font-size: 11px; opacity: 0.75; align-self: center; }` to `taskpane.css` **(working-tree caution — check `git diff src/taskpane/taskpane.css` first)**. `renderPanel` in taskpane.js: `document.getElementById("sp-cost").textContent = st.footer.costLabel; document.getElementById("sp-apply").title = st.footer.costLabel;`

- [ ] **Step 5: Verify + commit** — `npm test`; browser: caption updates when gap edited.

```bash
git add src/taskpane/settings-panel.js src/taskpane/taskpane.html src/taskpane/taskpane.css src/taskpane/taskpane.js tests/settings-panel.test.js
git commit -m "feat(panel): refresh-cost caption — rebuild vs re-justify visible before Apply"
```

---

### Task 7: §6 debug phase metrics

**Files:**
- Create: `src/taskpane/metrics.js`; Create: `tests/metrics.test.js`; Modify: `package.json`
- Modify: `src/taskpane/taskpane.js` (applyPanel router + reRender), `src/taskpane/taskpane.html` (script include)

**Interfaces:**
- Produces (UMD `AshaarMetrics`): `startRun(label)` → `{ phase(name), end(), report() }`. `phase("x")` closes the previous phase and opens "x"; `end()` closes the last; `report()` → `{ label, totalMs, phases: [{name, ms}] }`. Injectable clock for tests: `startRun(label, nowFn)`.

- [ ] **Step 1: Failing test** — `tests/metrics.test.js`:

```js
const assert = require("assert");
const M = require("../src/taskpane/metrics");
let t = 0; const now = () => t;
const run = M.startRun("apply", now);
t = 5; run.phase("tag write");
t = 25; run.phase("justify");
t = 100; run.end();
const r = run.report();
assert.strictEqual(r.label, "apply");
assert.strictEqual(r.totalMs, 100);
assert.deepStrictEqual(r.phases, [
  { name: "start", ms: 5 }, { name: "tag write", ms: 20 }, { name: "justify", ms: 75 },
]);
console.log("metrics tests passed");
```

- [ ] **Step 2: verify failure; Step 3: implement** (UMD, ES5; `nowFn` defaults to `Date.now`); **Step 4: integrate** — in `applyPanel`: when debug mode is on (`debugMode.checked`), `var run = AshaarMetrics.startRun("apply " + _panel.scopeLevel);` and call `run.phase(...)` at: after face gate, after tag write, after pipeline, after reflect; `run.end()` then append `JSON.stringify(run.report())` to the debug output (reuse `renderDebug` or append a line to `debugOutput.textContent`). Same wrap in `reRender` (label "re-render"). Probe/calibrate hit/miss from Task 5 diags appear in the same debug dump. Script include + npm chain.

- [ ] **Step 5: Verify + commit** — `npm test`; browser boot clean.

```bash
git add src/taskpane/metrics.js tests/metrics.test.js package.json src/taskpane/taskpane.js src/taskpane/taskpane.html
git commit -m "feat(debug): per-phase Apply/Re-render timing in debug mode"
```

---

### Task 8: §1 cascade descope

**Files:**
- Modify: `src/taskpane/taskpane.js` only.

**Interfaces:**
- Produces: `applyProfileToQaseeda(name, opts)` accepts optional `opts.onlyBlockTag` (string) — when set, `gatherQaseedaBlocks` result is filtered to blocks whose `cc.tag === opts.onlyBlockTag` (BOTH passes). The Apply-path delegation in `justifySelectionInner` (head: `var qname = await getQaseedaAtSelection(); if (qname && …) { await applyProfileToQaseeda(qname); return; }` ~line 2255) passes the current block's tag: capture it in `getQaseedaAtSelection` or re-read the enclosing CC's tag and call `applyProfileToQaseeda(qname, { onlyBlockTag: tag })`. Assign/Update/Restore call sites stay unfiltered (profile-wide is their meaning).

- [ ] **Step 1: Implement** — thread `opts` through `applyProfileToQaseeda`; in BOTH passes, after `gatherQaseedaBlocks(context, name)`, add:

```js
        if (opts && opts.onlyBlockTag) {
          blocks = blocks.filter(function (b) { return b.tag === opts.onlyBlockTag; });
        }
```

(`gatherQaseedaBlocks` already loads `items/tag`.) Update the summary message for the filtered case ("Applied to this poem."). In `justifySelectionInner`'s delegation head, obtain the enclosing CC tag (the function already locates the block — reuse its lookup, or `getQaseedaAtSelection` can return `{name, tag}`; pick the smaller change and document it) and pass `{ onlyBlockTag: tag }`.

- [ ] **Step 2: Audit call sites** — `grep -n "applyProfileToQaseeda(" src/taskpane/taskpane.js`: `assignProfile`, `updateProfile`, `saveAsProfile` keep NO filter. Any other Apply-path caller gets the filter. List each call site + decision in your report.

- [ ] **Step 3: Verify** — `npm test` (no node coverage of this glue — must stay green anyway). Browser boot check. This task's real verification is the user's Word checklist item ("sibling poems no longer refresh").

- [ ] **Step 4: Commit**

```bash
git add src/taskpane/taskpane.js
git commit -m "feat(apply): block-scoped apply — profile-wide refresh only on Assign/Update/Restore"
```

---

### Task 9: §2 Re-render button

**Files:**
- Modify: `src/taskpane/taskpane.html` (footer), `src/taskpane/taskpane.js` (bind + enable state + metrics label)

- [ ] **Step 1: HTML** — in the `sp-footer` div, before `sp-apply`:

```html
          <button id="sp-rerender" type="button" class="button--secondary" title="Re-render — rebuilds poem tables" disabled>Re-render</button>
```

- [ ] **Step 2: Wire** — `bind()`: `document.getElementById("sp-rerender").addEventListener("click", reRender);`. In `renderPanel`: `document.getElementById("sp-rerender").disabled = !( _panel.target && _panel.target.kind === "block");`

- [ ] **Step 3: Verify** — browser: button disabled at boot ("Selection" target); `npm test` green.

- [ ] **Step 4: Commit**

```bash
git add src/taskpane/taskpane.html src/taskpane/taskpane.js
git commit -m "feat(panel): Re-render button — rebuild current poem after native text/font edits"
```

---

### Task 10: §3 apply-to-all toggle

**Files:**
- Modify: `src/taskpane/bandh-cell-map.js` (pure enumeration), `src/taskpane/taskpane.html` (target selector), `src/taskpane/taskpane.js` (applyPanel cell/gap branches)
- Test: `tests/bandh-cell-map.test.js` (append)

**Interfaces:**
- Consumes: the cell map produced by `buildBandhCellMap(pattern)` (cells carry `kind` "content"|"spacing", labels/slots) and `AshaarOverrides.overrideKey(tableIndex, label)`.
- Produces: `AshaarBandhCellMap.keysForTarget(map, kind, mode, currentKey)` → array of override/decor keys. `kind` ∈ "content"|"spacing"; `mode` ∈ "this"|"bandh"|"poem". For "this" → `[currentKey]`. For "bandh"/"poem": enumerate the map's cells of that kind and produce keys via the SAME key scheme the current detection uses (`overrideKey(tableIndex, label)` for content, decor keys for spacing) — "bandh" limits to the current key's table index. **Read `buildBandhCellMap`'s actual output shape first and mirror the key derivation used in `reflectActiveCell` — do not invent labels. If the map lacks table indices, take a `tables` param; report the exact signature you land on.**

- [ ] **Step 1: Write the failing test** — append to `tests/bandh-cell-map.test.js` a block that builds a two-bandh pattern with content and spacing cells (reuse the fixtures already in that test file — read them first), then asserts: "this" returns exactly the current key; "bandh" returns all keys of that kind in the current table only; "poem" returns all tables' keys; content and spacing enumerations don't mix.

- [ ] **Step 2: verify failure; Step 3: implement; Step 4: `npm test`** (adapt assertions to the real shapes, never weaken the this/bandh/poem containment properties).

- [ ] **Step 5: UI + Apply routing** — `taskpane.html`: add to BOTH `sp-body-cell` and `sp-body-gap`:

```html
          <div class="field"><label>Apply to</label>
            <select id="sp-cell-target"><option value="this" selected>This cell</option><option value="bandh">This bandh</option><option value="poem">Whole poem</option></select></div>
```

(id `sp-gap-target` with "This gap" wording in the gap body). `applyPanel` cell branch: replace the single `setTagOverride(cc.tag, target.cellLabel, …)` with a loop over `keysForTarget(...)` re-encoding the tag once per key (setters compose — feed each output tag into the next call; ONE `cc.tag =` assignment and ONE `context.sync()` at the end). Same for the gap branch with `setTagSlotDecor`. ⟲-cleared fields already produce nulls via `dirtyOrNull` — nulls clear on every targeted key (setTagOverride deletes all-null overrides).

- [ ] **Step 6: Verify + commit** — `npm test`; browser boot; Word verification is a checklist item.

```bash
git add src/taskpane/bandh-cell-map.js tests/bandh-cell-map.test.js src/taskpane/taskpane.html src/taskpane/taskpane.js
git commit -m "feat(panel): apply cell/gap formatting to this / bandh / whole poem"
```

---

### Task 11: §4 cell fill + text color

**Files:**
- Modify: `src/taskpane/word-html.js` (`setTagOverride` accepts `fill`/`color`), `src/taskpane/taskpane.html` (cell body controls), `src/taskpane/taskpane.js` (dirtyOrNull keys, render application)
- Test: `tests/word-html.test.js` (append) + `tests/cell-overrides.test.js` if it pins the override shape

**Interfaces:**
- Produces: `setTagOverride(tag, key, {strength, widthPt, capEm, fill, color})` — `fill`/`color` persist like the others; all-null/empty still deletes the override. Cell overrides resolve as today (cell layer); `resolveSettings` cell branch passes `fill`/`color` through `values`/`source` ONLY if trivial — otherwise cell fill/color bypass the resolver and are read from `payload.overrides[key]` directly at render time (spec: "resolve at cell scope only"). Pick the simpler wiring, document it.

- [ ] **Step 1: Failing test** — append to `tests/word-html.test.js`:

```js
// ── §4 cell fill/color overrides ─────────────────────────────────────────────
{
  const base = AshaarWord.contentControlTag("متن", {});
  const t1 = AshaarWord.setTagOverride(base, "A2:3", { strength: null, widthPt: null, capEm: null, fill: "#F5F0E0", color: "#A7352A" });
  const p1 = AshaarWord.parseContentControlTag(t1);
  assert.deepStrictEqual(p1.overrides["A2:3"], { fill: "#F5F0E0", color: "#A7352A" });
  const t2 = AshaarWord.setTagOverride(t1, "A2:3", { strength: null, widthPt: null, capEm: null, fill: null, color: null });
  assert.strictEqual("A2:3" in AshaarWord.parseContentControlTag(t2).overrides, false, "all-null deletes");
}
```

- [ ] **Step 2: verify failure; Step 3: implement** — in `setTagOverride`, extend `has` and `clean` with `fill` and `color` (same `!= null` pattern; empty string counts as unset — use `override.fill != null && override.fill !== ""`).

- [ ] **Step 4: UI** — `sp-body-cell` gains (mirroring the gap body, no symbol):

```html
          <div class="field"><label>Fill color</label>
            <input id="sp-cell-fill" type="color" value="#f5f0e0">
            <label class="adopt-check"><input type="checkbox" id="sp-cell-fill-on"> on</label></div>
          <div class="field"><label>Text color</label>
            <input id="sp-cell-color" type="color" value="#a7352a">
            <label class="adopt-check"><input type="checkbox" id="sp-cell-color-on"> on</label></div>
```

`applyPanel` cell branch passes `fill: cellFillOn.checked ? cellFill.value : null, color: cellColorOn.checked ? cellColor.value : null` into the override object (alongside `dirtyOrNull` values).

- [ ] **Step 5: Render application** — in `applyProfileToQaseeda` pass 2, the CONTENT-cell branch (the `else` of the `c.kind === "spacing"` decor block ~1429): after justify, apply `var ov = info.overrides[c.key] || {}; c.cell.shadingColor = ov.fill || "#FFFFFF"; if (ov.color) c.cell.body.font.color = ov.color;` — note the `"#FFFFFF"`-to-clear quirk is already documented at the spacing branch; mirror it. Do the same in `justifySelectionInner`'s per-cell loop (locate where `ccOverrides` per-cell values are already consumed). Cite both sites.

- [ ] **Step 6: Verify + commit** — `npm test` (update `cell-overrides.test.js` if it whitelists override fields — extend, don't weaken); browser boot.

```bash
git add src/taskpane/word-html.js src/taskpane/taskpane.html src/taskpane/taskpane.js tests/word-html.test.js tests/cell-overrides.test.js
git commit -m "feat(cells): per-cell fill + text color overrides"
```

---

### Task 12: §5 capture from Word

**Files:**
- Modify: `src/taskpane/taskpane.html` (buttons), `src/taskpane/taskpane.js` (handler)

**Interfaces:**
- Consumes: `_panel.target` (cellLabel/gapKey), `_activeSlot`, `AshaarPanel.mergePending`, `refreshPanel`.
- Produces: `captureCellFormatting()` — reads the cursor cell's `shadingColor`, `font.color`, and (gap scope) its text; normalizes; merges into `_panel.pending` + updates the scope body's color inputs; never writes to the document.

- [ ] **Step 1: HTML** — add to BOTH `sp-body-cell` and `sp-body-gap`:

```html
          <button id="sp-cell-capture" type="button" class="button--secondary" title="Read this cell's formatting into the pane — Apply to persist">Capture formatting</button>
```

(id `sp-gap-capture` in the gap body).

- [ ] **Step 2: Implement** — one handler for both scopes:

```js
  // §5: read the cursor cell's native formatting into the pane as pending
  // values. Never writes. Normalizes Word's no-color/automatic quirks:
  // shadingColor of "#FFFFFF"/""/null ⇒ no fill; font.color of "" /
  // "Automatic" / "#000000"-as-auto ⇒ inherit (empty).
  async function captureCellFormatting() {
    if (typeof Word === "undefined") { setMessage("Open this task pane inside Word to capture."); return; }
    try {
      await Word.run(async function (context) {
        var sel = context.document.getSelection();
        var tcell = sel.parentTableCellOrNullObject;
        tcell.load("shadingColor,body/text");
        tcell.body.font.load("color");
        await context.sync();
        if (tcell.isNullObject) { setMessage("Click inside a table cell first."); return; }
        var fill = tcell.shadingColor;
        if (!fill || /^#?F{6}$/i.test(String(fill).replace("#", "")) || String(fill).toLowerCase() === "auto") fill = "";
        var color = tcell.body.font.color;
        if (!color || String(color).toLowerCase() === "automatic" || String(color).toLowerCase() === "auto") color = "";
        var isGap = _panel.scopeLevel === "gap";
        var fillEl = document.getElementById(isGap ? "sp-gap-fill" : "sp-cell-fill");
        var fillOn = document.getElementById(isGap ? "sp-gap-fill-on" : "sp-cell-fill-on");
        var colorEl = document.getElementById(isGap ? "sp-gap-color" : "sp-cell-color");
        if (fill) { fillEl.value = fill; }
        fillOn.checked = !!fill;
        if (color) colorEl.value = color;
        if (isGap) {
          var sym = (tcell.body.text || "").trim();
          document.getElementById("sp-gap-symbol").value = sym;
        } else {
          var colorOn = document.getElementById("sp-cell-color-on");
          colorOn.checked = !!color;
        }
        setMessage("Captured — Apply to persist" + (_panel.scopeLevel === "cell" ? " (choose Apply-to target first)." : "."));
      });
    } catch (e) {
      setMessage("Capture failed: " + (e && e.message ? e.message : e));
    }
  }
```

Wire both buttons to it in `bind()`. NOTE: `parentTableCellOrNullObject` + the load list must match this codebase's proven patterns — check how `reflectActiveCell` reads the current cell and reuse ITS lookup instead if it differs; cite what you used.

- [ ] **Step 3: Verify** — `npm test`; browser with mocked `Word.run` + a fake cell object: capture populates the inputs and message; no document writes (mock records calls). Word verification is a checklist item.

- [ ] **Step 4: Commit**

```bash
git add src/taskpane/taskpane.html src/taskpane/taskpane.js
git commit -m "feat(panel): capture native cell formatting into pending"
```

---

### Task 13: Manual checklist + docs

**Files:**
- Modify: `docs/superpowers/specs/2026-07-12-unified-settings-manual-checklist.md` (append a "Render workflow & batch formatting" section), `CLAUDE.md` (panel flow note)

- [ ] **Step 1: Append checklist items** (one `- [ ]` each): sibling poems untouched by cell/gap/poem Apply on a shared profile; Re-render picks up native text edit + font change; apply-to-all (bandh, poem) for cell strength and gap symbol; cell fill/text color render + clear; capture → apply → update-profile round trip; line height stops nastaliq clipping (set ~1.8× font size on a clipping poem); separator 1pt — tables nearly touch, never merge, raising the setting spreads them; cost caption matches what actually happens (rebuild vs re-justify); debug metrics table appears with phase timings + probe/calib hit-miss; second Apply of an unchanged poem is visibly faster (calibration memo hit).

- [ ] **Step 2: CLAUDE.md** — in the Settings-panel flow block, add one sentence: Re-render button, apply-to-all targets, capture, cost caption, and the two vertical-rhythm keys.

- [ ] **Step 3: Commit**

```bash
git add docs/superpowers/specs/2026-07-12-unified-settings-manual-checklist.md CLAUDE.md
git commit -m "docs: manual checklist + CLAUDE.md for render workflow & batch formatting"
```

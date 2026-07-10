# Nastaliq Kashida Fonts Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add first-class support for the Nastaliq fonts that can actually fill a poetry line — **Mehr Nastaliq** (tatweel), **Jameel Noori Kasheeda** (italic-run), **Gulzar** (whitespace) — with the justify engine picking the right mechanism per font.

**Architecture:** A pure font registry (`fonts.js`) tags each font with a `mechanism`. All the scattered `Noto Nastaliq Urdu` literals read from it. Justify dispatch branches on the tag: Mehr rides the existing canvas-tatweel engine with a per-font tatweel whitelist injected via `params.priorityTable` (no vendor edit); Jameel gets a new add-in module `kashida-italic.js` doing discrete connected-segment (fasl) subset-selection emitted as per-run italic OOXML; Gulzar/Noto are guarded off the tatweel path entirely (they shatter under injected tatweels) and fall back to spacing.

**Tech Stack:** Vanilla ES5 / UMD modules (no build step, no transpilation), Office.js v1, Node `assert` tests. Runs in Word's WebView (WKWebView on Mac, WebView2 on Windows) and Node for tests.

## Global Constraints

- **No build step / ES5 only.** Every module uses the repo UMD wrapper (`module.exports` for Node, `root.X` for browser). No `import`/`export`, no arrow-only APIs that break ES5 parse in old WebViews.
- **Never edit `src/vendor/*`.** The vendored `ashaar-justify.js` / `ashaar-autotune.js` / `ashaar.js` are synced from the submodule; all new behavior lives in the add-in layer (`src/taskpane/*`).
- **Every new pure module gets a Node test file wired into the `npm test` chain** in `package.json` (`&&`-joined list).
- **Mehr tatweel whitelist (Beta 2.0), verbatim:** medial/initial into `ب پ ت ٹ ث س ش ف ک گ`; word-final into `ب پ ت ٹ ث ف ک گ` (no س ش word-final).
- **Bundled-font licensing:** Mehr = CC-BY-SA (attribution required in `README.md`/`LICENSE`); Gulzar = OFL (notice required); **Jameel Kasheeda = private**, must be excluded from the public GitHub Pages deploy.
- **Reader-end policy:** set the run font name; the pane documents that readers need the font installed. No `.docx` font embedding (Office.js can't).
- **Justification chooser is the CURRENT one** (`kashida`/`css`/`spacing`/`none`); the guided-justification §5 four-mode reframe is a separate, unbuilt project. This plan wires against the current chooser plus a minimal per-font safety guard.

---

## Manual Gate (BLOCKING — do before Task 5)

### Gate G: Jameel Word-italic swap verification

Jameel Kasheeda's elongated forms live in the **italic style slot**. Whether Microsoft Word (not just InPage) triggers the kasheeda glyphs when a run's italic property is set is **unverified** and gates the entire `italic-run` mechanism (Task 5).

- [ ] **G.1** Install Jameel Noori Nastaleeq Kasheeda on a Mac and a Windows machine with Word.
- [ ] **G.2** In Word, type an Urdu line in Jameel Kasheeda. Duplicate it. Set the duplicate's font to italic (⌘I / Ctrl+I).
- [ ] **G.3** Observe: does the italic line show **elongated kasheeda connectors** (PASS) or merely **slanted/oblique** normal glyphs (FAIL)?
- [ ] **G.4** Record the result in `docs/superpowers/specs/2026-07-10-nastaliq-kashida-fonts-design.md` under a new "Gate G result" line, on both platforms.

**Decision:**
- **PASS (either/both platforms):** proceed with Task 5 as written; note any platform that failed in the reader-end note (Task 6).
- **FAIL (both platforms):** **cut Task 5.** Reclassify Jameel `mechanism:"whitespace"` in the registry (Task 1) and skip `kashida-italic.js`. Jameel remains a selectable render-only font. Tasks 1–4, 6 proceed unchanged.

---

## Task 1: Font registry module (`fonts.js`)

**Files:**
- Create: `src/taskpane/fonts.js`
- Test: `tests/fonts.test.js`
- Modify: `package.json` (add test to the `npm test` chain)

**Interfaces:**
- Produces:
  - `AshaarFonts.LIST` — object keyed by font id.
  - `AshaarFonts.get(id) → descriptor|null`
  - `AshaarFonts.mechanismOf(id) → "tatweel"|"italic-run"|"whitespace"` (default `"whitespace"` for unknown/plain ids)
  - `AshaarFonts.wordNameOf(id) → string|null` (the `<w:rFonts w:cs>` name; `null` when the font mode adds no cs font, e.g. `document`)
  - `AshaarFonts.cssFamilyOf(id) → string|null` (CSS `font-family` value, or `null` for `document`)
  - `AshaarFonts.tatweelRulesOf(id) → {medialInto:string[], finalInto:string[]}|null`

- [ ] **Step 1: Write the failing test**

Create `tests/fonts.test.js`:
```js
"use strict";
const assert = require("assert");
const AshaarFonts = require("../src/taskpane/fonts");

// mechanism tags
assert.strictEqual(AshaarFonts.mechanismOf("mehr"), "tatweel");
assert.strictEqual(AshaarFonts.mechanismOf("jameel"), "italic-run");
assert.strictEqual(AshaarFonts.mechanismOf("gulzar"), "whitespace");
assert.strictEqual(AshaarFonts.mechanismOf("noto"), "whitespace");
// unknown / plain modes default to whitespace (never the tatweel engine)
assert.strictEqual(AshaarFonts.mechanismOf("document"), "whitespace");
assert.strictEqual(AshaarFonts.mechanismOf("nope"), "whitespace");

// Word cs names line up with what callers emit
assert.strictEqual(AshaarFonts.wordNameOf("mehr"), "Mehr Nastaliq Web");
assert.strictEqual(AshaarFonts.wordNameOf("gulzar"), "Gulzar");
assert.strictEqual(AshaarFonts.wordNameOf("jameel"), "Jameel Noori Nastaleeq");
assert.strictEqual(AshaarFonts.wordNameOf("document"), null);

// css families
assert.ok(/Mehr Nastaliq Web/.test(AshaarFonts.cssFamilyOf("mehr")));
assert.strictEqual(AshaarFonts.cssFamilyOf("document"), null);

// Mehr whitelist verbatim
const r = AshaarFonts.tatweelRulesOf("mehr");
assert.deepStrictEqual(r.medialInto, ["ب","پ","ت","ٹ","ث","س","ش","ف","ک","گ"]);
assert.deepStrictEqual(r.finalInto,  ["ب","پ","ت","ٹ","ث","ف","ک","گ"]);
assert.strictEqual(AshaarFonts.tatweelRulesOf("gulzar"), null);

console.log("fonts tests passed");
```

- [ ] **Step 2: Run test to verify it fails**

Run: `node tests/fonts.test.js`
Expected: FAIL — `Cannot find module '../src/taskpane/fonts'`.

- [ ] **Step 3: Write minimal implementation**

Create `src/taskpane/fonts.js`:
```js
/**
 * AshaarFonts — single source of truth for the fonts the add-in offers.
 * Each descriptor carries the CSS family (preview), the Word cs font name
 * (OOXML <w:rFonts w:cs>), and the kashida `mechanism` that selects the
 * justify strategy. Pure (no DOM); safe to require in Node tests.
 */
(function (root, factory) {
  if (typeof module !== "undefined" && module.exports) module.exports = factory();
  else root.AshaarFonts = factory();
}(typeof globalThis !== "undefined" ? globalThis : this, function () {
  "use strict";

  var LIST = {
    document: { id: "document", label: "Document default",
      css: null, wordName: null, mechanism: "whitespace", bundled: false },
    "arabic-serif": { id: "arabic-serif", label: "Arabic serif",
      css: "'Scheherazade New','Amiri','Times New Roman',serif",
      wordName: "Scheherazade New", mechanism: "whitespace", bundled: false },
    noto: { id: "noto", label: "Noto Nastaliq Urdu",
      css: "'Noto Nastaliq Urdu',serif", wordName: "Noto Nastaliq Urdu",
      mechanism: "whitespace", bundled: false },
    mehr: { id: "mehr", label: "Mehr Nastaliq",
      css: "'Mehr Nastaliq Web','Noto Nastaliq Urdu',serif", wordName: "Mehr Nastaliq Web",
      mechanism: "tatweel", bundled: true, file: "MehrNastaliqWeb.woff2",
      tatweelRules: {
        version: "beta-2.0",
        medialInto: ["ب","پ","ت","ٹ","ث","س","ش","ف","ک","گ"],
        finalInto:  ["ب","پ","ت","ٹ","ث","ف","ک","گ"]
      } },
    jameel: { id: "jameel", label: "Jameel Noori Kasheeda",
      css: "'Jameel Noori Nastaleeq Kasheeda','Jameel Noori Nastaleeq',serif",
      wordName: "Jameel Noori Nastaleeq",
      mechanism: "italic-run", bundled: true, private: true,
      file: "JameelNooriNastaleeqKasheeda.ttf" },
    gulzar: { id: "gulzar", label: "Gulzar",
      css: "'Gulzar',serif", wordName: "Gulzar",
      mechanism: "whitespace", bundled: true, file: "Gulzar-Regular.woff2" }
  };

  function get(id) { return LIST[id] || null; }
  function mechanismOf(id) { var d = get(id); return d ? d.mechanism : "whitespace"; }
  function wordNameOf(id) { var d = get(id); return d && d.wordName ? d.wordName : null; }
  function cssFamilyOf(id) { var d = get(id); return d && d.css ? d.css : null; }
  function tatweelRulesOf(id) { var d = get(id); return d && d.tatweelRules ? d.tatweelRules : null; }

  return { LIST: LIST, get: get, mechanismOf: mechanismOf, wordNameOf: wordNameOf,
    cssFamilyOf: cssFamilyOf, tatweelRulesOf: tatweelRulesOf };
}));
```

- [ ] **Step 4: Run test to verify it passes**

Run: `node tests/fonts.test.js`
Expected: PASS — prints `fonts tests passed`.

- [ ] **Step 5: Wire into `npm test`**

In `package.json`, append ` && node tests/fonts.test.js` to the end of the `"test"` script string.
Run: `npm test`
Expected: all suites pass, including `fonts tests passed`.

- [ ] **Step 6: Commit**

```bash
git add src/taskpane/fonts.js tests/fonts.test.js package.json
git commit -m "feat(fonts): AshaarFonts registry with per-font kashida mechanism

Co-Authored-By: Claude Opus 4.8 (1M context) <noreply@anthropic.com>"
```

---

## Task 2: Route the scattered font literals through the registry

Replace the five hard-coded `Noto Nastaliq Urdu`/`Jameel` sites so all font naming flows from `AshaarFonts`. Behavior for existing modes (`nastaliq`→now `noto`, `arabic-serif`) is preserved; new modes become reachable.

**Files:**
- Modify: `src/taskpane/word-html.js:140-145` (`fontFamilyStyle`), `src/taskpane/word-html.js:1093-1094` (`misraParaXml` rFonts)
- Modify: `src/taskpane/word-tabstop.js:65-72` (`runPropsXml`)
- Modify: `src/taskpane/taskpane.js:255-259` (`previewFontFamily`)
- Modify: `src/taskpane/taskpane.html` (UMD load order: add `./fonts.js` before `word-html.js`)
- Test: `tests/word-html.test.js`, `tests/word-tabstop.test.js` (add assertions)

**Interfaces:**
- Consumes: `AshaarFonts.cssFamilyOf`, `AshaarFonts.wordNameOf` (Task 1).
- Produces: no new API; `fontMode` values now include `mehr`/`gulzar`/`jameel`/`noto`.

**Note on the UMD require:** `word-html.js`/`word-tabstop.js` already use the UMD `factory(require("../vendor/ashaar-justify"))` pattern. Add `AshaarFonts` as a factory dependency the same way (Node `require("./fonts")`; browser `root.AshaarFonts`).

- [ ] **Step 1: Write the failing test**

In `tests/word-html.test.js`, add (after existing requires/tests):
```js
// Registry-driven fonts: Mehr/Gulzar reach the OOXML cs name and preview stack.
{
  const AshaarFonts = require("../src/taskpane/fonts");
  assert.strictEqual(AshaarFonts.wordNameOf("mehr"), "Mehr Nastaliq Web");
  // fontFamilyStyle now delegates to the registry
  const mehrCss = AshaarWord.fontFamilyStyle({ fontMode: "mehr" });
  assert.ok(/Mehr Nastaliq Web/.test(mehrCss), "mehr css from registry");
  const gulzarCss = AshaarWord.fontFamilyStyle({ fontMode: "gulzar" });
  assert.ok(/Gulzar/.test(gulzarCss), "gulzar css from registry");
  // legacy "nastaliq" alias still resolves to a Nastaliq face (Noto)
  const notoCss = AshaarWord.fontFamilyStyle({ fontMode: "nastaliq" });
  assert.ok(/Noto Nastaliq Urdu/.test(notoCss), "nastaliq alias preserved");
  console.log("word-html registry-font tests passed");
}
```

(`fontFamilyStyle` must be exported from `AshaarWord` for this test — Step 3 adds it to the returned object if not already present.)

- [ ] **Step 2: Run test to verify it fails**

Run: `node tests/word-html.test.js`
Expected: FAIL — `AshaarWord.fontFamilyStyle is not a function` or Mehr css assertion fails (literal-based code returns `""` for `mehr`).

- [ ] **Step 3: Write minimal implementation**

In `src/taskpane/word-html.js`, add `AshaarFonts` to the UMD factory args (mirror the existing `AshaarJustify` dependency wiring at the top/bottom of the file), then replace `fontFamilyStyle` (`:140-145`):
```js
  function fontFamilyStyle(opts) {
    opts = opts || {};
    var mode = opts.fontMode === "nastaliq" ? "noto" : opts.fontMode; // legacy alias
    var css = AshaarFonts.cssFamilyOf(mode);
    return css ? "font-family:" + css : "";
  }
```
Replace the rFonts line in `misraParaXml` (`:1093-1094`):
```js
    var mode = (opts || {}).fontMode === "nastaliq" ? "noto" : (opts || {}).fontMode;
    var csName = AshaarFonts.wordNameOf(mode);
    if (csName) rpr += '<w:rFonts w:cs="' + csName + '"/>';
```
Ensure `fontFamilyStyle` is included in the object `factory` returns (so tests can call `AshaarWord.fontFamilyStyle`).

In `src/taskpane/word-tabstop.js` `runPropsXml` (`:65-72`), add `AshaarFonts` to the factory and replace the font branch:
```js
  function runPropsXml(opts, isRefrain) {
    var inner = "<w:rtl/>";
    var mode = (opts || {}).fontMode === "nastaliq" ? "noto" : (opts || {}).fontMode;
    var csName = AshaarFonts.wordNameOf(mode);
    if (csName) inner += '<w:rFonts w:cs="' + csName + '"/>';
    if (isRefrain) inner += '<w:color w:val="A7352A"/>';
    return "<w:rPr>" + inner + "</w:rPr>";
  }
```

In `src/taskpane/taskpane.js` `previewFontFamily` (`:255-259`):
```js
  function previewFontFamily(font) {
    var mode = font === "nastaliq" ? "noto" : font;
    var css = AshaarFonts.cssFamilyOf(mode);
    return css || "\"Times New Roman\", serif";
  }
```

In `src/taskpane/taskpane.html`, add `"./fonts.js"` to the `srcs` array (the loader at ~`:326-331`) **before** `"./word-html.js"` and `"./word-tabstop.js"`.

- [ ] **Step 4: Run tests to verify they pass**

Run: `node tests/word-html.test.js && node tests/word-tabstop.test.js`
Expected: PASS, including `word-html registry-font tests passed`. Existing assertions still green (Noto/arabic-serif unchanged).

- [ ] **Step 5: Full suite + commit**

Run: `npm test` (all pass).
```bash
git add src/taskpane/word-html.js src/taskpane/word-tabstop.js src/taskpane/taskpane.js src/taskpane/taskpane.html tests/word-html.test.js
git commit -m "refactor(fonts): route font-family + rFonts through AshaarFonts registry

Co-Authored-By: Claude Opus 4.8 (1M context) <noreply@anthropic.com>"
```

---

## Task 3: Bundle Mehr + Gulzar, add @font-face, populate the font picker from the registry

**Files:**
- Add: `assets/fonts/MehrNastaliqWeb.woff2`, `assets/fonts/Gulzar-Regular.woff2` (and `assets/fonts/JameelNooriNastaleeqKasheeda.ttf` + `-italic` face **only if Gate G passed**)
- Modify: `src/taskpane/taskpane.css:4-30` region (add `@font-face` blocks alongside FatemiMaqala)
- Modify: `src/taskpane/taskpane.html:43-46` (font `<select>` — populate from registry)
- Modify: `src/taskpane/taskpane.js` (populate `#font-mode` options from `AshaarFonts.LIST` on Office-ready)
- Modify: `README.md` / `LICENSE` (Mehr CC-BY-SA attribution, Gulzar OFL notice)

**Interfaces:**
- Consumes: `AshaarFonts.LIST` (Task 1).

**Fetch sources:** Mehr from the `mehr` npm package (`node_modules/mehr/…woff2`) or mehrtype.com; Gulzar from Google Fonts (`fonts.google.com/specimen/Gulzar`, OFL). This step is mechanical file placement; verified visually in Word (no Node test asserts binary font bytes).

- [ ] **Step 1: Add the font files**

Place the woff2/ttf files under `assets/fonts/` with the exact filenames in the registry `file` fields. Confirm: `ls assets/fonts/` shows `MehrNastaliqWeb.woff2` and `Gulzar-Regular.woff2`.

- [ ] **Step 2: Add @font-face blocks**

In `src/taskpane/taskpane.css`, after the FatemiMaqala block (`:4-17`), add:
```css
@font-face {
  font-family: "Mehr Nastaliq Web";
  src: url("../../assets/fonts/MehrNastaliqWeb.woff2") format("woff2");
  font-display: swap;
}
@font-face {
  font-family: "Gulzar";
  src: url("../../assets/fonts/Gulzar-Regular.woff2") format("woff2");
  font-display: swap;
}
```
**Only if Gate G passed**, also add the two Jameel faces (normal + italic) so canvas `ctx.font = "italic …"` and Word `<w:i/>` both resolve kasheeda glyphs:
```css
@font-face {
  font-family: "Jameel Noori Nastaleeq Kasheeda";
  src: url("../../assets/fonts/JameelNooriNastaleeqKasheeda.ttf") format("truetype");
  font-style: normal;
}
@font-face {
  font-family: "Jameel Noori Nastaleeq Kasheeda";
  src: url("../../assets/fonts/JameelNooriNastaleeqKasheeda-italic.ttf") format("truetype");
  font-style: italic;
}
```
(If Jameel ships as a single TTF whose italic bit swaps glyphs, point the italic `src` at the same file. Confirm during Gate G.)

- [ ] **Step 3: Populate the font picker from the registry**

Replace the hard-coded `<option>`s in `src/taskpane/taskpane.html` (`:44-46`) with an empty `<select id="font-mode"></select>`, then in `taskpane.js` (in `bind()` on Office-ready) add:
```js
  (function populateFontModes() {
    var order = ["document", "arabic-serif", "noto", "mehr", "jameel", "gulzar"];
    fontMode.innerHTML = "";
    order.forEach(function (id) {
      var d = AshaarFonts.get(id);
      if (!d) return;
      var o = document.createElement("option");
      o.value = id; o.textContent = d.label;
      if (id === "document") o.selected = true;
      fontMode.appendChild(o);
    });
  })();
```

- [ ] **Step 4: Manual verify in Word (checklist)**

- [ ] Preview pane: selecting **Mehr**, **Gulzar** renders each face (not a serif fallback).
- [ ] Insert a poem in Mehr → the Word document shows Mehr (with the font installed).
- [ ] **Mehr whitelist spot-check:** in the bundled Mehr, type a tatweel before each of `ب پ ت ٹ ث س ش ف ک گ` (medial) and confirm elongation; confirm a join into `ر`/`د`/`ه` does **not** elongate; note whether the bundled build is still Beta 2.0.

- [ ] **Step 5: Attribution + commit**

Add Mehr (CC-BY-SA) and Gulzar (OFL) notices to `README.md`.
```bash
git add assets/fonts/MehrNastaliqWeb.woff2 assets/fonts/Gulzar-Regular.woff2 src/taskpane/taskpane.css src/taskpane/taskpane.html src/taskpane/taskpane.js README.md
git commit -m "feat(fonts): bundle Mehr + Gulzar, @font-face + registry-driven picker

Co-Authored-By: Claude Opus 4.8 (1M context) <noreply@anthropic.com>"
```

---

## Task 4: Mehr tatweel whitelist + mechanism dispatch + whitespace guard

Two pieces: (a) a **pure** helper that turns Mehr's whitelist into a `priorityTable` the vendored `buildSlots` already honors via `canInsertTatweel`; (b) dispatch in `justifySelection` that applies it for `tatweel` fonts and **guards `whitespace` fonts off the tatweel path** so Gulzar never shatters.

**Files:**
- Create: `src/taskpane/tatweel-whitelist.js`
- Test: `tests/tatweel-whitelist.test.js`
- Modify: `src/taskpane/taskpane.js` (justify dispatch, ~`:1229` probe + `:1317` kashida branch)
- Modify: `package.json` (test chain)

**Interfaces:**
- Consumes: `AshaarFonts.tatweelRulesOf`, `AshaarFonts.mechanismOf` (Task 1).
- Produces:
  - `AshaarTatweel.buildPriorityTable(text, rules) → { "<prev><next>": {blocked:true}, ... }` — pure. Blocks every adjacent letter-pair in `text` whose `next` letter is NOT in `rules.medialInto`. (Word-final `finalInto` narrowing of س/ش is a documented on-device check, not table-expressible.)
  - The dispatch sets `calibParams.priorityTable = buildPriorityTable(corpusText, rules)` for Mehr before `justifyRuns`.

- [ ] **Step 1: Write the failing test**

Create `tests/tatweel-whitelist.test.js`:
```js
"use strict";
const assert = require("assert");
const AshaarTatweel = require("../src/taskpane/tatweel-whitelist");
const AshaarFonts = require("../src/taskpane/fonts");

const rules = AshaarFonts.tatweelRulesOf("mehr");

// "کتاب" = ک-ت-ا-ب. Pairs: کت (next ت ✓ allowed), تا (next ا ✗ block), اب (next ب ✓ allowed).
const t = AshaarTatweel.buildPriorityTable("کتاب", rules);
assert.ok(!t["کت"] || !t["کت"].blocked, "into ت allowed");
assert.ok(t["تا"] && t["تا"].blocked === true, "into ا blocked");
assert.ok(!t["اب"] || !t["اب"].blocked, "into ب allowed");

// join into ر is never whitelisted → blocked
const t2 = AshaarTatweel.buildPriorityTable("در", rules);
assert.ok(t2["در"] && t2["در"].blocked === true, "into ر blocked");

console.log("tatweel-whitelist tests passed");
```

- [ ] **Step 2: Run test to verify it fails**

Run: `node tests/tatweel-whitelist.test.js`
Expected: FAIL — `Cannot find module '../src/taskpane/tatweel-whitelist'`.

- [ ] **Step 3: Write minimal implementation**

Create `src/taskpane/tatweel-whitelist.js`:
```js
/**
 * AshaarTatweel — turn a font's tatweel whitelist into a priorityTable the
 * vendored buildSlots()/canInsertTatweel() already honor (a pair with
 * {blocked:true} is skipped). We block every adjacent letter-pair present in
 * the text whose *next* letter is not elongatable in this font, so the engine
 * only inserts tatweels the font actually renders. Pure; no DOM.
 */
(function (root, factory) {
  if (typeof module !== "undefined" && module.exports) module.exports = factory();
  else root.AshaarTatweel = factory();
}(typeof globalThis !== "undefined" ? globalThis : this, function () {
  "use strict";

  // Combining marks / non-letters we skip when finding adjacent base letters.
  var SKIP = /[ً-ٰٟـ\s]/; // harakat, superscript alef, tatweel, space

  function baseLetters(word) {
    var out = [];
    for (var i = 0; i < word.length; i++) {
      var ch = word.charAt(i);
      if (!SKIP.test(ch)) out.push(ch);
    }
    return out;
  }

  function buildPriorityTable(text, rules) {
    var table = {};
    if (!rules || !rules.medialInto) return table;
    var allowed = {};
    for (var a = 0; a < rules.medialInto.length; a++) allowed[rules.medialInto[a]] = true;
    var words = String(text).split(" ");
    words.forEach(function (w) {
      var letters = baseLetters(w);
      for (var i = 0; i < letters.length - 1; i++) {
        var next = letters[i + 1];
        if (!allowed[next]) table[letters[i] + next] = { blocked: true };
      }
    });
    return table;
  }

  return { buildPriorityTable: buildPriorityTable };
}));
```

- [ ] **Step 4: Run test to verify it passes**

Run: `node tests/tatweel-whitelist.test.js`
Expected: PASS — `tatweel-whitelist tests passed`.

- [ ] **Step 5: Wire dispatch + whitespace guard into `justifySelection`**

In `src/taskpane/taskpane.js`, in `justifySelection` (the run-aware block around `:1229`–`:1327`):

1. Near the top of the function, resolve the mechanism from the selected font:
```js
    var fontId = opts.fontMode === "nastaliq" ? "noto" : opts.fontMode;
    var mechanism = AshaarFonts.mechanismOf(fontId);
```
2. **Whitespace guard** — before the tatweel path runs, if a `whitespace` font is selected while the chooser is on `kashida`, downgrade to spacing and warn (Gulzar/Noto shatter under injected tatweels):
```js
    if (mechanism === "whitespace" && opts.justifyMode === "kashida") {
      opts = Object.assign({}, opts, { justifyMode: "spacing" });
      setMessage("“" + (AshaarFonts.get(fontId) || {}).label + "” has no stretch letters — filling by spacing instead.");
    }
```
3. **Mehr whitelist** — where `calibParams` is assembled (before `AshaarJustify.justifyRuns(primRuns, colPx, calibParams)` at `:1318`), inject the priorityTable for tatweel fonts that have rules:
```js
      if (mechanism === "tatweel") {
        var rules = AshaarFonts.tatweelRulesOf(fontId);
        if (rules) {
          var corpus = runs.map(function (r) { return r.text; }).join(" ");
          calibParams = Object.assign({}, calibParams, {
            priorityTable: AshaarTatweel.buildPriorityTable(corpus, rules)
          });
        }
      }
```
Add `AshaarTatweel` / `AshaarFonts` to the browser globals already used in `taskpane.js` (they are loaded via the `srcs` list; confirm `./tatweel-whitelist.js` is added to that list in `taskpane.html`, before `taskpane.js`).

- [ ] **Step 6: Add loader entry + test chain; full suite**

- Add `"./tatweel-whitelist.js"` to the `srcs` array in `taskpane.html` (before `taskpane.js`).
- Append ` && node tests/tatweel-whitelist.test.js` to `package.json`'s `test` script.
Run: `npm test`
Expected: all pass.

- [ ] **Step 7: Manual verify in Word**

- [ ] Select **Gulzar** + Kashida mode → message shows the spacing-downgrade note; the document is filled by spacing, **no tatweels injected** (Gulzar shaping intact).
- [ ] Select **Mehr** + Kashida → lines fill; tatweels appear only on whitelisted joins; ragged→clean.

- [ ] **Step 8: Commit**

```bash
git add src/taskpane/tatweel-whitelist.js tests/tatweel-whitelist.test.js src/taskpane/taskpane.js src/taskpane/taskpane.html package.json
git commit -m "feat(justify): per-font mechanism dispatch — Mehr whitelist + whitespace guard

Co-Authored-By: Claude Opus 4.8 (1M context) <noreply@anthropic.com>"
```

---

## Task 5: Jameel italic-run engine (`kashida-italic.js`) + OOXML italic emission

**GATED on Gate G = PASS.** If Gate G failed, skip this task (Jameel is already `whitespace` via the Task-G decision).

**Files:**
- Create: `src/taskpane/kashida-italic.js`
- Test: `tests/kashida-italic.test.js`
- Modify: `src/taskpane/taskpane.js` (dispatch `italic-run` branch → `insertOoxml`)
- Modify: `src/taskpane/word-html.js` (a small `runsToOoxml(runs, opts)` helper that emits `<w:i/>` per italic run, reusing `misraParaXml`'s rPr construction)
- Modify: `package.json` (test chain)

**Interfaces:**
- Consumes: `AshaarFonts.wordNameOf("jameel")` (Task 1).
- Produces:
  - `AshaarKashidaItalic.splitSpans(text) → string[]` — connected segments (fasl/PAW), order-preserving so `spans.join("") === text`.
  - `AshaarKashidaItalic.selectItalicRuns(spans, widthsNormal, widthsItalic, targetPx) → { runs:[{text,italic}], fill, reason }` — greedy discrete subset-selection.

- [ ] **Step 1: Write the failing test**

Create `tests/kashida-italic.test.js`:
```js
"use strict";
const assert = require("assert");
const K = require("../src/taskpane/kashida-italic");

// splitSpans: break at non-joining letters. "ستارہ" = س-ت-ا | ر | ہ
const spans = K.splitSpans("ستارہ");
assert.deepStrictEqual(spans, ["ستا", "ر", "ہ"]);
assert.strictEqual(spans.join(""), "ستارہ");

// selectItalicRuns: pick highest-gain spans until <= target, mark them italic.
// widthsNormal sum = 30; italic adds gain [+8, +2, +0]; target 36.
// greedy by gain: italicize span0 (+8 → 38 > 36? yes overshoot) → skip if strict;
// then span1 (+2 → 32 <= 36) italic. Result: only span1 italic, fill 32/36.
const r = K.selectItalicRuns(["a","b","c"], [10,10,10], [18,12,10], 36);
assert.strictEqual(r.runs.length, 3);
assert.strictEqual(r.runs[1].italic, true);
assert.strictEqual(r.runs[0].italic, false);
assert.strictEqual(r.runs[2].italic, false);
assert.ok(r.fill > 0 && r.fill <= 1);

// no elongatable spans (italic == normal) → reason set, nothing italic
const r2 = K.selectItalicRuns(["a","b"], [10,10], [10,10], 40);
assert.strictEqual(r2.runs.every(function (x){return !x.italic;}), true);
assert.strictEqual(r2.reason, "no elongatable spans");

console.log("kashida-italic tests passed");
```

- [ ] **Step 2: Run test to verify it fails**

Run: `node tests/kashida-italic.test.js`
Expected: FAIL — module not found.

- [ ] **Step 3: Write minimal implementation**

Create `src/taskpane/kashida-italic.js`:
```js
/**
 * AshaarKashidaItalic — Jameel Noori Kasheeda fills a line by italicizing whole
 * connected segments (fasl / piece-of-Arabic-word), each of which swaps to its
 * wider kasheeda form. Elongation is therefore DISCRETE: we choose which spans
 * to italicize to approach the target width. splitSpans + selectItalicRuns are
 * pure; width measurement (canvas) lives in the browser caller.
 */
(function (root, factory) {
  if (typeof module !== "undefined" && module.exports) module.exports = factory();
  else root.AshaarKashidaItalic = factory();
}(typeof globalThis !== "undefined" ? globalThis : this, function () {
  "use strict";

  // Letters that do NOT join to the following letter → a segment ends after them.
  var NONJOIN = "اأإآٱدذڈرزڑژوؤءے";

  function splitSpans(text) {
    var spans = [];
    var cur = "";
    for (var i = 0; i < text.length; i++) {
      var ch = text.charAt(i);
      if (ch === " ") { if (cur) { spans.push(cur); cur = ""; } spans.push(" "); continue; }
      cur += ch;
      if (NONJOIN.indexOf(ch) !== -1) { spans.push(cur); cur = ""; }
    }
    if (cur) spans.push(cur);
    return spans;
  }

  function selectItalicRuns(spans, widthsNormal, widthsItalic, targetPx) {
    var n = spans.length;
    var italic = new Array(n);
    var total = 0, i;
    for (i = 0; i < n; i++) { italic[i] = false; total += widthsNormal[i]; }

    // Candidate gains, largest first.
    var cand = [];
    for (i = 0; i < n; i++) {
      var gain = widthsItalic[i] - widthsNormal[i];
      if (gain > 0) cand.push({ i: i, gain: gain });
    }
    cand.sort(function (a, b) { return b.gain - a.gain; });

    var reason = null;
    if (!cand.length) reason = "no elongatable spans";

    for (var k = 0; k < cand.length; k++) {
      var add = cand[k].gain;
      if (total + add <= targetPx) { italic[cand[k].i] = true; total += add; }
    }
    if (!reason && total < targetPx) reason = "discrete steps underfill";

    var runs = [];
    for (i = 0; i < n; i++) runs.push({ text: spans[i], italic: italic[i] });
    return { runs: runs, fill: targetPx > 0 ? total / targetPx : 0, reason: reason };
  }

  return { splitSpans: splitSpans, selectItalicRuns: selectItalicRuns };
}));
```

- [ ] **Step 4: Run test to verify it passes**

Run: `node tests/kashida-italic.test.js`
Expected: PASS — `kashida-italic tests passed`.

- [ ] **Step 5: Add the OOXML italic-run emitter**

In `src/taskpane/word-html.js`, add and export a helper that renders a run list to a single misra paragraph, italic runs carrying `<w:i/>` (reuse the `jc`/`bidi`/rFonts logic from `misraParaXml`):
```js
  // runs: [{text, italic}]; opts.fontMode drives the cs font (jameel).
  function runsToMisraXml(runs, align, opts) {
    var jc = align === "right" ? "right" : align === "left" ? "left" : "center";
    var mode = (opts || {}).fontMode === "nastaliq" ? "noto" : (opts || {}).fontMode;
    var csName = AshaarFonts.wordNameOf(mode);
    var cs = csName ? '<w:rFonts w:cs="' + csName + '"/>' : "";
    var body = runs.map(function (r) {
      var rpr = "<w:rPr><w:rtl/>" + cs + (r.italic ? "<w:i/>" : "") + "</w:rPr>";
      return "<w:r>" + rpr + '<w:t xml:space="preserve">' + escapeXml(r.text) + "</w:t></w:r>";
    }).join("");
    return "<w:p><w:pPr><w:bidi/><w:spacing w:after=\"80\"/><w:jc w:val=\"" + jc + "\"/></w:pPr>" + body + "</w:p>";
  }
```
Add `runsToMisraXml` to the object `factory` returns.

- [ ] **Step 6: Write a failing test for the emitter**

In `tests/word-html.test.js` add:
```js
{
  const xml = AshaarWord.runsToMisraXml(
    [{text:"ابc", italic:false},{text:" ", italic:false},{text:"دe", italic:true}],
    "right", { fontMode: "jameel" });
  assert.ok(xml.indexOf('<w:rFonts w:cs="Jameel Noori Nastaleeq"/>') !== -1, "jameel cs name");
  assert.ok((xml.match(/<w:i\/>/g) || []).length === 1, "exactly one italic run");
  console.log("word-html italic-run tests passed");
}
```
Run: `node tests/word-html.test.js` → verify it fails first (helper absent), then passes after Step 5.

- [ ] **Step 7: Dispatch `italic-run` in `justifySelection`**

In `taskpane.js`, add an `italic-run` branch to the mechanism dispatch (Task 4, Step 5). For each cell, when `mechanism === "italic-run"`:
```js
      // Jameel: measure spans normal vs italic on the cell's own canvas ctx.
      var spans = AshaarKashidaItalic.splitSpans(stripJustification(current));
      var wn = [], wi = [];
      spans.forEach(function (s) {
        canvasCtx.font = runFontStr(repName, repSize, false, false); wn.push(canvasCtx.measureText(s).width);
        canvasCtx.font = runFontStr(repName, repSize, false, true);  wi.push(canvasCtx.measureText(s).width);
      });
      var sel = AshaarKashidaItalic.selectItalicRuns(spans, wn, wi, colPx);
      var xml = AshaarWord.runsToMisraXml(
        sel.runs.map(function (r){ return { text: r.text, italic: r.italic }; }),
        cellAlign, opts);
      plans.push({ cell: cell, ooxml: xml });
```
Then, where `plans` are applied, handle the `ooxml` case with `cell.body.clear(); cell.body.insertOoxml(plan.ooxml, Word.InsertLocation.replace);` (alongside the existing `flat`/text write-back). `runFontStr(name, size, bold, italic)` already exists in `taskpane.js` (used at `:1306`, `:1323`) and produces the CSS `font` string; confirm it sets the italic flag.

- [ ] **Step 8: Loader + test chain + full suite**

- Add `"./kashida-italic.js"` to `srcs` in `taskpane.html` (before `taskpane.js`).
- Append ` && node tests/kashida-italic.test.js` to `package.json` `test`.
Run: `npm test` → all pass.

- [ ] **Step 9: Manual verify in Word**

- [ ] Select **Jameel** + Kashida → a short misra fills by italicizing whole connected segments; segments shift to elongated kasheeda forms (not oblique); no mid-cluster italic.
- [ ] Underfilled line reports a sensible state (fill < 1, reason surfaced in the existing debug block).

- [ ] **Step 10: Commit**

```bash
git add src/taskpane/kashida-italic.js tests/kashida-italic.test.js src/taskpane/word-html.js tests/word-html.test.js src/taskpane/taskpane.js src/taskpane/taskpane.html package.json
git commit -m "feat(justify): Jameel italic-run kashida — fasl subset-selection + OOXML italic runs

Co-Authored-By: Claude Opus 4.8 (1M context) <noreply@anthropic.com>"
```

---

## Task 6: Reader-end note + Jameel private-deploy guard

**Files:**
- Modify: `src/taskpane/taskpane.js` (per-font install note on font-mode change)
- Modify: `src/taskpane/taskpane.html` (a note element under the font picker) + `taskpane.css` (note style)
- Modify: `.gitignore` and/or the Pages deploy path so `private` fonts never publish
- Modify: `README.md` (document the reader requirement)

**Interfaces:**
- Consumes: `AshaarFonts.get(id).private`, `.bundled`, `.label` (Task 1).

- [ ] **Step 1: Add the reader-end note**

In `taskpane.html`, add under the font `<select>`: `<p id="font-install-note" class="adopt-hint" hidden></p>`.
In `taskpane.js`, on `fontMode` change:
```js
  function updateFontNote() {
    var d = AshaarFonts.get(fontMode.value === "nastaliq" ? "noto" : fontMode.value);
    var note = document.getElementById("font-install-note");
    if (!note) return;
    if (d && d.wordName && d.mechanism !== "whitespace") {
      note.hidden = false;
      note.textContent = "Readers need “" + d.wordName + "” installed to see this correctly."
        + (d.id === "jameel" ? " Specifically the Kasheeda build, or italic runs won’t elongate." : "");
    } else { note.hidden = true; }
  }
```
Wire `updateFontNote` to the `fontMode` `change` listener and call once in `bind()`.

- [ ] **Step 2: Guard the private font out of the public deploy**

Add to `.gitignore` (or the deploy exclude list): `assets/fonts/JameelNooriNastaleeqKasheeda*.ttf`. If Jameel was committed for local use, instead document in `README.md` that the Pages deploy step must exclude `private` fonts, and keep the file out of `main`. Verify: the GitHub Pages URL for the Jameel file 404s after deploy (or the file is absent from the deployed tree).

- [ ] **Step 3: Manual verify + README**

- [ ] Changing the font picker to Mehr/Gulzar/Jameel shows the correct install note; `document`/`arabic-serif` show none.
- [ ] Add a "Fonts & reader requirements" section to `README.md` (which fonts bundle, which readers must install, Jameel privacy).

- [ ] **Step 4: Commit**

```bash
git add src/taskpane/taskpane.js src/taskpane/taskpane.html src/taskpane/taskpane.css .gitignore README.md
git commit -m "feat(fonts): reader-install note + Jameel private-deploy guard

Co-Authored-By: Claude Opus 4.8 (1M context) <noreply@anthropic.com>"
```

---

## Self-Review

**Spec coverage:**
- §1 Font registry → **Task 1** ✓
- §1 refactor the five literals → **Task 2** ✓
- §2 bundling + @font-face + preview → **Task 3** ✓
- §3 mechanism→mode interlock (whitespace guard against current chooser; full §5 reframe deferred per sequencing decision) → **Task 4** (guard) ✓; §5 presentation reframe explicitly out of scope (documented below)
- §4 Jameel kashida-italic.js + insertOoxml → **Task 5** (gated) ✓
- §4 Mehr whitelist / tatweelRules → **Task 4** ✓
- §5 dispatch by mechanism → **Task 4** (tatweel/whitespace) + **Task 5** (italic-run) ✓
- §6 reader-end + private-deploy guard → **Task 6** ✓
- Spike (Gate G) → **Manual Gate** ✓

**Deferred (per the "build engine now" sequencing decision), not gaps:** the guided-justification §5 four-mode chooser *presentation* reframe and its `modesFor(font)` UI filter. Task 4 delivers the *safety* interlock (whitespace fonts can't run the tatweel engine) against the current chooser; the fuller UX filter lands when guided-justification is implemented.

**Placeholder scan:** no TBD/TODO; every code step has real code; every command has expected output.

**Type consistency:** `AshaarFonts.{get,mechanismOf,wordNameOf,cssFamilyOf,tatweelRulesOf}` used identically across Tasks 1–6. `buildPriorityTable(text, rules)` signature matches between Task 4 test and dispatch. `selectItalicRuns(spans, widthsNormal, widthsItalic, targetPx)` and `splitSpans(text)` match between Task 5 test, module, and dispatch. `runsToMisraXml(runs, align, opts)` matches between Task 5 emitter, its test, and the dispatch caller. `priorityTable` shape (`{ "<prev><next>": {blocked:true} }`) matches the vendored `pairEntry`/`canInsertTatweel` contract read from `ashaar-justify.js:123-136`.

## Out of scope (this plan)

- Guided-justification §5 mode-chooser reframe / §3 native Word-kashida / §4 stretch-strength (separate specs & plans).
- Editing the vendored engine.
- `.docx` font embedding.
- Additional Nastaliq fonts (Awami/Nafees/Alvi/Fajer) — now one registry entry each.

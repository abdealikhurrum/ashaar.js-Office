# Text Conversion (Double-Press ⇄ Modern) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add an on-demand find-replace subsystem that converts Arabic-script text between the legacy LD "double-press"/AL-KANZ encoding and a modern encoding (Unicode where it exists, Fatemi `!keyword!` otherwise), exposed as a new "Convert" task-pane tab.

**Architecture:** A pure, node-testable data module (`word-conversion.js`) owns the mapping table and turns a direction + selected rows into an ordered list of literal find→replace ops (ordering encodes the kashida-escape rule). A thin Office.js layer (`conversion-pane.js`) applies those ops via `Range.search()`+range-replace so run formatting survives, and manages named presets. A generator script extracts the `mark`/`symbol` rows from on-disk fonts.

**Tech Stack:** Vanilla ES5/UMD JS (no build step), Office.js v1, Node `assert` tests, `fonttools` (via the Kanz-al-Marjaan venv) for the one-time extraction.

## Global Constraints

- No build step; ES5/UMD modules; `src/vendor/` is never edited.
- Pure logic is node-testable with `require()`; Office.js orchestration lives separately (mirror `word-styles.js` ↔ `styles-pane.js`).
- Tests use only Node's built-in `assert` (no jest/mocha), one `tests/<name>.test.js` file wired into the `"test"` script in `package.json`.
- New tab must hide the poetry chrome in `setMode()` exactly as Styles/Booklet do.
- Every pane-code push must bump `ASHAAR_ASSET_VERSION` (installed users otherwise run stale JS).
- Tatweel is `"ـ"` (ـ). Fonts: legacy `Al-Kanz for Windows.ttf` (`~/Downloads/Al_Kanz_Fonts_For_Windows/`), modern `assets/fonts/FatemiMaqala-Regular.ttf`; extraction python is `~/Kanz-al-Marjaan/venv/bin/python` (fonttools 4.55).

---

### Task 1: Pure conversion module — mapping table + engine

**Files:**
- Create: `src/taskpane/word-conversion.js`
- Test: `tests/word-conversion.test.js`

**Interfaces:**
- Produces:
  - `DIRECTIONS = { TO_MODERN: "toModern", TO_LEGACY: "toLegacy" }`
  - `TATWEEL = "ـ"`
  - `MAPPINGS: Array<{id, category:"letter"|"mark"|"symbol", legacy, modern, label, wholeWord:boolean, lossy:boolean}>`
  - `buildOperations(direction, enabledIds?) -> Array<{find, replaceWith, wholeWord, category}>` (ordered)
  - `convert(text, direction, enabledIds?) -> string`
  - `groupsForUi() -> Array<{category, rows:MAPPINGS[]}>`
- Consumes: nothing (leaf module).

- [ ] **Step 1: Write the failing test** — `tests/word-conversion.test.js`

```js
const assert = require("assert");
const C = require("../src/taskpane/word-conversion");
const T = C.TATWEEL;

// ── table shape ──
{
  assert.deepEqual(Object.keys(C.DIRECTIONS).sort(), ["TO_LEGACY", "TO_MODERN"]);
  assert.ok(Array.isArray(C.MAPPINGS) && C.MAPPINGS.length >= 10);
  const ids = C.MAPPINGS.map(m => m.id);
  assert.equal(new Set(ids).size, ids.length, "ids are unique");
  C.MAPPINGS.forEach(m => {
    assert.ok(["letter", "mark", "symbol"].includes(m.category), m.id + " category");
    assert.equal(typeof m.legacy, "string");
    assert.equal(typeof m.modern, "string");
    assert.equal(typeof m.wholeWord, "boolean");
    assert.equal(typeof m.lossy, "boolean");
  });
  const seen = C.MAPPINGS.find(m => m.id === "seen-baariye");
  assert.equal(seen.legacy, "سس"); assert.equal(seen.modern, "ے");
}

// ── letter tier: simple contiguous substitution both ways ──
{
  assert.equal(C.convert("حح", C.DIRECTIONS.TO_MODERN), "چ");
  assert.equal(C.convert("چ", C.DIRECTIONS.TO_LEGACY), "حح");
  assert.equal(C.convert("كك گگ", C.DIRECTIONS.TO_MODERN), "گ گگ".replace("گگ","گگ"));
}

// ── kashida-escape: genuine double letter round-trips ──
{
  // TO_MODERN: escaped double (سـس) becomes a genuine double seen (سس),
  // while a plain double (سس) becomes ے.
  assert.equal(C.convert("سس", C.DIRECTIONS.TO_MODERN), "ے");
  assert.equal(C.convert("س" + T + "س", C.DIRECTIONS.TO_MODERN), "سس");
  // TO_LEGACY: ے becomes سس (contiguous), while a genuine double seen (سس)
  // is protected with a tatweel so the old font won't merge it.
  assert.equal(C.convert("ے", C.DIRECTIONS.TO_LEGACY), "سس");
  assert.equal(C.convert("سس", C.DIRECTIONS.TO_LEGACY), "س" + T + "س");
  // Round-trip a word containing a genuine double seen ("مسس" style token).
  const word = "بسس"; // بـ + genuine double seen
  const toLegacy = C.convert(word, C.DIRECTIONS.TO_LEGACY);   // ب س ـ س
  assert.equal(toLegacy, "ب" + "س" + T + "س");
  assert.equal(C.convert(toLegacy, C.DIRECTIONS.TO_MODERN), word, "round-trips");
}

// ── whole-word: چھے ⇄ ؛ ──
{
  assert.equal(C.convert("؛", C.DIRECTIONS.TO_MODERN), "چھے");
  assert.equal(C.convert("چھے", C.DIRECTIONS.TO_LEGACY), "؛", "standalone word collapses");
  assert.equal(C.convert("اچھے", C.DIRECTIONS.TO_LEGACY), "اچھے",
    "چھے inside a larger word is NOT collapsed");
}

// ── enabledIds filters which rows run ──
{
  assert.equal(C.convert("حح كك", C.DIRECTIONS.TO_MODERN, ["cheh-hah"]), "چ كك",
    "only the enabled row converts");
}

// ── buildOperations ordering: TO_MODERN puts contiguous before escape-drop ──
{
  const ops = C.buildOperations(C.DIRECTIONS.TO_MODERN, ["seen-baariye"]);
  const iContig = ops.findIndex(o => o.find === "سس");
  const iEscape = ops.findIndex(o => o.find === "س" + T + "س");
  assert.ok(iContig >= 0 && iEscape >= 0 && iContig < iEscape,
    "contiguous سس→ے must run before سـس→سس");
}
// ── buildOperations ordering: TO_LEGACY protects doubles before ے→سس ──
{
  const ops = C.buildOperations(C.DIRECTIONS.TO_LEGACY, ["seen-baariye"]);
  const iProtect = ops.findIndex(o => o.find === "سس");
  const iSub = ops.findIndex(o => o.find === "ے");
  assert.ok(iProtect >= 0 && iSub >= 0 && iProtect < iSub,
    "protect سس→سـس must run before ے→سس");
}

// ── groupsForUi ──
{
  const groups = C.groupsForUi();
  assert.ok(groups.every(g => g.category && Array.isArray(g.rows)));
  assert.ok(groups.some(g => g.category === "letter"));
}

console.log("word-conversion.test.js: all assertions passed");
```

- [ ] **Step 2: Run it — expect failure** — `node tests/word-conversion.test.js` → fails (`Cannot find module '../src/taskpane/word-conversion'`).

- [ ] **Step 3: Implement `src/taskpane/word-conversion.js`**

```js
/**
 * AshaarConversion — pure data + engine for the "Convert" tab. Converts
 * Arabic-script text between the legacy LD double-press/AL-KANZ encoding and a
 * modern encoding (Unicode where it exists, Fatemi !keyword! otherwise).
 * No Office.js/DOM; the Word.run() orchestration lives in conversion-pane.js.
 * See docs/superpowers/specs/2026-07-16-text-conversion-design.md.
 */
(function (root, factory) {
  if (typeof module !== "undefined" && module.exports) module.exports = factory();
  else root.AshaarConversion = factory();
}(typeof globalThis !== "undefined" ? globalThis : this, function () {
  "use strict";

  var TATWEEL = "ـ";
  var DIRECTIONS = { TO_MODERN: "toModern", TO_LEGACY: "toLegacy" };

  // legacy = old double-press/placeholder form; modern = Unicode-or-!keyword!.
  // wholeWord: only collapse when standalone (Modern→Legacy). lossy: not safely
  // round-trippable (source char has a legitimate independent meaning).
  var MAPPINGS = [
    { id: "seen-baariye",  category: "letter", legacy: "سس", modern: "ے",   label: "ے baari ye",        wholeWord: false, lossy: false },
    { id: "zah-heh",       category: "letter", legacy: "ظظ", modern: "ہ",   label: "ہ gol he",          wholeWord: false, lossy: false },
    { id: "tah-noonghunna",category: "letter", legacy: "طط", modern: "ں",   label: "ں noon ghunna",     wholeWord: false, lossy: false },
    { id: "kaf-gaf",       category: "letter", legacy: "كك", modern: "گ",   label: "گ gaf",             wholeWord: false, lossy: false },
    { id: "cheh-hah",      category: "letter", legacy: "حح", modern: "چ",   label: "چ cheh",            wholeWord: false, lossy: false },
    { id: "tteh-dad",      category: "letter", legacy: "ضض", modern: "ٹ",   label: "ٹ tteh",            wholeWord: false, lossy: false },
    { id: "rreh-re",       category: "letter", legacy: "رٌ", modern: "ڑ", label: "ڑ rreh (rā+dammatan)", wholeWord: false, lossy: false },
    { id: "ddal-dal",      category: "letter", legacy: "دٌ", modern: "ڈ", label: "ڈ ddal (dāl+dammatan)", wholeWord: false, lossy: false },
    { id: "peh-theh",      category: "letter", legacy: "ثث", modern: "پ",   label: "پ peh",             wholeWord: false, lossy: false },
    { id: "chhay-semicolon",category:"letter", legacy: "؛",  modern: "چھے", label: "چھے ⇄ ؛ (semicolon)", wholeWord: true, lossy: false }
    // `mark` and `symbol` rows are appended by scripts/generate-conversion-table.mjs (Task 2).
  ];

  function isDoubledConsonant(m) {
    return m.category === "letter" && m.legacy.length === 2 && m.legacy[0] === m.legacy[1];
  }

  function enabledSet(ids) {
    if (!ids) return null;              // null = all enabled
    var s = {}; ids.forEach(function (i) { s[i] = true; }); return s;
  }

  // Ordered literal ops. Ordering (not runtime context) enforces the escape rule.
  function buildOperations(direction, enabledIds) {
    var on = enabledSet(enabledIds);
    var rows = MAPPINGS.filter(function (m) { return !on || on[m.id]; });
    var ops = [];

    if (direction === DIRECTIONS.TO_MODERN) {
      // 1) contiguous / direct legacy→modern (longest find first)
      rows.slice().sort(byFindLenDesc("legacy")).forEach(function (m) {
        ops.push({ find: m.legacy, replaceWith: m.modern, wholeWord: false, category: m.category });
      });
      // 2) escape-drop: سـس → سس (after, so the tatweel-separated form survived step 1)
      rows.filter(isDoubledConsonant).forEach(function (m) {
        var b = m.legacy[0];
        ops.push({ find: b + TATWEEL + b, replaceWith: b + b, wholeWord: false, category: m.category });
      });
    } else {
      // 1) protect genuine doubles: سس → سـس (before ے→سس creates new doubles)
      rows.filter(isDoubledConsonant).forEach(function (m) {
        var b = m.legacy[0];
        ops.push({ find: b + b, replaceWith: b + TATWEEL + b, wholeWord: false, category: m.category });
      });
      // 2) modern→legacy (longest find first; whole-word where flagged)
      rows.slice().sort(byFindLenDesc("modern")).forEach(function (m) {
        ops.push({ find: m.modern, replaceWith: m.legacy, wholeWord: !!m.wholeWord, category: m.category });
      });
    }
    return ops;
  }

  function byFindLenDesc(key) {
    return function (a, b) { return b[key].length - a[key].length; };
  }

  // Arabic-script "letter or mark" ranges, for whole-word boundary detection.
  var WORDCHAR = /[؀-ۿݐ-ݿࢠ-ࣿﭐ-﷿ﹰ-﻿]/;
  function isWordChar(ch) { return !!ch && WORDCHAR.test(ch); }

  function replaceAll(text, find, replaceWith, wholeWord) {
    if (!find) return text;
    var out = "", i = 0;
    while (i < text.length) {
      if (text.substr(i, find.length) === find) {
        if (wholeWord) {
          var before = i > 0 ? text[i - 1] : "";
          var after = text[i + find.length] || "";
          if (isWordChar(before) || isWordChar(after)) { out += text[i]; i += 1; continue; }
        }
        out += replaceWith; i += find.length;
      } else { out += text[i]; i += 1; }
    }
    return out;
  }

  function convert(text, direction, enabledIds) {
    var ops = buildOperations(direction, enabledIds);
    var s = String(text == null ? "" : text);
    ops.forEach(function (op) { s = replaceAll(s, op.find, op.replaceWith, op.wholeWord); });
    return s;
  }

  function groupsForUi() {
    var order = ["letter", "mark", "symbol"], byCat = {};
    MAPPINGS.forEach(function (m) { (byCat[m.category] = byCat[m.category] || []).push(m); });
    return order.filter(function (c) { return byCat[c]; })
      .map(function (c) { return { category: c, rows: byCat[c] }; });
  }

  return {
    TATWEEL: TATWEEL, DIRECTIONS: DIRECTIONS, MAPPINGS: MAPPINGS,
    buildOperations: buildOperations, convert: convert, groupsForUi: groupsForUi
  };
}));
```

- [ ] **Step 4: Run tests — expect pass** — `node tests/word-conversion.test.js` → `all assertions passed`.

- [ ] **Step 5: Wire into `package.json` `"test"`** — append ` && node tests/word-conversion.test.js` to the `"test"` script.

- [ ] **Step 6: Commit** — `git add src/taskpane/word-conversion.js tests/word-conversion.test.js package.json && git commit -m "feat(convert): pure conversion engine + letter tier"`

---

### Task 2: Extract `mark`/`symbol` rows from the fonts

**Files:**
- Create: `scripts/generate-conversion-table.mjs`
- Modify: `src/taskpane/word-conversion.js` (append generated rows into `MAPPINGS`, between a `// >>> GENERATED` / `// <<< GENERATED` marker pair added in this task)
- Modify: `tests/word-conversion.test.js` (add assertions for the concrete `mark` rows once known)

**Interfaces:**
- Consumes: `MAPPINGS` from Task 1.
- Produces: additional `MAPPINGS` rows with `category:"mark"|"symbol"`.

- [ ] **Step 1: Add the generated-rows marker to `word-conversion.js`** — replace the `// mark and symbol rows are appended…` comment with:

```js
    // >>> GENERATED mark/symbol rows (scripts/generate-conversion-table.mjs) — do not edit by hand
    // <<< GENERATED
```

- [ ] **Step 2: Inspect the fonts (extraction spike)** — run, and read the output to author the rows:

```bash
PY=~/Kanz-al-Marjaan/venv/bin/python
KANZ="$HOME/Downloads/Al_Kanz_Fonts_For_Windows/Al-Kanz for Windows.ttf"
FAT="assets/fonts/FatemiMaqala-Regular.ttf"
"$PY" - "$KANZ" "$FAT" <<'PY'
import sys
from fontTools.ttLib import TTFont
kanz = TTFont(sys.argv[1]); fat = TTFont(sys.argv[2])
# 1) AL-KANZ cmap: which placeholder code points map to named glyphs (symbols).
kmap = kanz.getBestCmap()
print("AL-KANZ cmap entries:", len(kmap))
for cp, name in sorted(kmap.items()):
    if cp < 0x0600 or 0x21 <= cp <= 0x7E:   # ASCII placeholders + non-Arabic
        print("KANZ", hex(cp), repr(chr(cp)), name)
# 2) Fatemi ligature inputs beginning with '!' (the !keyword! strings).
def liga_inputs(font):
    out = []
    g = font.get("GSUB")
    if not g: return out
    for lk in g.table.LookupList.Lookup:
        for st in lk.SubTable:
            if getattr(st, "LookupType", None) == 4 or st.__class__.__name__.startswith("Ligature"):
                for first, ligs in getattr(st, "ligatures", {}).items():
                    for lig in ligs:
                        out.append((first, lig.Component, lig.LigGlyph))
    return out
print("Fatemi ligatures:", len(liga_inputs(fat)))
PY
```

  Read the printed tables. For each honorific/symbol: the AL-KANZ placeholder character(s) is the `legacy` form; the modern form is the real Unicode codepoint if the glyph maps to one (prefer it), else the Fatemi `!keyword!` string. Confirm the three `mark` rows: Shift+X → sukun `ْ` legacy → khari zabar `ٰ` modern (`lossy:true`); Shift+C / Shift+V → high jeem / high noon (codepoint if present else `!keyword!`).

- [ ] **Step 3: Write `scripts/generate-conversion-table.mjs`** — reads the fonts, builds the row objects, and rewrites the region between the `>>> GENERATED` / `<<< GENERATED` markers in `src/taskpane/word-conversion.js`. (Mirror the file-rewrite approach of `scripts/sync-ashaar-vendor.mjs`; emit each row as a `MAPPINGS.push({...})` line with the fields from Step 2.) Include the known `mark` rows explicitly so they exist even if the symbol extraction yields nothing:

```js
// Known mark rows (keyboard repurposings), always emitted:
const MARK_ROWS = [
  { id: "sukun-kharizabar", category: "mark", legacy: "ْ", modern: "ٰ", label: "khari zabar (dagger alef)", wholeWord: false, lossy: true },
  // high-jeem / high-noon rows filled from Step 2 findings:
];
```

- [ ] **Step 4: Run the generator** — `node scripts/generate-conversion-table.mjs` → rewrites the GENERATED region.

- [ ] **Step 5: Add tests for the concrete rows** — in `tests/word-conversion.test.js`:

```js
{
  const sukun = C.MAPPINGS.find(m => m.id === "sukun-kharizabar");
  assert.ok(sukun && sukun.lossy === true, "sukun→khari zabar is lossy");
  assert.equal(C.convert("ْ", C.DIRECTIONS.TO_MODERN), "ٰ");
  assert.ok(C.MAPPINGS.some(m => m.category === "symbol"), "at least one symbol row extracted");
}
```

- [ ] **Step 6: Run tests — expect pass** — `node tests/word-conversion.test.js`.

- [ ] **Step 7: Commit** — `git add scripts/generate-conversion-table.mjs src/taskpane/word-conversion.js tests/word-conversion.test.js && git commit -m "feat(convert): generate mark/symbol rows from fonts"`

---

### Task 3: Office.js orchestration — `conversion-pane.js`

**Files:**
- Create: `src/taskpane/conversion-pane.js`
- Modify: `src/taskpane/taskpane.html` (add `<script>` tag near the other pane scripts)

**Interfaces:**
- Consumes: `AshaarConversion` (Task 1) via `window.AshaarConversion`; the DOM ids created in Task 4.
- Produces: `window.ConversionPane = { bind() }` — called from taskpane bootstrap; reads the panel controls, runs the ops, reports counts, manages presets.

- [ ] **Step 1: Implement `conversion-pane.js`** (mirror `styles-pane.js`: `byId`, roaming settings via `Office.context.roamingSettings` for presets, a `run(direction)` that opens `Word.run`, resolves scope, and for each op does `scope.search(op.find,{matchCase:true,matchWholeWord:op.wholeWord})` → load → replace → sync, accumulating per-category counts). Key `run`:

```js
function run(direction) {
  var enabled = checkedIds();               // from the checklist
  var ops = AshaarConversion.buildOperations(direction, enabled);
  var scopeSel = document.querySelector('input[name="conv-scope"]:checked').value;
  var counts = { letter: 0, mark: 0, symbol: 0 };
  Word.run(function (ctx) {
    var scope = scopeSel === "selection" ? ctx.document.getSelection() : ctx.document.body;
    var chain = ctx.sync();
    ops.forEach(function (op) {
      chain = chain.then(function () {
        var res = scope.search(op.find, { matchCase: true, matchWholeWord: op.wholeWord });
        res.load("items");
        return ctx.sync().then(function () {
          res.items.forEach(function (r) { r.insertText(op.replaceWith, Word.InsertLocation.replace); });
          counts[op.category] += res.items.length;
          return ctx.sync();
        });
      });
    });
    return chain;
  }).then(function () {
    setStatus("Converted " + counts.letter + " letters, " + counts.mark + " marks, " + counts.symbol + " symbols.");
  }).catch(function (e) { setStatus("Convert failed: " + e.message); });
}
```

- [ ] **Step 2: Add `<script src="conversion-pane.js"></script>`** to `taskpane.html` alongside the other pane scripts, and call `ConversionPane.bind()` where the other panes bind on Office ready.

- [ ] **Step 3: Manual smoke** — cannot node-test Office.js. Verify no syntax errors: `node --check src/taskpane/conversion-pane.js`.

- [ ] **Step 4: Commit** — `git add src/taskpane/conversion-pane.js src/taskpane/taskpane.html && git commit -m "feat(convert): Office.js search/replace orchestration"`

---

### Task 4: Convert tab — HTML, CSS, mode wiring

**Files:**
- Modify: `src/taskpane/taskpane.html` (mode tab button + `#convertConvertPanel` panel with direction buttons, scope radio, preset dropdown, grouped checklist rendered from `AshaarConversion.groupsForUi()`, status line)
- Modify: `src/taskpane/taskpane.js` (`setMode()` — add `isConvertText` branch; hide poetry chrome; add the tab button + panel refs and click handler)
- Modify: `src/taskpane/taskpane.css` (checklist + warning-marker styles, reuse existing pane styles)
- Modify: `src/taskpane/taskpane.html` (bump `ASHAAR_ASSET_VERSION`)

**Interfaces:**
- Consumes: `groupsForUi()`, `DIRECTIONS`; `ConversionPane.bind()`.
- Produces: the DOM ids `conversion-pane.js` reads (`convScopeDoc`/`convScopeSel` radios `name="conv-scope"`, `convPreset`, `convToModern`, `convToLegacy`, `convStatus`, `convChecklist`).

- [ ] **Step 1: Add the mode tab + panel markup** in `taskpane.html` (copy the Styles tab/panel structure; give the panel `id="convertTextPanel"` and `hidden`). Render the checklist container `#convChecklist` empty — `conversion-pane.js` populates it from `groupsForUi()`, one `<fieldset>` per group with a group select-all and a checkbox per row labeled `legacy ⇄ modern`, and a ⚠ marker when `row.lossy`.

- [ ] **Step 2: Extend `setMode()`** — add `var isConvertText = mode === "convertText";`, toggle the tab/panel `is-active`/`hidden`, include it in the non-poetry set so `settingsPanel`/`justifyActions`/`fontsStrip` hide (poetryMode stays `isTable || isConvert`), and add a status message branch.

- [ ] **Step 3: Wire the tab button** — `modeConvertText.addEventListener("click", function () { setMode("convertText"); });` and add its element ref alongside `modeStyles`.

- [ ] **Step 4: Bump `ASHAAR_ASSET_VERSION`** in `taskpane.html`.

- [ ] **Step 5: Syntax check + full test run** — `node --check src/taskpane/taskpane.js && npm test`.

- [ ] **Step 6: Commit** — `git add -A && git commit -m "feat(convert): Convert tab UI + mode wiring"`

---

### Task 5: Manual Word verification

**Files:** none (verification only). Record results in a checklist doc `docs/superpowers/specs/2026-07-16-text-conversion-manual-checklist.md`.

- [ ] Open a real AL-KANZ legacy `.docx`; **To Modern** whole-document → doubled consonants, escaped doubles, `؛→چھے`, sukun→khari zabar all convert; run formatting preserved; count line correct.
- [ ] **To Legacy** on the result (excluding lossy rows) → round-trips letters; genuine doubles gain a tatweel; `چھے` standalone → `؛` but not inside a word.
- [ ] Current-selection scope converts only the selection.
- [ ] Save a named preset (e.g. "Letters only"), reopen a different document → preset available (roaming) and applies.
- [ ] Lossy-row warning marker visible; unchecking sukun row leaves genuine sukuns intact.

## Self-Review

**Spec coverage:** §1 architecture → Tasks 1/3/4; §2 data model (letter/mark/symbol, modern-form rule, lossy flag) → Tasks 1–2; §3 extraction → Task 2; §4 engine/ordering/escape → Task 1; §5 UI → Task 4 (+ presets in Task 3); §6 safety (formatting via search-replace, lossy flag, feasibility spike) → Tasks 2–3; §7 testing → Tasks 1–2 (node) + Task 5 (manual); §8 non-goals → not built (correct). No gaps.

**Placeholder scan:** The `mark` high-jeem/high-noon exact chars are deliberately resolved by the Step-2 extraction inside Task 2 (a named, runnable spike with a concrete command), not left vague in code — acceptable. No "TBD/handle edge cases/write tests for the above" without code.

**Type consistency:** `buildOperations`/`convert`/`groupsForUi`/`MAPPINGS`/`DIRECTIONS`/`TATWEEL` names identical across Tasks 1–4; op object shape `{find,replaceWith,wholeWord,category}` consistent between Task 1 (producer) and Task 3 (consumer); DOM ids listed in Task 4 "Produces" match those read in Task 3.

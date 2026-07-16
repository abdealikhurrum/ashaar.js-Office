# Ashaar Styles Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a "Styles" tab to the task pane that provides named Word styles for prose documents (3 heading levels, Emphasis, Quote, Quran Quote), grouped into swappable "style groups" per use-case (Petition/Araz, Maqala, General, Waaz, plus custom), and a one-time RTL document setup action.

**Architecture:** A pure, node-testable data module (`word-styles.js`, UMD like `word-html.js`) owns the style-group recipe shape, defaults, and merge logic. A new orchestration module (`styles-pane.js`, following the `booklet-pane.js` precedent) owns all `Word.run()` calls: creating/reconfiguring the 6 named Word styles from the active group's recipe, applying a style to the current selection, instance-level overrides, group switching/storage, and the RTL setup action. `taskpane.html`/`taskpane.js` get the minimal glue to add a fourth mode tab, following exactly how the existing Booklet tab was wired in.

**Tech Stack:** Vanilla ES5 JavaScript (`var`/`function`, no arrow functions, no build step — matches every existing file in `src/taskpane/`), Office.js Word JavaScript API, Node `assert` for pure-module tests.

## Global Constraints

- **Feature requires `WordApiDesktop 1.3`.** Verified against the Word JavaScript API reference (`learn.microsoft.com/javascript/api/word`): `Style.borders`/`BorderCollection` (used for Quote's left/right borders) need `WordApiDesktop 1.1`; `Font.nameAscii`/`nameBidirectional`/`sizeBidirectional` and `PageSetup.sectionDirection` (used by the RTL setup action) need `WordApiDesktop 1.3`. `WordApiDesktop 1.3` covers both, and is the single gate to check. **This means the whole Styles tab is Desktop-Word-only (Windows/Mac) — unsupported on Word for the web.** Gate with `Office.context.requirements.isSetSupported("WordApiDesktop", "1.3")` and show a plain "Styles requires desktop Word" message otherwise, mirroring the existing `pl.width` WordApi-1.5 fallback pattern already in `insertTabStopPoem` (`taskpane.js:2592-2601`).
- **Known platform gap: no Office.js setter for paragraph/style bidi (reading order).** Confirmed by exhausting the Word JS API docs: `Word.ParagraphFormat` and `Word.Paragraph` expose `alignment`/`leftIndent`/`rightIndent`/`outlineLevel` etc., but reading-order/bidi (`Paragraph.ReadingOrder`, `RtlPara()`) exists **only** in the legacy VBA object model, not in Office.js. The RTL document setup action therefore implements Latin font, complex-script font+size, and section-level right-to-left layout — but **cannot** programmatically set the Normal style's paragraph bidi. This is documented in the UI (status text) rather than silently omitted. Do not attempt to work around this via raw OOXML injection into the styles part — Office.js has no supported way to replace/patch `word/styles.xml` of the currently-open document (unlike inserting OOXML into a *range*, which the existing `word-html.js`/`word-tabstop.js` already do for new content).
- **Never edit `src/vendor/`.** Pure modules stay DOM/Office-free and node-tested, per the existing project rule (CLAUDE.md).
- **UMD pattern for `word-styles.js`**, matching `word-tabstop.js`'s exact wrapper.
- **ES5 style**: `var`/`function` declarations only, no arrow functions, no `const`/`let` — matches every existing file in `src/taskpane/`.
- Every new test file must be added to the `test` script chain in `package.json`.
- Storage for group data (custom groups + which group is active) uses `Office.context.document.settings` (get/set/saveAsync), the same document-scoped mechanism the poem profile store already uses (`taskpane.js:416-427`) — so custom groups travel with the file.

---

### Task 1: `word-styles.js` — style group data model

**Files:**
- Create: `src/taskpane/word-styles.js`
- Test: `tests/word-styles.test.js`

**Interfaces:**
- Produces: `AshaarStyles.ROLES` (array of 6 role-key strings), `AshaarStyles.STYLE_NAME` (role key → Word style name string), `AshaarStyles.BASE_STYLE` (role key → Word built-in/ashaar style name it's based on), `AshaarStyles.STYLE_TYPE` (role key → `"Paragraph"|"Character"`), `AshaarStyles.defaultGroup(name)`, `AshaarStyles.mergeGroup(base, partial)`, `AshaarStyles.normalizeGroup(g)`, `AshaarStyles.BUILTIN_GROUPS` (object keyed by group name).

- [ ] **Step 1: Write the failing test for `defaultGroup`/`mergeGroup`/`normalizeGroup`**

Create `tests/word-styles.test.js`:

```js
const assert = require("assert");
const AshaarStyles = require("../src/taskpane/word-styles");

// ── ROLES / STYLE_NAME / BASE_STYLE / STYLE_TYPE ────────────────────────────

{
  assert.deepEqual(AshaarStyles.ROLES,
    ["heading1", "heading2", "heading3", "emphasis", "quote", "quranQuote"],
    "role order matters: quote must precede quranQuote (basedOn dependency)");
  assert.equal(AshaarStyles.STYLE_NAME.heading1, "Ashaar Heading 1");
  assert.equal(AshaarStyles.STYLE_NAME.quranQuote, "Ashaar Quran Quote");
  assert.equal(AshaarStyles.BASE_STYLE.heading1, "Heading 1");
  assert.equal(AshaarStyles.BASE_STYLE.emphasis, "Emphasis");
  assert.equal(AshaarStyles.BASE_STYLE.quote, "Quote");
  assert.equal(AshaarStyles.BASE_STYLE.quranQuote, "Ashaar Quote",
    "Quran Quote is based on OUR Quote style, not a Word built-in");
  assert.equal(AshaarStyles.STYLE_TYPE.emphasis, "Character");
  assert.equal(AshaarStyles.STYLE_TYPE.quote, "Paragraph");
}

// ── defaultGroup ─────────────────────────────────────────────────────────────

{
  const g = AshaarStyles.defaultGroup("General");
  assert.equal(g.name, "General");
  assert.equal(g.heading1.font, "Marjaan");
  assert.equal(typeof g.heading1.sizePt, "number");
  assert.ok(g.heading1.sizePt > g.heading2.sizePt, "heading1 larger than heading2");
  assert.ok(g.heading2.sizePt > g.heading3.sizePt, "heading2 larger than heading3");
  assert.equal(g.emphasis.color, "#FF0000");
  assert.equal(typeof g.emphasis.bumpPt, "number");
  assert.equal(g.quote.borderColor, "#000000");
  assert.equal(typeof g.quote.indentPt, "number");
  assert.equal(g.quranQuote.font, "Amiri Quran");
  assert.equal(g.quranQuote.lineHeightPt, null, "null = Word auto by default");
}

{
  // No name still normalizes to an empty-string-named group, not undefined.
  const g = AshaarStyles.defaultGroup();
  assert.equal(g.name, "");
}

// ── mergeGroup (deep-merge overrides onto a base, one level per role) ────────

{
  const base = AshaarStyles.defaultGroup("Base");
  const merged = AshaarStyles.mergeGroup(base, { heading1: { sizePt: 24 }, quote: { indentPt: 18 } });
  assert.equal(merged.heading1.sizePt, 24, "override applied");
  assert.equal(merged.heading1.font, base.heading1.font, "unset heading1 fields keep base");
  assert.equal(merged.quote.indentPt, 18);
  assert.equal(merged.quote.borderColor, base.quote.borderColor, "unset quote fields keep base");
  assert.equal(merged.emphasis.color, base.emphasis.color, "roles absent from partial are untouched");
  assert.notEqual(merged, base, "returns a new object, does not mutate base");
}

// ── normalizeGroup (fill missing roles/fields from defaults) ─────────────────

{
  const g = AshaarStyles.normalizeGroup({ name: "Partial", heading1: { sizePt: 30 } });
  assert.equal(g.name, "Partial");
  assert.equal(g.heading1.sizePt, 30, "keeps provided field");
  assert.equal(g.heading1.font, "Marjaan", "fills missing field from default");
  assert.ok(g.quote && typeof g.quote.indentPt === "number", "fills entirely-missing role");
}

// ── BUILTIN_GROUPS ────────────────────────────────────────────────────────────

{
  var names = Object.keys(AshaarStyles.BUILTIN_GROUPS);
  assert.deepEqual(names.sort(), ["General", "Maqala", "Petition", "Waaz"]);
  Object.keys(AshaarStyles.BUILTIN_GROUPS).forEach(function (k) {
    var g = AshaarStyles.BUILTIN_GROUPS[k];
    assert.equal(g.name, k, "built-in group's name matches its key");
    AshaarStyles.ROLES.forEach(function (role) {
      assert.ok(g[role], "built-in group '" + k + "' defines role '" + role + "'");
    });
  });
}

console.log("word-styles.test.js: all assertions passed");
```

- [ ] **Step 2: Run test to verify it fails**

Run: `node tests/word-styles.test.js`
Expected: `Error: Cannot find module '../src/taskpane/word-styles'`

- [ ] **Step 3: Write `word-styles.js`**

Create `src/taskpane/word-styles.js`:

```js
/**
 * AshaarStyles — style-group data model for the prose Styles tab (headings,
 * emphasis, block quotes, Quran quotes) and RTL document setup. Pure (no
 * Office.js/DOM); the Word.run() orchestration lives in styles-pane.js.
 *
 * See docs/superpowers/specs/2026-07-16-ashaar-styles-design.md.
 */
(function (root, factory) {
  if (typeof module !== "undefined" && module.exports) module.exports = factory();
  else root.AshaarStyles = factory();
}(typeof globalThis !== "undefined" ? globalThis : this, function () {
  "use strict";

  // Order matters: "quote" must be created/configured before "quranQuote",
  // since Ashaar Quran Quote is basedOn Ashaar Quote (not a Word built-in).
  var ROLES = ["heading1", "heading2", "heading3", "emphasis", "quote", "quranQuote"];

  var STYLE_NAME = {
    heading1: "Ashaar Heading 1",
    heading2: "Ashaar Heading 2",
    heading3: "Ashaar Heading 3",
    emphasis: "Ashaar Emphasis",
    quote: "Ashaar Quote",
    quranQuote: "Ashaar Quran Quote"
  };

  var BASE_STYLE = {
    heading1: "Heading 1",
    heading2: "Heading 2",
    heading3: "Heading 3",
    emphasis: "Emphasis",
    quote: "Quote",
    quranQuote: "Ashaar Quote"
  };

  var STYLE_TYPE = {
    heading1: "Paragraph",
    heading2: "Paragraph",
    heading3: "Paragraph",
    emphasis: "Character",
    quote: "Paragraph",
    quranQuote: "Paragraph"
  };

  function isObj(v) { return v && typeof v === "object" && !Array.isArray(v); }

  // The authoritative default group recipe. Headings are bold+centered
  // (fixed, not user-exposed); font/size are the adjustable style-level
  // fields. indentPt/lineHeightPt double as both the style-level default AND
  // the seed for a per-instance override (applied as direct formatting on
  // top of the named style).
  function defaultGroup(name) {
    return {
      name: typeof name === "string" ? name : "",
      heading1: { font: "Marjaan", sizePt: 18 },
      heading2: { font: "Marjaan", sizePt: 16 },
      heading3: { font: "Marjaan", sizePt: 14 },
      emphasis: { color: "#FF0000", bumpPt: 3 },
      quote: { borderColor: "#000000", borderWidth: "Pt050", indentPt: 0 },
      quranQuote: { font: "Amiri Quran", lineHeightPt: null } // null = Word auto
    };
  }

  // Shallow-merge `partial` onto `base` one level deep, per role. Returns a
  // new object; never mutates `base` or `partial`.
  function mergeGroup(base, partial) {
    var out = {};
    var b = base || {};
    var p = partial || {};
    out.name = ("name" in p) ? p.name : b.name;
    ROLES.forEach(function (role) {
      var br = isObj(b[role]) ? b[role] : {};
      var pr = isObj(p[role]) ? p[role] : {};
      var merged = {};
      Object.keys(br).forEach(function (k) { merged[k] = br[k]; });
      Object.keys(pr).forEach(function (k) { merged[k] = pr[k]; });
      out[role] = merged;
    });
    return out;
  }

  // Fill any missing roles/fields of `g` from the defaults (deep, via mergeGroup).
  function normalizeGroup(g) {
    return mergeGroup(defaultGroup((g && g.name) || ""), g || {});
  }

  var BUILTIN_GROUPS = {
    General: defaultGroup("General"),
    Petition: mergeGroup(defaultGroup("Petition"), {
      heading1: { sizePt: 16 }, heading2: { sizePt: 14 }, heading3: { sizePt: 12 },
      quote: { indentPt: 18 }
    }),
    Maqala: mergeGroup(defaultGroup("Maqala"), {
      heading1: { sizePt: 16 }, heading2: { sizePt: 14 }, heading3: { sizePt: 12 }
    }),
    Waaz: mergeGroup(defaultGroup("Waaz"), {
      heading1: { font: "Fatemi", sizePt: 20 }, heading2: { font: "Fatemi", sizePt: 17 }
    })
  };

  return {
    ROLES: ROLES,
    STYLE_NAME: STYLE_NAME,
    BASE_STYLE: BASE_STYLE,
    STYLE_TYPE: STYLE_TYPE,
    defaultGroup: defaultGroup,
    mergeGroup: mergeGroup,
    normalizeGroup: normalizeGroup,
    BUILTIN_GROUPS: BUILTIN_GROUPS
  };
}));
```

- [ ] **Step 4: Run test to verify it passes**

Run: `node tests/word-styles.test.js`
Expected: `word-styles.test.js: all assertions passed`

- [ ] **Step 5: Add to the test script chain**

Edit `package.json`'s `"test"` script, appending ` && node tests/word-styles.test.js` at the end of the existing chain.

- [ ] **Step 6: Commit**

```bash
git add src/taskpane/word-styles.js tests/word-styles.test.js package.json
git commit -m "Add word-styles.js style-group data model (roles, defaults, merge)"
```

---

### Task 2: `word-styles.js` — emphasis point-bump + role helpers

**Files:**
- Modify: `src/taskpane/word-styles.js`
- Test: `tests/word-styles.test.js`

**Interfaces:**
- Consumes: nothing new from Task 1 beyond what's already in the module.
- Produces: `AshaarStyles.computeEmphasisSize(baseSizePt, bumpPt)`, `AshaarStyles.clampIndentPt(pt)`, `AshaarStyles.clampLineHeightPt(pt)`.

- [ ] **Step 1: Write the failing test**

Append to `tests/word-styles.test.js` (before the final `console.log`):

```js
// ── computeEmphasisSize ───────────────────────────────────────────────────────

{
  assert.equal(AshaarStyles.computeEmphasisSize(12, 3), 15);
  assert.equal(AshaarStyles.computeEmphasisSize(10.5, 2.5), 13);
  assert.equal(AshaarStyles.computeEmphasisSize(null, 3), 15,
    "missing base size falls back to 12pt (Word's own default)");
  assert.equal(AshaarStyles.computeEmphasisSize(12, null), 12,
    "missing bump behaves as +0");
}

// ── clampIndentPt / clampLineHeightPt ────────────────────────────────────────

{
  assert.equal(AshaarStyles.clampIndentPt(18), 18);
  assert.equal(AshaarStyles.clampIndentPt(-5), 0, "negative indent clamps to 0");
  assert.equal(AshaarStyles.clampIndentPt(500), 200, "clamps to a 200pt ceiling");
  assert.equal(AshaarStyles.clampIndentPt(null), 0);

  assert.equal(AshaarStyles.clampLineHeightPt(null), null, "null (auto) passes through");
  assert.equal(AshaarStyles.clampLineHeightPt(24), 24);
  assert.equal(AshaarStyles.clampLineHeightPt(2), 6, "clamps to a 6pt floor");
  assert.equal(AshaarStyles.clampLineHeightPt(500), 200, "clamps to a 200pt ceiling");
}
```

- [ ] **Step 2: Run test to verify it fails**

Run: `node tests/word-styles.test.js`
Expected: `TypeError: AshaarStyles.computeEmphasisSize is not a function`

- [ ] **Step 3: Add the helpers**

In `src/taskpane/word-styles.js`, add above the final `return`:

```js
  // Emphasis has no style-level absolute size — only a bump amount. The
  // resulting absolute size is computed live from whatever the selection's
  // own base size already is (see §2 of the design spec).
  function computeEmphasisSize(baseSizePt, bumpPt) {
    var base = (typeof baseSizePt === "number" && baseSizePt > 0) ? baseSizePt : 12;
    var bump = (typeof bumpPt === "number") ? bumpPt : 0;
    return base + bump;
  }

  function clampIndentPt(pt) {
    var n = (typeof pt === "number" && !isNaN(pt)) ? pt : 0;
    return Math.max(0, Math.min(200, n));
  }

  // null means "Word auto" and must pass through unclamped.
  function clampLineHeightPt(pt) {
    if (pt == null) return null;
    var n = (typeof pt === "number" && !isNaN(pt)) ? pt : 0;
    return Math.max(6, Math.min(200, n));
  }
```

And add the three names to the returned object at the bottom of the factory function.

- [ ] **Step 4: Run test to verify it passes**

Run: `node tests/word-styles.test.js`
Expected: `word-styles.test.js: all assertions passed`

- [ ] **Step 5: Commit**

```bash
git add src/taskpane/word-styles.js tests/word-styles.test.js
git commit -m "Add emphasis point-bump and indent/line-height clamp helpers to word-styles.js"
```

---

### Task 3: Styles tab markup (HTML shell)

**Files:**
- Modify: `src/taskpane/taskpane.html`

**Interfaces:**
- Produces: DOM element IDs consumed by Task 4 (`mode-styles`, `styles-mode-panel`) and Task 5-9 (`styles-*` IDs listed below).

- [ ] **Step 1: Add the fourth mode-switch button**

In `src/taskpane/taskpane.html`, find:

```html
        <button id="mode-booklet" class="mode-button" type="button" role="tab" aria-selected="false" aria-controls="booklet-mode-panel">Booklet</button>
```

Add immediately after it:

```html
        <button id="mode-styles" class="mode-button" type="button" role="tab" aria-selected="false" aria-controls="styles-mode-panel">Styles</button>
```

- [ ] **Step 2: Add the panel markup**

Find the end of the booklet panel's closing `</section>` (search for `booklet-mode-panel` to locate it) and insert this new section immediately after it, before the closing `</main>`:

```html
      <section id="styles-mode-panel" class="mode-panel" role="tabpanel" aria-labelledby="mode-styles" hidden>
        <div class="section-title">
          <h2>Styles</h2>
          <p>Named Word styles for prose documents — headings, emphasis, block quotes, and Quran quotes — grouped by use case. Requires desktop Word (Windows/Mac).</p>
        </div>
        <p id="styles-unsupported" class="adopt-hint" hidden>This add-in's Styles tab requires desktop Word (Windows or Mac) — it isn't available on Word for the web.</p>

        <div id="styles-body">
          <section class="sp-profile-row">
            <label for="styles-group-select">Style group</label>
            <select id="styles-group-select"></select>
            <button id="styles-group-saveas" type="button" class="button--secondary">Save as new group&hellip;</button>
          </section>
          <div id="styles-saveas-row" class="sp-profile-row" hidden>
            <label for="styles-saveas-name">Name</label>
            <input id="styles-saveas-name" type="text" class="template-name-input" placeholder="New group name">
            <button id="styles-saveas-ok" type="button">Save</button>
            <button id="styles-saveas-cancel" type="button" class="button--secondary">Cancel</button>
          </div>

          <details class="import-panel" open>
            <summary>Heading 1</summary>
            <div class="field"><label for="styles-h1-font">Font</label>
              <input id="styles-h1-font" type="text" value="Marjaan"></div>
            <div class="field"><label for="styles-h1-size">Size (pt)</label>
              <input id="styles-h1-size" type="number" min="6" max="72" value="18"></div>
            <button id="styles-h1-update" type="button" class="button--secondary">Update style</button>
            <button id="styles-h1-apply" type="button">Apply to selection</button>
          </details>

          <details class="import-panel">
            <summary>Heading 2</summary>
            <div class="field"><label for="styles-h2-font">Font</label>
              <input id="styles-h2-font" type="text" value="Marjaan"></div>
            <div class="field"><label for="styles-h2-size">Size (pt)</label>
              <input id="styles-h2-size" type="number" min="6" max="72" value="16"></div>
            <button id="styles-h2-update" type="button" class="button--secondary">Update style</button>
            <button id="styles-h2-apply" type="button">Apply to selection</button>
          </details>

          <details class="import-panel">
            <summary>Heading 3</summary>
            <div class="field"><label for="styles-h3-font">Font</label>
              <input id="styles-h3-font" type="text" value="Marjaan"></div>
            <div class="field"><label for="styles-h3-size">Size (pt)</label>
              <input id="styles-h3-size" type="number" min="6" max="72" value="14"></div>
            <button id="styles-h3-update" type="button" class="button--secondary">Update style</button>
            <button id="styles-h3-apply" type="button">Apply to selection</button>
          </details>

          <details class="import-panel">
            <summary>Emphasis</summary>
            <div class="field"><label for="styles-emphasis-color">Color</label>
              <input id="styles-emphasis-color" type="text" value="#FF0000"></div>
            <div class="field"><label for="styles-emphasis-bump">Size bump (pt)</label>
              <input id="styles-emphasis-bump" type="number" min="0" max="20" value="3"></div>
            <button id="styles-emphasis-update" type="button" class="button--secondary">Update style</button>
            <button id="styles-emphasis-apply" type="button">Apply to selection</button>
          </details>

          <details class="import-panel">
            <summary>Quote</summary>
            <div class="field"><label for="styles-quote-color">Border color</label>
              <input id="styles-quote-color" type="text" value="#000000"></div>
            <div class="field"><label for="styles-quote-width">Border width</label>
              <select id="styles-quote-width">
                <option value="Pt025">0.25pt</option>
                <option value="Pt050" selected>0.5pt</option>
                <option value="Pt075">0.75pt</option>
                <option value="Pt100">1pt</option>
                <option value="Pt150">1.5pt</option>
              </select></div>
            <div class="field"><label for="styles-quote-indent">Default indent (pt)</label>
              <input id="styles-quote-indent" type="number" min="0" max="200" value="0"></div>
            <button id="styles-quote-update" type="button" class="button--secondary">Update style</button>
            <button id="styles-quote-apply" type="button">Apply to selection</button>
            <div class="field"><label for="styles-quote-indent-override">This selection's indent (pt)</label>
              <input id="styles-quote-indent-override" type="number" min="0" max="200" placeholder="leave blank to use default"></div>
            <button id="styles-quote-override-apply" type="button" class="button--secondary">Apply indent to this selection</button>
          </details>

          <details class="import-panel">
            <summary>Quran Quote</summary>
            <div class="field"><label for="styles-quranquote-font">Font</label>
              <input id="styles-quranquote-font" type="text" value="Amiri Quran" list="styles-quranquote-font-list">
              <datalist id="styles-quranquote-font-list">
                <option value="Amiri Quran">
                <option value="KFGQPC Uthman Taha Naskh">
              </datalist></div>
            <div class="field"><label for="styles-quranquote-lh">Default line height (pt, blank = Word auto)</label>
              <input id="styles-quranquote-lh" type="number" min="6" max="200" placeholder="auto"></div>
            <button id="styles-quranquote-update" type="button" class="button--secondary">Update style</button>
            <button id="styles-quranquote-apply" type="button">Apply to selection</button>
            <div class="field"><label for="styles-quranquote-lh-override">This selection's line height (pt)</label>
              <input id="styles-quranquote-lh-override" type="number" min="6" max="200" placeholder="leave blank to use default"></div>
            <button id="styles-quranquote-override-apply" type="button" class="button--secondary">Apply line height to this selection</button>
          </details>

          <details class="import-panel">
            <summary>RTL document setup</summary>
            <p id="styles-rtl-status" class="adopt-hint">Not yet applied in this document.</p>
            <div class="field"><label for="styles-rtl-latin-font">Latin font</label>
              <input id="styles-rtl-latin-font" type="text" value="Times New Roman"></div>
            <div class="field"><label for="styles-rtl-cs-font">Complex-script font</label>
              <input id="styles-rtl-cs-font" type="text" value="Fatemi Maqala"></div>
            <div class="field"><label for="styles-rtl-cs-size">Complex-script size (pt)</label>
              <input id="styles-rtl-cs-size" type="number" min="6" max="72" value="12"></div>
            <button id="styles-rtl-apply" type="button">Set up RTL document</button>
            <p class="adopt-hint">Note: Word's automatic RTL detection handles paragraph direction for typed Arabic/Urdu/Persian text. Office.js has no supported way to force Normal-style paragraph direction (a VBA-only capability) — if a blank paragraph's cursor/list behavior looks wrong, use Word's own Layout &rarr; Paragraph Direction control.</p>
          </details>
        </div>
      </section>
```

- [ ] **Step 3: Add the new script tags**

Find the script-list array in `taskpane.html` (search for `"./booklet-pane.js"`) and add two entries immediately before `"./taskpane.js"`:

```js
          "./word-styles.js",
          "./styles-pane.js",
          "./taskpane.js"
```

(`word-styles.js` before `styles-pane.js`, since the latter consumes the former; both before `taskpane.js`, matching every other module in the list.)

- [ ] **Step 4: Manual check — page loads without console errors**

Run `npm run dev-server`, open `https://localhost:3000/taskpane.html` directly in a browser (not inside Word yet — `styles-pane.js` in the next tasks guards all Office.js calls behind `Office.onReady`, so the tab should render even outside Word). Confirm: the "Styles" button appears in the mode switch, clicking it is inert for now (wiring comes in Task 4), and the browser console shows no script errors from the new `<script>` tags (missing files would 404; `word-styles.js` exists from Task 1, `styles-pane.js` does not exist yet — expect a 404 for it specifically, which is fine until Task 5 creates the file. If this 404 breaks the whole script-loading sequence for `taskpane.js`, create an empty placeholder `src/taskpane/styles-pane.js` containing just `(function () {}());` for now, and remove the placeholder body in Task 5).

- [ ] **Step 5: Commit**

```bash
git add src/taskpane/taskpane.html
git commit -m "Add Styles tab markup and script wiring to taskpane.html"
```

---

### Task 4: Wire the Styles mode tab into `taskpane.js`

**Files:**
- Modify: `src/taskpane/taskpane.js:1-16` (DOM refs), `taskpane.js:309-328` (`setMode`), `taskpane.js:4904-4920` (`bind`)

**Interfaces:**
- Consumes: `mode-styles`, `styles-mode-panel` element IDs from Task 3.
- Produces: `setMode("styles")` as a valid call other code can make.

- [ ] **Step 1: Add DOM refs**

In `src/taskpane/taskpane.js`, near line 8 (`var modeBooklet = document.getElementById("mode-booklet");`), add:

```js
  var modeStyles = document.getElementById("mode-styles");
```

Near line 11 (`var bookletPanel = document.getElementById("booklet-mode-panel");`), add:

```js
  var stylesPanel = document.getElementById("styles-mode-panel");
```

- [ ] **Step 2: Extend `setMode`**

Replace the `setMode` function (`taskpane.js:309-328`):

```js
  function setMode(mode) {
    var isTable = mode === "table";
    var isConvert = mode === "convert";
    var isBooklet = mode === "booklet";
    var isStyles = mode === "styles";
    modeTable.classList.toggle("is-active", isTable);
    modeConvert.classList.toggle("is-active", isConvert);
    modeBooklet.classList.toggle("is-active", isBooklet);
    modeStyles.classList.toggle("is-active", isStyles);
    modeTable.setAttribute("aria-selected", String(isTable));
    modeConvert.setAttribute("aria-selected", String(isConvert));
    modeBooklet.setAttribute("aria-selected", String(isBooklet));
    modeStyles.setAttribute("aria-selected", String(isStyles));
    tablePanel.classList.toggle("is-active", isTable);
    convertPanel.classList.toggle("is-active", isConvert);
    bookletPanel.classList.toggle("is-active", isBooklet);
    stylesPanel.classList.toggle("is-active", isStyles);
    tablePanel.hidden = !isTable;
    convertPanel.hidden = !isConvert;
    bookletPanel.hidden = !isBooklet;
    stylesPanel.hidden = !isStyles;
    setMessage(isTable ? "Table input mode: draw a blank grid, then type in Word."
      : isConvert ? "Ashaar.js conversion mode: paste source text, then insert a converted table."
      : isBooklet ? "Booklet mode: impose the open document into a print-ready booklet."
      : "Styles mode: apply named heading/quote/emphasis styles, grouped by document use case.");
  }
```

- [ ] **Step 3: Wire the click handler and (if present) an init hook**

Near `taskpane.js:4913-4915`, add after the `modeBooklet.addEventListener(...)` block:

```js
    modeStyles.addEventListener("click", function () {
      setMode("styles");
      if (typeof AshaarStylesPane !== "undefined" && AshaarStylesPane.onTabShown) {
        AshaarStylesPane.onTabShown();
      }
    });
```

(`AshaarStylesPane.onTabShown` is defined in Task 5 — it's a no-op-safe optional hook so this task compiles/runs even before Task 5 lands, via the `typeof` guard.)

- [ ] **Step 4: Manual check**

Reload the task pane in the browser (outside Word is fine for this check). Click "Styles" — the tab should become active, its panel visible, and the status message should update. Click back to "Table Input" — confirm the Styles panel hides again and nothing else broke.

- [ ] **Step 5: Commit**

```bash
git add src/taskpane/taskpane.js
git commit -m "Wire Styles mode tab into taskpane.js setMode/bind"
```

---

### Task 5: `styles-pane.js` — capability gate + `ensureAshaarStyles`

**Files:**
- Create (replacing the Task-3 placeholder if you made one): `src/taskpane/styles-pane.js`

**Interfaces:**
- Consumes: `AshaarStyles.ROLES`/`STYLE_NAME`/`BASE_STYLE`/`STYLE_TYPE`/`BUILTIN_GROUPS`/`normalizeGroup` (Task 1).
- Produces: global `AshaarStylesPane` object with `onTabShown()` (Task 4 dependency), `ensureAshaarStyles(context, group)` (internal, used by Tasks 6-9), `isDesktopCapable()`.

- [ ] **Step 1: Write the module skeleton, capability gate, and style-creation core**

Create `src/taskpane/styles-pane.js`:

```js
/* StylesPane — wires the "Styles" tab to AshaarStyles (word-styles.js).
 * Creates/reconfigures 6 named Word styles from the active style group's
 * recipe, applies them to the current selection, and runs the RTL document
 * setup action. See docs/superpowers/specs/2026-07-16-ashaar-styles-design.md.
 */
(function () {
  "use strict";

  var els = {};
  var bound = false;
  var activeGroupName = "General";
  var groupStore = {}; // name -> group recipe (built-ins + custom, merged at load time)

  function byId(id) { return document.getElementById(id); }

  function setStatus(el, msg, warn) {
    if (!el) return;
    el.textContent = msg;
    el.classList.toggle("warn", !!warn);
  }

  // WordApiDesktop 1.3 covers everything this tab needs: Style.borders
  // (WordApiDesktop 1.1) and Font.nameAscii/nameBidirectional/sizeBidirectional
  // + PageSetup.sectionDirection (WordApiDesktop 1.3). Word for the web does
  // not support this requirement set at all.
  function isDesktopCapable() {
    return !!(window.Office && Office.context && Office.context.requirements &&
      Office.context.requirements.isSetSupported("WordApiDesktop", "1.3"));
  }

  // Idempotent: reuse the style if it already exists, else create it.
  function ensureStyle(context, name, type) {
    var style = context.document.getStyles().getByNameOrNullObject(name);
    style.load("isNullObject");
    return context.sync().then(function () {
      if (style.isNullObject) {
        return context.document.addStyle(name, type);
      }
      return style;
    });
  }

  // Configure one role's style object from the group recipe. Split into its
  // own function (rather than inlined in ensureAshaarStyles) so Task 8/9 can
  // reuse the font-string application logic if needed.
  function configureRoleStyle(style, role, group) {
    var recipe = group[role];
    style.baseStyle = AshaarStyles.BASE_STYLE[role];
    style.unhideWhenUsed = true; // visible in Word's own Style gallery once used

    if (role === "heading1" || role === "heading2" || role === "heading3") {
      style.font.nameAscii = recipe.font;
      style.font.nameBidirectional = recipe.font;
      style.font.size = recipe.sizePt;
      style.font.sizeBidirectional = recipe.sizePt;
      style.font.bold = true;
      style.paragraphFormat.alignment = Word.Alignment.centered;
    } else if (role === "emphasis") {
      style.font.color = recipe.color;
      // No absolute size here — Emphasis's bump is computed live per
      // instance in Task 8 (computeEmphasisSize), not stored as a style size.
    } else if (role === "quote") {
      style.paragraphFormat.leftIndent = AshaarStyles.clampIndentPt(recipe.indentPt);
      style.paragraphFormat.rightIndent = AshaarStyles.clampIndentPt(recipe.indentPt);
      var leftBorder = style.borders.getByLocation(Word.BorderLocation.left);
      var rightBorder = style.borders.getByLocation(Word.BorderLocation.right);
      leftBorder.type = Word.BorderType.single;
      leftBorder.width = recipe.borderWidth;
      leftBorder.color = recipe.borderColor;
      rightBorder.type = Word.BorderType.single;
      rightBorder.width = recipe.borderWidth;
      rightBorder.color = recipe.borderColor;
    } else if (role === "quranQuote") {
      style.font.nameAscii = recipe.font;
      style.font.nameBidirectional = recipe.font;
      // lineSpacing has no clean "reset to Word auto" value in the object
      // model, so switching a group FROM a set line height back to null
      // (auto) will leave the previous numeric value in place rather than
      // truly reverting to auto. Flagged for the manual check in Task 9 —
      // if this matters in practice, the fix is to also set
      // style.paragraphFormat.lineSpacingRule = Word.LineSpacing.single
      // as the "auto" case, which needs confirming live before relying on it.
      if (recipe.lineHeightPt != null) {
        style.paragraphFormat.lineSpacing = AshaarStyles.clampLineHeightPt(recipe.lineHeightPt);
      }
    }
  }

  // Create/reconfigure all 6 named styles from `group`'s recipe. ROLES order
  // (quote before quranQuote) guarantees Ashaar Quote exists before Ashaar
  // Quran Quote's baseStyle references it.
  function ensureAshaarStyles(context, group) {
    var styleObjs = {};
    var chain = Promise.resolve();
    AshaarStyles.ROLES.forEach(function (role) {
      chain = chain.then(function () {
        return ensureStyle(context, AshaarStyles.STYLE_NAME[role], AshaarStyles.STYLE_TYPE[role]);
      }).then(function (style) {
        style.load("baseStyle"); // harmless preload; configureRoleStyle overwrites it
        styleObjs[role] = style;
        return context.sync();
      }).then(function () {
        configureRoleStyle(styleObjs[role], role, group);
        return context.sync();
      });
    });
    return chain.then(function () { return styleObjs; });
  }

  // Public: called when the Styles tab is first shown (Task 4's onTabShown
  // hook) and again whenever the group picker changes (Task 6).
  function onTabShown() {
    if (!els.body) cacheEls();
    if (!isDesktopCapable()) {
      els.unsupported.hidden = false;
      els.body.hidden = true;
      return;
    }
    els.unsupported.hidden = true;
    els.body.hidden = false;
  }

  function cacheEls() {
    els.unsupported = byId("styles-unsupported");
    els.body = byId("styles-body");
  }

  window.AshaarStylesPane = {
    onTabShown: onTabShown,
    isDesktopCapable: isDesktopCapable,
    ensureAshaarStyles: ensureAshaarStyles // exposed for Tasks 6-9 in this same file
  };
}());
```

- [ ] **Step 2: Manual check — styles get created in a live document**

This step cannot be automated (it drives live Office.js). In Word desktop (`npm start`), open the task pane, click the Styles tab (confirms Task 4 didn't regress), then in the browser devtools console (task pane webview) run:

```js
Word.run(function (context) {
  return AshaarStylesPane.ensureAshaarStyles(context, AshaarStyles.BUILTIN_GROUPS.General).then(function () {
    return context.sync();
  });
});
```

Open Word's Styles pane (Ctrl+Shift+S or Alt+Ctrl+Shift+S) — confirm "Ashaar Heading 1", "Ashaar Heading 2", "Ashaar Heading 3", "Ashaar Emphasis", "Ashaar Quote", "Ashaar Quran Quote" all appear. Type a line of text, apply "Ashaar Heading 1" from Word's own Style gallery — confirm it renders bold, centered, 18pt. Apply "Ashaar Quote" to a paragraph — confirm left+right borders appear.

- [ ] **Step 3: Commit**

```bash
git add src/taskpane/styles-pane.js
git commit -m "Add styles-pane.js capability gate and ensureAshaarStyles style creation"
```

---

### Task 6: Group picker + document-scoped storage

**Files:**
- Modify: `src/taskpane/styles-pane.js`

**Interfaces:**
- Consumes: `AshaarStyles.BUILTIN_GROUPS`/`normalizeGroup` (Task 1), `ensureAshaarStyles` (Task 5).
- Produces: `activeGroup()` (returns the currently-selected group recipe), populated `#styles-group-select`.

- [ ] **Step 1: Add storage helpers and picker population**

In `src/taskpane/styles-pane.js`, add (near the top, after `groupStore`/`activeGroupName` declarations):

```js
  var GROUP_STORE_KEY = "ashaar-style-groups"; // custom groups only, keyed by name
  var ACTIVE_GROUP_KEY = "ashaar-style-active-group";

  function loadGroupStore() {
    try {
      var raw = Office.context.document.settings.get(GROUP_STORE_KEY);
      return raw ? JSON.parse(raw) : {};
    } catch (e) { return {}; }
  }

  function saveGroupStore(store, done) {
    Office.context.document.settings.set(GROUP_STORE_KEY, JSON.stringify(store || {}));
    Office.context.document.settings.saveAsync(function () { if (done) done(); });
  }

  function loadActiveGroupName() {
    var raw = Office.context.document.settings.get(ACTIVE_GROUP_KEY);
    return (typeof raw === "string" && raw) ? raw : "General";
  }

  function saveActiveGroupName(name, done) {
    Office.context.document.settings.set(ACTIVE_GROUP_KEY, name);
    Office.context.document.settings.saveAsync(function () { if (done) done(); });
  }

  // All group names available in the picker: built-ins first, then custom
  // (custom groups can shadow a built-in name; custom wins).
  function allGroups() {
    var out = {};
    Object.keys(AshaarStyles.BUILTIN_GROUPS).forEach(function (k) { out[k] = AshaarStyles.BUILTIN_GROUPS[k]; });
    Object.keys(groupStore).forEach(function (k) { out[k] = AshaarStyles.normalizeGroup(groupStore[k]); });
    return out;
  }

  function activeGroup() {
    var groups = allGroups();
    return groups[activeGroupName] || AshaarStyles.BUILTIN_GROUPS.General;
  }

  function populateGroupPicker() {
    var select = byId("styles-group-select");
    if (!select) return;
    select.innerHTML = "";
    var groups = allGroups();
    Object.keys(groups).sort().forEach(function (name) {
      var opt = document.createElement("option");
      opt.value = name;
      opt.textContent = name;
      if (name === activeGroupName) opt.selected = true;
      select.appendChild(opt);
    });
  }

  // Re-syncs Word's actual style definitions to match the active group, then
  // re-populates form fields from it (Task 7/8 read these same fields).
  function applyActiveGroupToDocument(then) {
    Word.run(function (context) {
      return AshaarStylesPane.ensureAshaarStyles(context, activeGroup()).then(function () {
        return context.sync();
      });
    }).then(function () {
      if (then) then();
    }).catch(function (e) {
      setStatus(byId("styles-rtl-status"), "Error applying style group: " + (e.message || String(e)), true);
    });
  }
```

- [ ] **Step 2: Wire into `onTabShown` and bind the picker's change event**

Replace `onTabShown` in `src/taskpane/styles-pane.js`:

```js
  function onTabShown() {
    if (!els.body) cacheEls();
    if (!isDesktopCapable()) {
      els.unsupported.hidden = false;
      els.body.hidden = true;
      return;
    }
    els.unsupported.hidden = true;
    els.body.hidden = false;
    if (bound) return;
    bound = true;
    groupStore = loadGroupStore();
    activeGroupName = loadActiveGroupName();
    populateGroupPicker();
    byId("styles-group-select").addEventListener("change", function (e) {
      activeGroupName = e.target.value;
      saveActiveGroupName(activeGroupName);
      applyActiveGroupToDocument();
    });
    applyActiveGroupToDocument();
  }
```

- [ ] **Step 3: Manual check**

In Word desktop, open the Styles tab. Confirm the group picker lists "General, Maqala, Petition, Waaz" alphabetically, "General" selected. Switch to "Waaz" — confirm (via `Word.run` devtools probe as in Task 5 Step 2, or by applying "Ashaar Heading 1" to a test paragraph) that the style now uses Fatemi/20pt rather than Marjaan/18pt. Close and reopen the document — confirm the picker still shows "Waaz" as selected (proves `Office.context.document.settings` persisted it).

- [ ] **Step 4: Commit**

```bash
git add src/taskpane/styles-pane.js
git commit -m "Add style-group picker and document-scoped group storage"
```

---

### Task 7: Save as new group

**Files:**
- Modify: `src/taskpane/styles-pane.js`

**Interfaces:**
- Consumes: `#styles-group-saveas`, `#styles-saveas-row`, `#styles-saveas-name`, `#styles-saveas-ok`, `#styles-saveas-cancel` (Task 3), `groupStore`/`saveGroupStore`/`populateGroupPicker` (Task 6).
- Produces: a captured-from-current-fields group recipe added to `groupStore`.

- [ ] **Step 1: Add a "read current fields into a group recipe" helper**

This reads the same fields Task 8's "Update style" buttons write from, so add it once here and let Task 8 reuse it:

```js
  function readFieldsIntoGroup(name) {
    return {
      name: name,
      heading1: { font: byId("styles-h1-font").value, sizePt: Number(byId("styles-h1-size").value) },
      heading2: { font: byId("styles-h2-font").value, sizePt: Number(byId("styles-h2-size").value) },
      heading3: { font: byId("styles-h3-font").value, sizePt: Number(byId("styles-h3-size").value) },
      emphasis: { color: byId("styles-emphasis-color").value, bumpPt: Number(byId("styles-emphasis-bump").value) },
      quote: {
        borderColor: byId("styles-quote-color").value,
        borderWidth: byId("styles-quote-width").value,
        indentPt: Number(byId("styles-quote-indent").value)
      },
      quranQuote: {
        font: byId("styles-quranquote-font").value,
        lineHeightPt: byId("styles-quranquote-lh").value === "" ? null : Number(byId("styles-quranquote-lh").value)
      }
    };
  }

  function populateFieldsFromGroup(group) {
    byId("styles-h1-font").value = group.heading1.font;
    byId("styles-h1-size").value = group.heading1.sizePt;
    byId("styles-h2-font").value = group.heading2.font;
    byId("styles-h2-size").value = group.heading2.sizePt;
    byId("styles-h3-font").value = group.heading3.font;
    byId("styles-h3-size").value = group.heading3.sizePt;
    byId("styles-emphasis-color").value = group.emphasis.color;
    byId("styles-emphasis-bump").value = group.emphasis.bumpPt;
    byId("styles-quote-color").value = group.quote.borderColor;
    byId("styles-quote-width").value = group.quote.borderWidth;
    byId("styles-quote-indent").value = group.quote.indentPt;
    byId("styles-quranquote-font").value = group.quranQuote.font;
    byId("styles-quranquote-lh").value = group.quranQuote.lineHeightPt == null ? "" : group.quranQuote.lineHeightPt;
  }
```

- [ ] **Step 2: Wire the Save-as row**

Add to `onTabShown`'s one-time binding block (inside the `if (!bound) { ... }` body from Task 6 Step 2), and also call `populateFieldsFromGroup(activeGroup())` right after `applyActiveGroupToDocument()`:

```js
    byId("styles-group-saveas").addEventListener("click", function () {
      byId("styles-saveas-row").hidden = false;
      byId("styles-saveas-name").value = "";
      byId("styles-saveas-name").focus();
    });
    byId("styles-saveas-cancel").addEventListener("click", function () {
      byId("styles-saveas-row").hidden = true;
    });
    byId("styles-saveas-ok").addEventListener("click", function () {
      var name = String(byId("styles-saveas-name").value || "").trim();
      if (!name) return;
      groupStore[name] = readFieldsIntoGroup(name);
      saveGroupStore(groupStore, function () {
        byId("styles-saveas-row").hidden = true;
        activeGroupName = name;
        saveActiveGroupName(activeGroupName);
        populateGroupPicker();
      });
    });
```

And change the group-picker `change` handler (Task 6 Step 2) to also repopulate fields:

```js
    byId("styles-group-select").addEventListener("change", function (e) {
      activeGroupName = e.target.value;
      saveActiveGroupName(activeGroupName);
      populateFieldsFromGroup(activeGroup());
      applyActiveGroupToDocument();
    });
```

- [ ] **Step 3: Manual check**

In the Styles tab, change Heading 1's font to "Fatemi" and size to 22, click "Save as new group…", name it "MyGroup", Save. Confirm the group picker now shows "MyGroup" selected. Switch to "General" and back to "MyGroup" — confirm the Heading 1 fields show Fatemi/22 again (proves round-trip through `Office.context.document.settings`). Reopen the document — confirm "MyGroup" is still in the picker.

- [ ] **Step 4: Commit**

```bash
git add src/taskpane/styles-pane.js
git commit -m "Add Save-as-new-group capture for the Styles tab"
```

---

### Task 8: Apply styles to selection (6 roles + instance overrides)

**Files:**
- Modify: `src/taskpane/styles-pane.js`

**Interfaces:**
- Consumes: `AshaarStyles.STYLE_NAME`/`computeEmphasisSize`/`clampIndentPt`/`clampLineHeightPt` (Tasks 1-2), `readFieldsIntoGroup`/`activeGroup` (Tasks 6-7), the 6 role sections' "Update style" and "Apply to selection" buttons plus the 2 override buttons (Task 3).

- [ ] **Step 1: Add "Update style" handlers for the 6 roles**

Each writes the current field values into the active group, persists it (built-in groups get shadowed into `groupStore` under their own name on first edit — same as any custom group), and re-runs `ensureAshaarStyles`:

```js
  function updateActiveGroupFromFields() {
    var updated = readFieldsIntoGroup(activeGroupName);
    groupStore[activeGroupName] = updated;
    saveGroupStore(groupStore, function () {
      applyActiveGroupToDocument();
    });
  }

  function bindUpdateButtons() {
    ["styles-h1-update", "styles-h2-update", "styles-h3-update",
      "styles-emphasis-update", "styles-quote-update", "styles-quranquote-update"
    ].forEach(function (id) {
      byId(id).addEventListener("click", updateActiveGroupFromFields);
    });
  }
```

Call `bindUpdateButtons();` once inside the one-time binding block in `onTabShown`.

- [ ] **Step 2: Add "Apply to selection" for the paragraph-style roles (headings, quote, quranQuote)**

```js
  function applyParagraphStyle(styleName) {
    Word.run(function (context) {
      var selection = context.document.getSelection();
      var paragraphs = selection.paragraphs;
      paragraphs.load("items");
      return context.sync().then(function () {
        paragraphs.items.forEach(function (p) { p.style = styleName; });
        return context.sync();
      });
    }).catch(function (e) {
      setStatus(byId("styles-rtl-status"), "Error applying style: " + (e.message || String(e)), true);
    });
  }

  function bindParagraphApplyButtons() {
    byId("styles-h1-apply").addEventListener("click", function () { applyParagraphStyle(AshaarStyles.STYLE_NAME.heading1); });
    byId("styles-h2-apply").addEventListener("click", function () { applyParagraphStyle(AshaarStyles.STYLE_NAME.heading2); });
    byId("styles-h3-apply").addEventListener("click", function () { applyParagraphStyle(AshaarStyles.STYLE_NAME.heading3); });
    byId("styles-quote-apply").addEventListener("click", function () { applyParagraphStyle(AshaarStyles.STYLE_NAME.quote); });
    byId("styles-quranquote-apply").addEventListener("click", function () { applyParagraphStyle(AshaarStyles.STYLE_NAME.quranQuote); });
  }
```

Call `bindParagraphApplyButtons();` once inside the one-time binding block.

- [ ] **Step 3: Add "Apply to selection" for Emphasis (character style + live size bump)**

Emphasis is the one role that needs its resulting absolute size computed from whatever the selection's current size already is, per the design spec:

```js
  function applyEmphasis() {
    var bumpPt = Number(byId("styles-emphasis-bump").value) || 0;
    var color = byId("styles-emphasis-color").value;
    Word.run(function (context) {
      var selection = context.document.getSelection();
      selection.font.load("size");
      return context.sync().then(function () {
        var resultSize = AshaarStyles.computeEmphasisSize(selection.font.size, bumpPt);
        selection.style = AshaarStyles.STYLE_NAME.emphasis;
        selection.font.color = color;
        selection.font.size = resultSize;
        return context.sync();
      });
    }).catch(function (e) {
      setStatus(byId("styles-rtl-status"), "Error applying Emphasis: " + (e.message || String(e)), true);
    });
  }
```

Bind: `byId("styles-emphasis-apply").addEventListener("click", applyEmphasis);` inside the one-time binding block.

- [ ] **Step 4: Add the two instance-override buttons (Quote indent, Quran Quote line height)**

```js
  function applyQuoteIndentOverride() {
    var raw = byId("styles-quote-indent-override").value;
    if (raw === "") return; // blank = no override requested
    var pt = AshaarStyles.clampIndentPt(Number(raw));
    Word.run(function (context) {
      var paragraphs = context.document.getSelection().paragraphs;
      paragraphs.load("items");
      return context.sync().then(function () {
        paragraphs.items.forEach(function (p) { p.leftIndent = pt; p.rightIndent = pt; });
        return context.sync();
      });
    }).catch(function (e) {
      setStatus(byId("styles-rtl-status"), "Error applying indent override: " + (e.message || String(e)), true);
    });
  }

  function applyQuranQuoteLineHeightOverride() {
    var raw = byId("styles-quranquote-lh-override").value;
    if (raw === "") return;
    var pt = AshaarStyles.clampLineHeightPt(Number(raw));
    Word.run(function (context) {
      var paragraphs = context.document.getSelection().paragraphs;
      paragraphs.load("items");
      return context.sync().then(function () {
        paragraphs.items.forEach(function (p) { p.lineSpacing = pt; });
        return context.sync();
      });
    }).catch(function (e) {
      setStatus(byId("styles-rtl-status"), "Error applying line-height override: " + (e.message || String(e)), true);
    });
  }
```

Bind both to their `-override-apply` buttons inside the one-time binding block.

- [ ] **Step 5: Manual check**

In Word desktop: type a plain paragraph, select it, click Heading 1's "Apply to selection" — confirm bold/centered/18pt Marjaan. Select a word mid-sentence, click Emphasis's "Apply to selection" — confirm it turns red and grows by the bump amount relative to its *surrounding* text's size (test this on a sentence where you've manually set a different base size first, to confirm the bump is relative, not absolute). Select a paragraph, apply Quote — confirm left/right borders. Enter `12` into "This selection's indent" and click its override button — confirm just that paragraph narrows, while a second Quote paragraph elsewhere keeps the group's default indent. Repeat the analogous check for Quran Quote's line-height override.

- [ ] **Step 6: Commit**

```bash
git add src/taskpane/styles-pane.js
git commit -m "Add apply-to-selection for the 6 style roles and instance overrides"
```

---

### Task 9: RTL document setup action

**Files:**
- Modify: `src/taskpane/styles-pane.js`

**Interfaces:**
- Consumes: `#styles-rtl-latin-font`, `#styles-rtl-cs-font`, `#styles-rtl-cs-size`, `#styles-rtl-apply`, `#styles-rtl-status` (Task 3).

- [ ] **Step 1: Implement the setup action**

Per the Global Constraints note: this sets the Normal style's Latin/complex-script fonts+size and the document's section-level RTL layout. It explicitly does NOT (cannot) set Normal-style paragraph bidi — no Office.js API exists for that.

```js
  function runRtlSetup() {
    var latinFont = byId("styles-rtl-latin-font").value;
    var csFont = byId("styles-rtl-cs-font").value;
    var csSize = Number(byId("styles-rtl-cs-size").value) || 12;
    setStatus(byId("styles-rtl-status"), "Applying…");
    Word.run(function (context) {
      var normal = context.document.getStyles().getByNameOrNullObject("Normal");
      var section = context.document.sections.getFirst();
      normal.load("isNullObject");
      return context.sync().then(function () {
        if (!normal.isNullObject) {
          normal.font.nameAscii = latinFont;
          normal.font.nameBidirectional = csFont;
          normal.font.sizeBidirectional = csSize;
        }
        section.pageSetup.sectionDirection = Word.SectionDirection.rightToLeft;
        return context.sync();
      });
    }).then(function () {
      setStatus(byId("styles-rtl-status"),
        "Applied: Latin font, complex-script font/size, and right-to-left section layout (margins, column order, footnote numbering direction).");
    }).catch(function (e) {
      setStatus(byId("styles-rtl-status"), "Error: " + (e.message || String(e)), true);
    });
  }
```

Bind: `byId("styles-rtl-apply").addEventListener("click", runRtlSetup);` inside the one-time binding block.

- [ ] **Step 2: Manual check (this is the task's real verification — mark it explicitly as a manual-only check, no automated equivalent exists)**

In Word desktop, on a fresh document:
1. Type a few Arabic/Urdu paragraphs in Normal style. Note current appearance.
2. Open the Styles tab → RTL document setup, set the complex-script font to a font actually installed (e.g. "Fatemi Maqala" if bundled, or "Scheherazade New"), size 14. Click "Set up RTL document".
3. Confirm the status text updates to "Applied: …".
4. Confirm the typed paragraphs now render in the chosen complex-script font at the chosen size (Normal style's font changed).
5. Add a footnote (References → Insert Footnote) with Arabic text. Confirm the footnote reference number and footnote pane text now follow right-to-left layout (this is the concrete effect of `sectionDirection = rightToLeft` from the design spec's checklist item 5 — if it does NOT visibly flip, this is a real spec-vs-Word-behavior gap to record as a follow-up, not a bug in this task's code).
6. Insert a brand-new blank paragraph and start typing Arabic. Note whether cursor/list-bullet behavior looks correct or backwards — if backwards, this is the **known, documented gap** (no Office.js bidi setter) surfacing exactly as expected; confirm the in-pane note text is visible and accurate, and tell the user to use Word's own Layout → Paragraph Direction control there.
7. Click "Set up RTL document" a second time — confirm no error and no visible double-application (idempotent).

- [ ] **Step 3: Commit**

```bash
git add src/taskpane/styles-pane.js
git commit -m "Add RTL document setup action (Latin/CS font+size, section direction)"
```

---

### Task 10: Full manual Word verification pass + spec cross-check

**Files:** none (verification only)

- [ ] **Step 1: Run the automated suite**

Run: `npm test`
Expected: all tests pass, including the new `word-styles.test.js`.

- [ ] **Step 2: Walk the design spec's §5 Testing checklist end-to-end in Word desktop**

Using `docs/superpowers/specs/2026-07-16-ashaar-styles-design.md` §5 as the checklist, in one Word session:
- Switch groups (General → Waaz → Petition → Maqala → back to General) and confirm all 6 roles reflow each time without touching unrelated document content.
- Save a custom group, close and reopen the document, confirm the custom group and its field values survive.
- Apply each of the 6 roles to a selection from the pane; separately, apply "Ashaar Heading 1" directly from Word's native Style gallery (Ctrl+Shift+S) and confirm it matches the pane's applied look (proves `unhideWhenUsed`/gallery visibility works and the two entry points agree).
- Confirm Quran Quote selected text shows the same left/right border as a plain Quote (proves the `basedOn` inheritance).
- Confirm Heading 1/2/3 each show up as their respective outline levels in the Navigation Pane (View → Navigation Pane) — proves the `basedOn "Heading N"` inheritance carries outline level through.
- Run the RTL setup action per Task 9 Step 2, including the footnote-numbering check.

- [ ] **Step 3: Record findings**

If any check in Step 2 fails or behaves differently than documented (especially the footnote-numbering claim or the blank-paragraph bidi gap), add a dated note to `docs/superpowers/specs/2026-07-16-ashaar-styles-design.md` under a new "## Post-implementation notes" section (create it if absent) rather than silently reinterpreting the spec — this keeps the spec as the accurate record for the next person who reads it.

- [ ] **Step 4: Bump the asset version**

Per the `bump-asset-version-on-deploy` project convention: edit `window.ASHAAR_ASSET_VERSION` in `src/taskpane/taskpane.html` to a new value (e.g. today's date + "-styles") so installed users don't run stale cached JS.

- [ ] **Step 5: Final commit**

```bash
git add docs/superpowers/specs/2026-07-16-ashaar-styles-design.md src/taskpane/taskpane.html
git commit -m "Record Ashaar Styles manual verification results; bump asset version"
```

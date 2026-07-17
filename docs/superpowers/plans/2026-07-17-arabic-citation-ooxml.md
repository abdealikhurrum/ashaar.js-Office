# Arabic Citation OOXML Rendering (font + italic) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax.

**Goal:** Render Arabic footnote/endnote/bibliography citations via `Range.insertOoxml()` so they use the document's complex-script font and drop italic on Arabic runs (fixing the Times-New-Roman font + tofu-squares defects). English + inline citations keep `insertHtml` unchanged.

**Architecture:** A pure HTML→OOXML run converter in `cite-word.js` turns sanitized citeproc HTML into `<w:r>` runs (Arabic runs get `<w:rtl/>` + `<w:rFonts w:cs="…">`, italic suppressed; Latin runs keep `<w:i/>`/`<w:b/>`), wrapped in a `<w:bidi/>` paragraph. `cite-pane.js` reads the document CS font, wraps the paragraph via the existing `AshaarTabStop.wrapOoxml`, and inserts via `insertOoxml` for the Arabic note/bib paths only.

**Tech Stack:** Vanilla ES5/UMD JS, Office.js `Range.insertOoxml` (FlatOPC), citeproc-js, node `assert`.

**Design spec:** `docs/superpowers/specs/2026-07-17-arabic-citation-ooxml-design.md`

## Global Constraints

- **No new runtime deps; no build step.** `cite-word.js` stays dependency-free (the pane supplies the FlatOPC wrap via the already-loaded `AshaarTabStop.wrapOoxml`).
- **Arabic-only OOXML path.** Only Arabic-locale **footnote/endnote/bibliography** insertion switches to `insertOoxml`. **Inline** citations (any locale) and **all en-US** citations keep `insertHtml` byte-for-byte unchanged.
- **Italic suppressed on Arabic runs only** — an Arabic run's `<w:rPr>` never contains `<w:i/>`; Latin runs keep `<w:i/>`/`<w:b/>` from `<i>/<em>`/`<b>/<strong>`.
- **Arabic runs carry `<w:rtl/>` + `<w:rFonts w:cs="<csFont>"/>`.** Neutral/digit characters adjacent to Arabic belong to the Arabic (rtl) run (same grouping as `wrapRtlRuns`), and ASCII `,`/`;` inside an Arabic run localize to `،`/`؛` (same as `wrapRtlRuns`).
- **Document CS font** = `Ashaar Normal` style's `font.nameBidirectional`, falling back to the target range's `nameBidirectional`, then a default constant `DEFAULT_AR_CS_FONT` (use `"Arial"` — universally present with Arabic coverage; do NOT invent a font name).
- **Paragraph is RTL:** `<w:p><w:pPr><w:bidi/><w:jc w:val="right"/></w:pPr>…</w:p>`.
- **SP-A tagging preserved:** the inserted range is still wrapped in the content control with `CiteWord.buildCitationTag(...)` / `buildBibliographyTag(...)`.
- **Tests in the `npm test` chain.** OUT of scope: Fatemi CSL style changes (SP-4), SP-B/SP-C.

---

### Task 1: `cite-word.js` — pure HTML→OOXML run converter

**Files:**
- Modify: `src/taskpane/cite-word.js` (add converter + paragraph builder; export them; reuse the existing `RTL_CHAR`/`LTR_CHAR`/`AR_PUNCT`)
- Test: `tests/cite-word.test.js` (append)

**Interfaces:**
- Produces:
  - `CiteWord.htmlToOoxmlRuns(html, {csFont})` → OOXML `<w:r>…` run string.
  - `CiteWord.buildCitationParagraphOoxml(html, {csFont})` → `"<w:p><w:pPr><w:bidi/><w:jc w:val=\"right\"/></w:pPr>" + runs + "</w:p>"`.
- Input `html` is the **sanitized** citeproc output (the converter does its own RTL grouping — do NOT pre-run `wrapRtlRuns`).

- [ ] **Step 1: Write the failing tests** — append to `tests/cite-word.test.js`:

```js
// --- htmlToOoxmlRuns (Arabic OOXML) ---
var AR = "كتاب"; // كتاب
var runs = CiteWord.htmlToOoxmlRuns("<i>" + AR + "</i>", { csFont: "Amiri" });
assert.ok(runs.indexOf("<w:rtl/>") !== -1, "Arabic run is rtl");
assert.ok(runs.indexOf('<w:rFonts w:cs="Amiri"/>') !== -1, "Arabic run uses the cs font");
assert.ok(runs.indexOf("<w:i/>") === -1, "italic is SUPPRESSED on the Arabic run (the squares fix)");
// Latin italic keeps <w:i/>, no rtl/cs
var lat = CiteWord.htmlToOoxmlRuns("<i>Daftary</i>", { csFont: "Amiri" });
assert.ok(lat.indexOf("<w:i/>") !== -1 && lat.indexOf("<w:rtl/>") === -1 && lat.indexOf("w:cs") === -1,
  "Latin italic run keeps <w:i/> and is not rtl/cs");
// mixed → distinct runs (both an rtl run and a latin run present)
var mixed = CiteWord.htmlToOoxmlRuns(AR + " Daftary", { csFont: "Amiri" });
assert.ok(mixed.indexOf("<w:rtl/>") !== -1 && /<w:r>(?!.*<w:rtl\/>).*Daftary/.test(mixed.replace(/\n/g,"")),
  "mixed content yields both an rtl run and a non-rtl Latin run");
// xml-escape
assert.ok(CiteWord.htmlToOoxmlRuns("A &amp; B", {}).indexOf("A &amp; B") !== -1, "ampersand stays escaped");
assert.ok(CiteWord.htmlToOoxmlRuns("a < b", {}).indexOf("&lt;") !== -1, "raw < is escaped");
// superscript
assert.ok(CiteWord.htmlToOoxmlRuns("<sup>1</sup>", {}).indexOf('<w:vertAlign w:val="superscript"/>') !== -1);
// paragraph wrapper
var para = CiteWord.buildCitationParagraphOoxml(AR, { csFont: "Amiri" });
assert.ok(para.indexOf("<w:p><w:pPr><w:bidi/><w:jc w:val=\"right\"/></w:pPr>") === 0, "RTL paragraph wrapper");
assert.ok(para.lastIndexOf("</w:p>") === para.length - "</w:p>".length, "paragraph closed");
console.log("htmlToOoxmlRuns test passed");
```

- [ ] **Step 2: Run to verify it fails** — `node tests/cite-word.test.js` → FAIL (`htmlToOoxmlRuns` undefined).

- [ ] **Step 3: Implement** — add to `src/taskpane/cite-word.js` (before the `return {…}`), reusing the existing `RTL_CHAR`, `LTR_CHAR`, `AR_PUNCT`:

```js
  function xmlEsc(s) {
    return String(s).replace(/&(?!amp;|lt;|gt;)/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  }

  // Tokenise into HTML tags (opaque) and single characters — same shape wrapRtlRuns uses.
  function tokenizeHtml(html) {
    var items = [], re = /<[^>]*>|[\s\S]/g, m;
    while ((m = re.exec(html || ""))) {
      var tok = m[0];
      if (tok.charAt(0) === "<" && tok.charAt(tok.length - 1) === ">") { items.push({ tag: tok }); }
      else { items.push({ ch: tok }); }
    }
    return items;
  }

  // Per-char RTL-run membership, mirroring wrapRtlRuns: an Arabic char plus the
  // trailing digits/neutrals/tags up to (but not including) the next Latin letter.
  function computeRtlFlags(items) {
    var n = items.length, flags = [], i;
    for (i = 0; i < n; i++) { flags[i] = false; }
    i = 0;
    while (i < n) {
      if (!items[i].tag && RTL_CHAR.test(items[i].ch)) {
        var start = i, lastRtl = i, k = i + 1;
        while (k < n) {
          var t = items[k];
          if (t.tag) { k++; continue; }
          if (LTR_CHAR.test(t.ch)) { break; }
          if (RTL_CHAR.test(t.ch)) { lastRtl = k; }
          k++;
        }
        var end = k;
        if (k < n) { while (end - 1 > lastRtl && !items[end - 1].tag && /\s/.test(items[end - 1].ch)) { end--; } }
        for (var j = start; j < end; j++) { if (!items[j].tag) { flags[j] = true; } }
        i = end;
      } else { i++; }
    }
    return flags;
  }

  function emitRun(text, sig, csFont) {
    var rpr = "<w:rPr>";
    if (sig.rtl) { rpr += "<w:rtl/>"; if (csFont) { rpr += '<w:rFonts w:cs="' + csFont + '"/>'; } }
    if (sig.b) { rpr += "<w:b/>"; if (sig.rtl) { rpr += "<w:bCs/>"; } }
    if (sig.i && !sig.rtl) { rpr += "<w:i/>"; } // italic suppressed on Arabic runs
    if (sig.sup) { rpr += '<w:vertAlign w:val="superscript"/>'; }
    rpr += "</w:rPr>";
    return "<w:r>" + rpr + '<w:t xml:space="preserve">' + xmlEsc(text) + "</w:t></w:r>";
  }

  function htmlToOoxmlRuns(html, opts) {
    var csFont = (opts && opts.csFont) || "";
    var items = tokenizeHtml(html);
    var rtl = computeRtlFlags(items);
    var out = [], buf = "", cur = null, fmt = { i: 0, b: 0, sup: 0 };
    function flush() { if (buf !== "" && cur) { out.push(emitRun(buf, cur, csFont)); } buf = ""; }
    for (var idx = 0; idx < items.length; idx++) {
      var it = items[idx];
      if (it.tag) {
        var tg = it.tag.toLowerCase().replace(/\s+/g, "");
        if (/^<(i|em)>$/.test(tg)) { flush(); fmt.i++; }
        else if (/^<\/(i|em)>$/.test(tg)) { flush(); fmt.i = Math.max(0, fmt.i - 1); }
        else if (/^<(b|strong)>$/.test(tg)) { flush(); fmt.b++; }
        else if (/^<\/(b|strong)>$/.test(tg)) { flush(); fmt.b = Math.max(0, fmt.b - 1); }
        else if (/^<sup>$/.test(tg)) { flush(); fmt.sup++; }
        else if (/^<\/sup>$/.test(tg)) { flush(); fmt.sup = Math.max(0, fmt.sup - 1); }
        else if (/^<br\/?>$/.test(tg)) { flush(); out.push("<w:r><w:br/></w:r>"); }
        // span/sub and any other tag: transparent (no format change)
        continue;
      }
      var sig = { rtl: rtl[idx], i: fmt.i > 0, b: fmt.b > 0, sup: fmt.sup > 0 };
      if (!cur || cur.rtl !== sig.rtl || cur.i !== sig.i || cur.b !== sig.b || cur.sup !== sig.sup) {
        flush(); cur = sig;
      }
      var ch = it.ch;
      if (sig.rtl && AR_PUNCT[ch]) { ch = AR_PUNCT[ch]; }
      buf += ch;
    }
    flush();
    return out.join("");
  }

  function buildCitationParagraphOoxml(html, opts) {
    return '<w:p><w:pPr><w:bidi/><w:jc w:val="right"/></w:pPr>' + htmlToOoxmlRuns(html, opts) + "</w:p>";
  }
```

Add `htmlToOoxmlRuns` and `buildCitationParagraphOoxml` to the returned object (keep all existing exports).

- [ ] **Step 4: Run to verify it passes** — `node tests/cite-word.test.js` → PASS.

- [ ] **Step 5: Full suite** — `npm test` → all pass.

- [ ] **Step 6: Commit**

```bash
git add src/taskpane/cite-word.js tests/cite-word.test.js
git commit -m "feat(cite): HTML->OOXML run converter (Arabic cs font, italic-suppressed, rtl)"
```
(End with the two session trailer lines.)

---

### Task 2: `cite-pane.js` — Arabic notes/bibliography via `insertOoxml`

**Files:**
- Modify: `src/taskpane/cite-pane.js`
- Modify: `src/taskpane/taskpane.html` (bump `ASHAAR_ASSET_VERSION` to `"20260717-cite-ar-ooxml"`)
- No new node test (Office/DOM wiring; converter is covered by Task 1). Verify via `npm test` regression + the manual Word checklist.

**Interfaces (consumed):**
- `CiteWord.buildCitationParagraphOoxml(html, {csFont})` (Task 1), `CiteWord.buildCitationTag`/`buildBibliographyTag` (SP-A), `CiteWord.sanitize`, and `AshaarTabStop.wrapOoxml(bodyXml)` (already global — wraps a `<w:p>` in a FlatOPC package for `insertOoxml`).

- [ ] **Step 1: Add a CS-font reader + default constant.** In `cite-pane.js`, add near the top:

```js
  var DEFAULT_AR_CS_FONT = "Arial"; // universally present, has Arabic coverage
  // Read the document's complex-script font inside a Word.run: prefer the
  // "Ashaar Normal" style (created by the Styles-tab RTL setup), else the target
  // range's bidi font, else the default. Returns a Promise<string>.
  function readDocCsFont(ctx, range) {
    var style = ctx.document.getStyles().getByNameOrNullObject("Ashaar Normal");
    style.load("isNullObject,font/nameBidirectional");
    range.font.load("nameBidirectional");
    return ctx.sync().then(function () {
      if (!style.isNullObject && style.font.nameBidirectional) { return style.font.nameBidirectional; }
      if (range.font.nameBidirectional) { return range.font.nameBidirectional; }
      return DEFAULT_AR_CS_FONT;
    });
  }
```

- [ ] **Step 2: Route the Arabic footnote/endnote insert through OOXML.** In `insertCitation`, the note branch currently does `note.body.getRange().insertHtml(noteHtml, replace)`. For Arabic notes, insert OOXML instead. Restructure the `Word.run` note branch so that when `rtl && form !== "inline"`:

```js
        } else if (canNotes) {
          var note = form === "endnote" ? sel.insertEndnote() : sel.insertFootnote();
          var noteRange = note.body.getRange();
          if (rtl) {
            return readDocCsFont(ctx, noteRange).then(function (csFont) {
              var sanitized = CiteWord.sanitize(engine.cite(items));
              var pkg = AshaarTabStop.wrapOoxml(CiteWord.buildCitationParagraphOoxml(sanitized, { csFont: csFont }));
              var range = noteRange.insertOoxml(pkg, Word.InsertLocation.replace);
              var cc = range.insertContentControl();
              cc.tag = citeTag; cc.title = "Ashaar Citation";
              return ctx.sync();
            });
          }
          range = noteRange.insertHtml(noteHtml, Word.InsertLocation.replace);
        }
```

Keep the inline branch and the Word<1.5 fallback branch on `insertHtml` unchanged. Keep the en-US note branch (`rtl` false) on `insertHtml`. Ensure the non-OOXML branches still reach the shared `cc = range.insertContentControl()` + alignment code; the OOXML branch does its own CC + returns early (it already set bidi, so `rightAlignParas` is unnecessary there). Make sure the promise chain returns correctly in all branches.

- [ ] **Step 3: Route the Arabic bibliography through OOXML.** In `insertBibliography`, when `rtl`, build via OOXML the same way (read CS font from the selection range, `buildCitationParagraphOoxml(CiteWord.sanitize(engine.bibliography()), {csFont})`, `wrapOoxml`, `getSelection().getRange().insertOoxml(pkg, after)`, tag with `buildBibliographyTag`). en-US keeps `insertHtml`.

- [ ] **Step 4: Asset version.** In `taskpane.html`, bump `window.ASHAAR_ASSET_VERSION` to `"20260717-cite-ar-ooxml"`.

- [ ] **Step 5: Regression.** Run `npm test` → all pass (no node test covers the pane; this is the regression gate).

- [ ] **Step 6: FEASIBILITY VERIFY (live, do this before relying on the path).** Start the dev server, sideload/reload in Word (`npm start` or reload the pane). With Arabic locale, insert a footnote citation. Confirm: (a) `insertOoxml` into the footnote body succeeds (no error in the pane status), (b) the Arabic text renders in the document's Arabic font (not Times New Roman), (c) the title is NOT italic / shows no squares, (d) the paragraph is RTL, (e) the content control tag is still present. If `insertOoxml` into a footnote body throws, STOP and report — the fallback is to insert into the note body differently (e.g. clear + insert), but do not silently leave it broken. Also confirm an en-US footnote is unchanged.

- [ ] **Step 7: Commit**

```bash
git add src/taskpane/cite-pane.js src/taskpane/taskpane.html
git commit -m "feat(cite): insert Arabic notes/bibliography as OOXML (doc cs font, no italic squares)"
```

---

## Manual checklist (needs Word + the Arabic font)
1. Arabic footnote → renders in the document's Arabic font, title correct + non-italic (no squares), RTL paragraph, number on the right (with RTL doc setup).
2. Arabic bibliography → same.
3. en-US footnote + bibliography → unchanged (still Times/whatever, italics preserved for Latin titles).
4. Inline citation (Arabic + en) → unchanged (still `insertHtml`, in-flow).
5. The inserted Arabic citation's content control still decodes to the SP-A `{style,locale,keys,locators}` tag.

## Self-review (author)
- Spec coverage: converter font/italic/rtl (T1), Arabic note+bib OOXML insert + CS-font read (T2), inline/en-US unchanged (T2 constraints), tagging preserved (T2). Mapped.
- Type consistency: `htmlToOoxmlRuns(html,{csFont})` / `buildCitationParagraphOoxml(html,{csFont})` / `AshaarTabStop.wrapOoxml(paragraphXml)` consistent across tasks.
- Risk: footnote-body `insertOoxml` is verified live in T2 Step 6 before the path is trusted (mirrors SP-2's translator-name verification).

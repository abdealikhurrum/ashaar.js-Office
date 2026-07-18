# Citation Engine — Manual Word Verification Checklist

SP-1 Citation Engine, Task 7 (Cite tab UI + Word insertion). Subagents / CI cannot
open real Word, so the Word.run insertion paths (footnote, endnote, inline, tagged
bibliography, RTL-as-alignment, Word < 1.5 fallback) must be verified by hand.

## What is already automated (do NOT re-verify by hand)

- Node suite (`npm test`) — `citeproc-vendor`, `cite-engine` (en / ar / multilingual /
  Fatemi parity), `cite-word`.
- Browser smoke test (Playwright, headless, no Word) verified on 2026-07-16:
  - Zero JS console errors on load (citeproc + cite-engine + cite-word + cite-pane all parse).
  - Cite tab activates, item list populates from `fixtures/cite-sample.json` (3 items).
  - Live preview renders formatted citation + bibliography (pure JS, no Word):
    - Chicago footnote: `Farhad Daftary, *The Fatimid Empire* (Edinburgh University Press, 2018).`
    - APA author-date: `(Daftary, 2018; القاضي النعمان, 1951)`.
    - Arabic locale flips the preview container to `dir="rtl"`.

Everything below needs a real Word host (desktop Word recommended: Windows/Mac).

## Setup

1. `npm start` (loads the add-in into desktop Word with the dev manifest) — or sideload
   the manifest and open the task pane.
2. Confirm the pane header reads **Connected to Word**.
3. Click the **Cite** tab. Confirm: the item list shows three items
   (*The Fatimid Empire* (en-book), *Isma'ili History and Doctrine* (en-article),
   دعائم الإسلام (ar-book)) and the preview renders a formatted citation.

## Checks

### A. Chicago footnote (English) — `en-book`
1. Style = **Chicago (notes & bibliography)**, Locale = **English (en-US)**, Output form = **Footnote**.
2. Check only **The Fatimid Empire (en-book)**. Preview shows
   `Farhad Daftary, The Fatimid Empire (Edinburgh University Press, 2018).` with the
   title in italics.
3. Place the cursor in the document body. Click **Insert citation**.
4. Expected: a real Word footnote is created at the cursor; the footnote text reads
   `Farhad Daftary, The Fatimid Empire (Edinburgh University Press, 2018).` with
   *The Fatimid Empire* in italics. Status line reads "Inserted footnote citation."

### B. Endnote (English)
1. Same as A but Output form = **Endnote**. Click **Insert citation**.
2. Expected: a real Word **endnote** (not footnote) is created; same formatted text;
   status "Inserted endnote citation."

### C. Inline (English)
1. Output form = **Inline**. Put the cursor in a paragraph. Click **Insert citation**.
2. Expected: the citation text is inserted inline at the cursor (no note reference),
   italics preserved; status "Inserted inline citation."

### D. Arabic footnote + RTL treatment — `ar-book`
1. Locale = **Arabic (ar)**, Style = **Chicago (notes & bibliography)**, Output form = **Footnote**.
2. Check **دعائم الإسلام (ar-book)** (optionally also en-book to see a mixed cluster).
3. Preview container is right-aligned (`dir="rtl"`) and shows the Arabic title.
4. Click **Insert citation**. Expected: a footnote is created; the Arabic title renders;
   the parentheses, comma and period around the Arabic run sit in the **correct
   positions** — e.g. `(دار المعارف، 1951).`, NOT `)دار المعارف, ).1951` — and the
   in-run commas are **Arabic commas ،** (verified live in Word, 2026-07-16).
5. **How the bidi correctness is achieved / its boundary.** citeproc emits plain
   mixed-direction text with no directional markup, so each maximal Arabic run is
   wrapped in `<span dir="rtl">` (mapped by Word to an RTL run) — this fixes the
   neutral-punctuation ordering *within* the run, and ASCII comma/semicolon inside the
   run are localized to `،`/`؛` (CSL locales cannot localize inter-element delimiters).
   Separately, Office.js exposes **no paragraph reading-order setter** — only
   `paragraph.alignment` — so the whole footnote paragraph is **right-aligned** rather
   than set to true RTL paragraph direction. Per-run direction (the span) is what makes
   the punctuation correct; the paragraph-level right-to-left flag, if wanted, is still a
   manual step: **Layout → Paragraph Direction → Right-to-Left** (desktop Word).

### E. APA (author-date form)
1. Style = **APA (author-date)**, Locale = English, Output form = **Inline** (author-date
   styles are in-text, not note styles).
2. Check en-book. Preview shows `(Daftary, 2018)`. Click **Insert citation**.
3. Expected: `(Daftary, 2018)` inserted inline. (Note: the Fatemi variants —
   *Chicago — Fatemi*, *APA — Fatemi* — are wired but behave identically to their base
   styles until SP-4 activates Fatemi honorific handling.)

### F. Bibliography (tagged content control)
1. Any style/locale. Place cursor where the bibliography should go. Click **Insert bibliography**.
2. Expected: a bibliography block is inserted (all three sample entries, formatted for the
   chosen style). Status reads `Inserted bibliography (tagged "AshaarBibliography").`
3. Verify the block is wrapped in a **content control** whose **Tag = `AshaarBibliography`**
   and **Title = "Ashaar Bibliography"** (Developer tab → click the block → Properties, or
   hover the CC boundary). Under Arabic locale, the bibliography paragraphs are right-aligned
   (same RTL-as-alignment limitation as D).

### G. Word < 1.5 fallback (notes unavailable)
1. On a host **without WordApi 1.5** (older Word build / Word on the web where notes aren't
   supported): Output form = **Footnote** or **Endnote**, click **Insert citation**.
2. Expected: no error thrown; the citation is inserted **inline** as a graceful fallback,
   and the status line reads
   "Word < 1.5: footnotes/endnotes unavailable — inserted inline instead."
   (Feature detection: `Office.context.requirements.isSetSupported("WordApi","1.5")`.)

## Notes / caveats

- All inserted HTML is passed through `CiteWord.sanitize` first — only
  `i b em strong span sup sub br` survive, with all attributes stripped — so no CSL
  markup can inject styles/scripts into the document. `CiteWord.wrapRtlRuns` then runs
  *after* sanitize and adds our own trusted `<span dir="rtl">` around Arabic runs, so
  the only attribute in the final HTML is that controlled `dir`.
- RTL correctness = per-run `dir="rtl"` spans (fixes punctuation ordering) + Arabic
  in-run comma/semicolon + paragraph right-alignment; true paragraph-level RTL flag is
  the one manual step (see D5); document this to end users.
- The bibliography content-control tag (`AshaarBibliography`) is the anchor a future
  "refresh bibliography" / re-cite feature would key on.

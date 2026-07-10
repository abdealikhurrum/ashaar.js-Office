# Guided Justification UX — Design

**Date:** 2026-07-10
**Status:** Approved, ready for implementation plans
**Scope:** Make the task-pane justification workflow self-explanatory for **non-technical, regular** users, using **just-in-time discovery** — surface what's needed, when it's needed, in stable places (no live-churning "magical staircases"). Turns today's opaque, "thoughts-and-prayers" Justify into a transparent, explainable, recoverable action. Includes the engine/behavior fixes required to make the guidance honest.

## Guiding principles

1. **Discovery when needed, not upfront.** Guidance appears at the moment of the action it explains, then stays put until the next action. Nothing reflows as the cursor moves.
2. **Explain *why*, not just *how*.** For every unintuitive step (loading a font, choosing a mode), the user must understand the reason before the mechanics.
3. **Every shortfall has a recourse.** When justification underfills, name the limiting factor and offer the exact control to change.
4. **Honest controls.** A control that appears to do something must actually do it (today's "Word justify" and "Stretch strength" don't, on the Justify path — see Findings).

## Findings that reshape scope (verified in code)

- **"Word justify" never reaches the document.** `justifyMode:"css"` only adds a CSS class to the live preview (`taskpane.js:291`); inserted Word paragraphs are aligned right/left/center by column (`word-html.js:1090`), never `w:jc="both"/"distribute"`. In the document it behaves like "None."
- **The Stretch-strength slider does nothing on Justify.** `justifySelection` builds `targetFill:0.92` then overwrites it with an auto-calibrated value (`taskpane.js:1234–1251`); the slider (`tatweelCount`) is read only on insert-as-table and qaseeda paths. Dragging it and re-justifying changes nothing.
- **Fill is capped at the column edge by design.** All paths cap `targetFill ≤ 1.0`; calibrate searches `[0.80,1.00]` (`ashaar-autotune.js:386`); default `0.96` with a comment that breathing room *prevents* over-stretch. There is no expressive/override regime.
- **No programmatic undo.** Office.js has no undo/redo hook into Word's history. Native ⌘Z/Ctrl+Z is the only exact undo. A deterministic "reset to bare line" is possible because justification is reducible (`stripJustification` + re-justify).

## Audience

Non-technical, regular users (a small team running a publication workflow). They learn the basics but keep hitting the hard spots — selecting a target, loading fonts, and understanding why a line didn't fill. Favors quiet contextual cues that teach without nagging once learned.

---

## Components

### 1. Justification Result panel

A stable panel below the **Justify Selected Text** button. Populated only *after* an action; persists until the next action; never live-updates on cursor movement.

**Layout: "Summary + only what fell short," with a Show-all toggle.**
- **Headline:** `✓ Justified 6 of 8 lines · 2 could fill more`.
- **Shortfall list (default visible):** only lines that underfilled, each with: the misra text (RTL), a fill meter + %, a plain-language **reason**, and a **fix button**.
- **`Show all N lines ▾`** expands to the full per-line list (every line with its fill % and ✓/flag) — the "Option B" density on demand.

**Recourse — the four limiting reasons, each with a one-click fix:**
| Reason | Message | Fix action |
|---|---|---|
| Capped by stretch strength | "78% — capped by stretch strength" | Focus/raise the Stretch-strength slider (see §4). |
| Font not resolvable | "71% — font "X" not loaded (measuring is a guess)" | Open the Add-font flow (§2) with X's name prefilled. |
| No stretch joins in this line | "Kashida can't stretch this line" | Suggest **Space out the words** or **Let Word fill it** (§5). |
| Hit column / page width | "Already as wide as the column allows" | Turn on **Auto-fit to text** / widen the table. |

**Empty state (covers the "what do I select?" confusion, stably):** if Justify runs with no valid target (not inside a poem table/content control), the same panel shows `Nothing to justify — click inside a poem table, then Justify again.` No live pre-action button state (rejected as a "magical staircase").

**Undo affordance (see §6):** after applying, a quiet line: `Not what you wanted? Press ⌘Z to undo` (OS-aware) plus a **Reset to unstretched** button.

**Data source:** the metrics already computed in the debug path (`nat/target/fin/fill/tatweel`, `taskpane.js` debug block) become the panel's data, translated to plain language. The existing raw debug table remains available behind the pane's debug checkbox as an "advanced" view.

### 2. Font-loading flow (the "Add font" destination)

Triggered from the result panel's "font not loaded" fix, or opened directly. Reworked to answer **why**, then **why here/again**, then help **find the file**.

**Why (stated + shown):** kashida fills a line by stretching letters; to know how far, the engine must **measure how wide your letters are**. A small before/after shows a line falling short ("guessing") vs. filling the column ("measured").

**Why *here*, why *again* (the two-sealed-rooms model):** the pane is a separate, sandboxed mini-app running beside Word; for security it **cannot read the fonts on your computer or the one Word is using**. A compact diagram: `Word [has font ✓]` — 🧱 SEALED — `This pane [can't see it]`. Handing it the file once is the only way in.

**The steps:**
1. **Detected at your cursor** — auto-read the font at the selection and **prefill** the name (editable; "Re-detect" fallback). Removes the "name Word uses" guesswork.
2. **Give it the file** — a **drag-and-drop zone** (primary) + "Choose file" (fallback). An OS-aware **"Where do I find it?"** expander:
   - **macOS:** Font Book → find the font → right-click → *Show in Finder* → drag it in. Copyable `~/Library/Fonts`.
   - **Windows:** `Win+R` → `%WINDIR%\Fonts` → **copy it to Desktop first** (Windows won't let you select straight from the Fonts folder) → Choose the copy. Copyable `C:\Windows\Fonts`.
   - **Caveats laid out** in the expander:
     - The **file name often differs from the display name** (e.g. `ScheherazadeNew-Regular.ttf`).
     - Some **system fonts are protected/bundled** (macOS `.ttc` collections, certain Windows system fonts) and **can't be copied**.
     - **In either caveat case → recommend "Let Word fill it"** (§5): it needs no font loaded, so it sidesteps the whole problem.
3. **Add & re-justify** — confirmation: `✓ Loaded — the engine now measures its real letters.`
   Done once per font; stored on the machine; nothing uploaded.

**Result-panel inline reason:** at the "font not loaded" flag, a one-line version — "the pane can't see Word's fonts — give it the file once" — linking to this full explainer.

### 3. Fix "Let Word fill it" to actually justify (native Word kashida)

Today `justifyMode:"css"` never reaches the document. Rework the mode to emit real Word justification on the inserted/justified paragraphs (table path in `word-html.js`; tab-stop path where applicable), using Word's own font — no pane measurement, no sandbox wall (the §2 fallback).

**Word's justification values** (verified, OOXML `w:jc` / `JustificationValues`):
- `lowKashida` / `mediumKashida` / `highKashida` — **true kashida**: elongate the connecting strokes (letters stay joined), at three intensities. Arabic/complex-script only.
- `distribute` — **spacing, not kashida**: "distribute all characters equally" (inter-character + inter-word spacing). Script-agnostic; fills every line including the last.

**Mapping:**
- **Arabic:** emit a **kashida level** (mapped from strength — see §4) for genuine stroke elongation.
- **Last-line problem + fix:** kashida (like `both`) does **not** stretch the last line of a paragraph, and each misra cell is a single line = the last line. So the app **auto-appends a trailing soft line break (`<w:br/>`)** to the misra paragraph, demoting it from last-line so Word kashidas it — the user presses nothing. The break run is shrunk (e.g. ~2pt) / line spacing tightened to minimize the trailing empty line's height.
- **Non-Arabic fallback:** `distribute` (spacing) — the only value that fills a single last line without the break, for text kashida can't elongate. Not a kashida substitute for Arabic (it's loose spacing).

**Verify-on-device risk (flag):** whether the trailing `<w:br/>` reliably triggers kashida on line 1, the empty-line height after shrinking, and kashida behavior across Word builds — all require testing in Word. This is the second place a last-line/trailing-break subtlety bites; treat as the riskiest task.

### 4. Stretch strength — honest, mode-specific, expressive

- **Wire the slider into the Justify path.** Today the slider does nothing on Justify (auto-calibrated `targetFill`). Make the Stretch-strength value drive fill on **Justify Selected Text**, so the control — and the result panel's "raise strength" recourse — actually do something.
- **The fill lever is mode-specific** (the slider drives whatever the chosen mode uses):
  - **Ashaar.js engine (tatweel):** more **tatweels** at legal joins. Never word spacing.
  - **Space out the words:** more **spaces** (that mode's mechanism — spacing legitimately grows here).
  - **Let Word fill it:** the strength picks the Word **kashida level** (`low → medium → high`, §3).
- **Ashaar.js-engine two regimes (0–24 slider):**
  - **0–15:** today's fill behavior, now slider-driven (fills toward the column edge).
  - **15–24 (expressive):** raise the **tatweel cap** from **1× → 3×** the engine's normal limit along an **exponential** curve (eases in just past 15, then accelerates to 3× at 24). Extra tatweels distribute across **legal joins only** (illegal joins always skipped). Tick shown on the slider at 15; the slider reads its multiplier (e.g. "24 · 3×").
  - **Bounded by the cell edge — no auto-resize.** The tatweel cap is usually the first barrier (not column width), so lifting it fills lines that fell short *within the existing cell*. The table is **never** auto-widened. If a line still can't fit, that is the separate, user-initiated **Auto-fit** recourse (§1 "hit column/page width").
- Not word spacing anywhere in engine mode; not a column resize.

### 5. Mode chooser — plain language, Ashaar.js engine as the hero

Reframe the "Justification" control (header: **"How should lines fill the column?"**). Four modes; the smart path is branded and clearly recommended, the rest are plainer fallbacks. Lead with **what each does** — "kashida" is jargon almost no one knows, so it appears only as a learnable parenthetical.

| Label | Framing |
|---|---|
| **Ashaar.js engine** — *stretches the letters (kashida)* · **recommended** | The engine reads your actual font and elongates letters only at the joins a scribe would use, never breaking a ligature. Chips: *reads your font · best-join placement · fine strength control*. Owns the Stretch-strength slider (§4). |
| **Space out the words** | Adds gaps between words instead of stretching letters. Use when a font's letters can't stretch. |
| **Let Word fill it** | Word's own kashida — real letter elongation using **Word's installed font**, so it **needs no font loaded** (the fallback of §2). Strength picks the intensity (low/med/high); the app auto-inserts the trailing break so single misra lines actually stretch (§3). |
| **Leave as typed** | No filling. |

The Ashaar.js-engine framing is also what *motivates* loading a font (§2): the engine's font-measuring intelligence is the reason the file matters.

### 6. Undo / reset

- **Undo hint** in the result panel after applying: OS-aware `Press ⌘Z / Ctrl+Z to undo` — native, restores the exact prior state.
- **Reset to unstretched** button: a deterministic clean slate that undoes **every** fill artifact across all modes, not just tatweels. It must reverse:
  - **tatweels** (engine mode) — strip (`stripJustification`);
  - **micro-spaces** (spacing mode) — strip;
  - **uniform font scale** (spacing/scale) — reset run sizes to their originals;
  - **Word-fill artifacts (§3)** — remove the auto-inserted trailing `<w:br/>` **and** reset the paragraph's `w:jc` from the kashida/`distribute` value back to its natural alignment (right/left/center by column position).
  So `stripJustification` alone is insufficient for Word-fill mode; Reset is a mode-complete cleanup. (Native ⌘Z is unaffected — Word's own history reverses all of this regardless.)

---

## Suggested build order (each its own implementation plan)

Ordered so honesty-fixes land before the guidance that points at them:

1. **§3 Fix "Let Word fill it" → native Word kashida + auto trailing break** — unblocks the font fallback and mode framing. OOXML assertion for the `w:jc` value + trailing break; **but the real behavior (kashida on the single line, empty-line height) must be verified in Word** — riskiest task, do it first so surprises surface early.
2. **§4 Stretch strength wired + mode-specific fill + expressive tatweel cap (1×→3× exponential, cell-bounded)** — engine/behavior; makes the slider and the "raise strength" recourse real.
3. **§1 Justification Result panel** (+ empty state, recourse wiring, undo hint) — the centerpiece; depends on 3 & 4 for accurate recourse.
4. **§2 Font-loading flow** (why, detect/prefill, dropzone, OS locate, caveats, fallback) — links from §1's font recourse.
5. **§5 Mode chooser rename/reframe** + **§6 Reset action** — presentation + the reset button.

## Out of scope

- Pre-action live selection status on the button (rejected — churn).
- A true multi-step in-pane undo (Office.js can't; ⌘Z + Reset cover it).
- Redesigning Table Input / Conversion / qaseeda panels beyond what these components touch.
- The qaseeda profile path's own strength wiring (already uses `strengthToTargetFill`); revisit only if it conflicts with §4.

## Resolved during design (was open)

- **§4 expressive mapping:** tick at **15**; 15→24 raises the tatweel cap **1×→3× on an exponential curve**; **cell-bounded, no auto-resize**; the fill lever is **mode-specific** (tatweels / spaces / Word kashida level).
- **§3 Word fill:** native **kashida levels** + **auto trailing `<w:br/>`** (not plain `distribute`, which is spacing); `distribute` only as a non-Arabic fallback.
- **§6 Reset** must also remove the trailing break and reset `w:jc`.

## Open questions for spec review

- Whether the result panel should also appear after **Insert as Table**, or only after **Justify**.
- The exact strength→kashida-level thresholds in Word-fill mode (where low/medium/high switch on 0–24).

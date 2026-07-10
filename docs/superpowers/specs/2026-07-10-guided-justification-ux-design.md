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

### 3. Fix "Let Word fill it" to actually justify (→ `w:jc="distribute"`)

Make the mode emit `w:jc="distribute"` on the inserted/justified Word paragraphs (both the table path in `word-html.js` and, where applicable, the tab-stop path). `distribute` justifies **every** line including the last, so a single-line misra fills **without a trailing Shift-Enter**. This makes the §2 fallback real: Word justifies using its own installed font — no pane measurement, no sandbox wall.

### 4. Stretch strength — honest + expressive

- **Wire the slider into the Justify path.** The Stretch-strength value drives fill on **Justify Selected Text** (seeds/overrides the calibrated `targetFill`), so the control — and the result panel's "raise strength" recourse — actually do something.
- **Two regimes:**
  - **Low→mid:** fill `~90%→100%` of the current column (today's behavior, now slider-driven).
  - **High ("expressive") zone:** past 100%, **override the cap** — auto-widen the column (engage Auto-fit) and/or continue elongating beyond the current column, so letters visibly extend for dramatic effect that responds to the slider's magnitude. The regime boundary is shown on the control so the jump from "fill" to "stretch" is legible.
- Interacts with the "hit column/page width" recourse (§1): raising strength into the expressive zone is what widens past the column ceiling.

### 5. Mode chooser — plain language, Ashaar.js engine as the hero

Reframe the "Justification" control (header: **"How should lines fill the column?"**). Four modes; the smart path is branded and clearly recommended, the rest are plainer fallbacks. Lead with **what each does** — "kashida" is jargon almost no one knows, so it appears only as a learnable parenthetical.

| Label | Framing |
|---|---|
| **Ashaar.js engine** — *stretches the letters (kashida)* · **recommended** | The engine reads your actual font and elongates letters only at the joins a scribe would use, never breaking a ligature. Chips: *reads your font · best-join placement · fine strength control*. Owns the Stretch-strength slider (§4). |
| **Space out the words** | Adds gaps between words instead of stretching letters. Use when a font's letters can't stretch. |
| **Let Word fill it** | Word's own justify — fills every line, no Shift-Enter, **needs no font loaded** (the fallback of §2/§3). |
| **Leave as typed** | No filling. |

The Ashaar.js-engine framing is also what *motivates* loading a font (§2): the engine's font-measuring intelligence is the reason the file matters.

### 6. Undo / reset

- **Undo hint** in the result panel after applying: OS-aware `Press ⌘Z / Ctrl+Z to undo` — native, restores the exact prior state.
- **Reset to unstretched** button: strips the poem's tatweels/micro-spaces and resets any scaled sizes back to the bare lines via the reducible engine. A deterministic clean slate, clearly labeled as a reset (not a step-undo).

---

## Suggested build order (each its own implementation plan)

Ordered so honesty-fixes land before the guidance that points at them:

1. **§3 Fix "Let Word fill it" → distribute** — small, unblocks the font fallback and mode framing. Testable via OOXML assertion.
2. **§4 Stretch strength wired + expressive regime** — engine/behavior; makes the slider and the "raise strength" recourse real.
3. **§1 Justification Result panel** (+ empty state, recourse wiring, undo hint) — the centerpiece; depends on 3 & 4 for accurate recourse.
4. **§2 Font-loading flow** (why, detect/prefill, dropzone, OS locate, caveats, fallback) — links from §1's font recourse.
5. **§5 Mode chooser rename/reframe** + **§6 Reset action** — presentation + the reset button.

## Out of scope

- Pre-action live selection status on the button (rejected — churn).
- A true multi-step in-pane undo (Office.js can't; ⌘Z + Reset cover it).
- Redesigning Table Input / Conversion / qaseeda panels beyond what these components touch.
- The qaseeda profile path's own strength wiring (already uses `strengthToTargetFill`); revisit only if it conflicts with §4.

## Open questions for spec review

- Exact expressive-zone mapping in §4 (where the "widen/over-stretch" boundary sits on 0–24, and whether it auto-toggles Auto-fit vs. exceeds column directly).
- Whether the result panel should also appear after **insert** (Insert as Table), or only after **Justify**.

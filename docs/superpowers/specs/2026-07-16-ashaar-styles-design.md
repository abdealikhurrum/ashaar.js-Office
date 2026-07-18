# Ashaar Styles — Design Spec

**Date:** 2026-07-16
**Status:** Draft — user-approved (brainstorming); ready for implementation planning.
**Scope:** Prose/document styling (headings, emphasis, block quotes, Quran quotes) and one-time RTL document setup, exposed as a new "Styles" tab in the task pane. This is distinct from the existing poetry table/justification system, which is untouched by this spec.

**Explicitly out of scope (separate specs):**
- **Annotation/citation system** for bayan excerpts (notes on how a verbatim quote is cited, its relationship to the document, "potential new joins"). Deferred because it's a metadata/tagging subsystem, not a visual style, and deserves its own design once styles exist to annotate against. When it's built, marking a verbatim bayan excerpt should use a content-control/tag (consistent with how poem blocks are already tagged), not a Word style with no visual signature.
- **Double-press input conversion, symbol-to-letter conversion (e.g. Arabic semicolon ↔ cheh, brackets/mathematical symbols), and honorific-abbreviation expansion (maza lemaza, bhaisaheb, behnsaheb, etc.) via find-replace.** A distinct, unrelated subsystem (text conversion / find-replace tooling), to be brainstormed as its own spec immediately after this one.
- **Custom keyboard shortcuts** for invoking styles. Feasible (Office.js `KeyboardShortcuts 1.1` requirement set) but requires converting the add-in to a shared runtime (`SharedRuntime 1.1`) — a real architectural change — and only works on Word on the web, Word on Windows 2408+ (build 17928.20114+), or Word on Mac 16.88+ (build 24081116+). Treated as a fast-follow once the Styles tab's buttons work, not a blocker for this spec.

## Overview

Users writing prose documents (bayans, sermons, petitions, essays) in Word need consistent, reusable formatting for structural elements — headings, emphasized phrases, block quotations, and Quran citations (which sometimes carry vocalization/harakat) — plus a one-time setup pass that makes the document behave correctly as an RTL (Arabic/Urdu/Persian) document. Different genres of document (a petition vs. a sermon transcript vs. a general essay) conventionally look different, so the same abstract roles (Heading 1, Emphasis, Quote, …) need interchangeable concrete recipes per use case.

## §1 — Architecture

A new module `src/taskpane/word-styles.js` (UMD, same pattern as `word-html.js`/`word-tabstop.js`) provides:
- `ensureAshaarStyles(context, groupName)` — idempotently creates/updates the 6 named Word styles in the document, using the given style group's recipe.
- `ensureRtlDocumentSetup(context, opts)` — idempotently applies the RTL document checklist (§4).
- `applyHeading(selection, level)`, `applyEmphasis(selection)`, `applyQuote(selection)`, `applyQuranQuote(selection)` — apply a style to the current selection.

Pure logic (style-definition builders, OOXML fragments, group-recipe merging) is node-testable, same as the existing suite. The `Word.run()` orchestration lives alongside `taskpane.js`'s existing call sites into `word-html.js`.

A new **"Styles" tab** in the task pane sits alongside Table Input / Conversion.

## §2 — The 6 style roles

Each role is backed by a **named Word style**, `basedOn` the closest built-in (for outline-level/TOC/navigation-pane inheritance and Word-native discoverability), and marked as a Quick Style (gallery-visible) so it appears directly in Word's own Style gallery/Styles pane (Ctrl+Shift+S) — a user can apply it from Word's native UI without opening the task pane.

| Role | Based on | Style-level fields (edit → affects every use) | Instance-level fields (this selection only) |
|---|---|---|---|
| **Ashaar Heading 1** | Heading 1 | Font, size (bold + centered fixed, not exposed) | — |
| **Ashaar Heading 2** | Heading 2 | Font, size | — |
| **Ashaar Heading 3** | Heading 3 | Font, size | — |
| **Ashaar Emphasis** (character style) | Emphasis | Color (default red), point-bump amount (default +N pt, adjustable) | None — the *resulting* absolute size is computed live at apply time as "selection's current run size + the style's point-bump," since emphasized phrases sit inside body text of varying size and a named Word style can't express a relative delta itself |
| **Ashaar Quote** | Quote | Border left+right, default width | Width/indent — quotes vary in length/context, so width is direct paragraph formatting layered on top of the style |
| **Ashaar Quran Quote** | Ashaar Quote | Font (recommended Quran fonts: Amiri Quran, KFQC, or other installed), default line height | Line height — vocalized excerpts need more vertical room than unvocalized ones, so it isn't forced document-wide |

Ashaar Quran Quote being based on Ashaar Quote (not directly on Word's Quote) means clicking **Quran Quote** already yields a bordered block — Quran quotes in this domain are always block quotes, so there's no separate "apply Block Quote, then apply Quran Quote" composition step.

Editing a style-level field writes directly to the named Word style (Word propagates it to every paragraph using that style automatically — no separate "Update" step, unlike poem profiles, since exactly one instance of each role exists per active group).

## §3 — Style groups

A **style group** is a complete recipe for all 6 roles at once (analogous to Word's Design-tab Style Set switcher). Exactly **one group is active per document at a time**; switching groups rewrites the concrete definitions of the same 6 named Word styles non-destructively (font/size/color/border change; the `basedOn`/outline-level relationships and Quick-Style visibility do not).

- **Built-in groups:** Petition/Araz, Maqala, General, Waaz.
- **Custom groups:** users can save the currently-configured field values as a new named group (parallel to how poem profiles are saved), which then appears alongside the built-ins in the group picker. Custom groups are stored in the document (traveling with the file), consistent with how poem profiles already work.

**Styles tab layout:**
- Group picker (built-ins + any custom groups saved in this document).
- "Save as new group" — captures current field values as a new group.
- Below the picker: the editable fields from §2's table, scoped to the **active group**. Editing + applying updates that group's definition only; other groups are untouched until selected and edited.
- Six apply buttons/entries (Heading 1/2/3, Emphasis, Quote, Quran Quote), each showing its current resolved values and an Apply action for the current selection.

## §4 — RTL document setup

A single **"Set up RTL document"** action, separate from style groups (document-wide plumbing every group needs identically), idempotent and safe to re-run, showing an applied/not-applied status rather than blindly re-running on every tab open. Applies:

1. **Normal style → paragraph direction RTL** (`bidi`) — otherwise cursor/alignment/list-bullet behavior defaults to LTR even with RTL script typed in.
2. **Normal style → Latin font** (the `ascii`/`hAnsi` font slot) — for embedded Latin text/numerals inside RTL paragraphs.
3. **Normal style → Complex Script font + size** (`cs`/`szCs`, independent of the Latin font/size) — Word tracks Latin and complex-script size independently; setting Latin size alone leaves CS text at its own default.
4. *(No separate step — covered by the Heading roles in §2.)* Word doesn't propagate Normal's `cs` font to Heading 1-3 automatically; each Ashaar Heading role's own font field already fixes this per level.
5. **Section-level RTL layout** (`sectPr/bidi`) — flips margins/column order and, importantly, fixes **footnote/endnote numbering and layout direction**, which follow the section's layout direction rather than the Footnote Text paragraph style alone.

Fields needed for this action: Latin font choice, Normal complex-script font choice (a separate dropdown from any Heading role's font — body prose font and heading display font are typically different).

## §5 — Testing

- **Node (pure):** style-recipe merging (group → 6 role definitions), OOXML fragment builders for style defs, RTL-setup OOXML fragments, emphasis point-bump computation given a base size.
- **Manual (Word):** switching groups reflows all 6 roles without touching unrelated content; custom group save/reload; each role's Apply button on a selection; Quran Quote nested inside/matching Quote's border; RTL setup applied once is idempotent on re-run; footnote numbering renders RTL after setup; Ashaar Heading 1/2/3 show correctly in the navigation pane/TOC.

## §6 — Non-goals

- **Bayan-quote tagging for annotations** — deferred to the annotation spec (§ "Explicitly out of scope" above).
- **Text conversion / find-replace (double-press, symbol substitution, honorific expansion)** — deferred to its own spec.
- **Custom keyboard shortcuts** — deferred fast-follow, requires shared-runtime conversion.
- **Mixing style groups within a single document** — one group is active per document; a document that needs to show, e.g., a Maqala quoting Araz-styled text is out of scope for v1 (would need per-block group tagging, closer to the poetry profile system).

## Resolved decisions

1. Styles are backed by **named Word styles**, `basedOn` the closest built-in, marked Quick Style — not direct formatting, and not overwriting built-ins in place.
2. **6 roles**, not a flat list: Heading 1/2/3 (each independently configurable), Emphasis, Quote, Quran Quote (based on Quote).
3. **One style group active per document**, chosen from 4 built-ins + custom saved groups; switching a group reconfigures all 6 roles at once.
4. Instance-level overrides (Quote width, Quran Quote line height) are direct formatting layered on top of the named style, distinct from style-level fields that affect every use. Emphasis has no instance-level override — its point-bump amount is style-level, and only the resulting absolute size varies per instance (computed from each selection's own base size).
5. RTL document setup is a **separate, group-independent action** covering Normal style bidi/fonts and section-level layout (which also fixes footnote RTL numbering); the heading complex-script-font checklist item is satisfied by the Heading roles themselves, not a separate step.
6. Annotations, text-conversion/find-replace, and keyboard shortcuts are explicitly deferred to their own specs/fast-follows.

## Post-implementation notes (2026-07-16, from live Word verification)

Manual testing in desktop Word drove several changes to the design above. These
notes are the accurate record; where they differ from the sections above, these win.

1. **Arabic sizing uses the complex-script metrics.** Arabic-script text renders
   at the complex-script font/size (`nameBidirectional`/`sizeBidirectional`, i.e.
   OOXML `w:cs`/`w:szCs`), not the Latin `size`. Every place that sets a font or
   size sets BOTH — headings, the RTL body style, footnotes, and the Emphasis
   size bump (which initially set only the Latin size, so the bump was invisible
   on Arabic).
2. **Inherited italic is cancelled on Emphasis, Quote, and Quran Quote.** Word's
   built-in "Emphasis" and "Quote" styles are italic, and Office.js materializes
   a base style's italic into each derived style at creation. Italic doesn't
   render in Arabic-script fonts, so `configureRoleStyle` explicitly sets
   `italic`/`italicBidirectional = false` on all three.
3. **Emphasis apply is two-sync.** Applying the character style and writing the
   size in one `context.sync()` reset the run size; the style is applied and
   committed first, then size is written as direct formatting.
4. **"Ashaar Normal" is a real body style** created by RTL setup (not a mutation
   of the built-in Normal, per decision #1). **Ashaar Quote is now `basedOn`
   "Ashaar Normal"** (not built-in "Quote") so quotes follow the body
   font/size/RTL by default; Quran Quote inherits it transitively. ensureAshaarStyles
   ensures a minimal Ashaar Normal exists before the quote roles.
5. **Paragraph reading-order (bidi) is not settable via the object model** — only
   `alignment`. Two mitigations: RTL setup right-aligns body/footnote styles; and
   a **bidi-carrying "Ashaar Normal" (plus RTL Footnote Text/Reference) is merged
   in via `insertFileFromBase64({importStyles:true})`** from a generated carrier
   .docx (`scripts/make-ashaar-normal-carrier.mjs` → `src/taskpane/ashaar-normal-carrier.js`),
   giving true RTL reading order where the API can't. Best-effort with a
   right-align-only fallback when WordApi 1.5 is unavailable.
6. **The footnote separator line cannot be set by the add-in** — it's a special
   footnote in `footnotes.xml`, not a style, so `importStyles` can't reach it.
   Right-aligned in the test-doc generator and left as a Word Draft-view /
   template step for real documents.
7. **Footnote text is 10pt** (capped at the body size).
8. **Real installed font-family names**: heading default is **"Kanz Al Marjaan"**,
   Waaz uses **"Fatemi Maqala"** (the earlier placeholder "Marjaan"/"Fatemi"
   didn't resolve).
9. **UI**: RTL setup is the first, always-open section in the Styles panel and all
   sections are expanded; the poetry-only chrome (settings panel, justify actions,
   Fonts strip) is hidden in Styles and Booklet modes; the two color fields use
   native color pickers.

**Verified working in Word by the user:** heading levels + fonts, emphasis
(red + size bump), Quote/Quran Quote (borders, no italic, inherit body),
insertFileFromBase64 style import for Ashaar Normal, RTL footnote text/reference,
footnote size, and the mode-based UI hiding. Still template-only: the footnote
separator line (and, longer-term, the Word-Online `.dotx` template reusing the
same bidi `styles.xml`).

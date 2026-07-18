# Text Conversion — Manual Word Verification Checklist

**Date:** 2026-07-16
**Feature:** Convert tab (legacy double-press ⇄ modern). Spec: `2026-07-16-text-conversion-design.md`. Plan: `plans/2026-07-16-text-conversion.md`.

Automated (node) tests cover the pure engine (ordering, kashida-escape, whole-word, round-trips, the sukun→khari-zabar mark row). The items below need a running Word (desktop or web) and a real legacy `.docx`.

## Setup
- [ ] `npm start` (or load the add-in against the dev server) and open the **Convert** tab.
- [ ] Confirm the tab shows: two direction buttons, a Scope radio (Whole document / Current selection), a Preset picker, and a grouped checklist (Letters, Marks) with every row checked. The sukun→khari-zabar row shows a ⚠ marker.

## To Modern (Legacy → modern)
- [ ] Open a real AL-KANZ legacy `.docx`. With Whole document + all rows checked, click **⟵ To Modern**.
- [ ] Doubled consonants convert (e.g. `حح`→`چ`, `كك`→`گ`); `؛`→`چھے`; an escaped double `سـس` becomes a genuine `سس` (not `ے`); sukun→khari zabar.
- [ ] Run formatting (bold/color/font runs) is preserved on converted text.
- [ ] The status line reports non-zero counts, e.g. "Converted N letters, M marks, 0 symbols."

## To Legacy (modern → double-press)
- [ ] On modern text (or the just-converted result, with the lossy sukun row **unchecked**), click **To Legacy ⟶**.
- [ ] Letters round-trip (`چ`→`حح`, …); a genuine double `سس` gains a separator tatweel (`سـس`); a standalone `چھے` word → `؛`, but `چھے` inside a larger word does **not** collapse to `؛` (its `چ`/`ے` still convert).

## Scope
- [ ] Select a passage, choose **Current selection**, run a direction → only the selection changes.

## Presets (roaming / cross-document)
- [ ] Uncheck some rows, "Save current as preset…", name it (e.g. "Letters only"), Save.
- [ ] Open a **different** document → the preset appears in the picker and, when chosen, re-applies that row selection.

## Lossy row
- [ ] Confirm the ⚠ marker/tooltip on sukun→khari zabar; verify that unchecking it leaves genuine sukuns untouched during To Modern.

## Feasibility spike (record findings here)
- [ ] `Range.search()` matches literal tatweel/diacritics (the escape ops `سـس` fire).
- [ ] `matchWholeWord: true` respects Arabic word boundaries for `چھے → ؛`. If it does **not**, fall back to a manual neighbor check (search bare `چھے`, verify surrounding chars before replacing).
- [ ] Sequential per-op search/replace within one `Word.run()` behaves (no stale-range errors).

## Findings
_(record dates, Word build, pass/fail, and any surprises)_

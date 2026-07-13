# Unified settings panel — manual Word verification

Run with `npm start` against test-documents/marsiya-test.docx.

- [ ] Fresh insert from Conversion tab creates a v3-tagged block; panel header shows `Poem` on click-in.
- [ ] Gap 8 → Apply on poem A only; poem B untouched; dot shown on A's gap; ⟲ + Apply reverts.
- [ ] Assign both poems to "Test"; Update "Test" changes both EXCEPT A's tweaked gap.
- [ ] Save as… from a tweaked poem clears its dots and creates the profile.
- [ ] Delete the profile via DevTools: the profile store lives in `Office.context.document.settings`
      (NOT browser `localStorage` — a prior draft of this checklist said localStorage, which is wrong;
      `loadProfileStore`/`saveProfileStore` read/write `Office.context.document.settings` under the key
      `ashaar:profiles`). To delete a profile for this test, either call
      `Office.context.document.settings.remove("ashaar:profiles")` followed by `saveAsync` from the
      console, or edit the JSON blob under that key to drop the one profile name, then `saveAsync`.
      After that: click in the poem → header shows `(not on this machine)` → Restore heals it.
- [ ] Plain-text selection: header `Selection`, Apply justifies, no content control created.
- [ ] Cell chip: strength override applies to one cell; Clear via ⟲ + Apply falls back.
- [ ] Gap chip: symbol + colors render in the spacing cell; `Set as default for all bandhs` propagates.
- [ ] Unmeasurable font prompts once; after registering via Fonts strip it never prompts again.
- [ ] Legacy document (v2 tags, from main branch): opens, panel reflects old settings as local dots, Apply upgrades tag to v3, nothing visually moves except requested changes.
      Known migration-fidelity limit: v2 tags never persisted the auto-fit/fixed width flag, so a
      legacy "fixed %" block may render as auto-fit after migration, with `widthPct` still shown as a
      local dot — there is no source data to recover which mode the block was actually in.
- [ ] Adopt Existing Table → Replace Selection (Table Input tab) round-trips.
- [ ] Apply twice in a row is idempotent (Debug dump: no growing nat/target/nSp/segs).
- [ ] Advanced → Font correction factor persists via Apply and Save as…
- [ ] Advanced → Debug colors toggle on/off and persist

Tester note: if Apply reports an error AFTER the tag write completes (poem scope), the settings ARE
already persisted to the tag — the reported failure is from the render/justify pipeline that runs
after the write, not the write itself. Re-click Apply to retry the render; there is no need to re-enter
the settings.

## Additional items accumulated from Tasks 7-9 (PENDING USER, not yet run in Word)

These were flagged as manual-verification-only in the Task 7/8/9 reports and are folded in here so
there is a single checklist to work through, rather than four scattered ones.

- [ ] **Save as… relies on `window.prompt()`** (`saveAsProfile`, Task 8) — verify native `prompt()`
      actually shows a dialog in Word's WebView. This is a known Office WebView risk: if `prompt()` is
      blocked/suppressed there, `saveAsProfile`'s `(prompt(...) || "").trim()` guard returns `""` and the
      function silently no-ops (no error message, nothing happens). If confirmed blocked, "Save as…"
      needs an inline name `<input>` in the panel instead of `window.prompt()`.
- [ ] **Font gate: no false-positive prompt** — a poem entirely in a registered/system font → Apply →
      no font prompt appears at all (Task 9, item 1; distinct from the "prompts once" item above, which
      only checks the *positive* case).
- [ ] **Font gate: "Continue anyway"** — when the unmeasurable-font prompt appears, clicking
      *Continue anyway* (as opposed to *Add font file…*) lets Apply proceed with approximate metrics
      instead of blocking (Task 9, item 3).
- [ ] Scoped apply on one poem of a shared profile: siblings untouched (width may differ until Update — expected).
- [ ] Second consecutive scoped apply (Justify on a profiled poem, no edits between) shows rebuild=no in the debug dump.

## Render workflow & batch formatting (Tasks 1-12, 2026-07-13)

- [ ] Sibling poems untouched by cell/gap/poem Apply on a shared profile — edit poem A's cell strength, Apply to cell only: poem B unchanged, no re-render.
- [ ] Re-render button (block scope) picks up native text edits and font-family changes in the document.
- [ ] Apply-to-all: cell strength to whole bandh, gap symbol to whole poem — each propagates to the full scope.
- [ ] Cell fill/color render — check "on" and select a color; Apply: the cell visibly fills and text darkens (or custom color applies).
- [ ] Cell fill/color clearing — uncheck "on" and Apply: cell returns white, text color returns black (known limitation: pre-override manual colors are not restored).
- [ ] Cell fill/color cross-poem seeding check — set fill on poem A's first cell, click poem B's first cell at the same position: pane fill/color controls show B's state (empty/defaults), not A's.
- [ ] Cell color clear retry — apply a color override, clear it, then force a render failure (e.g. font gate cancel); re-click Apply: cell text returns black on retry (tag is colorless; pending clear retained).
- [ ] Capture buttons disabled by cursor context — cursor in a gap cell: cell-scope Capture disabled; cursor in a content cell: gap-scope Capture disabled.
- [ ] Capture reads cell formatting (fill, text color) into the pane as read-only; uncheck "on" or edit a field to activate pending changes.
- [ ] Capture from theme-picker-shaded cell — verify the pane color swatch updates to show the selected color (Word may return a color NAME instead of #RRGGBB; if swatch stays stale while "on" is checked, this is a known deferred normalization issue — report it).
- [ ] Capture → Apply → Update Profile round trip — capture a layout, apply it to define a profile, Update an existing profile to persist the captured shape.
- [ ] Apply-to fan-out with ⟲-cleared fields — clear a field with ⟲, select "Whole poem" as apply target, Apply: the clear (null) propagates to every cell and all-null overrides are deleted from the tag.
- [ ] "Apply to" select does not reset after Apply — verify a leftover "Whole poem" selection from a prior Apply doesn't affect the next cell edit (known Minor).
- [ ] Line height (lineHeightPt) stops nastaliq clipping — set to ~1.8× font size on a clipping poem, Apply: lines no longer collide vertically.
- [ ] Separator (separatorPt) sizing — set to 1pt, Apply: tables nearly touch, never merge; increase the setting: inter-poem gaps visibly widen.
- [ ] Cost caption accuracy — refresh-cost caption correctly distinguishes rebuild (full re-calibration) from re-justify (same params, new text) based on the apply scope and context.
- [ ] Debug metrics table — Debug toggle on; Apply: metrics table appears with phase timings (Probe, Calibrate, Justify); hit/miss recorded for calibration memos.
- [ ] Second Apply idempotency speed — Apply once to a poem, then Apply again without edits: second Apply is visibly faster (cached calibration memo hit).

## Notes for whoever runs this

- Task 9's report flagged `applyCellOverride`/`applyBandhWidth` as older apply entry points that were
  NOT covered by the font-measurability gate. Task 10 (this cleanup) deleted both functions entirely
  (dead code — the DOM controls they wrote to no longer exist in `taskpane.html`; the panel's Apply
  button is the only surviving entry point for cell/bandh edits), so that gap is now moot rather than
  outstanding.

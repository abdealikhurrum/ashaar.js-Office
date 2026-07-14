# Ashaar for Word — Quick How-To

A short guide to the everyday workflow. Deeper walkthroughs (adopting tables, grid layouts, Nastaliq fonts, bandh modes, decor): [HOW-TO-WORKFLOWS.md](HOW-TO-WORKFLOWS.md). Full reference: [USER_GUIDE.md](USER_GUIDE.md).

## 1. Enter a poem

Open the task pane (**Home → Ashaar Poetry**), go to the **Conversion** tab and paste your poetry:

```
first misra \ second misra      ← one bayt (couplet)
a single full-width line |      ← single-misra line
refrain line *                  ← marked as refrain

(blank line = new stanza/bandh)
---                             ← separates poems
```

The live preview updates as you type. Click **Insert as Table** to place the poem in the
document as a formatted table block (this is what the settings panel manages), or
**Insert as Paragraphs** for tab-stop layout instead of tables.

## 2. Format with the Settings panel

Click anywhere inside an inserted poem — the panel header shows **Poem** and the scope chips
(**Poem / Bandh / Cell / Gap**) light up for wherever the cursor is.

- Edit any field (Justification, Fill mode, Kashida strength, Misra gap, Width…). Nothing
  touches the document until you press **Apply**.
- A dot next to a field means it differs from the assigned profile. Press **⟲** on a field
  then Apply to clear that override and fall back.
- The caption above **Apply** tells you the cost in advance: *re-justify* (fast, in place)
  vs *rebuilds poem tables* (structural — gap, width, layout, line height, separator).
- **Apply only changes the poem you're in.** Other poems on the same profile are untouched
  until you press **Update** on the profile.

## 3. Profiles (shared styles)

- **Save as…** — save the current poem's settings as a named profile.
- **Assign** — put the selected profile on this poem.
- **Update** — push this poem's current settings into the profile *and refresh every poem
  that uses it*.
- **Restore profile from this poem** — heal a profile that exists in the document but not
  on this machine.

Profiles live inside the document, so they travel with the file.

## 4. Cells and gaps

Click a specific cell (or a spacing/gap cell) and pick the **Cell** / **Gap** chip:

- **Cell:** per-cell Strength, Target width, Cap lift, plus **Fill color** and **Text
  color** (tick **on** to set; untick + Apply to clear — cleared text color returns to
  black).
- **Gap:** separator **Symbol**, its fill and color.
- **Apply to** — fan one edit out to *This cell / This bandh / Whole poem*. Only the fields
  you actually changed are written to the other cells; it resets to "This cell" after each
  Apply.
- **Capture formatting** — reads the cursor cell's existing shading/text color (and gap
  symbol) into the pane without touching the document; press Apply to persist it.

## 5. Line height & separators (Advanced)

- **Min line height (pt)** — stops tall Nastaliq lines clipping; try ~1.8× the font size.
- **Table separator (pt)** — the exact gap between poem tables; 1pt keeps them nearly
  touching without merging.

## 6. After editing text in Word

If you typed inside the poem or changed its font natively, press **Re-render** (footer,
enabled when the cursor is in a poem). It rebuilds the tables with the current settings
while keeping font, size, and your per-cell formatting.

## 7. If something looks wrong

- **"Font not measurable" prompt:** register the font once via the Fonts strip (or *Continue
  anyway* for approximate metrics). It won't ask again.
- **Kashida looks broken after an Apply:** press Apply again (a failed render after a
  successful save is retryable), or Re-render the poem.
- **Start over on a poem:** Reset (clear kashida & spaces), then Apply your settings fresh.

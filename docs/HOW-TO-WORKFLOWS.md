# Ashaar for Word — Workflows Guide

The deeper companion to the [Quick How-To](HOW-TO.md): converting existing tables, building
layouts from the grid, Nastaliq fonts, couplets vs. complex bandhs, and decorating cells and gaps.

## 1. Convert an existing table

You have a poem already sitting in a Word table (typed by hand, or pasted from elsewhere) and
want it to become a managed Ashaar block.

1. In Word, click inside the table (or select several tables at once).
2. Task pane → **Table Input** tab → **Adopt Existing Table**.
3. The add-in reads the cells, reconstructs the poem text, switches to the Convert editor,
   and — by default — immediately replaces the table with a clean, justifiable Ashaar block.

Options (under **Adopt options**):

- **Review before replacing** — stops after step 3 so you can fix the recovered text in the
  editor first; click **Replace Selection** when satisfied.
- **Reading direction** — leave on *Auto (right-to-left)* unless the table reads oddly.
- **Scope** — *Table at cursor* or *All tables in selection* (each selected table becomes
  one stanza).

If the recovered lines pair up wrong, open the editor's import options: **Split hemistichs
on** lets you pick the separator the original used (dash, tab, double-space…), and **Pair
every 2 lines (no separator)** treats consecutive lines as couplets when the source has no
separator at all.

## 2. Build a table from the grid

For laying out *blank* tables you'll type into — the Table Input tab's main flow.

### Quick: Draw Table

Set **Bandhs** and **Misras per bandh**, pick a **Layout preset**:

| Preset | Shape |
|---|---|
| Paired rows | classic two-column couplets (`2 - 1`, `4 - 3`, …) |
| Centered stack | every misra centered, one per row |
| Alternating sides | misras alternate right / left |
| Indented stack | staircase indентation |
| 3 + center + refrain | the marsiya shape: a 3-misra row, a centered solo, then a pair |

Click **Draw Table**. The layout appears as an editable spec first — switch between the
**Grid** view (12-column bubbles you click to shape rows) and the **Numbers** view (text
spec) to customize before drawing. Spec syntax, one row per line:

```
2 - 1        two cells: right cell "1", left cell "2"
3 | 2 | 1    up to four cells in one row
<4>          centered solo misra
1 >          right-side-only misra        < 2   left-side-only
  indented   leading spaces = indent      (used by Indented stack)
```

### Freeform: Drop Grid → Capture

When no preset fits:

1. **Drop Grid** inserts a blank 12-column grid table.
2. In Word, **merge and reshape cells natively** until the row shapes match your layout.
3. Back in the pane, type a name under **Templates** and click **Capture from Word** —
   the add-in reads your merged column widths and saves the shape as a named template.
4. **Apply** inserts that template anywhere, any time. **Export JSON / Import JSON** moves
   templates between machines and documents.

## 3. Nastaliq fonts: Mehr, Jameel, and adding your own

There is no font picker in the pane — **set the font in Word itself** (Home tab), like any
other text. The add-in reads each run's real font and automatically picks the right
stretching mechanism:

| Font in Word | Mechanism | What stretching looks like |
|---|---|---|
| **Mehr Nastaliq Web** | trailing tatweel | eligible letters (ب ت ٹ ث ف ک گ …) gain a flowing tail |
| **Jameel Noori Nastaleeq** | font-swap | whole words swap to the wider *Jameel Kasheeda MarkSafe* face |
| Noto Nastaliq Urdu, Gulzar, Scheherazade New | word-spacing | space distribution only (these designs don't stretch well) |
| Anything else (Fatemi Maqala, Arial, …) | generic kashida | straight tatweel strokes inserted at legal joins |

Practical notes:

- **Mehr** gives the most even, calligraphic fills for Urdu — its tails are designed glyphs.
- **Jameel** fills by swapping words into its Kasheeda face; the effect is bolder and
  chunkier per word. Known limit: heavily **vocalized** (harakat) text barely stretches in
  Jameel — the wide glyph variants collapse under marks. Prefer Mehr or bare text there.
- Mixed-font poems are fine — the mechanism is chosen per run, so a Jameel misra and a
  Mehr misra in the same poem each justify their own way.

### Adding a font

Most fonts installed on your machine measure automatically — you don't need to do anything.
The **Fonts** strip (bottom of the pane) exists for fonts Word can render but the pane's
sandbox can't see (typical on Windows): pick the file, confirm the **Register as** name
matches what Word shows (auto-filled from the font file when readable), click **Add font**,
then **Verify at cursor** with the cursor in text using that font. The font is stored
locally, never uploaded. If a font is unmeasurable at Apply time, the pane prompts you once
with the same *Add font file…* flow.

Keep custom font family names **31 characters or fewer** — Word silently truncates longer
names, which breaks the name matching.

## 4. Poem mode: couplets vs. complex bandhs

Both come from the same Convert-tab text; the difference is line grouping and layout.

### Couplet mode (ghazal)

Type each bayt on one line, hemistichs separated by `\`, **no blank lines**:

```
دل ناداں تجھے ہوا کیا ہے \ آخر اس درد کی دوا کیا ہے
ہم ہیں مشتاق اور وہ بیزار \ یا الٰہی یہ ماجرا کیا ہے
```

Consecutive lines become one multi-row table — one bayt per row, two columns. With
**Fill mode: Natural-fit (harmony)**, each column equalizes to its own longest line across
all the couplets; **Cell-fit** fills every cell to its edge.

### Complex bandh mode (marsiya, musaddas…)

Separate stanzas with a **blank line**; use `\` for multi-misra rows and `|` for full-width
solo lines within a bandh:

```
پہلا مصرع \ دوسرا مصرع \ تیسرا مصرع
چوتھا مصرع|
پانچواں مصرع \ چھٹا مصرع
```

Each bandh becomes its own table with the row shapes your lines imply (a 3-misra row, a
solo, a pair — the classic 3/1/2). Other markers: `%` at a line's end marks a refrain,
`---` on its own line starts a new poem. For *blank* complex layouts you'll type into
later, use the Table Input presets instead (**3 + center + refrain** is the marsiya shape).

Harmony works across bandhs too: same-position cells in every bandh's tables share a
justification target, so the whole poem's columns line up.

## 5. Colors & symbols

Click into a cell or gap (the narrow separator column) and pick the **Cell** or **Gap**
chip in the Settings panel.

**Cells** — per-cell **Strength**, **Target width**, **Cap lift**, plus **Fill color** and
**Text color** (tick **on** to set; untick + Apply clears — cleared text returns to black).

**Gaps** — **Symbol** (e.g. `؎`; blank = none), **Fill color**, **Symbol color**. Gap decor
applies instantly, with or without a profile.

Shared tools in both bodies:

- **Apply to** — fan one edit out: *This cell/gap*, *This bandh*, **Same cell/gap in all
  bandhs** (every A1 across the poem), or *Whole poem*. Only the fields you actually
  changed propagate — other cells keep their own settings. The select snaps back to
  "This…" after each Apply.
- **Capture formatting** — reads the cursor cell's existing shading/text color (and gap
  symbol) into the pane without touching the document; Apply persists it.
- **Set as default for all bandhs** (gap body) — writes the current symbol/colors into the
  poem's assigned profile as the default for that slot, so every poem on the profile picks
  it up. Requires a profile to be assigned first.

Tip: decorate one bandh's gap the way you like, then *Same gap in all bandhs* → Apply — the
fastest way to a consistent poem.

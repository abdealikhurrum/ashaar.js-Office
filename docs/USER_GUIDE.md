# Ashaar Poetry for Word — User Guide

Ashaar Poetry is a Word task-pane add-in for typesetting Arabic, Urdu, and Persian
poetry: proper right-to-left couplets, stanzas, and refrains, with professional
kashida (tatweel) justification and flexible table layouts.

## Getting started

1. Open Word and load the add-in (your administrator or the install link provides it).
2. On the **Home** tab, click **Ashaar → Poetry** to open the task pane.
3. The pane has two modes, selected by the tabs at the top:
   - **Table Input** — draw a layout first, then type poetry into the cells (composing in Word).
   - **Ashaar.js Convert** — paste or load poetry text and convert it into a formatted block.
4. For testing the different modes, you can use test-documents/adopt-test.docx This file has a myriad of different bandhs to illustrate
how existing poetry can be converted and enhanced using the add-in.

**Which mode do I want?** If you already have the poem as text, use **Convert**. If you are
composing from scratch inside Word, use **Table Input**.

## The poetry markup

A few characters describe the structure of a poem. You only need these:

```
sadr \ ajuz      a couplet — two hemistichs split by a backslash
a bare line      a solo misra — full width, centred
(blank line)     a stanza break
---              a new poem
line %           a refrain (coloured)
```

`|` and `*` also work as hemistich separators, and a line with three or more parts
(`m1 \ m2 \ m3`) becomes a multi-misra row.

Example:

```
دل ناداں تجھے ہوا کیا ہے \ آخر اس درد کی دوا کیا ہے
ہم ہیں مشتاق اور وہ بیزار \ یا الٰہی یہ ماجرا کیا ہے
```

## Shared formatting controls

These sit at the top of the pane and apply to whatever you insert:

- **Font** — *Document default*, *Arabic serif* (Scheherazade/Amiri), or *Nastaliq*.
- **Justification** — *Kashida* (real tatweel stretching), *Word justify*, *Spacing preview*, or *None*.
- **Kashida strength** — how aggressively kashida fills each line (0 = off).
- **Middle gap** — the spacing between the two columns of a couplet.

## Convert poetry you already have

1. Switch to **Ashaar.js Convert**.
2. Get your text into the editor:
   - **Paste** it, or
   - **Load Selection** — select the poem in your Word document and click this to pull it in.
3. Watch the **Conversion Preview** update live.
4. Insert it:
   - **Insert as Table** — a formatted RTL table block.
   - **Insert as Paragraphs** — tab-stop paragraphs instead of a table.
   - **Replace Selection** — replace the selected text in your document, in place.

### Import options (separators)

Pasted or loaded text is **auto-converted** to standard `\` form — so poems that use a
dash, tab, asterisk, or a wide gap between hemistichs just work, and a short note tells
you what was detected. Open **Import options** if a poem didn't split correctly:

- **Split hemistichs on** — force a specific separator (backslash, asterisk, pipe, dash,
  tab, double space) or **Custom…** for your own.
- **Pair every 2 lines** — for files that store one hemistich per line with no separator;
  pairs consecutive lines into couplets (blank lines still break stanzas).
- **Apply to editor** — re-run the conversion after changing an option.

## Adopt an existing table

If your poem is already in a Word table:

1. Click inside the table (in **Table Input** mode).
2. Click **Adopt Existing Table**.
3. The table is read, cleaned, and replaced with a managed Ashaar block, and the recovered
   text appears in the editor.

Open **Adopt options** to adjust:

- **Review before replacing** — load the recovered text and preview first; you press
  *Replace Selection* yourself when it looks right.
- **Reading direction** — Auto (right-to-left) by default; switch to LTR if the hemistichs
  come out reversed.
- **Scope** — the table at the cursor, or all tables in the selection.

Adopting strips any old kashida, ignores empty/gap columns, and re-justifies using your
current settings. Use **Ctrl+Z** to undo if needed.

## Compose directly in Word (Table Input)

1. Choose the shape: **Bandhs** (stanzas), **Misras per bandh**, and a **Layout preset**
   (Paired rows, Centered stack, Alternating sides, Indented stack, or 3 + center + refrain).
2. Fine-tune the **Layout spec** box if you want exact control over rows.
3. Click **Draw Table** to insert a blank, editable table, then type your poetry into the cells.
4. Alternatively, **Drop Grid** inserts a blank 12-column grid you can merge and reshape
   natively in Word — then **Capture from Word** saves that layout as a reusable template.

## Justify and balance a poem

1. Click anywhere inside an inserted **Ashaar Poem** block.
2. Click **Justify Selected Text**.

The add-in measures your font and column widths and fills each hemistich with kashida,
balancing all misras to a consistent width. Justification is fully **re-justifiable**: run
it again after changing the font, column width, or kashida strength and it recomputes from
scratch — it never piles tatweels on top of old ones, and can reduce or remove them.

## Reusable layouts (templates)

In **Table Input → Templates**:

- **Capture from Word** — name a template and save the layout of the selected table.
- **Apply** — insert a saved template.
- **Export JSON / Import JSON** — move templates between documents or share them.

Templates turn a layout you tuned by hand (including column widths you dragged in Word)
into something you can reuse anywhere.

## Tips and limitations

- **Don't delete the content control.** Inserted poems are wrapped in an "Ashaar Poem"
  content control — that's how *Justify Selected Text* finds the poem. Clicking inside it
  is how you re-justify later.
- **Column widths are a starting estimate.** They're computed from font metrics; you can
  drag column borders in Word to fine-tune, then *Capture* the result as a template.
- **Nastaliq is for output, not editing.** The editor uses a legible naskh; choose Nastaliq
  from the Font control to render the inserted poem in Nastaliq.
- **Tabs and wide gaps** from a Word *selection* depend on Word preserving them; if a tab- or
  space-delimited poem doesn't split, paste it instead, or use *Import options → Split on*.
- **Undo** (Ctrl+Z) reverts any insert, replace, adopt, or justify action.

## Quick reference

| Control | Where | What it does |
| --- | --- | --- |
| Insert as Table | Convert | Insert poem as a formatted RTL table |
| Insert as Paragraphs | Convert | Insert as tab-stop paragraphs |
| Replace Selection | Convert | Replace selected text with the formatted poem |
| Load Selection | Convert | Pull selected document text into the editor |
| Import options | Convert | Separator auto-detect + overrides, pair-by-line |
| Adopt Existing Table | Table Input | Convert an existing table into a managed Ashaar block |
| Draw Table | Table Input | Insert a blank editable table from the layout spec |
| Drop Grid | Table Input | Insert a blank 12-column grid to reshape in Word |
| Capture from Word | Table Input | Save the selected table's layout as a template |
| Justify Selected Text | both | Apply / recompute kashida justification on a poem |

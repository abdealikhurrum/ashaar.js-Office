const assert = require("assert");
const { adoptTableToSource, blankSpacingCells, stripDecorCells } = require("../src/taskpane/table-adopt");

// rows: array of rows; each row = array of cell texts in LOGICAL (document) order.
// direction "rtl" (default): logical first cell = visual right = sadr.
// direction "ltr": reverse each row so the visual-right cell becomes sadr.

// Couplet (RTL): logical [sadr, ajuz] → "sadr \ ajuz"
assert.equal(
  adoptTableToSource([["سادر", "عجز"]]),
  "سادر \\ عجز",
  "2-cell row becomes a couplet"
);

// Solo misra: single cell → bare line, no separator
assert.equal(adoptTableToSource([["مصرع"]]), "مصرع", "1-cell row becomes a bare solo line");

// Multi-misra row (3+): joined with separators on one line
assert.equal(
  adoptTableToSource([["م١", "م٢", "م٣"]]),
  "م١ \\ م٢ \\ م٣",
  "3-cell row becomes a multi-misra line"
);

// Mixed stanza: couplet, solo, couplet → one stanza, newline-separated
assert.equal(
  adoptTableToSource([["a1", "a2"], ["solo"], ["b1", "b2"]]),
  "a1 \\ a2\nsolo\nb1 \\ b2",
  "rows join with newlines into one stanza"
);

// Gap / empty / whitespace-only cells are ignored when counting and emitting
assert.equal(
  adoptTableToSource([["sadr", "", "ajuz"]]),
  "sadr \\ ajuz",
  "empty gap cell between hemistichs is dropped"
);
assert.equal(
  adoptTableToSource([["sadr", "   ", "ajuz"]]),
  "sadr \\ ajuz",
  "whitespace-only gap cell is dropped"
);
assert.equal(
  adoptTableToSource([["m1", "", "m2", "", "m3"]]),
  "m1 \\ m2 \\ m3",
  "interleaved gap columns (our own generator's shape) are dropped"
);

// All-empty row is skipped entirely (no blank line emitted)
assert.equal(
  adoptTableToSource([["", ""], ["solo"]]),
  "solo",
  "all-empty rows are skipped"
);

// Justification artifacts stripped: tatweel (U+0640), hair (U+200A), thin (U+2009)
assert.equal(
  adoptTableToSource([["قفاــ نبك", "ذكرى"]]),
  "قفا نبك \\ ذكرى",
  "tatweels are stripped from adopted cells"
);
assert.equal(
  adoptTableToSource([["a\u200Ab\u2009c", "d"]]),
  "abc \\ d",
  "inserted hair/thin micro-spaces are stripped"
);

// Internal newlines and runs of whitespace collapse to a single space
assert.equal(
  adoptTableToSource([["line1\nline2", "x"]]),
  "line1 line2 \\ x",
  "multi-paragraph cell collapses newlines to a space"
);
assert.equal(
  adoptTableToSource([["a   b", "c"]]),
  "a b \\ c",
  "runs of spaces collapse to one"
);

// LTR direction reverses each row so the visual-right cell becomes sadr
assert.equal(
  adoptTableToSource([["X", "Y"]], { direction: "ltr" }),
  "Y \\ X",
  "ltr reverses cell order"
);
assert.equal(
  adoptTableToSource([["X", "Y"]], { direction: "rtl" }),
  "X \\ Y",
  "rtl keeps logical order"
);

// Empty input
assert.equal(adoptTableToSource([]), "", "no rows → empty source");
assert.equal(adoptTableToSource(null), "", "null rows → empty source");

// ── blankSpacingCells: gap-decor text must never re-enter the source ─────────
// (gap-corruption fix, Part A): a spacing cell carrying a decorative symbol
// ("٭") is structure, not poetry — blank it before adoptTableToSource so a
// bayt row [m1 | ٭ | m2] reconstructs as "m1 \ m2", never a 3-misra line.
{
  // Decorated gap blanked; content preserved.
  assert.deepStrictEqual(
    blankSpacingCells([["م١", "٭", "م٢"]], [["c", "g", "c"]]),
    [["م١", "", "م٢"]],
    "spacing cell text blanked per the pattern"
  );
  assert.equal(
    adoptTableToSource(blankSpacingCells([["م١", "٭", "م٢"]], [["c", "g", "c"]])),
    "م١ \\ م٢",
    "decorated gap no longer becomes a misra"
  );
  // Undecorated (already-blank) gaps unchanged; content untouched.
  assert.deepStrictEqual(
    blankSpacingCells([["م١", "", "م٢"], ["", "تنہا", ""]], [["c", "g", "c"], ["g", "c", "g"]]),
    [["م١", "", "م٢"], ["", "تنہا", ""]],
    "blank gaps and content cells pass through"
  );
  // Solo row with decorated pads.
  assert.deepStrictEqual(
    blankSpacingCells([["٭", "تنہا مصرعہ", "٭"]], [["g", "c", "g"]]),
    [["", "تنہا مصرعہ", ""]],
    "solo-row pad cells blanked"
  );
  // Shape mismatch (hand-edited table) → rows returned unchanged: the caller
  // must fall back to the decor-symbol strip, not trust a misaligned pattern.
  assert.deepStrictEqual(
    blankSpacingCells([["م١", "٭", "م٢"]], [["c", "g", "c", "g", "c"]]),
    [["م١", "٭", "م٢"]],
    "row-length mismatch → unchanged"
  );
  assert.deepStrictEqual(
    blankSpacingCells([["م١", "٭", "م٢"], ["س"]], [["c", "g", "c"]]),
    [["م١", "٭", "م٢"], ["س"]],
    "row-count mismatch → unchanged"
  );
  // No pattern → unchanged (unmanaged tables keep current behavior).
  assert.deepStrictEqual(
    blankSpacingCells([["a", "b"]], null),
    [["a", "b"]],
    "null pattern → unchanged"
  );
  // Input rows are not mutated.
  const inRows = [["م١", "٭", "م٢"]];
  blankSpacingCells(inRows, [["c", "g", "c"]]);
  assert.deepStrictEqual(inRows, [["م١", "٭", "م٢"]], "input rows not mutated");
}

// ── stripDecorCells: defense when the pattern can't be aligned ────────────────
// (gap-corruption fix, Part B): with no trustworthy cell map, a cell whose
// ENTIRE trimmed text equals one of the block's decor symbols is structure,
// never a legitimate misra — blank it before reconstruction.
{
  assert.deepStrictEqual(
    stripDecorCells([["م١", "٭", "م٢"]], ["٭"]),
    [["م١", "", "م٢"]],
    "symbol-only cell blanked"
  );
  // Symbol embedded INSIDE a misra is content — kept.
  assert.deepStrictEqual(
    stripDecorCells([["م١ ٭ م٢", "م٣"]], ["٭"]),
    [["م١ ٭ م٢", "م٣"]],
    "symbol inside real text is kept"
  );
  // Whitespace / tatweel padding around the symbol still matches (cleanCell rules).
  assert.deepStrictEqual(
    stripDecorCells([["  ٭  ", "م١"]], ["٭"]),
    [["", "م١"]],
    "padded symbol-only cell blanked"
  );
  // Multi-char decor strings match whole-cell too.
  assert.deepStrictEqual(
    stripDecorCells([["***", "م١"]], ["***"]),
    [["", "م١"]],
    "multi-char decor blanked"
  );
  // Empty/blank symbols never blank anything (an empty cell is already a gap).
  assert.deepStrictEqual(
    stripDecorCells([["م١", "", "م٢"]], ["", "   ", null]),
    [["م١", "", "م٢"]],
    "empty symbols are ignored"
  );
  // No symbols → unchanged.
  assert.deepStrictEqual(
    stripDecorCells([["م١", "٭"]], []),
    [["م١", "٭"]],
    "no decor symbols → unchanged"
  );
  // Input rows are not mutated.
  const inRows2 = [["٭", "م١"]];
  stripDecorCells(inRows2, ["٭"]);
  assert.deepStrictEqual(inRows2, [["٭", "م١"]], "input rows not mutated");
}

console.log("table-adopt tests passed");

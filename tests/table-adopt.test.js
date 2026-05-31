const assert = require("assert");
const { adoptTableToSource } = require("../src/taskpane/table-adopt");

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

console.log("table-adopt tests passed");

(function (root, factory) {
  if (typeof module !== "undefined" && module.exports) {
    module.exports = factory();
  } else {
    root.AshaarTableAdopt = factory();
  }
}(typeof globalThis !== "undefined" ? globalThis : this, function () {
  // Strip elongation artifacts (tatweel U+0640, hair space U+200A, thin space
  // U+2009), collapse any run of whitespace (incl. newlines) to one space, and
  // trim. A cell that becomes empty is treated as a gap/blank.
  function cleanCell(text) {
    return String(text == null ? "" : text)
      .replace(/[\u0640\u200A\u2009]/g, "")
      .replace(/\s+/g, " ")
      .trim();
  }

  // Reconstruct canonical Ashaar source from a table's cells.
  //   rows: array of rows; each row = array of cell texts in LOGICAL (document) order.
  //   opts.direction: "rtl" (default) keeps logical order (logical-first = visual right
  //     = sadr); "ltr" reverses each row so the visual-right cell becomes sadr.
  // One table → one stanza: rows join with "\n". Per row, by non-empty cell count K:
  //   0 → skip, 1 → bare solo line, 2 → "sadr \ ajuz", 3+ → "m1 \ m2 \ …".
  function adoptTableToSource(rows, opts) {
    opts = opts || {};
    var ltr = opts.direction === "ltr";
    var lines = [];
    (rows || []).forEach(function (row) {
      var cells = (row || []).map(cleanCell);
      if (ltr) cells = cells.slice().reverse();
      var nonEmpty = cells.filter(function (c) { return c.length > 0; });
      if (!nonEmpty.length) return;
      lines.push(nonEmpty.join(" \\ "));
    });
    return lines.join("\n");
  }

  return { adoptTableToSource: adoptTableToSource, cleanCell: cleanCell };
}));

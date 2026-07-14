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

  // Gap-corruption fix, Part A: spacing cells must never re-enter a
  // reconstructed source. Given a table's cell texts (rows, logical order)
  // and the block's persisted content/gap pattern for that table
  // (Array<Array<"c"|"g">>, same emission order), blank every "g" cell so a
  // decorated gap (e.g. "٭") is dropped by adoptTableToSource instead of
  // being read back as a misra. If the pattern does not align with the live
  // shape (row count / per-row cell count), rows are returned UNCHANGED —
  // the caller must fall back to stripDecorCells. Pure; never mutates rows.
  function blankSpacingCells(rows, pattern) {
    if (!rows || !pattern || pattern.length !== rows.length) return rows;
    for (var i = 0; i < rows.length; i++) {
      if (!pattern[i] || !rows[i] || pattern[i].length !== rows[i].length) return rows;
    }
    return rows.map(function (row, r) {
      return row.map(function (cell, c) { return pattern[r][c] === "g" ? "" : cell; });
    });
  }

  // Gap-corruption fix, Part B (defense in depth): when no trustworthy cell
  // map exists, blank any cell whose ENTIRE cleaned text equals one of the
  // block's decor symbols — a short decor string is structure, never a
  // legitimate misra. Symbols are compared under the same cleanCell rules as
  // the cells; empty/blank symbols are ignored. Pure; never mutates rows.
  function stripDecorCells(rows, symbols) {
    var syms = {};
    var any = false;
    (symbols || []).forEach(function (s) {
      var clean = cleanCell(s);
      if (clean) { syms[clean] = true; any = true; }
    });
    if (!rows || !any) return rows;
    return rows.map(function (row) {
      return (row || []).map(function (cell) {
        return syms[cleanCell(cell)] ? "" : cell;
      });
    });
  }

  return {
    adoptTableToSource: adoptTableToSource,
    cleanCell: cleanCell,
    blankSpacingCells: blankSpacingCells,
    stripDecorCells: stripDecorCells
  };
}));

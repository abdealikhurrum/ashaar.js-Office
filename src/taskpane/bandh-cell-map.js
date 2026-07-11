/**
 * AshaarCellMap — label a bandh's content/gap PATTERN into positional cell ids
 * and validate a live Word table's shape against a stored pattern (pure, UMD;
 * node-testable, no DOM/Office dependency).
 *
 * A pattern is Array<Array<"c"|"g">>: outer = rows top-to-bottom, inner = cells
 * in OOXML emission order. Rows map to letters A,B,C…; content cells ("c") are
 * numbered 1,2,3… in emission order (= misra reading order = rightmost-first in
 * RTL), giving labels A1,A2,B1,C1,C2. Gap cells ("g") get no number — a `slot`
 * key "<letter>#<n>" instead (the future anchor for hemistich symbols / bandh
 * numbers / annotations).
 *
 * See docs/superpowers/specs/2026-07-11-bandh-cell-map-design.md.
 */
(function (root, factory) {
  if (typeof module !== "undefined" && module.exports) module.exports = factory();
  else root.AshaarCellMap = factory();
}(typeof globalThis !== "undefined" ? globalThis : this, function () {
  "use strict";

  var LETTERS = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
  function rowLetter(r) { return LETTERS.charAt(r) || ("R" + (r + 1)); }

  // pattern → flat ordered cell-map (emission order preserved).
  function buildBandhCellMap(pattern) {
    var out = [];
    var idx = 0;
    (pattern || []).forEach(function (row, r) {
      var letter = rowLetter(r);
      var contentN = 0, spacingN = 0;
      (row || []).forEach(function (tok) {
        if (tok === "c") {
          contentN++;
          out.push({ index: idx++, row: r, kind: "content", label: letter + contentN, slot: null });
        } else {
          spacingN++;
          out.push({ index: idx++, row: r, kind: "spacing", label: null, slot: letter + "#" + spacingN });
        }
      });
    });
    return out;
  }

  // Does a stored pattern match the live table's shape? perRowCounts is the
  // number of cells in each live row, in order. A mismatch (hand-edited/adopted
  // table) tells the caller to fall back to geometric inference.
  function alignPatternToTable(perRowCounts, pattern) {
    if (!pattern || !perRowCounts) return false;
    if (pattern.length !== perRowCounts.length) return false;
    for (var i = 0; i < pattern.length; i++) {
      if (!pattern[i] || pattern[i].length !== perRowCounts[i]) return false;
    }
    return true;
  }

  return {
    buildBandhCellMap: buildBandhCellMap,
    alignPatternToTable: alignPatternToTable,
  };
}));

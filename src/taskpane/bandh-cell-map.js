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
  if (typeof module !== "undefined" && module.exports) {
    module.exports = factory(require("./cell-overrides"));
  } else {
    root.AshaarCellMap = factory(root.AshaarOverrides);
  }
}(typeof globalThis !== "undefined" ? globalThis : this, function (AshaarOverrides) {
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

  // A cell-map entry's override/decor key, in the SAME scheme reflectActiveCell
  // uses at the cursor: content → overrideKey(tableIndex, label); spacing →
  // overrideKey(tableIndex, slot).
  function keyForEntry(tableIndex, entry) {
    return AshaarOverrides.overrideKey(tableIndex, entry.kind === "content" ? entry.label : entry.slot);
  }

  // Harmony-pooling key: which cells across DIFFERENT ROWS of a mapped
  // (bandh-cell-map) table balance to one width. The label/slot alone (A1,
  // B1, C1…) can't be used directly — the row letter makes every row its own
  // pool, which was the bug (see natural-width-matrix.js positionKey). Simply
  // stripping the row letter isn't safe either: a marsiya bandh can stack
  // rows of DIFFERENT shape (3 content cells, then a solo, then a flanked
  // pair) where the ordinal "1" recurs in every row but refers to physically
  // different cells that must never share a width. So the group key is the
  // ordinal (content/spacing number within its row, i.e. the label/slot with
  // its row letter stripped) PLUS the row's own token-shape ("c"/"g" pattern
  // joined) — two cells only pool when both the ordinal AND the row shape
  // match, which holds exactly when the rows are structurally identical
  // (a couplet's rows repeated down a single table) and never holds across
  // rows of a different content layout. `kind` is folded in too so a content
  // ordinal and a spacing ordinal can never collide.
  function columnGroupKey(pattern, entry) {
    entry = entry || {};
    var rowShape = (pattern && pattern[entry.row]) ? pattern[entry.row].join("") : "";
    var raw = entry.kind === "content" ? entry.label : entry.slot;
    var ordinal = String(raw || "").replace(/^[A-Za-z]+#?/, "");
    return entry.kind + ":" + rowShape + ":" + ordinal;
  }

  // Fan one Apply out to the keys it targets. `map` is the current table's cell
  // map (buildBandhCellMap output for the table containing currentKey) — used
  // for "this"/"bandh". `tables` is the block's full per-table pattern list
  // (payload.cells) — needed for "poem" to enumerate every table. currentKey is
  // an override/decor key already in the "tableIndex:label" scheme (see
  // AshaarOverrides.overrideKey), so the table index is recovered from it —
  // the map itself carries no table index.
  //
  //   kind      ∈ "content" | "spacing"
  //   mode      ∈ "this" | "bandh" | "poem" | "position"
  //   currentKey  the key of the cell/gap under the cursor
  //   tables    payload.cells (array of patterns, one per table) — required
  //             for mode "poem"/"position"; ignored otherwise.
  //
  // Containment holds by construction: "this" ⊆ "bandh" (same table) ⊆ "poem"
  // (all tables); "position" ⊆ "poem" too (same tables, filtered to one
  // label). Content and spacing never mix because every fan-out filters
  // strictly on `kind`, and (for "position") on the label/slot field that
  // kind actually carries — a content label ("A1") and a spacing slot
  // ("A#1") never collide.
  function keysForTarget(map, kind, mode, currentKey, tables) {
    if (mode === "this") return [currentKey];
    var tableIndex = parseInt(String(currentKey).split(":")[0], 10) || 0;
    if (mode === "bandh") {
      return (map || [])
        .filter(function (e) { return e.kind === kind; })
        .map(function (e) { return keyForEntry(tableIndex, e); });
    }
    if (mode === "poem") {
      var out = [];
      (tables || []).forEach(function (pattern, ti) {
        buildBandhCellMap(pattern)
          .filter(function (e) { return e.kind === kind; })
          .forEach(function (e) { out.push(keyForEntry(ti, e)); });
      });
      return out;
    }
    if (mode === "position") {
      // "Same cell in all bandhs": every table's cell whose label (content)
      // or slot (spacing) equals currentKey's, regardless of table index.
      var curLabel = String(currentKey).slice(String(currentKey).indexOf(":") + 1);
      var outP = [];
      (tables || []).forEach(function (pattern, ti) {
        buildBandhCellMap(pattern)
          .filter(function (e) {
            if (e.kind !== kind) return false;
            return (kind === "content" ? e.label : e.slot) === curLabel;
          })
          .forEach(function (e) { outP.push(keyForEntry(ti, e)); });
      });
      return outP;
    }
    return [];
  }

  return {
    buildBandhCellMap: buildBandhCellMap,
    alignPatternToTable: alignPatternToTable,
    keysForTarget: keysForTarget,
    columnGroupKey: columnGroupKey,
  };
}));

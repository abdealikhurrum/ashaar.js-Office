(function (root, factory) {
  if (typeof module !== "undefined" && module.exports) {
    module.exports = factory();
  } else {
    root.AshaarLayoutGrid = factory();
  }
}(typeof globalThis !== "undefined" ? globalThis : this, function () {
  var COLS = 12;

  // Group a row of booleans (reading order: index 0 = first misra = visual right)
  // into contiguous runs: [{ on, span }, …].
  function rowRuns(rowBools) {
    var runs = [];
    var i = 0;
    while (i < rowBools.length) {
      var v = rowBools[i];
      var start = i;
      while (i < rowBools.length && rowBools[i] === v) i++;
      runs.push({ on: v, span: i - start });
    }
    return runs;
  }

  // Convert the bubble matrix to the 12-column span template that templateToOoxml
  // consumes: { columnCount, rows: [ [ {span, align, role} … ] ] }.
  // A run of "on" bubbles = a misra cell; "off" = a gap cell. Among content cells
  // (reading order, right-to-left): first = right, last = left, middle/lone = center.
  // All-off rows are skipped.
  function gridToTemplate(matrix) {
    matrix = matrix || [];
    var columnCount = (matrix[0] && matrix[0].length) || COLS;
    var rows = [];
    matrix.forEach(function (rowBools) {
      var runs = rowRuns(rowBools);
      var contentCount = runs.filter(function (r) { return r.on; }).length;
      if (!contentCount) return; // blank row — skip
      var ci = 0;
      var cells = runs.map(function (r) {
        if (!r.on) return { span: r.span, align: "center", role: "gap" };
        var align;
        if (contentCount === 1) align = "center";
        else if (ci === 0) align = "right";
        else if (ci === contentCount - 1) align = "left";
        else align = "center";
        ci++;
        return { span: r.span, align: align, role: "misra" };
      });
      rows.push(cells);
    });
    return { columnCount: columnCount, rows: rows };
  }

  // Serialize the matrix to the nearest text layout-spec (Numbers view). Content
  // cells are numbered sequentially in reading order across rows. Best-effort:
  // 1 cell → "<n>", 2 → "n1 - n2", 3+ → "n1 | n2 | …".
  function gridToSpec(matrix) {
    matrix = matrix || [];
    var n = 1;
    var lines = [];
    matrix.forEach(function (rowBools) {
      var content = rowRuns(rowBools).filter(function (r) { return r.on; });
      if (!content.length) return;
      var nums = content.map(function () { return n++; });
      if (nums.length === 1) lines.push("<" + nums[0] + ">");
      else if (nums.length === 2) lines.push(nums[0] + " - " + nums[1]);
      else lines.push(nums.join(" | "));
    });
    return lines.join("\n");
  }

  function buildRow(runs) {
    var row = [];
    runs.forEach(function (r) {
      for (var i = 0; i < r.span; i++) row.push(r.on);
    });
    return row;
  }

  // Parse a coarse text layout-spec into a bubble matrix (best-effort), so the
  // Numbers→Grid toggle shows an existing layout. Supports pair ("a - b"),
  // centred ("<n>"), and multi ("a | b | c"); anything else → one centred cell.
  function specToGrid(text) {
    var lines = String(text == null ? "" : text).split(/\r\n|\r|\n/);
    var rows = [];
    lines.forEach(function (raw) {
      var line = raw.trim();
      if (!line) return;
      if (line.indexOf("|") >= 0) {
        var K = line.split("|").length;
        var gaps = K - 1;
        var contentTotal = COLS - gaps;
        var base = Math.floor(contentTotal / K);
        var extra = contentTotal - base * K;
        var runs = [];
        for (var i = 0; i < K; i++) {
          runs.push({ on: true, span: base + (i >= K - extra ? 1 : 0) });
          if (i < K - 1) runs.push({ on: false, span: 1 });
        }
        rows.push(buildRow(runs));
      } else if (line.indexOf(" - ") >= 0) {
        rows.push(buildRow([{ on: true, span: 5 }, { on: false, span: 2 }, { on: true, span: 5 }]));
      } else {
        rows.push(buildRow([{ on: false, span: 3 }, { on: true, span: 6 }, { on: false, span: 3 }]));
      }
    });
    return rows;
  }

  return {
    gridToTemplate: gridToTemplate,
    gridToSpec: gridToSpec,
    specToGrid: specToGrid,
    COLS: COLS
  };
}));

(function (root, factory) {
  if (typeof module !== "undefined" && module.exports) {
    module.exports = factory(require("../vendor/ashaar-justify"));
  } else {
    root.AshaarWord = factory(root.AshaarJustify);
  }
}(typeof globalThis !== "undefined" ? globalThis : this, function (AshaarJustify) {
  function escapeHtml(value) {
    return String(value || "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  }

  function cssLengthPercent(value, fallback) {
    var n = Number(value);
    return Number.isFinite(n) ? Math.max(0, Math.min(20, n)) + "%" : fallback;
  }

  function clamp(n, min, max) {
    return Math.min(max, Math.max(min, n));
  }

  function visibleWeight(text) {
    var cleaned = String(text || "").replace(/[ـ\s،؛,.!?؟]/g, "");
    var weight = 0;
    for (var i = 0; i < cleaned.length; i++) {
      var cp = cleaned.charCodeAt(i);
      weight += cp >= 0x0600 && cp <= 0x06FF ? 1 : 0.72;
    }
    return Math.max(1, weight);
  }

  function justifyText(text, opts) {
    opts = opts || {};
    var count = Number(opts.tatweelCount || 0);
    if (!AshaarJustify || opts.justifyMode !== "kashida" || count <= 0) return text;
    return AshaarJustify.spreadTatweels(text, count);
  }

  function parsePoetry(text, Ashaar) {
    if (!Ashaar || typeof Ashaar.parse !== "function") {
      throw new Error("Ashaar.js is not loaded.");
    }
    return Ashaar.parse(text || "");
  }

  function tableColumns(stanza, opts) {
    opts = opts || {};
    var layout = opts.layoutMode || "balanced";
    var gap = clamp(Number(opts.gapWidth || 4), 2, layout === "compact" ? 6 : 10);

    if (layout === "compact") {
      return { outer: 8, sadr: 40, gap: gap, ajuz: 40, mode: "compact" };
    }

    var maxMisraCount = 0;
    stanza.bayts.forEach(function (b) {
      if (b.type === "row" && b.misras) maxMisraCount = Math.max(maxMisraCount, b.misras.length);
    });
    if (maxMisraCount >= 3) {
      var colWidth = 100 / maxMisraCount;
      var widths = [];
      for (var i = 0; i < maxMisraCount; i++) widths.push(colWidth);
      return { mode: "normalized", widths: widths, count: maxMisraCount };
    }

    if (layout === "equal" || opts.widthMode === "fixed") {
      var equalText = (100 - gap) / 2;
      return { sadr: equalText, gap: gap, ajuz: equalText, mode: "three" };
    }

    var maxSadr = 1;
    var maxAjuz = 1;
    stanza.bayts.forEach(function (bayt) {
      if (bayt.type === "row" && bayt.misras && bayt.misras.length >= 2) {
        maxSadr = Math.max(maxSadr, visibleWeight(bayt.misras[0].text));
        maxAjuz = Math.max(maxAjuz, visibleWeight(bayt.misras[bayt.misras.length - 1].text));
        return;
      }
      if (!bayt.ajuz) return;
      maxSadr = Math.max(maxSadr, visibleWeight(bayt.sadr));
      maxAjuz = Math.max(maxAjuz, visibleWeight(bayt.ajuz));
    });

    var available = 100 - gap;
    var sadr = available * maxSadr / (maxSadr + maxAjuz);
    sadr = clamp(sadr, 36, 56);
    return { sadr: sadr, gap: gap, ajuz: available - sadr, mode: "three" };
  }

  function fontFamilyStyle(opts) {
    opts = opts || {};
    if (opts.fontMode === "nastaliq") return "font-family:'Noto Nastaliq Urdu','Jameel Noori Nastaleeq',serif";
    if (opts.fontMode === "arabic-serif") return "font-family:'Scheherazade New','Amiri','Times New Roman',serif";
    return "";
  }

  function tableStyle(opts) {
    var font = fontFamilyStyle(opts);
    opts = opts || {};
    return [
      "width:100%",
      "table-layout:fixed",
      "border-collapse:collapse",
      "direction:rtl",
      font,
      "font-size:16pt",
      "line-height:2.1",
      "margin:0 0 14pt 0"
    ].filter(Boolean).join(";");
  }

  function widthStyle(width) {
    return "width:" + Math.round(width * 100) / 100 + "%";
  }

  function colGroup(cols) {
    if (cols.mode === "normalized") {
      return "<colgroup>" + cols.widths.map(function (width) {
        return "<col style=\"" + widthStyle(width) + "\">";
      }).join("") + "</colgroup>";
    }
    if (cols.mode === "four") {
      return "<colgroup>" +
        "<col style=\"" + widthStyle(cols.c4) + "\">" +
        "<col style=\"" + widthStyle(cols.c3) + "\">" +
        "<col style=\"" + widthStyle(cols.c2) + "\">" +
        "<col style=\"" + widthStyle(cols.c1) + "\">" +
        "</colgroup>";
    }
    if (cols.mode === "compact") {
      return "<colgroup>" +
        "<col style=\"" + widthStyle(cols.outer) + "\">" +
        "<col style=\"" + widthStyle(cols.sadr) + "\">" +
        "<col style=\"" + widthStyle(cols.gap) + "\">" +
        "<col style=\"" + widthStyle(cols.ajuz) + "\">" +
        "<col style=\"" + widthStyle(cols.outer) + "\">" +
        "</colgroup>";
    }
    return "<colgroup>" +
      "<col style=\"" + widthStyle(cols.ajuz) + "\">" +
      "<col style=\"" + widthStyle(cols.gap) + "\">" +
      "<col style=\"" + widthStyle(cols.sadr) + "\">" +
      "</colgroup>";
  }

  function cellStyle(align, opts) {
    opts = opts || {};
    var justify = opts.justifyMode === "css" ? "text-align:justify;text-justify:inter-word;" : "text-align:" + align + ";";
    return [
      "vertical-align:baseline",
      "padding:0 0 3pt 0",
      justify,
      "direction:rtl"
    ].join(";");
  }

  function rightSideCellStyle(opts) {
    return cellStyle("left", opts);
  }

  function leftSideCellStyle(opts) {
    return cellStyle("right", opts);
  }

  function gapStyle() {
    return [
      "padding:0 5pt",
      "text-align:center",
      "vertical-align:baseline"
    ].join(";");
  }

  function blankMisraLabel(n) {
    return String(n || "\u00a0");
  }

  function misraCellStyle(spec, opts) {
    var align = spec.align || "center";
    var pieces = [
      "vertical-align:baseline",
      "padding:0 0 3pt 0",
      "text-align:" + align,
      "direction:rtl"
    ];
    if (spec.width) pieces.push("width:" + cssLengthPercent(spec.width, "auto"));
    if (spec.indent) pieces.push("padding-right:" + Number(spec.indent) + "pt");
    return pieces.join(";");
  }

  function renderGridRow(row, opts, cols) {
    cols = cols || {};
    if (row.type === "triple") {
      return "<tr>" +
        "<td style=\"padding:0\"></td>" +
        "<td style=\"" + misraCellStyle({ align: "right" }, opts) + "\">" + escapeHtml(blankMisraLabel(row.left)) + "</td>" +
        "<td style=\"" + misraCellStyle({ align: "center" }, opts) + "\">" + escapeHtml(blankMisraLabel(row.middle)) + "</td>" +
        "<td style=\"" + misraCellStyle({ align: "left" }, opts) + "\">" + escapeHtml(blankMisraLabel(row.right)) + "</td>" +
        "</tr>";
    }
    if (row.type === "multi") {
      var cells = row.cells.slice(0, 4);
      while (cells.length < 4) cells.unshift("\u00a0");
      return "<tr>" + cells.map(function (cell, index) {
        var align = index === cells.length - 1 ? "left" : index === 0 ? "right" : "center";
        return "<td style=\"" + misraCellStyle({ align: align }, opts) + "\">" + escapeHtml(blankMisraLabel(cell)) + "</td>";
      }).join("") + "</tr>";
    }
    if (row.type === "refrain") {
      return "<tr>" +
        "<td style=\"" + misraCellStyle({ align: "right" }, opts) + "\">" + escapeHtml(blankMisraLabel(row.left)) + "</td>" +
        "<td colspan=\"2\" style=\"" + gapStyle() + "\"></td>" +
        "<td style=\"" + misraCellStyle({ align: "left" }, opts) + "\">" + escapeHtml(blankMisraLabel(row.right)) + "</td>" +
        "</tr>";
    }
    if (row.type === "center") {
      return "<tr><td colspan=\"" + (row.colspan || 3) + "\" style=\"" + misraCellStyle(row, opts) + "\">" + escapeHtml(blankMisraLabel(row.misra)) + "</td></tr>";
    }
    if (row.type === "right") {
      if (cols.mode === "four") {
        return "<tr>" +
          "<td style=\"padding:0\"></td>" +
          "<td style=\"padding:0\"></td>" +
          "<td style=\"padding:0\"></td>" +
          "<td style=\"" + misraCellStyle(Object.assign({ align: "left" }, row), opts) + "\">" + escapeHtml(blankMisraLabel(row.misra)) + "</td>" +
          "</tr>";
      }
      return "<tr>" +
        "<td style=\"padding:0\"></td>" +
        "<td style=\"" + gapStyle() + "\"></td>" +
        "<td style=\"" + misraCellStyle(Object.assign({ align: "left" }, row), opts) + "\">" + escapeHtml(blankMisraLabel(row.misra)) + "</td>" +
        "</tr>";
    }
    if (row.type === "left") {
      if (cols.mode === "four") {
        return "<tr>" +
          "<td style=\"" + misraCellStyle(Object.assign({ align: "right" }, row), opts) + "\">" + escapeHtml(blankMisraLabel(row.misra)) + "</td>" +
          "<td style=\"padding:0\"></td>" +
          "<td style=\"padding:0\"></td>" +
          "<td style=\"padding:0\"></td>" +
          "</tr>";
      }
      return "<tr>" +
        "<td style=\"" + misraCellStyle(Object.assign({ align: "right" }, row), opts) + "\">" + escapeHtml(blankMisraLabel(row.misra)) + "</td>" +
        "<td style=\"" + gapStyle() + "\"></td>" +
        "<td style=\"padding:0\"></td>" +
        "</tr>";
    }
    if (cols.mode === "four") {
      return "<tr>" +
        "<td style=\"" + misraCellStyle({ align: "right" }, opts) + "\">" + escapeHtml(blankMisraLabel(row.left)) + "</td>" +
        "<td colspan=\"2\" style=\"" + gapStyle() + "\"></td>" +
        "<td style=\"" + misraCellStyle({ align: "left" }, opts) + "\">" + escapeHtml(blankMisraLabel(row.right)) + "</td>" +
        "</tr>";
    }
    return "<tr>" +
      "<td style=\"" + misraCellStyle({ align: "right" }, opts) + "\">" + escapeHtml(blankMisraLabel(row.left)) + "</td>" +
      "<td style=\"" + gapStyle() + "\"></td>" +
      "<td style=\"" + misraCellStyle({ align: "left" }, opts) + "\">" + escapeHtml(blankMisraLabel(row.right)) + "</td>" +
      "</tr>";
  }

  function replaceMisraRefs(row, misras) {
    function resolve(value) {
      var index = Number(value);
      if (Number.isFinite(index) && index > 0) return misras[index - 1] || "\u00a0";
      return value;
    }

    if (row.type === "multi") return Object.assign({}, row, { cells: row.cells.map(resolve) });
    if (row.type === "pair") return Object.assign({}, row, { left: resolve(row.left), right: resolve(row.right) });
    if (row.type === "center") return Object.assign({}, row, { misra: resolve(row.misra) });
    if (row.type === "right" || row.type === "left") return Object.assign({}, row, { misra: resolve(row.misra) });
    if (row.type === "triple") return Object.assign({}, row, {
      right: resolve(row.right),
      middle: resolve(row.middle),
      left: resolve(row.left)
    });
    if (row.type === "refrain") return Object.assign({}, row, { right: resolve(row.right), left: resolve(row.left) });
    return row;
  }

  function alignmentForColumn(index, columnCount) {
    if (columnCount <= 1) return "center";
    if (index === 0) return "right";
    if (index === columnCount - 1) return "left";
    return "center";
  }

  function rowColumnCount(row) {
    if (row.type === "multi") return row.cells.length;
    if (row.type === "triple") return 3;
    if (row.type === "pair" || row.type === "refrain") return 3;
    return 1;
  }

  function layoutColumnCount(rows) {
    var count = rows.reduce(function (max, row) {
      return Math.max(max, rowColumnCount(row));
    }, 1);
    return clamp(count, 1, 12);
  }

  function emptyCells(columnCount) {
    var cells = [];
    for (var i = 0; i < columnCount; i++) {
      cells.push({ text: "", align: alignmentForColumn(i, columnCount), role: "empty" });
    }
    return cells;
  }

  function normalizeLayoutRow(row, columnCount) {
    var cells = emptyCells(columnCount);
    var center = Math.floor(columnCount / 2);

    function put(index, text, align, role) {
      if (index < 0 || index >= columnCount) return;
      cells[index] = {
        text: String(text || ""),
        align: align || alignmentForColumn(index, columnCount),
        role: role || "misra"
      };
    }

    if (row.type === "multi") {
      var start = Math.max(0, columnCount - row.cells.length);
      row.cells.forEach(function (cell, index) {
        var target = start + index;
        put(target, cell, alignmentForColumn(target, columnCount), "misra");
      });
      return cells;
    }

    if (row.type === "triple") {
      put(0, row.left, "right", "left");
      put(center, row.middle, "center", "center");
      put(columnCount - 1, row.right, "left", "right");
      return cells;
    }

    if (row.type === "pair" || row.type === "refrain") {
      put(0, row.left, "right", "left");
      put(columnCount - 1, row.right, "left", "right");
      return cells;
    }

    if (row.type === "right") {
      put(columnCount - 1, row.misra, "left", "right");
      return cells;
    }

    if (row.type === "left") {
      put(0, row.misra, "right", "left");
      return cells;
    }

    put(center, row.misra, row.align || "center", "center");
    return cells;
  }

  function normalizedLayoutRows(rows) {
    var columnCount = layoutColumnCount(rows);
    return {
      columnCount: columnCount,
      rows: rows.map(function (row) {
        return normalizeLayoutRow(row, columnCount);
      })
    };
  }

  function layoutColumnWidths(columnCount, opts) {
    if (columnCount <= 1) return [100];
    if (columnCount === 2) return [50, 50];
    var gap = clamp(Number((opts || {}).gapWidth || 4), 0, 20);
    var outer = (100 - gap) / 2;
    if (columnCount === 3) return [outer, gap, outer];
    var each = 100 / columnCount;
    var widths = [];
    for (var i = 0; i < columnCount; i++) widths.push(each);
    return widths;
  }

  function layoutTablesForText(text, opts) {
    opts = Object.assign({}, opts || {});
    var baseRows = templateGrid(opts);
    var bandhs = splitBandhs(text);
    if (!bandhs.length) return [];
    return bandhs.map(function (bandh) {
      var misras = extractMisras(bandh);
      var resolvedRows = baseRows.map(function (row) { return replaceMisraRefs(row, misras); });
      var table = normalizedLayoutRows(resolvedRows);
      table.widths = layoutColumnWidths(table.columnCount, opts);
      return table;
    });
  }

  function layoutTablesForTemplate(opts) {
    opts = Object.assign({}, opts || {});
    var bandhCount = clamp(Number(opts.bandhCount || 1), 1, 20);
    var tables = [];
    for (var i = 0; i < bandhCount; i++) {
      var table = normalizedLayoutRows(templateGrid(opts));
      table.widths = layoutColumnWidths(table.columnCount, opts);
      tables.push(table);
    }
    return tables;
  }

  function stackedMisras(misras, spanCount, opts) {
    return "<tr><td colspan=\"" + spanCount + "\" style=\"text-align:center;direction:rtl;padding:0 0 5pt 0;\">" +
      misras.map(function (m) {
        var mc = m.isRefrain ? "color:#a7352a;" : "";
        return "<span style=\"padding-right:32pt;" + mc + "\">" + escapeHtml(justifyText(m.text, opts)) + "</span>";
      }).join("<br>") + "</td></tr>";
  }

  function renderBaytRow(bayt, opts, cols) {
    var N = cols.mode === "normalized" && cols.count ? cols.count : 0;

    if (bayt.type === "row") {
      var misras = bayt.misras || [];
      var K = misras.length;
      if (K === 0) return "";

      if (!N) {
        var fallbackSpan = cols.mode === "compact" ? 5 : 3;
        return stackedMisras(misras, fallbackSpan, opts);
      }

      if (K === 1 || opts.layoutMode === "stacked") {
        return stackedMisras(misras, N, opts);
      }

      if (K === 2) {
        var sadr2 = escapeHtml(justifyText(misras[0].text, opts));
        var ajuz2 = escapeHtml(justifyText(misras[1].text, opts));
        var sadrColor2 = misras[0].isRefrain ? "color:#a7352a;" : "";
        var ajuzColor2 = misras[1].isRefrain ? "color:#a7352a;" : "";
        var gapSpan = N - 2;
        return "<tr>" +
          "<td style=\"" + leftSideCellStyle(opts) + ajuzColor2 + "\">" + ajuz2 + "</td>" +
          (gapSpan > 0 ? "<td colspan=\"" + gapSpan + "\" style=\"" + gapStyle() + "\"></td>" : "") +
          "<td style=\"" + rightSideCellStyle(opts) + sadrColor2 + "\">" + sadr2 + "</td>" +
          "</tr>";
      }

      if (K === N) {
        return "<tr>" + misras.slice().reverse().map(function (m, i) {
          var align = i === 0 ? "right" : i === N - 1 ? "left" : "center";
          var mc = m.isRefrain ? "color:#a7352a;" : "";
          return "<td style=\"" + misraCellStyle({ align: align }, opts) + mc + "\">" + escapeHtml(justifyText(m.text, opts)) + "</td>";
        }).join("") + "</tr>";
      }

      // K > 2 and K < N: rare mixed-count stanza — stacked fallback
      return stackedMisras(misras, N, opts);
    }

    var sadr = escapeHtml(justifyText(bayt.sadr, opts));
    var ajuz = bayt.ajuz ? escapeHtml(justifyText(bayt.ajuz, opts)) : "";
    var sadrColor = bayt.sadrRefrain ? "color:#a7352a;" : "";
    var ajuzColor = bayt.ajuzRefrain ? "color:#a7352a;" : "";

    if (!bayt.ajuz || opts.layoutMode === "stacked") {
      var second = bayt.ajuz ? "<br><span style=\"padding-right:32pt;" + ajuzColor + "\">" + ajuz + "</span>" : "";
      var spanCount = cols.mode === "compact" ? 5 : N || 3;
      return "<tr><td colspan=\"" + spanCount + "\" style=\"text-align:center;direction:rtl;padding:0 0 5pt 0;" + sadrColor + "\">" + sadr + second + "</td></tr>";
    }

    if (cols.mode === "compact") {
      return "<tr>" +
        "<td style=\"padding:0\"></td>" +
        "<td style=\"" + leftSideCellStyle(opts) + ajuzColor + "\">" + ajuz + "</td>" +
        "<td style=\"" + gapStyle() + "\"></td>" +
        "<td style=\"" + rightSideCellStyle(opts) + sadrColor + "\">" + sadr + "</td>" +
        "<td style=\"padding:0\"></td>" +
        "</tr>";
    }

    if (N) {
      var gapSpanBayt = N - 2;
      return "<tr>" +
        "<td style=\"" + leftSideCellStyle(opts) + ajuzColor + "\">" + ajuz + "</td>" +
        (gapSpanBayt > 0 ? "<td colspan=\"" + gapSpanBayt + "\" style=\"" + gapStyle() + "\"></td>" : "") +
        "<td style=\"" + rightSideCellStyle(opts) + sadrColor + "\">" + sadr + "</td>" +
        "</tr>";
    }

    return "<tr>" +
      "<td style=\"" + leftSideCellStyle(opts) + ajuzColor + "\">" + ajuz + "</td>" +
      "<td style=\"" + gapStyle() + "\"></td>" +
      "<td style=\"" + rightSideCellStyle(opts) + sadrColor + "\">" + sadr + "</td>" +
      "</tr>";
  }

  function renderForWord(text, opts, Ashaar) {
    opts = opts || {};
    var poems = parsePoetry(String(text || ""), Ashaar);
    if (!poems.length) return "";
    return poems.map(function (poem) {
      return poem.stanzas.map(function (stanza) {
        var cols = tableColumns(stanza, opts);
        return "<table dir=\"rtl\" data-ashaar-layout=\"" + escapeHtml(opts.layoutMode || "balanced") + "\" style=\"" + tableStyle(opts) + "\">" +
          colGroup(cols) +
          "<tbody>" +
          stanza.bayts.map(function (bayt) { return renderBaytRow(bayt, opts, cols); }).join("") +
          "</tbody></table>";
      }).join("<p style=\"margin:10pt 0\"></p>");
    }).join("<p style=\"margin:18pt 0\"></p>");
  }

  function splitBandhs(text) {
    return String(text || "").split(/\n\s*(?:---|—|–)\s*\n/g).map(function (chunk) {
      return chunk.trim();
    }).filter(Boolean);
  }

  function extractMisras(text) {
    var misras = [];
    String(text || "").split(/\r?\n/).forEach(function (raw) {
      var line = raw.trim();
      if (!line || /^(?:---|—|–)$/.test(line)) return;
      line.split(/\s*[\\*|]\s*/).forEach(function (part) {
        var clean = part.replace(/\s*%\s*$/, "").trim();
        if (clean) misras.push(clean);
      });
    });
    return misras;
  }

  function renderTextWithLayoutForWord(text, opts) {
    opts = Object.assign({}, opts || {});
    var tables = layoutTablesForText(text, opts);
    if (!tables.length) return "";

    return tables.map(function (table) {
      var cols = { mode: "normalized", widths: table.widths };
      return "<table dir=\"rtl\" data-ashaar-layout=\"custom-spec\" data-ashaar-template=\"false\" style=\"" + tableStyle(opts) + "\">" +
        colGroup(cols) +
        "<tbody>" +
        table.rows.map(function (row) {
          return "<tr>" + row.map(function (cell) {
            return "<td style=\"" + misraCellStyle({ align: cell.align }, opts) + "\">" + escapeHtml(cell.text || "\u00a0") + "</td>";
          }).join("") + "</tr>";
        }).join("") +
        "</tbody></table>";
    }).join("<p style=\"margin:10pt 0\"></p>");
  }

  function templateGrid(opts) {
    opts = opts || {};
    var count = clamp(Number(opts.misraCount || 4), 1, 80);
    if (opts.layoutSpec && String(opts.layoutSpec).trim()) {
      return parseLayoutSpec(opts.layoutSpec);
    }
    var pattern = opts.misraPattern || "paired";
    if (pattern === "multi-misra-row" || pattern === "karbala-refrain") {
      pattern = "three-plus-center-refrain";
    }
    var rows = [];
    var i;

    if (pattern === "centered-stack") {
      for (i = 1; i <= count; i++) rows.push({ type: "center", misra: i, align: "center" });
      return rows;
    }

    if (pattern === "alternate-right" || pattern === "alternate-left") {
      var startsRight = pattern === "alternate-right";
      for (i = 1; i <= count; i++) {
        var isRight = (i % 2 === 1) === startsRight;
        rows.push({ type: isRight ? "right" : "left", misra: i });
      }
      return rows;
    }

    if (pattern === "indented-stack") {
      for (i = 1; i <= count; i++) rows.push({ type: "center", misra: i, align: "right", indent: (count - i) * 12 });
      return rows;
    }

    if (pattern === "four-plus-centered") {
      rows.push({ type: "pair", right: 1, left: 2 });
      if (count >= 4) rows.push({ type: "pair", right: 3, left: 4 });
      for (i = 5; i <= count; i++) rows.push({ type: "center", misra: i, align: "center" });
      return rows;
    }

    if (pattern === "three-plus-center-refrain") {
      rows.push({ type: "triple", right: 1, middle: 2, left: 3 });
      if (count >= 4) rows.push({ type: "center", misra: 4, align: "center", colspan: 4 });
      if (count >= 6) rows.push({ type: "refrain", right: 5, left: 6 });
      for (i = 7; i <= count; i++) rows.push({ type: "center", misra: i, align: "center", colspan: 4 });
      return rows;
    }

    for (i = 1; i <= count; i += 2) {
      rows.push({ type: "pair", right: i, left: i + 1 <= count ? i + 1 : "\u00a0" });
    }
    return rows;
  }

  function misraToken(value) {
    var token = String(value || "").trim();
    return token || "\u00a0";
  }

  function parseLayoutSpec(spec) {
    return String(spec || "").split(/\r?\n/).map(function (raw) {
      var line = raw.trim();
      if (!line) return null;

      var centered = line.match(/^<\s*([^>]+?)\s*>$/);
      if (centered) return { type: "center", misra: misraToken(centered[1]), align: "center", colspan: 4 };

      if (/\|/.test(line)) {
        var parts = line.split("|").map(misraToken);
        return { type: "multi", cells: parts };
      }

      var pair = line.split(/\s*-\s*/);
      if (pair.length === 2) return { type: "pair", left: misraToken(pair[0]), right: misraToken(pair[1]) };

      if (/>\s*$/.test(line)) return { type: "right", misra: misraToken(line.replace(/>\s*$/, "")) };
      if (/^<\s*/.test(line)) return { type: "left", misra: misraToken(line.replace(/^<\s*/, "")) };

      var indent = raw.match(/^\s*/)[0].length;
      return { type: "center", misra: misraToken(line), align: indent ? "right" : "center", indent: indent * 4, colspan: 4 };
    }).filter(Boolean);
  }

  function templateColumns(opts) {
    var gap = clamp(Number((opts || {}).gapWidth || 4), 2, 10);
    var pattern = (opts || {}).misraPattern;
    var spec = (opts || {}).layoutSpec || "";
    if (/\|/.test(spec) || pattern === "three-plus-center-refrain" || pattern === "multi-misra-row" || pattern === "karbala-refrain") {
      return { c1: 30, c2: 30, c3: 30, c4: 10, mode: "four" };
    }
    var side = (100 - gap) / 2;
    return { sadr: side, gap: gap, ajuz: side, mode: "three" };
  }

  function renderTemplateForWord(opts) {
    opts = Object.assign({}, opts || {});
    var layout = opts.layoutMode || "equal";
    opts.layoutMode = layout;
    opts.widthMode = opts.widthMode || "fixed";

    return layoutTablesForTemplate(opts).map(function (table) {
      var cols = { mode: "normalized", widths: table.widths };
      return "<table dir=\"rtl\" data-ashaar-layout=\"" + escapeHtml(layout) + "\" data-ashaar-template=\"true\" style=\"" + tableStyle(opts) + "\">" +
        colGroup(cols) +
        "<tbody>" +
        table.rows.map(function (row) {
          return "<tr>" + row.map(function (cell) {
            return "<td style=\"" + misraCellStyle({ align: cell.align }, opts) + "\">" + escapeHtml(cell.text || "\u00a0") + "</td>";
          }).join("") + "</tr>";
        }).join("") +
        "</tbody></table>";
    }).join("<p style=\"margin:10pt 0\"></p>");
  }

  function contentControlTag(text, opts) {
    opts = opts || {};
    var source = String(text || "");
    var hash = 0;
    for (var i = 0; i < source.length; i++) {
      hash = ((hash << 5) - hash + source.charCodeAt(i)) | 0;
    }
    var payload = {
      k: "ashaar-poem",
      v: 1,
      layoutMode: opts.layoutMode || "balanced",
      widthMode: opts.widthMode || "optimized",
      justifyMode: opts.justifyMode || "kashida",
      tatweelCount: Number(opts.tatweelCount || 0),
      gapWidth: Number(opts.gapWidth || 4),
      misraPattern: opts.misraPattern || "paired",
      misraCount: Number(opts.misraCount || 4),
      fontMode: opts.fontMode || "document",
      sourceHash: (hash >>> 0).toString(16)
    };
    return "ashaar:" + encodeURIComponent(JSON.stringify(payload));
  }

  function justifyPlainTextBlock(text, opts) {
    return String(text || "").split(/\r?\n/).map(function (line) {
      return line.trim() ? justifyText(line, opts) : line;
    }).join("\n");
  }

  return {
    renderForWord: renderForWord,
    renderTextWithLayoutForWord: renderTextWithLayoutForWord,
    renderTemplateForWord: renderTemplateForWord,
    layoutTablesForText: layoutTablesForText,
    layoutTablesForTemplate: layoutTablesForTemplate,
    extractMisras: extractMisras,
    templateGrid: templateGrid,
    templateColumns: templateColumns,
    parseLayoutSpec: parseLayoutSpec,
    tableColumns: tableColumns,
    contentControlTag: contentControlTag,
    justifyPlainTextBlock: justifyPlainTextBlock
  };
}));

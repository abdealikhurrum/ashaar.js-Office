(function (root, factory) {
  if (typeof module !== "undefined" && module.exports) {
    module.exports = factory(require("../vendor/ashaar-justify"), require("./fonts"));
  } else {
    root.AshaarWord = factory(root.AshaarJustify, root.AshaarFonts);
  }
}(typeof globalThis !== "undefined" ? globalThis : this, function (AshaarJustify, AshaarFonts) {
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
      // Skip Arabic combining diacritics (harakat U+064B–U+065F, superscript alef U+0670)
      if ((cp >= 0x064B && cp <= 0x065F) || cp === 0x0670) continue;
      weight += cp >= 0x0600 && cp <= 0x06FF ? 1 : 0.72;
    }
    return Math.max(1, weight);
  }

  // Distribute hair spaces (U+200A) between words to reach targetWidth * targetFill.
  // Falls back to thin spaces (U+2009) if the font returns zero width for hair space.
  // Uses the same binary-search pattern as AshaarJustify.justifyLine for tatweels.
  function justifyWordSpacing(text, targetWidth, ctx, targetFill) {
    if (!text || !ctx || targetWidth <= 0) return text;
    var target = targetWidth * (targetFill || 0.92);
    var natural = ctx.measureText(String(text)).width;
    if (natural >= target) return text;
    var words = String(text).split(" ");
    var gaps = words.length - 1;
    if (gaps <= 0) return text;

    var SPACE = " "; // hair space — ~1/24 em
    var spaceW = ctx.measureText(SPACE).width;
    if (spaceW <= 0) {
      SPACE = " "; // thin space — ~1/5 em
      spaceW = ctx.measureText(SPACE).width;
    }
    if (spaceW <= 0) return text;

    var maxN = Math.min(Math.floor((target - natural) / spaceW), 1000);

    function applyN(n) {
      if (n <= 0) return text;
      var base = Math.floor(n / gaps);
      var extra = n % gaps;
      var out = "";
      for (var i = 0; i < words.length; i++) {
        if (i > 0) {
          out += " ";
          var add = base + (i - 1 < extra ? 1 : 0);
          for (var j = 0; j < add; j++) out += SPACE;
        }
        out += words[i];
      }
      return out;
    }

    var lo = 0, hi = maxN;
    while (lo < hi) {
      var mid = Math.ceil((lo + hi) / 2);
      if (ctx.measureText(applyN(mid)).width <= target) lo = mid;
      else hi = mid - 1;
    }
    return applyN(lo);
  }

  // Map tatweelCount slider (1–10) to targetFill:
  //   1  → 0.90
  //   10 → 1.0 (full fill, "living on the edge")
  function sliderToFill(count) {
    var s = Math.max(1, Math.min(10, Number(count) || 1));
    return 0.90 + ((s - 1) / 9) * 0.10;
  }

  // Poetry Kashida strength (1–10) → elongation share φ ∈ [0,1]: the fraction of a
  // line's fill-gap closed by the font's elongation mechanism (the rest by
  // spacing). s=1 → 0 (all spacing); s=10 → 1 (all elongation, minor spacing).
  function strengthToElongationShare(strength) {
    var s = Number(strength);
    if (!isFinite(s)) s = 1;
    s = Math.max(1, Math.min(10, s));
    return (s - 1) / 9;
  }

  // Poetry Kashida strength (1–10) → max elongation positions K(s) for the
  // concentrated generic path. Low strengths stay deliberately sparse:
  // K(1)=1, K(2)=2, K(3)=3; s>=4 is unbounded (0, the engine's "no cap").
  function strengthToMaxPositions(strength) {
    var s = Number(strength);
    if (!isFinite(s)) s = 1;
    s = Math.max(1, Math.min(10, s));
    if (s <= 3) return s;
    return 0;
  }

  function justifyText(text, opts, colWidthPx) {
    opts = opts || {};
    var count = Number(opts.tatweelCount || 0);
    if (opts._justifyCtx && colWidthPx > 0) {
      if (AshaarJustify && opts.justifyMode === "kashida") {
        if (count === 0) return text; // slider at 0 = off
        var params = { targetFill: sliderToFill(count) };
        if (opts._fontProfile) params.fontQualityBoost = 1.8;
        return AshaarJustify.justifyLine(text, colWidthPx, opts._justifyCtx, params, opts._fontProfile || null);
      }
      if (opts.justifyMode === "spacing") {
        if (count === 0) return text;
        return justifyWordSpacing(text, colWidthPx, opts._justifyCtx, sliderToFill(count));
      }
    }
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

    if (layout === "compact") {
      var gap = clamp(Number(opts.gapWidth || 4), 2, 6);
      return { outer: 8, sadr: 40, gap: gap, ajuz: 40, mode: "compact" };
    }

    // 12-column fixed grid: every stanza uses 12 equal columns regardless of N.
    // Each misra gets floor(12/N) columns; any remainder is absorbed by the last misra.
    var N = 0;
    stanza.bayts.forEach(function (b) {
      if (b.type === "row" && b.misras) N = Math.max(N, b.misras.length);
      else if (b.ajuz) N = Math.max(N, 2);
      else N = Math.max(N, 1);
    });
    N = Math.max(N, 2); // minimum 2-column layout
    var cpm = Math.floor(12 / N); // cols per misra
    var rem = 12 - cpm * N;      // extra cols absorbed into last misra
    return { mode: "grid12", N: N, cpm: cpm, rem: rem };
  }

  function fontFamilyStyle(opts) {
    opts = opts || {};
    var mode = opts.fontMode === "nastaliq" ? "noto" : opts.fontMode; // legacy alias
    var css = AshaarFonts.cssFamilyOf(mode);
    return css ? "font-family:" + css : "";
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
    if (cols.mode === "grid12") {
      var cg = "<colgroup>";
      for (var i = 0; i < 12; i++) cg += '<col style="width:8.3333%">';
      return cg + "</colgroup>";
    }
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

  // Convert rows-with-indent to span-based cells for structural indentation.
  // Returns null when no row has indent (caller falls back to normalizedLayoutRows).
  function indentedRowsToSpans(rawRows) {
    var N = rawRows.length;
    if (N < 2) return null;
    var maxIndent = rawRows.reduce(function (m, r) { return Math.max(m, r.indent || 0); }, 0);
    if (!maxIndent) return null;
    return rawRows.map(function (row) {
      // GRID = 2N: content fixed at N cols, rightSpan + leftSpan = N so both sides are adjustable.
      var rightSpan = Math.round((row.indent || 0) / maxIndent * (N - 1));
      var leftSpan = N - rightSpan;
      var cells = [];
      if (rightSpan > 0) cells.push({ span: rightSpan, text: " ", align: "center" });
      cells.push({ span: N, text: blankMisraLabel(String(row.misra || " ")), align: "right" });
      if (leftSpan > 0) cells.push({ span: leftSpan, text: " ", align: "center" });
      return cells;
    });
  }

  function layoutTablesForTemplate(opts) {
    opts = Object.assign({}, opts || {});
    var bandhCount = clamp(Number(opts.bandhCount || 1), 1, 20);
    var tables = [];
    for (var i = 0; i < bandhCount; i++) {
      var rawRows = templateGrid(opts);
      var spanRows = indentedRowsToSpans(rawRows);
      if (spanRows) {
        tables.push({ columnCount: 2 * rawRows.length, rows: spanRows, spanBased: true });
      } else {
        var table = normalizedLayoutRows(rawRows);
        table.widths = layoutColumnWidths(table.columnCount, opts);
        table.rows = table.rows.map(function (row) {
          return row.map(function (cell) {
            return Object.assign({}, cell, { text: blankMisraLabel(cell.text) });
          });
        });
        tables.push(table);
      }
    }
    return tables;
  }

  function stackedMisras(misras, spanCount, opts, colWidthPx) {
    return "<tr><td colspan=\"" + spanCount + "\" style=\"text-align:center;direction:rtl;padding:0 0 5pt 0;\">" +
      misras.map(function (m) {
        var mc = m.isRefrain ? "color:#a7352a;" : "";
        return "<span style=\"padding-right:32pt;" + mc + "\">" + escapeHtml(justifyText(m.text, opts, colWidthPx)) + "</span>";
      }).join("<br>") + "</td></tr>";
  }

  function renderBaytRowGrid12(bayt, opts, cols) {
    var textWidthPx = (opts || {})._textWidthPx || 0;
    var N = cols.N, cpm = cols.cpm, rem = cols.rem;
    var halfPx = textWidthPx / 2;
    var colPx = cpm / 12 * textWidthPx;

    function pairRow(sadrText, ajuzText, sadrColor, ajuzColor) {
      return "<tr>" +
        "<td colspan=\"6\" style=\"" + leftSideCellStyle(opts) + sadrColor + "\">" + sadrText + "</td>" +
        "<td colspan=\"" + (6 + rem) + "\" style=\"" + rightSideCellStyle(opts) + ajuzColor + "\">" + ajuzText + "</td>" +
        "</tr>";
    }

    if (bayt.type === "row") {
      var misras = bayt.misras || [];
      var K = misras.length;
      if (K === 0) return "";

      if (K === 1 || opts.layoutMode === "stacked") {
        return stackedMisras(misras, 12, opts, textWidthPx);
      }

      if (K === 2) {
        return pairRow(
          escapeHtml(justifyText(misras[0].text, opts, halfPx)),
          escapeHtml(justifyText(misras[1].text, opts, halfPx)),
          misras[0].isRefrain ? "color:#a7352a;" : "",
          misras[1].isRefrain ? "color:#a7352a;" : ""
        );
      }

      if (K >= N) {
        return "<tr>" + misras.slice(0, N).map(function (m, i) {
          var align = i === 0 ? "right" : i === N - 1 ? "left" : "center";
          var mc = m.isRefrain ? "color:#a7352a;" : "";
          var span = (i === N - 1) ? cpm + rem : cpm;
          return "<td colspan=\"" + span + "\" style=\"" + misraCellStyle({ align: align }, opts) + mc + "\">" +
            escapeHtml(justifyText(m.text, opts, colPx)) + "</td>";
        }).join("") + "</tr>";
      }

      // K between 2 and N — stacked fallback
      return stackedMisras(misras, 12, opts, textWidthPx);
    }

    // Old-format bayt (sadr / ajuz)
    var sadr = escapeHtml(justifyText(bayt.sadr, opts, halfPx));
    var sadrColor = bayt.sadrRefrain ? "color:#a7352a;" : "";

    if (!bayt.ajuz || opts.layoutMode === "stacked") {
      var ajuzInline = "";
      if (bayt.ajuz) {
        var ajuzColor0 = bayt.ajuzRefrain ? "color:#a7352a;" : "";
        var ajuzEsc = escapeHtml(justifyText(bayt.ajuz, opts, halfPx));
        ajuzInline = "<br><span style=\"padding-right:32pt;" + ajuzColor0 + "\">" + ajuzEsc + "</span>";
      }
      return "<tr><td colspan=\"12\" style=\"text-align:center;direction:rtl;padding:0 0 5pt 0;" + sadrColor + "\">" + sadr + ajuzInline + "</td></tr>";
    }

    return pairRow(
      sadr,
      escapeHtml(justifyText(bayt.ajuz, opts, halfPx)),
      sadrColor,
      bayt.ajuzRefrain ? "color:#a7352a;" : ""
    );
  }

  function renderBaytRow(bayt, opts, cols) {
    if (cols.mode === "grid12") {
      return renderBaytRowGrid12(bayt, opts, cols);
    }

    // Non-grid12 modes: compact layout and template rendering
    var textWidthPx = (opts || {})._textWidthPx || 0;
    var sadr = escapeHtml(justifyText(bayt.sadr || "", opts, 0));
    var ajuz = bayt.ajuz ? escapeHtml(justifyText(bayt.ajuz, opts, 0)) : "";
    var sadrColor = bayt.sadrRefrain ? "color:#a7352a;" : "";
    var ajuzColor = bayt.ajuzRefrain ? "color:#a7352a;" : "";

    if (bayt.type === "row") {
      var misras = bayt.misras || [];
      var K = misras.length;
      if (K === 0) return "";
      if (K === 1 || opts.layoutMode === "stacked") {
        return stackedMisras(misras, cols.mode === "compact" ? 5 : 3, opts, textWidthPx);
      }
      // Compact 2-misra
      if (cols.mode === "compact") {
        var s2 = escapeHtml(justifyText(misras[0].text, opts, 0));
        var a2 = escapeHtml(justifyText(misras[1].text, opts, 0));
        var sc2 = misras[0].isRefrain ? "color:#a7352a;" : "";
        var ac2 = misras[1].isRefrain ? "color:#a7352a;" : "";
        return "<tr>" +
          "<td style=\"padding:0\"></td>" +
          "<td style=\"" + leftSideCellStyle(opts) + sc2 + "\">" + s2 + "</td>" +
          "<td style=\"" + gapStyle() + "\"></td>" +
          "<td style=\"" + rightSideCellStyle(opts) + ac2 + "\">" + a2 + "</td>" +
          "<td style=\"padding:0\"></td>" +
          "</tr>";
      }
      return stackedMisras(misras, 3, opts, textWidthPx);
    }

    if (!bayt.ajuz || opts.layoutMode === "stacked") {
      var second = bayt.ajuz ? "<br><span style=\"padding-right:32pt;" + ajuzColor + "\">" + ajuz + "</span>" : "";
      var spanCount = cols.mode === "compact" ? 5 : 3;
      return "<tr><td colspan=\"" + spanCount + "\" style=\"text-align:center;direction:rtl;padding:0 0 5pt 0;" + sadrColor + "\">" + sadr + second + "</td></tr>";
    }

    if (cols.mode === "compact") {
      return "<tr>" +
        "<td style=\"padding:0\"></td>" +
        "<td style=\"" + leftSideCellStyle(opts) + sadrColor + "\">" + sadr + "</td>" +
        "<td style=\"" + gapStyle() + "\"></td>" +
        "<td style=\"" + rightSideCellStyle(opts) + ajuzColor + "\">" + ajuz + "</td>" +
        "<td style=\"padding:0\"></td>" +
        "</tr>";
    }

    return "<tr>" +
      "<td style=\"" + leftSideCellStyle(opts) + sadrColor + "\">" + sadr + "</td>" +
      "<td style=\"" + gapStyle() + "\"></td>" +
      "<td style=\"" + rightSideCellStyle(opts) + ajuzColor + "\">" + ajuz + "</td>" +
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

  function layoutTablesForPoem(text, opts, Ashaar) {
    opts = opts || {};
    var poems = parsePoetry(String(text || ""), Ashaar);
    if (!poems.length) return null;

    var hasMultiMisra = false;
    var tables = [];

    poems.forEach(function (poem) {
      poem.stanzas.forEach(function (stanza) {
        var N = 0;
        stanza.bayts.forEach(function (b) {
          if (b.type === "row" && b.misras) N = Math.max(N, b.misras.length);
        });
        if (N >= 3) {
          hasMultiMisra = true;
        } else {
          N = 3;
        }

        // 2N-1 interleaved columns: content at even indices, gap at odd indices.
        // This gives each gap its own fixed-width column so maqta spacing
        // doesn't depend on middle-content column width.
        var M = 2 * N - 1;
        var gap = clamp(Number(opts.gapWidth || 4), 1, 10);
        var cw = (100 - (N - 1) * gap) / N;
        var widths = [];
        for (var i = 0; i < M; i++) widths.push(i % 2 === 0 ? cw : gap);

        // Center content column in M-col grid (always an even index)
        var centerMCol = 2 * Math.floor(N / 2);

        var rows = stanza.bayts.map(function (bayt) {
          var row = [];
          for (var i = 0; i < M; i++) row.push({ text: "", align: "center" });

          if (bayt.type === "row") {
            var K = bayt.misras ? bayt.misras.length : 0;
            if (K === 0) return row;
            if (K === 1) {
              row[centerMCol] = { text: bayt.misras[0].text, align: "center" };
            } else if (K >= N) {
              // Full N-misra row: misras[0] (sadr) in col 0 = visual right in RTL table
              for (var i = 0; i < N; i++) {
                var align = i === 0 ? "right" : i === N - 1 ? "left" : "center";
                row[2 * i] = { text: bayt.misras[i].text, align: align };
              }
            } else if (K === 2) {
              // Pair within multi-misra stanza: sadr in col 0 (visual right), ajuz in col M-1
              row[0] = { text: bayt.misras[0].text, align: "right" };
              row[M - 1] = { text: bayt.misras[1].text, align: "left" };
            } else {
              // Partial K-misra row (3 <= K < N): occupy leftmost K content cols (visual right in RTL)
              for (var j = 0; j < K; j++) {
                var align = j === 0 ? "right" : j === K - 1 ? "left" : "center";
                row[2 * j] = { text: bayt.misras[j].text, align: align };
              }
            }
          } else if (!bayt.ajuz) {
            row[centerMCol] = { text: bayt.sadr, align: "center" };
          } else {
            // sadr in col 0 (visual right in RTL table), ajuz in col M-1 (visual left)
            row[0] = { text: bayt.sadr, align: "right" };
            row[M - 1] = { text: bayt.ajuz, align: "left" };
          }

          return row;
        });

        tables.push({ columnCount: M, rows: rows, widths: widths });
      });
    });

    return hasMultiMisra ? tables : null;
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
      for (i = 1; i <= count; i++) rows.push({ type: "center", misra: i, align: "right", indent: (i - 1) * 12 });
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

  function contentControlTag(text, opts, cellPatterns) {
    opts = opts || {};
    var source = String(text || "");
    var hash = 0;
    for (var i = 0; i < source.length; i++) {
      hash = ((hash << 5) - hash + source.charCodeAt(i)) | 0;
    }
    var payload = {
      k: "ashaar-poem",
      v: 2,
      layoutMode: opts.layoutMode || "balanced",
      widthMode: opts.widthMode || "optimized",
      justifyMode: opts.justifyMode || "kashida",
      tatweelCount: Number(opts.tatweelCount || 0),
      gapWidth: Number(opts.gapWidth || 4),
      misraPattern: opts.misraPattern || "paired",
      misraCount: Number(opts.misraCount || 4),
      fontMode: opts.fontMode || "document",
      tableWidthPct: Number(opts.tableWidthPct || 100),
      qaseeda: opts.qaseeda || "",
      sourceHash: (hash >>> 0).toString(16)
    };
    if (cellPatterns && cellPatterns.length) payload.cells = cellPatterns;
    return "ashaar:" + encodeURIComponent(JSON.stringify(payload));
  }

  // Decode an "ashaar:" content-control tag back into its payload object.
  // Returns null for empty/non-ashaar/malformed tags. Guarantees a string
  // `qaseeda` field so callers can read it without a presence check.
  function parseContentControlTag(tag) {
    if (typeof tag !== "string" || tag.indexOf("ashaar:") !== 0) return null;
    try {
      var payload = JSON.parse(decodeURIComponent(tag.slice("ashaar:".length)));
      if (!payload || typeof payload !== "object") return null;
      if (typeof payload.qaseeda !== "string") payload.qaseeda = "";
      payload.cells = payload.cells || null;
      payload.overrides = (payload.overrides && typeof payload.overrides === "object") ? payload.overrides : {};
      payload.slotDecor = (payload.slotDecor && typeof payload.slotDecor === "object") ? payload.slotDecor : {};
      payload.runFonts = (payload.runFonts && typeof payload.runFonts === "object") ? payload.runFonts : null;
      payload.widthPt = (typeof payload.widthPt === "number" && payload.widthPt > 0) ? payload.widthPt : null;
      return payload;
    } catch (e) {
      return null;
    }
  }

  // Return a copy of an "ashaar:" tag with only its qaseeda name replaced.
  // Non-ashaar / malformed tags are returned unchanged.
  function setTagQaseeda(tag, name) {
    var payload = parseContentControlTag(tag);
    if (!payload) return tag;
    payload.qaseeda = name || "";
    return "ashaar:" + encodeURIComponent(JSON.stringify(payload));
  }

  // Return a copy of an "ashaar:" tag with one per-cell override set or removed.
  // A null/empty override deletes the key. Non-ashaar tags returned unchanged.
  function setTagOverride(tag, key, override) {
    var payload = parseContentControlTag(tag);
    if (!payload) return tag;
    var ov = payload.overrides && typeof payload.overrides === "object" ? payload.overrides : {};
    var has = override && typeof override === "object" &&
      (override.strength != null || override.widthPt != null || override.capEm != null);
    if (has) {
      var clean = {};
      if (override.strength != null) clean.strength = override.strength;
      if (override.widthPt != null) clean.widthPt = override.widthPt;
      if (override.capEm != null) clean.capEm = override.capEm;
      ov[key] = clean;
    } else {
      delete ov[key];
    }
    payload.overrides = ov;
    return "ashaar:" + encodeURIComponent(JSON.stringify(payload));
  }

  // Return a copy of an "ashaar:" tag with one per-slot decoration set or removed.
  // A null/all-empty decor deletes the key. Non-ashaar tags returned unchanged.
  function setTagSlotDecor(tag, key, decor) {
    var payload = parseContentControlTag(tag);
    if (!payload) return tag;
    var sd = payload.slotDecor && typeof payload.slotDecor === "object" ? payload.slotDecor : {};
    var clean = {};
    if (decor && typeof decor === "object") {
      if (decor.symbol) clean.symbol = decor.symbol;
      if (decor.fill) clean.fill = decor.fill;
      if (decor.color) clean.color = decor.color;
    }
    if (clean.symbol || clean.fill || clean.color) sd[key] = clean; else delete sd[key];
    payload.slotDecor = sd;
    return "ashaar:" + encodeURIComponent(JSON.stringify(payload));
  }

  // ── Per-word font persistence (content-control tag) ──────────────────────
  // Office.js Font.name reads the COMPLEX-SCRIPT face for Arabic runs and ""
  // when a word's runs carry mixed cs faces — which is exactly what the Jameel
  // font-swap mechanism produces (some fasls base, some Kasheeda). Re-reading
  // fonts from a justified cell is therefore structurally lossy. The tag holds
  // the authoritative per-word fonts: packRunWords stores them at apply time,
  // reconcileRunWords heals ambiguous ("") reads on the next capture.

  // Compact per-cell encoding: consecutive same-styled words collapse to
  // [wordCount, name, sizePt] — with flags (bold=1|italic=2) appended when any
  // style is set, and the hex color after that when present. Word-aligned,
  // order = document order. Styleless entries stay 3-tuples, so pre-style tags
  // and new tags share one format.
  function packRunWords(words) {
    var out = [];
    (words || []).forEach(function (w) {
      var flags = (w.bold ? 1 : 0) | (w.italic ? 2 : 0);
      var color = w.color || 0;
      var prev = out[out.length - 1];
      if (prev && prev[1] === w.name && prev[2] === w.size &&
          (prev[3] || 0) === flags && (prev[4] || 0) === color) prev[0]++;
      else {
        var e = [1, w.name, w.size];
        if (flags || color) { e.push(flags); if (color) e.push(color); }
        out.push(e);
      }
    });
    return out;
  }

  // Merge a cell's captured words with its tag entry, PER FIELD: each field
  // heals from the tag only when its document read is ambiguous (Office.js
  // returns null for any property that is mixed within the range):
  //   name  — raw ""/null (mixed cs faces, e.g. a partially-swapped word);
  //           a real raw name wins (user re-font, or a fully-swapped word
  //           cleanly reading the Kasheeda face — same family).
  //   size  — rawSize null/0 (mixed sizes within the word).
  //   bold/italic — null read (word partially bolded / italicised).
  //   color — rawColor null (mixed); rawColor "" means EXPLICITLY uncolored —
  //           a real state (user cleared it), never healed.
  // Whole-word user changes therefore stick; sub-word ambiguity heals to the
  // last locked state instead of silently coercing to defaults. Returns a NEW
  // words array (bold/italic coerced to booleans), or null when the pack is
  // absent or its word count no longer matches (text was edited — the tag
  // entry is stale, caller keeps the document reads).
  function reconcileRunWords(words, packed) {
    if (!packed || !packed.length || !words || !words.length) return null;
    var total = packed.reduce(function (a, p) { return a + (p[0] || 0); }, 0);
    if (total !== words.length) return null;
    var perWord = [];
    packed.forEach(function (p) {
      for (var k = 0; k < p[0]; k++) {
        perWord.push({ name: p[1], size: p[2], flags: p[3] || 0, color: p[4] || undefined });
      }
    });
    return words.map(function (w, i) {
      var t = perWord[i];
      var out = {};
      for (var key in w) if (w.hasOwnProperty(key)) out[key] = w[key];
      if (!w.raw) out.name = t.name || w.name;
      if (!(w.rawSize > 0) && t.size) out.size = t.size;
      if (w.bold == null) out.bold = !!(t.flags & 1);
      if (w.italic == null) out.italic = !!(t.flags & 2);
      if (w.rawColor === null && t.color) out.color = t.color;
      out.bold = !!out.bold;
      out.italic = !!out.italic;
      return out;
    });
  }

  // Return a copy of an "ashaar:" tag with the BANDH-level misra width (pt) set
  // or removed (null/0 deletes). Sits between the qaseeda profile's
  // justify.widthPt and per-cell overrides: cell > bandh > qaseeda > computed.
  function setTagBandhWidth(tag, widthPt) {
    var payload = parseContentControlTag(tag);
    if (!payload) return tag;
    if (typeof widthPt === "number" && widthPt > 0) payload.widthPt = widthPt;
    else delete payload.widthPt;
    return "ashaar:" + encodeURIComponent(JSON.stringify(payload));
  }

  // Return a copy of an "ashaar:" tag with its per-cell runFonts map replaced
  // ({"table:cell": packRunWords(...)}). Non-ashaar tags returned unchanged.
  function setTagRunFonts(tag, runFonts) {
    var payload = parseContentControlTag(tag);
    if (!payload) return tag;
    if (runFonts && typeof runFonts === "object" && Object.keys(runFonts).length) payload.runFonts = runFonts;
    else delete payload.runFonts;
    return "ashaar:" + encodeURIComponent(JSON.stringify(payload));
  }

  function justifyPlainTextBlock(text, opts, colWidthPx) {
    return String(text || "").split(/\r?\n/).map(function (line) {
      return line.trim() ? justifyText(line, opts, colWidthPx) : line;
    }).join("\n");
  }

  // ── Run-aware justification consumers ────────────────────────────────────
  // Merge adjacent words that share an identical style tuple into runs. A run's
  // text is its words joined by a single space. Word-aligned (exact per the
  // run-aware kashida foundation spec); mid-word style changes are not
  // represented here — the caller reads one style per whole word.
  function coalesceRuns(words) {
    var runs = [];
    (words || []).forEach(function (w) {
      var prev = runs[runs.length - 1];
      if (prev && prev.name === w.name && prev.size === w.size &&
          prev.bold === w.bold && prev.italic === w.italic && prev.color === w.color) {
        prev.text += " " + w.text;
        prev.refs.push(w);
      } else {
        runs.push({
          text: w.text, name: w.name, size: w.size, bold: w.bold, italic: w.italic, color: w.color,
          refs: [w] // source word objects in order — caller maps these back (e.g. to Word ranges)
        });
      }
    });
    return runs;
  }

  // Insert n micro-space glyphs across the intra-run word gaps, round-robin so
  // distribution stays even. Each glyph is placed at a gap that belongs to one
  // run, so it stays within that run's range on write-back. Inter-run gaps (the
  // split delimiters between runs) are outside every run's text and are left
  // unstretched — a documented limitation of the run-aware spacing path.
  function distributeMicroSpaces(runTexts, n, spaceChar) {
    var texts = (runTexts || []).map(String);
    if (n <= 0) return texts;
    var slots = [];
    texts.forEach(function (t, ri) {
      var gaps = t.split(" ").length - 1;
      for (var g = 0; g < gaps; g++) slots.push({ ri: ri, gap: g });
    });
    if (!slots.length) return texts;
    var counts = {}; // "ri:gap" -> extra glyphs
    for (var i = 0; i < n; i++) {
      var s = slots[i % slots.length];
      var key = s.ri + ":" + s.gap;
      counts[key] = (counts[key] || 0) + 1;
    }
    return texts.map(function (t, ri) {
      var parts = t.split(" ");
      var out = parts[0] || "";
      for (var g = 0; g < parts.length - 1; g++) {
        var extra = counts[ri + ":" + g] || 0;
        out += " " + new Array(extra + 1).join(spaceChar) + parts[g + 1];
      }
      return out;
    });
  }

  // Map Strength 1–10 to Word's three native kashida jc levels, in thirds
  // (1–3 low, 4–6 medium, 7–10 high). Used by "Let Word fill it" (§3).
  function strengthToKashidaLevel(strength) {
    var s = Number(strength);
    if (!isFinite(s)) return "mediumKashida";
    if (s <= 3) return "lowKashida";
    if (s <= 6) return "mediumKashida";
    return "highKashida";
  }

  // Any Arabic-script character (Arabic, Arabic Supplement, presentation forms).
  function containsArabic(text) {
    return /[؀-ۿݐ-ݿﭐ-﷿ﹰ-﻿]/.test(String(text || ""));
  }

  // jc for "Let Word fill it": native kashida (level from strength) for Arabic,
  // else distribute (spacing — fills the last line without a trailing break).
  function wordFillJc(text, strength) {
    return containsArabic(text) ? strengthToKashidaLevel(strength) : "distribute";
  }

  // Column expansion allowance for "Let Word fill it": 0 at strength 1, ~15% at
  // full strength (10). Applied qaseeda-proportionally by the justify path.
  function kashidaExpansionFraction(strength) {
    var s = Math.max(1, Math.min(10, Number(strength) || 1));
    return Math.round((0.15 * (s - 1) / 9) * 1000) / 1000;
  }

  // ── Mehr per-font tatweel (form-aware) ───────────────────────────────────
  // Minimal Arabic shaping just for Mehr's discrete trailing tatweel.
  function isArabicMark(cp) {
    return (cp >= 0x0610 && cp <= 0x061A) || (cp >= 0x064B && cp <= 0x065F) ||
      cp === 0x0670 || (cp >= 0x06D6 && cp <= 0x06ED);
  }
  function isArabicLetter(cp) {
    return (cp >= 0x0621 && cp <= 0x063A) || (cp >= 0x0641 && cp <= 0x064A) ||
      (cp >= 0x066E && cp <= 0x066F) || (cp >= 0x0671 && cp <= 0x06D3) || cp === 0x06D5;
  }
  // Letters that connect only from the right (don't join onward to the left):
  // after one of these, the next letter is in ISOLATED form.
  var RIGHT_JOIN_ONLY = (function () {
    var s = {};
    [0x0622,0x0623,0x0624,0x0625,0x0627,0x062F,0x0630,0x0631,0x0632,0x0698,
     0x0648,0x06C1,0x06C3,0x06BA,0x0671,0x0672,0x0673,0x0675,0x0677,0x06D5,
     0x0688,0x0691,0x06D2,0x06D3].forEach(function (cp) { s[cp] = 1; });
    return s;
  }());
  function connectsLeftward(cp) {
    return isArabicLetter(cp) && !isArabicMark(cp) && !RIGHT_JOIN_ONLY[cp] && cp !== 0x0621;
  }

  // Add ONE trailing tatweel after a word's base final letter, but only when that
  // letter is allowed for its shaping form: isolated (a lone letter, or one after
  // a right-only joiner / non-joiner) uses `isolatedInto`; final (joined from the
  // previous letter) uses `finalInto`. Trailing diacritics are skipped, and the
  // tatweel goes right after the base letter (before any diacritic).
  // isolatedSet/finalSet are maps { char: true }.
  function mehrElongate(word, isolatedSet, finalSet) {
    var s = String(word);
    var i = s.length - 1;
    while (i >= 0 && isArabicMark(s.charCodeAt(i))) i--;          // skip trailing marks
    if (i < 0 || !isArabicLetter(s.charCodeAt(i))) return s;
    var base = s.charAt(i);
    var j = i - 1;
    while (j >= 0 && isArabicMark(s.charCodeAt(j))) j--;          // previous base char
    var finalForm = j >= 0 && connectsLeftward(s.charCodeAt(j));
    var allowed = finalForm ? finalSet : isolatedSet;
    if (!allowed || !allowed[base]) return s;
    return s.slice(0, i + 1) + "ـ" + s.slice(i + 1);
  }

  // ── OOXML table rendering ────────────────────────────────────────────────

  var BASE_CPM = 6; // baseline grid columns per misra

  function escapeXml(s) {
    return String(s || "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  }

  // Distribute contentCols across positions proportional to their weights (≥1 each).
  function allocateSpans(weights, contentCols) {
    var total = weights.reduce(function (a, b) { return a + b; }, 0) || 1;
    var spans = weights.map(function (w) {
      return Math.max(1, Math.round(w / total * contentCols));
    });
    var delta = contentCols - spans.reduce(function (a, b) { return a + b; }, 0);
    spans[spans.length - 1] = Math.max(1, spans[spans.length - 1] + delta);
    return spans;
  }

  function textWeight(text, ctx) {
    return ctx ? Math.max(1, ctx.measureText(String(text || "")).width) : visibleWeight(text);
  }

  // Compute proportional column spans for misra texts summing to contentCols.
  // Pass a canvas 2D context as ctx for accurate font-metric measurement; falls back to visibleWeight.
  function misraSpans(texts, contentCols, ctx) {
    return allocateSpans(texts.map(function (t) { return textWeight(t, ctx); }), contentCols);
  }

  // Compute ONE column-span vector for a whole stanza, shared by every full row
  // (rows that have exactly N misras). Each position's span is proportional to the
  // WIDEST misra in that position across the stanza — so every sadr column is the
  // same width and every ajuz column is the same width, while the sadr>ajuz
  // asymmetry is preserved. Returns null when no full rows exist.
  function stanzaColSpans(stanza, si, opts) {
    var N = si.N;
    var ctx = (opts || {})._justifyCtx || null;
    var maxW = [];
    for (var j = 0; j < N; j++) maxW.push(0);
    var found = false;
    stanza.bayts.forEach(function (b) {
      var texts = null;
      if (b.type === "row" && b.misras) texts = b.misras.map(function (m) { return m.text; });
      else if (b.ajuz) texts = [b.sadr, b.ajuz];
      if (!texts || texts.length !== N) return;
      found = true;
      for (var i = 0; i < N; i++) {
        var w = textWeight(texts[i], ctx);
        if (w > maxW[i]) maxW[i] = w;
      }
    });
    return found ? allocateSpans(maxW, si.contentCols) : null;
  }

  function stanzaGridInfo(stanza, opts, textWidthTwips) {
    var gapCols = Math.max(1, Math.round(Number((opts || {}).gapWidth || 1)));
    var N = 0;
    stanza.bayts.forEach(function (b) {
      if (b.type === "row" && b.misras) N = Math.max(N, b.misras.length);
      else if (b.ajuz) N = Math.max(N, 2);
      else N = Math.max(N, 1);
    });
    N = Math.max(N, 2);
    var GRID = N * BASE_CPM + (N - 1) * gapCols;
    var contentCols = N * BASE_CPM;
    var colWidthTwips = Math.max(1, Math.round(textWidthTwips / GRID));
    return { N: N, GRID: GRID, gapCols: gapCols, contentCols: contentCols, cwt: colWidthTwips };
  }

  function tblBordersXml() {
    var none = 'w:val="none" w:sz="0" w:space="0" w:color="auto"';
    return "<w:tblBorders>" +
      "<w:top " + none + "/><w:left " + none + "/>" +
      "<w:bottom " + none + "/><w:right " + none + "/>" +
      "<w:insideH " + none + "/><w:insideV " + none + "/>" +
      "</w:tblBorders>";
  }

  function tblBordersGridXml() {
    var thin = 'w:val="single" w:sz="4" w:space="0" w:color="auto"';
    return "<w:tblBorders>" +
      "<w:top " + thin + "/><w:left " + thin + "/>" +
      "<w:bottom " + thin + "/><w:right " + thin + "/>" +
      "<w:insideH " + thin + "/><w:insideV " + thin + "/>" +
      "</w:tblBorders>";
  }

  function tblGridXml(GRID, cwt) {
    var col = '<w:gridCol w:w="' + cwt + '"/>';
    var g = "<w:tblGrid>";
    for (var i = 0; i < GRID; i++) g += col;
    return g + "</w:tblGrid>";
  }

  function tcXml(span, cwt, paraXml) {
    return "<w:tc>" +
      "<w:tcPr>" +
        '<w:gridSpan w:val="' + span + '"/>' +
        '<w:tcW w:w="' + (span * cwt) + '" w:type="dxa"/>' +
        (paraXml ? '<w:vAlign w:val="bottom"/>' : "") +
      "</w:tcPr>" +
      (paraXml || "<w:p/>") +
      "</w:tc>";
  }

  // indTwips: optional right-indent in twips (used by stacked layout to offset ajuz from sadr)
  function misraParaXml(text, align, isRefrain, opts, indTwips) {
    opts = opts || {};
    var jc = align === "right" ? "right" : align === "left" ? "left" : "center";
    var trailingBreak = "";
    if (opts.justifyMode === "css") {
      // "Let Word fill it": native Word justification. Arabic → kashida level +
      // a shrunk trailing break so the single (last) line actually stretches;
      // non-Arabic → distribute (fills the last line without a break).
      jc = wordFillJc(text, Number(opts.tatweelCount || 0));
      if (containsArabic(text)) {
        trailingBreak = '<w:r><w:rPr><w:sz w:val="4"/><w:szCs w:val="4"/></w:rPr><w:br/></w:r>';
      }
    }
    var rpr = "<w:rPr><w:rtl/>";
    if (isRefrain) rpr += '<w:color w:val="A7352A"/>';
    // Font routed through the AshaarFonts registry (nastaliq engine merge).
    // opts.fontCsName pins an EXACT cs font (Re-render preserving an arbitrary
    // existing font — e.g. Fatemi Maqala — that isn't in the registry),
    // overriding the fontMode → wordName lookup.
    var mode = opts.fontMode === "nastaliq" ? "noto" : opts.fontMode;
    var csName = opts.fontCsName || AshaarFonts.wordNameOf(mode);
    if (csName) rpr += '<w:rFonts w:cs="' + csName + '"/>';
    // Size-preserving re-render: when the caller supplies the cell's point size
    // (gap/width apply that rebuilds the poem), emit it so the rebuilt run keeps
    // its size instead of reverting to Word's document default. Omitted on plain
    // insert (no fontSizePt) so that path is unchanged.
    if (opts.fontSizePt) {
      var szHp = Math.round(Number(opts.fontSizePt) * 2);
      rpr += '<w:sz w:val="' + szHp + '"/><w:szCs w:val="' + szHp + '"/>';
    }
    rpr += "</w:rPr>";
    var ind = indTwips ? '<w:ind w:left="' + indTwips + '"/>' : "";
    // Shrink the paragraph mark when word-fill trailing break is emitted, so the
    // empty line after the break is ~2pt high instead of full paragraph height.
    var paraMark = trailingBreak ? "<w:rPr><w:sz w:val=\"4\"/><w:szCs w:val=\"4\"/></w:rPr>" : "";
    return "<w:p>" +
      "<w:pPr><w:bidi/><w:spacing w:after=\"80\"/><w:jc w:val=\"" + jc + "\"/>" + ind + paraMark + "</w:pPr>" +
      "<w:r>" + rpr + '<w:t xml:space="preserve">' + escapeXml(text) + "</w:t></w:r>" +
      trailingBreak +
      "</w:p>";
  }

  // runs: [{text, swap}]; base vs Kasheeda cs name chosen per run (Jameel
  // font-swap kashida — Task 8). swap:true fasls render in the wider
  // "Kasheeda" face; the rest stay in the base face.
  function runsToMisraXml(runs, align, opts, sizePt) {
    var jc = align === "right" ? "right" : align === "left" ? "left" : "center";
    var mode = (opts || {}).fontMode === "nastaliq" ? "noto" : (opts || {}).fontMode;
    var baseName = AshaarFonts.wordNameOf(mode);
    var wideName = AshaarFonts.kasheedaNameOf(mode) || baseName;
    // Emit the cell's size on every run — a full-paragraph insertOoxml replace
    // does NOT inherit the previous runs' size, so without this the swapped
    // paragraph reverts to Word's document default (12pt).
    var szXml = sizePt ? '<w:sz w:val="' + Math.round(sizePt * 2) + '"/><w:szCs w:val="' + Math.round(sizePt * 2) + '"/>' : "";
    var body = (runs || []).map(function (r) {
      var cs = r.swap ? wideName : baseName;
      var rpr = "<w:rPr><w:rtl/>" + (cs ? '<w:rFonts w:cs="' + cs + '"/>' : "") + szXml + "</w:rPr>";
      return "<w:r>" + rpr + '<w:t xml:space="preserve">' + escapeXml(r.text) + "</w:t></w:r>";
    }).join("");
    return "<w:p><w:pPr><w:bidi/><w:spacing w:after=\"80\"/><w:jc w:val=\"" + jc + "\"/></w:pPr>" + body + "</w:p>";
  }

  // Cell-fit residual is Word Distributed justification: emit the (tatweel'd)
  // misra as a paragraph with <w:jc w:val="distribute"/> so Word stretches the
  // inter-word gaps to the true cell edge. Each run keeps its own cs font (+
  // size); NO micro-spaces are injected (that is the Natural-fit residual).
  // runs: [{text, csName, sizePt?}].
  function misraDistributeXml(runs, sizePtFallback) {
    var body = (runs || []).map(function (r) {
      var sz = r.sizePt || sizePtFallback;
      var szXml = sz ? '<w:sz w:val="' + Math.round(sz * 2) + '"/><w:szCs w:val="' + Math.round(sz * 2) + '"/>' : "";
      var cs = r.csName ? '<w:rFonts w:cs="' + r.csName + '"/>' : "";
      return "<w:r><w:rPr><w:rtl/>" + cs + szXml + "</w:rPr>" +
        '<w:t xml:space="preserve">' + escapeXml(r.text) + "</w:t></w:r>";
    }).join("");
    return "<w:p><w:pPr><w:bidi/><w:spacing w:after=\"80\"/><w:jc w:val=\"distribute\"/></w:pPr>" + body + "</w:p>";
  }

  // Natural-fit residual for a font-swap / discrete-tatweel / mixed-font misra:
  // emit the runs with their EXPLICIT per-run cs face (+size +color) and a REAL
  // jc (right/center/left) — NOT `distribute` (which no-ops on a single line).
  // The injected hair-spaces do the residual filling and the line keeps its
  // visual side. Unlike runsToMisraXml, the face is per-run (r.csName) rather
  // than derived from a pane fontMode — so the apply path can name each cell's
  // OWN per-word font, which is what lets a mixed-font misra survive a rebuild.
  //   runs: [{text, csName, sizePt?, color?}]   color = "#RRGGBB" or "RRGGBB"
  //   opts: { indentTwips? }  paragraph w:ind w:left (stacked-layout ajuz offset)
  function misraRunsXml(runs, align, sizePtFallback, opts) {
    opts = opts || {};
    var jc = align === "right" ? "right" : align === "left" ? "left" : "center";
    var body = (runs || []).map(function (r) {
      var sz = r.sizePt || sizePtFallback;
      var szXml = sz ? '<w:sz w:val="' + Math.round(sz * 2) + '"/><w:szCs w:val="' + Math.round(sz * 2) + '"/>' : "";
      // Set ascii+hAnsi as WELL as cs. cs alone renders correctly, but Office.js
      // Font.name reads back the ASCII font, so on a re-apply the captured font
      // would be the document default (not this run's) and the cell would flatten
      // to one font — breaking idempotency. Naming all three makes the per-word
      // font survive a capture → write → capture round-trip.
      //
      // asciiName (when given) decouples the read-back name from the rendered
      // face: Arabic renders via cs on these rtl runs, but Font.name reads ascii,
      // and a word whose runs carry DIFFERENT ascii fonts reads back "" (mixed).
      // A font-swap word (some fasls base, some Kasheeda) must therefore pin
      // ascii+hAnsi to the BASE face on every run — one family per word — or the
      // next apply misclassifies it as generic and the round-trip diverges.
      var asc = r.asciiName || r.csName;
      var cs = r.csName
        ? '<w:rFonts w:ascii="' + asc + '" w:hAnsi="' + asc + '" w:cs="' + r.csName + '"/>'
        : "";
      var col = "";
      if (r.color && /^#?[0-9a-fA-F]{6}$/.test(r.color)) col = '<w:color w:val="' + r.color.replace(/^#/, "") + '"/>';
      // Bold/italic need BOTH variants: Word styles rtl/Arabic runs via the
      // complex-script bCs/iCs, and b/i keeps any Latin content consistent.
      var bi = (r.bold ? "<w:b/><w:bCs/>" : "") + (r.italic ? "<w:i/><w:iCs/>" : "");
      // Residual gap spacing: rPr w:spacing is CHARACTER spacing (twips added
      // to each glyph's advance) — on a single-space run it widens just that
      // gap, pixel-exact, with no injected characters to strip on re-capture.
      var spc = (r.spacingTwips > 0) ? '<w:spacing w:val="' + Math.round(r.spacingTwips) + '"/>' : "";
      // Debug tint for spacing runs: w:shd takes a hex fill (w:highlight is a
      // named-enum and can't carry the profile's color-picker value).
      var shd = (r.shdFill && /^#?[0-9a-fA-F]{6}$/.test(r.shdFill))
        ? '<w:shd w:val="clear" w:fill="' + r.shdFill.replace(/^#/, "").toUpperCase() + '"/>' : "";
      return "<w:r><w:rPr><w:rtl/>" + bi + col + spc + shd + cs + szXml + "</w:rPr>" +
        '<w:t xml:space="preserve">' + escapeXml(r.text) + "</w:t></w:r>";
    }).join("");
    var ind = opts.indentTwips ? '<w:ind w:left="' + Math.round(opts.indentTwips) + '"/>' : "";
    return "<w:p><w:pPr><w:bidi/><w:spacing w:after=\"80\"/><w:jc w:val=\"" + jc + "\"/>" + ind + "</w:pPr>" + body + "</w:p>";
  }

  function baytRowsOoxml(bayt, si, opts) {
    var N = si.N, gapCols = si.gapCols, cwt = si.cwt;
    var textWidthPx = (opts || {})._textWidthPx || 0;

    function justify(text, span) {
      var px = textWidthPx > 0 ? (span / si.GRID) * textWidthPx : 0;
      return justifyText(text, opts, px);
    }

    function gapTc() { return tcXml(gapCols, cwt, null); }
    function padTc(p) { return p > 0 ? tcXml(p, cwt, null) : ""; }

    // Solo misra is the SAME width as an ordinary misra (BASE_CPM columns),
    // centred, with empty pad cells flanking it — not a single full-grid cell.
    // In a marsiya (centered solo line + paired refrain lines) a full-grid cell
    // stretched those lines across the whole table width, which looked wrong;
    // they should match the width of the other misras. Spans still sum to
    // si.GRID (left pad + BASE_CPM + right pad), so fixed table layout stays valid.
    // indTwips: right-indent passed through to misraParaXml for stacked ajuz offset
    function soloRow(text, isRefrain, indTwips) {
      var span = BASE_CPM;
      var pad = Math.max(0, si.GRID - span);
      var left = Math.floor(pad / 2);
      var right = pad - left;
      var para = misraParaXml(justify(text, span), "center", isRefrain, opts, indTwips || 0);
      return "<w:tr>" + padTc(left) + tcXml(span, cwt, para) + padTc(right) + "</w:tr>";
    }

    function misraRow(texts, refrains) {
      var K = texts.length;
      // Full available content width (GRID minus gaps) so partial rows fill the table width.
      // The gap boundary lands on a gridCol boundary, giving Word layout flexibility.
      var kContentCols = si.GRID - (K - 1) * gapCols;
      // Full rows (K === N) use the stanza-wide shared spans so columns align across
      // rows; partial rows fall back to their own proportional split.
      var spans = (K === si.N && si.colSpans)
        ? si.colSpans
        : misraSpans(texts, kContentCols, (opts || {})._justifyCtx || null);
      var cells = "";
      for (var i = 0; i < K; i++) {
        var align = i === 0 ? "right" : i === K - 1 ? "left" : "center";
        var px = textWidthPx > 0 ? (spans[i] / si.GRID) * textWidthPx : 0;
        cells += tcXml(spans[i], cwt, misraParaXml(justifyText(texts[i], opts, px), align, refrains[i], opts));
        if (i < K - 1) cells += gapTc();
      }
      return "<w:tr>" + cells + "</w:tr>";
    }

    if (bayt.type === "row") {
      var misras = bayt.misras || [];
      var K = misras.length;
      if (K === 0) return "";
      if (K === 1) {
        return soloRow(misras[0].text, !!misras[0].isRefrain);
      }
      if ((opts || {}).layoutMode === "stacked") {
        // Each misra gets its own centered row; first at true center, rest with 640-twips (32pt) indent
        return misras.slice(0, N).map(function (m, i) {
          return soloRow(m.text, !!m.isRefrain, i > 0 ? 640 : 0);
        }).join("");
      }
      var texts = misras.map(function (m) { return m.text; });
      var refs = misras.map(function (m) { return !!m.isRefrain; });
      return misraRow(texts, refs);
    }

    // Old-format bayt (sadr / ajuz)
    if (!bayt.ajuz) {
      return soloRow(bayt.sadr || "", !!bayt.sadrRefrain);
    }
    if ((opts || {}).layoutMode === "stacked") {
      // Two rows: sadr at center, ajuz indented 640 twips (32pt) to match HTML stacked offset
      return soloRow(bayt.sadr || "", !!bayt.sadrRefrain) +
             soloRow(bayt.ajuz, !!bayt.ajuzRefrain, 640);
    }
    return misraRow(
      [bayt.sadr, bayt.ajuz],
      [!!bayt.sadrRefrain, !!bayt.ajuzRefrain]
    );
  }

  // The content/gap KIND sequence a bayt contributes, per row — mirrors
  // baytRowsOoxml's branches (solo → [g,c,g]; K-misra row → [c,g,c,…,c]; stacked
  // → one solo row per misra). Spans/alignment/indents don't change cell order,
  // so only the kind matters here. This is the single derivation the persisted
  // pattern and (via AshaarCellMap) all labels come from; the cross-check test
  // (tests/word-html.test.js) locks it to the actual OOXML the generator emits.
  function baytCellPatternRows(bayt, opts) {
    var stacked = (opts || {}).layoutMode === "stacked";
    function solo() { return ["g", "c", "g"]; }
    function misra(K) { var r = []; for (var i = 0; i < K; i++) { r.push("c"); if (i < K - 1) r.push("g"); } return r; }
    if (bayt.type === "row") {
      var K = (bayt.misras || []).length;
      if (K === 0) return [];
      if (K === 1) return [solo()];
      if (stacked) { var rows = []; for (var i = 0; i < K; i++) rows.push(solo()); return rows; }
      return [misra(K)];
    }
    if (!bayt.ajuz) return [solo()];
    if (stacked) return [solo(), solo()];
    return [misra(2)];
  }

  function stanzaCellPattern(stanza, opts) {
    var rows = [];
    (stanza.bayts || []).forEach(function (b) {
      baytCellPatternRows(b, opts).forEach(function (r) { rows.push(r); });
    });
    return rows;
  }

  function poemCellPatterns(text, opts, Ashaar) {
    var poems = parsePoetry(String(text || ""), Ashaar);
    var pats = [];
    poems.forEach(function (poem) {
      (poem.stanzas || []).forEach(function (stanza) { pats.push(stanzaCellPattern(stanza, opts)); });
    });
    return pats;
  }

  // The grid SPANS + content/gap KIND a bayt contributes, per row — mirrors
  // baytRowsOoxml's branches (solo → pad/content/pad; K-misra row → content+gap;
  // stacked → one solo row per misra) but emits geometry instead of OOXML. The
  // apply engine needs each cell's grid span because Word cannot report column
  // geometry for span (merged-cell) tables. A cross-check test locks this to the
  // generator's actual <w:gridSpan> values.
  function baytCellGeometryRows(bayt, si, opts) {
    var N = si.N, gapCols = si.gapCols;
    var stacked = (opts || {}).layoutMode === "stacked";

    function soloCells() {
      var span = BASE_CPM;
      var pad = Math.max(0, si.GRID - span);
      var left = Math.floor(pad / 2), right = pad - left;
      var out = [];
      if (left > 0) out.push({ kind: "spacing", span: left });
      out.push({ kind: "content", span: span });
      if (right > 0) out.push({ kind: "spacing", span: right });
      return out;
    }
    function misraCells(texts) {
      var K = texts.length;
      var kContentCols = si.GRID - (K - 1) * gapCols;
      var spans = (K === si.N && si.colSpans)
        ? si.colSpans
        : misraSpans(texts, kContentCols, (opts || {})._justifyCtx || null);
      var out = [];
      for (var i = 0; i < K; i++) {
        out.push({ kind: "content", span: spans[i] });
        if (i < K - 1) out.push({ kind: "spacing", span: gapCols });
      }
      return out;
    }

    if (bayt.type === "row") {
      var misras = bayt.misras || [];
      var K = misras.length;
      if (K === 0) return [];
      if (K === 1) return [soloCells()];
      if (stacked) return misras.slice(0, N).map(function () { return soloCells(); });
      return [misraCells(misras.map(function (m) { return m.text; }))];
    }
    if (!bayt.ajuz) return [soloCells()];
    if (stacked) return [soloCells(), soloCells()];
    return [misraCells([bayt.sadr, bayt.ajuz])];
  }

  // Per-cell grid geometry for a stanza, in the SAME emission order as
  // baytRowsOoxml / stanzaCellPattern (so it zips against a live table's
  // row-major cells and the persisted pattern). Returns
  //   { GRID, rows: [ [ { kind:"content"|"spacing", col, span } … ] … ] }
  // where col/span are grid-column coordinates (col = running left→right sum).
  // textWidthTwips only feeds stanzaGridInfo's cwt (unused here); any positive
  // value works.
  function stanzaCellGeometry(stanza, opts, textWidthTwips) {
    var si = stanzaGridInfo(stanza, opts, textWidthTwips || 9360);
    si.colSpans = stanzaColSpans(stanza, si, opts);
    var rows = [];
    (stanza.bayts || []).forEach(function (bayt) {
      baytCellGeometryRows(bayt, si, opts).forEach(function (rowCells) {
        var col = 0;
        rows.push(rowCells.map(function (c) {
          var e = { kind: c.kind, col: col, span: c.span };
          col += c.span;
          return e;
        }));
      });
    });
    return { GRID: si.GRID, rows: rows };
  }

  // Per-stanza geometry for a whole poem source (parallels poemCellPatterns).
  function poemCellGeometry(text, opts, Ashaar, textWidthTwips) {
    var poems = parsePoetry(String(text || ""), Ashaar);
    var out = [];
    poems.forEach(function (poem) {
      (poem.stanzas || []).forEach(function (stanza) {
        out.push(stanzaCellGeometry(stanza, opts, textWidthTwips));
      });
    });
    return out;
  }

  function stanzaTableOoxml(stanza, opts, textWidthTwips) {
    var si = stanzaGridInfo(stanza, opts, textWidthTwips);
    si.colSpans = stanzaColSpans(stanza, si, opts); // shared spans → columns align across rows
    // Fixed layout + a definite width make the shared grid rigid: Word uses the
    // gridCol widths verbatim instead of auto-fitting columns to content across
    // rows. Without this, a wide solo misra widens the centre columns and drags
    // the couplet's split with it — coupling rows within a stanza.
    var totalW = si.GRID * si.cwt;
    var tblPr = "<w:tblPr>" +
      '<w:tblW w:w="' + totalW + '" w:type="dxa"/>' +
      '<w:jc w:val="center"/>' +
      '<w:tblLayout w:type="fixed"/>' +
      tblBordersXml() +
      "<w:bidiVisual/>" +
      "</w:tblPr>";
    var rows = stanza.bayts.map(function (bayt) {
      return baytRowsOoxml(bayt, si, opts);
    }).filter(Boolean).join("");
    return "<w:tbl>" + tblPr + tblGridXml(si.GRID, si.cwt) + rows + "</w:tbl>";
  }

  // Generate a blank 12-column grid table for the user to merge/resize in Word.
  // Thin borders are shown so cells are clearly visible.
  function generateBareGrid12Ooxml(textWidthTwips) {
    var twips = textWidthTwips > 0 ? textWidthTwips : 9360;
    var cwt = Math.round(twips / 12);
    var tblPr = "<w:tblPr>" +
      '<w:tblW w:w="0" w:type="auto"/>' +
      '<w:jc w:val="center"/>' +
      tblBordersGridXml() +
      "<w:bidiVisual/>" +
      "</w:tblPr>";
    var emptyPara = "<w:p><w:pPr><w:bidi/><w:spacing w:after=\"200\"/></w:pPr></w:p>";
    var cells = "";
    for (var i = 0; i < 12; i++) cells += tcXml(1, cwt, emptyPara);
    return "<w:tbl>" + tblPr + tblGridXml(12, cwt) + "<w:tr>" + cells + "</w:tr></w:tbl>";
  }

  // Generate an OOXML table from a captured template (rows of {span} cells).
  // GRID defaults to templateData.columnCount (12 for captured grids).
  function templateToOoxml(templateData, textWidthTwips, opts) {
    var twips = textWidthTwips > 0 ? textWidthTwips : 9360;
    var GRID = templateData.columnCount || 12;
    var cwt = Math.round(twips / GRID);
    var tblPr = "<w:tblPr>" +
      '<w:tblW w:w="0" w:type="auto"/>' +
      '<w:jc w:val="center"/>' +
      tblBordersXml() +
      "<w:bidiVisual/>" +
      "</w:tblPr>";
    var emptyPara = "<w:p><w:pPr><w:bidi/><w:spacing w:after=\"80\"/></w:pPr></w:p>";
    var rpr = "<w:rPr><w:rtl/>";
    var mode = (opts || {}).fontMode === "nastaliq" ? "noto" : (opts || {}).fontMode;
    var csName = AshaarFonts.wordNameOf(mode);
    if (csName) rpr += '<w:rFonts w:cs="' + csName + '"/>';
    rpr += "</w:rPr>";
    var rows = (templateData.rows || []).map(function (row) {
      var cells = row.map(function (cell, cIdx) {
        var jc = cell.align || (cIdx === 0 ? "right" : cIdx === row.length - 1 ? "left" : "center");
        var para = "<w:p><w:pPr><w:bidi/><w:spacing w:after=\"80\"/><w:jc w:val=\"" + jc + "\"/></w:pPr>" +
          "<w:r>" + rpr + "<w:t/></w:r></w:p>";
        return tcXml(cell.span, cwt, para);
      }).join("");
      return "<w:tr>" + cells + "</w:tr>";
    }).join("");
    return "<w:tbl>" + tblPr + tblGridXml(GRID, cwt) + rows + "</w:tbl>";
  }

  // Non-span layout table (the "Numbers" view and other plain, non-indented
  // layouts). Mirrors the fixed-width RTL builder above but honours per-column
  // widths and fills each cell with its (placeholder) text. The <w:bidiVisual/>
  // makes this a genuine right-to-left table: cell/tab order runs right-to-left
  // and the logical-first cell (column 0) is displayed on the right — unlike the
  // native Word.insertTable path, which produces an LTR table whose cell order
  // (and tab order) runs left-to-right even when the content looks correct.
  function layoutTableToOoxml(layoutTable, textWidthTwips, opts) {
    var twips = textWidthTwips > 0 ? textWidthTwips : 9360;
    var GRID = layoutTable.columnCount || 1;
    var pct = (layoutTable.widths && layoutTable.widths.length === GRID)
      ? layoutTable.widths
      : layoutColumnWidths(GRID, opts);
    var colTwips = pct.map(function (p) { return Math.max(8, Math.round(twips * p / 100)); });
    var totalW = colTwips.reduce(function (a, b) { return a + b; }, 0);
    var tblPr = "<w:tblPr>" +
      '<w:tblW w:w="' + totalW + '" w:type="dxa"/>' +
      '<w:jc w:val="center"/>' +
      '<w:tblLayout w:type="fixed"/>' +
      tblBordersXml() +
      "<w:bidiVisual/>" +
      "</w:tblPr>";
    var grid = "<w:tblGrid>" + colTwips.map(function (w) {
      return '<w:gridCol w:w="' + w + '"/>';
    }).join("") + "</w:tblGrid>";
    var rpr = "<w:rPr><w:rtl/>";
    var mode = (opts || {}).fontMode === "nastaliq" ? "noto" : (opts || {}).fontMode;
    var csName = AshaarFonts.wordNameOf(mode);
    if (csName) rpr += '<w:rFonts w:cs="' + csName + '"/>';
    rpr += "</w:rPr>";
    var rows = (layoutTable.rows || []).map(function (row) {
      var cells = row.map(function (cell, cIdx) {
        var jc = cell.align || (cIdx === 0 ? "right" : cIdx === row.length - 1 ? "left" : "center");
        var para = "<w:p><w:pPr><w:bidi/><w:spacing w:after=\"80\"/><w:jc w:val=\"" + jc + "\"/></w:pPr>" +
          "<w:r>" + rpr + '<w:t xml:space="preserve">' + escapeXml(blankMisraLabel(cell.text)) + "</w:t></w:r></w:p>";
        return tcXml(1, colTwips[cIdx] || colTwips[0], para);
      }).join("");
      return "<w:tr>" + cells + "</w:tr>";
    }).join("");
    return "<w:tbl>" + tblPr + grid + rows + "</w:tbl>";
  }

  function renderForWordOoxml(text, opts, Ashaar, textWidthTwips) {
    opts = opts || {};
    var twips = (textWidthTwips > 0) ? textWidthTwips : 9360;
    var poems = parsePoetry(String(text || ""), Ashaar);
    if (!poems.length) return "";
    var tables = [];
    poems.forEach(function (poem) {
      poem.stanzas.forEach(function (stanza) {
        tables.push(stanzaTableOoxml(stanza, opts, twips));
      });
    });
    return tables.join("<w:p/>");
  }

  function wrapOoxml(bodyContent) {
    var wns = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
    var pkgns = 'xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage"';
    var rels = '<pkg:part pkg:name="/_rels/.rels"' +
      ' pkg:contentType="application/vnd.openxmlformats-package.relationships+xml"' +
      ' pkg:padding="512"><pkg:xmlData>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1"' +
      ' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"' +
      ' Target="word/document.xml"/>' +
      '</Relationships></pkg:xmlData></pkg:part>';
    var docPart = '<pkg:part pkg:name="/word/document.xml"' +
      ' pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">' +
      '<pkg:xmlData><w:document ' + wns + '><w:body>' +
      bodyContent + '<w:sectPr/></w:body></w:document></pkg:xmlData></pkg:part>';
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<pkg:package ' + pkgns + '>' + rels + docPart + '</pkg:package>';
  }

  // Wrap body content in a block-level Structured Document Tag (content control)
  // carrying the title + tag, then package it. insertOoxml on this creates a
  // Word content control that spans the ENTIRE content (all table rows) — unlike
  // insertContentControl() on an insertOoxml-returned range, which on Mac Word
  // wraps only the first block element (the row-1-only bug). w:id is a stable
  // hash of the tag so re-applied blocks don't collide.
  function wrapOoxmlControl(bodyContent, title, tag) {
    var s = String(tag || title || ""), idNum = 0;
    for (var i = 0; i < s.length; i++) idNum = ((idNum << 5) - idNum + s.charCodeAt(i)) | 0;
    idNum = (Math.abs(idNum) % 2000000000) || 1;
    var sdt = "<w:sdt><w:sdtPr>" +
      '<w:alias w:val="' + escapeXml(title || "") + '"/>' +
      '<w:tag w:val="' + escapeXml(tag || "") + '"/>' +
      '<w:id w:val="' + idNum + '"/>' +
      "</w:sdtPr><w:sdtContent>" + (bodyContent || "") + "</w:sdtContent></w:sdt>";
    return wrapOoxml(sdt);
  }

  return {
    fontFamilyStyle: fontFamilyStyle,
    renderForWord: renderForWord,
    renderTextWithLayoutForWord: renderTextWithLayoutForWord,
    renderTemplateForWord: renderTemplateForWord,
    layoutTablesForText: layoutTablesForText,
    layoutTablesForTemplate: layoutTablesForTemplate,
    layoutTablesForPoem: layoutTablesForPoem,
    extractMisras: extractMisras,
    templateGrid: templateGrid,
    templateColumns: templateColumns,
    parseLayoutSpec: parseLayoutSpec,
    tableColumns: tableColumns,
    contentControlTag: contentControlTag,
    parseContentControlTag: parseContentControlTag,
    setTagQaseeda: setTagQaseeda,
    setTagOverride: setTagOverride,
    setTagSlotDecor: setTagSlotDecor,
    packRunWords: packRunWords,
    reconcileRunWords: reconcileRunWords,
    setTagRunFonts: setTagRunFonts,
    setTagBandhWidth: setTagBandhWidth,
    justifyPlainTextBlock: justifyPlainTextBlock,
    coalesceRuns: coalesceRuns,
    distributeMicroSpaces: distributeMicroSpaces,
    strengthToElongationShare: strengthToElongationShare,
    strengthToMaxPositions: strengthToMaxPositions,
    strengthToKashidaLevel: strengthToKashidaLevel,
    containsArabic: containsArabic,
    wordFillJc: wordFillJc,
    kashidaExpansionFraction: kashidaExpansionFraction,
    mehrElongate: mehrElongate,
    renderForWordOoxml: renderForWordOoxml,
    stanzaCellPattern: stanzaCellPattern,
    poemCellPatterns: poemCellPatterns,
    stanzaCellGeometry: stanzaCellGeometry,
    poemCellGeometry: poemCellGeometry,
    runsToMisraXml: runsToMisraXml,
    misraDistributeXml: misraDistributeXml,
    misraRunsXml: misraRunsXml,
    wrapOoxml: wrapOoxml,
    wrapOoxmlControl: wrapOoxmlControl,
    misraSpans: misraSpans,
    generateBareGrid12Ooxml: generateBareGrid12Ooxml,
    templateToOoxml: templateToOoxml,
    layoutTableToOoxml: layoutTableToOoxml,
    misraParaXml: misraParaXml
  };
}));

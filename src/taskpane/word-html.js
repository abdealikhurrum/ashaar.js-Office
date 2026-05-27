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

    if (layout === "equal" || opts.widthMode === "fixed") {
      var equalText = (100 - gap) / 2;
      return { sadr: equalText, gap: gap, ajuz: equalText, mode: "three" };
    }

    var maxSadr = 1;
    var maxAjuz = 1;
    stanza.bayts.forEach(function (bayt) {
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
      "<col style=\"" + widthStyle(cols.sadr) + "\">" +
      "<col style=\"" + widthStyle(cols.gap) + "\">" +
      "<col style=\"" + widthStyle(cols.ajuz) + "\">" +
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

  function renderGridRow(row, opts) {
    if (row.type === "center") {
      return "<tr><td colspan=\"3\" style=\"" + misraCellStyle(row, opts) + "\">" + escapeHtml(blankMisraLabel(row.misra)) + "</td></tr>";
    }
    if (row.type === "right") {
      return "<tr>" +
        "<td style=\"" + misraCellStyle(Object.assign({ align: "left" }, row), opts) + "\">" + escapeHtml(blankMisraLabel(row.misra)) + "</td>" +
        "<td style=\"" + gapStyle() + "\"></td>" +
        "<td style=\"padding:0\"></td>" +
        "</tr>";
    }
    if (row.type === "left") {
      return "<tr>" +
        "<td style=\"padding:0\"></td>" +
        "<td style=\"" + gapStyle() + "\"></td>" +
        "<td style=\"" + misraCellStyle(Object.assign({ align: "right" }, row), opts) + "\">" + escapeHtml(blankMisraLabel(row.misra)) + "</td>" +
        "</tr>";
    }
    return "<tr>" +
      "<td style=\"" + misraCellStyle({ align: "left" }, opts) + "\">" + escapeHtml(blankMisraLabel(row.right)) + "</td>" +
      "<td style=\"" + gapStyle() + "\"></td>" +
      "<td style=\"" + misraCellStyle({ align: "right" }, opts) + "\">" + escapeHtml(blankMisraLabel(row.left)) + "</td>" +
      "</tr>";
  }

  function renderBaytRow(bayt, opts, cols) {
    var sadr = escapeHtml(justifyText(bayt.sadr, opts));
    var ajuz = bayt.ajuz ? escapeHtml(justifyText(bayt.ajuz, opts)) : "";
    var sadrColor = bayt.sadrRefrain ? "color:#a7352a;" : "";
    var ajuzColor = bayt.ajuzRefrain ? "color:#a7352a;" : "";

    if (!bayt.ajuz || opts.layoutMode === "stacked") {
      var second = bayt.ajuz ? "<br><span style=\"padding-right:32pt;" + ajuzColor + "\">" + ajuz + "</span>" : "";
      var spanCount = cols.mode === "compact" ? 5 : 3;
      return "<tr><td colspan=\"" + spanCount + "\" style=\"text-align:center;direction:rtl;padding:0 0 5pt 0;" + sadrColor + "\">" + sadr + second + "</td></tr>";
    }

    if (cols.mode === "compact") {
      return "<tr>" +
        "<td style=\"padding:0\"></td>" +
        "<td style=\"" + rightSideCellStyle(opts) + sadrColor + "\">" + sadr + "</td>" +
        "<td style=\"" + gapStyle() + "\"></td>" +
        "<td style=\"" + leftSideCellStyle(opts) + ajuzColor + "\">" + ajuz + "</td>" +
        "<td style=\"padding:0\"></td>" +
        "</tr>";
    }

    return "<tr>" +
      "<td style=\"" + rightSideCellStyle(opts) + sadrColor + "\">" + sadr + "</td>" +
      "<td style=\"" + gapStyle() + "\"></td>" +
      "<td style=\"" + leftSideCellStyle(opts) + ajuzColor + "\">" + ajuz + "</td>" +
      "</tr>";
  }

  function renderForWord(text, opts, Ashaar) {
    opts = opts || {};
    var poems = parsePoetry(text, Ashaar);
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

  function templateGrid(opts) {
    opts = opts || {};
    var count = clamp(Number(opts.misraCount || 4), 1, 80);
    var pattern = opts.misraPattern || "paired";
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

    for (i = 1; i <= count; i += 2) {
      rows.push({ type: "pair", right: i, left: i + 1 <= count ? i + 1 : "\u00a0" });
    }
    return rows;
  }

  function templateColumns(opts) {
    var gap = clamp(Number((opts || {}).gapWidth || 4), 2, 10);
    var side = (100 - gap) / 2;
    return { sadr: side, gap: gap, ajuz: side, mode: "three" };
  }

  function renderTemplateForWord(opts) {
    opts = Object.assign({}, opts || {});
    var bandhCount = clamp(Number(opts.bandhCount || 1), 1, 20);
    var layout = opts.layoutMode || "equal";
    opts.layoutMode = layout;
    opts.widthMode = opts.widthMode || "fixed";

    var chunks = [];
    for (var i = 0; i < bandhCount; i++) {
      var rows = templateGrid(opts);
      var cols = templateColumns(opts);
      chunks.push("<table dir=\"rtl\" data-ashaar-layout=\"" + escapeHtml(layout) + "\" data-ashaar-template=\"true\" style=\"" + tableStyle(opts) + "\">" +
        colGroup(cols) +
        "<tbody>" +
        rows.map(function (row) { return renderGridRow(row, opts); }).join("") +
        "</tbody></table>");
    }
    return chunks.join("<p style=\"margin:10pt 0\"></p>");
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
    renderTemplateForWord: renderTemplateForWord,
    templateGrid: templateGrid,
    templateColumns: templateColumns,
    tableColumns: tableColumns,
    contentControlTag: contentControlTag,
    justifyPlainTextBlock: justifyPlainTextBlock
  };
}));

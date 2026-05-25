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

  function tableStyle(opts) {
    opts = opts || {};
    return [
      "width:100%",
      "border-collapse:collapse",
      "direction:rtl",
      "font-family:'Noto Nastaliq Urdu','Scheherazade New','Amiri',serif",
      "font-size:16pt",
      "line-height:2.1",
      "margin:0 0 14pt 0"
    ].join(";");
  }

  function cellStyle(align, opts) {
    opts = opts || {};
    var justify = opts.justifyMode === "css" ? "text-align:justify;text-justify:inter-word;" : "text-align:" + align + ";";
    return [
      "width:48%",
      "vertical-align:baseline",
      "padding:0 0 3pt 0",
      justify,
      "direction:rtl"
    ].join(";");
  }

  function gapStyle(opts) {
    return [
      "width:" + cssLengthPercent(opts.gapWidth, "4%"),
      "padding:0 5pt",
      "text-align:center",
      "vertical-align:baseline"
    ].join(";");
  }

  function renderBaytRow(bayt, opts) {
    var sadr = escapeHtml(justifyText(bayt.sadr, opts));
    var ajuz = bayt.ajuz ? escapeHtml(justifyText(bayt.ajuz, opts)) : "";
    var sadrColor = bayt.sadrRefrain ? "color:#a7352a;" : "";
    var ajuzColor = bayt.ajuzRefrain ? "color:#a7352a;" : "";

    if (!bayt.ajuz || opts.layoutMode === "stacked") {
      var second = bayt.ajuz ? "<br><span style=\"padding-right:32pt;" + ajuzColor + "\">" + ajuz + "</span>" : "";
      return "<tr><td colspan=\"3\" style=\"text-align:center;direction:rtl;padding:0 0 5pt 0;" + sadrColor + "\">" + sadr + second + "</td></tr>";
    }

    return "<tr>" +
      "<td style=\"" + cellStyle("left", opts) + sadrColor + "\">" + sadr + "</td>" +
      "<td style=\"" + gapStyle(opts) + "\"></td>" +
      "<td style=\"" + cellStyle("right", opts) + ajuzColor + "\">" + ajuz + "</td>" +
      "</tr>";
  }

  function renderForWord(text, opts, Ashaar) {
    opts = opts || {};
    var poems = parsePoetry(text, Ashaar);
    if (!poems.length) return "";
    return poems.map(function (poem) {
      return poem.stanzas.map(function (stanza) {
        return "<table dir=\"rtl\" style=\"" + tableStyle(opts) + "\"><tbody>" +
          stanza.bayts.map(function (bayt) { return renderBaytRow(bayt, opts); }).join("") +
          "</tbody></table>";
      }).join("<p style=\"margin:10pt 0\"></p>");
    }).join("<p style=\"margin:18pt 0\"></p>");
  }

  function justifyPlainTextBlock(text, opts) {
    return String(text || "").split(/\r?\n/).map(function (line) {
      return line.trim() ? justifyText(line, opts) : line;
    }).join("\n");
  }

  return {
    renderForWord: renderForWord,
    justifyPlainTextBlock: justifyPlainTextBlock
  };
}));

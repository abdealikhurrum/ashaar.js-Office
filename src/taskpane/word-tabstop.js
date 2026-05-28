(function (root, factory) {
  if (typeof module !== "undefined" && module.exports) {
    module.exports = factory(require("../vendor/ashaar-justify"));
  } else {
    root.AshaarTabStop = factory(root.AshaarJustify);
  }
}(typeof globalThis !== "undefined" ? globalThis : this, function (AshaarJustify) {

  // US-Letter with 1-inch margins: 6.5" × 1440 twips/inch = 9360 twips.
  // Used as the fallback when no measured page width is available (tests, browser preview).
  var TEXT_WIDTH_DEFAULT = 9360;

  function escapeXml(s) {
    return String(s || "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  }

  function justifyMisra(text, opts, colWidthPx) {
    opts = opts || {};
    if (AshaarJustify && opts.justifyMode === "kashida" && opts._justifyCtx && colWidthPx > 0) {
      return AshaarJustify.justifyLine(text, colWidthPx, opts._justifyCtx, { targetFill: 0.92 });
    }
    var count = Number(opts.tatweelCount || 0);
    if (!AshaarJustify || opts.justifyMode !== "kashida" || count <= 0) return text;
    return AshaarJustify.spreadTatweels(text, count);
  }

  // Tab stop positions for N content columns (in twips from left margin).
  //
  // The paragraph is LTR with RTL text runs.  Column order (left → right):
  //   col 0 = ajuz side (leftmost visually), col N-1 = sadr side (rightmost).
  //
  // Col 0: no stop needed — text starts at the left margin automatically.
  // Cols 1..N-2: CENTER stops at each column's midpoint.
  // Col N-1: RIGHT stop at W (sadr right-aligns to the right margin).
  //
  // textWidth: measured text-area width in twips (pageWidth - leftMargin - rightMargin,
  //            all in points × 20).  Falls back to TEXT_WIDTH_DEFAULT when omitted.
  function tabStopsForN(N, textWidth) {
    var W = (textWidth > 0) ? textWidth : TEXT_WIDTH_DEFAULT;
    var colW = W / N;
    var stops = [];
    for (var i = 1; i < N; i++) {
      if (i === N - 1) {
        stops.push({ pos: W, val: "right" });
      } else {
        // Midpoint of column i: (i + 0.5) × colW = (2i+1) × colW / 2
        stops.push({ pos: Math.round((2 * i + 1) * colW / 2), val: "center" });
      }
    }
    return stops;
  }

  function tabStopsXml(stops) {
    if (!stops.length) return "";
    return "<w:tabs>" + stops.map(function (s) {
      return '<w:tab w:val="' + s.val + '" w:pos="' + s.pos + '"/>';
    }).join("") + "</w:tabs>";
  }

  function runPropsXml(opts, isRefrain) {
    var inner = "<w:rtl/>";
    var font = (opts || {}).fontMode;
    if (font === "nastaliq") inner += '<w:rFonts w:cs="Noto Nastaliq Urdu"/>';
    else if (font === "arabic-serif") inner += '<w:rFonts w:cs="Scheherazade New"/>';
    if (isRefrain) inner += '<w:color w:val="A7352A"/>';
    return "<w:rPr>" + inner + "</w:rPr>";
  }

  function textRun(text, opts, isRefrain, colWidthPx) {
    return "<w:r>" + runPropsXml(opts, isRefrain) +
      '<w:t xml:space="preserve">' + escapeXml(justifyMisra(text, opts, colWidthPx)) + "</w:t></w:r>";
  }

  function tabRun() { return "<w:r><w:tab/></w:r>"; }

  // Convert one bayt to a row descriptor.
  // Returns { solo:true, text, isRefrain }  — centered paragraph
  //      or { solo:false, cols:[{text,isRefrain}|null, …N] } — tab-stop paragraph
  //      or null for empty/skipped rows.
  function baytToRow(bayt, N) {
    var cols, i, j;

    if (bayt.type === "row") {
      var misras = bayt.misras || [];
      var K = misras.length;
      if (K === 0) return null;
      if (K === 1) return { solo: true, text: misras[0].text, isRefrain: !!misras[0].isRefrain };

      cols = [];
      for (i = 0; i < N; i++) cols.push(null);

      if (K >= N) {
        // Full N-misra row: col 0 = misras[N-1] (ajuz/left), col N-1 = misras[0] (sadr/right)
        for (i = 0; i < N; i++) {
          cols[i] = { text: misras[N - 1 - i].text, isRefrain: !!misras[N - 1 - i].isRefrain };
        }
      } else if (K === 2) {
        // Pair inside a wider stanza: span to the outer columns
        cols[0] = { text: misras[1].text, isRefrain: !!misras[1].isRefrain };
        cols[N - 1] = { text: misras[0].text, isRefrain: !!misras[0].isRefrain };
      } else {
        // Partial row (3 ≤ K < N): fill rightmost K content columns
        for (j = 0; j < K; j++) {
          cols[N - 1 - j] = { text: misras[j].text, isRefrain: !!misras[j].isRefrain };
        }
      }
      return { solo: false, cols: cols };
    }

    // type:'bayt'
    if (!bayt.ajuz) return { solo: true, text: bayt.sadr, isRefrain: !!bayt.sadrRefrain };

    cols = [];
    for (i = 0; i < N; i++) cols.push(null);
    cols[0] = { text: bayt.ajuz, isRefrain: !!bayt.ajuzRefrain };
    cols[N - 1] = { text: bayt.sadr, isRefrain: !!bayt.sadrRefrain };
    return { solo: false, cols: cols };
  }

  // Build OOXML runs for a column-based row.
  // Null columns contribute nothing but still advance the cursor via the inter-column tab.
  function buildRuns(cols, opts, colWidthPx) {
    var parts = [];
    for (var i = 0; i < cols.length; i++) {
      var col = cols[i];
      if (col) parts.push(textRun(col.text, opts, col.isRefrain, colWidthPx));
      if (i < cols.length - 1) parts.push(tabRun());
    }
    return parts.join("");
  }

  function spacingAttr(afterPt) {
    return '<w:spacing w:after="' + Math.round(afterPt * 20) + '"/>';
  }

  function columnParaXml(row, stopsXml, opts, afterPt, colWidthPx) {
    return "<w:p><w:pPr>" + stopsXml + spacingAttr(afterPt) + "</w:pPr>" +
      buildRuns(row.cols, opts, colWidthPx) + "</w:p>";
  }

  function soloParaXml(row, opts, afterPt, textWidthPx) {
    return '<w:p><w:pPr><w:jc w:val="center"/>' + spacingAttr(afterPt) + "</w:pPr>" +
      textRun(row.text, opts, row.isRefrain, textWidthPx) + "</w:p>";
  }

  // Convert poetry source text to OOXML paragraph markup (no document wrapper).
  // textWidth: actual text-area width in twips, read from Word's section page layout.
  //            Omit (or pass 0) to fall back to the US-Letter default.
  function poemToOoxml(source, opts, Ashaar, textWidth) {
    opts = opts || {};
    if (!Ashaar || typeof Ashaar.parse !== "function") throw new Error("Ashaar.js not loaded.");
    var poems = Ashaar.parse(String(source || ""));
    var paras = [];

    poems.forEach(function (poem, poemIdx) {
      poem.stanzas.forEach(function (stanza, stanzaIdx) {
        var isLastStanza = poemIdx === poems.length - 1 && stanzaIdx === poem.stanzas.length - 1;

        // N = max misra count in stanza; fall back to 2 for regular couplets
        var N = 2;
        stanza.bayts.forEach(function (b) {
          if (b.type === "row" && b.misras) N = Math.max(N, b.misras.length);
        });

        var stops = tabStopsForN(N, textWidth);
        var stopsXml = tabStopsXml(stops);
        var W = (textWidth > 0) ? textWidth : TEXT_WIDTH_DEFAULT;
        var textWidthPx = W / 20 * (96 / 72);
        var colWidthPx = textWidthPx / N;

        stanza.bayts.forEach(function (bayt, baytIdx) {
          var isLast = baytIdx === stanza.bayts.length - 1;
          // 8 pt gap between stanzas; no gap within a stanza or after the last stanza
          var afterPt = isLast ? (isLastStanza ? 0 : 8) : 0;
          var row = baytToRow(bayt, N);
          if (!row) return;
          paras.push(row.solo
            ? soloParaXml(row, opts, afterPt, textWidthPx)
            : columnParaXml(row, stopsXml, opts, afterPt, colWidthPx));
        });
      });
    });

    return paras.join("");
  }

  // Wrap paragraph markup in a minimal OOXML document for Range.insertOoxml().
  function wrapOoxml(bodyContent) {
    var ns = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<w:document ' + ns + '><w:body>' + bodyContent + '<w:sectPr/></w:body></w:document>';
  }

  return {
    poemToOoxml: poemToOoxml,
    wrapOoxml: wrapOoxml,
    tabStopsForN: tabStopsForN
  };
}));

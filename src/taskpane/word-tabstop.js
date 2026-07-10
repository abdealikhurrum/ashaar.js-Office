(function (root, factory) {
  if (typeof module !== "undefined" && module.exports) {
    module.exports = factory(require("../vendor/ashaar-justify"), require("./fonts"));
  } else {
    root.AshaarTabStop = factory(root.AshaarJustify, root.AshaarFonts);
  }
}(typeof globalThis !== "undefined" ? globalThis : this, function (AshaarJustify, AshaarFonts) {

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

  // Tab stop positions for N content columns in an RTL (<w:bidi/>) paragraph.
  //
  // The paragraph is RTL: cursor starts at the right margin (W) and moves leftward.
  // Column order in XML (first run → last run): sadr (col 0, rightmost) → ajuz (col N-1, leftmost).
  //
  // Col 0 (sadr): no stop — text starts at the right margin automatically.
  // Cols 1..N-2: CENTER stops at each column's midpoint going leftward.
  // Col N-1 (ajuz): LEFT stop at 0 so ajuz anchors to the left margin.
  //
  // For equal-width columns the CENTER stop positions are symmetric around W/2
  // and happen to equal the LTR midpoints; only the final stop changes
  // (RIGHT at W → LEFT at 0).
  //
  // textWidth: measured text-area width in twips. Falls back to TEXT_WIDTH_DEFAULT.
  function tabStopsForN(N, textWidth) {
    var W = (textWidth > 0) ? textWidth : TEXT_WIDTH_DEFAULT;
    var colW = W / N;
    var stops = [];
    if (N >= 2) stops.push({ pos: 0, val: "left" }); // ajuz anchors to left margin
    for (var i = N - 2; i >= 1; i--) {
      // Midpoint of column i (from the right): W - (i + 0.5) × colW
      stops.push({ pos: Math.round(W - (i + 0.5) * colW), val: "center" });
    }
    stops.sort(function (a, b) { return a.pos - b.pos; }); // ascending per OOXML convention
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
    var mode = (opts || {}).fontMode === "nastaliq" ? "noto" : (opts || {}).fontMode;
    var csName = AshaarFonts.wordNameOf(mode);
    if (csName) inner += '<w:rFonts w:cs="' + csName + '"/>';
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
        // Full N-misra row: col 0 = misras[0] (sadr/right in RTL), col N-1 = misras[N-1] (ajuz/left)
        for (i = 0; i < N; i++) {
          cols[i] = { text: misras[i].text, isRefrain: !!misras[i].isRefrain };
        }
      } else if (K === 2) {
        // Pair inside a wider stanza: sadr in col 0 (right), ajuz in col N-1 (left)
        cols[0] = { text: misras[0].text, isRefrain: !!misras[0].isRefrain };
        cols[N - 1] = { text: misras[1].text, isRefrain: !!misras[1].isRefrain };
      } else {
        // Partial row (3 ≤ K < N): fill leftmost K cols (= rightmost visually in RTL)
        for (j = 0; j < K; j++) {
          cols[j] = { text: misras[j].text, isRefrain: !!misras[j].isRefrain };
        }
      }
      return { solo: false, cols: cols };
    }

    // type:'bayt'
    if (!bayt.ajuz) return { solo: true, text: bayt.sadr, isRefrain: !!bayt.sadrRefrain };

    cols = [];
    for (i = 0; i < N; i++) cols.push(null);
    // RTL paragraph: sadr in col 0 (right margin), ajuz in col N-1 (left margin)
    cols[0] = { text: bayt.sadr, isRefrain: !!bayt.sadrRefrain };
    cols[N - 1] = { text: bayt.ajuz, isRefrain: !!bayt.ajuzRefrain };
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
    // Schema order: tabs(9) → bidi(17) → spacing(20)
    return "<w:p><w:pPr>" + stopsXml + "<w:bidi/>" + spacingAttr(afterPt) + "</w:pPr>" +
      buildRuns(row.cols, opts, colWidthPx) + "</w:p>";
  }

  function soloParaXml(row, opts, afterPt, textWidthPx) {
    // Schema order: bidi(17) → spacing(20) → jc(25)
    return '<w:p><w:pPr><w:bidi/>' + spacingAttr(afterPt) + '<w:jc w:val="center"/>' + "</w:pPr>" +
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

  // Wrap paragraph markup in a FlatOpc package for Range.insertOoxml().
  // The bare <w:document> wrapper is rejected by Word for Mac; the full pkg:package
  // format (same as getOoxml() returns) is required on all platforms.
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

  return {
    poemToOoxml: poemToOoxml,
    wrapOoxml: wrapOoxml,
    tabStopsForN: tabStopsForN
  };
}));

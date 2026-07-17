(function (root, factory) {
  if (typeof module !== "undefined" && module.exports) {
    module.exports = factory();
  } else {
    root.CiteWord = factory();
  }
}(typeof globalThis !== "undefined" ? globalThis : this, function () {
  "use strict";

  var KEEP = { i: 1, b: 1, em: 1, strong: 1, span: 1, sup: 1, sub: 1, br: 1 };

  // Strip all attributes from kept tags; unwrap (keep inner content of) all other tags.
  function sanitize(html) {
    if (!html) { return ""; }
    return String(html).replace(/<(\/?)([a-zA-Z0-9]+)([^>]*)>/g, function (m, slash, name) {
      var tag = name.toLowerCase();
      if (KEEP[tag]) { return "<" + slash + tag + ">"; }
      return ""; // unwrap: remove the tag, keep its text content
    });
  }

  function direction(rtl) { return rtl ? "Rtl" : "Ltr"; }

  // citeproc emits plain mixed-direction text with NO directional markup (it
  // delegates RTL to the rendering surface). In an LTR paragraph, Word's bidi
  // algorithm then mis-orders the neutral punctuation — ( ) , . — that sits
  // around Arabic runs (e.g. "(دار المعارف, 1951)" renders as ")دار المعارف, ).1951").
  // wrapRtlRuns supplies that guidance: it wraps each maximal Arabic run (its
  // letters plus the digits/neutrals that belong to it) in <span dir="rtl">, which
  // Word's HTML import maps to an RTL run so the punctuation resolves correctly.
  // Latin (LTR) runs are left untouched. HTML tags are treated as transparent so
  // an italicised Arabic title (<i>…</i>) is wrapped as one run.
  // Within each Arabic run it also localizes ASCII comma/semicolon to their Arabic
  // forms (،/؛) — Chicago (and most CSL styles) hard-code ASCII delimiters in the
  // style, so even under the `ar` locale the raw output carries Latin ", "/"; ".
  var RTL_CHAR = /[؀-ۿݐ-ݿࢠ-ࣿﭐ-﷿ﹰ-﻿]/;
  var LTR_CHAR = /[A-Za-zÀ-ɏ]/;
  var AR_PUNCT = { ",": "،", ";": "؛" };

  function wrapRtlRuns(html) {
    if (!html || !RTL_CHAR.test(html)) { return html || ""; }
    // Tokenise into items: whole HTML tags (opaque) or single characters.
    var items = [];
    var re = /<[^>]*>|[\s\S]/g;
    var m;
    while ((m = re.exec(html))) {
      var tok = m[0];
      if (tok.charAt(0) === "<" && tok.charAt(tok.length - 1) === ">") { items.push({ tag: tok }); }
      else { items.push({ ch: tok }); }
    }
    var n = items.length;
    var out = [];
    var i = 0;
    while (i < n) {
      var it = items[i];
      if (!it.tag && RTL_CHAR.test(it.ch)) {
        // Start of an RTL run at the first Arabic char; extend forward over
        // Arabic + digits + neutrals + tags, stopping before the next Latin letter.
        var start = i;
        var lastRtl = i;
        var k = i + 1;
        while (k < n) {
          var t = items[k];
          if (t.tag) { k++; continue; }
          if (LTR_CHAR.test(t.ch)) { break; }
          if (RTL_CHAR.test(t.ch)) { lastRtl = k; }
          k++;
        }
        var end = k; // exclusive
        // If we stopped at a Latin letter, trailing whitespace before it belongs
        // to the outer flow, not the Arabic clause — trim it back to the last Arabic char.
        if (k < n) {
          while (end - 1 > lastRtl && !items[end - 1].tag && /\s/.test(items[end - 1].ch)) { end--; }
        }
        out.push('<span dir="rtl">');
        for (var j = start; j < end; j++) {
          if (items[j].tag) { out.push(items[j].tag); }
          else { out.push(AR_PUNCT[items[j].ch] || items[j].ch); }
        }
        out.push("</span>");
        i = end;
      } else {
        out.push(it.tag ? it.tag : it.ch);
        i++;
      }
    }
    return out.join("");
  }

  function buildNotePayload(o) {
    return { html: wrapRtlRuns(sanitize(o.html)), direction: direction(o.rtl) };
  }

  function buildBibliographyPayload(o) {
    return { html: wrapRtlRuns(sanitize(o.html)), direction: direction(o.rtl), tag: o.tag || "AshaarBibliography" };
  }

  function citationTag(itemKeys, style) {
    return "AshaarCite:" + style + ":" + (itemKeys || []).join(",");
  }

  function parseCitationTag(tag) {
    var m = /^AshaarCite:([^:]*):(.*)$/.exec(tag || "");
    if (!m) { return null; }
    return { style: m[1], itemKeys: m[2] ? m[2].split(",") : [] };
  }

  return {
    sanitize: sanitize,
    wrapRtlRuns: wrapRtlRuns,
    buildNotePayload: buildNotePayload,
    buildBibliographyPayload: buildBibliographyPayload,
    citationTag: citationTag,
    parseCitationTag: parseCitationTag
  };
}));

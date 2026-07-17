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
    return { html: wrapRtlRuns(sanitize(o.html)), direction: direction(o.rtl) };
  }

  // UTF-8-safe base64 that works in Node (Buffer) and the Word WebView (btoa).
  function b64encode(str) {
    if (typeof Buffer !== "undefined") { return Buffer.from(str, "utf8").toString("base64"); }
    return btoa(unescape(encodeURIComponent(str)));
  }
  function b64decode(b64) {
    if (typeof Buffer !== "undefined") { return Buffer.from(b64, "base64").toString("utf8"); }
    return decodeURIComponent(escape(atob(b64)));
  }

  function buildCitationTag(o) {
    var payload = {
      v: 1,
      style: o.style,
      locale: o.locale,
      keys: (o.items || []).map(function (i) {
        return { id: i.id, locator: i.locator || null, label: i.label || null };
      })
    };
    return "AshaarCite:" + b64encode(JSON.stringify(payload));
  }

  function parseCitationTag(tag) {
    var s = String(tag || "");
    if (s.indexOf("AshaarCite:") !== 0) { return null; }
    try {
      var obj = JSON.parse(b64decode(s.slice("AshaarCite:".length)));
      if (!obj || !Array.isArray(obj.keys)) { return null; }
      return obj;
    } catch (e) { return null; }
  }

  function buildBibliographyTag(o) {
    return "AshaarBib:" + b64encode(JSON.stringify({ v: 1, style: o.style, locale: o.locale }));
  }

  // Maps a parsed AshaarCite: tag's keys back to the {id, locator?, label?}
  // shape engine.cite() expects, dropping null locator/label. Tolerates
  // missing/null parsed input (returns []).
  function citationItemsFromTag(parsed) {
    var keys = (parsed && parsed.keys) || [];
    return keys.map(function (k) {
      var it = { id: k.id };
      if (k.locator) { it.locator = k.locator; }
      if (k.label) { it.label = k.label; }
      return it;
    });
  }

  function xmlEsc(s) {
    return String(s).replace(/&(?!amp;|lt;|gt;)/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  }

  // Tokenise into HTML tags (opaque) and single characters — same shape wrapRtlRuns uses.
  function tokenizeHtml(html) {
    var items = [], re = /<[^>]*>|[\s\S]/g, m;
    while ((m = re.exec(html || ""))) {
      var tok = m[0];
      if (tok.charAt(0) === "<" && tok.charAt(tok.length - 1) === ">") { items.push({ tag: tok }); }
      else { items.push({ ch: tok }); }
    }
    return items;
  }

  // Per-char RTL-run membership, mirroring wrapRtlRuns: an Arabic char plus the
  // trailing digits/neutrals/tags up to (but not including) the next Latin letter.
  function computeRtlFlags(items) {
    var n = items.length, flags = [], i;
    for (i = 0; i < n; i++) { flags[i] = false; }
    i = 0;
    while (i < n) {
      if (!items[i].tag && RTL_CHAR.test(items[i].ch)) {
        var start = i, lastRtl = i, k = i + 1;
        while (k < n) {
          var t = items[k];
          if (t.tag) { k++; continue; }
          if (LTR_CHAR.test(t.ch)) { break; }
          if (RTL_CHAR.test(t.ch)) { lastRtl = k; }
          k++;
        }
        var end = k;
        if (k < n) { while (end - 1 > lastRtl && !items[end - 1].tag && /\s/.test(items[end - 1].ch)) { end--; } }
        for (var j = start; j < end; j++) { if (!items[j].tag) { flags[j] = true; } }
        i = end;
      } else { i++; }
    }
    return flags;
  }

  function emitRun(text, sig, csFont) {
    var rpr = "<w:rPr>";
    if (sig.rtl) { rpr += "<w:rtl/>"; if (csFont) { rpr += '<w:rFonts w:cs="' + csFont + '"/>'; } }
    if (sig.b) { rpr += "<w:b/>"; if (sig.rtl) { rpr += "<w:bCs/>"; } }
    if (sig.i && !sig.rtl) { rpr += "<w:i/>"; } // italic suppressed on Arabic runs
    if (sig.sup) { rpr += '<w:vertAlign w:val="superscript"/>'; }
    rpr += "</w:rPr>";
    return "<w:r>" + rpr + '<w:t xml:space="preserve">' + xmlEsc(text) + "</w:t></w:r>";
  }

  function htmlToOoxmlRuns(html, opts) {
    var csFont = (opts && opts.csFont) || "";
    var items = tokenizeHtml(html);
    var rtl = computeRtlFlags(items);
    var out = [], buf = "", cur = null, fmt = { i: 0, b: 0, sup: 0 };
    function flush() { if (buf !== "" && cur) { out.push(emitRun(buf, cur, csFont)); } buf = ""; }
    for (var idx = 0; idx < items.length; idx++) {
      var it = items[idx];
      if (it.tag) {
        var tg = it.tag.toLowerCase().replace(/\s+/g, "");
        if (/^<(i|em)>$/.test(tg)) { flush(); fmt.i++; }
        else if (/^<\/(i|em)>$/.test(tg)) { flush(); fmt.i = Math.max(0, fmt.i - 1); }
        else if (/^<(b|strong)>$/.test(tg)) { flush(); fmt.b++; }
        else if (/^<\/(b|strong)>$/.test(tg)) { flush(); fmt.b = Math.max(0, fmt.b - 1); }
        else if (/^<sup>$/.test(tg)) { flush(); fmt.sup++; }
        else if (/^<\/sup>$/.test(tg)) { flush(); fmt.sup = Math.max(0, fmt.sup - 1); }
        else if (/^<br\/?>$/.test(tg)) { flush(); out.push("<w:r><w:br/></w:r>"); }
        // span/sub and any other tag: transparent (no format change)
        continue;
      }
      var sig = { rtl: rtl[idx], i: fmt.i > 0, b: fmt.b > 0, sup: fmt.sup > 0 };
      if (!cur || cur.rtl !== sig.rtl || cur.i !== sig.i || cur.b !== sig.b || cur.sup !== sig.sup) {
        flush(); cur = sig;
      }
      var ch = it.ch;
      if (sig.rtl && AR_PUNCT[ch]) { ch = AR_PUNCT[ch]; }
      buf += ch;
    }
    flush();
    return out.join("");
  }

  function buildCitationParagraphOoxml(html, opts) {
    return '<w:p><w:pPr><w:bidi/><w:jc w:val="right"/></w:pPr>' + htmlToOoxmlRuns(html, opts) + "</w:p>";
  }

  return {
    sanitize: sanitize,
    wrapRtlRuns: wrapRtlRuns,
    buildNotePayload: buildNotePayload,
    buildBibliographyPayload: buildBibliographyPayload,
    buildCitationTag: buildCitationTag,
    parseCitationTag: parseCitationTag,
    buildBibliographyTag: buildBibliographyTag,
    citationItemsFromTag: citationItemsFromTag,
    htmlToOoxmlRuns: htmlToOoxmlRuns,
    buildCitationParagraphOoxml: buildCitationParagraphOoxml
  };
}));

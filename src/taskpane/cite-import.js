(function (root, factory) {
  if (typeof module !== "undefined" && module.exports) {
    module.exports = factory();
  } else {
    root.CiteImport = factory();
  }
}(typeof globalThis !== "undefined" ? globalThis : this, function () {
  "use strict";

  // Detect the format of imported text. Filename extension wins; otherwise sniff
  // the content. Returns "csljson" | "bibtex" | "ris" | null.
  function sniffFormat(text, filename) {
    var fn = String(filename || "").toLowerCase();
    if (/\.json$/.test(fn)) { return "csljson"; }
    if (/\.(bib|bibtex)$/.test(fn)) { return "bibtex"; }
    if (/\.ris$/.test(fn)) { return "ris"; }
    var s = String(text || "").replace(/^﻿/, "").trim();
    if (s.charAt(0) === "[" || s.charAt(0) === "{") { return "csljson"; }
    if (/^TY\s*-\s/m.test(s)) { return "ris"; }
    if (/@[A-Za-z]+\s*\{/.test(s)) { return "bibtex"; }
    return null;
  }

  // Parse imported text into an array of CSL-JSON items. MVP supports CSL JSON
  // (from Zotero's "Better CSL JSON" — drag, file, or paste). BibTeX/RIS are
  // detected but not yet converted; they throw a clear message rather than
  // silently importing nothing.
  function parseImport(text, format) {
    var fmt = format || sniffFormat(text);
    if (fmt === "csljson") {
      var data;
      try { data = JSON.parse(String(text)); }
      catch (e) { throw new Error("That doesn't look like valid CSL JSON."); }
      var arr = Array.isArray(data) ? data : (data && typeof data === "object" ? [data] : []);
      var out = [];
      for (var i = 0; i < arr.length; i++) {
        var it = arr[i];
        if (!it || typeof it !== "object") { continue; }
        if (it.id === undefined || it.id === null || it.id === "") { it.id = "import-" + (i + 1); }
        out.push(it);
      }
      return out;
    }
    if (fmt === "bibtex" || fmt === "ris") {
      throw new Error("BibTeX/RIS import is coming soon — for now export as \"Better CSL JSON\" from Zotero.");
    }
    throw new Error("Unrecognized format — expected CSL JSON.");
  }

  return { sniffFormat: sniffFormat, parseImport: parseImport };
}));

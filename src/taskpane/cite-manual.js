(function (root, factory) {
  if (typeof module !== "undefined" && module.exports) {
    module.exports = factory();
  } else {
    root.CiteManual = factory();
  }
}(typeof globalThis !== "undefined" ? globalThis : this, function () {
  "use strict";

  var TYPE_MAP = { book: "book", chapter: "chapter", article: "article-journal", webpage: "webpage" };

  // "Family, Given" per line -> CSL name objects. A line without a comma is an
  // organization / single-name literal. Blank lines are skipped.
  function parseNames(text) {
    var out = [];
    var lines = String(text || "").split(/\r?\n/);
    for (var i = 0; i < lines.length; i++) {
      var line = lines[i].trim();
      if (!line) { continue; }
      var comma = line.indexOf(",");
      if (comma === -1) { out.push({ literal: line }); continue; }
      var family = line.slice(0, comma).trim();
      var given = line.slice(comma + 1).trim();
      var name = { family: family };
      if (given) { name.given = given; }
      out.push(name);
    }
    return out;
  }

  // "yyyy" | "yyyy-mm" | "yyyy-mm-dd" -> CSL date-parts [[y]] / [[y,m]] / [[y,m,d]].
  // null when there is no leading numeric year.
  function parseDateParts(str) {
    var s = String(str || "").trim();
    if (!/^\d{1,4}/.test(s)) { return null; }
    var parts = s.split("-");
    var nums = [];
    for (var i = 0; i < parts.length; i++) {
      var n = parseInt(parts[i], 10);
      if (isNaN(n)) { break; }
      nums.push(n);
    }
    return nums.length ? [nums] : null;
  }

  function setStr(item, key, val) {
    var v = (val === undefined || val === null) ? "" : String(val).trim();
    if (v) { item[key] = v; }
  }

  // Build a CSL-JSON item from flat form values. Empty fields are omitted.
  function buildManualItem(values) {
    var v = values || {};
    var item = { id: v.id, type: TYPE_MAP[v.type] || "document" };
    setStr(item, "title", v.title);
    setStr(item, "container-title", v.containerTitle);
    setStr(item, "publisher", v.publisher);
    setStr(item, "publisher-place", v.place);
    setStr(item, "volume", v.volume);
    setStr(item, "issue", v.issue);
    setStr(item, "page", v.pages);
    setStr(item, "URL", v.url);
    var issued = parseDateParts(v.year);
    if (issued) { item.issued = { "date-parts": issued }; }
    var accessed = parseDateParts(v.accessed);
    if (accessed) { item.accessed = { "date-parts": accessed }; }
    var authors = parseNames(v.authors);
    if (authors.length) { item.author = authors; }
    var editors = parseNames(v.editors);
    if (editors.length) { item.editor = editors; }
    return item;
  }

  return { parseNames: parseNames, parseDateParts: parseDateParts, buildManualItem: buildManualItem };
}));

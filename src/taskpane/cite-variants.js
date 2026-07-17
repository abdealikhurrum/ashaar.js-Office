(function (root, factory) {
  if (typeof module !== "undefined" && module.exports) {
    module.exports = factory();
  } else {
    root.CiteVariants = factory();
  }
}(typeof globalThis !== "undefined" ? globalThis : this, function () {
  "use strict";

  // Invisible bidi control chars (RLE/LRE/PDF/LRM/RLM/LRO/RLO) that Juris-M and
  // some editors embed around Arabic strings — strip from variant values.
  var BIDI = /[‎‏‪‫‬‭‮]/g;
  var VARIANT_TAGS = { romanized: "ar-Latn", translated: "en" }; // 'original' -> skip (real field holds it)

  function stripBidi(s) { return String(s == null ? "" : s).replace(BIDI, "").trim(); }

  // Parse one "cne-<rest>" key (prefix already removed) + value into the
  // normalized intermediate, mutating `out`. Ignores unknown-variant lines.
  function addCneKey(out, rest, value) {
    var segs = rest.split("-");
    if (segs.length < 2) { return; }
    var variant = segs[segs.length - 1];
    var tag = VARIANT_TAGS[variant];
    if (!tag) { return; } // unknown/original variant -> ignore
    var body = segs.slice(0, segs.length - 1); // field or creator body

    // creator line: contains a numeric segment followed by last|first
    var numIdx = -1;
    for (var i = 0; i < body.length; i++) { if (/^\d+$/.test(body[i])) { numIdx = i; break; } }
    if (numIdx !== -1 && numIdx + 1 < body.length &&
        (body[numIdx + 1] === "last" || body[numIdx + 1] === "first")) {
      var creatorType = body.slice(0, numIdx).join("-");
      var index = body[numIdx];
      var part = body[numIdx + 1];
      out.creators[creatorType] = out.creators[creatorType] || {};
      out.creators[creatorType][index] = out.creators[creatorType][index] || {};
      out.creators[creatorType][index][tag] = out.creators[creatorType][index][tag] || {};
      out.creators[creatorType][index][tag][part === "last" ? "family" : "given"] = stripBidi(value);
      return;
    }

    // simple field
    var field = body.join("-");
    out.fields[field] = out.fields[field] || {};
    out.fields[field][tag] = stripBidi(value);
  }

  function parseCne(text) {
    var s = String(text || "");
    if (s.indexOf("cne-") === -1) { return null; }
    var out = { fields: {}, creators: {} };
    var lines = s.split(/\r?\n/);
    var seen = false;
    for (var i = 0; i < lines.length; i++) {
      var m = /^\s*cne-([^:]+):\s*([\s\S]*)$/.exec(lines[i]);
      if (!m) { continue; }
      var before = JSON.stringify(out);
      addCneKey(out, m[1].trim(), m[2]);
      if (JSON.stringify(out) !== before) { seen = true; }
    }
    return seen ? out : null;
  }

  // Shallow clone + attach multi models from parsed cne-* variants.
  function applyVariantsToItem(item) {
    if (!item || typeof item !== "object") { return item; }
    var parsed = parseCne(item.note);
    if (!parsed) { return item; }

    var out = {};
    var k;
    for (k in item) { if (Object.prototype.hasOwnProperty.call(item, k)) { out[k] = item[k]; } }

    // fields
    var keys = {};
    var main = {};
    for (var f in parsed.fields) {
      if (Object.prototype.hasOwnProperty.call(parsed.fields, f)) {
        keys[f] = parsed.fields[f];
        main[f] = item.language || "ar";
      }
    }
    if (Object.keys(keys).length) { out.multi = { main: main, _keys: keys }; }

    // creators — clone the target creator array + entry before attaching multi
    for (var cv in parsed.creators) {
      if (!Object.prototype.hasOwnProperty.call(parsed.creators, cv)) { continue; }
      if (!Array.isArray(out[cv])) { continue; }
      out[cv] = out[cv].slice();
      var byIdx = parsed.creators[cv];
      for (var idx in byIdx) {
        if (!Object.prototype.hasOwnProperty.call(byIdx, idx)) { continue; }
        var i = parseInt(idx, 10);
        if (!out[cv][i]) { continue; }
        var c = {};
        for (var ck in out[cv][i]) { if (Object.prototype.hasOwnProperty.call(out[cv][i], ck)) { c[ck] = out[cv][i][ck]; } }
        c.multi = { main: item.language || "ar", _key: byIdx[idx] };
        out[cv][i] = c;
      }
    }
    return out;
  }

  function enrichItemMap(items) {
    var out = {};
    for (var id in items) {
      if (Object.prototype.hasOwnProperty.call(items, id)) { out[id] = applyVariantsToItem(items[id]); }
    }
    return out;
  }

  return {
    parseCne: parseCne,
    stripBidi: stripBidi,
    applyVariantsToItem: applyVariantsToItem,
    enrichItemMap: enrichItemMap
  };
}));

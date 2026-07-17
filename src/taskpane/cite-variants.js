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

  return { parseCne: parseCne, stripBidi: stripBidi };
}));

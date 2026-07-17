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

  // --- migration converter (used only by scripts/migrate-mlzsync-to-cne.mjs) ---

  function parseMlzsync(text) {
    var s = String(text || "");
    var at = s.indexOf("mlzsync1:");
    if (at === -1) { return null; }
    var rest = s.slice(at + "mlzsync1:".length);
    // 4-digit zero-padded length prefix, then JSON of that length.
    var m = /^(\d{4})/.exec(rest);
    if (!m) { return null; }
    var len = parseInt(m[1], 10);
    var json = rest.slice(4, 4 + len);
    var blob;
    try { blob = JSON.parse(json); } catch (e) {
      try { blob = JSON.parse(rest.slice(4)); } catch (e2) { return null; }
    }
    var out = { fields: {}, creators: {} };
    var mf = (blob && blob.multifields && blob.multifields._keys) || {};
    for (var f in mf) {
      if (!Object.prototype.hasOwnProperty.call(mf, f)) { continue; }
      out.fields[f] = {};
      for (var tag in mf[f]) {
        if (Object.prototype.hasOwnProperty.call(mf[f], tag)) { out.fields[f][tag] = stripBidi(mf[f][tag]); }
      }
    }
    var mc = (blob && blob.multicreators) || {};
    for (var idx in mc) {
      if (!Object.prototype.hasOwnProperty.call(mc, idx)) { continue; }
      var keyObj = mc[idx]._key || {};
      out.creators[idx] = {};
      for (var t in keyObj) {
        if (!Object.prototype.hasOwnProperty.call(keyObj, t)) { continue; }
        var nm = keyObj[t];
        var v = {};
        if (nm.lastName) { v.family = stripBidi(nm.lastName); }
        if (nm.firstName) { v.given = stripBidi(nm.firstName); }
        out.creators[idx][t] = v;
      }
    }
    return out;
  }

  function mlzsyncToCneLines(parsed, creators) {
    var lines = [];
    if (!parsed) { return lines; }
    // fields: any source tag -> romanized (mlzsync 'en' holds transliteration)
    var fnames = Object.keys(parsed.fields).sort();
    fnames.forEach(function (f) {
      var byTag = parsed.fields[f];
      var tags = Object.keys(byTag).sort();
      if (tags.length) { lines.push("cne-" + f + "-romanized: " + byTag[tags[0]]); }
    });
    // creators: resolve flat index -> creatorType + within-type index
    var typeCount = {};
    (creators || []).forEach(function (c, flat) {
      var type = (c && c.creatorType) || "author";
      var within = typeCount[type] || 0;
      typeCount[type] = within + 1;
      var byTag = parsed.creators[String(flat)];
      if (!byTag) { return; }
      var tags = Object.keys(byTag).sort();
      if (!tags.length) { return; }
      var nm = byTag[tags[0]];
      if (nm.family) { lines.push("cne-" + type + "-" + within + "-last-romanized: " + nm.family); }
      if (nm.given) { lines.push("cne-" + type + "-" + within + "-first-romanized: " + nm.given); }
    });
    return lines;
  }

  var SEGMENTS = ["persons", "institutions", "titles", "journals", "publishers", "places", "number", "title-short"];

  function variantToLangPrefs(variant) {
    var slots;
    if (variant === "translit") { slots = ["translit"]; }
    else if (variant === "both") { slots = ["orig", "translit"]; }
    else { return null; } // orig / unknown -> no override
    var lp = { translit: ["ar-Latn"], translat: ["en"] };
    for (var i = 0; i < SEGMENTS.length; i++) { lp[SEGMENTS[i]] = slots.slice(); }
    return lp;
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
    enrichItemMap: enrichItemMap,
    variantToLangPrefs: variantToLangPrefs,
    parseMlzsync: parseMlzsync,
    mlzsyncToCneLines: mlzsyncToCneLines
  };
}));

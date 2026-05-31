(function (root, factory) {
  if (typeof module !== "undefined" && module.exports) {
    module.exports = factory();
  } else {
    root.AshaarSeparators = factory();
  }
}(typeof globalThis !== "undefined" ? globalThis : this, function () {
  var CANONICAL = " \\ ";
  var POEM_MARKER = /^(?:---|—|–)$/;

  // Candidate hemistich separators, in tiebreak-priority order (explicit poetry
  // separators first; the ambiguous wide-gap rule last). Each `re` is global, used
  // for both detection and replacement.
  var CANDIDATES = [
    { id: "backslash", re: /\s*\\\s*/g, minRatio: 0.5 },
    { id: "asterisk",  re: /\s*\*\s*/g, minRatio: 0.5 },
    { id: "pipe",      re: /\s*\|\s*/g, minRatio: 0.5 },
    { id: "dash",      re: /\s+[-–—]\s+/g, minRatio: 0.5 }, // spaced dash only (not intra-word hyphens)
    { id: "tab",       re: /\t+/g, minRatio: 0.5 },
    { id: "spaces",    re: / {2,}/g, minRatio: 0.6 }        // wide gap — stricter to avoid false positives
  ];

  function escapeRegex(s) {
    return String(s == null ? "" : s).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  }

  function contentLines(lines) {
    return lines.filter(function (l) {
      var t = l.trim();
      return t && !POEM_MARKER.test(t);
    });
  }

  // Pick the separator that best explains the text, or null if none is confident.
  function detectSeparator(text) {
    var lines = String(text == null ? "" : text).split(/\r\n|\r|\n/);
    var content = contentLines(lines);
    if (!content.length) return null;
    var best = null;
    CANDIDATES.forEach(function (c) {
      var count = content.filter(function (l) { return new RegExp(c.re.source).test(l); }).length;
      if (!count || count / content.length < c.minRatio) return;
      // Higher match count wins; ties broken by CANDIDATES order (earlier = preferred).
      if (!best || count > best.count) best = { id: c.id, count: count };
    });
    return best ? { id: best.id } : null;
  }

  function candidateById(id) {
    for (var i = 0; i < CANDIDATES.length; i++) if (CANDIDATES[i].id === id) return CANDIDATES[i];
    return null;
  }

  function normalizeLine(line, re) {
    var t = line.trim();
    if (!t || POEM_MARKER.test(t)) return line; // preserve blank lines and poem markers
    return line
      .replace(new RegExp(re.source, "g"), CANONICAL)
      .replace(/ {2,}/g, " ")
      .trim();
  }

  // Pair consecutive content lines into couplets, respecting stanza/poem breaks.
  function pairLineMode(text) {
    var lines = String(text == null ? "" : text).split(/\r\n|\r|\n/);
    var out = [];
    var buf = [];
    function flush() {
      for (var i = 0; i < buf.length; i += 2) {
        out.push(i + 1 < buf.length ? buf[i] + CANONICAL + buf[i + 1] : buf[i]);
      }
      buf = [];
    }
    lines.forEach(function (line) {
      var t = line.trim();
      if (!t || POEM_MARKER.test(t)) { flush(); out.push(line); }
      else buf.push(t);
    });
    flush();
    return out.join("\n");
  }

  // Normalize hemistich separators to the canonical "\".
  //   opts.separator: "auto" (default) | candidate id | "custom"
  //   opts.customPattern: literal separator string (used when separator === "custom")
  //   opts.pairLines: if true, ignore separators and pair consecutive lines
  // Returns { text, detected, changed }.
  function normalizeSeparators(text, opts) {
    opts = opts || {};
    // Normalize line endings first (Word's Range.text uses CR) so detection,
    // pairing, and the `changed` flag all work on canonical LF text.
    var input = String(text == null ? "" : text).replace(/\r\n?/g, "\n");

    if (opts.pairLines) {
      var paired = pairLineMode(input);
      return { text: paired, detected: "pairLines", changed: paired !== input };
    }

    var re = null;
    var detected = null;
    var sep = opts.separator || "auto";

    if (sep === "custom") {
      var pat = escapeRegex(opts.customPattern);
      if (!pat) return { text: input, detected: null, changed: false };
      re = new RegExp("\\s*" + pat + "\\s*", "g");
      detected = "custom";
    } else if (sep === "auto") {
      var found = detectSeparator(input);
      if (!found) return { text: input, detected: null, changed: false };
      re = candidateById(found.id).re;
      detected = found.id;
    } else {
      var c = candidateById(sep);
      if (!c) return { text: input, detected: null, changed: false };
      re = c.re;
      detected = sep;
    }

    var result = input.split(/\r\n|\r|\n/).map(function (line) {
      return normalizeLine(line, re);
    }).join("\n");

    return { text: result, detected: detected, changed: result !== input };
  }

  return {
    normalizeSeparators: normalizeSeparators,
    detectSeparator: detectSeparator
  };
}));

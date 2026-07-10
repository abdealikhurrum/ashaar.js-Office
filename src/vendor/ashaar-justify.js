(function (root, factory) {
  // Expose the module for CommonJS and browser usage.
  if (typeof module !== 'undefined' && module.exports) {
    module.exports = factory();
  } else {
    root.AshaarJustify = factory();
  }
}(typeof globalThis !== 'undefined' ? globalThis : this, function () {

  // The tatweel is the Arabic stretching character used for justification.
  var TATWEEL = 'ـ';
  var ZWNJ = 0x200C;
  var DEFAULT_PRIORITY = 7;

  // Tatweels are elongation, not content. Strip them at every public entry so
  // justification always re-derives from the bare line: it is idempotent and
  // reducible — re-justify any number of times and the count follows the target
  // width / font / fill rather than compounding on previously inserted tatweels.
  function stripTatweels(text) {
    return String(text == null ? '' : text).replace(/ـ/g, '');
  }

  // Letters that can connect from the previous character but not onward.
  var RIGHT_JOIN = (function () {
    var s = {};
    [0x0622,0x0623,0x0624,0x0625,0x0627,
     0x062F,0x0630,
     0x0631,0x0632,0x0698,
     0x0648,0x06C1,0x06C3,0x06BA,
     0x0671,0x0672,0x0673,0x0675,0x0677,
     0x06D5].forEach(function (cp) { s[cp] = 1; });
    return s;
  }());

  // Bare hamza does not connect on either side.
  var NON_JOIN = { 0x0621: 1 };

  // Lam + alef forms a required ligature and must never be split by tatweel.
  var ALEF_VARIANTS = {
    0x0627: 1, 0x0622: 1, 0x0623: 1, 0x0624: 1, 0x0625: 1
  };

  // Arabic combining marks (harakat) should not be treated as letters.
  function isArabicMark(cp) {
    return (cp >= 0x0610 && cp <= 0x061A) ||
      (cp >= 0x064B && cp <= 0x065F) ||
      cp === 0x0670 ||
      (cp >= 0x06D6 && cp <= 0x06ED);
  }

  // ZWNJ is a hard connection boundary in Persian/Arabic shaping.
  function isZwnj(cp) {
    return cp === ZWNJ;
  }

  function isArabicLetter(cp) {
    return (cp >= 0x0621 && cp <= 0x063A) ||
      (cp >= 0x0641 && cp <= 0x064A) ||
      (cp >= 0x066E && cp <= 0x066F) ||
      (cp >= 0x0671 && cp <= 0x06D3) ||
      cp === 0x06D5;
  }

  function isJoiningLetter(cp) {
    return isArabicLetter(cp) && !isArabicMark(cp) && !isZwnj(cp) && !NON_JOIN[cp];
  }

  // Dual-joining letters are the ones that can connect onward to the left.
  function isDualJoining(cp) {
    return isJoiningLetter(cp) && !RIGHT_JOIN[cp];
  }

  // Returns true for alef variants used in lam-alef ligature detection.
  function isAlifVariant(cp) {
    return !!ALEF_VARIANTS[cp];
  }

  function getLetters(word) {
    var letters = [];
    for (var i = 0; i < word.length; i++) {
      var cp = word.charCodeAt(i);
      if (isJoiningLetter(cp)) letters.push({ index: i, cp: cp });
    }
    return letters;
  }

  function onlyMarksBetween(word, start, end) {
    for (var i = start + 1; i < end; i++) {
      var cp = word.charCodeAt(i);
      if (isArabicMark(cp)) continue;
      return false;
    }
    return true;
  }

  function connectsToNext(word, current, next) {
    return isDualJoining(current.cp) &&
      isJoiningLetter(next.cp) &&
      onlyMarksBetween(word, current.index, next.index);
  }

  // Detects lam + alef variants that should not receive an extra stretch.
  function isLamAlefSequence(prevCp, nextCp) {
    return prevCp === 0x0644 && isAlifVariant(nextCp);
  }

  // Allah in the common spelling لله uses a special shaping sequence.
  function isAllahWord(letters) {
    return letters.length === 3 &&
      letters[0].cp === 0x0644 &&
      letters[1].cp === 0x0644 &&
      letters[2].cp === 0x0647;
  }

  function canInsertTatweel(word, current, next) {
    if (!connectsToNext(word, current, next)) return false;
    if (isLamAlefSequence(current.cp, next.cp)) return false;
    return true;
  }

  function insertionPosition(next) {
    return next.index;
  }

  // Discover all legal tatweel insertion points for a single word.
  function tatweelSlots(word) {
    var slots = [];
    var letters = getLetters(word);
    if (isAllahWord(letters)) return slots;
    for (var i = 0; i < letters.length - 1; i++) {
      if (canInsertTatweel(word, letters[i], letters[i + 1])) {
        slots.push({ pos: insertionPosition(letters[i + 1]), priority: DEFAULT_PRIORITY });
      }
    }
    return slots;
  }

  // Spread a fixed number of tatweels across the legal slots.
  function spreadTatweels(text, n) {
    text = stripTatweels(text);
    if (n <= 0) return text;
    var words = text.split(' ');
    var allSlots = [];
    words.forEach(function (w, wi) {
      tatweelSlots(w).forEach(function (s) {
        allSlots.push({ wi: wi, pos: s.pos, priority: s.priority });
      });
    });
    if (!allSlots.length) return text;
    allSlots.sort(function (a, b) {
      if (b.priority !== a.priority) return b.priority - a.priority;
      return a.wi - b.wi || a.pos - b.pos;
    });

    // Rotate through the sorted slots so extra tatweels do not pile up on the same
    // position before other valid insertions have been used.
    var slotOrder = allSlots.slice();
    var chosen = [];
    for (var i = 0; i < n; i++) {
      chosen.push(slotOrder[0]);
      slotOrder.push(slotOrder.shift());
    }

    // Aggregate insertions by word and character position.
    var insertMap = {};
    chosen.forEach(function (s) {
      var key = s.wi + ':' + s.pos;
      insertMap[key] = (insertMap[key] || 0) + 1;
    });

    return words.map(function (w, wi) {
      var chars = w.split('');
      var offset = 0;
      var ins = [];
      for (var key in insertMap) {
        var kp = key.split(':');
        if (+kp[0] === wi) ins.push({ pos: +kp[1], count: insertMap[key] });
      }
      ins.sort(function (a, b) { return a.pos - b.pos; });
      ins.forEach(function (entry) {
        var tatweels = new Array(entry.count + 1).join(TATWEEL).split('');
        Array.prototype.splice.apply(chars, [entry.pos + offset, 0].concat(tatweels));
        offset += tatweels.length;
      });
      return chars.join('');
    }).join(' ');
  }

  // Build a ranked list of insertion slots for a line, optionally using font quality.
  function buildSlots(text, params, fontProfile) {
    params = params || {};
    var words = text.split(' ');
    var slots = [];
    words.forEach(function (w, wi) {
      var letters = getLetters(w);
      if (isAllahWord(letters)) return;
      for (var i = 0; i < letters.length - 1; i++) {
        var current = letters[i];
        var next = letters[i + 1];
        if (!canInsertTatweel(w, current, next)) continue;
        var base = DEFAULT_PRIORITY;
        var bonus = 0;
        if (fontProfile) {
          var q = fontProfile.getQuality(w[current.index], w[next.index]);
          bonus = (q - 0.5) * (params.fontQualityBoost || 0);
        }
        slots.push({ wi: wi, pos: insertionPosition(next), score: base + bonus });
      }
    });
    slots.sort(function (a, b) { return b.score - a.score; });
    return slots;
  }

  // Splice fill characters into one text. insertMap maps "wi:pos" -> count,
  // where wi is the word index (split on ' ') and pos is the char offset within
  // that word. Shared by applySlots (single string) and applySlotsMulti (runs).
  function insertIntoWords(text, insertMap, fillChar) {
    var words = String(text).split(' ');
    return words.map(function (w, wi) {
      var chars = w.split('');
      var ins = [];
      for (var key in insertMap) {
        var kp = key.split(':');
        if (+kp[0] === wi) ins.push({ pos: +kp[1], count: insertMap[key] });
      }
      ins.sort(function (a, b) { return a.pos - b.pos; });
      var offset = 0;
      ins.forEach(function (e) {
        var fill = new Array(e.count + 1).join(fillChar).split('');
        Array.prototype.splice.apply(chars, [e.pos + offset, 0].concat(fill));
        offset += fill.length;
      });
      return chars.join('');
    }).join(' ');
  }

  // Round-robin the top-n slots (tagged with a run index ri) across runs and
  // splice tatweels into each run's text. Returns a new same-length array.
  function applySlotsMulti(runTexts, slots, n) {
    var out = runTexts.slice();
    if (!n || !slots.length) return out;
    var byRun = {}; // ri -> { "wi:pos": count }
    for (var i = 0; i < n; i++) {
      var s = slots[i % slots.length];
      var map = byRun[s.ri] || (byRun[s.ri] = {});
      var key = s.wi + ':' + s.pos;
      map[key] = (map[key] || 0) + 1;
    }
    for (var ri in byRun) {
      out[ri] = insertIntoWords(runTexts[ri], byRun[ri], TATWEEL);
    }
    return out;
  }

  // Natural (tatweel-free) total width of a run array, each measured in its
  // own font via its measure() callback.
  function measureRunsNatural(runs) {
    var total = 0;
    for (var i = 0; i < (runs || []).length; i++) {
      var r = runs[i];
      if (!r || typeof r.measure !== 'function') {
        throw new TypeError('run[' + i + '] is missing a measure() function');
      }
      total += r.measure(stripTatweels(r.text || ''));
    }
    return total;
  }

  // Apply the top-scoring slots to a text string, inserting the requested number of tatweels.
  function applySlots(text, slots, n) {
    if (!n || !slots.length) return text;
    var insertMap = {};
    for (var i = 0; i < n; i++) {
      var s = slots[i % slots.length];
      var key = s.wi + ':' + s.pos;
      insertMap[key] = (insertMap[key] || 0) + 1;
    }
    return insertIntoWords(text, insertMap, TATWEEL);
  }

  // Run-aware kashida: justify a misra whose pieces may each use a different
  // font. runs = [{ text, measure, fontProfile? }]; returns a same-length array
  // of { text } with tatweels inserted, measured per run in its own font.
  function justifyRuns(runs, targetWidth, params) {
    params = params || {};
    runs = runs || [];
    var texts = runs.map(function (r) {
      if (!r || typeof r.measure !== 'function') {
        throw new TypeError('run is missing a measure() function');
      }
      return stripTatweels(r.text || '');
    });
    var results = texts.map(function (t) { return { text: t }; });

    // Nothing to stretch: no visible text.
    var hasText = texts.some(function (t) { return t.trim(); });
    if (!hasText) return results;

    // Already fills the target width (measureRunsNatural sums per-run widths).
    var natural = measureRunsNatural(runs);
    var target = targetWidth * (params.targetFill || 1);
    if (natural >= target) return results;

    // Build a global, quality-ranked slot list tagged with the owning run.
    var slots = [];
    texts.forEach(function (t, ri) {
      buildSlots(t, params, runs[ri].fontProfile || null).forEach(function (s) {
        slots.push({ ri: ri, wi: s.wi, pos: s.pos, score: s.score });
      });
    });
    if (!slots.length) return results;
    slots.sort(function (a, b) {
      if (b.score !== a.score) return b.score - a.score;
      if (a.ri !== b.ri) return a.ri - b.ri;
      if (a.wi !== b.wi) return a.wi - b.wi;
      return a.pos - b.pos;
    });

    // Binary-search the total tatweel count; width is monotonic in the count.
    var totalLetters = texts.reduce(function (acc, t) {
      return acc + t.replace(/\s/g, '').length;
    }, 0);
    var lo = 1, hi = totalLetters * 8, best = texts;
    while (lo <= hi) {
      var mid = (lo + hi) >> 1;
      var cand = applySlotsMulti(texts, slots, mid);
      var w = 0;
      for (var k = 0; k < runs.length; k++) w += runs[k].measure(cand[k]);
      if (w <= target) { best = cand; lo = mid + 1; }
      else { hi = mid - 1; }
    }
    return best.map(function (t) { return { text: t }; });
  }

  // Find the maximum acceptable number of tatweels for a single line.
  // Thin wrapper over justifyRuns (single run) so the single- and multi-font
  // paths share one algorithm.
  function justifyLine(text, targetWidth, ctx, params, fontProfile) {
    var run = {
      text: text,
      measure: function (s) { return ctx.measureText(s).width; },
      fontProfile: fontProfile || null
    };
    return justifyRuns([run], targetWidth, params)[0].text;
  }

  // Apply justification to each non-empty line in the input array.
  function justifyLines(lines, containerWidth, ctx, params, fontProfile) {
    return lines.map(function (l) {
      return l.trim() ? justifyLine(l, containerWidth, ctx, params, fontProfile) : l;
    });
  }

  return {
    tatweelSlots: tatweelSlots,
    spreadTatweels: spreadTatweels,
    buildSlots: buildSlots,
    applySlots: applySlots,
    justifyLine: justifyLine,
    justifyLines: justifyLines,
    insertIntoWords: insertIntoWords,
    applySlotsMulti: applySlotsMulti,
    measureRunsNatural: measureRunsNatural,
    justifyRuns: justifyRuns
  };
}));

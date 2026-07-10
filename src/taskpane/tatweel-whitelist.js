/**
 * AshaarTatweel — turn a font's tatweel whitelist into a priorityTable the
 * vendored buildSlots()/canInsertTatweel() already honor (a pair with
 * {blocked:true} is skipped). We block every adjacent letter-pair present in
 * the text whose *next* letter is not elongatable in this font, so the engine
 * only inserts tatweels the font actually renders. Pure; no DOM.
 */
(function (root, factory) {
  if (typeof module !== "undefined" && module.exports) module.exports = factory();
  else root.AshaarTatweel = factory();
}(typeof globalThis !== "undefined" ? globalThis : this, function () {
  "use strict";

  // Combining marks / non-letters we skip when finding adjacent base letters.
  var SKIP = /[ً-ٰٟـ\s]/; // harakat, superscript alef, tatweel, space

  function baseLetters(word) {
    var out = [];
    for (var i = 0; i < word.length; i++) {
      var ch = word.charAt(i);
      if (!SKIP.test(ch)) out.push(ch);
    }
    return out;
  }

  function buildPriorityTable(text, rules) {
    var table = {};
    if (!rules || !rules.medialInto) return table;
    var allowed = {};
    for (var a = 0; a < rules.medialInto.length; a++) allowed[rules.medialInto[a]] = true;
    var words = String(text).split(" ");
    words.forEach(function (w) {
      var letters = baseLetters(w);
      for (var i = 0; i < letters.length - 1; i++) {
        var next = letters[i + 1];
        if (!allowed[next]) table[letters[i] + next] = { blocked: true };
      }
    });
    return table;
  }

  return { buildPriorityTable: buildPriorityTable };
}));

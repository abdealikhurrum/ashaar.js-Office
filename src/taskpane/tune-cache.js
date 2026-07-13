/**
 * tune-cache.js — memoization for AshaarTune probe/calibrate results.
 * probeFont depends only on (fontFamily, engine build) → persisted to storage.
 * calibrate depends on poem texts/width/font/size → in-memory only (texts churn).
 * Pure: no DOM, no Office.js. See spec §8.
 */
(function (root, factory) {
  if (typeof module !== "undefined" && module.exports) module.exports = factory();
  else root.AshaarTuneCache = factory();
}(typeof globalThis !== "undefined" ? globalThis : this, function () {
  "use strict";
  var PREFIX = "ashaar:fontProbe:";
  function hash32(s) {
    var h = 0; s = String(s || "");
    for (var i = 0; i < s.length; i++) h = ((h << 5) - h + s.charCodeAt(i)) | 0;
    return (h >>> 0).toString(16);
  }
  function probeKey(family, engineVersion) { return "probe|" + family + "|" + engineVersion; }
  function calibKey(family, sizePt, containerPx, texts) {
    // Floor (not round) to a 25px bucket: 412 and 420 must land in the SAME
    // bucket (test-pinned) — round(412/25)*25=400 but round(420/25)*25=425,
    // which fails that assertion; floor gives 400 for both.
    var bucket = Math.floor(Number(containerPx || 0) / 25) * 25;
    return "calib|" + family + "|" + Number(sizePt || 0) + "|" + bucket + "|" + hash32((texts || []).join(""));
  }
  function makeCache(storage) {
    var calib = {};
    // Key-tracking (not storage enumeration): the test's storage shim has no
    // length/key, and real localStorage may hold unrelated keys we must not
    // touch. Track exactly the keys THIS instance has written via putProbe,
    // and remove exactly those on bustAll.
    var written = {};
    function sKey(k) { return PREFIX + k; }
    return {
      getProbe: function (k) {
        if (!storage) return null;
        try { var raw = storage.getItem(sKey(k)); return raw ? JSON.parse(raw) : null; }
        catch (e) { return null; }
      },
      putProbe: function (k, profile) {
        if (!storage) return;
        try { storage.setItem(sKey(k), JSON.stringify(profile)); written[k] = true; }
        catch (e) { /* quota — skip */ }
      },
      getCalib: function (k) { return (k in calib) ? calib[k] : null; },
      putCalib: function (k, params) { calib[k] = params; },
      bustAll: function () {
        calib = {};
        if (!storage) return;
        try {
          Object.keys(written).forEach(function (k) { storage.removeItem(sKey(k)); });
        } catch (e) { /* ignore */ }
        written = {};
      },
    };
  }
  return { probeKey: probeKey, calibKey: calibKey, makeCache: makeCache, _hash32: hash32 };
}));

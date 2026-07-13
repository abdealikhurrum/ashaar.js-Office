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
  // Cross-session index of probe keys this cache family has written, so a
  // NEW session's bustAll (e.g. user replaces a font file under the same
  // family name) can clear probes persisted by a PREVIOUS session.
  var INDEX_KEY = PREFIX + "__index";
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
  /**
   * Reattach the function properties that JSON round-tripping strips from a
   * cached FontProfile. Mirrors ashaar-autotune's FontProfile methods
   * (src/vendor/ashaar-autotune.js:198-208) byte-for-behavior; the versioned
   * probe key busts this cache on engine updates, bounding drift. The vendor
   * closures read only `pairQualities`/`tierQuality`, which the profile also
   * exposes as plain data properties — so rehydration is purely data-driven
   * (no closure state is lost to JSON). Pure: mutates+returns `data` only.
   */
  function rehydrateFontProfile(data) {
    if (!data || typeof data !== "object") return data;
    var pairQualities = data.pairQualities || {};
    var tierQuality = data.tierQuality || {};
    /** Quality 0–1 for a specific character pair. (autotune:199-202) */
    data.getQuality = function (prev, next) {
      var e = pairQualities[prev + next];
      return e ? e.quality : 0.5;
    };
    /** Average quality for a HarfBuzz priority tier. (autotune:205-208) */
    data.getTierQuality = function (priority) {
      var tier = tierQuality[priority];
      return tier ? tier.avg : 0.5;
    };
    return data;
  }
  function makeCache(storage) {
    var calib = {};
    // Key-tracking (not storage enumeration): the test's storage shim has no
    // length/key, and real localStorage may hold unrelated keys we must not
    // touch. Track exactly the keys written via putProbe — seeded from the
    // persisted __index so bustAll also clears keys written by PRIOR sessions
    // (font replacement after a reload must not keep serving stale probes).
    var written = {};
    if (storage) {
      try {
        var idxRaw = storage.getItem(INDEX_KEY);
        if (idxRaw) JSON.parse(idxRaw).forEach(function (k) { written[k] = true; });
      } catch (e) { /* corrupt/unavailable index — start empty */ }
    }
    function sKey(k) { return PREFIX + k; }
    function saveIndex() {
      try { storage.setItem(INDEX_KEY, JSON.stringify(Object.keys(written))); }
      catch (e) { /* quota — index may lag; bustAll still clears in-memory set */ }
    }
    return {
      getProbe: function (k) {
        if (!storage) return null;
        try { var raw = storage.getItem(sKey(k)); return raw ? JSON.parse(raw) : null; }
        catch (e) { return null; }
      },
      putProbe: function (k, profile) {
        if (!storage) return;
        try { storage.setItem(sKey(k), JSON.stringify(profile)); written[k] = true; saveIndex(); }
        catch (e) { /* quota — skip */ }
      },
      getCalib: function (k) { return (k in calib) ? calib[k] : null; },
      putCalib: function (k, params) { calib[k] = params; },
      bustAll: function () {
        calib = {};
        if (!storage) { written = {}; return; }
        try {
          Object.keys(written).forEach(function (k) { storage.removeItem(sKey(k)); });
          storage.removeItem(INDEX_KEY);
        } catch (e) { /* ignore */ }
        written = {};
      },
    };
  }
  return { probeKey: probeKey, calibKey: calibKey, makeCache: makeCache, rehydrateFontProfile: rehydrateFontProfile, _hash32: hash32 };
}));

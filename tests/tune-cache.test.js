const assert = require("assert");
const TC = require("../src/taskpane/tune-cache");

assert.strictEqual(TC.probeKey("Jameel", "v7"), "probe|Jameel|v7");
assert.strictEqual(TC.probeKey("Jameel", "v8") === TC.probeKey("Jameel", "v7"), false, "engine bump busts");

const k1 = TC.calibKey("Jameel", 14, 412, ["الف", "ب"]);
assert.strictEqual(TC.calibKey("Jameel", 14, 420, ["الف", "ب"]), k1, "same 25px bucket");
assert.notStrictEqual(TC.calibKey("Jameel", 14, 460, ["الف", "ب"]), k1, "different bucket");
assert.notStrictEqual(TC.calibKey("Jameel", 14, 412, ["الف", "ج"]), k1, "texts hash differs");

// storage-backed probes, in-memory calib
const mem = {};
const storage = { getItem: k => (k in mem ? mem[k] : null), setItem: (k, v) => { mem[k] = v; }, removeItem: k => { delete mem[k]; } };
const c = TC.makeCache(storage);
assert.strictEqual(c.getProbe("probe|J|v7"), null);
c.putProbe("probe|J|v7", { pairs: { AB: 1 } });
assert.deepStrictEqual(c.getProbe("probe|J|v7"), { pairs: { AB: 1 } });
assert.deepStrictEqual(TC.makeCache(storage).getProbe("probe|J|v7"), { pairs: { AB: 1 } }, "survives new instance via storage");
c.putCalib("k", { targetFill: 0.9 });
assert.deepStrictEqual(c.getCalib("k"), { targetFill: 0.9 });
assert.strictEqual(TC.makeCache(storage).getCalib("k"), null, "calib is in-memory only");
c.bustAll();
assert.strictEqual(c.getProbe("probe|J|v7"), null, "bust clears storage too");

// probe round-trip is JSON: data fields survive, function properties are
// (expectedly) dropped by JSON.stringify.
const c2 = TC.makeCache(storage);
const liveProfile = {
  fontFamily: "J", fontSize: 64, standaloneWidth: 10,
  pairQualities: { "بب": { quality: 0.75, priority: 9 } },
  tierQuality: { 9: { sum: 0.75, count: 1, avg: 0.75 } },
  getQuality: function () { return 1; }
};
c2.putProbe("probe|J|v9", liveProfile);
const roundTripped = c2.getProbe("probe|J|v9");
assert.deepStrictEqual(roundTripped.pairQualities, liveProfile.pairQualities, "data fields survive round-trip");
assert.deepStrictEqual(roundTripped.tierQuality, liveProfile.tierQuality, "tierQuality survives round-trip");
assert.strictEqual(roundTripped.standaloneWidth, 10);
assert.strictEqual(typeof roundTripped.getQuality, "undefined", "functions dropped by JSON");

// rehydrateFontProfile reattaches the vendor FontProfile methods over the
// JSON-surviving data. The vendor's probeFont (src/vendor/ashaar-autotune.js)
// needs document.fonts + canvas so it can't build a reference profile in
// node; expected values below are hand-computed from the vendor source:
//   getQuality (ashaar-autotune.js:199-202): pairQualities[prev+next].quality, else 0.5
//   getTierQuality (ashaar-autotune.js:205-208): tierQuality[priority].avg, else 0.5
const rh = TC.rehydrateFontProfile(roundTripped);
assert.strictEqual(rh.getQuality("ب", "ب"), 0.75, "known pair → stored quality");
assert.strictEqual(rh.getQuality("x", "y"), 0.5, "unknown pair → vendor fallback 0.5");
assert.strictEqual(rh.getTierQuality(9), 0.75, "known tier → stored avg");
assert.strictEqual(rh.getTierQuality(7), 0.5, "unknown tier → vendor fallback 0.5");
assert.strictEqual(rh, roundTripped, "rehydrate mutates+returns the same object");
assert.strictEqual(TC.rehydrateFontProfile(null), null, "null passthrough");

// bustAll works ACROSS sessions: the written-keys index is persisted, so a
// second makeCache over the same storage busts keys the first instance wrote.
const mem2 = {};
const storage2 = { getItem: k => (k in mem2 ? mem2[k] : null), setItem: (k, v) => { mem2[k] = v; }, removeItem: k => { delete mem2[k]; } };
const s1 = TC.makeCache(storage2);
s1.putProbe("probe|X|v1", { pairs: 1 });
s1.putProbe("probe|Y|v1", { pairs: 2 });
const s2 = TC.makeCache(storage2); // "new session"
assert.deepStrictEqual(s2.getProbe("probe|X|v1"), { pairs: 1 }, "second instance reads persisted probe");
s2.bustAll();
assert.strictEqual(s2.getProbe("probe|X|v1"), null, "cross-session bust: X cleared");
assert.strictEqual(s1.getProbe("probe|Y|v1"), null, "cross-session bust: Y cleared (visible to first instance too)");
assert.strictEqual(Object.keys(mem2).length, 0, "storage fully empty after cross-session bust (index removed)");

console.log("tune-cache tests passed");

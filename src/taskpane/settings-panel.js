/**
 * settings-panel.js — pure panel-state logic for the unified Settings panel.
 *
 * Computes what the panel should show (values, provenance, dirty flags,
 * visible actions) from a resolveSettings() result + the pending-edits buffer
 * + the cursor target. NO DOM and NO Office.js — taskpane.js renders this.
 * See docs/superpowers/specs/2026-07-12-unified-settings-design.md.
 */
(function (root, factory) {
  if (typeof module !== "undefined" && module.exports) module.exports = factory();
  else root.AshaarPanel = factory();
}(typeof globalThis !== "undefined" ? globalThis : this, function () {
  "use strict";

  // Canonical keys shown per scope. Gap scope edits slotDecor (not
  // settings-keyed), so its list is empty here.
  var SCOPE_FIELDS = {
    poem: ["justifyMode", "fillMode", "strength", "gap", "widthMode", "widthPct",
           "layoutMode", "colWidthMode", "misraWidthPt", "fontCorrections", "debugColors"],
    bandh: ["misraWidthPt"],
    cell: ["strength", "misraWidthPt", "capEm"],
    gap: [],
  };

  var LEVELS = ["poem", "bandh", "cell", "gap"];

  function scopeTitle(target, resolved) {
    if (target.kind !== "block") return "Selection";
    var name = resolved.profileName
      ? resolved.profileName + (resolved.profileMissing ? " (not on this machine)" : "")
      : "(no profile)";
    var head = "Poem — " + name;
    if (!resolved.profileName) head = "Poem";
    var lvl = target.scope.level;
    if (lvl === "bandh") return head + " › Bandh" + (target.bandhLabel ? " " + target.bandhLabel : "");
    if (lvl === "cell") return head + " › Cell " + (target.cellLabel || target.scope.key || "");
    if (lvl === "gap") return head + " › Gap " + (target.gapLabel || target.scope.key || "");
    return head;
  }

  function panelStateFor(args) {
    var resolved = args.resolved;
    var pending = args.pending || { set: {}, clear: [] };
    var target = args.target || { kind: "selection", scope: { level: "poem" } };
    var level = target.scope.level || "poem";

    var chips = target.kind === "block" ? LEVELS.map(function (l) {
      return {
        level: l,
        enabled: l === "poem" || l === "bandh" ? true : (l === "cell" ? !!target.cellEnabled : !!target.gapEnabled),
        active: l === level,
      };
    }) : [];

    var controls = SCOPE_FIELDS[level].map(function (key) {
      var dirty = (key in pending.set) || pending.clear.indexOf(key) !== -1;
      var value = (key in pending.set) ? pending.set[key]
        : (pending.clear.indexOf(key) !== -1 ? inheritedValue(resolved, key) : resolved.values[key]);
      return { key: key, value: value, source: resolved.source[key], dirty: dirty };
    });

    var anyDirty = Object.keys(pending.set).length > 0 || pending.clear.length > 0;

    return {
      header: { title: scopeTitle(target, resolved) },
      chips: chips,
      controls: controls,
      profileRow: {
        name: resolved.profileName,
        missing: resolved.profileMissing,
        assignEnabled: target.kind === "block",
        updateVisible: !!resolved.profileName && !resolved.profileMissing && anyDirty && level === "poem",
        restoreVisible: resolved.profileMissing && resolved.usedCache,
      },
      footer: {
        applyEnabled: true,
        revertLabel: resolved.profileName ? "Revert to profile" : "Reset to defaults",
      },
    };
  }

  // What the value would be if the local delta for `key` were removed:
  // profile layer if it set the key, else the default. (Bandh/cell layers are
  // not consulted — clears happen at the scope that owns the key.)
  function inheritedValue(resolved, key) {
    // resolved.source tells where the CURRENT value came from; for a clear we
    // need the layer below "local". The resolver exposes this indirectly: a
    // key whose source is "local" inherits the profile value when the profile
    // owns it, else the default. panelStateFor callers pass resolvedBase for
    // exactness; this fallback recomputation is display-only (Apply re-resolves).
    if (resolved.inherited && (key in resolved.inherited)) return resolved.inherited[key];
    return resolved.values[key];
  }

  function mergePending(pending, key, value) {
    var set = {}, clear = pending.clear.slice();
    Object.keys(pending.set).forEach(function (k) { set[k] = pending.set[k]; });
    var ci = clear.indexOf(key);
    if (value === null || value === undefined) {   // ⟲ reset → clear the delta
      delete set[key];
      if (ci === -1) clear.push(key);
    } else {
      set[key] = value;
      if (ci !== -1) clear.splice(ci, 1);
    }
    return { set: set, clear: clear };
  }

  // Compute the NEW full local map: apply sets, drop clears; only keys legal
  // for the scope. Never mutates inputs.
  function pendingToLocal(local, pending, scopeKeys) {
    var out = {};
    Object.keys(local || {}).forEach(function (k) {
      if (pending.clear.indexOf(k) === -1) out[k] = local[k];
    });
    Object.keys(pending.set).forEach(function (k) {
      if (!scopeKeys || scopeKeys.indexOf(k) !== -1) out[k] = pending.set[k];
    });
    return out;
  }

  return {
    SCOPE_FIELDS: SCOPE_FIELDS,
    panelStateFor: panelStateFor,
    mergePending: mergePending,
    pendingToLocal: pendingToLocal,
  };
}));

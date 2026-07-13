/**
 * AshaarOverrides — per-cell justify override key + merge (pure, UMD;
 * node-testable, no DOM/Office). An override deviates one cell from its block's
 * justify defaults on strength / target width / cap-lift. fillMode is NOT
 * overridable (block-wide choice). See
 * docs/superpowers/specs/2026-07-11-per-cell-overrides-design.md.
 */
(function (root, factory) {
  if (typeof module !== "undefined" && module.exports) module.exports = factory();
  else root.AshaarOverrides = factory();
}(typeof globalThis !== "undefined" ? globalThis : this, function () {
  "use strict";

  // Per-cell key: the SP1 label prefixed by the cell's table (bandh) index in
  // the block, so an override targets exactly one cell.
  function overrideKey(tableIndex, label) {
    return (tableIndex || 0) + ":" + label;
  }

  // Merge a cell override onto the block's base justify settings. Absent fields
  // inherit. fillMode always comes from base. widthPt inherits base.widthPt
  // (the bandh- or qaseeda-level misra width) when the cell doesn't set one;
  // null when unset everywhere (the caller then uses the matrix target).
  // capEm is null when unset (the caller applies its own default).
  function resolveCellOverride(base, override) {
    base = base || {};
    override = override || {};
    var s = (override.strength != null) ? override.strength : base.strength;
    return {
      strength: s,
      fillMode: base.fillMode,
      widthPt: (override.widthPt != null) ? override.widthPt
        : (base.widthPt != null) ? base.widthPt : null,
      capEm: (override.capEm != null) ? override.capEm : null
    };
  }

  // Merge a per-slot decoration override onto the profile default. Per field: an
  // override key present wins (empty string = explicit none); else inherit the
  // profile; else "". Fields: symbol / fill / color.
  function resolveSlotDecor(profileDecor, override) {
    profileDecor = profileDecor || {};
    override = override || {};
    function pick(k) { return (k in override) ? (override[k] || "") : (profileDecor[k] || ""); }
    return { symbol: pick("symbol"), fill: pick("fill"), color: pick("color") };
  }

  // §4 transition-clear: the targeted keys whose persisted override HAD a text
  // color that the incoming override no longer carries. Word's font.color has
  // no "no color" clear value (unlike shadingColor's "#FFFFFF" quirk) and the
  // renderer is otherwise set-only for color — so the Apply that deletes a
  // color override must know exactly which cells are transitioning color→none
  // to reset them explicitly. Empty string counts as "no color" on both sides
  // (same convention as setTagOverride's codec).
  function colorClearKeys(oldOverrides, keys, newOverride) {
    oldOverrides = oldOverrides || {};
    if (newOverride && newOverride.color != null && newOverride.color !== "") return [];
    return (keys || []).filter(function (k) {
      var o = oldOverrides[k];
      return !!(o && o.color != null && o.color !== "");
    });
  }

  // Final review I2: fan-out merge for a bandh/poem-target cell-scope Apply.
  // The pane's `incoming` override is derived from the CURRENT cell's state,
  // so writing it verbatim onto every fanned-out sibling key would delete
  // fields the user never touched this Apply (e.g. "apply fill to whole
  // poem" wiping every cell's strength/color). Per field: if the user
  // touched it this Apply (`touched[field]` true — dirty in the pending
  // buffer, INCLUDING a ⟲-clear, whose incoming value is null and must still
  // clear on every targeted key), use `incoming`; otherwise keep the key's
  // OWN existing value (still null if it never had one). The current key
  // bypasses this (full replace — the pane is the full truth for it); see
  // applyPanel.
  function mergeFanOutOverride(existing, incoming, touched) {
    existing = existing || {};
    incoming = incoming || {};
    touched = touched || {};
    function pick(field) {
      return touched[field] ? (incoming[field] != null ? incoming[field] : null)
        : (existing[field] != null ? existing[field] : null);
    }
    return {
      strength: pick("strength"),
      widthPt: pick("widthPt"),
      capEm: pick("capEm"),
      fill: pick("fill"),
      color: pick("color")
    };
  }

  return {
    overrideKey: overrideKey,
    resolveCellOverride: resolveCellOverride,
    resolveSlotDecor: resolveSlotDecor,
    colorClearKeys: colorClearKeys,
    mergeFanOutOverride: mergeFanOutOverride
  };
}));

const assert = require("assert");
const AshaarStyles = require("../src/taskpane/word-styles");

// ── ROLES / STYLE_NAME / BASE_STYLE / STYLE_TYPE ────────────────────────────

{
  assert.deepEqual(AshaarStyles.ROLES,
    ["heading1", "heading2", "heading3", "emphasis", "quote", "quranQuote"],
    "role order matters: quote must precede quranQuote (basedOn dependency)");
  assert.equal(AshaarStyles.STYLE_NAME.heading1, "Ashaar Heading 1");
  assert.equal(AshaarStyles.STYLE_NAME.quranQuote, "Ashaar Quran Quote");
  assert.equal(AshaarStyles.BASE_STYLE.heading1, "Heading 1");
  assert.equal(AshaarStyles.BASE_STYLE.emphasis, "Emphasis");
  assert.equal(AshaarStyles.BASE_STYLE.quote, "Ashaar Normal",
    "Quote follows the Ashaar Normal body style by default");
  assert.equal(AshaarStyles.BASE_STYLE.quranQuote, "Ashaar Quote",
    "Quran Quote is based on OUR Quote style, not a Word built-in");
  assert.equal(AshaarStyles.NORMAL_STYLE_NAME, "Ashaar Normal");
  assert.equal(AshaarStyles.STYLE_TYPE.emphasis, "Character");
  assert.equal(AshaarStyles.STYLE_TYPE.quote, "Paragraph");
}

// ── defaultGroup ─────────────────────────────────────────────────────────────

{
  const g = AshaarStyles.defaultGroup("General");
  assert.equal(g.name, "General");
  assert.equal(g.heading1.font, "Kanz Al Marjaan");
  assert.equal(typeof g.heading1.sizePt, "number");
  assert.ok(g.heading1.sizePt > g.heading2.sizePt, "heading1 larger than heading2");
  assert.ok(g.heading2.sizePt > g.heading3.sizePt, "heading2 larger than heading3");
  assert.equal(g.emphasis.color, "#FF0000");
  assert.equal(typeof g.emphasis.bumpPt, "number");
  assert.equal(g.quote.borderColor, "#000000");
  assert.equal(typeof g.quote.indentPt, "number");
  assert.equal(g.quranQuote.font, "Amiri Quran");
  assert.equal(g.quranQuote.lineHeightPt, null, "null = Word auto by default");
}

{
  // No name still normalizes to an empty-string-named group, not undefined.
  const g = AshaarStyles.defaultGroup();
  assert.equal(g.name, "");
}

// ── mergeGroup (deep-merge overrides onto a base, one level per role) ────────

{
  const base = AshaarStyles.defaultGroup("Base");
  const merged = AshaarStyles.mergeGroup(base, { heading1: { sizePt: 24 }, quote: { indentPt: 18 } });
  assert.equal(merged.heading1.sizePt, 24, "override applied");
  assert.equal(merged.heading1.font, base.heading1.font, "unset heading1 fields keep base");
  assert.equal(merged.quote.indentPt, 18);
  assert.equal(merged.quote.borderColor, base.quote.borderColor, "unset quote fields keep base");
  assert.equal(merged.emphasis.color, base.emphasis.color, "roles absent from partial are untouched");
  assert.notEqual(merged, base, "returns a new object, does not mutate base");
}

// ── normalizeGroup (fill missing roles/fields from defaults) ─────────────────

{
  const g = AshaarStyles.normalizeGroup({ name: "Partial", heading1: { sizePt: 30 } });
  assert.equal(g.name, "Partial");
  assert.equal(g.heading1.sizePt, 30, "keeps provided field");
  assert.equal(g.heading1.font, "Kanz Al Marjaan", "fills missing field from default");
  assert.ok(g.quote && typeof g.quote.indentPt === "number", "fills entirely-missing role");
}

// ── BUILTIN_GROUPS ────────────────────────────────────────────────────────────

{
  var names = Object.keys(AshaarStyles.BUILTIN_GROUPS);
  assert.deepEqual(names.sort(), ["General", "Maqala", "Petition", "Waaz"]);
  Object.keys(AshaarStyles.BUILTIN_GROUPS).forEach(function (k) {
    var g = AshaarStyles.BUILTIN_GROUPS[k];
    assert.equal(g.name, k, "built-in group's name matches its key");
    AshaarStyles.ROLES.forEach(function (role) {
      assert.ok(g[role], "built-in group '" + k + "' defines role '" + role + "'");
    });
  });
}

// ── computeEmphasisSize ───────────────────────────────────────────────────────

{
  assert.equal(AshaarStyles.computeEmphasisSize(12, 3), 15);
  assert.equal(AshaarStyles.computeEmphasisSize(10.5, 2.5), 13);
  assert.equal(AshaarStyles.computeEmphasisSize(null, 3), 15,
    "missing base size falls back to 12pt (Word's own default)");
  assert.equal(AshaarStyles.computeEmphasisSize(12, null), 12,
    "missing bump behaves as +0");
}

// ── clampIndentPt / clampLineHeightPt ────────────────────────────────────────

{
  assert.equal(AshaarStyles.clampIndentPt(18), 18);
  assert.equal(AshaarStyles.clampIndentPt(-5), 0, "negative indent clamps to 0");
  assert.equal(AshaarStyles.clampIndentPt(500), 200, "clamps to a 200pt ceiling");
  assert.equal(AshaarStyles.clampIndentPt(null), 0);

  assert.equal(AshaarStyles.clampLineHeightPt(null), null, "null (auto) passes through");
  assert.equal(AshaarStyles.clampLineHeightPt(24), 24);
  assert.equal(AshaarStyles.clampLineHeightPt(2), 6, "clamps to a 6pt floor");
  assert.equal(AshaarStyles.clampLineHeightPt(500), 200, "clamps to a 200pt ceiling");
}

console.log("word-styles.test.js: all assertions passed");

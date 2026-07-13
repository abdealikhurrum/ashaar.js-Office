# Font surgery: mark-safe Jameel Noori Nastaleeq Kasheeda

**Date:** 2026-07-12 · **Status:** requirements for a SEPARATE font repo (own workflow) · **Consumer:** `ashaar.js-Office` apply engine (no engine changes needed).

## 1. Goal

Make the Kasheeda face's wide ligatures match **vocalized** text, with harakat correctly positioned, so the add-in's font-swap kashida works on fully-vocalized Arabic (qasidas, marsiyas) the way it already works on plain Urdu. "Real deal kashidas."

## 2. The defect (proven 2026-07-12, fonttools + live measurement)

- Both Jameel faces map harakat (U+064B–U+0652, U+0670) to real glyphs (names `.notdefNNN` — the font is a converted InPage-era build) and GDEF classes them as marks (class 3). GPOS has `MarkBasePos`, `MarkLigPos`, `MarkMarkPos`, `CursivePos`.
- The bulk `rlig` LigatureSubst lookups — **18,664 rules (Regular) / 24,621 rules (Kasheeda)** — carry **LookupFlag `0x0001` (RightToLeft only): IgnoreMarks NOT set**. Only ~1,100 rules (lookups 26–37, 41 in Regular) set `0x0009`.
- Consequence (spec-conformant in every shaper — HarfBuzz/Chrome, Word, WebView): a haraka *between* ligature components blocks the match → vocalized words fall back to unwidened forms.
- Measured @12pt: `قَدر` gains 0.1px vocalized vs **12.0px** stripped; `العَزائم` 1.3 → **15.0**; `تَعرفُ` 13.4 → **25.6**; `الحَمراء` 1.7 → **17.8**; `صِغارها` 0.2 → **12.2**.
- Even stripped, some words gain ~0 (`في`, `عَين`, `الصغير`, `العظيم`, `المكارم`, `لونها`) — the Kasheeda face only widens a specific ligature vocabulary. Surgery restores the *vocalized* share of the existing vocabulary; it does not invent new wide forms.

## 3. Prototype result (what one flag flip buys — and breaks)

Flipping IgnoreMarks on all 33 ligature/contextual GSUB lookups of the Kasheeda face ("JNNK MarkSafe" prototype):

- ✅ **Widths unlock**: vocalized `وذلكَ ما لا تَدَّعيهِ الضراغم` fill 120→137px; `يُفدّي أتَمُّ الطَيرِ...` 107→138px; 10 of 14 test misras get 29–68px of newly-available gain.
- ❌ **Harakat displace**: skipped marks attach at the ligature's END, not over their base letter — the fatha of `قَدر` slides toward the د, the kasra of `صِغارها` lands under the ا, `سِلاحَهُ`'s kasra drifts onto لا. This changes the reading — unshippable as-is.
- Benefit and harm coincide exactly: only mark-carrying fasls gain from the flip, and precisely those get displaced marks. **The missing piece is per-component mark anchors.**

## 4. Requirements for the real fix

1. **Scope: the Kasheeda face only.** The engine renders unswapped text in Regular (correct marks today) and swaps fasls *into* Kasheeda; only glyphs reachable via the swap need surgery. (Optionally repeat on Regular later for visual consistency of vocalized ligatures.)
2. **Flip IgnoreMarks** (`LookupFlag |= 0x0008`) on the LigatureSubst + Context/ChainContext GSUB lookups (33 in the Kasheeda face).
3. **Per-component mark anchors** so skipped marks re-attach over the right letter:
   - Component count per ligature comes from its GSUB rule (the input glyph sequence).
   - Preferred: `MarkLigPos` (GPOS type 5) LigatureAttach records — one anchor row per component per mark class.
   - x-anchors: approximate from the proportional advance of each component's standalone form scaled to the ligature advance; refine per-ligature where the kashida stroke concentrates the stretch in ONE joint (marks before the stretch keep near-original x from the right edge — RTL).
   - y-anchors: Nastaliq slant means components climb; derive y from outline extrema per component zone, or from the base font's mark anchors carried along the slant line. This is the genuinely hard part — budget iteration with visual proofing.
   - Pragmatic sequencing: don't anchor all ~24k ligatures up front. Extract the subset that is (a) actually wider than Regular AND (b) has ≥1 mark-carrying component in a target corpus (the add-in can dump swap candidates); anchor those first.
4. **Ligature carets** (GDEF LigCaretList) optional but useful for the same component geometry.
5. **Naming/versioning**: bump `head`/`name` version; keep the family name **identical** ("Jameel Noori Nastaleeq Kasheeda") for a drop-in with zero add-in changes, or use a new family name and add one registry line in `src/taskpane/fonts.js` (`kasheedaName` for the Jameel entry). Same-name is preferred; the add-in measures at runtime, so improved gains are discovered automatically.
6. **Licensing**: Jameel Noori Nastaleeq is distributed as freeware; confirm modification/redistribution terms before publishing the patched build anywhere public. Private/internal use for typesetting is the assumed context.

## 5. Verification harness (reuse from this session)

- **Width A/B**: browser page with `@font-face` for Regular/Kasheeda/patched + canvas `measureText`; assert vocalized gain ≈ stripped gain per golden word. (Chrome canvas was proven pixel-identical to the Word-WebView measurement pipeline on this exact corpus.)
- **Mark placement**: large-size render table (word × face) — every haraka visually on its base letter. Compare against Regular as ground truth.
- **Regression**: an unvocalized Urdu corpus renders byte-identical widths before/after (IgnoreMarks is a no-op without marks, but contextual-lookup flips must be sanity-checked).
- **In Word**: insert the vocalized golden list styled in the patched face; confirm widths and marks match the browser (Word's shaper honors the same OTL data).
- **Golden words** (vocalized, @12pt, expected unlocked gain ≈ stripped gain): `قَدر` +12, `صِغارها` +12, `العَزائم` +15, `الحَمراء` +17.8, `تَعرفُ` +25.6 (13.4 already available), plus zero-gain sentinels `في`, `عَينِ`, `الصغيرِ` (must stay zero — no invented forms).

## 6. Reference scripts (from the diagnosis session)

### 6a. Forensics — `inspect_jameel.py`

```python
"""Checks GDEF mark classes, GSUB ligature-lookup IgnoreMarks flags,
ligature inventory, and GPOS mark-attachment presence."""
from fontTools.ttLib import TTFont

HARAKAT = [0x064B, 0x064C, 0x064D, 0x064E, 0x064F, 0x0650, 0x0651, 0x0652, 0x0670]
IGNORE_MARKS = 0x0008

def inspect(path):
    f = TTFont(path)
    cmap = f.getBestCmap()
    have = {h: cmap.get(h) for h in HARAKAT}
    cd = f["GDEF"].table.GlyphClassDef.classDefs if "GDEF" in f and f["GDEF"].table.GlyphClassDef else {}
    print("harakat marks:", {g: cd.get(g) for g in have.values() if g})
    gsub = f["GSUB"].table
    for i, lk in enumerate(gsub.LookupList.Lookup):
        nlig = 0
        for st in lk.SubTable:
            if type(st).__name__ == "ExtensionSubst":
                st = st.ExtSubTable
            if hasattr(st, "ligatures"):
                nlig += sum(len(v) for v in st.ligatures.values())
        if nlig:
            print("lig lookup %d: flag=0x%04X IgnoreMarks=%s rules=%d"
                  % (i, lk.LookupFlag, bool(lk.LookupFlag & IGNORE_MARKS), nlig))
```

### 6b. Prototype patch — `patch_jnnk.py`

```python
"""Flips IgnoreMarks on ligature+contextual GSUB lookups and renames the
family so the patched face loads side-by-side for A/B. This alone unlocks
widths but DISPLACES marks (see §3) — anchors (§4.3) are the real work."""
from fontTools.ttLib import TTFont

IGNORE_MARKS = 0x0008
f = TTFont(SRC)
for lk in f["GSUB"].table.LookupList.Lookup:
    kinds = set()
    for st in lk.SubTable:
        t = type(st).__name__
        if t == "ExtensionSubst":
            t = type(st.ExtSubTable).__name__
        kinds.add(t)
    if kinds & {"LigatureSubst", "ChainContextSubst", "ContextSubst"}:
        lk.LookupFlag |= IGNORE_MARKS
# ...rename family in name IDs 1,3,4,6,16, then f.save(DST)
```

## 7. Payoff in the add-in (already wired)

Nothing to change: `applyProfileToQaseeda` measures fasl widths in both faces at apply time, so a mark-safe Kasheeda automatically converts today's spacing-filled residuals into real kashida. The adaptive-harmony fill mode's shared target `T*` will also rise on its own (weakest lines gain reach). Expected end state on the Mutanabbi test: most lines kashida-flush at the box with ≤1–2px of trim spacing.

## 8. Related

- Memory: `jameel-kasheeda-harakat-limit` (the ceiling this removes), `justification-modes-state` (engine state).
- Handoff: `docs/superpowers/2026-07-12-fable-handoff-mixed-font-justify.md` (the round-trip saga that preceded this).

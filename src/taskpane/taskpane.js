(function () {
  var input = document.getElementById("poem-input");
  var preview = document.getElementById("preview");
  var message = document.getElementById("message");
  var hostStatus = document.getElementById("host-status");
  var modeTable = document.getElementById("mode-table");
  var modeConvert = document.getElementById("mode-convert");
  var tablePanel = document.getElementById("table-mode-panel");
  var convertPanel = document.getElementById("convert-mode-panel");
  var bandhCount = document.getElementById("bandh-count");
  var misraCount = document.getElementById("misra-count");
  var layoutPreset = document.getElementById("layout-preset");
  var layoutSpec = document.getElementById("layout-spec");

  // Table width as a fraction of the page text column (centred). 100% = full
  // width. pct comes from the caller's resolved opts.tableWidthPct (the
  // Settings panel's widthMode/widthPct, via options()) — there is no
  // standalone table-width DOM control anymore.
  function scaledTextWidth(twips, pct) {
    var clampedPct = Math.max(25, Math.min(100, Number(pct || 50)));
    return Math.max(1, Math.round(twips * clampedPct / 100));
  }

  // Smallest table width (twips) that fits the widest misra in each column at the
  // given font (ctx), leaving kashida headroom. Drives Auto-fit and the width nudge.
  function neededTableTwips(source, ctx, opts, pageTwips) {
    if (!ctx || typeof Ashaar === "undefined" || !Ashaar.parse) return pageTwips;
    var maxByPos = [];
    (Ashaar.parse(source) || []).forEach(function (p) {
      (p.stanzas || []).forEach(function (s) {
        (s.bayts || []).forEach(function (b) {
          var texts = (b.type === "row" && b.misras) ? b.misras.map(function (m) { return m.text; })
                    : (b.ajuz ? [b.sadr, b.ajuz] : [b.sadr]);
          texts.forEach(function (t, i) {
            var w = ctx.measureText(stripJustification(String(t || ""))).width;
            if (!(maxByPos[i] >= w)) maxByPos[i] = w;
          });
        });
      });
    });
    if (!maxByPos.length) return pageTwips;
    var kashidaOn = (opts.justifyMode === "kashida" || opts.justifyMode === "spacing") && Number(opts.tatweelCount || 0) > 0;
    var headroom = kashidaOn ? 0.9 : 0.98;        // leave room for kashida to fill
    var N = Math.max(maxByPos.length, 2);
    var gapCols = Math.max(1, Math.round(Number(opts.gapWidth || 1)));
    // 6 must mirror BASE_CPM in word-html.js (grid columns per misra) so the
    // gap-to-content ratio here matches the actual render; otherwise the needed
    // width (which now floors the table width) is overestimated and tables come
    // out too wide.
    var GRID = N * 6 + (N - 1) * gapCols, contentCols = N * 6;
    var sumContent = maxByPos.reduce(function (a, b) { return a + b; }, 0) / headroom;
    return Math.round(sumContent * GRID / contentCols * 1440 / 96);
  }
  var templateNameInput = document.getElementById("template-name");
  var templateList = document.getElementById("template-list");
  var importFileInput = document.getElementById("import-file");
  var sepMode = document.getElementById("sep-mode");
  var sepCustom = document.getElementById("sep-custom");
  var sepPair = document.getElementById("sep-pair");
  var layoutGridEl = document.getElementById("layout-grid");
  var layoutViewGridBtn = document.getElementById("layout-view-grid");
  var layoutViewNumbersBtn = document.getElementById("layout-view-numbers");
  var debugMode = document.getElementById("debug-mode");
  var debugOutput = document.getElementById("debug-output");

  // §8 probe/calibration memoization (tune-cache.js). Probes are keyed by
  // font+engine build and persisted to localStorage (machine-scoped font
  // metrics); calibration sessions are keyed by font+size+container-bucket+
  // texts-hash and kept in-memory only (poem text churns too much to persist
  // usefully). Busted wholesale when the fonts strip registers/replaces a font.
  var _tuneCache = (typeof AshaarTuneCache !== "undefined")
    ? AshaarTuneCache.makeCache(typeof localStorage !== "undefined" ? localStorage : null)
    : null;
  // Engine-build id for probe cache-busting. NOT read from
  // src/vendor/ASHAAR_UPSTREAM_VERSION at runtime (that file is a build-time
  // stamp written by scripts/sync-ashaar-vendor.mjs; it's never loaded into
  // the browser, and this vanilla/no-build-step add-in has no existing
  // fetch()/XHR path to read a sibling asset). Mirrors that file's `commit=`
  // field by hand — bump this string whenever `npm run update:ashaar` moves
  // the pointer, until sync-ashaar-vendor.mjs is taught to stamp this value
  // into a loadable JS file instead.
  var ASHAAR_UPSTREAM_VERSION = "caf103f1"; // src/vendor/ASHAAR_UPSTREAM_VERSION: commit=caf103f1c8...

  // Format collected per-cell justification metrics into the Debug panel.
  // `meta` (optional) surfaces session-level probe/calibrate cache hit|miss.
  function renderDebug(diags, meta) {
    if (!debugOutput) return;
    if (!diags.length) { debugOutput.textContent = "(no kashida cells measured)"; return; }
    var metaLine = meta ? ("probe=" + meta.probe + " calib=" + meta.calib + "\n") : "";
    var head = "cell  font        res  col(in)   nat  target  final  fill  tw/cap  text";
    var rows = diags.map(function (d) {
      return [
        String(d.i).padEnd(4),
        d.font.padEnd(11),
        String(d.res || "?").padStart(3),
        (d.colPx + "(" + d.colIn + ")").padEnd(9),
        String(d.nat).padStart(4),
        String(d.target).padStart(6),
        String(d.fin).padStart(6),
        (d.fill + "%").padStart(5),
        (d.tw + "/" + d.cap).padStart(6),
        "  " + d.text
      ].join(" ");
    });
    debugOutput.textContent = metaLine + head + "\n" + rows.join("\n");
  }

  // Whether the WebView can actually render `name` vs silently falling back.
  // document.fonts.check is unreliable for system fonts (true for unknown names),
  // so compare measured widths against generic families — if the font changes the
  // width over a generic baseline, it resolved; if it matches all three, it fell back.
  var _fontProbeCtx = null;
  function fontAvailable(name) {
    if (!name) return false;
    try {
      if (!_fontProbeCtx) _fontProbeCtx = document.createElement("canvas").getContext("2d");
      var ctx = _fontProbeCtx;
      if (!ctx) return true;
      var test = "mMgwiحيبٹكطولِ ظہور";
      var generics = ["monospace", "sans-serif", "serif"];
      for (var i = 0; i < generics.length; i++) {
        ctx.font = "72px " + generics[i];
        var base = ctx.measureText(test).width;
        ctx.font = "72px '" + name + "'," + generics[i];
        if (Math.abs(ctx.measureText(test).width - base) > 0.5) return true;
      }
      return false;
    } catch (e) { return true; }
  }

  var GRID_COLS = 12;
  var layoutView = "numbers";   // "grid" | "numbers"
  var gridMatrix = [];          // rows of 12 booleans, reading order (index 0 = visual right)

  var SEP_LABELS = {
    backslash: "\\", asterisk: "*", pipe: "|", dash: "dash",
    tab: "tab", spaces: "double space", custom: "custom", pairLines: "paired lines"
  };

  // Apply separator normalization to the current editor text using the import
  // options. Auto-detects by default; explicit choices and pair-lines override.
  // Re-renders the preview and notes what changed.
  function applyImportNormalization() {
    if (typeof AshaarSeparators === "undefined") { renderPreview(); return; }
    var res = AshaarSeparators.normalizeSeparators(input.value, {
      separator: sepMode ? sepMode.value : "auto",
      customPattern: sepCustom ? sepCustom.value : "",
      pairLines: sepPair ? sepPair.checked : false
    });
    if (res.changed) {
      input.value = res.text;
      setMessage("Converted separators (" + (SEP_LABELS[res.detected] || res.detected) + ") to standard \\ form.");
    }
    renderPreview();
  }

  // ── Visual layout grid (Grid mode) ──────────────────────────────────────────

  // Default row: a paired couplet — 5 on / 2 gap / 5 on (reading order).
  function defaultGridRow() {
    var r = [];
    for (var i = 0; i < 5; i++) r.push(true);
    for (i = 0; i < 2; i++) r.push(false);
    for (i = 0; i < 5; i++) r.push(true);
    return r;
  }

  function renderLayoutGrid() {
    if (!layoutGridEl || typeof AshaarLayoutGrid === "undefined") return;
    var html = "";
    gridMatrix.forEach(function (row, ri) {
      html += '<div class="lg-row">';
      html += '<button type="button" class="lg-rm" data-row="' + ri + '" title="Remove row" aria-label="Remove row">✕</button>';
      html += '<span class="lg-bubbles">';
      for (var c = 0; c < GRID_COLS; c++) {
        html += '<span class="lg-b' + (row[c] ? " on" : "") + '" data-row="' + ri + '" data-col="' + c + '"></span>';
      }
      html += "</span></div>";
    });
    html += '<button type="button" class="lg-add" id="lg-add-row">+ Add row</button>';

    var tpl = AshaarLayoutGrid.gridToTemplate(gridMatrix);
    html += '<div class="lg-prev"><span class="lg-prev-label">Preview</span>';
    tpl.rows.forEach(function (cells) {
      html += '<div class="lg-prev-row">';
      cells.forEach(function (cell) {
        if (cell.role === "gap") html += '<span class="lg-prev-gap" style="flex:' + cell.span + '"></span>';
        else html += '<span class="lg-prev-cell" style="flex:' + cell.span + '">·</span>';
      });
      html += "</div>";
    });
    html += "</div>";
    layoutGridEl.innerHTML = html;
  }

  function onLayoutGridClick(e) {
    var t = e.target;
    if (!t) return;
    if (t.id === "lg-add-row") { gridMatrix.push(defaultGridRow()); renderLayoutGrid(); return; }
    if (t.classList.contains("lg-rm")) {
      var ri = Number(t.getAttribute("data-row"));
      if (gridMatrix.length > 1) gridMatrix.splice(ri, 1); else gridMatrix = [defaultGridRow()];
      renderLayoutGrid();
      return;
    }
    if (t.classList.contains("lg-b")) {
      var r = Number(t.getAttribute("data-row"));
      var c = Number(t.getAttribute("data-col"));
      if (gridMatrix[r]) { gridMatrix[r][c] = !gridMatrix[r][c]; renderLayoutGrid(); }
    }
  }

  // Toggle between the visual Grid and the Numbers (text) view of the layout spec.
  // The two are kept in sync best-effort: Numbers→Grid parses the spec into bubbles;
  // Grid→Numbers serializes the bubbles back to the text spec.
  function setLayoutView(view) {
    if (typeof AshaarLayoutGrid === "undefined") return;
    layoutView = view === "grid" ? "grid" : "numbers";
    var grid = layoutView === "grid";
    layoutViewGridBtn.classList.toggle("is-active", grid);
    layoutViewNumbersBtn.classList.toggle("is-active", !grid);
    layoutViewGridBtn.setAttribute("aria-pressed", String(grid));
    layoutViewNumbersBtn.setAttribute("aria-pressed", String(!grid));
    layoutGridEl.hidden = !grid;
    layoutSpec.hidden = grid;
    if (grid) {
      var m = AshaarLayoutGrid.specToGrid(layoutSpec.value);
      gridMatrix = m.length ? m : [defaultGridRow()];
      renderLayoutGrid();
    } else {
      layoutSpec.value = AshaarLayoutGrid.gridToSpec(gridMatrix);
      renderPreview();
    }
  }

  // Thin adapter over the Settings panel: every insert/justify/re-render path
  // still calls options() for its knobs, so this is the ONE place the panel's
  // resolved+pending values (panelValues()) reach the rest of the file. Table
  // Input-only fields (bandhCount, layoutSpec) still read their own (non-panel)
  // controls directly — the panel governs poem justify/layout/width.
  function options() {
    var v = panelValues();
    return {
      justifyMode: v.justifyMode,
      justify: v.justifyMode === "none" ? false : v.justifyMode,
      fillMode: v.fillMode,
      layoutMode: v.layoutMode,
      layout: v.layoutMode,
      widthMode: v.colWidthMode,
      bandhCount: Number(bandhCount.value || 1),
      misraCount: Number(misraCount.value || 4),
      // Table-Input misra pattern control; the panel has no field for this.
      misraPattern: layoutPreset.value,
      layoutSpec: (!tablePanel.hidden) ? layoutSpec.value : "",
      fontMode: "document",
      tatweelCount: Number(v.strength || 6),
      gapWidth: Number(v.gap != null ? v.gap : 4),
      lineHeightPt: v.lineHeightPt,
      separatorPt: v.separatorPt,
      tableWidthPct: v.widthMode === "fixed" ? Number(v.widthPct || 50) : 100,
      autoFitWidth: v.widthMode !== "fixed",
      // v3 tag fields for fresh inserts. When the panel is focused on an
      // EXISTING block, its pending/profile belong to that block — a fresh
      // insert must not inherit them (cross-block leak).
      profile: (_panel.target && _panel.target.kind === "block") ? "" : (_panel.resolved ? _panel.resolved.profileName : ""),
      local: (_panel.target && _panel.target.kind === "block") ? {} : AshaarPanel.pendingToLocal({}, _panel.pending, AshaarPanel.SCOPE_FIELDS.poem),
    };
  }

  function previewFontFamily(font) {
    var mode = font === "nastaliq" ? "noto" : font;
    var css = AshaarFonts.cssFamilyOf(mode);
    return css || "\"Times New Roman\", serif";
  }

  function setMessage(text) {
    message.textContent = text || "";
  }

  // Office.js collapses many failures into the opaque string "GeneralException".
  // The useful detail (which API statement threw) lives in error.debugInfo. This
  // surfaces code + errorLocation to the pane and dumps the full error + debugInfo
  // to the WebView console (open devtools to see the stack / surrounding statements).
  function describeError(error) {
    try { console.error("[ashaar] error:", error, error && error.debugInfo); } catch (e) {}
    if (!error) return "Unknown error.";
    var di = error.debugInfo || {};
    var parts = [];
    var head = error.message || error.code || String(error);
    parts.push(head);
    if (di.code && di.code !== error.message) parts.push("code=" + di.code);
    if (di.errorLocation) parts.push("at " + di.errorLocation);
    // surroundingStatements pinpoints the API calls around the throw when
    // errorLocation is an internal resolver (e.g. _GetObjectByReferenceId).
    var ss = di.surroundingStatements;
    if (ss && ss.length) parts.push("near: " + ss.join(" | "));
    return parts.join(" — ");
  }

  function setMode(mode) {
    var isTable = mode === "table";
    modeTable.classList.toggle("is-active", isTable);
    modeConvert.classList.toggle("is-active", !isTable);
    modeTable.setAttribute("aria-selected", String(isTable));
    modeConvert.setAttribute("aria-selected", String(!isTable));
    tablePanel.classList.toggle("is-active", isTable);
    convertPanel.classList.toggle("is-active", !isTable);
    tablePanel.hidden = !isTable;
    convertPanel.hidden = isTable;
    setMessage(isTable ? "Table input mode: draw a blank grid, then type in Word." : "Ashaar.js conversion mode: paste source text, then insert a converted table.");
  }

  function renderPreview() {
    var opts = options();
    preview.className = "ashaar preview";
    // Mirror the chosen table width: a narrower, centred preview previews the insert.
    preview.style.maxWidth = opts.autoFitWidth ? "100%" : (opts.tableWidthPct + "%");
    preview.style.marginInline = "auto";
    preview.style.setProperty("--ashaar-font-family", previewFontFamily(opts.fontMode));
    preview.innerHTML = Ashaar.renderText(String(input.value || ""), { gapWidth: opts.gapWidth + "%" });
    Ashaar.applyRenderOptions(preview, { gapWidth: opts.gapWidth + "%" });
    if (opts.layout === "stacked") preview.classList.add("ashaar--stacked");
    if (opts.layout === "auto" || opts.layout === "compact") Ashaar.applyAutoLayout(preview, { layout: "auto" });
    if (opts.justify === "css") {
      preview.classList.add("ashaar--justify");
    } else if (opts.justify === "spacing") {
      Ashaar.justifyEl(preview, { method: "spacing", tatweel: false });
    } else if (opts.justify === "kashida") {
      // tatweelCount=0 suppresses tatweels; otherwise justifyEl fills to available width.
      // Fine-grained strength is applied on OOXML insertion (not controllable here without
      // modifying the vendor ashaar.js justifyMisra function).
      Ashaar.justifyEl(preview, opts.tatweelCount === 0 ? { tatweel: false } : {});
    }
  }

  function layoutSpecForPreset(preset, count) {
    count = Math.max(1, Number(count || 4));
    var rows = [];
    var i;
    if (preset === "centered-stack") {
      for (i = 1; i <= count; i++) rows.push("<" + i + ">");
      return rows.join("\n");
    }
    if (preset === "alternate-right") {
      for (i = 1; i <= count; i++) rows.push(i % 2 ? i + " >" : "< " + i);
      return rows.join("\n");
    }
    if (preset === "indented-stack") {
      for (i = 1; i <= count; i++) rows.push(new Array(i).join("  ") + i);
      return rows.join("\n");
    }
    if (preset === "karbala-refrain") {
      return "3 | 2 | 1\n<4>\n6 - 5";
    }
    for (i = 1; i <= count; i += 2) rows.push((i + 1 <= count ? i + 1 : "") + " - " + i);
    return rows.join("\n");
  }

  function applyLayoutPreset() {
    layoutSpec.value = layoutSpecForPreset(layoutPreset.value, misraCount.value);
  }

  // Final review I3: returns true/false so a caller further up the chain
  // (e.g. reRender) can gate ITS OWN trailing message instead of
  // unconditionally overwriting whatever this set. Existing callers that
  // don't need the signal are unaffected (they just don't await the value).
  async function withWord(callback) {
    if (typeof Word === "undefined") {
      setMessage("Open this task pane inside Word to update the document.");
      return false;
    }
    try {
      await Word.run(callback);
      setMessage("Done.");
      return true;
    } catch (error) {
      setMessage(describeError(error));
      return false;
    }
  }

  // Strict variant for the panel Apply paths: failures must propagate so
  // applyPanel keeps pending edits for retry (spec: apply failure keeps edits).
  async function withWordStrict(callback) {
    if (typeof Word === "undefined") throw new Error("Open this task pane inside Word to update the document.");
    await Word.run(callback);
  }

  // ── Qaseeda profiles — document store + block tagging (P2) ────────────────
  // Profiles live in Word document settings (one authoritative copy that travels
  // with the .docx), keyed by name. A block is linked to a qaseeda by the
  // `qaseeda` field in its content-control tag. Pure profile math is in
  // profiles.js (AshaarProfiles); this layer is the Office.js orchestration.
  var PROFILE_STORE_KEY = "ashaar:profiles";

  function loadProfileStore() {
    try {
      if (typeof Office === "undefined" || !Office.context || !Office.context.document) return {};
      var raw = Office.context.document.settings.get(PROFILE_STORE_KEY);
      var obj = raw ? (typeof raw === "string" ? JSON.parse(raw) : raw) : {};
      return (obj && typeof obj === "object") ? obj : {};
    } catch (e) { return {}; }
  }

  function saveProfileStore(store) {
    return new Promise(function (resolve) {
      try {
        if (typeof Office === "undefined" || !Office.context || !Office.context.document) { resolve(false); return; }
        Office.context.document.settings.set(PROFILE_STORE_KEY, JSON.stringify(store || {}));
        Office.context.document.settings.saveAsync(function (res) {
          resolve(!!(res && res.status === Office.AsyncResultStatus.Succeeded));
        });
      } catch (e) { resolve(false); }
    });
  }

  function getProfile(name) {
    var p = loadProfileStore()[name];
    return p ? AshaarProfiles.normalizeProfile(p) : AshaarProfiles.defaultProfile(name);
  }

  async function putProfile(profile) {
    var p = AshaarProfiles.normalizeProfile(profile);
    if (!p.name) return false;
    var store = loadProfileStore();
    store[p.name] = p;
    return await saveProfileStore(store);
  }

  function listProfileNames() {
    return Object.keys(loadProfileStore());
  }

  // Read the profile name assigned to the Ashaar Poem block at the cursor
  // ({name:"", tag:""} if none). `tag` is the block's CURRENT raw content-
  // control tag at the moment of the read — callers that need to re-identify
  // this exact physical block later (e.g. applyProfileToQaseeda's onlyBlockTag
  // scoping) capture it here rather than opening a second Word.run.
  async function getQaseedaAtSelection() {
    var result = { name: "", tag: "" };
    if (typeof Word === "undefined") return result;
    try {
      await Word.run(async function (context) {
        var cc = context.document.getSelection().parentContentControlOrNullObject;
        cc.load("title,tag");
        await context.sync();
        if (!cc.isNullObject && cc.title === "Ashaar Poem") {
          var payload = AshaarWord.parseContentControlTag(cc.tag);
          result.name = (payload && payload.profile) || "";
          result.tag = cc.tag;
        }
      });
    } catch (e) { /* leave result empty */ }
    return result;
  }

  // Read-only: show the bandh cell-map (labels + gaps) for the Ashaar Poem block
  // at the cursor. No document mutation. Labels can't be shown on the Word page
  // itself (no native per-cell text overlay), so the pane is their home.
  async function showCellMap() {
    var view = document.getElementById("cell-map-view");
    if (!view) return;
    if (typeof Word === "undefined") { setMessage("Open this task pane inside Word."); return; }
    var patterns = null;
    try {
      await Word.run(async function (context) {
        var cc = context.document.getSelection().parentContentControlOrNullObject;
        cc.load("title,tag");
        await context.sync();
        if (!cc.isNullObject && cc.title === "Ashaar Poem") {
          var payload = AshaarWord.parseContentControlTag(cc.tag);
          patterns = payload && payload.cells;
        }
      });
    } catch (e) { /* leave patterns null */ }

    if (!patterns || !patterns.length) {
      view.hidden = false;
      view.textContent = "No cell map on the block at the cursor (older or hand-drawn table).";
      return;
    }
    var html = "";
    patterns.forEach(function (pattern, bi) {
      var map = AshaarCellMap.buildBandhCellMap(pattern);
      html += "<div class=\"cell-map-bandh\"><b>Bandh " + (bi + 1) + "</b>";
      var lastRow = -1, rowHtml = "";
      function flush() { if (rowHtml) html += "<div class=\"cell-map-row\">" + rowHtml + "</div>"; rowHtml = ""; }
      map.forEach(function (e) {
        if (e.row !== lastRow) { flush(); lastRow = e.row; }
        rowHtml += e.kind === "content"
          ? "<span class=\"cell-map-cell\">" + e.label + "</span>"
          : "<span class=\"cell-map-gap\">(gap)</span>";
      });
      flush();
      html += "</div>";
    });
    view.hidden = false;
    view.innerHTML = html;
  }

  // ── Active-block / active-cell reflection (SP2) ───────────────────────────
  var _lastBlockTag = null;          // last-reflected block tag (resync only on change)
  var _reflectPending = false;       // debounce guard
  var _reflectBusy = false;          // suppress reflection while our own justify runs
  var _activeOvKey = null;           // override key of the content cell at the cursor (or null)
  var _activeDecorKey = null;        // slot-decor key of the spacing cell at the cursor (or null)
  var _activeSlot = null;            // slot-position (e.g. "A#1") of the spacing cell at the cursor

  // §4 transition-clear: override keys whose text COLOR was just removed by
  // the current cell-scope Apply. Word's font.color has no "no color" write
  // (unlike shadingColor, where "#FFFFFF" is the documented clear value), and
  // the render sites are otherwise set-only for color — worse, the qaseeda
  // capture reads live run colors as "original", so an applied color would be
  // baked into the captured runs and survive override deletion forever.
  // applyPanel's cell branch records the affected keys here right after the
  // tag write; the render pass it triggers resets those cells to "black" and
  // the success tail empties the map (one-shot). Kept on a failed render so a
  // retry still clears — the tag is already colorless at that point.
  //   blockId scoping (fix round 2): override keys ("0:A1") are IDENTICAL
  // across poems, so a retained map + a later render on a DIFFERENT poem
  // would blacken the colliding cell there. blockId is
  // AshaarWord.tagIdentity(tag) — the tag minus the per-apply runFonts heal,
  // the ONE mutation the pipeline makes between Apply's write and
  // consumption. cc.id can't serve: the size-rebuild path deletes and
  // re-inserts the control (fresh id), while both consumers already resolve
  // blocks by tag. A consumer applies a clear ONLY when its block's identity
  // matches; retry-after-failure retention still works because the retried
  // Apply writes the same payload → same identity. Accepted limitation:
  // retained clears silently no-op if the user's follow-up after a failed
  // render mutates a non-runFonts tag field first (Assign writes profile;
  // poem-scope Apply writes local/profileCache) — identity differs, the
  // clear is skipped, and the stale color persists until manually
  // re-cleared. Safe direction (skip, not corruption).
  var _pendingColorClears = { blockId: "", keys: {} };

  // §4 (fix rounds 2+3): the composite "block tag|override key" the decor
  // inputs were last seeded for. Word fires DocumentSelectionChanged
  // constantly (see the pendingProfile note in _panel) — an unconditional
  // reseed on every reflection would wipe the user's checked-but-not-yet-
  // Applied fill/color state mid-edit. But the override key ALONE can't be
  // the tracker: "0:A1" exists in every poem, so clicking poem A's 0:A1 then
  // poem B's 0:A1 would skip the reseed and leak A's decor into an Apply on
  // B. The composite includes cc.tag, so a block change (or any tag rewrite
  // — changed tag can mean changed overrides) reseeds, while a spurious
  // same-cell reflection carries an unchanged tag and is skipped. The Apply
  // success tail nulls the tracker to force a reseed from the fresh tag.
  var _lastSeededDecorKey = null;

  // Final review I2: snapshot of what seedCellDecorInputs last put in the
  // fill/color controls, so applyPanel's cell branch can tell whether the
  // user actually TOUCHED fill/color this Apply (vs. the seeded value simply
  // sitting there) — untouched fields must not fan out onto sibling keys.
  // Reset to "nothing seeded" defaults whenever seedCellDecorInputs runs, so
  // an empty/no-override cell reads as untouched until the user edits it.
  var _seededCellDecor = { fillOn: false, fill: null, colorOn: false, color: null };

  // ── Unified settings panel state ──────────────────────────────────────────
  var _panel = {
    pending: { set: {}, clear: [] },
    // The profile dropdown's chosen-but-not-Assigned value. Like pending, it
    // must survive re-renders: DocumentSelectionChanged fires constantly in
    // Word and renderPanel resets the dropdown to the tag's resolved profile —
    // without this, the user's choice vanished before they could click Assign.
    pendingProfile: null,
    scopeLevel: "poem",
    target: null,      // { kind:"block"|"selection", cc?, payload?, scope, cellEnabled, gapEnabled, cellLabel?, gapKey? }
    resolved: null,
  };

  var SP_BODIES = { poem: "sp-body-poem", bandh: "sp-body-bandh", cell: "sp-body-cell", gap: "sp-body-gap" };

  function panelValues() {
    var out = {};
    var base = _panel.resolved ? _panel.resolved.values : AshaarProfiles.defaultSettings();
    Object.keys(base).forEach(function (k) { out[k] = base[k]; });
    _panel.pending.clear.forEach(function (k) {
      if (_panel.resolved && _panel.resolved.inherited) out[k] = _panel.resolved.inherited[k];
    });
    Object.keys(_panel.pending.set).forEach(function (k) { out[k] = _panel.pending.set[k]; });
    return out;
  }

  function refreshPanel() {
    var target = _panel.target || { kind: "selection", scope: { level: "poem" } };
    var scope = { level: _panel.scopeLevel, key: target.kind === "block"
      ? (_panel.scopeLevel === "cell" ? target.cellLabel : _panel.scopeLevel === "gap" ? target.gapKey : undefined)
      : undefined };
    _panel.resolved = AshaarProfiles.resolveSettings({
      payload: target.kind === "block" ? target.payload : null,
      profileStore: loadProfileStore(),
      scope: scope,
    });
    var st = AshaarPanel.panelStateFor({ resolved: _panel.resolved, pending: _panel.pending,
      target: { kind: target.kind, scope: scope, cellEnabled: !!target.cellEnabled,
                gapEnabled: !!target.gapEnabled, cellLabel: target.cellLabel, gapLabel: target.gapKey } });
    renderPanel(st);
  }

  function renderPanel(st) {
    document.getElementById("sp-target").textContent = st.header.title;
    var chipsWrap = document.getElementById("sp-chips");
    chipsWrap.hidden = st.chips.length === 0;
    st.chips.forEach(function (c) {
      var el = document.getElementById("sp-chip-" + c.level);
      el.disabled = !c.enabled;
      el.classList.toggle("is-active", c.active);
    });
    Object.keys(SP_BODIES).forEach(function (lvl) {
      document.getElementById(SP_BODIES[lvl]).hidden = lvl !== _panel.scopeLevel;
    });
    // Values + provenance dots. Controls carry data-key; skip ones the user is
    // mid-editing (focused).
    st.controls.forEach(function (c) {
      var body = document.getElementById(SP_BODIES[_panel.scopeLevel]);
      // NOTE (bug found in Task 6 browser verification, fixed here): the
      // provenance dot (.sp-src) and its form control share the same data-key
      // and the dot comes first in DOM order for every field markup, so an
      // unscoped selector grabbed the span instead of the input/select and
      // silently no-opped every value write. Exclude .sp-src explicitly.
      var input = body.querySelector('[data-key="' + c.key + '"]:not(.sp-src)');
      if (input && document.activeElement !== input) {
        input.value = c.value == null ? "" : c.value;
        if (c.key === "strength") document.getElementById("sp-strength-value").textContent = String(c.value);
      }
      var src = body.querySelector('.sp-src[data-key="' + c.key + '"]');
      if (src) {
        // c.resettable (settings-panel.js) covers dirty pending edits AND the
        // already-committed delta layer the CURRENT scope can actually clear
        // (poem→local, bandh→bandh, cell→cell) — a committed override must
        // keep showing the reset affordance, but only where resetting isn't
        // a wrong-layer no-op (review R5).
        src.textContent = c.resettable ? "•" : "";
        src.title = !c.resettable ? ""
          : c.dirty ? "edited — Apply to commit; click to reset"
          : c.source === "cell" ? "cell override — click to reset to inherited"
          : c.source === "bandh" ? "bandh width override — click to reset to inherited"
          : c.source === "local" ? "local tweak — click to reset to inherited" : "";
      }
    });
    // Profile row.
    var sel = document.getElementById("sp-profile");
    var names = listProfileNames();
    sel.innerHTML = "<option value=\"\">(none)</option>" + names.map(function (n) {
      return "<option value=\"" + String(n).replace(/"/g, "&quot;") + "\">" + String(n) + "</option>";
    }).join("");
    // An un-Assigned dropdown choice wins over the resolved name — reflection
    // re-renders constantly in Word and must not wipe the user's selection.
    sel.value = _panel.pendingProfile != null ? _panel.pendingProfile
      : (st.profileRow.missing ? "" : st.profileRow.name);
    document.getElementById("sp-profile-assign").disabled = !st.profileRow.assignEnabled;
    document.getElementById("sp-profile-update").hidden = !st.profileRow.updateVisible;
    document.getElementById("sp-profile-update").textContent = "Update \"" + st.profileRow.name + "\"";
    document.getElementById("sp-profile-restore").hidden = !st.profileRow.restoreVisible;
    document.getElementById("sp-revert").textContent = st.footer.revertLabel;
    document.getElementById("sp-cost").textContent = st.footer.costLabel;
    document.getElementById("sp-apply").title = st.footer.costLabel;
    document.getElementById("sp-rerender").disabled = !( _panel.target && _panel.target.kind === "block");
    // §5 Capture gating: _panel.scopeLevel is sticky (it survives the cursor
    // moving from a content cell to a sibling gap cell and vice versa), so the
    // visible body — and its Capture button — can outlive its target. Mirror
    // the chip-enablement conditions (settings-panel.js panelStateFor: chips
    // exist only for kind==="block", and cell/gap chips gate on
    // !!target.cellEnabled / !!target.gapEnabled) so Capture is clickable
    // exactly when its scope's chip is.
    var capTgt = _panel.target;
    document.getElementById("sp-cell-capture").disabled =
      !(capTgt && capTgt.kind === "block" && capTgt.cellEnabled);
    document.getElementById("sp-gap-capture").disabled =
      !(capTgt && capTgt.kind === "block" && capTgt.gapEnabled);
  }

  // Reflect the Ashaar block (and the cell, Task 4) at the cursor in the pane.
  async function reflectActiveContext() {
    if (typeof Word === "undefined" || _reflectBusy) return;
    try {
      await Word.run(async function (context) {
        var sel = context.document.getSelection();
        var cc = sel.parentContentControlOrNullObject;
        cc.load("title,tag");
        await context.sync();
        var isBlock = !cc.isNullObject && cc.title === "Ashaar Poem";
        var payload = isBlock ? AshaarWord.parseContentControlTag(cc.tag) : null;
        await reflectActiveCell(context, sel, cc, isBlock, payload);
        // Guarded reseed: only when the composite (block tag + cell key)
        // CHANGED since the last seed — spurious same-cell reflections carry
        // an unchanged tag and must not wipe mid-edit state, while the same
        // OVERRIDE KEY in a different poem (keys repeat across poems) or a
        // rewritten tag on the same cell MUST reseed. Never read cc.tag
        // unless isBlock (null-object proxy throws — same rule as below).
        var seedKey = (isBlock ? cc.tag : "") + "|" + _activeOvKey;
        if (seedKey !== _lastSeededDecorKey) {
          if (_activeOvKey) seedCellDecorInputs(payload);
          _lastSeededDecorKey = seedKey;
        }
        _panel.target = isBlock
          ? { kind: "block", cc: cc, payload: payload, tag: cc.tag,
              cellEnabled: !!_activeOvKey, gapEnabled: !!_activeDecorKey,
              cellLabel: _activeOvKey, gapKey: _activeDecorKey,
              scope: { level: _panel.scopeLevel } }
          : { kind: "selection", scope: { level: "poem" } };
        if (!isBlock) _panel.scopeLevel = "poem";
        // A new block target drops stale pending edits; same block keeps them.
        // Never read cc.tag unless isBlock — with the cursor in plain text cc
        // is a null-object proxy and property access throws inside Word.run.
        if (isBlock ? cc.tag !== _lastBlockTag : _lastBlockTag !== null) {
          _panel.pending = { set: {}, clear: [] };
          _panel.pendingProfile = null;
          _lastBlockTag = isBlock ? cc.tag : null;
        }
        refreshPanel();
      });
    } catch (e) { /* selection transient — ignore */ }
  }

  // §4 (fix): the cell fill/color controls live OUTSIDE the pending/data-key
  // system (same as the gap-decor inputs), so they are raw shared DOM state —
  // never re-rendered by renderPanel. Seed all four from the ACTIVE cell's
  // persisted override whenever the active cell CHANGES (guard in
  // reflectActiveContext via _lastSeededDecorKey); otherwise a checked box left
  // over from cell A would leak A's colors into an Apply on cell B, and an
  // unchecked box would silently DELETE B's persisted fill/color
  // (setTagOverride replaces the whole per-key override object). Called only
  // from reflectActiveContext (selection-driven), NOT refreshPanel — refresh
  // runs on every in-pane edit and re-seeding there would wipe the user's
  // not-yet-applied checkbox/color changes mid-edit.
  function seedCellDecorInputs(payload) {
    var ov = (payload && payload.overrides && payload.overrides[_activeOvKey]) || {};
    var hasFill = ov.fill != null && ov.fill !== "";
    var hasColor = ov.color != null && ov.color !== "";
    document.getElementById("sp-cell-fill-on").checked = hasFill;
    document.getElementById("sp-cell-fill").value = hasFill ? ov.fill : "#f5f0e0";
    document.getElementById("sp-cell-color-on").checked = hasColor;
    document.getElementById("sp-cell-color").value = hasColor ? ov.color : "#a7352a";
    // Final review I2: remember exactly what was just seeded, so applyPanel
    // can diff the CURRENT control state against it to decide whether the
    // user touched fill/color this Apply (fan-out must not overwrite sibling
    // keys' fill/color just because the seeded value is still sitting there).
    _seededCellDecor = {
      fillOn: hasFill, fill: hasFill ? ov.fill : null,
      colorOn: hasColor, color: hasColor ? ov.color : null
    };
  }

  // §5: read the cursor cell's native formatting into the pane as pending
  // values. Never writes. Normalizes Word's no-color/automatic quirks:
  // shadingColor of "#FFFFFF"/""/null ⇒ no fill; font.color of "" /
  // "Automatic"/"auto" ⇒ inherit (empty). Uses the same
  // sel.parentTableCellOrNullObject lookup as reflectActiveCell (line 710)
  // — the proven pattern for "which table cell is the cursor in" in this
  // codebase.
  async function captureCellFormatting() {
    if (typeof Word === "undefined") { setMessage("Open this task pane inside Word to capture."); return; }
    // Belt-and-braces re-check of the renderPanel button gating (covers any
    // render race): the sticky scopeLevel must still have a LIVE matching
    // target, or the fresh cell lookup below would stuff the real cell's
    // formatting into the wrong scope's inputs.
    var capTgt = _panel.target;
    var capOk = capTgt && capTgt.kind === "block" &&
      (_panel.scopeLevel === "gap" ? capTgt.gapEnabled : capTgt.cellEnabled);
    if (!capOk) {
      setMessage(_panel.scopeLevel === "gap"
        ? "Click inside a spacing (gap) cell first."
        : "Click inside a poem content cell first.");
      return;
    }
    try {
      await Word.run(async function (context) {
        var sel = context.document.getSelection();
        var tcell = sel.parentTableCellOrNullObject;
        tcell.load("shadingColor,body/text");
        tcell.body.font.load("color");
        await context.sync();
        if (tcell.isNullObject) { setMessage("Click inside a table cell first."); return; }
        var fill = tcell.shadingColor;
        if (!fill || /^#?F{6}$/i.test(String(fill).replace("#", "")) || String(fill).toLowerCase() === "auto") fill = "";
        var color = tcell.body.font.color;
        if (!color || String(color).toLowerCase() === "automatic" || String(color).toLowerCase() === "auto") color = "";
        var isGap = _panel.scopeLevel === "gap";
        var fillEl = document.getElementById(isGap ? "sp-gap-fill" : "sp-cell-fill");
        var fillOn = document.getElementById(isGap ? "sp-gap-fill-on" : "sp-cell-fill-on");
        var colorEl = document.getElementById(isGap ? "sp-gap-color" : "sp-cell-color");
        if (fill) { fillEl.value = fill; }
        fillOn.checked = !!fill;
        if (color) colorEl.value = color;
        if (isGap) {
          var sym = (tcell.body.text || "").trim();
          document.getElementById("sp-gap-symbol").value = sym;
        } else {
          var colorOn = document.getElementById("sp-cell-color-on");
          colorOn.checked = !!color;
        }
        setMessage("Captured — Apply to persist" + (_panel.scopeLevel === "cell" ? " (choose Apply-to target first)." : "."));
      });
    } catch (e) {
      setMessage("Capture failed: " + (e && e.message ? e.message : e));
    }
  }

  // Detect the content/spacing cell at the cursor. Resolves (tableIndex, label)
  // via the SP1 cells map and sets _activeOvKey/_activeDecorKey/_activeSlot for
  // the settings panel (chip enablement, scope key) — no DOM writes here; the
  // removed per-cell/per-gap editors were retired in favor of the panel.
  async function reflectActiveCell(context, sel, cc, isBlock, payload) {
    _activeOvKey = null;
    if (!isBlock || !payload || !payload.cells) return;

    var tcell = sel.parentTableCellOrNullObject;
    tcell.load("rowIndex,cellIndex,isNullObject");
    var tbls = cc.getRange().tables;
    tbls.load("items");
    await context.sync();
    if (tcell.isNullObject) return;

    // §6a: which block table contains the selection? (No stable table id, so
    // match by range intersection.)
    var selRange = sel.getRange();
    var inters = tbls.items.map(function (tbl) {
      var r = tbl.getRange().intersectWithOrNullObject(selRange); r.load("isNullObject"); return r;
    });
    await context.sync();
    var tIdx = -1;
    for (var k = 0; k < inters.length; k++) { if (!inters[k].isNullObject) { tIdx = k; break; } }
    if (tIdx < 0 || !payload.cells[tIdx]) return;

    var map = AshaarCellMap.buildBandhCellMap(payload.cells[tIdx]);
    var inRow = map.filter(function (e) { return e.row === tcell.rowIndex; });
    var entry = inRow[tcell.cellIndex];
    if (!entry) { _activeOvKey = null; _activeDecorKey = null; return; }
    if (entry.kind === "content") {
      _activeDecorKey = null;
      _activeOvKey = AshaarOverrides.overrideKey(tIdx, entry.label);
    } else { // spacing → decoration slot
      _activeOvKey = null;
      _activeDecorKey = AshaarOverrides.overrideKey(tIdx, entry.slot);
      _activeSlot = entry.slot;
    }
  }

  // Debounced entry point for the DocumentSelectionChanged event.
  function onSelectionChanged() {
    if (_reflectPending) return;
    _reflectPending = true;
    window.setTimeout(function () { _reflectPending = false; reflectActiveContext(); }, 150);
  }

  // ── Qaseeda apply/refresh engine (P3) ─────────────────────────────────────
  // Find every Ashaar Poem block linked to a qaseeda name.
  async function gatherQaseedaBlocks(context, name) {
    var ccs = context.document.contentControls;
    ccs.load("items/title,items/tag");
    await context.sync();
    return ccs.items.filter(function (cc) {
      if (cc.title !== "Ashaar Poem") return false;
      var p = AshaarWord.parseContentControlTag(cc.tag);
      return !!(name && p && p.profile === name);
    });
  }

  // Per-qaseeda cache (this pane session) of the width signature last applied.
  // Rebuilding a block's OOXML is destructive (delete + re-insert its content
  // control) and heavy, so we only do it when the target WIDTH changes. Justify-
  // only changes (strength, fill mode, per-cell override) keep the same signature
  // → skip the rebuild and just re-justify the existing (already-sized) tables.
  var _appliedSizeSig = {};

  // Capture a qaseeda's blocks as measured tables ready for width sizing + justify.
  // Loads each block's tables/rows/cells (text + real font), captures plain values
  // (later edits invalidate proxies, so we never re-read body.*), reconstructs each
  // block's SOURCE, derives each content cell's grid span from that source (Word
  // can't report span-table column geometry), and measures natural (tatweel-free)
  // widths on a canvas. Geometry uses each block's OWN stored opts (from its tag)
  // so it matches how the block was/will be rendered. Returns everything both the
  // rebuild and justify passes need.
  // Every decor symbol this block could have written into a gap cell: the
  // profile's per-slot spacingDecor plus the block tag's slotDecor overrides.
  // Feeds AshaarTableAdopt.stripDecorCells (gap-corruption fix, Part B) when a
  // table's shape can't be aligned to its persisted pattern.
  function blockDecorSymbols(profile, slotDecor) {
    var syms = [];
    [profile && profile.spacingDecor, slotDecor].forEach(function (map) {
      Object.keys(map || {}).forEach(function (k) {
        var s = map[k] && map[k].symbol;
        if (s) syms.push(s);
      });
    });
    return syms;
  }

  async function captureQaseedaTables(context, blocks, profile) {
    var fallbackName = "Times New Roman";

    var section = context.document.sections.getFirst();
    section.load("pageLayout/width,pageLayout/leftMargin,pageLayout/rightMargin");

    var blockTables = blocks.map(function (cc) { var t = cc.getRange().tables; t.load("items"); return t; });
    await context.sync();
    blockTables.forEach(function (t) { t.items.forEach(function (tbl) { tbl.rows.load("items"); }); });
    await context.sync();
    blockTables.forEach(function (t) { t.items.forEach(function (tbl) { tbl.rows.items.forEach(function (row) { row.cells.load("items"); }); }); });
    await context.sync();
    blockTables.forEach(function (t) { t.items.forEach(function (tbl) { tbl.rows.items.forEach(function (row) {
      row.cells.items.forEach(function (cell) {
        cell.body.load("text"); cell.body.font.load("name,size");
        // First-paragraph alignment + indent — the run-aware justify pass rebuilds
        // each cell via OOXML and must re-assert its visual side (sadr/ajuz/solo)
        // and any stacked-layout indent.
        cell.body.paragraphs.load("alignment,leftIndent");
      });
    }); }); });
    await context.sync();

    var pl = section.pageLayout;
    var pagePt = pl && pl.width ? (pl.width - (pl.leftMargin || 0) - (pl.rightMargin || 0)) : 468;
    var pageTwips = Math.round(pagePt * 20);

    var blockInfos = blocks.map(function (cc, bi) {
      var payload = AshaarWord.parseContentControlTag(cc.tag) || {};
      var pattern = payload.cells || null;
      var overrides = payload.overrides || {};
      var slotDecor = payload.slotDecor || {};
      var decorSyms = blockDecorSymbols(profile, slotDecor);
      var repFont = "", repSize = 0;
      var tableInfos = blockTables[bi].items.map(function (tbl, j) {
        var perRowCounts = tbl.rows.items.map(function (row) { return row.cells.items.length; });
        var tablePattern = pattern ? pattern[j] : null;
        var tblMap = AshaarCellMap.alignPatternToTable(perRowCounts, tablePattern)
          ? AshaarCellMap.buildBandhCellMap(tablePattern) : null;
        var seq = 0, cells = [], rowsText = [];
        tbl.rows.items.forEach(function (row, ri) {
          var cols = row.cells.items.length, rowText = [];
          row.cells.items.forEach(function (cell, ci) {
            var f = cell.body.font;
            var current = (cell.body.text || "").trim();
            var base = stripJustification(current);
            var mapped = tblMap ? tblMap[seq] : null;
            seq++;
            rowText.push(cell.body.text || "");
            // §7: never let a justification-artifact-only cell (e.g. a blank
            // spacing/gutter cell, or one reduced to stray tatweels) become the
            // block's representative font — such cells' fonts are stale/
            // meaningless. isArtifactRun (not stripJustification/base) is the
            // right check here: it also flags plain-space-only cells, which
            // stripJustification alone would not strip.
            if (f && f.name && !repFont && !AshaarFonts.isArtifactRun(current)) {
              repFont = f.name; if (f.size) repSize = f.size;
            }
            var p0 = cell.body.paragraphs.items && cell.body.paragraphs.items[0];
            var alv = p0 && p0.alignment;
            cells.push({
              cell: cell, current: current, base: base,
              measure: base.replace(/\s+/g, " ").trim(),
              matKey: mapped ? AshaarCellMap.columnGroupKey(tablePattern, mapped) : AshaarMatrix.positionKey({ row: ri, col: ci, span: cols }),
              kind: mapped ? mapped.kind : null,
              slot: (mapped && mapped.kind === "spacing") ? mapped.slot : null,
              decorKey: (mapped && mapped.kind === "spacing" && mapped.slot) ? AshaarOverrides.overrideKey(j, mapped.slot) : null,
              ovKey: (mapped && mapped.kind === "content" && mapped.label) ? AshaarOverrides.overrideKey(j, mapped.label) : null,
              fontName: (f && f.name) || "", fontSize: (f && f.size) || 0,
              align: alv === "Right" ? "right" : alv === "Left" ? "left" : "center",
              indentTwips: (p0 && p0.leftIndent) ? Math.round(p0.leftIndent * 20) : 0
            });
          });
          rowsText.push(rowText);
        });
        // Gap-corruption fix (Parts A+B): spacing cells must never re-enter the
        // reconstructed source. Aligned pattern → blank every gap cell (a
        // decorated gap like "٭" would otherwise read back as a misra and get
        // baked into the next rebuild as CONTENT). No alignment → defense in
        // depth: blank cells that are exactly one of the block's decor symbols.
        rowsText = tblMap
          ? AshaarTableAdopt.blankSpacingCells(rowsText, tablePattern)
          : AshaarTableAdopt.stripDecorCells(rowsText, decorSyms);
        return { tbl: tbl, cells: cells, rowsText: rowsText, overrides: overrides, slotDecor: slotDecor,
          bandhWidthPt: payload.widthPt, blockIdx: j, grid: 0 };
      });
      var source = tableInfos.map(function (ti) {
        return AshaarTableAdopt.adoptTableToSource(ti.rowsText, { direction: "rtl" });
      }).filter(function (s) { return s.trim(); }).join("\n\n");
      return { cc: cc, oldTag: cc.tag, payload: payload, source: source, repFont: repFont, repSize: repSize, tableInfos: tableInfos };
    });

    // Representative font for the canvas baseline. §7: skip artifact-only
    // cells (blank spacing cells, stray-tatweel cells) — same rationale as
    // the per-block repFont pick above.
    var repName = fallbackName, repSize = 16;
    for (var b0 = 0; b0 < blockInfos.length && repName === fallbackName; b0++) {
      var tis = blockInfos[b0].tableInfos;
      for (var t0 = 0; t0 < tis.length && repName === fallbackName; t0++) {
        var cs0 = tis[t0].cells;
        for (var c0 = 0; c0 < cs0.length; c0++) {
          if (cs0[c0].fontName && !AshaarFonts.isArtifactRun(cs0[c0].current)) {
            repName = cs0[c0].fontName; if (cs0[c0].fontSize) repSize = cs0[c0].fontSize; break;
          }
        }
      }
    }

    var canvasCtx = document.createElement("canvas").getContext("2d");
    if (canvasCtx) {
      canvasCtx.font = repSize + "pt \"" + repName + "\"";
      if (document.fonts && document.fonts.load) { try { await document.fonts.load(repSize + "pt \"" + repName + "\""); } catch (e) {} }
    }

    // Grid geometry per table from its OWN source + the block's stored opts (so it
    // matches the rendered table). Zips onto captured cells by emission order.
    var geoProfileStore = loadProfileStore();
    blockInfos.forEach(function (blk) {
      var p = blk.payload;
      var eff = AshaarProfiles.resolveSettings({ payload: p, profileStore: geoProfileStore, scope: { level: "poem" } }).values;
      var geomOpts = { gapWidth: eff.gap, layoutMode: eff.layoutMode };
      blk.tableInfos.forEach(function (info) {
        var flatGeo = [], grid = 0;
        try {
          var src = AshaarTableAdopt.adoptTableToSource(info.rowsText, { direction: "rtl" });
          AshaarWord.poemCellGeometry(src, geomOpts, Ashaar, pageTwips).forEach(function (st) {
            grid = Math.max(grid, st.GRID || 0);
            st.rows.forEach(function (row) { row.forEach(function (g) { flatGeo.push(g); }); });
          });
        } catch (e) { flatGeo = []; }
        if (flatGeo.length === info.cells.length) {
          info.cells.forEach(function (c, i) { c.gridCol = flatGeo[i].col; c.gridSpan = flatGeo[i].span; });
          info.grid = grid;
        } else {
          info.cells.forEach(function (c, i) { c.gridSpan = 1; c.gridCol = i; });
          info.grid = info.cells.length;
        }
      });
    });

    // Force-load EVERY distinct cell font (and its Kasheeda face, for font-swap)
    // before measuring — measureText silently substitutes an unloaded font and
    // returns wrong widths (mis-ranking swaps / mis-sizing the matrix). The rep
    // font alone isn't enough for multi-font poems or Jameel's wide face. See
    // memory font-measurement-model.
    // faceNameSet (name-only, no size) also feeds the Task 9 JIT measurement
    // gate — collected here since this is the one pass over every cell.
    var faceNameSet = {};
    if (canvasCtx && typeof document !== "undefined" && document.fonts && document.fonts.load) {
      var faceSet = {};
      blockInfos.forEach(function (blk) {
        blk.tableInfos.forEach(function (info) {
          info.cells.forEach(function (c) {
            var fnm = c.fontName || repName, fsz = c.fontSize || repSize;
            if (fnm) { faceSet[fsz + "pt \"" + fnm + "\""] = true; faceNameSet[fnm] = true; }
            var kn = (typeof AshaarFonts !== "undefined" && AshaarFonts.descriptorForFontName)
              ? AshaarFonts.descriptorForFontName(fnm).kasheedaName : null;
            if (kn) { faceSet[fsz + "pt \"" + kn + "\""] = true; faceNameSet[kn] = true; }
          });
        });
      });
      var faceLoads = [];
      Object.keys(faceSet).forEach(function (s) { faceLoads.push(document.fonts.load(s).catch(function () {})); });
      try { await Promise.all(faceLoads); } catch (e) {}
    }

    // Measure natural widths + build the cross-block harmony matrix.
    var qMatrixCells = [];
    if (canvasCtx) {
      blockInfos.forEach(function (blk) {
        blk.tableInfos.forEach(function (info) {
          info.cells.forEach(function (c) {
            if (!AshaarMatrix.isContentCell(c.measure)) return;
            var fnm = c.fontName || repName, fsz = c.fontSize || repSize;
            // Measure in the BASE face: a re-applied Jameel cell reports the wider
            // Kasheeda face, which would inflate the harmony width and drift the
            // target on each apply. wordName maps Kasheeda→base; leaves others.
            var mfnm = (typeof AshaarFonts !== "undefined" && AshaarFonts.descriptorForFontName)
              ? (AshaarFonts.descriptorForFontName(fnm).wordName || fnm) : fnm;
            canvasCtx.font = fsz + "pt \"" + mfnm + "\"";
            c.natPx = AshaarProfiles.applyFontCorrection(canvasCtx.measureText(c.measure).width, mfnm, profile.fontCorrections);
            qMatrixCells.push({ key: c.matKey, natural: c.natPx });
          });
        });
      });
    }
    var qMatrix = AshaarMatrix.buildMatrix(qMatrixCells);

    return { blockInfos: blockInfos, pagePt: pagePt, pageTwips: pageTwips, repName: repName, repSize: repSize, canvasCtx: canvasCtx, qMatrix: qMatrix, faceNames: Object.keys(faceNameSet) };
  }

  // Read each content cell's ORIGINAL per-word fonts (name/size/style/color)
  // BEFORE the SIZE rebuild flattens them to one representative font. Returns a
  // plain map keyed by block:table:cell index → { runs, align, indentTwips },
  // where runs are coalesced same-style segments. The rebuild regenerates cells
  // from font-less source text, so this is the only place the per-word fonts of
  // a mixed-font misra (e.g. Mehr + Amiri) still exist — pass 2 re-emits them.
  // Returns { cells, blockPacks }: cells = the per-cell run map; blockPacks =
  // per-block {"t:c": packRunWords(...)} of the RECONCILED words, ready to
  // persist in each block's tag. Reconciliation heals ambiguous font reads from
  // the tag's stored pack: Font.name reads the cs face for Arabic runs and ""
  // when a word's fasls carry mixed cs (base+Kasheeda) — proven in Word
  // 2026-07-12 — so a justified cell can NOT be read back reliably; the tag is
  // the source of truth and a clean per-word read (user re-font) wins over it.
  async function captureQaseedaCellRuns(context, cap) {
    var refs = [];
    cap.blockInfos.forEach(function (blk, b) {
      blk.tableInfos.forEach(function (info, t) {
        info.cells.forEach(function (c, i) {
          if (c.kind === "spacing" || !c.base) return;
          var wr = c.cell.body.getRange().getTextRanges([" "], true);
          wr.load("items");
          refs.push({
            key: b + ":" + t + ":" + i, b: b, cellKey: t + ":" + i, c: c, wr: wr,
            pack: (cap.blockInfos[b].payload.runFonts || {})[t + ":" + i] || null
          });
        });
      });
    });
    var blockPacks = cap.blockInfos.map(function () { return {}; });
    if (!refs.length) return { cells: {}, blockPacks: blockPacks };
    await context.sync();
    refs.forEach(function (r) {
      r.wr.items.forEach(function (w) { w.load("text"); w.font.load("name,size,bold,italic,color"); });
    });
    await context.sync();
    var out = {};
    refs.forEach(function (r) {
      var words = [];
      r.wr.items.forEach(function (w) {
        var txt = stripJustification(w.text || "");
        if (!txt) return;
        var f = w.font;
        var col = f && f.color;
        var hexCol = (col && /^#?[0-9a-fA-F]{6}$/.test(col)) ? col : undefined;
        words.push({
          text: txt,
          // raw* = what Office.js actually reported, BEFORE any fallback —
          // Office.js returns null for ANY property that is mixed within the
          // range, so null is "ambiguous, heal from the tag" while a real
          // value (incl. "" color = Automatic) is a state the user set.
          raw: (f && f.name != null) ? f.name : null,
          name: (f && f.name) || r.c.fontName || "",
          rawSize: (f && typeof f.size === "number" && f.size > 0) ? f.size : null,
          size: (f && f.size) || r.c.fontSize || 0,
          bold: (f && typeof f.bold === "boolean") ? f.bold : null,
          italic: (f && typeof f.italic === "boolean") ? f.italic : null,
          rawColor: (col == null) ? null : (hexCol || ""),
          color: hexCol
        });
      });
      if (!words.length) return;
      // Heal ""/mixed reads from the tag's stored pack (null = no/stale pack —
      // keep the document reads), then re-pack the result for persistence.
      // Without a pack there is nothing to heal FROM: coerce the tri-state
      // bold/italic ambiguity to off (the pre-style behavior).
      var reconciled = AshaarWord.reconcileRunWords(words, r.pack);
      var healed = (reconciled || words).map(function (w) {
        if (w.bold !== null && w.italic !== null) return w;
        var out = {};
        for (var k in w) if (w.hasOwnProperty(k)) out[k] = w[k];
        out.bold = !!out.bold; out.italic = !!out.italic;
        return out;
      });
      // Legacy-face migration (MarkSafe rename): words read from the document
      // or healed from an old tag pack may still name the retired plain
      // Kasheeda face. Normalize to the base face BEFORE re-packing/emitting —
      // the swap engine then re-decides and re-emits the current Kasheeda
      // target, so documents self-heal on their next justify.
      healed.forEach(function (w) {
        w.name = AshaarFonts.normalizeLegacyFontName(w.name);
        if (w.raw) w.raw = AshaarFonts.normalizeLegacyFontName(w.raw);
      });
      var nAmbig = words.reduce(function (a, w) { return a + (w.raw ? 0 : 1); }, 0);
      blockPacks[r.b][r.cellKey] = AshaarWord.packRunWords(healed);
      out[r.key] = {
        // Heal state for the debug dump: was a tag pack found, and did it apply?
        healInfo: !r.pack ? "no-pack" : reconciled ? "heal=" + nAmbig + "/" + words.length : "PACK-STALE",
        runs: AshaarWord.coalesceRuns(healed), align: r.c.align, indentTwips: r.c.indentTwips || 0,
        // Per-word RAW font names as read from Word (no fallback), for the debug dump.
        rawWords: words.map(function (w) {
          return (w.raw === null ? "∅" : w.raw === "" ? '""' : w.raw) + "«" + w.text.slice(0, 8) + "»";
        }).join("  ")
      };
    });
    return { cells: out, blockPacks: blockPacks };
  }

  // Force-load every font (+ Kasheeda face) referenced by the captured cell
  // runs so canvas measurement uses real metrics — measureText silently
  // substitutes an unloaded face (see memory font-measurement-model).
  async function loadOrigContentFaces(origContent, repSize) {
    if (typeof document === "undefined" || !document.fonts || !document.fonts.load) return;
    var faces = {};
    Object.keys(origContent).forEach(function (k) {
      (origContent[k].runs || []).forEach(function (r) {
        var sz = r.size || repSize;
        if (r.name) faces[sz + "pt \"" + r.name + "\""] = true;
        var kn = AshaarFonts.descriptorForFontName(r.name).kasheedaName;
        if (kn) faces[sz + "pt \"" + kn + "\""] = true;
      });
    });
    var loads = [];
    Object.keys(faces).forEach(function (s) { loads.push(document.fonts.load(s).catch(function () {})); });
    try { await Promise.all(loads); } catch (e) {}
  }

  // Repair the representative fonts from the reconciled runs. After a justify,
  // the CELL-level font read resolves through the paragraph-mark theme default
  // (observed: "Aptos"), so blk.repFont / cap.repName captured from cells would
  // pin a rebuild — and measure every natural width — in the wrong font. The
  // first reconciled run's BASE face is the trustworthy representative.
  function repairRepFonts(cap, origContent) {
    cap.blockInfos.forEach(function (blk, b) {
      var found = null;
      blk.tableInfos.forEach(function (info, t) {
        info.cells.forEach(function (c, i) {
          if (found) return;
          var oc = origContent[b + ":" + t + ":" + i];
          if (oc && oc.runs && oc.runs.length && oc.runs[0].name) found = oc.runs[0];
        });
      });
      if (found) {
        blk.repFont = AshaarFonts.descriptorForFontName(found.name).wordName || found.name;
        if (found.size) blk.repSize = found.size;
      }
    });
    for (var b = 0; b < cap.blockInfos.length; b++) {
      if (cap.blockInfos[b].repFont) {
        cap.repName = cap.blockInfos[b].repFont;
        cap.repSize = cap.blockInfos[b].repSize || cap.repSize;
        break;
      }
    }
  }

  // Rebuild each content cell's natural width (natPx) — and the cross-block
  // harmony matrix — from the reconciled per-word runs, measured per family
  // segment in its BASE face (mirrors buildContentCellOoxml's segment model).
  // The capture-time natPx measured the cell text in the CELL's reported font,
  // which after a justify is the theme default — an inflated natural drives an
  // inflated target, a changed size signature, and a runaway rebuild each apply.
  function recomputeQaseedaNaturals(cap, origContent, profile) {
    var ctx = cap.canvasCtx;
    if (!ctx) return;
    function measIn(t, nm, sz) { ctx.font = sz + "pt \"" + nm + "\""; return ctx.measureText(t).width; }
    var qCells = [];
    cap.blockInfos.forEach(function (blk, b) {
      blk.tableInfos.forEach(function (info, t) {
        info.cells.forEach(function (c, i) {
          if (!AshaarMatrix.isContentCell(c.measure)) return;
          var oc = origContent[b + ":" + t + ":" + i];
          if (oc && oc.runs && oc.runs.length) {
            var segs = [];
            oc.runs.forEach(function (r) {
              var d = AshaarFonts.descriptorForFontName(r.name);
              var fam = d.wordName || r.name;
              var sz = r.size || c.fontSize || cap.repSize;
              var prev = segs[segs.length - 1];
              if (prev && prev.fam === fam && prev.size === sz) prev.text += " " + r.text;
              else segs.push({ fam: fam, size: sz, text: r.text });
            });
            var w = segs.reduce(function (a, s) {
              return a + AshaarProfiles.applyFontCorrection(measIn(s.text, s.fam, s.size), s.fam, profile.fontCorrections);
            }, 0);
            if (segs.length > 1) w += (segs.length - 1) * measIn(" ", segs[0].fam, segs[0].size);
            c.natPx = w;
          }
          if (c.natPx != null) qCells.push({ key: c.matKey, natural: c.natPx });
        });
      });
    });
    cap.qMatrix = AshaarMatrix.buildMatrix(qCells);
  }

  // Apply a qaseeda's profile across ALL its blocks so they stay consistent. Two
  // passes: (1) SIZE — rebuild every block's table OOXML at one shared target
  // width (the only way to resize span tables; columns.setWidth garbles them —
  // see memory width-engine-rebuild-not-setwidth); same width for all bandhs →
  // same-GRID bandhs get an identical gridCol (harmony). (2) JUSTIFY — re-gather
  // the fresh bare tables and fill each cell to its box = span × (target/GRID).
  async function applyProfileToQaseeda(name, opts) {
    if (typeof Word === "undefined") { setMessage("Open this task pane inside Word to apply a qaseeda."); return false; }
    var profile = getProfile(name);
    var CELL_MARGIN_PT = 5.76;
    var MARGIN_PX = CELL_MARGIN_PT * 96 / 72;
    var strength = AshaarProfiles.normalizeStrength(profile.justify.strength);
    var fillMode = AshaarProfiles.normalizeFillMode(profile.justify.fillMode);
    var doKashida = profile.justify.mode === "kashida";
    // Both kashida and spacing fill the cell to its box; only "none"/"css" skip.
    var doFill = doKashida || profile.justify.mode === "spacing";
    var summary = "";
    var blockCount = 0, targetTwips = 0, sizeSig = "";
    // Per-word fonts of every content cell, captured before the rebuild flattens
    // them (block:table:cell → {runs, align, indentTwips}). Pass 2 re-emits these
    // so a mixed-font misra survives. Keyed by position; pass 1 and pass 2 gather
    // blocks in the same document order and rebuild from the same source, so the
    // indices line up (pass 2 also text-matches before trusting an entry).
    var origContent = {};
    // Debug-mode diagnostic: per-cell justify metrics, dumped to the Debug panel.
    // Idempotency check = apply twice, diff the two dumps; any value that grows
    // between passes (nat/target/nSp/segs) is the leak.
    var qDebug = !!(debugMode && debugMode.checked);
    var qDiags = [];
    var qMeta = {};
    // Final review I3: true only when the outer catch fires (a real pipeline
    // exception) — the honest success/failure signal returned below so
    // reRender can gate its own trailing message instead of unconditionally
    // overwriting this one.
    var qFailed = false;
    // §1 cascade descope: opts.onlyBlockTag scopes both passes to a single
    // physical block instead of every block tagged with this profile. The tag
    // string CANNOT be reused as-is across passes — pass 1 rewrites cc.tag
    // (setTagRunFonts healing the runFonts pack, on BOTH the rebuild path and
    // the sizeSig-skip path below) before pass 2 re-gathers, so a fresh gather
    // in pass 2 would see a different tag than the one captured at delegation
    // time and silently match nothing. Instead, pass 1 threads out the tag it
    // ITSELF ends up writing (blk.oldTag, the value both continuation paths
    // persist to cc.tag) via this closure var, and pass 2 filters against
    // that resolved value, falling back to the original opts.onlyBlockTag only
    // for the defensive case where pass 1 never reached that point (which
    // always exits before pass 2 runs anyway — see the two early summary
    // returns below).
    var onlyBlockResolvedTags = null;
    // Rebuild-skip cache key (fix round 1): a scoped sig (ONE block's sources)
    // is structurally different from a profile-wide sig (ALL blocks' sources),
    // so sharing the plain profile-name key would poison the cache — an
    // alternating scoped/unscoped sequence on the same profile would flap
    // needRebuild=true forever. Scoped applies get their own entry, keyed by
    // the delegation tag (unique per block; in-memory map only).
    var sigKey = (opts && opts.onlyBlockTag) ? name + "|" + opts.onlyBlockTag : name;

    try {
      // ── Pass 1: SIZE — rebuild each block at one shared target width ───────────
      await Word.run(async function (context) {
        var blocks = await gatherQaseedaBlocks(context, name);
        if (opts && opts.onlyBlockTag) {
          // True block-scope (user decision 2026-07-13): a scoped apply computes
          // width from THIS block alone — sibling poems on the same profile keep
          // their width until the next profile-wide Assign/Update re-harmonizes.
          // Accepted trade-off for per-apply cost.
          blocks = blocks.filter(function (b) { return b.tag === opts.onlyBlockTag; });
        }
        if (!blocks.length) {
          summary = (opts && opts.onlyBlockTag)
            ? "This poem is not tagged with qaseeda “" + name + "”."
            : "No blocks are tagged with qaseeda “" + name + "”.";
          return;
        }
        blockCount = blocks.length;
        var cap = await captureQaseedaTables(context, blocks, profile);
        // Just-in-time font-measurement gate (Task 9): once the distinct-face
        // set is built (captureQaseedaTables → cap.faceNames), confirm the
        // WebView can measure every one BEFORE any structural rebuild or tag
        // write below. "cancel" aborts pass 1 entirely — pass 2 never runs.
        var qGate = await ensureFacesMeasurable(cap.faceNames);
        if (qGate === "cancel") { summary = "Add the font, then Apply again."; return; }
        // Snapshot per-word fonts NOW — the rebuild below discards them. Reads
        // are reconciled against each block's tag (the source of truth; the
        // document read is lossy after a justify), and the healed pack goes
        // back into blk.oldTag so BOTH the rebuild path (wrapOoxmlControl) and
        // the skip path (cc.tag below) persist it for the next apply.
        var capRuns = await captureQaseedaCellRuns(context, cap);
        origContent = capRuns.cells;
        cap.blockInfos.forEach(function (blk, b) {
          blk.oldTag = AshaarWord.setTagRunFonts(blk.oldTag, capRuns.blockPacks[b]);
        });
        // Both continuation paths below (sizeSig skip → `blk.cc.tag = blk.oldTag`;
        // rebuild → `wrapOoxmlControl(ooxmlBody, "Ashaar Poem", blk.oldTag)`)
        // persist blk.oldTag as the block's on-disk tag, so this is the value
        // pass 2 must filter against — capture it now, before either path runs.
        if (opts && opts.onlyBlockTag) {
          onlyBlockResolvedTags = cap.blockInfos.map(function (blk) { return blk.oldTag; });
        }
        // Re-measure with the real run fonts: repair the representative fonts
        // and rebuild natPx/qMatrix from the reconciled runs — the cell-level
        // font read resolves to the theme default after a justify, which
        // inflated the naturals → target → a runaway rebuild on every apply.
        await loadOrigContentFaces(origContent, cap.repSize);
        repairRepFonts(cap, origContent);
        recomputeQaseedaNaturals(cap, origContent, profile);

        // Shared slot (px): auto-fit sizes it to hold every cell's natural text
        // (+kashida headroom, +cell margins → no wrap); fixed uses pct-of-page.
        var HEADROOM = doKashida ? 0.18 : 0.06;
        var bandhs = [];
        cap.blockInfos.forEach(function (blk) {
          blk.tableInfos.forEach(function (info) {
            var gcells = [];
            info.cells.forEach(function (c) { if (c.natPx != null) gcells.push({ natural: c.natPx, span: c.gridSpan || 1 }); });
            bandhs.push({ GRID: info.grid || 0, cells: gcells });
          });
        });
        var maxGRID = bandhs.reduce(function (m, b) { return Math.max(m, b.GRID || 0); }, 0);
        var pagePx = cap.pagePt * 96 / 72;
        var slotPx = AshaarMatrix.uniformSlotPx(bandhs, {
          mode: profile.width.mode === "fixed" ? "fixed" : "auto-fit",
          pct: profile.width.pct, pagePx: pagePx, headroom: HEADROOM, marginPx: MARGIN_PX
        });
        // One target width for all bandhs → same-GRID bandhs share an identical
        // cwt (harmony); a smaller-GRID bandh gets a wider cwt (still no wrap).
        // Capped at the page.
        targetTwips = Math.min(cap.pageTwips, Math.round(slotPx * maxGRID * 1440 / 96));
        if (targetTwips <= 0) targetTwips = cap.pageTwips;

        // Width signature: the target + each block's source. If unchanged since the
        // last apply this session, the tables are already sized correctly — skip
        // the destructive rebuild and let pass 2 just re-justify. Only a real width
        // change (mode/pct/text) triggers the rebuild.
        var sizeSigProfileStore = loadProfileStore();
        sizeSig = AshaarWord.applySizeSignature({
          targetTwips: targetTwips,
          sources: cap.blockInfos.map(function (b) { return b.source; }),
          // Structural inputs: any block's EFFECTIVE (resolver) gap/pattern
          // participates — not just the local delta, so a profile-driven gap
          // change also triggers the rebuild.
          gap: cap.blockInfos.map(function (b) {
            return AshaarProfiles.resolveSettings({ payload: b.payload, profileStore: sizeSigProfileStore, scope: { level: "poem" } }).values.gap;
          }).join(","),
          misraPattern: cap.blockInfos.map(function (b) { return b.payload.misraPattern || ""; }).join(","),
          // §9 threading (final review C2b): lineHeightPt/separatorPt are now
          // emitted by the rebuild (renderOpts above) — fold the resolved
          // values in here too, or a profile/local change to either would
          // never flip needRebuild and the rebuild would silently be skipped.
          lineHeightPt: cap.blockInfos.map(function (b) {
            return AshaarProfiles.resolveSettings({ payload: b.payload, profileStore: sizeSigProfileStore, scope: { level: "poem" } }).values.lineHeightPt;
          }).join(","),
          separatorPt: cap.blockInfos.map(function (b) {
            return AshaarProfiles.resolveSettings({ payload: b.payload, profileStore: sizeSigProfileStore, scope: { level: "poem" } }).values.separatorPt;
          }).join(","),
        });
        var needRebuild = _appliedSizeSig[sigKey] !== sizeSig;
        if (qDebug) { qMeta.targetTwips = targetTwips; qMeta.rebuild = needRebuild; qMeta.repName = cap.repName; }
        if (!needRebuild) {
          // No rebuild → the controls survive; persist the healed runFonts on them.
          cap.blockInfos.forEach(function (blk) { blk.cc.tag = blk.oldTag; });
          return; // pass 2 re-justifies the already-sized tables
        }

        // Rebuild LAST block first so earlier blocks' ranges don't shift. Render
        // BARE (justifyMode none) with the block's own structural opts + pinned
        // font/size; the pattern is unchanged, so the old tag stays valid.
        for (var bi = cap.blockInfos.length - 1; bi >= 0; bi--) {
          var blk = cap.blockInfos[bi];
          if (!blk.source.trim()) {
            // Skipped rebuild = skipped re-tag: this block's physical tag stays
            // pre-apply, so its healed oldTag would never match in pass 2 —
            // drop it from the filter set (fix round 1: the set must never
            // contain a tag pass 2 can't see).
            if (onlyBlockResolvedTags) {
              var obrIdx = onlyBlockResolvedTags.indexOf(blk.oldTag);
              if (obrIdx !== -1) onlyBlockResolvedTags.splice(obrIdx, 1);
            }
            continue;
          }
          var p = blk.payload;
          var eff = AshaarProfiles.resolveSettings({ payload: p, profileStore: sizeSigProfileStore, scope: { level: "poem" } }).values;
          var renderOpts = {
            layoutMode: eff.layoutMode,
            gapWidth: eff.gap,
            fontMode: "document",
            misraPattern: p.misraPattern || "paired",
            misraCount: Number(p.misraCount || 4),
            tatweelCount: 0,
            justifyMode: "none",
            // §9 threading (final review C2a): eff already carries the resolved
            // vertical-rhythm values — without these the bare rebuild always
            // fell back to renderForWordOoxml's defaults (auto line height,
            // 1pt separator), so a profile/local lineHeightPt/separatorPt was
            // never emitted for the flagship profiled-poem path.
            lineHeightPt: eff.lineHeightPt,
            separatorPt: eff.separatorPt
          };
          if (blk.repFont) renderOpts.fontCsName = blk.repFont;
          if (blk.repSize) renderOpts.fontSizePt = blk.repSize;
          var ooxmlBody = AshaarWord.renderForWordOoxml(blk.source, renderOpts, Ashaar, targetTwips);
          if (!ooxmlBody) continue;
          // Gap-corruption fix, Part C: the old tag is only valid for the
          // rebuilt table if the pattern really is unchanged — VERIFY instead
          // of assuming. Mint the patterns the reconstructed source produces
          // (baytCellPatternRows depends only on layoutMode) and compare with
          // the persisted payload.cells; on drift, re-mint the tag from the
          // rebuilt source (as insertPoem does) so tag and table can never
          // desync, carrying every persisted layer (profile/local/cache,
          // overrides, slotDecor, widthPt, healed runFonts). Absent stored
          // patterns (pre-pattern tags) keep the old tag as before.
          var tagForRebuild = blk.oldTag;
          var newPats = null;
          try { newPats = AshaarWord.poemCellPatterns(blk.source, { layoutMode: eff.layoutMode }, Ashaar); } catch (ePat) { newPats = null; }
          if (p.cells && newPats && !AshaarWord.cellPatternsEqual(p.cells, newPats)) {
            // Review fix (R3): overrides/slotDecor keys are POSITIONAL
            // ("tableIndex:label" / "tableIndex:slot") against the OLD
            // pattern — after a genuine shape change some of them no longer
            // exist in the new pattern's key space. Carrying them verbatim
            // left orphaned dead weight in the tag, so prune every key that
            // the re-minted patterns can't address. RECORDED HEALING
            // LIMITATION: keys that DO survive the prune may still target a
            // renumbered cell (e.g. "0:A2" landing on what used to be A3) —
            // labels are positional and the pre-drift→post-drift cell mapping
            // is unknown here, so true label migration is not possible
            // without capturing that mapping before the drift. Backlog.
            var validDriftKeys = {};
            newPats.forEach(function (patN, tiN) {
              AshaarCellMap.buildBandhCellMap(patN).forEach(function (e) {
                validDriftKeys[AshaarOverrides.overrideKey(
                  tiN, e.kind === "content" ? e.label : e.slot)] = true;
              });
            });
            var prunedKeyCount = 0;
            var pruneToValidKeys = function (mapObj) {
              var out = {};
              Object.keys(mapObj || {}).forEach(function (k) {
                if (validDriftKeys[k]) out[k] = mapObj[k];
                else prunedKeyCount++;
              });
              return out;
            };
            var driftOverrides = pruneToValidKeys(p.overrides);
            var driftSlotDecor = pruneToValidKeys(p.slotDecor);
            tagForRebuild = AshaarWord.contentControlTag(blk.source, {
              profile: p.profile, local: p.local, profileCache: p.profileCache,
              overrides: driftOverrides, slotDecor: driftSlotDecor, widthPt: p.widthPt,
              misraPattern: p.misraPattern || "paired", misraCount: Number(p.misraCount || 4)
            }, newPats);
            tagForRebuild = AshaarWord.setTagRunFonts(tagForRebuild, capRuns.blockPacks[bi]);
            // Pass 2 filters blocks by the tag pass 1 persisted — swap the
            // stale entry for the re-minted tag or pass 2 would skip this block.
            if (onlyBlockResolvedTags) {
              var driftIdx = onlyBlockResolvedTags.indexOf(blk.oldTag);
              if (driftIdx !== -1) onlyBlockResolvedTags[driftIdx] = tagForRebuild;
            }
            if (qDebug) {
              qMeta.patternDrift = (qMeta.patternDrift || 0) + 1;
              qMeta.patternDriftPrunedKeys = (qMeta.patternDriftPrunedKeys || 0) + prunedKeyCount;
            }
          }
          // Embed the content control IN the OOXML (block-level w:sdt spanning all
          // tables) so insertOoxml creates a control over the WHOLE poem. Wrapping
          // the insertOoxml-returned range with insertContentControl() instead only
          // caught row 1 on Mac Word. insertOoxml("Replace") on a whole control
          // throws, so insert just after the old control, then delete the old.
          var ooxml = AshaarWord.wrapOoxmlControl(ooxmlBody, "Ashaar Poem", tagForRebuild);
          var afterRange = blk.cc.getRange("After");
          afterRange.insertOoxml(ooxml, Word.InsertLocation.start);
          blk.cc.delete(false);
          await context.sync();
        }
      });
      // Final review I3: an honest success/failure signal so callers (reRender)
      // can gate their own trailing message instead of unconditionally
      // overwriting this one.
      if (summary) { setMessage(summary); return false; }

      // ── Pass 2: JUSTIFY — fill each cell of the fresh tables to its box ────────
      var changed = 0, coloured = 0;
      await Word.run(async function (context) {
        var blocks = await gatherQaseedaBlocks(context, name);
        if (opts && opts.onlyBlockTag) {
          var filterTags = onlyBlockResolvedTags || [opts.onlyBlockTag];
          blocks = blocks.filter(function (b) { return filterTags.indexOf(b.tag) !== -1; });
        }
        if (!blocks.length) return;
        // Ensure the rebuilt (SDT-created) controls show the block outline.
        blocks.forEach(function (cc) { cc.appearance = "BoundingBox"; });
        var cap = await captureQaseedaTables(context, blocks, profile);
        if (!cap.canvasCtx) { summary = "Canvas unavailable; cannot measure."; return; }
        // §9 threading (final review C2c): resolved per-block lineHeightPt for
        // the justify-pass emitters below — pass 1's rebuild emits it into the
        // bare tables, but misraRunsXml (used by both the no-fill early return
        // and emitContentCell) re-emits every misra paragraph from scratch and
        // would silently drop it without this.
        var qProfileStore = loadProfileStore();

        // Force-load every ORIGINAL per-run font (+ Kasheeda face), then repair
        // the rep baseline and rebuild natPx/qMatrix from the reconciled runs —
        // same reasons as pass 1 (rebuilt cells report only the flattened font;
        // un-rebuilt justified cells read the theme default at cell level).
        await loadOrigContentFaces(origContent, cap.repSize);
        repairRepFonts(cap, origContent);
        recomputeQaseedaNaturals(cap, origContent, profile);
        var canvasCtx = cap.canvasCtx, repName = cap.repName, repSize = cap.repSize, qMatrix = cap.qMatrix;

        // Debug tint for the residual-spacing runs (w:shd hex, no '#').
        var spcColorHex = ((profile.debugColors && profile.debugColors.space) || "").replace(/^#/, "");

        // Every content cell is (re)written as run-aware OOXML so each word keeps
        // its ORIGINAL font — the SIZE rebuild flattened the whole block to one
        // representative font, and a plain insertText would flatten it again.
        // Collected here, written after the spacing batch, one sync per cell.
        var cellPlans = [];
        var preps = []; // prepped content cells awaiting target resolution

        // Canvas measurement cache — the adaptive-harmony scan re-fills every
        // cell once per candidate target, so identical (text,font,size) spans
        // are measured hundreds of times. Fonts are loaded before any measure
        // (loadOrigContentFaces above), so cached widths stay valid.
        var measCache = Object.create(null);

        // Prepare one content cell for filling: rebuild its original runs,
        // regroup into family segments, measure naturals, resolve overrides —
        // everything that does NOT depend on the fill target. Returns
        // { xml } directly for the no-fill case, else a prep whose
        // fillToTarget(cTargetPx, strict) elongates each segment by its own
        // mechanism toward the target (strict = never overshoot, for the
        // adaptive mode where every line must land AT the shared target).
        function prepContentCell(c, info, key, colPx, lineHeightPt) {
          var repFallback = c.fontName || repName;
          var sizeFallback = c.fontSize || repSize;
          // Original per-word runs; fall back to a single run when capture missed
          // or the reconstructed text no longer matches (rebuild changed words).
          var oc = origContent[key];
          var joined = (oc && oc.runs) ? oc.runs.map(function (r) { return r.text; }).join(" ").replace(/\s+/g, " ").trim() : "";
          var origRuns;
          if (oc && oc.runs && oc.runs.length && joined === c.base.replace(/\s+/g, " ").trim()) {
            origRuns = oc.runs.map(function (r) {
              return { text: r.text, name: r.name || repFallback, size: r.size || sizeFallback, color: r.color, bold: !!r.bold, italic: !!r.italic };
            });
          } else {
            origRuns = [{ text: c.base, name: repFallback, size: sizeFallback, color: undefined, bold: false, italic: false }];
          }
          var align = (oc && oc.align) || c.align || "center";
          var indentTwips = (oc && oc.indentTwips) || 0;
          var rep0Size = origRuns[0].size;
          function measIn(text, nm, sz) {
            var mk = sz + "|" + nm + "|" + text;
            if (mk in measCache) return measCache[mk];
            canvasCtx.font = sz + "pt \"" + nm + "\"";
            return (measCache[mk] = canvasCtx.measureText(text).width);
          }

          // Regroup consecutive runs of the same FONT FAMILY into segments: both
          // Jameel faces map to one wordName (so a re-applied Jameel word, read
          // back with mixed base+Kasheeda faces, regroups into one segment and
          // re-derives its swap from clean text — idempotent); generic fonts key
          // on their own name so Amiri ≠ Fatemi. Each segment fills by its family's
          // mechanism. Natural widths are measured in the BASE face so a widened
          // Kasheeda read-back doesn't inflate the target on re-apply.
          function baseFaceOf(name) { return AshaarFonts.descriptorForFontName(name).wordName || name; }
          var segs = [];
          origRuns.forEach(function (r) {
            var d = AshaarFonts.descriptorForFontName(r.name);
            var fam = d.wordName || r.name;
            var prev = segs[segs.length - 1];
            if (prev && prev.fam === fam && prev.size === r.size && prev.color === r.color &&
                prev.bold === r.bold && prev.italic === r.italic) {
              prev.text += " " + r.text;
            } else {
              segs.push({ fam: fam, mech: d.mechanism, desc: d, name: r.name, size: r.size, color: r.color, bold: r.bold, italic: r.italic, text: r.text });
            }
          });
          var interSpacePx = measIn(" ", baseFaceOf(segs[0].name), segs[0].size);

          // Flatten per-segment output runs, re-inserting the inter-segment space
          // (in the left segment's base face) that coalescing drops — so words
          // across a font boundary keep their gap instead of touching.
          function flattenSegs(segOut) {
            var out = [];
            segOut.forEach(function (arr, si) {
              if (si > 0) out.push({ text: " ", csName: baseFaceOf(segs[si - 1].name), sizePt: segs[si - 1].size, color: undefined });
              (arr || []).forEach(function (rr) { out.push(rr); });
            });
            return out;
          }

          // No fill (justify none/css) → re-emit each segment's text unchanged
          // (fonts + inter-segment spaces preserved), no elongation.
          if (!doFill || !(colPx > 0)) {
            var passOut = flattenSegs(segs.map(function (seg) {
              return [{ text: seg.text, csName: baseFaceOf(seg.name), sizePt: seg.size, color: seg.color, bold: seg.bold, italic: seg.italic }];
            }));
            return { xml: AshaarWord.misraRunsXml(passOut, align, rep0Size, { indentTwips: indentTwips, lineHeightPt: lineHeightPt }) };
          }

          // Natural width in BASE faces + the inter-segment spaces. Stable on
          // re-apply. Shared fill target (a per-cell width override wins): cell-fit
          // → toward the CELL EDGE; natural-fit → toward the position's HARMONY;
          // adaptive → the qaseeda-wide reachable target (computed by the caller).
          var cNatural = segs.reduce(function (a, seg) { return a + measIn(seg.text, baseFaceOf(seg.name), seg.size); }, 0)
            + Math.max(0, segs.length - 1) * interSpacePx;
          var cOv = c.ovKey ? info.overrides[c.ovKey] : null;
          // Misra width precedence: cell override > bandh (tag) > qaseeda profile
          // > computed (harmony/cell-fit/adaptive). Base carries the bandh/qaseeda level.
          var baseWidthPt = (info.bandhWidthPt != null) ? info.bandhWidthPt
            : (profile.justify.widthPt != null ? profile.justify.widthPt : null);
          var cRes = AshaarOverrides.resolveCellOverride({ strength: strength, fillMode: fillMode, widthPt: baseWidthPt }, cOv);
          var cPhi = AshaarWord.strengthToElongationShare(cRes.strength);
          // Residual-space cap: per-cell Cap lift wins; otherwise 1em/gap in the
          // flush modes — the qaseeda apply is the HARMONY context, where a
          // visibly short line misaligns the whole column and reads worse than
          // roomier word gaps. Adaptive inverts that bargain: its target is BY
          // CONSTRUCTION reachable by every line, so spacing is capped tight
          // (0.25em/gap) and stays a trim, not a filler.
          var cCapEm = cRes.capEm != null ? cRes.capEm : (fillMode === "adaptive" ? 0.25 : 1.0);
          var cMaxPos = AshaarWord.strengthToMaxPositions(cRes.strength);

          // Elongate each segment by its OWN mechanism toward a proportional share
          // of the target; generic segments then absorb the slack the discrete
          // mechanisms (Jameel swap / Mehr tatweel) leave. Whitespace-shaping
          // segments never get tatweels. Under a spacing profile nothing elongates
          // and the capped gap spacing does all the filling. strict=true forbids
          // overshoot (adaptive: every line lands AT the shared target and the
          // ≤capEm spacing trims the rest exactly).
          function fillToTarget(cTarget, strict) {
            var extra = Math.max(0, cTarget - cNatural);
            var segOut = new Array(segs.length);
            var genericIdx = [];
            var nonGenAchieved = Math.max(0, segs.length - 1) * interSpacePx; // inter-seg spaces
            // Overshoot budget for the discrete mechanisms (Jameel swap / Mehr
            // tatweel): the gap between the fill target and the cell box. Discrete
            // steps may land past the target (closer beats short — alignment), but
            // the SUM of overshoots across segments must stay inside the box, so
            // each segment consumes what it actually oversteps.
            var swapSlack = strict ? 0 : Math.max(0, colPx - cTarget);
            segs.forEach(function (seg, si) {
              var bf = baseFaceOf(seg.name);
              var segNat = measIn(seg.text, bf, seg.size);
              var subTarget = cNatural > 0 ? segNat + extra * (segNat / cNatural) : segNat;
              var segMaxPx = strict ? null : subTarget + swapSlack;
              if (doKashida && seg.mech === "font-swap") {
                var jw = seg.desc.kasheedaName || bf;
                var fss = AshaarKashidaFontswap.splitSpans(seg.text);
                var wb = [], ww = [];
                fss.forEach(function (s) { wb.push(measIn(s, bf, seg.size)); ww.push(measIn(s, jw, seg.size)); });
                // What the residual spacing can still close for this segment
                // (capEm per gap) — overshoot yields to spacing, which lands exactly.
                var swSpaceClose = cCapEm * (seg.size * 96 / 72) * Math.max(0, seg.text.split(" ").length - 1);
                var sel = AshaarKashidaFontswap.selectSwapRuns(fss, wb, ww, subTarget, segMaxPx, swSpaceClose);
                // asciiName pins ascii+hAnsi to the BASE face on every fasl (cs
                // carries the actual face). Word renders Arabic via cs, so the
                // Kasheeda face still shows — but Font.name (=ascii) reads back
                // ONE family for the whole word. Without this, a partially-swapped
                // word reads "" (mixed ascii) on the next apply and is
                // misclassified as generic: U+0640 shatter + runaway growth.
                segOut[si] = sel.runs.map(function (rr) { return { text: rr.text, csName: rr.swap ? jw : bf, asciiName: bf, sizePt: seg.size, color: seg.color, bold: seg.bold, italic: seg.italic }; });
                var swAch = fss.reduce(function (a, s, k) { return a + (sel.runs[k].swap ? ww[k] : wb[k]); }, 0);
                nonGenAchieved += swAch;
                swapSlack = Math.max(0, swapSlack - Math.max(0, swAch - subTarget));
              } else if (doKashida && seg.mech === "tatweel") {
                var mr = seg.desc.tatweelRules || {};
                var iso = {}, fin = {};
                (mr.isolatedInto || []).forEach(function (x) { iso[x] = true; });
                (mr.finalInto || []).forEach(function (x) { fin[x] = true; });
                var parts = seg.text.split(" "), toks = [];
                parts.forEach(function (wd, k) { if (k) toks.push(" "); toks.push(wd); });
                var elong = toks.map(function (t) { return t !== " " ? AshaarWord.mehrElongate(t, iso, fin) : t; });
                var mwb = [], mww = [];
                toks.forEach(function (t, k) { mwb.push(measIn(t, bf, seg.size)); mww.push(measIn(elong[k], bf, seg.size)); });
                var mSpaceClose = cCapEm * (seg.size * 96 / 72) * Math.max(0, seg.text.split(" ").length - 1);
                var msel = AshaarKashidaFontswap.selectSwapRuns(toks, mwb, mww, subTarget, segMaxPx, mSpaceClose);
                var mtext = msel.runs.map(function (rr, k) { return (rr.swap && mww[k] > mwb[k]) ? elong[k] : toks[k]; }).join("");
                segOut[si] = [{ text: mtext, csName: bf, sizePt: seg.size, color: seg.color, bold: seg.bold, italic: seg.italic }];
                var mAch = measIn(mtext, bf, seg.size);
                nonGenAchieved += mAch;
                swapSlack = Math.max(0, swapSlack - Math.max(0, mAch - subTarget));
              } else if (!doKashida || seg.mech === "whitespace") {
                segOut[si] = [{ text: seg.text, csName: bf, sizePt: seg.size, color: seg.color, bold: seg.bold, italic: seg.italic }];
                nonGenAchieved += segNat;
              } else {
                genericIdx.push(si); // generic — elongated jointly below
              }
            });
            if (genericIdx.length) {
              var primRuns = genericIdx.map(function (si) {
                var seg = segs[si], fstr = seg.size + "pt \"" + seg.name + "\"";
                return { text: seg.text, fontSize: seg.size, fontProfile: null,
                  measure: function (s) { canvasCtx.font = fstr; return canvasCtx.measureText(s).width; } };
              });
              var genNat = primRuns.reduce(function (a, r) { return a + r.measure(r.text); }, 0);
              var genTarget = Math.max(genNat, cTarget - nonGenAchieved);
              var conc = AshaarJustify.justifyRunsConcentrated(primRuns, genTarget, { perPositionEm: 0.5, maxPositions: cMaxPos });
              genericIdx.forEach(function (si, k) {
                segOut[si] = [{ text: conc.runs[k].text, csName: segs[si].name, sizePt: segs[si].size, color: segs[si].color, bold: segs[si].bold, italic: segs[si].italic }];
              });
            }
            var outRuns = flattenSegs(segOut);
            var achievedTot = outRuns.reduce(function (a, rr) { return a + measIn(rr.text, rr.csName, rr.sizePt); }, 0);
            var gaps = outRuns.reduce(function (a, rr) { return a + (rr.text.split(" ").length - 1); }, 0);
            return { outRuns: outRuns, achievedTot: achievedTot, gaps: gaps };
          }

          return {
            c: c, key: key, colPx: colPx, oc: oc, align: align, indentTwips: indentTwips,
            rep0Size: rep0Size, segs: segs, cNatural: cNatural, cOv: cOv, cRes: cRes,
            cPhi: cPhi, cCapEm: cCapEm, fillToTarget: fillToTarget, lineHeightPt: lineHeightPt
          };
        }

        // Fill a prepped cell to its resolved target and emit the OOXML. The
        // residual is closed with pixel-exact per-gap character spacing (rPr
        // w:spacing on single-space runs): gaps come out EVEN, the line lands on
        // target to the twip, and the cell text round-trips clean (no injected
        // glyphs to strip). capEm accept-short cap still applies.
        function emitContentCell(p, adaptT) {
          var cTarget, strict = false;
          if (p.cRes.widthPt != null) cTarget = p.cRes.widthPt * 96 / 72;
          else if (fillMode === "adaptive" && adaptT != null) { cTarget = adaptT; strict = true; }
          else if (fillMode === "cell-fit") cTarget = AshaarMatrix.cellFitBudget(p.cNatural, p.colPx, p.cPhi);
          else {
            var cReach = Math.max(p.cNatural, p.colPx - 0.28 * p.rep0Size * 96 / 72);
            var cWpos = qMatrix[p.c.matKey] || p.cNatural;
            cTarget = AshaarMatrix.naturalFitTarget(cWpos, cReach, p.cPhi);
          }
          cTarget = Math.min(cTarget, p.colPx); // no-wrap invariant
          var fr = p.fillToTarget(cTarget, strict);
          var spread = AshaarResidual.spreadResidualSpacing(fr.outRuns, cTarget - fr.achievedTot, p.rep0Size * 96 / 72, p.cCapEm);
          if (spcColorHex) spread.runs.forEach(function (rr) { if (rr.spacingTwips > 0) rr.shdFill = spcColorHex; });
          if (qDebug) qDiags.push({
            key: p.key, colPx: Math.round(p.colPx), nat: Math.round(p.cNatural),
            target: Math.round(cTarget), achieved: Math.round(fr.achievedTot), addPx: Math.round(spread.appliedPx),
            fin: Math.round(fr.achievedTot + spread.appliedPx),
            segs: p.segs.map(function (s) { return (s.fam || "?") + "/" + s.mech; }).join(" · "),
            text: (p.c.base || "").slice(0, 16),
            cellFont: p.c.fontName === "" ? '""' : p.c.fontName,
            natPx: Math.round(p.c.natPx || 0),
            healInfo: (p.oc && p.oc.healInfo) || "?",
            ov: p.cOv ? "OVERRIDE " + JSON.stringify(p.cOv) : "",
            rawWords: (p.oc && p.oc.rawWords) || "(no capture)"
          });
          return AshaarWord.misraRunsXml(spread.runs, p.align, p.rep0Size, { indentTwips: p.indentTwips, lineHeightPt: p.lineHeightPt });
        }

        cap.blockInfos.forEach(function (blk, bIdx) {
          // §9 threading (final review C2c): this block's resolved line
          // height, so the justify-pass emitters below re-emit the same
          // vertical rhythm pass 1's rebuild used (not "auto").
          var blkEff = AshaarProfiles.resolveSettings({ payload: blk.payload, profileStore: qProfileStore, scope: { level: "poem" } }).values;
          // §4 transition-clear consumption gate: pending clears apply ONLY to
          // the block that recorded them. Override keys ("0:A1") repeat across
          // poems, so without this a retained map (render failed after a tag
          // write) would blacken the colliding cell of whatever poem renders
          // next. Identity = tagIdentity (tag minus the runFonts heal), the
          // one field the pipeline mutates between Apply's write (blockId
          // source) and this pass's capture (blk.oldTag).
          var blkClears = (_pendingColorClears.blockId &&
            AshaarWord.tagIdentity(blk.oldTag) === _pendingColorClears.blockId)
            ? _pendingColorClears.keys : null;
          blk.tableInfos.forEach(function (info, tIdx) {
            // Cell box comes from the width we just rebuilt to: cwt = target/GRID.
            var cwtPx = info.grid > 0 ? (targetTwips / info.grid) * 96 / 1440 : 0;
            info.cells.forEach(function (c, cIdx) {
              if (c.kind === "spacing") {
                // Decorate (not justify) a structural gap.
                var pDecor = c.slot ? (profile.spacingDecor || {})[c.slot] : null;
                var oDecor = c.decorKey ? info.slotDecor[c.decorKey] : null;
                var decor = AshaarOverrides.resolveSlotDecor(pDecor, oDecor);
                c.cell.body.clear();
                if (decor.symbol) {
                  c.cell.body.insertText(decor.symbol, Word.InsertLocation.replace);
                  c.cell.body.font.color = decor.color || "black";
                }
                // shadingColor rejects "" / "No color"; use "#FFFFFF" to clear.
                c.cell.shadingColor = decor.fill || "#FFFFFF";
                c.cell.body.paragraphs.getFirst().alignment = Word.Alignment.centered;
                changed++;
                return;
              }
              if (!c.base) return;
              // Fill box = span × cwt − cell margins (the text area we rebuilt to).
              var colPx = cwtPx > 0
                ? Math.max(1, (c.gridSpan || 1) * cwtPx - 2 * MARGIN_PX)
                : Math.max(1, (c.cell.columnWidth || 0) - 2 * CELL_MARGIN_PT) * 96 / 72;
              var p = prepContentCell(c, info, bIdx + ":" + tIdx + ":" + cIdx, colPx, blkEff.lineHeightPt);
              if (!p) return;
              // §4 fill/color: the raw override (unfiltered by resolveCellOverride,
              // which only carries strength/widthPt/capEm) — read straight off the
              // tag so the write loop below can (re)assert or clear it.
              var cellOv = c.ovKey ? (info.overrides[c.ovKey] || {}) : {};
              var colorClear = !!(blkClears && c.ovKey && blkClears[c.ovKey]);
              if (p.xml) cellPlans.push({ cell: c.cell, ooxml: p.xml, ov: cellOv, colorClear: colorClear }); // no-fill emit
              else { p.colorClear = colorClear; preps.push(p); }
            });
          });
        });

        // Adaptive harmony: one shared target every misra can REACH — the
        // largest T (≤ the smallest cell box) where each cell's own kashida
        // plus at most capEm/gap of spacing lands ON T. Equal line widths and
        // small gaps, at the cost of stopping short of the cell edge when the
        // font can't stretch (Jameel + harakat). Width-overridden cells keep
        // their own target and sit out of the search.
        var adaptT = null;
        if (fillMode === "adaptive" && preps.length) {
          var partic = preps.filter(function (p) { return p.cRes.widthPt == null; });
          if (partic.length) {
            var aHi = Math.min.apply(null, partic.map(function (p) { return p.colPx; }));
            var aLo = Math.max.apply(null, partic.map(function (p) { return p.cNatural; }));
            adaptT = AshaarMatrix.adaptiveSharedTarget(aHi, aLo, function (T) {
              return partic.every(function (p) {
                if (T <= p.cNatural + 0.5) return true; // already there
                var r = p.fillToTarget(T, true);
                return r.achievedTot + p.cCapEm * (p.rep0Size * 96 / 72) * r.gaps >= T - 0.5;
              });
            });
            if (qDebug) qMeta.adaptT = adaptT;
          }
        }

        preps.forEach(function (p) {
          var x = emitContentCell(p, adaptT);
          if (x) cellPlans.push({ cell: p.c.cell, ooxml: x, ov: p.cOv || {}, colorClear: !!p.colorClear });
        });
        await context.sync(); // commit the spacing-cell decorations

        // Write each content cell's run-aware OOXML: clear + insert. One sync per
        // cell so a single OOXML failure leaves that cell as its bare rebuild
        // instead of aborting the whole batch.
        var writeFails = 0;
        for (var cpi = 0; cpi < cellPlans.length; cpi++) {
          var cp = cellPlans[cpi];
          try {
            cp.cell.body.clear();
            cp.cell.body.insertOoxml(AshaarWord.wrapOoxml(cp.ooxml), Word.InsertLocation.replace);
            // §4 fill/color: must come AFTER the clear+insert above — clear()
            // wipes any formatting set on the (now-empty) body, so setting
            // body.font.color earlier would have no effect once the new runs
            // land. shadingColor rejects "" / "No color"; "#FFFFFF" clears it
            // (same quirk as the spacing-cell decor branch above). Color has
            // no clear value, so a JUST-REMOVED color override (recorded in
            // _pendingColorClears by the Apply that deleted it, block-scoped
            // at plan time) resets to "black" — necessary because the capture
            // reads live run colors as "original", so the old override color
            // is baked into this cell's re-emitted runs. Accepted limitations:
            // (1) "black", not any pre-override manual text color (no source
            // data to recover it); (2) if THIS cell's write fails (catch
            // below), the queued reset dies with the batch but the success
            // tail still consumes the pending clear — the stale color
            // survives and is re-baked by the next capture. Narrow, accepted.
            var ov = cp.ov || {};
            cp.cell.shadingColor = ov.fill || "#FFFFFF";
            if (ov.color) cp.cell.body.font.color = ov.color;
            else if (cp.colorClear) cp.cell.body.font.color = "black";
            await context.sync();
            changed++;
          } catch (e) { writeFails++; /* leave the cell as its bare rebuild */ }
        }
        if (qDebug) qMeta.writeFails = writeFails;

        // Ground truth for the font round-trip: read back the FIRST written
        // cell's stored OOXML and list the distinct <w:rFonts> Word actually
        // kept. If ascii="<base face>" is missing here, Word rewrote our runs
        // on insert; if it's present but the next capture still reads ""/default,
        // Office.js Font.name doesn't read ascii for rtl runs.
        if (qDebug && cellPlans.length) {
          try {
            var gtOox = cellPlans[0].cell.body.getOoxml();
            await context.sync();
            var gtFonts = (gtOox.value.match(/<w:rFonts[^>]*\/>/g) || []);
            var seen = {}, uniq = [];
            gtFonts.forEach(function (s) { if (!seen[s]) { seen[s] = true; uniq.push(s); } });
            qMeta.storedRFonts = uniq;
          } catch (e) { qMeta.storedRFonts = ["(getOoxml failed: " + (e && e.message) + ")"]; }
        }

        // Debug colouring: tint inserted tatweels / micro-spaces so they're visible.
        var tatColor = (profile.debugColors && profile.debugColors.tatweel) || "";
        var spcColor = (profile.debugColors && profile.debugColors.space) || "";
        if (tatColor || spcColor) {
          var hits = [];
          cap.blockInfos.forEach(function (blk) {
            blk.tableInfos.forEach(function (info) {
              info.cells.forEach(function (c) {
                if (tatColor) { var st = c.cell.body.search("ـ"); st.load("items"); hits.push({ s: st, color: tatColor, hl: false }); }
                if (spcColor) {
                  var sh = c.cell.body.search(" "); sh.load("items"); hits.push({ s: sh, color: spcColor, hl: true });
                  var sn = c.cell.body.search(" "); sn.load("items"); hits.push({ s: sn, color: spcColor, hl: true });
                }
              });
            });
          });
          await context.sync();
          // Spaces have no ink, so tint them with the HIGHLIGHT (background); tatweels
          // are ink, so use the font color.
          hits.forEach(function (h) {
            h.s.items.forEach(function (r) {
              if (h.hl) r.font.highlightColor = h.color; else r.font.color = h.color;
              coloured++;
            });
          });
          await context.sync();
        }
      });

      // Remember the width we sized to, so a later justify-only apply (strength,
      // fill mode, per-cell override) skips the destructive rebuild.
      // Fix round 2: write the scoped entry under the tag physically ON the
      // block after healing (onlyBlockResolvedTags[0]) — that is what the next
      // scoped delegation will read and key by. Writing under the pre-heal key
      // left the scoped cache permanently cold (the tag churns every apply).
      // If the empty-source skip spliced the entry out, the physical tag never
      // changed, so the pre-heal sigKey is still the right one — keep it.
      if (opts && opts.onlyBlockTag && onlyBlockResolvedTags && onlyBlockResolvedTags.length) {
        sigKey = name + "|" + onlyBlockResolvedTags[0];
      }
      if (sizeSig) _appliedSizeSig[sigKey] = sizeSig;
      summary = (opts && opts.onlyBlockTag)
        ? "Applied to this poem; justified " + changed + " cell(s)"
          + (coloured ? "; coloured " + coloured + " artifact(s)" : "") + "."
        : "Applied qaseeda “" + name + "” to " + blockCount + " block(s); justified " + changed + " cell(s)"
          + (coloured ? "; coloured " + coloured + " artifact(s)" : "") + ".";
    } catch (error) {
      summary = "Apply failed: " + describeError(error);
      qFailed = true;
    }
    if (qDebug && debugOutput) {
      var qHead = "dbg=v5(gap-spacing)  target=" + (qMeta.targetTwips || 0) + "tw  rebuild=" + (qMeta.rebuild ? "YES" : "no")
        + (qMeta.patternDrift ? "  PATTERN-DRIFT=" + qMeta.patternDrift : "")
        + "  repName=" + (qMeta.repName || "?") + "  writeFails=" + (qMeta.writeFails || 0)
        + (qMeta.adaptT != null ? "  adaptT=" + qMeta.adaptT + "px" : "")
        // Registry beacon: proves WHICH bundle the WebView is actually running
        // (stale-cache flapping burned a MarkSafe test session, 2026-07-13).
        + "\nregistry: jameelKasheeda=" + AshaarFonts.kasheedaNameOf("jameel")
        + "\nkey      col   nat   tgt  achv  +px   fin  segs (family/mechanism)";
      debugOutput.textContent = !qDiags.length ? "(no content cells justified)" :
        qHead + "\n" + qDiags.map(function (d) {
          return [String(d.key).padEnd(8), String(d.colPx).padStart(4),
            String(d.nat).padStart(5), String(d.target).padStart(5),
            String(d.achieved).padStart(5), String(d.addPx).padStart(4), String(d.fin).padStart(5),
            "  " + d.segs + "  «" + d.text + "»",
            "\n         cellFont=" + d.cellFont + " capNatPx=" + d.natPx + " " + d.healInfo + (d.ov ? " " + d.ov : ""),
            "\n         words: " + d.rawWords].join(" ");
        }).join("\n")
        + (qMeta.storedRFonts ? "\n\nstored rFonts (cell " + (qDiags[0] ? qDiags[0].key : "?") + "):\n  "
           + qMeta.storedRFonts.join("\n  ") : "");
    }
    setMessage(summary);
    return !qFailed;
  }


  // ── Custom fonts (AshaarFontStore) ─────────────────────────────────────────
  // Let the user load a font from their machine so the justify canvas measures
  // the real outlines (see font-store.js). Registered under the exact name Word
  // reports for the text; persisted in IndexedDB and re-registered on startup.
  var fontUpload = document.getElementById("font-upload");
  var fontUploadName = document.getElementById("font-upload-name");
  var fontUploadStatus = document.getElementById("font-upload-status");
  var fontList = document.getElementById("font-list");

  function setFontUploadStatus(text, kind) {
    if (!fontUploadStatus) return;
    fontUploadStatus.textContent = text;
    fontUploadStatus.className = "qaseeda-font-status" + (kind === "ok" ? " is-ok" : kind === "warn" ? " is-warn" : "");
  }

  function renderFontList(fonts) {
    if (!fontList) return;
    fontList.innerHTML = "";
    (fonts || []).forEach(function (f) {
      var li = document.createElement("li");
      var span = document.createElement("span");
      span.textContent = f.family + (f.filename ? " (" + f.filename + ")" : "");
      var btn = document.createElement("button");
      btn.type = "button";
      btn.className = "button--secondary font-remove";
      btn.textContent = "×";
      btn.setAttribute("aria-label", "Remove " + f.family);
      btn.addEventListener("click", function () { removeFont(f.family); });
      li.appendChild(span);
      li.appendChild(btn);
      fontList.appendChild(li);
    });
  }

  function refreshFontList() {
    if (typeof AshaarFontStore === "undefined") return Promise.resolve();
    return AshaarFontStore.listFonts().then(renderFontList, function () {});
  }

  // Just-in-time font-measurement gate (Task 9). Force-load every face; when
  // one is invisible to the WebView (Word renders it but canvas can't measure
  // it) surface the inline prompt. Resolves "ok" when all faces measure,
  // "continue" when the user accepts fallback metrics, "cancel" when they go
  // add the font. See memory font-measurement-model.
  async function ensureFacesMeasurable(faceNames) {
    var missing = [];
    for (var i = 0; i < faceNames.length; i++) {
      var f = faceNames[i];
      try { await document.fonts.load('16pt "' + f + '"'); } catch (e) {}
      if (!document.fonts.check('16pt "' + f + '"')) missing.push(f);
    }
    if (!missing.length) return "ok";
    return await new Promise(function (resolve) {
      var box = document.getElementById("sp-font-prompt");
      document.getElementById("sp-font-prompt-name").textContent = missing.join('", "');
      box.hidden = false;
      document.getElementById("sp-font-prompt-add").onclick = function () {
        box.hidden = true;
        var strip = document.getElementById("fonts-strip");
        strip.open = true;
        document.getElementById("font-upload-name").value = missing[0];
        resolve("cancel");
      };
      document.getElementById("sp-font-prompt-continue").onclick = function () {
        box.hidden = true;
        resolve("continue");
      };
    });
  }

  // Distinct font faces used by the current block's cells (read-only capture
  // for the JIT gate above) — no tag write, so a "cancel" leaves nothing to
  // undo. Errors (transient selection state) degrade to an empty face list,
  // which makes the gate resolve "ok" and let the caller's own routing surface
  // any real problem.
  async function collectBlockFaceNames() {
    var names = {};
    if (typeof Word === "undefined") return [];
    try {
      await Word.run(async function (context) {
        var cc = await findBlockAt(context);
        var tables = cc.getRange().tables;
        tables.load("items");
        await context.sync();
        tables.items.forEach(function (tbl) { tbl.rows.load("items"); });
        await context.sync();
        tables.items.forEach(function (tbl) { tbl.rows.items.forEach(function (row) { row.cells.load("items"); }); });
        await context.sync();
        var cells = [];
        tables.items.forEach(function (tbl) { tbl.rows.items.forEach(function (row) { row.cells.items.forEach(function (cell) {
          cell.body.font.load("name");
          cells.push(cell);
        }); }); });
        await context.sync();
        cells.forEach(function (cell) { if (cell.body.font && cell.body.font.name) names[cell.body.font.name] = true; });
      });
    } catch (e) { /* transient selection — gate degrades to no known faces */ }
    return Object.keys(names);
  }

  // Auto-detect the family from the picked file and prefill the name field.
  function onFontFilePicked() {
    setFontUploadStatus("", "");
    var file = fontUpload && fontUpload.files && fontUpload.files[0];
    if (!file || typeof AshaarFontStore === "undefined") return;
    file.arrayBuffer().then(function (buf) {
      var names = AshaarFontStore.parseNames(buf);
      if (names && names.family) {
        fontUploadName.value = names.family;
      } else if (!fontUploadName.value) {
        fontUploadName.value = (file.name || "").replace(/\.[^.]+$/, "");
        setFontUploadStatus("Couldn’t read the font’s name — confirm it matches what Word shows (Verify at cursor).", "warn");
      }
    }, function () {});
  }

  async function addFont() {
    if (typeof AshaarFontStore === "undefined") return;
    var file = fontUpload && fontUpload.files && fontUpload.files[0];
    var family = (fontUploadName.value || "").trim();
    if (!file) { setFontUploadStatus("Choose a font file first.", "warn"); return; }
    if (!family) { setFontUploadStatus("Enter the name Word uses for this font.", "warn"); return; }
    try {
      var res = await AshaarFontStore.addUserFont(family, file);
      setFontUploadStatus("Loaded “" + family + "”" +
        (res.persisted ? " — saved for future sessions." : " — this session only (storage unavailable)."), "ok");
      fontUpload.value = "";
      // Newly registered/replaced font changes the measurement basis for any
      // probe/calibration already cached under this family name — bust so the
      // next Apply re-probes/re-calibrates against the real, now-loaded font.
      if (_tuneCache) _tuneCache.bustAll();
      await refreshFontList();
    } catch (e) {
      setFontUploadStatus("Couldn’t load that file — is it a valid .ttf/.otf/.woff?", "warn");
    }
  }

  function removeFont(family) {
    if (typeof AshaarFontStore === "undefined") return;
    AshaarFontStore.deleteFont(family).then(function () {
      setFontUploadStatus("Removed “" + family + "”. Reload the add-in to fully unload it.", "");
      return refreshFontList();
    }, function () {});
  }

  // Compare the registered name against the font Word applies at the cursor.
  async function verifyFontAtCursor() {
    if (typeof Word === "undefined") { setFontUploadStatus("Open in Word to verify.", "warn"); return; }
    var want = (fontUploadName.value || "").trim();
    var fontName = "";
    try {
      await Word.run(async function (context) {
        var sel = context.document.getSelection();
        sel.font.load("name");
        await context.sync();
        fontName = sel.font.name || "";
      });
    } catch (e) { /* ignore */ }
    if (!fontName) { setFontUploadStatus("Put the cursor in the styled text, then Verify.", "warn"); return; }
    if (document.fonts && document.fonts.load) { try { await document.fonts.load("16pt \"" + fontName + "\""); } catch (e) {} }
    var resolves = fontAvailable(fontName);
    if (want !== fontName) {
      fontUploadName.value = fontName;
      setFontUploadStatus("Word uses “" + fontName + "” here — I set the name to match. Pick its file and click Add.", "warn");
    } else if (resolves) {
      setFontUploadStatus("✓ “" + fontName + "” matches and resolves — justify will be accurate.", "ok");
    } else {
      setFontUploadStatus("“" + fontName + "” isn’t resolvable yet — pick its font file above and click Add.", "warn");
    }
  }

  // After embedding the SDT directly in the OOXML (AshaarWord.wrapOoxmlControl)
  // there is no insertContentControl() return value to configure — unlike the
  // insertContentControl() path, which hands back the control it just made.
  // Locate the freshly inserted control inside the RANGE insertOoxml returned
  // (review hardening): contentControlTag is fully deterministic (no nonce),
  // so two same-settings inserts mint byte-identical tags and a document-wide
  // title+tag search could match the WRONG poem's control. Scoping the lookup
  // to the returned range removes that ambiguity class entirely; the title+tag
  // match within it stays as a guard against unrelated nested controls. Used
  // to restore the visible bounding-box outline the insertContentControl()
  // path used to set explicitly.
  async function styleInsertedPoemControl(context, insertedRange, tag) {
    var ccs = insertedRange.contentControls;
    ccs.load("items/title,items/tag");
    await context.sync();
    var items = ccs.items;
    for (var i = items.length - 1; i >= 0; i--) {
      if (items[i].title === "Ashaar Poem" && items[i].tag === tag) {
        items[i].appearance = "BoundingBox";
        return items[i];
      }
    }
    return null;
  }

  async function insertPoem(replaceSelection, optsOverride) {
    // Final review I3: propagate withWord's success flag so reRender (which
    // delegates its rebuild step to this function) can gate its own message.
    return await withWord(async function (context) {
      var opts = options();
      // Re-render passes overrides (e.g. justifyMode:"none" for a bare rebuild,
      // fontCsName to pin the poem's existing font). Merge over the pane opts.
      if (optsOverride) opts = Object.assign({}, opts, optsOverride);
      var source = String(input.value || "");

      var sectionP = context.document.sections.getFirst();
      sectionP.load("pageLayout/width,pageLayout/leftMargin,pageLayout/rightMargin");
      var selFontP = context.document.getSelection();
      selFontP.load("font/size,font/name");
      await context.sync();

      // pageLayout requires WordApi 1.5; fall back to US-Letter 6.5" on older builds
      var plP = sectionP.pageLayout;
      var pageTwips = plP && plP.width
        ? Math.round((plP.width - (plP.leftMargin || 0) - (plP.rightMargin || 0)) * 20)
        : 9360;

      // Measurement canvas at the selection's font — used for auto-fit, the nudge, and kashida.
      var fontSizeP = selFontP.font.size || 12;
      // Size-preserving re-render: when REPLACING an existing poem (justify /
      // adjust re-render), carry the selection's real font size into the rebuild
      // so the replaced table keeps its size instead of reverting to Word's
      // default. Only when the selection reports a single size (mixed → null →
      // leave the default). Fresh insert (replaceSelection false) is unchanged.
      if (replaceSelection && selFontP.font.size) opts.fontSizePt = selFontP.font.size;
      var modeP = opts.fontMode === "nastaliq" ? "noto" : opts.fontMode;
      var fontNameP = AshaarFonts.wordNameOf(modeP)
                    || selFontP.font.name || "Times New Roman";
      var ctxP = document.createElement("canvas").getContext("2d");
      if (ctxP) ctxP.font = fontSizeP + "pt \"" + fontNameP + "\"";

      var neededTwips = ctxP ? neededTableTwips(source, ctxP, opts, pageTwips) : pageTwips;
      var textWidthTwips;
      if (opts.autoFitWidth) {
        textWidthTwips = Math.min(pageTwips, neededTwips);
      } else {
        textWidthTwips = scaledTextWidth(pageTwips, opts.tableWidthPct);
        // Never render narrower than needed to avoid misra word-wrap — floor at
        // the needed width (still capped at the page width), even when the user
        // hasn't enabled Auto-fit. A wider user preference is still honored.
        if (ctxP && neededTwips > textWidthTwips) {
          textWidthTwips = Math.min(pageTwips, neededTwips);
        }
      }

      // "Let Word fill it" (§4): expand the render width by a strength-driven
      // fraction (0..15%) so higher strength visibly fills a wider table. The
      // base width above still comes from the Table-width slider / auto-fit;
      // this only scales it up, capped at the page width.
      if (opts.justifyMode === "css") {
        var expandFrac = AshaarWord.kashidaExpansionFraction(opts.tatweelCount);
        textWidthTwips = Math.min(pageTwips, Math.round(textWidthTwips * (1 + expandFrac)));
      }

      if ((opts.justifyMode === "kashida" || opts.justifyMode === "spacing") && ctxP) {
        opts._textWidthPx = textWidthTwips * 96 / 1440;
        opts._justifyCtx = ctxP;
      }

      var ooxmlBody;
      try {
        ooxmlBody = AshaarWord.renderForWordOoxml(source, opts, Ashaar, textWidthTwips);
      } catch (err) {
        setMessage("Render error: " + (err.message || String(err)));
        return;
      }
      if (!ooxmlBody) { setMessage("No content generated."); return; }

      var newTag = AshaarWord.contentControlTag(source, opts, AshaarWord.poemCellPatterns(source, opts, Ashaar));
      var selection = context.document.getSelection();

      // In-place re-render (Re-render / word-fill re-render). insertOoxml("Replace")
      // on a Range that spans an ENTIRE "Ashaar Poem" content control — its boundary
      // markers plus the poem tables — throws GeneralException. Replace the control's
      // CONTENT instead (ContentControl.insertOoxml keeps the wrapper), so nothing
      // crosses the boundary and we don't nest a second control inside the old one.
      // Fresh inserts and manual "Replace Selection" over plain text fall through
      // to the original selection-scope insert unchanged.
      if (replaceSelection) {
        // Robustly find the enclosing "Ashaar Poem" control. We can't use
        // parentContentControlOrNullObject: Re-render selects the control's WHOLE
        // range (boundary markers included), so the selection has no strict parent
        // control and that lookup returns null (which is why the boundary-crossing
        // Range.insertOoxml still fired). Intersect each control's range with the
        // selection instead — the SP2 detection pattern — which matches even when
        // the selection spans the whole control.
        var poemCCs = context.document.contentControls;
        poemCCs.load("items/title");
        await context.sync();
        var xs = poemCCs.items.map(function (c) {
          return { cc: c, hit: c.getRange().intersectWithOrNullObject(selection) };
        });
        xs.forEach(function (x) { x.hit.load("isNullObject"); });
        await context.sync();
        var poemCC = null;
        for (var xi = 0; xi < xs.length; xi++) {
          if (xs[xi].cc.title === "Ashaar Poem" && !xs[xi].hit.isNullObject) { poemCC = xs[xi].cc; break; }
        }
        if (poemCC) {
          // This build also rejects ContentControl.insertOoxml("Replace") with
          // tables (GeneralException) — replacing INTO a control with block tables
          // fails at every scope. Only body-scope table insert works (proven by the
          // End path). So insert the rebuilt poem into the BODY just after the old
          // control, then delete the old control and its content. The fresh
          // "Ashaar Poem" control is embedded in the OOXML itself (w:sdt via
          // wrapOoxmlControl — the same pattern the profile-apply rebuild uses):
          // wrapping the insertOoxml RETURN range with insertContentControl is
          // unreliable for multi-block content — Word clamps a control applied to
          // a range starting inside a table to the FIRST ROW, so only the first
          // row of the bandh ended up inside the control.
          var afterRange = poemCC.getRange("After");
          afterRange.insertOoxml(
            AshaarWord.wrapOoxmlControl(ooxmlBody, "Ashaar Poem", newTag),
            Word.InsertLocation.start);
          poemCC.delete(false); // remove old control + its content
          await context.sync();
          return;
        }
      }

      // Fresh insert / manual "Replace Selection" over plain text: embed the
      // SDT directly in the OOXML (wrapOoxmlControl) instead of wrapping the
      // insertOoxml-returned range with insertContentControl() — on Mac Word
      // a control applied post-hoc to a range that starts inside/touches a
      // table boundary clamps to the FIRST ROW only (the row-1-only bug fixed
      // above for Re-render/rebuild). newTag is embedded as the control's
      // w:tag, so title/tag no longer need setting afterward.
      var inserted = selection.insertOoxml(
        AshaarWord.wrapOoxmlControl(ooxmlBody, "Ashaar Poem", newTag),
        replaceSelection ? Word.InsertLocation.replace : Word.InsertLocation.end);
      await context.sync();
      await styleInsertedPoemControl(context, inserted, newTag);
      await context.sync();
    });
  }

  async function insertStructure() {
    await withWord(async function (context) {
      var opts = options();

      // Grid mode: build the 12-column span template from the bubble grid and
      // insert it (repeated per bandh) via the existing template-OOXML path.
      if (layoutView === "grid" && typeof AshaarLayoutGrid !== "undefined") {
        var sectionG = context.document.sections.getFirst();
        sectionG.load("pageLayout/width,pageLayout/leftMargin,pageLayout/rightMargin");
        await context.sync();
        var plG = sectionG.pageLayout;
        var twG = plG && plG.width
          ? Math.round((plG.width - (plG.leftMargin || 0) - (plG.rightMargin || 0)) * 20)
          : 9360;
        var tmplG = AshaarLayoutGrid.gridToTemplate(gridMatrix);
        if (!tmplG.rows.length) { setMessage("Draw at least one row of bubbles in the grid."); return; }
        var countG = Math.max(1, Math.min(20, Number(opts.bandhCount || 1)));
        var bodyG = [];
        var twGs = scaledTextWidth(twG, opts.tableWidthPct);
        for (var bi = 0; bi < countG; bi++) bodyG.push(AshaarWord.templateToOoxml(tmplG, twGs, opts));
        var selG = context.document.getSelection();
        var tagG = AshaarWord.contentControlTag("grid", opts);
        // Embed the SDT in the OOXML (wrapOoxmlControl) — insertContentControl()
        // on the insertOoxml-returned range clamps to the first row on Mac Word
        // when the range touches a table boundary (see insertPoem above).
        var insG = selG.insertOoxml(
          AshaarWord.wrapOoxmlControl(bodyG.join("<w:p/>"), "Ashaar Poem", tagG),
          Word.InsertLocation.end);
        await context.sync();
        await styleInsertedPoemControl(context, insG, tagG);
        await context.sync();
        return;
      }

      var tables = AshaarWord.layoutTablesForTemplate(opts);

      // Both span-based and plain (Numbers-view) layouts go through OOXML so
      // every table carries <w:bidiVisual/> and is a genuine RTL table. (The
      // native Word.insertTable API has no per-table RTL flag, so it always
      // yields an LTR table whose cell/tab order runs left-to-right even when
      // the content looks right — hence OOXML for every layout.)
      if (tables.length) {
        var section = context.document.sections.getFirst();
        section.load("pageLayout/width,pageLayout/leftMargin,pageLayout/rightMargin");
        await context.sync();
        var pl = section.pageLayout;
        var textWidthTwips = pl && pl.width
          ? Math.round((pl.width - (pl.leftMargin || 0) - (pl.rightMargin || 0)) * 20)
          : 9360;
        var scaled = scaledTextWidth(textWidthTwips, opts.tableWidthPct);
        var ooxmlBody = tables.map(function (t) {
          return t.spanBased
            ? AshaarWord.templateToOoxml(t, scaled, opts)
            : AshaarWord.layoutTableToOoxml(t, scaled, opts);
        }).join("<w:p/>");
        var selection = context.document.getSelection();
        var tagS = AshaarWord.contentControlTag("template", opts);
        // Embed the SDT in the OOXML (wrapOoxmlControl) — see insertPoem above
        // for why insertContentControl() on the insertOoxml return range is
        // unreliable (first-row clamp on Mac Word when tables are involved).
        var inserted = selection.insertOoxml(
          AshaarWord.wrapOoxmlControl(ooxmlBody, "Ashaar Poem", tagS),
          Word.InsertLocation.end);
        await context.sync();
        await styleInsertedPoemControl(context, inserted, tagS);
        await context.sync();
        return;
      }

      // HTML fallback (no span-table/plain-layout tables generated): this
      // still goes through insertHtml + a post-hoc insertContentControl(), so
      // it carries the SAME first-row-clamp risk as the OOXML paths above on
      // Mac Word when the inserted HTML renders as a table that touches a
      // boundary. Left as-is because wrapOoxmlControl only wraps OOXML bodies
      // (w:sdt), not HTML — converting this path to OOXML is out of scope here.
      var html = AshaarWord.renderTemplateForWord(opts);
      var selection = context.document.getSelection();
      var inserted = selection.insertHtml(html, Word.InsertLocation.end);
      var control = inserted.insertContentControl();
      control.title = "Ashaar Poem";
      control.tag = AshaarWord.contentControlTag("template", opts);
      control.appearance = "BoundingBox";
      await context.sync();
    });
  }

  async function insertTabStopPoem() {
    await withWord(async function (context) {
      // Read the actual text-area width so tab stops fit the document's page size and margins.
      // pageLayout properties are in points; multiply by 20 to convert to twips.
      var section = context.document.sections.getFirst();
      section.load("pageLayout/width,pageLayout/leftMargin,pageLayout/rightMargin");
      var selFont = context.document.getSelection();
      selFont.load("font/size,font/name");
      await context.sync();
      var pl = section.pageLayout;
      // pageLayout requires WordApi 1.5; fall back to US-Letter 6.5" text width on older builds
      var textWidthTwips = pl && pl.width
        ? Math.round((pl.width - (pl.leftMargin || 0) - (pl.rightMargin || 0)) * 20)
        : 9360;

      var opts = options();
      if (opts.justifyMode === "kashida" || opts.justifyMode === "spacing") {
        var fontSize = selFont.font.size || 12;
        var mode = opts.fontMode === "nastaliq" ? "noto" : opts.fontMode;
        var fontName = AshaarFonts.wordNameOf(mode)
                     || selFont.font.name || "Times New Roman";
        var canvas = document.createElement("canvas");
        var ctx = canvas.getContext("2d");
        if (ctx) {
          ctx.font = fontSize + "pt \"" + fontName + "\"";
          opts._justifyCtx = ctx;
        }
      }
      var source = String(input.value || "");
      var content;
      try {
        content = AshaarTabStop.poemToOoxml(source, opts, Ashaar, textWidthTwips);
      } catch (e) {
        setMessage("Paragraph engine error: " + (e.message || String(e)));
        return;
      }
      if (!content) {
        setMessage("No content generated.");
        return;
      }
      var ooxml = AshaarTabStop.wrapOoxml(content);
      var selection = context.document.getSelection();
      var inserted = selection.insertOoxml(ooxml, Word.InsertLocation.end);
      var control = inserted.insertContentControl();
      control.title = "Ashaar Poem";
      control.tag = AshaarWord.contentControlTag(source, opts);
      control.appearance = "BoundingBox";
      await context.sync();
    });
  }

  async function loadSelection() {
    var picked = null;
    await withWord(async function (context) {
      var selection = context.document.getSelection();
      selection.load("text");
      await context.sync();
      picked = selection.text || "";
    });
    if (picked) { setMode("convert"); input.value = picked; }
    // Run after withWord (which sets "Done.") so the conversion note is visible.
    applyImportNormalization();
  }

  // Kashida (U+0640) and the micro-spaces inserted by spacing justification
  // (hair U+200A, thin U+2009) are elongation artifacts, not part of the source
  // line. Strip them so every justification pass re-derives from the bare text —
  // making justification idempotent and reducible: re-justify any number of times
  // and the kashida count follows the current width / font / fill, never compounds.
  function stripJustification(s) {
    return String(s || "").replace(/[ـ  ]/g, "");
  }

  // "Let Word fill it" mode: no probe/calibrate/tatweel — Word's own kashida
  // renderer does the work via the paragraph's jc (justification) value.
  //
  // A per-cell insertOoxml replace (the original approach) is WRONG: verified
  // in Word that it drops the kashida jc — Word resets a cell paragraph's
  // alignment to the cell's own positional value on a body-scope replace.
  // insertOoxml at SELECTION scope, by contrast, preserves jc (that's exactly
  // what insertPoem/adoptTable already rely on). So instead of rebuilding
  // OOXML per cell, we reconstruct the poem's plain-text SOURCE from the
  // table(s)' cells (mirrors adoptTable's AshaarTableAdopt.adoptTableToSource
  // reconstruction) and re-run the normal insertPoem(true) pipeline — which
  // renders via renderForWordOoxml → wrapOoxml → selection.insertOoxml(...,
  // replace) and re-wraps the "Ashaar Poem" content control. Because the
  // §6 "Reset to unstretched" (core): strip justification artifacts — tatweels
  // (U+0640) and hairline/thin micro-spaces (U+200A / U+2009) — from the poem or
  // table at the cursor, IN PLACE. Never rebuilds the table, so font size and
  // run formatting survive. Reuses stripJustification (the same strip re-justify
  // already runs before each pass, so this leaves the exact bare text). The
  // fuller §6 reset (uniform font-scale, column widening, word-fill jc/break) is
  // a separate follow-up; this covers the "clear kashida & spaces" case.
  async function resetJustification() {
    if (typeof Word === "undefined") { setMessage("Open this task pane inside Word to reset."); return; }
    await withWord(async function (context) {
      var selection = context.document.getSelection();
      var cc = selection.parentContentControlOrNullObject;
      cc.load("title");
      await context.sync();
      var workRange = (!cc.isNullObject && cc.title === "Ashaar Poem") ? cc.getRange() : selection;
      var tables = workRange.tables;
      tables.load("items");
      await context.sync();
      if (!tables.items.length) { setMessage("Click inside an Ashaar table to reset."); return; }

      tables.items.forEach(function (tbl) { tbl.rows.load("items"); });
      await context.sync();
      tables.items.forEach(function (tbl) { tbl.rows.items.forEach(function (row) { row.cells.load("items"); }); });
      await context.sync();
      var cells = [];
      tables.items.forEach(function (tbl) {
        tbl.rows.items.forEach(function (row) {
          row.cells.items.forEach(function (cell) { cells.push(cell); cell.body.load("text"); });
        });
      });
      await context.sync();

      // One sync per changed cell so a single-cell failure falls out of the
      // batch without aborting the rest (mirrors justifySelection's write phase).
      var cleared = 0;
      for (var i = 0; i < cells.length; i++) {
        var cur = cells[i].body.text || "";
        var stripped = stripJustification(cur);
        if (stripped === cur) continue; // already clean — no tatweels/micro-spaces
        try {
          cells[i].body.paragraphs.getFirst().insertText(stripped, Word.InsertLocation.replace);
          await context.sync();
          cleared++;
        } catch (e) { /* skip a cell that fails to rewrite */ }
      }
      setMessage("Reset " + cleared + " cell(s) — kashida & micro-spaces cleared.");
    });
  }

  // Single-button Re-render for a managed poem: reconstruct it from the current
  // cells and rebuild through the normal render path applying the pane's gap /
  // table-width, but BARE (justifyMode "none") — renderForWordOoxml is
  // mechanism-unaware, so baking kashida here would inject generic tatweels into
  // a Jameel/Mehr poem. Font + size are preserved: size rides insertPoem's
  // replace-capture; font is the dropdown font when one is explicitly chosen
  // (so a mode that needs a specific font adopts it), else the poem's existing
  // font (pinned via opts.fontCsName). After the bare rebuild, fill in place
  // with the correct per-cell mechanism by delegating to justifySelection.
  async function reRender() {
    if (typeof Word === "undefined") { setMessage("Open this task pane inside Word to re-render."); return; }
    var opts = options();
    var source = "";
    var existingFont = "";
    var ccPayload = null;   // the block's OWN persisted tag payload (v3)
    var debug = !!(debugMode && debugMode.checked);
    var run = debug ? AshaarMetrics.startRun("re-render") : null;

    await withWord(async function (context) {
      var selection = context.document.getSelection();
      var cc = selection.parentContentControlOrNullObject;
      cc.load("title,tag");
      await context.sync();
      var isBlock = !cc.isNullObject && cc.title === "Ashaar Poem";
      if (isBlock) ccPayload = AshaarWord.parseContentControlTag(cc.tag);
      var workRange = isBlock ? cc.getRange() : selection;
      var tables = workRange.tables;
      tables.load("items");
      await context.sync();
      if (!tables.items.length) { setMessage("Click inside an Ashaar table to re-render."); return; }

      tables.items.forEach(function (tbl) { tbl.rows.load("items"); });
      await context.sync();
      tables.items.forEach(function (tbl) { tbl.rows.items.forEach(function (row) { row.cells.load("items"); }); });
      await context.sync();
      tables.items.forEach(function (tbl) {
        tbl.rows.items.forEach(function (row) {
          row.cells.items.forEach(function (cell) { cell.body.load("text"); cell.body.font.load("name"); });
        });
      });
      await context.sync();

      // Reconstruct source (same reconstruction Adopt / word-fill use).
      // Gap-corruption fix (Parts A+B): for a MANAGED block, spacing cells
      // must never re-enter the source — blank them via the persisted pattern
      // (or, unaligned, strip decor-symbol-only cells) exactly as
      // captureQaseedaTables does. Unmanaged tables keep raw cells.
      var rrDecorSyms = ccPayload
        ? blockDecorSymbols(loadProfileStore()[ccPayload.profile], ccPayload.slotDecor)
        : [];
      source = tables.items.map(function (tbl, ti) {
        var rows = tbl.rows.items.map(function (row) {
          return row.cells.items.map(function (cell) { return cell.body.text || ""; });
        });
        if (ccPayload) {
          var pat = ccPayload.cells ? ccPayload.cells[ti] : null;
          rows = AshaarCellMap.alignPatternToTable(rows.map(function (r) { return r.length; }), pat)
            ? AshaarTableAdopt.blankSpacingCells(rows, pat)
            : AshaarTableAdopt.stripDecorCells(rows, rrDecorSyms);
        }
        return AshaarTableAdopt.adoptTableToSource(rows, { direction: "rtl" });
      }).filter(function (s) { return s.trim(); }).join("\n\n");

      // Representative existing font (first cell reporting one) to preserve.
      // §7: skip a cell whose text is nothing but justification artifacts (a
      // blank spacing cell, or a leftover-tatweel cell) — its font is stale/
      // meaningless and must never pin the rebuilt poem's font.
      tables.items.forEach(function (tbl) {
        tbl.rows.items.forEach(function (row) {
          row.cells.items.forEach(function (cell) {
            if (!existingFont && cell.body.font && cell.body.font.name &&
                !AshaarFonts.isArtifactRun(cell.body.text || "")) {
              existingFont = cell.body.font.name;
            }
          });
        });
      });

      if (!source.trim()) { setMessage("That table didn't contain any text to re-render."); return; }
      workRange.select();
      await context.sync();
    });

    if (run) run.phase("capture");
    if (!source.trim()) return; // a friendly message was already shown

    // Font pin: an explicitly-chosen dropdown font wins (so a mode that needs a
    // specific font adopts it); otherwise preserve the poem's existing font.
    var dropdownFont = AshaarFonts.wordNameOf(opts.fontMode === "nastaliq" ? "noto" : opts.fontMode);
    var fontCsName = dropdownFont || existingFont || null;

    input.value = source;
    // Step 1: bare rebuild — gap/width from the pane, font pinned, size preserved.
    // The fresh contentControlTag write inside insertPoem must carry the block's
    // OWN persisted settings (profile / local / profileCache from the tag we just
    // read), not whatever options() supplies for fresh inserts — otherwise a
    // rebuild would discard the tag layers Apply (or an earlier session) wrote.
    var rebuildOverride = { justifyMode: "none", fontCsName: fontCsName };
    if (ccPayload) {
      rebuildOverride.profile = ccPayload.profile;
      rebuildOverride.local = ccPayload.local;
      rebuildOverride.profileCache = ccPayload.profileCache;
      // Final review I1: the fresh tag insertPoem mints via contentControlTag
      // must also carry these three or a Re-render (and structural poem-scope
      // Apply, which reuses this same path) silently destroys per-cell
      // overrides, gap decor, and bandh width — actively so for fill/color,
      // since the next justify pass's always-assert fill resets shading to
      // white once the override is gone. Deliberately extends spec §2's carry
      // list (which predates fill/color living in overrides) — recorded as a
      // spec deviation in the final-fixes report.
      rebuildOverride.overrides = ccPayload.overrides;
      rebuildOverride.slotDecor = ccPayload.slotDecor;
      rebuildOverride.widthPt = ccPayload.widthPt;
    }
    // Final review I3: track both steps' own success flags — the trailing
    // "Re-rendered" message below must not overwrite either step's failure
    // message (withWord's describeError(...), or applyProfileToQaseeda's
    // "Apply failed: …" when the hybrid qaseeda trigger fires inside
    // justifySelection). Message-last rule: success text only when the
    // pipeline actually succeeded.
    var rebuildOk = await insertPoem(true, rebuildOverride);
    if (run) run.phase("rebuild");
    // Step 2: fill in place with the correct per-cell mechanism for the chosen
    // mode (skipped when the pane mode is "none").
    var justifyOk = true;
    if (opts.justifyMode && opts.justifyMode !== "none") justifyOk = await justifySelection();
    if (run) run.phase("justify");
    if (run) {
      run.end();
      debugOutput.textContent += "\n" + JSON.stringify(run.report());
    }
    if (rebuildOk && justifyOk) setMessage("Re-rendered (font & size preserved).");
  }

  // justify-mode dropdown is "Word justify" (opts.justifyMode === "css") on
  // this path, that re-render emits the word-fill kashida jc + shrunk break
  // (misraParaXml), with width %/gap/strength/fontMode flowing through
  // exactly as they do for a fresh insert.
  async function justifySelectionWordFill(opts) {
    setMessage("Justifying…");
    var source = "";
    // Hoisted like plainGateCancelled below: withWord unconditionally
    // overwrites the callback's setMessage with "Done." on success, so the
    // honest message for this branch has to be re-asserted AFTER withWord
    // returns (see the re-assert comment further down).
    var wfPlainNative = false;

    await withWord(async function (context) {
      var selection = context.document.getSelection();

      // Find enclosing Ashaar Poem content control (mirrors the kashida/spacing path).
      var cc = selection.parentContentControlOrNullObject;
      cc.load("title,tag");
      await context.sync();

      var wfIsBlock = !cc.isNullObject && cc.title === "Ashaar Poem";
      var wfPayload = wfIsBlock ? AshaarWord.parseContentControlTag(cc.tag) : null;
      var workRange = wfIsBlock ? cc.getRange() : selection;

      var tables = workRange.tables;
      tables.load("items");
      workRange.load("text"); // for the plain-selection artifact strip below
      await context.sync();

      if (!tables.items.length) {
        // "Let Word fill it" (css mode) has no Ashaar table/width to fill
        // toward on a plain selection — "Select an Ashaar table to fill."
        // was misleading for prose that was never going to be in a table.
        // Word's native paragraph justify is exactly the mechanism css mode
        // already delegates to for tables (via jc), so apply it directly.
        // Review fix (R4): strip prior kashida artifacts (tatweels /
        // hair-spaces from an earlier kashida Apply) BEFORE justifying, same
        // as the spacing-mode plain path — otherwise the old artifacts stay
        // in the text and get space-stretched around. insertText returns the
        // replacing range; justify THAT, not the stale pre-replace proxy.
        var wfStripped = stripJustification(workRange.text);
        var wfJustRange = workRange;
        if (wfStripped !== workRange.text) {
          wfJustRange = workRange.insertText(wfStripped, Word.InsertLocation.replace);
        }
        wfJustRange.paragraphs.load("items");
        await context.sync();
        wfJustRange.paragraphs.items.forEach(function (p) { p.alignment = Word.Alignment.justified; });
        await context.sync();
        wfPlainNative = true;
        return;
      }

      tables.items.forEach(function (tbl) { tbl.rows.load("items"); });
      await context.sync();
      tables.items.forEach(function (tbl) {
        tbl.rows.items.forEach(function (row) { row.cells.load("items"); });
      });
      await context.sync();
      tables.items.forEach(function (tbl) {
        tbl.rows.items.forEach(function (row) {
          row.cells.items.forEach(function (cell) { cell.body.load("text"); });
        });
      });
      await context.sync();

      // Each table → one stanza; multiple tables in scope → stanza-separated.
      // Same reconstruction adoptTable uses, so the round-trip parsing rules
      // (misra/bayt/refrain detection from row layout) stay in one place.
      // Gap-corruption fix (Parts A+B): managed blocks blank their spacing
      // cells (pattern-aligned) or strip decor-symbol-only cells (unaligned)
      // before reconstruction — same treatment as captureQaseedaTables.
      var wfDecorSyms = wfPayload
        ? blockDecorSymbols(loadProfileStore()[wfPayload.profile], wfPayload.slotDecor)
        : [];
      source = tables.items.map(function (tbl, ti) {
        var rows = tbl.rows.items.map(function (row) {
          return row.cells.items.map(function (cell) { return cell.body.text || ""; });
        });
        if (wfPayload) {
          var pat = wfPayload.cells ? wfPayload.cells[ti] : null;
          rows = AshaarCellMap.alignPatternToTable(rows.map(function (r) { return r.length; }), pat)
            ? AshaarTableAdopt.blankSpacingCells(rows, pat)
            : AshaarTableAdopt.stripDecorCells(rows, wfDecorSyms);
        }
        return AshaarTableAdopt.adoptTableToSource(rows, { direction: "rtl" });
      }).filter(function (s) { return s.trim(); }).join("\n\n");

      if (!source.trim()) {
        setMessage("That table didn't contain any text to fill.");
        return;
      }

      // Put the selection on the content insertPoem(true) will replace.
      workRange.select();
      await context.sync();
    });

    // Re-assert AFTER withWord's unconditional "Done." — the native-justify
    // branch above already did the real work; withWord's Word.run resolving
    // without throwing means its "Done." would otherwise silently replace
    // this honest message.
    if (wfPlainNative) { setMessage("Justified paragraph (Word native)."); return true; }

    if (!source.trim()) return false; // a friendly message was already shown

    input.value = source;
    return await insertPoem(true);
  }

  // Wrapper: suppress active-context reflection while our own justify mutates
  // the document/selection, and guarantee the flag resets on every path.
  async function justifySelection() {
    _reflectBusy = true;
    try { return await justifySelectionInner(); }
    finally { _reflectBusy = false; }
  }

  async function justifySelectionInner() {
    var opts = options();
    var fontId = opts.fontMode === "nastaliq" ? "noto" : opts.fontMode;
    var mechanism = AshaarFonts.mechanismOf(fontId);

    // Hybrid qaseeda trigger: if the cursor's block belongs to a qaseeda that has
    // a stored profile, justify by applying that profile across ALL its blocks so
    // they stay consistent — instead of the free-form local justify below. Only
    // fires for tagged blocks; untagged blocks justify exactly as before.
    try {
      var qsel = await getQaseedaAtSelection();
      if (qsel.name && loadProfileStore()[qsel.name]) {
        // Final review I3: propagate the honest success flag to the caller
        // (justifySelection → reRender) instead of an unconditional return.
        return await applyProfileToQaseeda(qsel.name, { onlyBlockTag: qsel.tag });
      }
    } catch (e) { /* fall through to normal justify */ }

    // "Let Word fill it": native Word kashida (jc) instead of manual tatweel
    // insertion / spacing math. Entirely different code path — skip the
    // probe/calibrate/tatweel machinery below and delegate.
    if (opts.justifyMode === "css") { return await justifySelectionWordFill(opts); }

    // Fallback font from the pane — used only when a cell reports no explicit font.
    var fbMode = opts.fontMode === "nastaliq" ? "noto" : opts.fontMode;
    var fallbackName = AshaarFonts.wordNameOf(fbMode) || "Times New Roman";
    var doKashida = opts.justifyMode === "kashida" || opts.justifyMode === "spacing";
    var CELL_MARGIN_PT = 5.76; // Word default cell side margin (0.08") reserved for text
    var debug = !!(debugMode && debugMode.checked);
    var diags = [];
    var probeCacheStatus = "skip", calibCacheStatus = "skip"; // §8 debug visibility

    setMessage("Justifying…");

    // Kashida mechanism is resolved PER CELL/RUN from each run's REAL font
    // (see the generic run-aware path below), NOT from the pane dropdown.
    // Mehr (tatweel) and Jameel (font-swap) still get their explicit
    // dropdown-driven branches; every other pane selection ("Document default",
    // Noto, Gulzar, Arabic serif) is "whitespace" as an id but falls through to
    // the generic path, which decides kashida-vs-spacing from the actual font
    // of each run: arbitrary Arabic fonts (e.g. Fatemi Maqala) run the generic
    // tatweel engine; true whitespace-shaping fonts (Noto/Gulzar/Scheherazade)
    // fall back to spacing. There is deliberately no blanket dropdown-based
    // downgrade here — it forced "document" and every unrecognised font to
    // spacing, which is the regression this restores.

    // Task 9 fix round 1: withWord unconditionally setMessages "Done." after
    // its callback (see adoptTable's note), which would clobber the gate's
    // cancel message. Hoisted flag; re-set the message after (last write wins).
    var plainGateCancelled = false;
    // Same idiom, for the plain-selection Word-native-justify branch below
    // (spacing mode / kashida→spacing downgrade have no width to fill toward
    // on a plain selection, so they hand off to Word's own paragraph justify
    // instead of a silent no-op) — the honest message needs re-asserting
    // after withWord's unconditional "Done.".
    var plainWordNativeMsg = null;

    var ranOk = await withWord(async function (context) {
      var selection = context.document.getSelection();

      // Find enclosing Ashaar Poem content control (the poem is the calibration unit)
      var cc = selection.parentContentControlOrNullObject;
      cc.load("title,tag");
      await context.sync();

      var workRange = (!cc.isNullObject && cc.title === "Ashaar Poem")
        ? cc.getRange() : selection;

      // Persisted bandh cell-map (content/spacing tag + labels) for this block,
      // when present — one pattern per stanza table, in document order.
      var ccCells = null, ccOverrides = {}, ccColorClears = {};
      if (!cc.isNullObject && cc.title === "Ashaar Poem") {
        var ccPayload = AshaarWord.parseContentControlTag(cc.tag);
        ccCells = ccPayload && ccPayload.cells;
        ccOverrides = (ccPayload && ccPayload.overrides) || {};
        // §4 transition-clear consumption gate: pending color clears apply
        // ONLY to the block that recorded them — override keys ("0:A1")
        // repeat across poems, so a retained map (failed render kept for
        // retry) must not blacken a colliding cell of a different poem.
        if (_pendingColorClears.blockId &&
            AshaarWord.tagIdentity(cc.tag) === _pendingColorClears.blockId) {
          ccColorClears = _pendingColorClears.keys;
        }
      }

      var tables = workRange.tables;
      tables.load("items");
      await context.sync();

      if (!tables.items.length) {
        // No tables — justify plain selection text, measuring with the selection's own font.
        selection.load("text");
        selection.font.load("name,size");
        await context.sync();
        // Just-in-time font-measurement gate (Task 9): a plain selection has no
        // block/cell fonts to gather, so gate on this single face. "cancel"
        // aborts before insertText below — nothing in the document changes.
        var plainGate = await ensureFacesMeasurable(selection.font.name ? [selection.font.name] : []);
        if (plainGate === "cancel") { plainGateCancelled = true; return; }
        // Resolve the mechanism from the selection's REAL font: true
        // whitespace-shaping fonts (Noto/Gulzar/Scheherazade) shatter under
        // injected tatweels, so downgrade kashida→spacing for them; generic /
        // arbitrary Arabic fonts (Fatemi Maqala, …) keep kashida.
        var plainOpts = opts;
        var wordNativeReason = "";
        if (opts.justifyMode === "kashida" &&
            AshaarFonts.mechanismForFontName(selection.font.name) === "whitespace") {
          plainOpts = Object.assign({}, opts, { justifyMode: "spacing" });
          wordNativeReason = "“" + (selection.font.name || "This font") +
            "” can’t stretch letters in Word — ";
        }
        // "spacing" mode (whether requested directly or reached via the
        // kashida→whitespace-font downgrade above) has no width to fill
        // toward on a plain selection: AshaarWord.justifyPlainTextBlock /
        // justifyText (word-html.js) return the text UNCHANGED without a
        // colWidthPx, which looked like a successful no-op ("Done."). Word's
        // own paragraph justify stretches inter-word spaces to the margins —
        // exactly the spacing-mode behavior, done honestly — so use it
        // instead. kashida mode on a generic/tatweel-mechanism font (not
        // downgraded) is unaffected and keeps the width-blind tatweel
        // insertion below.
        if (plainOpts.justifyMode === "spacing") {
          // Review fix (R4 regression): the pre-native-justify code path
          // always ran the text through stripJustification before replacing —
          // which removed tatweels/hair-spaces left by an earlier kashida
          // Apply. Restore that strip first, or a previously-tatweeled
          // selection would keep its tatweels AND get space-stretched around
          // them. insertText returns the replacing range — justify THAT
          // (the original selection proxy no longer covers the new text).
          var strippedPlain = stripJustification(selection.text);
          var plainJustRange = selection;
          if (strippedPlain !== selection.text) {
            plainJustRange = selection.insertText(strippedPlain, Word.InsertLocation.replace);
          }
          plainJustRange.paragraphs.load("items");
          await context.sync();
          plainJustRange.paragraphs.items.forEach(function (p) { p.alignment = Word.Alignment.justified; });
          await context.sync();
          plainWordNativeMsg = wordNativeReason + "Justified paragraph (Word native).";
          return;
        }
        if (doKashida) {
          var pc = document.createElement("canvas").getContext("2d");
          if (pc) {
            pc.font = (selection.font.size || 16) + "pt \"" + (selection.font.name || fallbackName) + "\"";
            plainOpts._justifyCtx = pc;
          }
        }
        var justifiedText = AshaarWord.justifyPlainTextBlock(stripJustification(selection.text), plainOpts);
        selection.insertText(justifiedText, Word.InsertLocation.replace);
        await context.sync();
        return;
      }

      // Load rows → cells, including each cell's REAL font name/size (not a guess).
      tables.items.forEach(function (tbl) { tbl.rows.load("items"); });
      await context.sync();
      tables.items.forEach(function (tbl) {
        tbl.rows.items.forEach(function (row) { row.cells.load("items/columnWidth"); });
      });
      await context.sync();

      var allCells = [];
      tables.items.forEach(function (tbl, ti) {
        // Prefer the persisted bandh map for this table; fall back to geometry
        // (adopted/hand-drawn tables, older v1 tags → no map).
        var pattern = ccCells && ccCells[ti];
        var perRowCounts = tbl.rows.items.map(function (row) { return row.cells.items.length; });
        var tblMap = AshaarCellMap.alignPatternToTable(perRowCounts, pattern)
          ? AshaarCellMap.buildBandhCellMap(pattern) : null;
        var cellSeq = 0;
        tbl.rows.items.forEach(function (row, ri) {
          var cols = row.cells.items.length;
          row.cells.items.forEach(function (cell, ci) {
            allCells.push(cell);
            // Harmony key + content/spacing: the column-group key derived from
            // the persisted map (pools A1/B1/C1… across ROWS of one table when
            // their row shapes genuinely match — see columnGroupKey) when
            // available, else the geometric signature. `__kind` lets an empty
            // content cell stay content and a tagged gap be skipped regardless
            // of its text.
            var mapped = tblMap ? tblMap[cellSeq] : null;
            cellSeq++;
            if (mapped) {
              cell.__kind = mapped.kind;
              cell.__matKey = AshaarCellMap.columnGroupKey(pattern, mapped);
              cell.__ovKey = (mapped.kind === "content" && mapped.label)
                ? AshaarOverrides.overrideKey(ti, mapped.label) : null;
            } else {
              cell.__kind = null;
              cell.__matKey = AshaarMatrix.positionKey({ row: ri, col: ci, span: cols });
              cell.__ovKey = null;
            }
            cell.body.load("text");
            cell.body.font.load("name,size");
            // Alignment of the cell's own first paragraph — used by the
            // font-swap (Jameel) path to preserve the misra's visual side
            // (sadr/ajuz/solo) when it rebuilds the cell via OOXML.
            cell.body.paragraphs.load("alignment");
          });
        });
      });
      await context.sync();

      // Split each cell into word-ranges so justify can read a font per word and
      // rebuild the cell as an ordered list of runs (run-aware justification).
      allCells.forEach(function (cell) {
        cell.__wordRanges = cell.body.getRange().getTextRanges([" "], true);
        cell.__wordRanges.load("items");
      });
      await context.sync();
      allCells.forEach(function (cell) {
        cell.__wordRanges.items.forEach(function (wr) {
          wr.load("text");
          wr.font.load("name,size,bold,italic");
        });
      });
      await context.sync();

      // Representative font taken from the cells themselves (fall back to the
      // pane). §7: skip a cell whose text is nothing but justification
      // artifacts (blank spacing cell / stray tatweels) — never let its stale
      // font become the representative for the whole selection.
      var repName = fallbackName, repSize = 16;
      for (var ci = 0; ci < allCells.length; ci++) {
        var rf = allCells[ci].body.font;
        if (rf && rf.name && !AshaarFonts.isArtifactRun(allCells[ci].body.text || "")) {
          repName = rf.name; if (rf.size) repSize = rf.size; break;
        }
      }

      // Content width = cell width minus the side margins Word reserves for text.
      function contentPx(cell) {
        return Math.max(1, (cell.columnWidth || 0) - 2 * CELL_MARGIN_PT) * 96 / 72;
      }

      // The cell's own paragraph alignment ("Right"/"Left"/"Centered"/…) maps
      // to the "right"/"left"/"center" jc the font-swap path needs when it
      // rebuilds the cell as fresh OOXML — so the misra keeps its visual side.
      function cellAlignOf(cell) {
        var p0 = cell.body.paragraphs.items && cell.body.paragraphs.items[0];
        var al = p0 && p0.alignment;
        if (al === "Right") return "right";
        if (al === "Left") return "left";
        return "center";
      }

      // Build the measurement canvas with the REAL font + size.
      var canvasCtx = null;
      if (doKashida) {
        var c = document.createElement("canvas").getContext("2d");
        if (c) { c.font = repSize + "pt \"" + repName + "\""; canvasCtx = c; opts._justifyCtx = c; }
      }

      // Ensure a bundled @font-face (e.g. Fatemi Maqala) finishes loading before we
      // measure, so the canvas measures the same outlines Word renders. @font-face
      // fonts load lazily on first use; this forces the load and awaits it.
      if (canvasCtx && typeof document !== "undefined" && document.fonts && document.fonts.load) {
        try { await document.fonts.load(repSize + "pt \"" + repName + "\""); } catch (e) {}
      }

      // Jameel font-swap also measures fasls in the Kasheeda (wide) face on
      // this same canvas — force that @font-face to finish loading too, or
      // measureText silently falls back to a substitute font and corrupts
      // selectSwapRuns' gain ranking (wrong fasls get swapped).
      // Both font-swap (Jameel) and tatweel (Mehr) measure a specific Arabic
      // w:cs face on this canvas, not repName (the cell's reported Latin/hAnsi
      // font). If that face isn't force-loaded, measureText silently falls back
      // to a substitute and the elongation measures ~zero width — so Mehr's
      // trailing tatweel never registers as wider (no final tatweels selected)
      // and Jameel's gain ranking picks the wrong fasls.
      if (canvasCtx && (mechanism === "font-swap" || mechanism === "tatweel") &&
          typeof document !== "undefined" && document.fonts && document.fonts.load) {
        var bName = AshaarFonts.wordNameOf(fontId);
        if (bName) { try { await document.fonts.load(repSize + "pt \"" + bName + "\""); } catch (e) {} }
        if (mechanism === "font-swap") {
          var kName = AshaarFonts.kasheedaNameOf(fontId);
          if (kName) { try { await document.fonts.load(repSize + "pt \"" + kName + "\""); } catch (e) {} }
        }
      }

      // Auto-fit (in place): widen each table's columns so the widest misra has
      // kashida headroom, then justify into the new widths. Uses the desktop-only
      // TableColumn API (WordApiDesktop 1.3); on hosts without it, justify proceeds
      // at the current widths (no resize).
      var canResize = (typeof Office !== "undefined" && Office.context && Office.context.requirements
        && Office.context.requirements.isSetSupported
        && Office.context.requirements.isSetSupported("WordApiDesktop", "1.3"));
      if (opts.autoFitWidth && canvasCtx && canResize) {
        var sectionA = context.document.sections.getFirst();
        sectionA.load("pageLayout/width,pageLayout/leftMargin,pageLayout/rightMargin");
        await context.sync();
        var plA = sectionA.pageLayout;
        var pagePt = plA && plA.width ? (plA.width - (plA.leftMargin || 0) - (plA.rightMargin || 0)) : 468;
        var kOn = (opts.justifyMode === "kashida" || opts.justifyMode === "spacing") && Number(opts.tatweelCount || 0) > 0;
        var headroom = kOn ? 0.9 : 0.98;

        // Per table: the scale needed so the tightest cell gains headroom, capped at page width.
        var scaleByTable = tables.items.map(function (tbl) {
          var maxScale = 1, tableWpt = 0;
          tbl.rows.items.forEach(function (row, ri) {
            row.cells.items.forEach(function (cell) {
              if (ri === 0) tableWpt += (cell.columnWidth || 0);
              var t = stripJustification(cell.body.text || "").replace(/\s+/g, " ").trim();
              if (!t) return;
              var cf = cell.body.font;
              canvasCtx.font = ((cf && cf.size) || repSize) + "pt \"" + ((cf && cf.name) || repName) + "\"";
              var colWpx = (cell.columnWidth || 0) * 96 / 72;
              if (colWpx > 0) maxScale = Math.max(maxScale, canvasCtx.measureText(t).width / (headroom * colWpx));
            });
          });
          if (tableWpt > 0 && tableWpt * maxScale > pagePt) maxScale = pagePt / tableWpt;
          return maxScale;
        });

        if (scaleByTable.some(function (s) { return s > 1.01; })) {
          var colSets = tables.items.map(function (tbl, i) {
            if (scaleByTable[i] <= 1.01) return null;
            var cols = tbl.columns; cols.load("items/width"); return cols;
          });
          await context.sync();
          colSets.forEach(function (cols, i) {
            if (!cols) return;
            cols.items.forEach(function (col) { col.width = Math.round(col.width * scaleByTable[i] * 100) / 100; });
          });
          await context.sync();
          // Re-read cell widths (now changed) so justify targets the resized columns.
          tables.items.forEach(function (tbl) {
            tbl.rows.items.forEach(function (row) { row.cells.load("items/columnWidth"); });
          });
          await context.sync();
        }
      } else if (!opts.autoFitWidth && canvasCtx && canResize && opts.tableWidthPct) {
        // Table-width % applied IN PLACE (no rebuild): scale each table's columns
        // so its total width = tableWidthPct% of the page content width. Mirrors
        // how insert treats the slider when auto-fit is off, but on the EXISTING
        // table — so a width change no longer requires copy-and-replace. Uniform
        // scale preserves the gap:content proportions (layout shape intact).
        var sectionW = context.document.sections.getFirst();
        sectionW.load("pageLayout/width,pageLayout/leftMargin,pageLayout/rightMargin");
        await context.sync();
        var plW = sectionW.pageLayout;
        var pageW = plW && plW.width ? (plW.width - (plW.leftMargin || 0) - (plW.rightMargin || 0)) : 468;
        var targetW = Math.max(1, (Number(opts.tableWidthPct) / 100) * pageW);
        var colSetsW = tables.items.map(function (tbl) { var c = tbl.columns; c.load("items/width"); return c; });
        await context.sync();
        var didResize = false;
        colSetsW.forEach(function (cols) {
          var cur = 0; cols.items.forEach(function (col) { cur += (col.width || 0); });
          if (cur <= 0) return;
          var scale = targetW / cur;
          if (Math.abs(scale - 1) < 0.005) return;
          cols.items.forEach(function (col) { col.width = Math.round(col.width * scale * 100) / 100; });
          didResize = true;
        });
        if (didResize) {
          await context.sync();
          tables.items.forEach(function (tbl) {
            tbl.rows.items.forEach(function (row) { row.cells.load("items/columnWidth"); });
          });
          await context.sync();
        }
      }

      // Probe + calibrate using the real font/size and content widths.
      // §8: both memoized via _tuneCache — probe by (font, engine build),
      // persisted; calibrate by (font, size, container bucket, texts hash),
      // in-memory only. On a cache MISS this runs exactly the pre-existing
      // probeFont/calibrate calls and stores the result; behavior is
      // unchanged from before caching existed.
      var fontProfile = null;
      if (canvasCtx && typeof AshaarTune !== "undefined") {
        var pk = _tuneCache ? AshaarTuneCache.probeKey(repName, ASHAAR_UPSTREAM_VERSION) : null;
        fontProfile = pk ? _tuneCache.getProbe(pk) : null;
        if (fontProfile) {
          probeCacheStatus = "hit";
          // JSON storage strips the FontProfile's getQuality/getTierQuality
          // methods (buildSlots calls getQuality unconditionally when a
          // profile is present) — reattach them before use.
          fontProfile = AshaarTuneCache.rehydrateFontProfile(fontProfile);
        } else {
          probeCacheStatus = "miss";
          try { fontProfile = await AshaarTune.probeFont({ fontFamily: repName, fontSize: 64 }); }
          catch (e) { /* degrade gracefully */ }
          if (fontProfile && pk) _tuneCache.putProbe(pk, fontProfile);
        }
      }
      if (fontProfile) opts._fontProfile = fontProfile;

      var calibParams = { targetFill: 0.92 };
      if (fontProfile) calibParams.fontQualityBoost = 1.8;
      if (canvasCtx && typeof AshaarTune !== "undefined") {
        var lineTexts = [];
        var totalPx = 0, n = 0;
        allCells.forEach(function (cell) {
          var t = stripJustification(cell.body.text || "").replace(/[\r\n]+/g, " ").trim();
          if (t) lineTexts.push(t);
          if (cell.columnWidth > 0) { totalPx += contentPx(cell); n++; }
        });
        var avgPx = n ? totalPx / n : 300;
        if (lineTexts.length) {
          var ck = _tuneCache ? AshaarTuneCache.calibKey(repName, repSize, avgPx, lineTexts) : null;
          var cached = ck ? _tuneCache.getCalib(ck) : null;
          if (cached) {
            calibCacheStatus = "hit";
            calibParams = Object.assign({}, cached);
            if (fontProfile) calibParams.fontQualityBoost = calibParams.fontQualityBoost || 1.8;
          } else {
            calibCacheStatus = "miss";
            try {
              var session = await AshaarTune.calibrate({
                texts: lineTexts, fontFamily: repName, fontSize: repSize,
                containerWidth: avgPx, mode: "poetry", fontProfile: fontProfile, iterations: 50
              });
              calibParams = Object.assign({}, session.params);
              if (fontProfile) calibParams.fontQualityBoost = calibParams.fontQualityBoost || 1.8;
              if (ck) _tuneCache.putCalib(ck, calibParams);
            } catch (e) { /* keep defaults */ }
          }
        }
      }

      // Canvas font shorthand for one run: "[italic] [bold] Npt \"Family\"".
      function runFontStr(name, size, bold, italic) {
        return (italic ? "italic " : "") + (bold ? "bold " : "") +
          ((size || repSize)) + "pt \"" + (name || repName) + "\"";
      }
      // Micro-space glyph used to realize word-spacing in Word (text-mutating).
      var MICRO_SPACE = " "; // hair space
      if (canvasCtx) {
        canvasCtx.font = runFontStr(repName, repSize, false, false);
        if (canvasCtx.measureText(MICRO_SPACE).width <= 0) MICRO_SPACE = " "; // thin space
      }

      // Force-load EVERY distinct run font across all cells before measuring —
      // not just repName. A font the WebView CAN see (system-exposed, or a
      // bundled/uploaded @font-face) loads lazily on first use; without this,
      // measureText in a mixed-font cell silently falls back to a substitute
      // for the runs whose face isn't loaded yet, so those runs' metrics are
      // wrong. This is what lets dual-accessible fonts auto-measure correctly;
      // fonts the sandbox can't reach are supplied via the Custom-fonts
      // uploader, which registers an @font-face loaded the exact same way.
      if (canvasCtx && typeof document !== "undefined" && document.fonts && document.fonts.load) {
        var faceStrs = {};
        allCells.forEach(function (cell) {
          (cell.__wordRanges.items || []).forEach(function (wr) {
            var f = wr.font;
            var nm = (f && f.name) || repName, sz = (f && f.size) || repSize;
            faceStrs[runFontStr(nm, sz, !!(f && f.bold), !!(f && f.italic))] = true;
            // A font-swap font (Jameel) also measures fasls in its wider Kasheeda
            // face — load that too, or measureText falls back to a substitute and
            // mis-ranks the swaps. Needed under any dropdown (incl. Document
            // default), since per-cell dispatch may pick font-swap here.
            var kn = AshaarFonts.descriptorForFontName(nm).kasheedaName;
            if (kn) faceStrs[runFontStr(kn, sz, false, false)] = true;
          });
        });
        var faceLoads = [];
        Object.keys(faceStrs).forEach(function (s) { faceLoads.push(document.fonts.load(s).catch(function () {})); });
        try { await Promise.all(faceLoads); } catch (e) {}
      }

      // Natural-width matrix (harmony): the longest tatweel-free width per grid
      // position across every content cell in the work range. Natural-fit fills
      // each cell up to its position's Wpos (φ=1 pushes further, to the edge).
      var fillMode = opts.fillMode === "cell-fit" ? "cell-fit" : "natural-fit";
      var matrixCells = [];
      allCells.forEach(function (cell) {
        var base = stripJustification(cell.body.text || "").replace(/\s+/g, " ").trim();
        var isContent = cell.__kind === "content" || (cell.__kind == null && AshaarMatrix.isContentCell(base));
        if (!isContent) return; // tagged spacing (even with stray text) excluded from the matrix
        var mf = cell.body.font;
        var mnm = (mf && mf.name) || repName, msz = (mf && mf.size) || repSize;
        var natPx = 0;
        if (canvasCtx) { canvasCtx.font = runFontStr(mnm, msz, false, false); natPx = canvasCtx.measureText(base).width; }
        cell.__natPx = natPx;
        matrixCells.push({ key: cell.__matKey, natural: natPx });
      });
      var widthMatrix = AshaarMatrix.buildMatrix(matrixCells);

      // Phase 1 (pure, no sync): rebuild each cell as an ordered list of style
      // runs and justify measuring each run in its OWN font. Produces per-cell
      // write plans consumed in phase 2.
      var plans = [];
      allCells.forEach(function (cell) {
        var current = (cell.body.text || "").trim();
        if (cell.__kind === "spacing") return; // structural gap — never justified
        if (!stripJustification(current)) return;
        var colPx = contentPx(cell);

        // Per-cell override (SP2): strength / target width / cap-lift deviations
        // for this one cell, merged onto the block's justify defaults.
        var cellOv = cell.__ovKey ? ccOverrides[cell.__ovKey] : null;
        var resolved = AshaarOverrides.resolveCellOverride({ strength: opts.tatweelCount, fillMode: fillMode }, cellOv);
        var cellPhi = AshaarWord.strengthToElongationShare(resolved.strength);
        var cellMaxPos = AshaarWord.strengthToMaxPositions(resolved.strength);
        var cellCapEm = resolved.capEm != null ? resolved.capEm : undefined;

        // Resolve THIS cell's mechanism from its OWN real font (per-cell
        // dispatch), not the pane dropdown — so any dropdown (incl. Document
        // default) routes each cell to its font's correct mechanism instead of
        // shattering Jameel/Mehr with generic tatweels. The dropdown font is
        // only the fallback when a cell reports no resolvable font.
        // §7: read the font per WORD-RANGE and pick the first non-artifact
        // one (dominantRunFont) instead of cell.body.font.name — a mixed cell
        // (real word + a stale-font tatweel/space tail) reads null at the
        // whole-cell level and would silently fall through to repName, which
        // can itself be an unrelated cell's font; going straight to the word
        // ranges finds the real word's font directly.
        var cellWordFonts = (cell.__wordRanges.items || []).map(function (wr) {
          return { text: wr.text, font: wr.font && wr.font.name };
        });
        var cellFontName = AshaarFonts.dominantRunFont(cellWordFonts) || repName;
        var cellDesc = AshaarFonts.descriptorForFontName(cellFontName);
        var cellMech = cellDesc.mechanism;

        // Jameel font-swap: measure each fasl (connected segment) in the base
        // vs Kasheeda face, greedily swap the highest-gain fasls to the wider
        // face until the misra fills the column, and rebuild the cell as
        // OOXML with a per-run w:cs. Its own path — no word-range/tatweel/
        // spacing handling applies, so it returns before that machinery.
        if (cellMech === "font-swap") {
          if (!canvasCtx || colPx <= 0) return; // no measurement context — leave the cell as-is
          var cellAlign = cellAlignOf(cell);
          var wideCss = "\"" + (cellDesc.kasheedaName || repName) + "\"";
          var baseCss = "\"" + (cellDesc.wordName || repName) + "\"";
          var fss = AshaarKashidaFontswap.splitSpans(stripJustification(current));
          var wb = [], ww = [];
          fss.forEach(function (s) {
            canvasCtx.font = repSize + "pt " + baseCss; wb.push(canvasCtx.measureText(s).width);
            canvasCtx.font = repSize + "pt " + wideCss; ww.push(canvasCtx.measureText(s).width);
          });
          var jNatural = wb.reduce(function (a, b) { return a + b; }, 0);
          if (fillMode === "cell-fit") {
            // Cell-fit: swap fasls up to the φ elongation budget (no buffer),
            // then let Word distribute the residual to the true edge.
            var jBudget = AshaarMatrix.cellFitBudget(jNatural, colPx, cellPhi);
            var jSelC = AshaarKashidaFontswap.selectSwapRuns(fss, wb, ww, jBudget);
            var jRunsC = jSelC.runs.map(function (r) {
              return { text: r.text, csName: r.swap ? (cellDesc.kasheedaName || repName) : (cellDesc.wordName || repName), sizePt: repSize };
            });
            plans.push({ cell: cell, ooxml: AshaarWord.misraDistributeXml(jRunsC, repSize, opts), ov: cellOv });
            return;
          }
          // Natural-fit: fill to the position's matrix width (φ pushes toward the
          // buffered edge); capped hair-spaces backfill what the swaps miss.
          var jReach = colPx - 0.28 * repSize * 96 / 72;
          var jWpos = widthMatrix[cell.__matKey] || jNatural;
          var jTarget = (resolved.widthPt != null) ? resolved.widthPt * 96 / 72
            : AshaarMatrix.naturalFitTarget(jWpos, jReach, cellPhi);
          var sel = AshaarKashidaFontswap.selectSwapRuns(fss, wb, ww, jTarget);
          // Hybrid fill: font-swap elongation undershoots (only fasls with a
          // Kasheeda variant widen) — close the residual with capped hair-spaces
          // in the inter-word gap runs. Accept-short if the cap binds.
          var jGaps = 0;
          for (var jgi = 0; jgi < sel.runs.length; jgi++) { if (sel.runs[jgi].text === " ") jGaps++; }
          canvasCtx.font = repSize + "pt " + baseCss;
          var jSpacePx = canvasCtx.measureText(MICRO_SPACE).width || 1;
          var jn = AshaarResidual.capMicroSpaces(jTarget - sel.fill * jTarget, jGaps, jSpacePx, repSize * 96 / 72, cellCapEm);
          var jRuns = AshaarResidual.injectSpaceRuns(sel.runs, jn, MICRO_SPACE);
          var swapXml = AshaarWord.runsToMisraXml(jRuns, cellAlign, opts, repSize);
          plans.push({ cell: cell, ooxml: swapXml, ov: cellOv });
          return; // handled — skip the tatweel/spacing paths for this cell
        }

        // Mehr tatweel: DISCRETE trailing elongation. Mehr renders a clean
        // kashida only from ONE trailing tatweel after a word ending in a
        // whitelisted final letter (medial U+0640 is zero-width on the canvas
        // we measure with; Word-native highKashida does nothing for Mehr). So
        // Mehr fits by the SAME discrete subset-selection as Jameel — choose
        // which eligible words get a trailing tatweel. Single-font text output.
        if (cellMech === "tatweel" && opts.justifyMode === "kashida") {
          if (!canvasCtx || colPx <= 0) return;
          var mehrFont = repSize + "pt \"" + (cellDesc.wordName || repName) + "\"";
          var mRules = cellDesc.tatweelRules || {};
          var isoSet = {}, finSet = {};
          (mRules.isolatedInto || []).forEach(function (c) { isoSet[c] = true; });
          (mRules.finalInto || []).forEach(function (c) { finSet[c] = true; });
          var mline = stripJustification(current);
          var mparts = mline.split(" "), mtoks = [];
          mparts.forEach(function (wd, i) { if (i) mtoks.push(" "); mtoks.push(wd); });
          // Form-aware: trailing tatweel only on allowed isolated/final letters.
          var melong = mtoks.map(function (t) { return t !== " " ? AshaarWord.mehrElongate(t, isoSet, finSet) : t; });
          var mwb = [], mww = [];
          canvasCtx.font = mehrFont;
          for (var mi = 0; mi < mtoks.length; mi++) { mwb.push(canvasCtx.measureText(mtoks[mi]).width); mww.push(canvasCtx.measureText(melong[mi]).width); }
          var mNatural = mwb.reduce(function (a, b) { return a + b; }, 0);
          if (fillMode === "cell-fit") {
            var mBudget = AshaarMatrix.cellFitBudget(mNatural, colPx, cellPhi);
            var mselC = AshaarKashidaFontswap.selectSwapRuns(mtoks, mwb, mww, mBudget);
            var moutC = mselC.runs.map(function (r, i) { return (r.swap && mww[i] > mwb[i]) ? melong[i] : mtoks[i]; }).join("");
            plans.push({ cell: cell, ooxml: AshaarWord.misraDistributeXml([{ text: moutC, csName: cellDesc.wordName || repName, sizePt: repSize }], repSize, opts), ov: cellOv });
            return;
          }
          var mReach = colPx - 0.28 * repSize * 96 / 72;
          var mWpos = widthMatrix[cell.__matKey] || mNatural;
          var mTarget = (resolved.widthPt != null) ? resolved.widthPt * 96 / 72
            : AshaarMatrix.naturalFitTarget(mWpos, mReach, cellPhi);
          var msel = AshaarKashidaFontswap.selectSwapRuns(mtoks, mwb, mww, mTarget);
          var mout = msel.runs.map(function (r, i) { return (r.swap && mww[i] > mwb[i]) ? melong[i] : mtoks[i]; }).join("");
          // Hybrid fill: Mehr elongates only at whitelisted word-endings, so it
          // undershoots — close the residual with capped hair-spaces at the word
          // gaps (reusing distributeMicroSpaces). Accept-short if the cap binds.
          var mGaps = mout.split(" ").length - 1;
          canvasCtx.font = mehrFont;
          var mSpacePx = canvasCtx.measureText(MICRO_SPACE).width || 1;
          var mn = AshaarResidual.capMicroSpaces(mTarget - msel.fill * mTarget, mGaps, mSpacePx, repSize * 96 / 72, cellCapEm);
          var mfinal = AshaarWord.distributeMicroSpaces([mout], mn, MICRO_SPACE)[0];
          if (mfinal !== current) plans.push({ cell: cell, flat: mfinal, align: cellAlignOf(cell), ov: cellOv });
          // §4: text unchanged, but EVERY managed cell still gets its decor
          // (re)asserted or cleared — fill is always-assert ("#FFFFFF" clears
          // a deleted override), and a null cellOv is exactly the deleted case.
          else if (cell.__ovKey) plans.push({ cell: cell, ov: cellOv });
          return;
        }

        // Per-word style tuples from the word ranges, then coalesce to runs.
        var words = [];
        (cell.__wordRanges.items || []).forEach(function (wr) {
          var t = stripJustification(wr.text || "");
          if (!t) return;
          var f = wr.font;
          words.push({
            text: t,
            name: (f && f.name) || repName,
            size: (f && f.size) || repSize,
            bold: !!(f && f.bold),
            italic: !!(f && f.italic),
            range: wr
          });
        });
        if (!words.length) return;
        var runs = AshaarWord.coalesceRuns(words);

        // Fallback: without a measurement canvas we cannot do run-aware work —
        // justify the flattened line as before (single-font behavior).
        if (!canvasCtx || colPx <= 0) {
          var flat = AshaarWord.justifyPlainTextBlock(stripJustification(current), opts, colPx);
          if (flat !== current) plans.push({ cell: cell, flat: flat, ov: cellOv });
          else if (cell.__ovKey) plans.push({ cell: cell, ov: cellOv }); // §4 decor always-asserts
          return;
        }

        // Primitive runs: each carries a measure() bound to its own font.
        var primRuns = runs.map(function (r) {
          var fstr = runFontStr(r.name, r.size, r.bold, r.italic);
          return {
            text: r.text,
            fontSize: r.size,
            fontProfile: fontProfile || null,
            measure: function (s) { canvasCtx.font = fstr; return canvasCtx.measureText(s).width; }
          };
        });

        var outTexts; // per-run text to write back (null when spacing writes properties only)
        var sp = null;

        // Per-cell mechanism from the runs' REAL fonts. When kashida is chosen
        // and every run is a font the tatweel engine can elongate (generic /
        // arbitrary Arabic fonts like Fatemi Maqala), kashida-fill via
        // justifyRuns — each run measured in its OWN font, so mixed-font misras
        // stretch correctly (Task A3 Step 3). If ANY run is a whitespace-shaping
        // font (Noto/Gulzar/Scheherazade), where injected tatweels shatter the
        // shaping, or the user chose spacing, fall back to run-aware spacing.
        // (Mehr/Jameel cells never reach here — handled by the branches above.)
        var anyWhitespaceRun = runs.some(function (r) {
          return AshaarFonts.mechanismForFontName(r.name) === "whitespace";
        });
        if (opts.justifyMode === "kashida" && !anyWhitespaceRun) {
          var gNatural = primRuns.reduce(function (a, r) { return a + r.measure(r.text); }, 0);
          var gMax = { perPositionEm: 0.5, maxPositions: cellMaxPos };
          if (fillMode === "cell-fit") {
            // Cell-fit: concentrate tatweels to the φ budget (no buffer); Word's
            // distribute jc stretches the inter-word gaps to the true edge.
            var gBudgetC = AshaarMatrix.cellFitBudget(gNatural, colPx, cellPhi);
            var concC = AshaarJustify.justifyRunsConcentrated(primRuns, gBudgetC, Object.assign({}, calibParams, gMax));
            var cfRuns = concC.runs.map(function (r, i) { return { text: r.text, csName: runs[i].name, sizePt: runs[i].size }; });
            plans.push({ cell: cell, ooxml: AshaarWord.misraDistributeXml(cfRuns, repSize, opts), ov: cellOv });
            return;
          }
          // Natural-fit: fill to the position's matrix width; capped micro-spaces
          // backfill whatever the concentrated tatweels didn't cover — so low
          // strength = spacing-dominant, harmony baseline at φ=0.
          var gReach = colPx - 0.28 * repSize * 96 / 72;
          var gWpos = widthMatrix[cell.__matKey] || gNatural;
          var gTarget = (resolved.widthPt != null) ? resolved.widthPt * 96 / 72
            : AshaarMatrix.naturalFitTarget(gWpos, gReach, cellPhi);
          var conc = AshaarJustify.justifyRunsConcentrated(primRuns, gTarget, Object.assign({}, calibParams, gMax));
          outTexts = conc.runs.map(function (r) { return r.text; });
          var gGaps = primRuns.reduce(function (a, r) { return a + (r.text.split(" ").length - 1); }, 0);
          canvasCtx.font = runFontStr(repName, repSize, false, false);
          var gSpacePx = canvasCtx.measureText(MICRO_SPACE).width || 1;
          var gN = AshaarResidual.capMicroSpaces(gTarget - conc.achievedPx, gGaps, gSpacePx, repSize * 96 / 72, cellCapEm);
          outTexts = AshaarWord.distributeMicroSpaces(outTexts, gN, MICRO_SPACE);
        } else {
          // spacing/scale: single wordSpacing + uniform fontScale from run-aware widths.
          sp = AshaarJustify.computeRunSpacing(primRuns, colPx, calibParams);
          var gaps = runs.reduce(function (a, r) { return a + (r.text.split(" ").length - 1); }, 0);
          canvasCtx.font = runFontStr(repName, repSize, false, false);
          var spaceGlyphPx = canvasCtx.measureText(MICRO_SPACE).width || 1;
          var n = Math.max(0, Math.round(sp.wordSpacing * gaps / spaceGlyphPx));
          outTexts = AshaarWord.distributeMicroSpaces(runs.map(function (r) { return r.text; }), n, MICRO_SPACE);
        }

        if (debug) {
          var natSum = 0, finSum = 0, twCount = 0;
          primRuns.forEach(function (pr, i) { natSum += pr.measure(runs[i].text); finSum += pr.measure(outTexts[i]); });
          outTexts.forEach(function (t) { twCount += (t.match(/ـ/g) || []).length; });
          diags.push({
            i: diags.length,
            font: runs.length + " run(s), " + repSize + "pt " + repName,
            res: fontAvailable(runs[0].name) ? "yes" : "NO",
            colPx: Math.round(colPx),
            colIn: (colPx / 96).toFixed(2),
            nat: Math.round(natSum),
            // gTarget is the ACTUAL value passed to justifyRunsConcentrated on
            // the natural-fit kashida path (harmony-pooled via widthMatrix).
            // The old colPx*targetFill readout ignored the matrix entirely and
            // actively misled debugging of the row-pooling bug (M4). gTarget
            // is only assigned on that branch; the spacing/scale branch never
            // sets it, so keep the previous readout there unchanged.
            target: Math.round(gTarget != null ? gTarget : colPx * (calibParams.targetFill || 1)),
            fin: Math.round(finSum),
            fill: colPx ? Math.round(finSum / colPx * 100) : 0,
            tw: twCount + (sp ? " ws" + sp.wordSpacing + " x" + sp.fontScale : ""),
            cap: runs.reduce(function (a, r) { return a + r.text.replace(/\s/g, "").length; }, 0),
            text: runs.map(function (r) { return r.text; }).join(" ").slice(0, 14)
          });
        }

        // Each run's justified text must split 1:1 back onto its source word
        // ranges (tatweels/micro-spaces never add ASCII spaces). If that ever
        // fails, route the cell to the flattened path instead of a partial write.
        var alignedOk = runs.every(function (r, i) {
          return outTexts[i].split(" ").length === r.refs.length;
        });
        if (!alignedOk) { plans.push({ cell: cell, flat: outTexts.join(" "), align: cellAlignOf(cell), ov: cellOv }); return; }

        plans.push({ cell: cell, runs: runs, outTexts: outTexts, sp: sp, align: cellAlignOf(cell), ov: cellOv });
      });

      // Phase 2 (write): one context.sync() per cell so a range failure on one
      // cell falls back to a flattened whole-cell replace without aborting the
      // batch (the run-aware write can only error at sync, not synchronously).
      // Map a cell's own alignment → an Office enum. Applied on flat/run-aware
      // writes so re-justifying a cell that was previously Cell-fit (paragraph
      // jc=distribute) clears the distribute — Office.js has no "distribute"
      // Alignment, so we re-assert the cell's intended side.
      function officeAlign(a) {
        if (a === "right") return Word.Alignment.right;
        if (a === "left") return Word.Alignment.left;
        return Word.Alignment.centered;
      }

      // §4 cell fill/color: read straight off the tag override (bypasses
      // resolveCellOverride, which only carries strength/widthPt/capEm) — a
      // no-op for cells outside a persisted bandh map (__ovKey null: plain
      // selections, adopted tables with no cell pattern) so this never touches
      // shading on tables our override system doesn't own. shadingColor
      // rejects "" / "No color"; "#FFFFFF" clears it (same quirk documented at
      // applyProfileToQaseeda's spacing-cell decor branch) — so fill is
      // ALWAYS-assert: a deleted fill override clears on the next pass.
      // Color has no such clear value, so it clears by TRANSITION: the Apply
      // that removed the override recorded the key in _pendingColorClears
      // (block-scoped — ccColorClears is non-empty only when this block's
      // identity matches) and we reset those cells to "black". Accepted
      // limitations (same class as the width-drift trade-off): (1) the reset
      // restores "black", not any pre-override manual text color — the tag
      // never captured one, and the qaseeda capture reads live run colors as
      // "original", so there is no source data to recover it from; (2) if the
      // one cell carrying a pending clear hits a per-cell write failure, the
      // success tail consumes the clear anyway and the stale color survives
      // (re-baked by the next capture). Narrow, accepted.
      function applyCellDecor(cell, ov) {
        if (!cell.__ovKey) return;
        ov = ov || {};
        cell.shadingColor = ov.fill || "#FFFFFF";
        if (ov.color) cell.body.font.color = ov.color;
        else if (ccColorClears[cell.__ovKey]) cell.body.font.color = "black";
      }

      var changed = 0;
      for (var pi = 0; pi < plans.length; pi++) {
        var p = plans[pi];
        if (p.ooxml) {
          try {
            p.cell.body.clear();
            p.cell.body.insertOoxml(AshaarWord.wrapOoxml(p.ooxml), Word.InsertLocation.replace);
            applyCellDecor(p.cell, p.ov);
            await context.sync();
            changed++;
          } catch (e) {
            if (debug) diags.push({ i: diags.length, font: "OOXML-FAIL", text: (e && e.message || "").slice(0, 14) });
          }
          continue;
        }
        if (p.flat != null) {
          var flatPara = p.cell.body.paragraphs.getFirst();
          flatPara.insertText(p.flat, Word.InsertLocation.replace);
          if (p.align) flatPara.alignment = officeAlign(p.align);
          applyCellDecor(p.cell, p.ov);
          await context.sync();
          changed++;
          continue;
        }
        if (!p.runs) {
          // §4 decor-only: justified text is unchanged this pass, but the
          // cell's fill/color still needs (re)asserting or clearing. No sync
          // of its own — property writes queue and commit at Word.run's final
          // implicit sync (no added round-trips for decor-only cells).
          applyCellDecor(p.cell, p.ov);
          changed++;
          continue;
        }
        try {
          var cellChanged = false;
          p.runs.forEach(function (r, i) {
            // outTexts[i] splits 1:1 onto the run's original word ranges
            // (validated in phase 1) — write each word range independently
            // (disjoint; no union/expand needed).
            var pieces = p.outTexts[i].split(" ");
            r.refs.forEach(function (w, j) {
              if (p.sp && p.sp.fontScale !== 1) { w.range.font.size = r.size * p.sp.fontScale; cellChanged = true; }
              if (pieces[j] !== w.text) { w.range.insertText(pieces[j], Word.InsertLocation.replace); cellChanged = true; }
            });
          });
          if (p.align) { p.cell.body.paragraphs.getFirst().alignment = officeAlign(p.align); cellChanged = true; }
          // Decor always-asserts for managed cells (fill clears via "#FFFFFF"
          // when the override is gone). The writes queue into this plan's
          // existing sync below (or Word.run's final implicit sync when the
          // text turned out unchanged) — no added round-trips.
          if (p.cell.__ovKey) applyCellDecor(p.cell, p.ov);
          if (cellChanged) { await context.sync(); changed++; }
        } catch (e) {
          // Queued range write failed at sync (or count mismatch) — flatten.
          p.cell.body.paragraphs.getFirst().insertText(p.outTexts.join(" "), Word.InsertLocation.replace);
          applyCellDecor(p.cell, p.ov);
          await context.sync();
          changed++;
          if (debug) diags.push({ i: diags.length, font: "RANGE-FALLBACK", text: (e && e.message || "").slice(0, 14) });
        }
      }

      setMessage("Justified " + changed + " cell(s) across " + tables.items.length + " table(s).");
      if (debug) renderDebug(diags, { probe: probeCacheStatus, calib: calibCacheStatus });
    });
    // Re-assert the cancel message AFTER withWord's unconditional "Done.".
    if (plainGateCancelled) setMessage("Add the font, then Apply again.");
    // Same re-assert for the plain-selection Word-native-justify branch —
    // its work genuinely succeeded (ranOk stays true), only the message
    // needs restoring over "Done.".
    if (plainWordNativeMsg) setMessage(plainWordNativeMsg);
    // Final review I3: a gate cancel is a soft early-return inside the
    // callback — withWord itself still reports true (Word.run didn't throw)
    // — so override the signal to false here for an honest result.
    return plainGateCancelled ? false : ranOk;
  }

  // Adopt an existing Word table of poetry: read its cells, reconstruct the
  // canonical Ashaar source, and (by default) replace the table in place with a
  // managed, content-controlled Ashaar block. Uses Word.run directly so we keep
  // control of messaging (withWord forces a "Done." message).
  async function adoptTable() {
    if (typeof Word === "undefined") {
      setMessage("Open this task pane inside Word to adopt a table.");
      return;
    }
    var reviewOnly = document.getElementById("adopt-review").checked;
    var dirChoice = document.getElementById("adopt-direction").value;     // auto | rtl | ltr
    var direction = dirChoice === "ltr" ? "ltr" : "rtl";                   // auto → rtl
    var scope = document.getElementById("adopt-scope").value;             // cursor | selection
    var source = "";

    try {
      await Word.run(async function (context) {
        var selection = context.document.getSelection();
        var targetTables;

        if (scope === "selection") {
          var tbls = selection.tables;
          tbls.load("items");
          await context.sync();
          targetTables = tbls.items;
          if (!targetTables.length) { setMessage("Select one or more tables to adopt."); return; }
        } else {
          var t = selection.parentTableOrNullObject;
          t.load("rows");
          await context.sync();
          if (t.isNullObject) { setMessage("Place the cursor inside a table to adopt it."); return; }
          targetTables = [t];
        }

        targetTables.forEach(function (tbl) { tbl.rows.load("items"); });
        await context.sync();
        targetTables.forEach(function (tbl) {
          tbl.rows.items.forEach(function (row) { row.cells.load("items"); });
        });
        await context.sync();
        targetTables.forEach(function (tbl) {
          tbl.rows.items.forEach(function (row) {
            row.cells.items.forEach(function (cell) { cell.body.load("text"); });
          });
        });
        await context.sync();

        // Each table → one stanza; multiple selected tables → stanza-separated.
        source = targetTables.map(function (tbl) {
          var rows = tbl.rows.items.map(function (row) {
            return row.cells.items.map(function (cell) { return cell.body.text || ""; });
          });
          return AshaarTableAdopt.adoptTableToSource(rows, { direction: direction });
        }).filter(function (s) { return s.trim(); }).join("\n\n");

        if (!source.trim()) { setMessage("That table didn't contain any text to adopt."); return; }

        // Put the selection on the content we'll replace, so insertPoem(true) targets it.
        var range = (scope === "cursor") ? targetTables[0].getRange() : selection.getRange();
        range.select();
        await context.sync();
      });
    } catch (e) {
      setMessage("Adopt failed: " + (e && e.message ? e.message : String(e)));
      return;
    }

    if (!source.trim()) return; // a friendly message was already shown

    // Show the recovered source + preview (transparent, editable).
    setMode("convert");
    input.value = source;
    renderPreview();

    if (reviewOnly) {
      setMessage("Adopted the table into the editor. Review the text, then click Replace Selection.");
      return;
    }

    // One-click: replace the selected table with the regenerated Ashaar block.
    await insertPoem(true);
    setMessage("Table adopted and replaced with a formatted Ashaar block.");
  }

  // ── Template persistence helpers ───────────────────────────────────────────

  function loadTemplates() {
    try { return JSON.parse(localStorage.getItem("ashaar-templates") || "[]"); }
    catch (e) { return []; }
  }

  function saveTemplates(templates) {
    localStorage.setItem("ashaar-templates", JSON.stringify(templates));
  }

  function renderTemplateList() {
    var templates = loadTemplates();
    templateList.innerHTML = "";
    if (!templates.length) {
      var opt = document.createElement("option");
      opt.value = "";
      opt.textContent = "— no saved templates —";
      opt.disabled = true;
      opt.selected = true;
      templateList.appendChild(opt);
      return;
    }
    templates.forEach(function (t) {
      var opt = document.createElement("option");
      opt.value = t.id;
      opt.textContent = t.name;
      templateList.appendChild(opt);
    });
  }

  // ── Drop bare 12-column grid ───────────────────────────────────────────────

  async function insertBareGrid() {
    await withWord(async function (context) {
      var opts = options();
      var section = context.document.sections.getFirst();
      section.load("pageLayout/width,pageLayout/leftMargin,pageLayout/rightMargin");
      await context.sync();
      var pl = section.pageLayout;
      var textWidthTwips = pl && pl.width
        ? Math.round((pl.width - (pl.leftMargin || 0) - (pl.rightMargin || 0)) * 20)
        : 9360;
      var bodyBG = AshaarWord.generateBareGrid12Ooxml(scaledTextWidth(textWidthTwips, opts.tableWidthPct));
      var tagBG = AshaarWord.contentControlTag("grid12", opts);
      var selection = context.document.getSelection();
      // Embed the SDT in the OOXML (wrapOoxmlControl) — see insertPoem for why
      // a post-hoc insertContentControl() clamps to the first row on Mac Word.
      var inserted = selection.insertOoxml(
        AshaarWord.wrapOoxmlControl(bodyBG, "Ashaar Poem", tagBG),
        Word.InsertLocation.end);
      await context.sync();
      await styleInsertedPoemControl(context, inserted, tagBG);
      await context.sync();
      setMessage("12-column grid inserted. Merge cells in Word, then Capture as a template.");
    });
  }

  // ── Capture selected table layout ─────────────────────────────────────────

  async function captureSelectedTableLayout() {
    var name = (templateNameInput.value || "").trim();
    if (!name) { setMessage("Enter a template name first."); return; }

    await withWord(async function (context) {
      var selection = context.document.getSelection();
      var table = selection.parentTableOrNullObject;
      table.load("rows");
      await context.sync();
      if (table.isNullObject) { setMessage("Click inside a table first, then capture."); return; }

      table.rows.load("items");
      await context.sync();
      table.rows.items.forEach(function (row) { row.cells.load("items/columnWidth"); });
      await context.sync();

      // Infer total table width (sum of first row cell widths in points)
      var firstRow = table.rows.items[0];
      var totalWidthPt = 0;
      firstRow.cells.items.forEach(function (cell) { totalWidthPt += (cell.columnWidth || 0); });
      if (totalWidthPt <= 0) { setMessage("Could not read table cell widths."); return; }

      var GRID = 12;
      var baseColPt = totalWidthPt / GRID;
      var rows = table.rows.items.map(function (row) {
        return row.cells.items.map(function (cell) {
          var span = Math.max(1, Math.min(GRID, Math.round((cell.columnWidth || baseColPt) / baseColPt)));
          return { span: span };
        });
      });

      var id = String(Date.now());
      var v = panelValues();
      var template = {
        id: id,
        name: name,
        columnCount: GRID,
        rows: rows,
        justifyMode: v.justifyMode,
        tatweelCount: Number(v.strength || 6),
        gapWidth: Number(v.gap != null ? v.gap : 4)
      };

      var templates = loadTemplates();
      templates.push(template);
      saveTemplates(templates);
      renderTemplateList();
      // Select the newly saved template
      for (var i = 0; i < templateList.options.length; i++) {
        if (templateList.options[i].value === id) { templateList.selectedIndex = i; break; }
      }
      templateNameInput.value = "";
      setMessage("Template \"" + name + "\" saved.");
    });
  }

  // ── Apply saved template ───────────────────────────────────────────────────

  async function applyTemplate() {
    var id = templateList.value;
    if (!id) { setMessage("Select a template first."); return; }
    var templates = loadTemplates();
    var tmpl = null;
    for (var i = 0; i < templates.length; i++) {
      if (templates[i].id === id) { tmpl = templates[i]; break; }
    }
    if (!tmpl) { setMessage("Template not found."); return; }

    await withWord(async function (context) {
      var section = context.document.sections.getFirst();
      section.load("pageLayout/width,pageLayout/leftMargin,pageLayout/rightMargin");
      await context.sync();
      var pl = section.pageLayout;
      var textWidthTwips = pl && pl.width
        ? Math.round((pl.width - (pl.leftMargin || 0) - (pl.rightMargin || 0)) * 20)
        : 9360;

      var opts = options();
      var bodyAT = AshaarWord.templateToOoxml(tmpl, scaledTextWidth(textWidthTwips, opts.tableWidthPct), opts);
      var tagAT = AshaarWord.contentControlTag("template:" + tmpl.name, opts);
      var selection = context.document.getSelection();
      // Embed the SDT in the OOXML (wrapOoxmlControl) — see insertPoem for why
      // a post-hoc insertContentControl() clamps to the first row on Mac Word.
      var inserted = selection.insertOoxml(
        AshaarWord.wrapOoxmlControl(bodyAT, "Ashaar Poem", tagAT),
        Word.InsertLocation.end);
      await context.sync();
      await styleInsertedPoemControl(context, inserted, tagAT);
      await context.sync();
      setMessage("Template \"" + tmpl.name + "\" inserted.");
    });
  }

  // ── Delete template ────────────────────────────────────────────────────────

  function deleteTemplate() {
    var id = templateList.value;
    if (!id) return;
    var templates = loadTemplates().filter(function (t) { return t.id !== id; });
    saveTemplates(templates);
    renderTemplateList();
    setMessage("Template deleted.");
  }

  // ── Export / Import ────────────────────────────────────────────────────────

  function exportTemplates() {
    var templates = loadTemplates();
    if (!templates.length) { setMessage("No templates to export."); return; }
    var json = JSON.stringify(templates, null, 2);
    var blob = new Blob([json], { type: "application/json" });
    var url = URL.createObjectURL(blob);
    var a = document.createElement("a");
    a.href = url;
    a.download = "ashaar-templates.json";
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }

  function importTemplates() {
    importFileInput.click();
  }

  function onImportFile(event) {
    var file = event.target.files && event.target.files[0];
    if (!file) return;
    var reader = new FileReader();
    reader.onload = function (e) {
      try {
        var incoming = JSON.parse(e.target.result);
        if (!Array.isArray(incoming)) throw new Error("Expected array");
        var existing = loadTemplates();
        var existingIds = {};
        existing.forEach(function (t) { existingIds[t.id] = true; });
        var added = 0;
        incoming.forEach(function (t) {
          if (t && t.id && t.name && t.rows && !existingIds[t.id]) {
            existing.push(t);
            added++;
          }
        });
        saveTemplates(existing);
        renderTemplateList();
        setMessage("Imported " + added + " template(s).");
      } catch (err) {
        setMessage("Import failed: " + (err.message || String(err)));
      }
      importFileInput.value = "";
    };
    reader.readAsText(file);
  }

  // Settings-panel action handlers (Apply/Profile pipelines, Tasks 7-8).
  // One Apply: route by target and scope. Writes deltas to the owning tag
  // slot, then re-renders/justifies. Pending clears only on success.
  async function applyPanel() {
    var target = _panel.target;
    var values = panelValues();
    var debug = !!(debugMode && debugMode.checked);
    var run = debug ? AshaarMetrics.startRun("apply " + _panel.scopeLevel) : null;
    try {
      if (!target || target.kind !== "block") {
        // Plain selection: one-shot justify with panel values; nothing persisted.
        // (The single-face gate for this route lives inside justifySelection's
        // own no-tables capture — it reads the selection's real font there.)
        await justifySelection();   // justifySelection reads options() → panelValues()
        _panel.pending = { set: {}, clear: [] };
        refreshPanel();
        // This branch returns before the block path's tail below, which is
        // where run.end() normally fires — without this, a plain-selection
        // Apply's metrics run was started (line above) and never closed.
        if (run) {
          run.phase("justify");
          run.end();
          debugOutput.textContent += "\n" + JSON.stringify(run.report());
        }
        return;
      }
      // Just-in-time font-measurement gate (Task 9), BEFORE any routing/tag
      // write: gather this block's distinct cell font faces and confirm the
      // WebView can measure them. "cancel" aborts here — pending edits survive
      // for retry (the routing branches below, which write tags, haven't run).
      var gateFaces = await collectBlockFaceNames();
      var gateResult = await ensureFacesMeasurable(gateFaces);
      if (gateResult === "cancel") { setMessage("Add the font, then Apply again."); return; }
      if (run) run.phase("face gate");
      if (_panel.scopeLevel === "poem") {
        await applyPoemScope(target, values, run);
      } else if (_panel.scopeLevel === "bandh") {
        await withWordStrict(async function (context) {
          var cc = await findBlockAt(context);           // helper below
          cc.tag = AshaarWord.setTagBandhWidth(cc.tag, values.misraWidthPt || 0);
          await context.sync();
        });
        if (run) run.phase("tag write");
        tagWritten();
        await reapplyBlock();                            // re-justify in place
        if (run) run.phase("pipeline");
      } else if (_panel.scopeLevel === "cell") {
        var cellFillOnEl = document.getElementById("sp-cell-fill-on");
        var cellFillEl = document.getElementById("sp-cell-fill");
        var cellColorOnEl = document.getElementById("sp-cell-color-on");
        var cellColorEl = document.getElementById("sp-cell-color");
        var override = {
          strength: dirtyOrNull("strength"), widthPt: dirtyOrNull("misraWidthPt"), capEm: dirtyOrNull("capEm"),
          fill: cellFillOnEl.checked ? cellFillEl.value : null,
          color: cellColorOnEl.checked ? cellColorEl.value : null,
        };
        // Final review I2: which fields the user actually TOUCHED this
        // Apply — only these fan out onto sibling keys; every other key
        // must keep its own existing value (a bandh/poem-target Apply must
        // not delete unrelated per-cell formatting on siblings). Numeric
        // fields: dirty in the pending buffer — a ⟲-clear counts as touched
        // too (its incoming value is already null via dirtyOrNull, and that
        // null must still clear on every targeted key per spec). fill/color
        // have no pending-buffer entry (raw shared DOM state, per
        // seedCellDecorInputs) — touched is "control state differs from
        // what was last seeded" for the ACTIVE (current) cell.
        var touched = {
          strength: ("strength" in _panel.pending.set) || _panel.pending.clear.indexOf("strength") !== -1,
          widthPt: ("misraWidthPt" in _panel.pending.set) || _panel.pending.clear.indexOf("misraWidthPt") !== -1,
          capEm: ("capEm" in _panel.pending.set) || _panel.pending.clear.indexOf("capEm") !== -1,
          fill: cellFillOnEl.checked !== _seededCellDecor.fillOn ||
            (cellFillOnEl.checked && cellFillEl.value !== _seededCellDecor.fill),
          color: cellColorOnEl.checked !== _seededCellDecor.colorOn ||
            (cellColorOnEl.checked && cellColorEl.value !== _seededCellDecor.color)
        };
        await withWordStrict(async function (context) {
          var cc = await findBlockAt(context);
          var keys = cellTargetKeys(cc.tag, "content", target.cellLabel, "sp-cell-target");
          var oldPayload = AshaarWord.parseContentControlTag(cc.tag) || {};
          // The current key still fully replaces (the pane is the full
          // truth for the seeded/reflected cell); every OTHER fanned-out key
          // merges the touched fields onto ITS OWN existing override —
          // AshaarOverrides.mergeFanOutOverride (cell-overrides.js) is the
          // pure composition, pinned by tests/cell-overrides.test.js.
          var mergedByKey = {};
          keys.forEach(function (k) {
            if (k === target.cellLabel) { mergedByKey[k] = override; return; }
            var existing = (oldPayload.overrides && oldPayload.overrides[k]) || null;
            mergedByKey[k] = AshaarOverrides.mergeFanOutOverride(existing, override, touched);
          });
          // §4 transition-clear: diff old-vs-new color for EVERY fan-out key,
          // using THAT key's own merged result (untouched keys' merged color
          // equals their existing color, so they never appear here) — BEFORE
          // the tag write replaces the overrides. Keys losing their color
          // are recorded for the render pass below to reset to black.
          var clearKeys = [];
          keys.forEach(function (k) {
            clearKeys = clearKeys.concat(AshaarOverrides.colorClearKeys(oldPayload.overrides, [k], mergedByKey[k]));
          });
          var newTag = cc.tag;
          keys.forEach(function (k) { newTag = AshaarWord.setTagOverride(newTag, k, mergedByKey[k]); });
          cc.tag = newTag;
          await context.sync();
          // Record only after the write committed — a strict-throw above must
          // not leave clears queued for a tag that still carries the color.
          // Scoped to this block's identity; a retry-Apply on the SAME block
          // (same payload → same identity) MERGES so clears retained from a
          // failed render survive (the retried diff finds nothing — the tag
          // is already colorless), while a different block starts fresh.
          var blockId = AshaarWord.tagIdentity(newTag);
          if (_pendingColorClears.blockId !== blockId) _pendingColorClears = { blockId: blockId, keys: {} };
          clearKeys.forEach(function (k) { _pendingColorClears.keys[k] = true; });
        });
        if (run) run.phase("tag write");
        tagWritten();
        await reapplyBlock();
        if (run) run.phase("pipeline");
      } else if (_panel.scopeLevel === "gap") {
        await withWordStrict(async function (context) {
          var cc = await findBlockAt(context);
          var decor = {
            symbol: document.getElementById("sp-gap-symbol").value,
            fill: document.getElementById("sp-gap-fill-on").checked ? document.getElementById("sp-gap-fill").value : "",
            color: document.getElementById("sp-gap-color").value,
          };
          var keys = cellTargetKeys(cc.tag, "spacing", target.gapKey, "sp-gap-target");
          var newTag = cc.tag;
          keys.forEach(function (k) { newTag = AshaarWord.setTagSlotDecor(newTag, k, decor); });
          cc.tag = newTag;
          await context.sync();
        });
        if (run) run.phase("tag write");
        tagWritten();
        await reapplyBlock();
        if (run) run.phase("pipeline");
      }
      // Success tail — must run AFTER the pipelines: they consume pending via
      // options() → panelValues() (resolved old tag + pending overlay), so the
      // clear can only happen once they've rendered with the new values. A
      // strict tag-write throw above skips this tail (pending survives for
      // retry).
      _panel.pending = { set: {}, clear: [] };
      _pendingColorClears = { blockId: "", keys: {} };  // one-shot: consumed above
      // Task 10 m2 (upgraded by final review — compounds with I2): a stale
      // "This bandh"/"Whole poem" selection left over from a fan-out Apply
      // would silently fan out the NEXT, unrelated edit too. Reset both
      // Apply-to selects to "this" every successful Apply.
      var spCellTargetEl = document.getElementById("sp-cell-target");
      if (spCellTargetEl) spCellTargetEl.value = "this";
      var spGapTargetEl = document.getElementById("sp-gap-target");
      if (spGapTargetEl) spGapTargetEl.value = "this";
      _lastSeededDecorKey = null;  // force the decor inputs to reseed from the fresh tag
      _lastBlockTag = null;   // force reflection to re-read the updated tag
      await reflectActiveContext();
      if (run) run.phase("reflect");
      if (run) {
        run.end();
        debugOutput.textContent += "\n" + JSON.stringify(run.report());
      }
      // No blanket "Applied." here: the render/justify pipeline above sets the
      // FINAL message ("Done." / its own error), and overwriting it would mask
      // pipeline errors. Recorded limitation (final-review item): a render-
      // pipeline failure AFTER the successful tag write is not detected by
      // applyPanel — the tag state is already persisted at that point.
    } catch (e) {
      // Keep pending for retry (spec: apply failure keeps edits).
      if (run) {
        run.end();
        debugOutput.textContent += "\n" + JSON.stringify(run.report());
      }
      setMessage("Apply failed: " + (e && e.message ? e.message : e));
    }
  }

  // Tag write persisted: interim status only. Pending must NOT be cleared here
  // — the render/justify pipeline that follows consumes it (options() →
  // panelValues() overlays pending on the resolved values); applyPanel's
  // success tail clears it after the pipeline has rendered.
  function tagWritten() {
    setMessage("Settings saved — re-rendering…");
  }

  function dirtyOrNull(key) {
    // Regression guard: a pending CLEAR (⟲ on a cell-sourced key) must win over
    // the cell-source fallback below, or Apply would re-write the old override
    // value right back onto the cell and the ⟲ could never actually clear it.
    // setTagOverride treats an all-null overrides object as "delete this
    // override" — returning null here is what makes that delete happen.
    if (_panel.pending.clear.indexOf(key) !== -1) return null;
    return (key in _panel.pending.set) ? _panel.pending.set[key]
      : (_panel.resolved && _panel.resolved.source[key] === "cell" ? _panel.resolved.values[key] : null);
  }

  // Task 10: fan the current override/decor key out to "this"/"bandh"/"poem"
  // per the panel's Apply-to selector, via AshaarCellMap.keysForTarget — the
  // SAME key scheme reflectActiveCell used to derive currentKey in the first
  // place (AshaarOverrides.overrideKey(tableIndex, label|slot)). Reads
  // payload.cells fresh off the live tag (not the closed-over target.payload)
  // so it can't drift from what's about to be written.
  function cellTargetKeys(liveTag, kind, currentKey, selectId) {
    var payload = AshaarWord.parseContentControlTag(liveTag) || {};
    var tables = payload.cells || [];
    var tableIndex = parseInt(String(currentKey).split(":")[0], 10) || 0;
    var map = AshaarCellMap.buildBandhCellMap(tables[tableIndex]);
    var sel = document.getElementById(selectId);
    var mode = (sel && sel.value) || "this";
    return AshaarCellMap.keysForTarget(map, kind, mode, currentKey, tables);
  }

  // Locate the enclosing Ashaar Poem control at the cursor (throws if none).
  async function findBlockAt(context) {
    var sel = context.document.getSelection();
    var cc = sel.parentContentControlOrNullObject;
    cc.load("title,tag");
    await context.sync();
    if (cc.isNullObject || cc.title !== "Ashaar Poem") throw new Error("Click inside an Ashaar poem first.");
    return cc;
  }

  // Poem scope: persist local deltas, then rebuild-if-structural + justify.
  // Reuses reRender()'s bare-rebuild + justifySelection() fill, both of which
  // now read options() → panelValues(), i.e. the resolved values.
  async function applyPoemScope(target, values, run) {
    // Reuse AshaarPanel.STRUCTURAL_KEYS (settings-panel.js) instead of a local
    // duplicate — the duplicate previously omitted separatorPt, so a
    // separator-only Apply silently routed to justify-only and no-opped
    // (final review C1).
    var structuralDirty = AshaarPanel.STRUCTURAL_KEYS.some(function (k) {
      return (k in _panel.pending.set) || _panel.pending.clear.indexOf(k) !== -1;
    });
    await withWordStrict(async function (context) {
      var cc = await findBlockAt(context);
      var payload = AshaarWord.parseContentControlTag(cc.tag);
      var newLocal = AshaarPanel.pendingToLocal(payload.local, _panel.pending, AshaarPanel.SCOPE_FIELDS.poem);
      var tag = AshaarWord.setTagLocal(cc.tag, newLocal);
      // Snapshot the profile layer for cross-machine portability.
      var prof = payload.profile ? loadProfileStore()[payload.profile] : null;
      if (prof) tag = AshaarWord.setTagProfileCache(tag, AshaarProfiles.settingsFromProfile(prof));
      cc.tag = tag;
      await context.sync();
    });
    if (run) run.phase("tag write");
    tagWritten();
    if (structuralDirty) await reRender();     // bare rebuild + in-place justify
    else await justifySelection();             // justify only — no destructive rebuild
    if (run) run.phase("pipeline");
  }

  // Re-justify the block in place (non-structural scopes).
  async function reapplyBlock() { await justifySelection(); }

  // Assign: link the block to the selected profile; local tweaks survive.
  // Fix round 1 (reviewer): strict tag write — a failed write must not run the
  // success tail; and the message goes BEFORE applyProfileToQaseeda so the
  // pipeline's own success/failure summary stays last (same rule as applyPanel).
  async function assignProfile() {
    // Prefer the pending (chosen-but-not-yet-Assigned) value: reflection may
    // have re-rendered the dropdown between the user's choice and this click.
    var name = _panel.pendingProfile != null
      ? _panel.pendingProfile
      : document.getElementById("sp-profile").value;
    try {
      await withWordStrict(async function (context) {
        var cc = await findBlockAt(context);
        cc.tag = AshaarWord.setTagProfile(cc.tag, name);
        await context.sync();
      });
    } catch (e) {
      setMessage("Assign failed: " + (e && e.message ? e.message : e));
      return;
    }
    _panel.pendingProfile = null;  // the choice is now the tag's resolved profile
    setMessage(name ? "Assigned to \"" + name + "\" — refreshing…" : "Profile link removed.");
    if (name) await applyProfileToQaseeda(name);
    _lastBlockTag = null; await reflectActiveContext();
  }

  // Save as…: panel's resolved+pending values → new profile; block assigned;
  // local cleared (the tweaks just became the profile). The name comes from
  // the inline #sp-saveas-row — window.prompt is disallowed in Office add-in
  // webviews (modal interactions blocked by the runtime), so a prompt() here
  // silently returns null inside Word.
  async function saveAsProfile(name) {
    name = (name || "").trim();
    if (!name) return;
    var values = panelValues();
    // The store write stays first and unconditional — the profile exists even
    // if linking the block fails below.
    await putProfile(AshaarProfiles.profileFromSettings(name, values));
    if (_panel.target && _panel.target.kind === "block") {
      // Fix round 1 (reviewer): strict block-linking write. On failure the
      // pending edits SURVIVE (no clear), the user gets a retry path, and we
      // return before the apply pipeline. NOTE (adaptation from the reviewer's
      // literal instruction, which said to keep `_lastBlockTag = null` here):
      // nulling _lastBlockTag on the FAILURE path makes the subsequent
      // reflection see tag !== _lastBlockTag and wipe _panel.pending (the
      // reflection's stale-pending drop, ~line 648) — exactly the state loss
      // this fix round exists to prevent. The write failed, so the tag is
      // unchanged and there is nothing new to re-read: keep _lastBlockTag as
      // is so reflection preserves pending.
      try {
        await withWordStrict(async function (context) {
          var cc = await findBlockAt(context);
          var tag = AshaarWord.setTagProfile(cc.tag, name);
          tag = AshaarWord.setTagLocal(tag, {});
          tag = AshaarWord.setTagProfileCache(tag, AshaarProfiles.settingsFromProfile(AshaarProfiles.profileFromSettings(name, values)));
          cc.tag = tag;
          await context.sync();
        });
      } catch (e) {
        setMessage("Profile \"" + name + "\" saved, but linking this poem failed: " +
          (e && e.message ? e.message : e) + " — select the profile and click Assign to retry.");
        await reflectActiveContext();
        return;
      }
      _panel.pending = { set: {}, clear: [] };  // success only — the tweaks became the profile
      setMessage("Profile \"" + name + "\" saved.");
      await applyProfileToQaseeda(name);        // its own message stays last
      _lastBlockTag = null; await reflectActiveContext();
      return;
    }
    // Selection target: nothing to link, no apply pipeline runs — the saved
    // message cannot overwrite a pipeline failure here.
    _panel.pending = { set: {}, clear: [] };
    _lastBlockTag = null; await reflectActiveContext();
    refreshPanel();   // reflect early-returns outside Word; repaint (dropdown) either way
    setMessage("Profile \"" + name + "\" saved.");
  }

  // Update "name": push panel values into the stored profile; re-apply to all
  // its blocks (each block's own local map survives via the resolver).
  async function updateProfile() {
    var name = _panel.resolved ? _panel.resolved.profileName : "";
    if (!name) return;
    // Store write is safe (no Word round-trip) — putProfile + pending-clear
    // stay as-is. Fix round 1 (reviewer): message BEFORE applyProfileToQaseeda
    // so the pipeline's own success/failure summary stays last.
    await putProfile(AshaarProfiles.profileFromSettings(name, panelValues()));
    _panel.pending = { set: {}, clear: [] };
    setMessage("Profile \"" + name + "\" updated — refreshing its poems…");
    await applyProfileToQaseeda(name);
    _lastBlockTag = null; await reflectActiveContext();
  }

  // Restore a missing profile from the tag's cached snapshot.
  async function restoreProfileFromPoem() {
    var t = _panel.target;
    if (!t || t.kind !== "block" || !t.payload || !t.payload.profileCache) return;
    var name = t.payload.profile;
    await putProfile(AshaarProfiles.profileFromSettings(name, t.payload.profileCache));
    _lastBlockTag = null; await reflectActiveContext();
    setMessage("Profile \"" + name + "\" restored from this poem.");
  }

  // Revert to profile / Reset to defaults: clear the whole local map.
  // Pending clears up-front by design: revert = the user chose to discard edits.
  // Fix round 1 (reviewer): strict tag write — on failure skip reRender; on
  // success the message goes BEFORE reRender so the pipeline's own message
  // (or its error) stays last.
  async function revertToProfile() {
    _panel.pending = { set: {}, clear: [] };
    if (_panel.target && _panel.target.kind === "block") {
      try {
        await withWordStrict(async function (context) {
          var cc = await findBlockAt(context);
          cc.tag = AshaarWord.setTagLocal(cc.tag, {});
          await context.sync();
        });
      } catch (e) {
        setMessage("Revert failed: " + (e && e.message ? e.message : e));
        return;
      }
      setMessage("Reverted — re-rendering…");
      await reRender();     // structure may change (gap/width may fall back)
      _lastBlockTag = null; await reflectActiveContext();
    } else {
      refreshPanel();
      setMessage("Reverted.");
    }
  }

  // Gap body's "Set as default for all bandhs": writes the current sp-gap-*
  // decoration into the ASSIGNED profile's spacingDecor (explicit-mutation
  // rule — this button edits the profile, not the block). Adaptation from the
  // brief: profile.spacingDecor is keyed by the bare slot label (e.g. "A#1" —
  // see profiles.js defaultProfile() comment and the render-time lookup
  // `profile.spacingDecor[c.slot]` a few hundred lines up), NOT by
  // `_panel.target.gapKey`, which is the table-index-prefixed override key
  // (`AshaarOverrides.overrideKey(tIdx, slot)`, e.g. "0:A#1") used for the
  // block-tag-level slotDecor override. Storing under gapKey would silently
  // never resolve at render time (no cell's `c.slot` is ever prefixed). The
  // bare slot is available as the module-level `_activeSlot`, set in lockstep
  // with `_panel.target.gapKey` inside reflectActiveCell, so it's still fresh
  // whenever gapKey is truthy. Kept the old slot-decor-save-profile handler's
  // shape (profile lookup by assigned name, delete-if-empty else set, save,
  // re-apply) — only the input source and key changed.
  async function saveGapDefaultToProfile() {
    var name = _panel.resolved ? _panel.resolved.profileName : "";
    if (!name || !_panel.target || _panel.target.kind !== "block" || !_panel.target.gapKey) {
      setMessage("Assign this poem to a profile first to set a profile-wide default.");
      return;
    }
    var profile = getProfile(name);
    profile.spacingDecor = profile.spacingDecor || {};
    var d = {
      symbol: document.getElementById("sp-gap-symbol").value,
      fill: document.getElementById("sp-gap-fill-on").checked ? document.getElementById("sp-gap-fill").value : "",
      color: document.getElementById("sp-gap-color").value,
    };
    if (d.symbol || d.fill || d.color) profile.spacingDecor[_activeSlot] = d;
    else delete profile.spacingDecor[_activeSlot];
    await putProfile(profile);
    await applyProfileToQaseeda(name);
    setMessage("Default gap decoration saved to \"" + name + "\".");
  }

  var isBound = false;

  function bind() {
    if (isBound) return;
    isBound = true;
    [input, bandhCount, misraCount, layoutPreset, layoutSpec].forEach(function (el) {
      el.addEventListener("input", renderPreview);
      el.addEventListener("change", renderPreview);
    });
    layoutPreset.addEventListener("change", applyLayoutPreset);
    misraCount.addEventListener("change", applyLayoutPreset);
    modeTable.addEventListener("click", function () { setMode("table"); });
    modeConvert.addEventListener("click", function () { setMode("convert"); });
    document.getElementById("insert-structure").addEventListener("click", insertStructure);
    document.getElementById("insert-poem").addEventListener("click", function () { insertPoem(false); });
    document.getElementById("insert-tabstop").addEventListener("click", insertTabStopPoem);
    document.getElementById("replace-selection").addEventListener("click", function () { insertPoem(true); });
    // "justify-selection" / "re-render" / cell-ov / bandh-ov / slot-decor /
    // qaseeda-panel button bindings removed here — those controls no longer
    // exist in taskpane.html (retired for the unified Settings panel). The
    // functions they called (justifySelection, reRender) are still very much
    // alive: the panel's "Apply" button (sp-apply, below) calls them.
    var showMapBtn = document.getElementById("show-cell-map");
    if (showMapBtn) showMapBtn.addEventListener("click", showCellMap);
    if (typeof Office !== "undefined" && Office.context && Office.context.document &&
        Office.context.document.addHandlerAsync && typeof Word !== "undefined") {
      Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, onSelectionChanged);
    }
    document.getElementById("reset-justification").addEventListener("click", resetJustification);
    document.getElementById("load-selection").addEventListener("click", loadSelection);
    // Import-options (separator flexibility): auto-normalize on paste; manual overrides.
    input.addEventListener("paste", function () { setTimeout(applyImportNormalization, 0); });
    sepMode.addEventListener("change", function () {
      sepCustom.hidden = sepMode.value !== "custom";
      applyImportNormalization();
    });
    sepCustom.addEventListener("change", applyImportNormalization);
    sepPair.addEventListener("change", applyImportNormalization);
    document.getElementById("sep-apply").addEventListener("click", applyImportNormalization);
    document.getElementById("drop-grid").addEventListener("click", insertBareGrid);
    layoutViewGridBtn.addEventListener("click", function () { setLayoutView("grid"); });
    layoutViewNumbersBtn.addEventListener("click", function () { setLayoutView("numbers"); });
    layoutGridEl.addEventListener("click", onLayoutGridClick);
    document.getElementById("adopt-table").addEventListener("click", adoptTable);
    document.getElementById("capture-template").addEventListener("click", captureSelectedTableLayout);
    document.getElementById("apply-template").addEventListener("click", applyTemplate);
    document.getElementById("delete-template").addEventListener("click", deleteTemplate);
    document.getElementById("export-templates").addEventListener("click", exportTemplates);
    document.getElementById("import-templates").addEventListener("click", importTemplates);
    importFileInput.addEventListener("change", onImportFile);

    // Settings panel: every data-key control feeds the pending buffer.
    // (:not(.sp-src) — the provenance dots share data-key with their controls
    // and must not get dead change listeners.)
    document.querySelectorAll("#settings-panel [data-key]:not(.sp-src)").forEach(function (input) {
      input.addEventListener("change", function () {
        var key = input.getAttribute("data-key");
        var raw = input.value;
        var val = input.type === "number" || input.type === "range"
          ? (raw === "" ? null : Number(raw)) : raw;
        _panel.pending = AshaarPanel.mergePending(_panel.pending, key, val);
        refreshPanel();
      });
    });
    // Provenance dots double as per-setting reset.
    document.querySelectorAll("#settings-panel .sp-src").forEach(function (span) {
      span.addEventListener("click", function () {
        _panel.pending = AshaarPanel.mergePending(_panel.pending, span.getAttribute("data-key"), null);
        refreshPanel();
      });
    });
    // Advanced: font correction (fontCorrections) and debug colors (debugColors)
    // are compound values (a map / a two-color object) — the generic scalar
    // [data-key] listener above can't build them, so these inputs deliberately
    // carry no data-key and get their own listeners here.
    var corrFontEl = document.getElementById("sp-corr-font");
    var corrFactorEl = document.getElementById("sp-corr-factor");
    function applyCorrFactor() {
      var fontName = (corrFontEl.value || "").trim();
      if (!fontName) return; // nothing to key the correction on
      var factor = Number(corrFactorEl.value);
      var map = {};
      var cur = panelValues().fontCorrections || {};
      Object.keys(cur).forEach(function (k) { map[k] = cur[k]; });
      // 1.0 (or garbage) = no correction — delete rather than store a no-op.
      if (!isFinite(factor) || factor === 1) delete map[fontName];
      else map[fontName] = factor;
      _panel.pending = AshaarPanel.mergePending(_panel.pending, "fontCorrections", map);
      refreshPanel();
    }
    corrFontEl.addEventListener("change", applyCorrFactor);
    corrFactorEl.addEventListener("change", applyCorrFactor);

    var debugTatweelEl = document.getElementById("sp-debug-tatweel");
    var debugTatweelOnEl = document.getElementById("sp-debug-tatweel-on");
    var debugSpaceEl = document.getElementById("sp-debug-space");
    var debugSpaceOnEl = document.getElementById("sp-debug-space-on");
    function applyDebugColors() {
      var obj = {
        tatweel: debugTatweelOnEl.checked ? debugTatweelEl.value : "",
        space: debugSpaceOnEl.checked ? debugSpaceEl.value : "",
      };
      _panel.pending = AshaarPanel.mergePending(_panel.pending, "debugColors", obj);
      refreshPanel();
    }
    [debugTatweelEl, debugTatweelOnEl, debugSpaceEl, debugSpaceOnEl].forEach(function (el) {
      el.addEventListener("change", applyDebugColors);
    });
    ["poem", "bandh", "cell", "gap"].forEach(function (lvl) {
      document.getElementById("sp-chip-" + lvl).addEventListener("click", function () {
        _panel.scopeLevel = lvl;
        _panel.pending = { set: {}, clear: [] }; // scope switch discards unapplied edits
        refreshPanel();
      });
    });
    document.getElementById("sp-revert").addEventListener("click", revertToProfile);
    document.getElementById("sp-rerender").addEventListener("click", reRender);
    document.getElementById("sp-apply").addEventListener("click", applyPanel);
    // The dropdown choice is pending until Assign — reflection re-renders the
    // panel on every Word selection change and must not wipe it.
    document.getElementById("sp-profile").addEventListener("change", function () {
      var resolvedName = _panel.resolved ? _panel.resolved.profileName : "";
      var v = document.getElementById("sp-profile").value;
      _panel.pendingProfile = (v === resolvedName) ? null : v;
    });
    document.getElementById("sp-profile-assign").addEventListener("click", assignProfile);
    // Save as… uses an inline name row (window.prompt is disallowed in Office
    // add-in webviews; it silently returns null inside Word).
    var saveasRow = document.getElementById("sp-saveas-row");
    var saveasName = document.getElementById("sp-saveas-name");
    function hideSaveasRow() { saveasRow.hidden = true; saveasName.value = ""; }
    document.getElementById("sp-profile-saveas").addEventListener("click", function () {
      saveasRow.hidden = false;
      saveasName.focus();
    });
    document.getElementById("sp-saveas-ok").addEventListener("click", function () {
      var name = saveasName.value.trim();
      if (!name) { saveasName.focus(); return; }
      hideSaveasRow();
      saveAsProfile(name);
    });
    document.getElementById("sp-saveas-cancel").addEventListener("click", hideSaveasRow);
    saveasName.addEventListener("keydown", function (ev) {
      if (ev.key === "Enter") { ev.preventDefault(); document.getElementById("sp-saveas-ok").click(); }
      if (ev.key === "Escape") hideSaveasRow();
    });
    document.getElementById("sp-profile-update").addEventListener("click", updateProfile);
    document.getElementById("sp-profile-restore").addEventListener("click", restoreProfileFromPoem);
    var gapDefaultBtn = document.getElementById("sp-gap-default");
    if (gapDefaultBtn) gapDefaultBtn.addEventListener("click", saveGapDefaultToProfile);
    var cellCaptureBtn = document.getElementById("sp-cell-capture");
    if (cellCaptureBtn) cellCaptureBtn.addEventListener("click", captureCellFormatting);
    var gapCaptureBtn = document.getElementById("sp-gap-capture");
    if (gapCaptureBtn) gapCaptureBtn.addEventListener("click", captureCellFormatting);
    document.getElementById("adopt-replace-selection").addEventListener("click", function () { insertPoem(true); });

    // Custom fonts: register any stored fonts before measurement, wire the UI.
    if (typeof AshaarFontStore !== "undefined") {
      AshaarFontStore.registerAll().then(refreshFontList, function () {});
    }
    if (fontUpload) fontUpload.addEventListener("change", onFontFilePicked);
    var fontAddBtn = document.getElementById("font-upload-add");
    if (fontAddBtn) fontAddBtn.addEventListener("click", addFont);
    var fontVerifyBtn = document.getElementById("font-upload-verify");
    if (fontVerifyBtn) fontVerifyBtn.addEventListener("click", verifyFontAtCursor);

    applyLayoutPreset();
    renderPreview();
    setMode("table");
    renderTemplateList();
    refreshPanel(); // initial render: no active block yet → "Selection", defaults
  }

  if (window.Office && Office.onReady) {
    Office.onReady(function (info) {
      hostStatus.textContent = info.host === Office.HostType.Word ? "Connected to Word" : "Preview mode";
      bind();
    });
    window.setTimeout(function () {
      if (!isBound) {
        hostStatus.textContent = "Browser preview mode";
        bind();
      }
    }, 1200);
  } else {
    hostStatus.textContent = "Browser preview mode";
    bind();
  }
}());

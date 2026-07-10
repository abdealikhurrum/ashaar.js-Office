(function () {
  var input = document.getElementById("poem-input");
  var preview = document.getElementById("preview");
  var message = document.getElementById("message");
  var hostStatus = document.getElementById("host-status");
  var modeTable = document.getElementById("mode-table");
  var modeConvert = document.getElementById("mode-convert");
  var tablePanel = document.getElementById("table-mode-panel");
  var convertPanel = document.getElementById("convert-mode-panel");
  var justifyMode = document.getElementById("justify-mode");
  var layoutMode = document.getElementById("layout-mode");
  var widthMode = document.getElementById("width-mode");
  var bandhCount = document.getElementById("bandh-count");
  var misraCount = document.getElementById("misra-count");
  var layoutPreset = document.getElementById("layout-preset");
  var layoutSpec = document.getElementById("layout-spec");
  var fontMode = document.getElementById("font-mode");
  var tatweelCount = document.getElementById("tatweel-count");
  var tatweelValue = document.getElementById("tatweel-value");
  var gapWidth = document.getElementById("gap-width");
  var tableWidth = document.getElementById("table-width");
  var tableWidthValue = document.getElementById("table-width-value");
  var autoFitWidth = document.getElementById("auto-fit-width");

  // Table width as a fraction of the page text column (centred). 100% = full width.
  function tableWidthPct() {
    return Math.max(25, Math.min(100, Number((tableWidth && tableWidth.value) || 50)));
  }
  function scaledTextWidth(twips) {
    return Math.max(1, Math.round(twips * tableWidthPct() / 100));
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
    var GRID = N * 3 + (N - 1) * gapCols, contentCols = N * 3;
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
  var qaseedaName = document.getElementById("qaseeda-name");
  var qaseedaNames = document.getElementById("qaseeda-names");
  var qaseedaWidthMode = document.getElementById("qaseeda-width-mode");
  var qaseedaWidthPct = document.getElementById("qaseeda-width-pct");
  var qaseedaJustifyMode = document.getElementById("qaseeda-justify-mode");
  var qaseedaStrength = document.getElementById("qaseeda-strength");
  var qaseedaStrengthValue = document.getElementById("qaseeda-strength-value");
  var qaseedaCorrFont = document.getElementById("qaseeda-corr-font");
  var qaseedaCorrFactor = document.getElementById("qaseeda-corr-factor");
  var qaseedaDebugTatweel = document.getElementById("qaseeda-debug-tatweel");
  var qaseedaDebugTatweelOn = document.getElementById("qaseeda-debug-tatweel-on");
  var qaseedaDebugSpace = document.getElementById("qaseeda-debug-space");
  var qaseedaDebugSpaceOn = document.getElementById("qaseeda-debug-space-on");
  var qaseedaFontStatus = document.getElementById("qaseeda-font-status");

  // Format collected per-cell justification metrics into the Debug panel.
  function renderDebug(diags) {
    if (!debugOutput) return;
    if (!diags.length) { debugOutput.textContent = "(no kashida cells measured)"; return; }
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
    debugOutput.textContent = head + "\n" + rows.join("\n");
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

  function options() {
    return {
      justifyMode: justifyMode.value,
      justify: justifyMode.value === "none" ? false : justifyMode.value,
      layoutMode: layoutMode.value,
      layout: layoutMode.value,
      widthMode: widthMode.value,
      bandhCount: Number(bandhCount.value || 1),
      misraCount: Number(misraCount.value || 4),
      misraPattern: layoutPreset.value,
      layoutSpec: (!tablePanel.hidden) ? layoutSpec.value : "",
      fontMode: fontMode.value,
      tatweelCount: Number(tatweelCount.value || 0),
      gapWidth: Number(gapWidth.value || 4),
      tableWidthPct: tableWidthPct(),
      autoFitWidth: !!(autoFitWidth && autoFitWidth.checked),
      qaseeda: (qaseedaName && qaseedaName.value ? qaseedaName.value.trim() : "")
    };
  }

  function previewFontFamily(font) {
    var mode = font === "nastaliq" ? "noto" : font;
    var css = AshaarFonts.cssFamilyOf(mode);
    return css || "\"Times New Roman\", serif";
  }

  function updateFontNote() {
    var d = AshaarFonts.get(fontMode.value === "nastaliq" ? "noto" : fontMode.value);
    var note = document.getElementById("font-install-note");
    if (!note) return;
    if (d && d.wordName && d.mechanism !== "whitespace") {
      note.hidden = false;
      note.textContent = "Readers need “" + d.wordName + "” installed to see this correctly."
        + (d.id === "jameel" ? " Specifically the Kasheeda build, or italic runs won’t elongate." : "");
    } else { note.hidden = true; }
  }

  function setMessage(text) {
    message.textContent = text || "";
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
    tatweelValue.textContent = String(opts.tatweelCount);
    if (tableWidthValue) tableWidthValue.textContent = String(opts.tableWidthPct);
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

  async function withWord(callback) {
    if (typeof Word === "undefined") {
      setMessage("Open this task pane inside Word to update the document.");
      return;
    }
    try {
      await Word.run(callback);
      setMessage("Done.");
    } catch (error) {
      setMessage(error && error.message ? error.message : String(error));
    }
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

  // Assign (or clear, with "") a qaseeda name onto the Ashaar Poem block at the cursor.
  async function setQaseedaOnSelection(name) {
    var done = false;
    await withWord(async function (context) {
      var cc = context.document.getSelection().parentContentControlOrNullObject;
      cc.load("title,tag");
      await context.sync();
      if (cc.isNullObject || cc.title !== "Ashaar Poem") {
        setMessage("Place the cursor inside an Ashaar Poem block first.");
        return;
      }
      cc.tag = AshaarWord.setTagQaseeda(cc.tag, name);
      await context.sync();
      done = true;
    });
    return done;
  }

  // Read the qaseeda name on the Ashaar Poem block at the cursor ("" if none).
  async function getQaseedaAtSelection() {
    var qaseeda = "";
    if (typeof Word === "undefined") return qaseeda;
    try {
      await Word.run(async function (context) {
        var cc = context.document.getSelection().parentContentControlOrNullObject;
        cc.load("title,tag");
        await context.sync();
        if (!cc.isNullObject && cc.title === "Ashaar Poem") {
          var payload = AshaarWord.parseContentControlTag(cc.tag);
          qaseeda = (payload && payload.qaseeda) || "";
        }
      });
    } catch (e) { /* leave qaseeda empty */ }
    return qaseeda;
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
      return !!(name && p && p.qaseeda === name);
    });
  }

  // Apply a qaseeda's profile across ALL its blocks so they stay consistent:
  // size every block's table to one shared width (auto-fit to the widest cell
  // across all blocks with kashida headroom, capped at the page) and re-justify
  // each cell with the profile's params. Additive — leaves justifySelection
  // untouched. In-place resize needs WordApiDesktop 1.3; without it, only the
  // re-justify runs (at current widths).
  async function applyProfileToQaseeda(name) {
    if (typeof Word === "undefined") { setMessage("Open this task pane inside Word to apply a qaseeda."); return; }
    var profile = getProfile(name);
    var CELL_MARGIN_PT = 5.76;
    var targetFill = AshaarProfiles.strengthToTargetFill(profile.justify.strength);
    var doKashida = profile.justify.mode === "kashida";
    var fallbackName = profile.font || "Times New Roman";
    var summary = "";

    try {
      await Word.run(async function (context) {
        var blocks = await gatherQaseedaBlocks(context, name);
        if (!blocks.length) { summary = "No blocks are tagged with qaseeda “" + name + "”."; return; }

        var section = context.document.sections.getFirst();
        section.load("pageLayout/width,pageLayout/leftMargin,pageLayout/rightMargin");

        // Load each block's tables → rows → cells (text, real font, columnWidth).
        var blockTables = blocks.map(function (cc) { var t = cc.getRange().tables; t.load("items"); return t; });
        await context.sync();
        blockTables.forEach(function (t) { t.items.forEach(function (tbl) { tbl.rows.load("items"); }); });
        await context.sync();
        var allTables = [];
        blockTables.forEach(function (t) { t.items.forEach(function (tbl) { allTables.push(tbl); }); });
        if (!allTables.length) { summary = "Qaseeda “" + name + "” has no tables to size."; return; }

        // Stage 1: load each row's cells (must sync before reading cells.items).
        allTables.forEach(function (tbl) {
          tbl.rows.items.forEach(function (row) { row.cells.load("items/columnWidth"); });
        });
        await context.sync();
        // Stage 2: now cells.items is available — load each cell's text + real font.
        allTables.forEach(function (tbl) {
          tbl.rows.items.forEach(function (row) {
            row.cells.items.forEach(function (cell) { cell.body.load("text"); cell.body.font.load("name,size"); });
          });
        });
        await context.sync();

        // Capture cell proxies + their text/font as plain values ONCE. Later
        // collection reloads (for the resized columnWidth) drop body.text from
        // freshly-derived items, so we never re-read body.* — only columnWidth
        // (which survives on the captured proxy) and insertText (a method).
        var tableInfos = allTables.map(function (tbl) {
          var cells = [];
          tbl.rows.items.forEach(function (row) {
            row.cells.items.forEach(function (cell) {
              var f = cell.body.font;
              var current = (cell.body.text || "").trim();
              var base = stripJustification(current);
              cells.push({
                cell: cell,
                current: current,
                base: base,
                measure: base.replace(/\s+/g, " ").trim(),
                fontName: (f && f.name) || "",
                fontSize: (f && f.size) || 0
              });
            });
          });
          return { tbl: tbl, cells: cells };
        });

        var pl = section.pageLayout;
        var pagePt = pl && pl.width ? (pl.width - (pl.leftMargin || 0) - (pl.rightMargin || 0)) : 468;

        // Representative font for the canvas baseline (first captured cell with one).
        var repName = fallbackName, repSize = 16;
        for (var ti = 0; ti < tableInfos.length && repName === fallbackName; ti++) {
          var cs0 = tableInfos[ti].cells;
          for (var ci = 0; ci < cs0.length; ci++) {
            if (cs0[ci].fontName) { repName = cs0[ci].fontName; if (cs0[ci].fontSize) repSize = cs0[ci].fontSize; break; }
          }
        }

        var canvasCtx = document.createElement("canvas").getContext("2d");
        if (!canvasCtx) { summary = "Canvas unavailable; cannot measure."; return; }
        canvasCtx.font = repSize + "pt \"" + repName + "\"";
        if (document.fonts && document.fonts.load) { try { await document.fonts.load(repSize + "pt \"" + repName + "\""); } catch (e) {} }

        var canResize = profile.width.mode === "auto-fit"
          && (typeof Office !== "undefined" && Office.context && Office.context.requirements
            && Office.context.requirements.isSetSupported
            && Office.context.requirements.isSetSupported("WordApiDesktop", "1.3"));

        // Resize. Each block's tightest cell gives the scale it needs for kashida
        // headroom. When every block has the same column count we equalise PER
        // COLUMN (column j identical across all blocks); otherwise we fall back to
        // an equal total width with proportional columns. Capped at the page.
        if (canResize) {
          var headroom = doKashida ? 0.9 : 0.98;
          var colColls = tableInfos.map(function (info) { var c = info.tbl.columns; c.load("items/width"); return c; });
          await context.sync();
          var needScales = tableInfos.map(function (info) {
            var needScale = 1;
            info.cells.forEach(function (c) {
              if (!c.measure) return;
              var fname = c.fontName || repName;
              canvasCtx.font = (c.fontSize || repSize) + "pt \"" + fname + "\"";
              var colWpx = (c.cell.columnWidth || 0) * 96 / 72;
              // Per-font correction nudges the measured width for fonts the WebView
              // can't resolve accurately (see profile.fontCorrections).
              var measured = AshaarProfiles.applyFontCorrection(canvasCtx.measureText(c.measure).width, fname, profile.fontCorrections);
              if (colWpx > 0) needScale = Math.max(needScale, measured / (headroom * colWpx));
            });
            return needScale;
          });
          var colCounts = colColls.map(function (c) { return c.items.length; });
          var sameShape = colCounts[0] > 0 && colCounts.every(function (n) { return n === colCounts[0]; });

          if (sameShape) {
            // Per-column equality: column j width = max over blocks of (its width × that
            // block's needed scale). Every block then gets the identical width vector.
            var N = colCounts[0];
            var sharedCol = [];
            for (var j = 0; j < N; j++) {
              var m = 0;
              colColls.forEach(function (c, idx) { m = Math.max(m, (c.items[j].width || 0) * needScales[idx]); });
              sharedCol.push(m);
            }
            var total = sharedCol.reduce(function (a, b) { return a + b; }, 0);
            if (total > pagePt && total > 0) { var k = pagePt / total; sharedCol = sharedCol.map(function (w) { return w * k; }); }
            colColls.forEach(function (c) {
              c.items.forEach(function (col, jj) { col.width = Math.round(sharedCol[jj] * 100) / 100; });
            });
          } else {
            // Mixed shapes: equalise total width, scale each block's columns proportionally.
            var perTotal = colColls.map(function (c, idx) {
              var t = 0; c.items.forEach(function (col) { t += (col.width || 0); });
              return t * needScales[idx];
            });
            var targetWidthPt = Math.min(pagePt, Math.max.apply(null, perTotal.concat([0])));
            colColls.forEach(function (c, idx) {
              var cur = 0; c.items.forEach(function (col) { cur += (col.width || 0); });
              if (cur <= 0 || targetWidthPt <= 0) return;
              var scale = targetWidthPt / cur;
              if (Math.abs(scale - 1) < 0.005) return;
              c.items.forEach(function (col) { col.width = Math.round(col.width * scale * 100) / 100; });
            });
          }
          await context.sync();
          // Re-read columnWidth on the captured proxies (justifySelection-proven).
          tableInfos.forEach(function (info) { info.tbl.rows.items.forEach(function (row) { row.cells.load("items/columnWidth"); }); });
          await context.sync();
        }

        // Re-justify every cell with the profile's params (captured values only).
        var changed = 0;
        var calibParams = { targetFill: targetFill };
        tableInfos.forEach(function (info) {
          info.cells.forEach(function (c) {
            if (!c.base) return;
            var colPx = Math.max(1, (c.cell.columnWidth || 0) - 2 * CELL_MARGIN_PT) * 96 / 72;
            var justified = c.base;
            if (doKashida && colPx > 0) {
              var fname = c.fontName || repName;
              canvasCtx.font = (c.fontSize || repSize) + "pt \"" + fname + "\"";
              justified = AshaarJustify.justifyLine(c.base, colPx, canvasCtx, calibParams, null);
            }
            if (justified !== c.current) {
              c.cell.body.paragraphs.getFirst().insertText(justified, Word.InsertLocation.replace);
              changed++;
            }
          });
        });
        await context.sync();

        // Debug colouring: tint the tatweels / micro-spaces that justification
        // inserted, so they're visible. Best-effort via Word search (the spaces
        // are exact code points; some hosts may not surface them).
        var tatColor = (profile.debugColors && profile.debugColors.tatweel) || "";
        var spcColor = (profile.debugColors && profile.debugColors.space) || "";
        var coloured = 0;
        if (tatColor || spcColor) {
          var hits = [];
          tableInfos.forEach(function (info) {
            info.cells.forEach(function (c) {
              if (tatColor) { var st = c.cell.body.search("ـ"); st.load("items"); hits.push({ s: st, color: tatColor }); }
              if (spcColor) {
                var sh = c.cell.body.search(" "); sh.load("items"); hits.push({ s: sh, color: spcColor });
                var sn = c.cell.body.search(" "); sn.load("items"); hits.push({ s: sn, color: spcColor });
              }
            });
          });
          await context.sync();
          hits.forEach(function (h) { h.s.items.forEach(function (r) { r.font.color = h.color; coloured++; }); });
          await context.sync();
        }

        summary = "Applied qaseeda “" + name + "” to " + blocks.length + " block(s); justified " + changed + " cell(s)"
          + (coloured ? "; coloured " + coloured + " artifact(s)" : "")
          + (canResize ? "." : " (widths unchanged — desktop Word needed to resize).");
      });
      // Hybrid refresh: remember that this qaseeda was applied (widths cached lazily next pass).
      if (profile.width.mode === "auto-fit") { profile.derived = profile.derived || {}; await putProfile(profile); }
    } catch (error) {
      summary = "Apply failed: " + (error && error.message ? error.message : String(error));
    }
    setMessage(summary);
  }

  // ── Qaseeda panel ↔ profile (P4 UI) ───────────────────────────────────────
  function panelToProfile() {
    var p = AshaarProfiles.defaultProfile((qaseedaName.value || "").trim());
    p.width.mode = qaseedaWidthMode.value;
    p.width.pct = Number(qaseedaWidthPct.value || 50);
    p.justify.mode = qaseedaJustifyMode.value;
    p.justify.strength = Number(qaseedaStrength.value || 0);
    var corrFont = (qaseedaCorrFont.value || "").trim();
    p.fontCorrections = {};
    if (corrFont) p.fontCorrections[corrFont] = Number(qaseedaCorrFactor.value || 1);
    p.debugColors = {
      tatweel: (qaseedaDebugTatweelOn && qaseedaDebugTatweelOn.checked) ? (qaseedaDebugTatweel.value || "") : "",
      space: (qaseedaDebugSpaceOn && qaseedaDebugSpaceOn.checked) ? (qaseedaDebugSpace.value || "") : ""
    };
    return AshaarProfiles.normalizeProfile(p);
  }

  function profileToPanel(profile) {
    var p = AshaarProfiles.normalizeProfile(profile || {});
    qaseedaWidthMode.value = p.width.mode;
    qaseedaWidthPct.value = p.width.pct;
    qaseedaJustifyMode.value = p.justify.mode;
    qaseedaStrength.value = p.justify.strength;
    qaseedaStrengthValue.textContent = p.justify.strength;
    var fonts = Object.keys(p.fontCorrections || {});
    qaseedaCorrFont.value = fonts[0] || "";
    qaseedaCorrFactor.value = fonts[0] ? p.fontCorrections[fonts[0]] : 1;
    var dc = p.debugColors || {};
    if (qaseedaDebugTatweelOn) qaseedaDebugTatweelOn.checked = !!dc.tatweel;
    if (dc.tatweel) qaseedaDebugTatweel.value = dc.tatweel;
    if (qaseedaDebugSpaceOn) qaseedaDebugSpaceOn.checked = !!dc.space;
    if (dc.space) qaseedaDebugSpace.value = dc.space;
  }

  function populateQaseedaNames() {
    if (!qaseedaNames) return;
    var names = listProfileNames();
    qaseedaNames.innerHTML = names.map(function (n) {
      return "<option value=\"" + String(n).replace(/"/g, "&quot;") + "\">";
    }).join("");
  }

  function loadQaseedaIntoPanel() {
    var name = (qaseedaName.value || "").trim();
    if (!name) return;
    var store = loadProfileStore();
    if (store[name]) profileToPanel(store[name]);
  }

  async function saveAndApplyQaseeda() {
    var p = panelToProfile();
    if (!p.name) { setMessage("Name the qaseeda first."); return; }
    await putProfile(p);
    populateQaseedaNames();
    await applyProfileToQaseeda(p.name);
  }

  async function assignBlockToQaseeda() {
    var name = (qaseedaName.value || "").trim();
    if (!name) { setMessage("Name the qaseeda first."); return; }
    if (!loadProfileStore()[name]) await putProfile(panelToProfile());
    await setQaseedaOnSelection(name);
    populateQaseedaNames();
  }

  function setQaseedaFontStatus(text, kind) {
    if (!qaseedaFontStatus) return;
    qaseedaFontStatus.textContent = text;
    qaseedaFontStatus.className = "qaseeda-font-status" + (kind === "ok" ? " is-ok" : kind === "warn" ? " is-warn" : "");
  }

  // Check whether the font of the block/selection at the cursor resolves in the
  // WebView; if not, justify metrics for it are only approximate.
  async function checkQaseedaFont() {
    if (typeof Word === "undefined") { setQaseedaFontStatus("Open in Word to check.", "warn"); return; }
    var fontName = "";
    try {
      await Word.run(async function (context) {
        var sel = context.document.getSelection();
        var cc = sel.parentContentControlOrNullObject;
        cc.load("title");
        await context.sync();
        var range = (!cc.isNullObject && cc.title === "Ashaar Poem") ? cc.getRange() : sel;
        range.font.load("name");
        await context.sync();
        fontName = range.font.name || "";
      });
    } catch (e) { /* ignore */ }
    if (!fontName) { setQaseedaFontStatus("No font found at the cursor.", "warn"); return; }
    if (document.fonts && document.fonts.load) { try { await document.fonts.load("16pt \"" + fontName + "\""); } catch (e) {} }
    if (fontAvailable(fontName)) setQaseedaFontStatus("“" + fontName + "” resolves — metrics are accurate.", "ok");
    else setQaseedaFontStatus("“" + fontName + "” is NOT resolvable here — metrics are approximate. Add its font file under Custom fonts.", "warn");
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

  async function insertPoem(replaceSelection) {
    var pendingMsg = "";
    await withWord(async function (context) {
      var opts = options();
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
        textWidthTwips = scaledTextWidth(pageTwips);
        if (ctxP && textWidthTwips < neededTwips) {
          pendingMsg = "Inserted — but this width is tight for " + fontSizeP +
            "pt; widen to ~" + Math.min(100, Math.round(neededTwips / pageTwips * 100)) +
            "% or turn on Auto-fit for full kashida.";
        }
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

      var ooxml = AshaarWord.wrapOoxml(ooxmlBody);
      var selection = context.document.getSelection();
      var inserted = selection.insertOoxml(ooxml,
        replaceSelection ? Word.InsertLocation.replace : Word.InsertLocation.end);
      var control = inserted.insertContentControl();
      control.title = "Ashaar Poem";
      control.tag = AshaarWord.contentControlTag(source, opts);
      control.appearance = "BoundingBox";
      await context.sync();
    });
    // withWord sets "Done."; surface the width nudge after it if one was raised.
    if (pendingMsg) setMessage(pendingMsg);
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
        var twGs = scaledTextWidth(twG);
        for (var bi = 0; bi < countG; bi++) bodyG.push(AshaarWord.templateToOoxml(tmplG, twGs, opts));
        var selG = context.document.getSelection();
        var insG = selG.insertOoxml(AshaarWord.wrapOoxml(bodyG.join("<w:p/>")), Word.InsertLocation.end);
        var ccG = insG.insertContentControl();
        ccG.title = "Ashaar Poem";
        ccG.tag = AshaarWord.contentControlTag("grid", opts);
        ccG.appearance = "BoundingBox";
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
        var scaled = scaledTextWidth(textWidthTwips);
        var ooxmlBody = tables.map(function (t) {
          return t.spanBased
            ? AshaarWord.templateToOoxml(t, scaled, opts)
            : AshaarWord.layoutTableToOoxml(t, scaled, opts);
        }).join("<w:p/>");
        var selection = context.document.getSelection();
        var inserted = selection.insertOoxml(AshaarWord.wrapOoxml(ooxmlBody), Word.InsertLocation.end);
        var control = inserted.insertContentControl();
        control.title = "Ashaar Poem";
        control.tag = AshaarWord.contentControlTag("template", opts);
        control.appearance = "BoundingBox";
        await context.sync();
        return;
      }

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

  async function justifySelection() {
    var opts = options();
    var fontId = opts.fontMode === "nastaliq" ? "noto" : opts.fontMode;
    var mechanism = AshaarFonts.mechanismOf(fontId);

    // Hybrid qaseeda trigger: if the cursor's block belongs to a qaseeda that has
    // a stored profile, justify by applying that profile across ALL its blocks so
    // they stay consistent — instead of the free-form local justify below. Only
    // fires for tagged blocks; untagged blocks justify exactly as before.
    try {
      var qname = await getQaseedaAtSelection();
      if (qname && loadProfileStore()[qname]) { await applyProfileToQaseeda(qname); return; }
    } catch (e) { /* fall through to normal justify */ }

    // Fallback font from the pane — used only when a cell reports no explicit font.
    var fbMode = opts.fontMode === "nastaliq" ? "noto" : opts.fontMode;
    var fallbackName = AshaarFonts.wordNameOf(fbMode) || "Times New Roman";
    var doKashida = opts.justifyMode === "kashida" || opts.justifyMode === "spacing";
    var CELL_MARGIN_PT = 5.76; // Word default cell side margin (0.08") reserved for text
    var debug = !!(debugMode && debugMode.checked);
    var diags = [];

    setMessage("Justifying…");

    // Whitespace-mechanism fonts (Gulzar, Noto, …) have no elongatable joins —
    // injected tatweels shatter their shaping. Downgrade to spacing and warn
    // instead of running the tatweel path against them.
    if (mechanism === "whitespace" && opts.justifyMode === "kashida") {
      opts = Object.assign({}, opts, { justifyMode: "spacing" });
      setMessage("“" + (AshaarFonts.get(fontId) || {}).label + "” has no stretch letters — filling by spacing instead.");
    }

    await withWord(async function (context) {
      var selection = context.document.getSelection();

      // Find enclosing Ashaar Poem content control (the poem is the calibration unit)
      var cc = selection.parentContentControlOrNullObject;
      cc.load("title");
      await context.sync();

      var workRange = (!cc.isNullObject && cc.title === "Ashaar Poem")
        ? cc.getRange() : selection;

      var tables = workRange.tables;
      tables.load("items");
      await context.sync();

      if (!tables.items.length) {
        // No tables — justify plain selection text, measuring with the selection's own font.
        selection.load("text");
        selection.font.load("name,size");
        await context.sync();
        if (doKashida) {
          var pc = document.createElement("canvas").getContext("2d");
          if (pc) {
            pc.font = (selection.font.size || 16) + "pt \"" + (selection.font.name || fallbackName) + "\"";
            opts._justifyCtx = pc;
          }
        }
        var justifiedText = AshaarWord.justifyPlainTextBlock(stripJustification(selection.text), opts);
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
      tables.items.forEach(function (tbl) {
        tbl.rows.items.forEach(function (row) {
          row.cells.items.forEach(function (cell) {
            allCells.push(cell);
            cell.body.load("text");
            cell.body.font.load("name,size");
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

      // Representative font taken from the cells themselves (fall back to the pane).
      var repName = fallbackName, repSize = 16;
      for (var ci = 0; ci < allCells.length; ci++) {
        var rf = allCells[ci].body.font;
        if (rf && rf.name) { repName = rf.name; if (rf.size) repSize = rf.size; break; }
      }

      // Content width = cell width minus the side margins Word reserves for text.
      function contentPx(cell) {
        return Math.max(1, (cell.columnWidth || 0) - 2 * CELL_MARGIN_PT) * 96 / 72;
      }

      // Build the measurement canvas with the REAL font + size.
      var canvasCtx = null;
      if (doKashida) {
        var c = document.createElement("canvas").getContext("2d");
        if (c) { c.font = repSize + "pt \"" + repName + "\""; canvasCtx = c; opts._justifyCtx = c; }
      }

      // Ensure a bundled @font-face (e.g. FatemiMaqala) finishes loading before we
      // measure, so the canvas measures the same outlines Word renders. @font-face
      // fonts load lazily on first use; this forces the load and awaits it.
      if (canvasCtx && typeof document !== "undefined" && document.fonts && document.fonts.load) {
        try { await document.fonts.load(repSize + "pt \"" + repName + "\""); } catch (e) {}
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
      }

      // Probe + calibrate using the real font/size and content widths.
      var fontProfile = null;
      if (canvasCtx && typeof AshaarTune !== "undefined") {
        try { fontProfile = await AshaarTune.probeFont({ fontFamily: repName, fontSize: 64 }); }
        catch (e) { /* degrade gracefully */ }
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
          try {
            var session = await AshaarTune.calibrate({
              texts: lineTexts, fontFamily: repName, fontSize: repSize,
              containerWidth: avgPx, mode: "poetry", fontProfile: fontProfile, iterations: 50
            });
            calibParams = Object.assign({}, session.params);
            if (fontProfile) calibParams.fontQualityBoost = calibParams.fontQualityBoost || 1.8;
          } catch (e) { /* keep defaults */ }
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

      // Phase 1 (pure, no sync): rebuild each cell as an ordered list of style
      // runs and justify measuring each run in its OWN font. Produces per-cell
      // write plans consumed in phase 2.
      var plans = [];
      allCells.forEach(function (cell) {
        var current = (cell.body.text || "").trim();
        if (!stripJustification(current)) return;
        var colPx = contentPx(cell);

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
          if (flat !== current) plans.push({ cell: cell, flat: flat });
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

        // Mehr's tatweel whitelist: block the engine from inserting a tatweel
        // into any join the font can't actually render (medialInto only —
        // finalInto's narrower س/ش rule is a rendering-time check, not
        // table-expressible here). Rebuilt per cell from that cell's own text.
        if (mechanism === "tatweel") {
          var rules = AshaarFonts.tatweelRulesOf(fontId);
          if (rules) {
            var corpus = runs.map(function (r) { return r.text; }).join(" ");
            calibParams = Object.assign({}, calibParams, {
              priorityTable: AshaarTatweel.buildPriorityTable(corpus, rules)
            });
          }
        }

        if (opts.justifyMode === "kashida") {
          outTexts = AshaarJustify.justifyRuns(primRuns, colPx, calibParams).map(function (o) { return o.text; });
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
            target: Math.round(colPx * (calibParams.targetFill || 1)),
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
        if (!alignedOk) { plans.push({ cell: cell, flat: outTexts.join(" ") }); return; }

        plans.push({ cell: cell, runs: runs, outTexts: outTexts, sp: sp });
      });

      // Phase 2 (write): one context.sync() per cell so a range failure on one
      // cell falls back to a flattened whole-cell replace without aborting the
      // batch (the run-aware write can only error at sync, not synchronously).
      var changed = 0;
      for (var pi = 0; pi < plans.length; pi++) {
        var p = plans[pi];
        if (p.flat != null) {
          p.cell.body.paragraphs.getFirst().insertText(p.flat, Word.InsertLocation.replace);
          await context.sync();
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
          if (cellChanged) { await context.sync(); changed++; }
        } catch (e) {
          // Queued range write failed at sync (or count mismatch) — flatten.
          p.cell.body.paragraphs.getFirst().insertText(p.outTexts.join(" "), Word.InsertLocation.replace);
          await context.sync();
          changed++;
          if (debug) diags.push({ i: diags.length, font: "RANGE-FALLBACK", text: (e && e.message || "").slice(0, 14) });
        }
      }

      setMessage("Justified " + changed + " cell(s) across " + tables.items.length + " table(s).");
      if (debug) renderDebug(diags);
    });
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
      var section = context.document.sections.getFirst();
      section.load("pageLayout/width,pageLayout/leftMargin,pageLayout/rightMargin");
      await context.sync();
      var pl = section.pageLayout;
      var textWidthTwips = pl && pl.width
        ? Math.round((pl.width - (pl.leftMargin || 0) - (pl.rightMargin || 0)) * 20)
        : 9360;
      var ooxml = AshaarWord.wrapOoxml(AshaarWord.generateBareGrid12Ooxml(scaledTextWidth(textWidthTwips)));
      var selection = context.document.getSelection();
      var inserted = selection.insertOoxml(ooxml, Word.InsertLocation.end);
      var control = inserted.insertContentControl();
      control.title = "Ashaar Poem";
      control.tag = AshaarWord.contentControlTag("grid12", options());
      control.appearance = "BoundingBox";
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
      var template = {
        id: id,
        name: name,
        columnCount: GRID,
        rows: rows,
        fontMode: fontMode.value,
        justifyMode: justifyMode.value,
        tatweelCount: Number(tatweelCount.value || 0),
        gapWidth: Number(gapWidth.value || 4)
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
      var ooxml = AshaarWord.wrapOoxml(AshaarWord.templateToOoxml(tmpl, scaledTextWidth(textWidthTwips), opts));
      var selection = context.document.getSelection();
      var inserted = selection.insertOoxml(ooxml, Word.InsertLocation.end);
      var control = inserted.insertContentControl();
      control.title = "Ashaar Poem";
      control.tag = AshaarWord.contentControlTag("template:" + tmpl.name, opts);
      control.appearance = "BoundingBox";
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

  var isBound = false;

  function bind() {
    if (isBound) return;
    isBound = true;
    (function populateFontModes() {
      var order = ["document", "arabic-serif", "noto", "mehr", "jameel", "gulzar"];
      fontMode.innerHTML = "";
      order.forEach(function (id) {
        var d = AshaarFonts.get(id);
        if (!d) return;
        var o = document.createElement("option");
        o.value = id; o.textContent = d.label;
        if (id === "document") o.selected = true;
        fontMode.appendChild(o);
      });
    })();
    autoFitWidth.addEventListener("change", function () {
      tableWidth.disabled = autoFitWidth.checked;
      renderPreview();
    });
    [input, justifyMode, layoutMode, widthMode, bandhCount, misraCount, layoutPreset, layoutSpec, fontMode, tatweelCount, gapWidth, tableWidth, autoFitWidth].forEach(function (el) {
      el.addEventListener("input", renderPreview);
      el.addEventListener("change", renderPreview);
    });
    fontMode.addEventListener("change", updateFontNote);
    layoutPreset.addEventListener("change", applyLayoutPreset);
    misraCount.addEventListener("change", applyLayoutPreset);
    modeTable.addEventListener("click", function () { setMode("table"); });
    modeConvert.addEventListener("click", function () { setMode("convert"); });
    document.getElementById("insert-structure").addEventListener("click", insertStructure);
    document.getElementById("insert-poem").addEventListener("click", function () { insertPoem(false); });
    document.getElementById("insert-tabstop").addEventListener("click", insertTabStopPoem);
    document.getElementById("replace-selection").addEventListener("click", function () { insertPoem(true); });
    document.getElementById("justify-selection").addEventListener("click", justifySelection);
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
    // Qaseeda profile panel
    qaseedaStrength.addEventListener("input", function () { qaseedaStrengthValue.textContent = qaseedaStrength.value; });
    qaseedaName.addEventListener("change", loadQaseedaIntoPanel);
    document.getElementById("qaseeda-assign").addEventListener("click", assignBlockToQaseeda);
    document.getElementById("qaseeda-apply").addEventListener("click", saveAndApplyQaseeda);
    document.getElementById("qaseeda-font-check").addEventListener("click", checkQaseedaFont);

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
    updateFontNote();
    setMode("table");
    renderTemplateList();
    populateQaseedaNames();
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

/* StylesPane — wires the "Styles" tab to AshaarStyles (word-styles.js).
 * Creates/reconfigures 6 named Word styles from the active style group's
 * recipe, applies them to the current selection, and runs the RTL document
 * setup action. See docs/superpowers/specs/2026-07-16-ashaar-styles-design.md.
 */
(function () {
  "use strict";

  var els = {};
  var bound = false;
  var activeGroupName = "General";
  var groupStore = {}; // name -> group recipe (built-ins + custom, merged at load time)

  var GROUP_STORE_KEY = "ashaar-style-groups"; // custom groups only, keyed by name
  var ACTIVE_GROUP_KEY = "ashaar-style-active-group";

  function byId(id) { return document.getElementById(id); }

  function loadGroupStore() {
    try {
      var raw = Office.context.document.settings.get(GROUP_STORE_KEY);
      return raw ? JSON.parse(raw) : {};
    } catch (e) { return {}; }
  }

  function saveGroupStore(store, done) {
    Office.context.document.settings.set(GROUP_STORE_KEY, JSON.stringify(store || {}));
    Office.context.document.settings.saveAsync(function () { if (done) done(); });
  }

  function loadActiveGroupName() {
    var raw = Office.context.document.settings.get(ACTIVE_GROUP_KEY);
    return (typeof raw === "string" && raw) ? raw : "General";
  }

  function saveActiveGroupName(name, done) {
    Office.context.document.settings.set(ACTIVE_GROUP_KEY, name);
    Office.context.document.settings.saveAsync(function () { if (done) done(); });
  }

  // All group names available in the picker: built-ins first, then custom
  // (custom groups can shadow a built-in name; custom wins).
  function allGroups() {
    var out = {};
    Object.keys(AshaarStyles.BUILTIN_GROUPS).forEach(function (k) { out[k] = AshaarStyles.BUILTIN_GROUPS[k]; });
    Object.keys(groupStore).forEach(function (k) { out[k] = AshaarStyles.normalizeGroup(groupStore[k]); });
    return out;
  }

  function activeGroup() {
    var groups = allGroups();
    return groups[activeGroupName] || AshaarStyles.BUILTIN_GROUPS.General;
  }

  function populateGroupPicker() {
    var select = byId("styles-group-select");
    if (!select) return;
    select.innerHTML = "";
    var groups = allGroups();
    Object.keys(groups).sort().forEach(function (name) {
      var opt = document.createElement("option");
      opt.value = name;
      opt.textContent = name;
      if (name === activeGroupName) opt.selected = true;
      select.appendChild(opt);
    });
  }

  // Re-syncs Word's actual style definitions to match the active group, then
  // re-populates form fields from it (Task 7/8 read these same fields).
  function applyActiveGroupToDocument(then) {
    Word.run(function (context) {
      return AshaarStylesPane.ensureAshaarStyles(context, activeGroup()).then(function () {
        return context.sync();
      });
    }).then(function () {
      if (then) then();
    }).catch(function (e) {
      setStatus(byId("styles-status"), "Error applying style group: " + (e.message || String(e)), true);
    });
  }

  function setStatus(el, msg, warn) {
    if (!el) return;
    el.textContent = msg;
    el.classList.toggle("warn", !!warn);
  }

  // WordApiDesktop 1.3 covers everything this tab needs: Style.borders
  // (WordApiDesktop 1.1) and Font.nameAscii/nameBidirectional/sizeBidirectional
  // + PageSetup.sectionDirection (WordApiDesktop 1.3). Word for the web does
  // not support this requirement set at all.
  function isDesktopCapable() {
    return !!(window.Office && Office.context && Office.context.requirements &&
      Office.context.requirements.isSetSupported("WordApiDesktop", "1.3"));
  }

  // Idempotent: reuse the style if it already exists, else create it.
  function ensureStyle(context, name, type) {
    var style = context.document.getStyles().getByNameOrNullObject(name);
    style.load("isNullObject");
    return context.sync().then(function () {
      if (style.isNullObject) {
        return context.document.addStyle(name, type);
      }
      return style;
    });
  }

  // Configure one role's style object from the group recipe. Split into its
  // own function (rather than inlined in ensureAshaarStyles) so Task 8/9 can
  // reuse the font-string application logic if needed.
  function configureRoleStyle(style, role, group) {
    var recipe = group[role];
    style.baseStyle = AshaarStyles.BASE_STYLE[role];
    style.unhideWhenUsed = true; // visible in Word's own Style gallery once used

    if (role === "heading1" || role === "heading2" || role === "heading3") {
      style.font.nameAscii = recipe.font;
      style.font.nameBidirectional = recipe.font;
      style.font.size = recipe.sizePt;
      style.font.sizeBidirectional = recipe.sizePt;
      style.font.bold = true;
      style.paragraphFormat.alignment = Word.Alignment.centered;
    } else if (role === "emphasis") {
      style.font.color = recipe.color;
      // Word's built-in "Emphasis" (our basedOn) is ITALIC, and italic does not
      // render in Arabic-script fonts — cancel it explicitly. Emphasis's signal
      // is red + a live per-instance size bump (applyEmphasis), not slant.
      style.font.italic = false;
      style.font.italicBidirectional = false;
      // No absolute size here — the bump is computed live per instance
      // (computeEmphasisSize) and written to sizeBidirectional in applyEmphasis.
    } else if (role === "quote") {
      style.paragraphFormat.leftIndent = AshaarStyles.clampIndentPt(recipe.indentPt);
      style.paragraphFormat.rightIndent = AshaarStyles.clampIndentPt(recipe.indentPt);
      var leftBorder = style.borders.getByLocation(Word.BorderLocation.left);
      var rightBorder = style.borders.getByLocation(Word.BorderLocation.right);
      leftBorder.type = Word.BorderType.single;
      leftBorder.width = recipe.borderWidth;
      leftBorder.color = recipe.borderColor;
      rightBorder.type = Word.BorderType.single;
      rightBorder.width = recipe.borderWidth;
      rightBorder.color = recipe.borderColor;
    } else if (role === "quranQuote") {
      style.font.nameAscii = recipe.font;
      style.font.nameBidirectional = recipe.font;
      // lineSpacing has no clean "reset to Word auto" value in the object
      // model, so switching a group FROM a set line height back to null
      // (auto) will leave the previous numeric value in place rather than
      // truly reverting to auto. Flagged for the manual check in Task 9 —
      // if this matters in practice, the fix is to also set
      // style.paragraphFormat.lineSpacingRule = Word.LineSpacing.single
      // as the "auto" case, which needs confirming live before relying on it.
      if (recipe.lineHeightPt != null) {
        style.paragraphFormat.lineSpacing = AshaarStyles.clampLineHeightPt(recipe.lineHeightPt);
      }
    }
  }

  // Create/reconfigure all 6 named styles from `group`'s recipe. ROLES order
  // (quote before quranQuote) guarantees Ashaar Quote exists before Ashaar
  // Quran Quote's baseStyle references it.
  function ensureAshaarStyles(context, group) {
    var styleObjs = {};
    var chain = Promise.resolve();
    AshaarStyles.ROLES.forEach(function (role) {
      chain = chain.then(function () {
        return ensureStyle(context, AshaarStyles.STYLE_NAME[role], AshaarStyles.STYLE_TYPE[role]);
      }).then(function (style) {
        style.load("baseStyle"); // harmless preload; configureRoleStyle overwrites it
        styleObjs[role] = style;
        return context.sync();
      }).then(function () {
        configureRoleStyle(styleObjs[role], role, group);
        return context.sync();
      });
    });
    return chain.then(function () { return styleObjs; });
  }

  // Reads current field values into a group recipe (Task 7-8).
  function readFieldsIntoGroup(name) {
    return {
      name: name,
      heading1: { font: byId("styles-h1-font").value, sizePt: Number(byId("styles-h1-size").value) },
      heading2: { font: byId("styles-h2-font").value, sizePt: Number(byId("styles-h2-size").value) },
      heading3: { font: byId("styles-h3-font").value, sizePt: Number(byId("styles-h3-size").value) },
      emphasis: { color: byId("styles-emphasis-color").value, bumpPt: Number(byId("styles-emphasis-bump").value) },
      quote: {
        borderColor: byId("styles-quote-color").value,
        borderWidth: byId("styles-quote-width").value,
        indentPt: Number(byId("styles-quote-indent").value)
      },
      quranQuote: {
        font: byId("styles-quranquote-font").value,
        lineHeightPt: byId("styles-quranquote-lh").value === "" ? null : Number(byId("styles-quranquote-lh").value)
      }
    };
  }

  // Populates fields from a group recipe (Task 7-8).
  function populateFieldsFromGroup(group) {
    byId("styles-h1-font").value = group.heading1.font;
    byId("styles-h1-size").value = group.heading1.sizePt;
    byId("styles-h2-font").value = group.heading2.font;
    byId("styles-h2-size").value = group.heading2.sizePt;
    byId("styles-h3-font").value = group.heading3.font;
    byId("styles-h3-size").value = group.heading3.sizePt;
    byId("styles-emphasis-color").value = group.emphasis.color;
    byId("styles-emphasis-bump").value = group.emphasis.bumpPt;
    byId("styles-quote-color").value = group.quote.borderColor;
    byId("styles-quote-width").value = group.quote.borderWidth;
    byId("styles-quote-indent").value = group.quote.indentPt;
    byId("styles-quranquote-font").value = group.quranQuote.font;
    byId("styles-quranquote-lh").value = group.quranQuote.lineHeightPt == null ? "" : group.quranQuote.lineHeightPt;
  }

  // Task 8: "Update style" handler — writes current field values into the
  // active group and persists it, shadowing built-ins on first edit.
  function updateActiveGroupFromFields() {
    var updated = readFieldsIntoGroup(activeGroupName);
    groupStore[activeGroupName] = updated;
    saveGroupStore(groupStore, function () {
      applyActiveGroupToDocument();
    });
  }

  function bindUpdateButtons() {
    ["styles-h1-update", "styles-h2-update", "styles-h3-update",
      "styles-emphasis-update", "styles-quote-update", "styles-quranquote-update"
    ].forEach(function (id) {
      byId(id).addEventListener("click", updateActiveGroupFromFields);
    });
  }

  // Task 8: Apply paragraph styles (headings, quote, quranQuote) to selection.
  function applyParagraphStyle(styleName) {
    Word.run(function (context) {
      var selection = context.document.getSelection();
      var paragraphs = selection.paragraphs;
      paragraphs.load("items");
      return context.sync().then(function () {
        paragraphs.items.forEach(function (p) { p.style = styleName; });
        return context.sync();
      });
    }).catch(function (e) {
      setStatus(byId("styles-status"), "Error applying style: " + (e.message || String(e)), true);
    });
  }

  function bindParagraphApplyButtons() {
    byId("styles-h1-apply").addEventListener("click", function () { applyParagraphStyle(AshaarStyles.STYLE_NAME.heading1); });
    byId("styles-h2-apply").addEventListener("click", function () { applyParagraphStyle(AshaarStyles.STYLE_NAME.heading2); });
    byId("styles-h3-apply").addEventListener("click", function () { applyParagraphStyle(AshaarStyles.STYLE_NAME.heading3); });
    byId("styles-quote-apply").addEventListener("click", function () { applyParagraphStyle(AshaarStyles.STYLE_NAME.quote); });
    byId("styles-quranquote-apply").addEventListener("click", function () { applyParagraphStyle(AshaarStyles.STYLE_NAME.quranQuote); });
  }

  // Task 8: Apply emphasis (character style + live size bump).
  function applyEmphasis() {
    var bumpPt = Number(byId("styles-emphasis-bump").value) || 0;
    var color = byId("styles-emphasis-color").value;
    Word.run(function (context) {
      var selection = context.document.getSelection();
      selection.font.load("size, sizeBidirectional");
      return context.sync().then(function () {
        // Arabic-script text renders at the complex-script size (szCs), so base
        // the bump on sizeBidirectional and write it back there — writing only
        // .size (the Latin size) leaves Arabic text visually unchanged. Keep
        // .size in sync too for any Latin runs in the selection.
        var base = selection.font.sizeBidirectional || selection.font.size;
        var resultSize = AshaarStyles.computeEmphasisSize(base, bumpPt);
        selection.style = AshaarStyles.STYLE_NAME.emphasis;
        selection.font.color = color;
        selection.font.size = resultSize;
        selection.font.sizeBidirectional = resultSize;
        return context.sync();
      });
    }).catch(function (e) {
      setStatus(byId("styles-status"), "Error applying Emphasis: " + (e.message || String(e)), true);
    });
  }

  // Task 8: Instance-level overrides.
  function applyQuoteIndentOverride() {
    var raw = byId("styles-quote-indent-override").value;
    if (raw === "") return; // blank = no override requested
    var pt = AshaarStyles.clampIndentPt(Number(raw));
    Word.run(function (context) {
      var paragraphs = context.document.getSelection().paragraphs;
      paragraphs.load("items");
      return context.sync().then(function () {
        paragraphs.items.forEach(function (p) { p.leftIndent = pt; p.rightIndent = pt; });
        return context.sync();
      });
    }).catch(function (e) {
      setStatus(byId("styles-status"), "Error applying indent override: " + (e.message || String(e)), true);
    });
  }

  function applyQuranQuoteLineHeightOverride() {
    var raw = byId("styles-quranquote-lh-override").value;
    if (raw === "") return;
    var pt = AshaarStyles.clampLineHeightPt(Number(raw));
    Word.run(function (context) {
      var paragraphs = context.document.getSelection().paragraphs;
      paragraphs.load("items");
      return context.sync().then(function () {
        paragraphs.items.forEach(function (p) { p.lineSpacing = pt; });
        return context.sync();
      });
    }).catch(function (e) {
      setStatus(byId("styles-status"), "Error applying line-height override: " + (e.message || String(e)), true);
    });
  }

  // Task 9 (+ manual-test refinements): RTL document setup. Creates a named
  // "Ashaar Normal" body style (rather than mutating the built-in Normal —
  // keeps built-ins pristine, spec decision #1), right-aligns the built-in
  // "Footnote Text" style so footnotes sit on the right, and flips the section
  // to right-to-left layout. Office.js has NO paragraph reading-order/bidi
  // setter (VBA-only), so right-alignment is the closest RTL lever for body +
  // footnote paragraphs; and the footnote SEPARATOR line (a special id=-1
  // footnote) is not exposed by the footnote API at all — it can only be
  // right-aligned via Word's Draft view (Show Notes → Footnote Separator) or a
  // prebuilt template.
  var ASHAAR_NORMAL_NAME = "Ashaar Normal";
  function runRtlSetup() {
    var latinFont = byId("styles-rtl-latin-font").value;
    var csFont = byId("styles-rtl-cs-font").value;
    var csSize = Number(byId("styles-rtl-cs-size").value) || 12;
    setStatus(byId("styles-status"), "Applying…");
    Word.run(function (context) {
      var styles = context.document.getStyles();
      var ashaarNormal = styles.getByNameOrNullObject(ASHAAR_NORMAL_NAME);
      // Built-in footnote text style; present once the document has a footnote.
      var footnote = styles.getByNameOrNullObject("Footnote Text");
      var section = context.document.sections.getFirst();
      ashaarNormal.load("isNullObject");
      footnote.load("isNullObject");
      return context.sync().then(function () {
        var an = ashaarNormal.isNullObject
          ? context.document.addStyle(ASHAAR_NORMAL_NAME, "Paragraph")
          : ashaarNormal;
        return context.sync().then(function () {
          an.baseStyle = "Normal";
          an.unhideWhenUsed = true;
          an.font.nameAscii = latinFont;
          an.font.nameBidirectional = csFont;
          an.font.size = csSize;
          an.font.sizeBidirectional = csSize;
          an.paragraphFormat.alignment = Word.Alignment.right;
          if (!footnote.isNullObject) {
            footnote.font.nameBidirectional = csFont;
            footnote.font.sizeBidirectional = csSize;
            footnote.paragraphFormat.alignment = Word.Alignment.right;
          }
          section.pageSetup.sectionDirection = Word.SectionDirection.rightToLeft;
          return context.sync();
        });
      });
    }).then(function () {
      setStatus(byId("styles-status"),
        "Applied: created “Ashaar Normal” body style (right-aligned, complex-script font/size), right-aligned footnote text, and right-to-left section layout. Apply “Ashaar Normal” to body paragraphs from Word’s Styles gallery. The add-in API can’t set true paragraph reading order or move the footnote separator line — for those use Word’s Layout → Paragraph Direction and Draft view → Show Notes → Footnote Separator.");
    }).catch(function (e) {
      setStatus(byId("styles-status"), "Error: " + (e.message || String(e)), true);
    });
  }

  // Public: called when the Styles tab is first shown (Task 4's onTabShown
  // hook) and again whenever the group picker changes (Task 6).
  function onTabShown() {
    if (!els.body) cacheEls();
    if (!isDesktopCapable()) {
      els.unsupported.hidden = false;
      els.body.hidden = true;
      return;
    }
    els.unsupported.hidden = true;
    els.body.hidden = false;
    if (bound) return;
    bound = true;
    groupStore = loadGroupStore();
    activeGroupName = loadActiveGroupName();
    populateGroupPicker();
    byId("styles-group-saveas").addEventListener("click", function () {
      byId("styles-saveas-row").hidden = false;
      byId("styles-saveas-name").value = "";
      byId("styles-saveas-name").focus();
    });
    byId("styles-saveas-cancel").addEventListener("click", function () {
      byId("styles-saveas-row").hidden = true;
    });
    byId("styles-saveas-ok").addEventListener("click", function () {
      var name = String(byId("styles-saveas-name").value || "").trim();
      if (!name) return;
      groupStore[name] = readFieldsIntoGroup(name);
      saveGroupStore(groupStore, function () {
        byId("styles-saveas-row").hidden = true;
        activeGroupName = name;
        saveActiveGroupName(activeGroupName);
        populateGroupPicker();
      });
    });
    byId("styles-group-select").addEventListener("change", function (e) {
      activeGroupName = e.target.value;
      saveActiveGroupName(activeGroupName);
      populateFieldsFromGroup(activeGroup());
      applyActiveGroupToDocument();
    });
    // Task 8 bindings: Update, paragraph apply, emphasis apply, override applies
    bindUpdateButtons();
    bindParagraphApplyButtons();
    byId("styles-emphasis-apply").addEventListener("click", applyEmphasis);
    byId("styles-quote-override-apply").addEventListener("click", applyQuoteIndentOverride);
    byId("styles-quranquote-override-apply").addEventListener("click", applyQuranQuoteLineHeightOverride);
    // Task 9 binding: RTL document setup
    byId("styles-rtl-apply").addEventListener("click", runRtlSetup);
    applyActiveGroupToDocument();
    populateFieldsFromGroup(activeGroup());
  }

  function cacheEls() {
    els.unsupported = byId("styles-unsupported");
    els.body = byId("styles-body");
  }

  window.AshaarStylesPane = {
    onTabShown: onTabShown,
    isDesktopCapable: isDesktopCapable,
    ensureAshaarStyles: ensureAshaarStyles // exposed for Tasks 6-9 in this same file
  };
}());

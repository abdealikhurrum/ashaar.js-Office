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

  function byId(id) { return document.getElementById(id); }

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
      // No absolute size here — Emphasis's bump is computed live per
      // instance in Task 8 (computeEmphasisSize), not stored as a style size.
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

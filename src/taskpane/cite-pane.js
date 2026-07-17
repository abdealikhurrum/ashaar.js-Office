/* CitePane — wires the "Cite" tab to CiteEngine (cite-engine.js) + CiteWord
 * (cite-word.js). Loads a bundled sample library (fixtures/cite-sample.json)
 * plus the vendored CSL style/locale XML, lets the user pick items + style +
 * locale + output form, renders a live formatted preview (pure JS — works in a
 * bare browser without Word), and inserts the formatted citation as a
 * footnote/endnote/inline note and the bibliography as a tagged content control.
 *
 * RTL note: Office.js exposes NO paragraph reading-order (bidi) setter — only
 * alignment. For Arabic locales we right-align the inserted note/bibliography
 * paragraphs as the practical RTL treatment; true bidi ordering must be set in
 * Word (Layout → Paragraph Direction). See the manual checklist.
 *
 * See docs/superpowers/specs/2026-07-16-citation-engine-manual-checklist.md.
 */
(function () {
  "use strict";

  var bound = false;
  var itemsPopulated = false;

  // In-memory caches so keystroke-fast preview re-renders don't re-fetch.
  var cache = { styles: {}, locales: null, items: null };

  function byId(id) { return document.getElementById(id); }

  function setStatus(msg, warn) {
    var el = byId("cite-status");
    if (!el) { return; }
    el.textContent = msg || "";
    el.classList.toggle("warn", !!warn);
  }

  function isRtlLang(lang) { return /^ar\b/i.test(lang || ""); }

  function fetchText(url) {
    return fetch(url).then(function (r) {
      if (!r.ok) { throw new Error(url + " → HTTP " + r.status); }
      return r.text();
    });
  }

  // Load (once, then cached) the locale XML pair, the sample library, and the
  // selected style's CSL XML.
  function ensureAssets(styleFile) {
    var jobs = [];
    if (!cache.locales) {
      jobs.push(Promise.all([
        fetchText("../vendor/csl-locales/locales-en-US.xml"),
        fetchText("../vendor/csl-locales/locales-ar.xml")
      ]).then(function (pair) {
        cache.locales = { "us": pair[0], "en-US": pair[0], "ar": pair[1] };
      }));
    }
    if (!cache.items) {
      jobs.push(fetchText("fixtures/cite-sample.json").then(function (txt) {
        cache.items = JSON.parse(txt);
      }));
    }
    if (!cache.styles[styleFile]) {
      jobs.push(fetchText("../vendor/csl-styles/" + styleFile + ".csl").then(function (xml) {
        cache.styles[styleFile] = xml;
      }));
    }
    return Promise.all(jobs);
  }

  function currentStyleFile() { return (byId("cite-style") || {}).value || "chicago-notes-bibliography"; }
  function currentLang() { return (byId("cite-locale") || {}).value || "en-US"; }
  function currentForm() { return (byId("cite-form") || {}).value || "footnote"; }

  // Build a fresh engine for this style+locale. Cheap enough per render, and
  // avoids citeproc citation-registry state leaking across selections.
  function buildEngine(styleFile, lang) {
    return CiteEngine.build({
      styleXml: cache.styles[styleFile],
      locales: cache.locales,
      items: cache.items,
      lang: lang
    });
  }

  function populateItems(skipSeed) {
    if (itemsPopulated || !cache.items) { return; }
    var host = byId("cite-items");
    if (!host) { return; }
    host.innerHTML = "";
    Object.keys(cache.items).forEach(function (id, idx) {
      var item = cache.items[id] || {};
      var li = document.createElement("li");
      var cb = document.createElement("input");
      cb.type = "checkbox";
      cb.id = "cite-item-" + id;
      cb.setAttribute("data-cite-id", id);
      if (idx === 0 && !skipSeed) { cb.checked = true; } // seed a non-empty preview
      cb.addEventListener("change", renderPreview);
      var lbl = document.createElement("label");
      lbl.setAttribute("for", cb.id);
      var title = item.title || id;
      lbl.innerHTML = "";
      lbl.appendChild(document.createTextNode(title + " "));
      var idSpan = document.createElement("span");
      idSpan.className = "cite-item-id";
      idSpan.textContent = "(" + id + ")";
      lbl.appendChild(idSpan);
      li.appendChild(cb);
      li.appendChild(lbl);

      // remove (x) button
      var rm = document.createElement("button");
      rm.type = "button";
      rm.className = "cite-item-remove";
      rm.setAttribute("aria-label", "Remove " + title);
      rm.textContent = "×";
      rm.addEventListener("click", function () { removeItem(id); });
      li.appendChild(rm);

      // locator row (type + value), shown only when the item is checked
      var loc = document.createElement("div");
      loc.className = "cite-locator-row";
      loc.hidden = !cb.checked;
      var prefix = document.createElement("span");
      prefix.className = "cite-locator-prefix";
      prefix.textContent = "cite at:";
      var sel = document.createElement("select");
      sel.className = "cite-locator-type";
      sel.setAttribute("data-cite-loc-type", id);
      ["page", "chapter", "section", "verse"].forEach(function (t) {
        var opt = document.createElement("option");
        opt.value = t; opt.textContent = t;
        sel.appendChild(opt);
      });
      var val = document.createElement("input");
      val.type = "text";
      val.className = "cite-locator-value";
      val.setAttribute("data-cite-loc-value", id);
      val.placeholder = "e.g. 42";
      val.addEventListener("input", renderPreview);
      sel.addEventListener("change", renderPreview);
      loc.appendChild(prefix); loc.appendChild(sel); loc.appendChild(val);
      li.appendChild(loc);

      // toggle the locator row + refresh preview when checked state changes
      // (in addition to the renderPreview binding above — both handlers fire)
      cb.addEventListener("change", function () { loc.hidden = !cb.checked; });

      host.appendChild(li);
    });
    itemsPopulated = true;
  }

  function selectedIds() {
    var ids = [];
    var boxes = document.querySelectorAll("#cite-items input[data-cite-id]");
    for (var i = 0; i < boxes.length; i++) {
      if (boxes[i].checked) { ids.push(boxes[i].getAttribute("data-cite-id")); }
    }
    return ids;
  }

  function removeItem(id) {
    if (cache.items && cache.items[id]) { delete cache.items[id]; }
    itemsPopulated = false;
    populateItems(true);
    renderPreview();
  }

  // Checked items + their locator inputs -> [{id, locator, label}] for cite().
  function selectedCitationItems() {
    var out = [];
    var boxes = document.querySelectorAll("#cite-items input[data-cite-id]");
    for (var i = 0; i < boxes.length; i++) {
      if (!boxes[i].checked) { continue; }
      var id = boxes[i].getAttribute("data-cite-id");
      var esc = (window.CSS && CSS.escape) ? CSS.escape(id) : id;
      var vEl = document.querySelector('[data-cite-loc-value="' + esc + '"]');
      var tEl = document.querySelector('[data-cite-loc-type="' + esc + '"]');
      var item = { id: id };
      if (vEl && vEl.value.trim()) { item.locator = vEl.value.trim(); item.label = (tEl && tEl.value) || "page"; }
      out.push(item);
    }
    return out;
  }

  function block(labelText, bodyHtml) {
    return '<div class="cite-preview-block"><div class="cite-preview-label">' + labelText +
      '</div><div class="cite-preview-body">' + bodyHtml + "</div></div>";
  }

  function renderPreview() {
    var styleFile = currentStyleFile();
    var lang = currentLang();
    var preview = byId("cite-preview");
    if (!preview) { return Promise.resolve(); }
    // Returned (not fire-and-forget) so callers — e.g. addFromZotero(),
    // pingZotero() — can sequence their own status message to land *after*
    // this render's internal setStatus("") / setStatus(error) settles,
    // instead of racing it.
    return ensureAssets(styleFile).then(function () {
      populateItems();
      var engine = buildEngine(styleFile, lang);
      var items = selectedCitationItems();
      var citeHtml = items.length
        ? CiteWord.wrapRtlRuns(CiteWord.sanitize(engine.cite(items)))
        : "<em>Select one or more items to preview a citation.</em>";
      var bibHtml = CiteWord.wrapRtlRuns(CiteWord.sanitize(engine.bibliography())) || "<em>No bibliography.</em>";
      var rtl = isRtlLang(lang);
      preview.setAttribute("dir", rtl ? "rtl" : "ltr");
      preview.innerHTML =
        block(currentForm() + " citation", citeHtml) +
        block("bibliography", bibHtml);
      setStatus("");
    }).catch(function (e) {
      setStatus("Couldn't load citation assets: " + (e && e.message ? e.message : String(e)), true);
    });
  }

  function notesSupported() {
    try {
      return !!(typeof Office !== "undefined" && Office.context && Office.context.requirements &&
        Office.context.requirements.isSetSupported("WordApi", "1.5"));
    } catch (e) { return false; }
  }

  function rightAlignParas(paras) {
    for (var i = 0; i < paras.items.length; i++) {
      paras.items[i].alignment = Word.Alignment.right;
    }
  }

  function insertCitation() {
    var styleFile = currentStyleFile();
    var lang = currentLang();
    var form = currentForm();
    var rtl = isRtlLang(lang);
    ensureAssets(styleFile).then(function () {
      var items = selectedCitationItems();
      if (!items.length) { setStatus("Select at least one item to cite.", true); return; }
      var engine = buildEngine(styleFile, lang);
      var html = CiteWord.wrapRtlRuns(CiteWord.sanitize(engine.cite(items)));
      var citeTag = CiteWord.buildCitationTag({ style: styleFile, locale: lang, items: items });
      if (typeof Word === "undefined" || !Word.run) {
        setStatus("Word isn't available — this is preview-only in a browser.", true);
        return;
      }
      var canNotes = notesSupported();
      var fellBack = false;
      Word.run(function (ctx) {
        var sel = ctx.document.getSelection();
        var range;
        if (form === "inline") {
          range = sel.insertHtml(html, Word.InsertLocation.replace);
        } else if (canNotes) {
          var note = form === "endnote" ? sel.insertEndnote() : sel.insertFootnote();
          // insertHtml(replace) returns the range of the newly inserted content —
          // use it (not the pre-insert note range) so alignment hits real paragraphs.
          range = note.body.getRange().insertHtml(html, Word.InsertLocation.replace);
        } else {
          range = sel.insertHtml(html, Word.InsertLocation.replace);
          fellBack = true;
        }
        var cc = range.insertContentControl();
        cc.tag = citeTag;
        cc.title = "Ashaar Citation";
        var paras = range.paragraphs;
        paras.load("items");
        return ctx.sync().then(function () {
          if (rtl) { rightAlignParas(paras); }
          return ctx.sync();
        });
      }).then(function () {
        setStatus(fellBack
          ? "Word < 1.5: footnotes/endnotes unavailable — inserted inline instead."
          : "Inserted " + form + " citation.");
        // Clear locator inputs (not the checkboxes) so the same item can be
        // re-cited immediately with a different locator.
        var vals = document.querySelectorAll("#cite-items .cite-locator-value");
        for (var i = 0; i < vals.length; i++) { vals[i].value = ""; }
        renderPreview();
      }).catch(function (e) {
        setStatus("Insert failed: " + (e && e.message ? e.message : String(e)), true);
      });
    }).catch(function (e) {
      setStatus("Couldn't load citation assets: " + (e && e.message ? e.message : String(e)), true);
    });
  }

  function insertBibliography() {
    var styleFile = currentStyleFile();
    var lang = currentLang();
    var rtl = isRtlLang(lang);
    ensureAssets(styleFile).then(function () {
      var engine = buildEngine(styleFile, lang);
      var payload = CiteWord.buildBibliographyPayload({ html: engine.bibliography(), rtl: rtl });
      var bibTag = CiteWord.buildBibliographyTag({ style: styleFile, locale: lang });
      if (typeof Word === "undefined" || !Word.run) {
        setStatus("Word isn't available — this is preview-only in a browser.", true);
        return;
      }
      Word.run(function (ctx) {
        var range = ctx.document.getSelection().insertHtml(payload.html, Word.InsertLocation.after);
        var cc = range.insertContentControl();
        cc.tag = bibTag;
        cc.title = "Ashaar Bibliography";
        var paras = range.paragraphs;
        paras.load("items");
        return ctx.sync().then(function () {
          if (rtl) { rightAlignParas(paras); }
          return ctx.sync();
        });
      }).then(function () {
        setStatus("Inserted bibliography.");
      }).catch(function (e) {
        setStatus("Insert failed: " + (e && e.message ? e.message : String(e)), true);
      });
    }).catch(function (e) {
      setStatus("Couldn't load citation assets: " + (e && e.message ? e.message : String(e)), true);
    });
  }

  var ZOTERO_HINT = "Start Zotero (with Better BibTeX) to cite from your library.";

  // Availability ping, fired after the tab's own render settles (see
  // renderPreview()'s comment) so a warn status here can't be clobbered by —
  // and doesn't race — the fixture render's own status update. Guarded for
  // cite-zotero.js being absent so the pane still works without it.
  function pingZotero() {
    if (typeof CiteZotero === "undefined") { return; }
    CiteZotero.ping().then(function (ok) {
      if (!ok) { setStatus(ZOTERO_HINT, true); }
      // else: Zotero is reachable — do not disturb the existing preview/status.
    });
  }

  function addFromZotero() {
    if (typeof CiteZotero === "undefined") {
      setStatus(ZOTERO_HINT, true);
      return;
    }
    setStatus("Picking in Zotero…");
    // caywPick() resolves [{citekey, locator?, label?}] — fetchCslJson still
    // wants bare citekeys, so map those out; the objects themselves carry the
    // locator info used to pre-fill the per-item locator inputs below.
    CiteZotero.caywPick().then(function (picks) {
      if (!picks || !picks.length) {
        setStatus(""); // cancelled — no-op
        return;
      }
      var citekeys = picks.map(function (p) { return p.citekey; });
      return CiteZotero.fetchCslJson(citekeys).then(function (items) {
        if (!cache.items) { cache.items = {}; }
        // Capture already-checked boxes BEFORE the rebuild below wipes them —
        // populateItems(true) below skips the index-0 seed, so re-checking
        // relies entirely on previouslySelected + citekeys below.
        var previouslySelected = selectedIds();
        Object.keys(items).forEach(function (id) {
          cache.items[id] = items[id];
        });
        // Force a full re-render of the item list so the new entries appear,
        // then re-check the union of previously-checked items and the freshly
        // added Zotero citekeys, so Insert still cites everything the user had
        // selected.
        itemsPopulated = false;
        populateItems(true); // skipSeed: rely on previouslySelected+citekeys re-check below
        previouslySelected.concat(citekeys).forEach(function (id) {
          var cb = byId("cite-item-" + id);
          if (cb) { cb.checked = true; }
        });
        // Pre-fill the locator row for picks that carried a page/chapter/etc.
        // from the CAYW popup, and make sure that row is visible.
        picks.forEach(function (p) {
          var cb = byId("cite-item-" + p.citekey);
          if (!cb) { return; }
          var row = cb.parentNode ? cb.parentNode.querySelector(".cite-locator-row") : null;
          if (row) { row.hidden = false; }
          if (p.locator) {
            var esc = (window.CSS && CSS.escape) ? CSS.escape(p.citekey) : p.citekey;
            var vEl = document.querySelector('[data-cite-loc-value="' + esc + '"]');
            var tEl = document.querySelector('[data-cite-loc-type="' + esc + '"]');
            if (vEl) { vEl.value = p.locator; }
            if (tEl && p.label) { tEl.value = p.label; }
          }
        });
        // Sequence the "Added…" status after renderPreview()'s own
        // setStatus("") settles, so it isn't immediately overwritten.
        return renderPreview().then(function () {
          setStatus("Added " + citekeys.length + " item(s) from Zotero.");
        });
      });
    }).catch(function (e) {
      // Proxy 502 / JSON-RPC error, network failure, or a post-merge rebuild
      // error (populateItems()/renderPreview() above) all land here — the
      // generic hint covers all of them; log the real error for debugging.
      console.error("[CitePane] addFromZotero failed:", e);
      setStatus(ZOTERO_HINT, true);
    });
  }

  function bind() {
    if (bound) { return; }
    var styleSel = byId("cite-style");
    var insertBtn = byId("cite-insert");
    if (!styleSel || !insertBtn) { return; } // markup not present yet
    bound = true;

    styleSel.addEventListener("change", renderPreview);
    byId("cite-locale").addEventListener("change", renderPreview);
    byId("cite-form").addEventListener("change", renderPreview);
    insertBtn.addEventListener("click", insertCitation);
    byId("cite-insert-bib").addEventListener("click", insertBibliography);
    var addZoteroBtn = byId("cite-add-zotero");
    if (addZoteroBtn) { addZoteroBtn.addEventListener("click", addFromZotero); }

    renderPreview().then(pingZotero);
  }

  function onTabShown() {
    if (!bound) { bind(); return; }
    renderPreview().then(pingZotero);
  }

  window.CitePane = { bind: bind, onTabShown: onTabShown };
}());

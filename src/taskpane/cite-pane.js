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

  function populateItems() {
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
      if (idx === 0) { cb.checked = true; } // seed a non-empty preview
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

  function block(labelText, bodyHtml) {
    return '<div class="cite-preview-block"><div class="cite-preview-label">' + labelText +
      '</div><div class="cite-preview-body">' + bodyHtml + "</div></div>";
  }

  function renderPreview() {
    var styleFile = currentStyleFile();
    var lang = currentLang();
    var preview = byId("cite-preview");
    if (!preview) { return; }
    ensureAssets(styleFile).then(function () {
      populateItems();
      var engine = buildEngine(styleFile, lang);
      var ids = selectedIds();
      var citeHtml = ids.length
        ? CiteWord.sanitize(engine.cite(ids))
        : "<em>Select one or more items to preview a citation.</em>";
      var bibHtml = CiteWord.sanitize(engine.bibliography()) || "<em>No bibliography.</em>";
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
      var ids = selectedIds();
      if (!ids.length) { setStatus("Select at least one item to cite.", true); return; }
      var engine = buildEngine(styleFile, lang);
      var html = CiteWord.sanitize(engine.cite(ids));
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
      if (typeof Word === "undefined" || !Word.run) {
        setStatus("Word isn't available — this is preview-only in a browser.", true);
        return;
      }
      Word.run(function (ctx) {
        var range = ctx.document.getSelection().insertHtml(payload.html, Word.InsertLocation.after);
        var cc = range.insertContentControl();
        cc.tag = payload.tag;
        cc.title = "Ashaar Bibliography";
        var paras = range.paragraphs;
        paras.load("items");
        return ctx.sync().then(function () {
          if (rtl) { rightAlignParas(paras); }
          return ctx.sync();
        });
      }).then(function () {
        setStatus("Inserted bibliography (tagged \"" + payload.tag + "\").");
      }).catch(function (e) {
        setStatus("Insert failed: " + (e && e.message ? e.message : String(e)), true);
      });
    }).catch(function (e) {
      setStatus("Couldn't load citation assets: " + (e && e.message ? e.message : String(e)), true);
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

    renderPreview();
  }

  function onTabShown() {
    if (!bound) { bind(); return; }
    renderPreview();
  }

  window.CitePane = { bind: bind, onTabShown: onTabShown };
}());

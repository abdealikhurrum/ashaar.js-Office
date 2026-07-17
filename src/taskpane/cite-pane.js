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

  // Fire-and-forget persistence of the reference set into the document; a save
  // failure is a non-fatal hint (never blocks the UI). No-op without CiteStore/Office.
  function persistRefs() {
    if (typeof CiteStore === "undefined") { return; }
    CiteStore.saveRefs(cache.items).catch(function () {
      setStatus("Couldn't save your reference list to the document.", true);
    });
  }

  function isRtlLang(lang) { return /^ar\b/i.test(lang || ""); }

  var DEFAULT_AR_CS_FONT = "Arial"; // universally present, has Arabic coverage
  // Read the document's complex-script font inside a Word.run: prefer the
  // "Ashaar Normal" style (created by the Styles-tab RTL setup), else the target
  // range's bidi font, else the default. Returns a Promise<string>.
  function readDocCsFont(ctx, range) {
    var style = ctx.document.getStyles().getByNameOrNullObject("Ashaar Normal");
    style.load("isNullObject,font/nameBidirectional");
    range.font.load("nameBidirectional");
    return ctx.sync().then(function () {
      if (!style.isNullObject && style.font.nameBidirectional) { return style.font.nameBidirectional; }
      if (range.font.nameBidirectional) { return range.font.nameBidirectional; }
      return DEFAULT_AR_CS_FONT;
    });
  }

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
      var loadRefs = (typeof CiteStore !== "undefined")
        ? CiteStore.loadRefs()
        : Promise.resolve({});
      jobs.push(loadRefs.then(function (saved) {
        if (saved && Object.keys(saved).length) { cache.items = enrich(saved); return; }
        return fetchText("fixtures/cite-sample.json").then(function (txt) { cache.items = enrich(JSON.parse(txt)); });
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
  function currentVariant() { return (byId("cite-variant") || {}).value || "orig"; }

  // Bake cne-* variant data into an {id:item} map's multi model (no-op if the
  // module is absent, e.g. in a stripped test harness).
  function enrich(items) {
    return (typeof CiteVariants !== "undefined") ? CiteVariants.enrichItemMap(items) : items;
  }

  // Build a fresh engine for this style+locale+variant. Cheap enough per render,
  // and avoids citeproc citation-registry state leaking across selections.
  function buildEngine(styleFile, lang) {
    return CiteEngine.build({
      styleXml: cache.styles[styleFile],
      locales: cache.locales,
      items: cache.items,
      lang: lang,
      langPrefs: (typeof CiteVariants !== "undefined")
        ? CiteVariants.variantToLangPrefs(currentVariant())
        : null
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
      val.addEventListener("input", debouncedRenderPreview);
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
    persistRefs();
    // Snapshot the OTHER checked rows' checkbox + locator state BEFORE the
    // rebuild below wipes every row back to unchecked/empty — mirrors the
    // capture/restore pattern addFromZotero() uses. Without this, removing
    // one item silently discards every other in-progress checked row.
    var preserved = selectedCitationItems().filter(function (it) { return it.id !== id; });
    itemsPopulated = false;
    populateItems(true);
    preserved.forEach(function (it) {
      var cb = byId("cite-item-" + it.id);
      if (!cb) { return; }
      cb.checked = true;
      var row = cb.parentNode ? cb.parentNode.querySelector(".cite-locator-row") : null;
      if (row) { row.hidden = false; }
      if (it.locator) {
        var esc = (window.CSS && CSS.escape) ? CSS.escape(it.id) : it.id;
        var vEl = document.querySelector('[data-cite-loc-value="' + esc + '"]');
        var tEl = document.querySelector('[data-cite-loc-type="' + esc + '"]');
        if (vEl) { vEl.value = it.locator; }
        if (tEl && it.label) { tEl.value = it.label; }
      }
    });
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

  // Locator value fields fire on every keystroke; each renderPreview() rebuilds
  // a fresh citeproc engine, so re-rendering per keystroke is noticeably slow.
  // Debounce so we only re-render once typing pauses.
  var previewTimer = null;
  function debouncedRenderPreview() {
    if (previewTimer) { clearTimeout(previewTimer); }
    previewTimer = setTimeout(function () { previewTimer = null; renderPreview(); }, 200);
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

  // Build the citation body for a range and insert it, honoring RTL (OOXML) vs
  // LTR (HTML). Returns a Promise<Word.Range> resolving to the inserted range.
  // `ctx` + `range` are live Word objects. For rtl it reads the doc cs font and
  // inserts OOXML (bidi + right-justified paragraph, non-italic Arabic runs);
  // else it inserts sanitized/RTL-run-wrapped HTML. Shared by insertCitation's
  // footnote/endnote branch and refreshCitations() so both paths format a
  // citation identically.
  // `csFont` is optional: pass a pre-resolved complex-script font (read once by
  // a caller that's iterating many ranges, e.g. refreshCitations()) to skip the
  // per-call readDocCsFont() sync; omitted, it reads it itself (insertCitation's
  // single-CC path — unchanged behavior).
  function renderCitationInto(ctx, range, items, styleFile, lang, csFont) {
    var engine = buildEngine(styleFile, lang);
    if (isRtlLang(lang)) {
      var csFontPromise = csFont ? Promise.resolve(csFont) : readDocCsFont(ctx, range);
      return csFontPromise.then(function (resolvedCsFont) {
        var pkg = AshaarTabStop.wrapOoxml(
          CiteWord.buildCitationParagraphOoxml(CiteWord.sanitize(engine.cite(items)), { csFont: resolvedCsFont }));
        return range.insertOoxml(pkg, Word.InsertLocation.replace);
      });
    }
    var html = CiteWord.wrapRtlRuns(CiteWord.sanitize(engine.cite(items)));
    return Promise.resolve(range.insertHtml(html, Word.InsertLocation.replace));
  }

  // Same render strategy as renderCitationInto, but for the whole-library
  // bibliography (engine.bibliography() takes no items). Only used by
  // refreshCitations() — insertBibliography() has its own payload/CC-title
  // handling that isn't otherwise shared. `csFont` is optional — see
  // renderCitationInto's comment.
  function renderBibliographyInto(ctx, range, styleFile, lang, csFont) {
    var engine = buildEngine(styleFile, lang);
    if (isRtlLang(lang)) {
      var csFontPromise = csFont ? Promise.resolve(csFont) : readDocCsFont(ctx, range);
      return csFontPromise.then(function (resolvedCsFont) {
        var pkg = AshaarTabStop.wrapOoxml(
          CiteWord.buildCitationParagraphOoxml(CiteWord.sanitize(engine.bibliography()), { csFont: resolvedCsFont }));
        return range.insertOoxml(pkg, Word.InsertLocation.replace);
      });
    }
    var html = CiteWord.wrapRtlRuns(CiteWord.sanitize(engine.bibliography()));
    return Promise.resolve(range.insertHtml(html, Word.InsertLocation.replace));
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
      var citeTag = CiteWord.buildCitationTag({ style: styleFile, locale: lang, variant: currentVariant(), items: items });
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
          var noteRange = note.body.getRange();
          // Arabic footnote/endnote: renderCitationInto inserts OOXML runs (not
          // insertHtml) so the document's complex-script font is set and the
          // Arabic title isn't left italic (Word renders italic Arabic as tofu
          // squares in fonts lacking an italic style). The paragraph already
          // carries <w:bidi/>, so no separate alignment pass is needed for rtl.
          return renderCitationInto(ctx, noteRange, items, styleFile, lang).then(function (insertedRange) {
            var occ = insertedRange.insertContentControl();
            occ.tag = citeTag;
            occ.title = "Ashaar Citation";
            if (rtl) { return ctx.sync(); }
            var paras = insertedRange.paragraphs;
            paras.load("items");
            return ctx.sync();
          });
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
        // Reset the locator type selects back to "page" too, so a fresh
        // per-insertion citation starts clean rather than keeping whatever
        // type was left over from the citation just inserted.
        var types = document.querySelectorAll("#cite-items .cite-locator-type");
        for (var j = 0; j < types.length; j++) { types[j].value = "page"; }
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
      var bibTag = CiteWord.buildBibliographyTag({ style: styleFile, locale: lang, variant: currentVariant() });
      // Arabic: wrap in a block <p dir="rtl"> for true RTL paragraph reading order.
      var bibHtml = rtl ? '<p dir="rtl">' + payload.html + "</p>" : payload.html;
      if (typeof Word === "undefined" || !Word.run) {
        setStatus("Word isn't available — this is preview-only in a browser.", true);
        return;
      }
      Word.run(function (ctx) {
        var selRange = ctx.document.getSelection().getRange();
        if (rtl) {
          // Arabic bibliography: OOXML runs (doc CS font, non-italic Arabic
          // titles) instead of insertHtml — same rationale as the note branch.
          return readDocCsFont(ctx, selRange).then(function (csFont) {
            var sanitized = CiteWord.sanitize(engine.bibliography());
            var pkg = AshaarTabStop.wrapOoxml(CiteWord.buildCitationParagraphOoxml(sanitized, { csFont: csFont }));
            var oRange = selRange.insertOoxml(pkg, Word.InsertLocation.after);
            var occ = oRange.insertContentControl();
            occ.tag = bibTag;
            occ.title = "Ashaar Bibliography";
            return ctx.sync();
          });
        }
        var range = selRange.insertHtml(bibHtml, Word.InsertLocation.after);
        var cc = range.insertContentControl();
        cc.tag = bibTag;
        cc.title = "Ashaar Bibliography";
        return ctx.sync();
      }).then(function () {
        setStatus("Inserted bibliography.");
      }).catch(function (e) {
        setStatus("Insert failed: " + (e && e.message ? e.message : String(e)), true);
      });
    }).catch(function (e) {
      setStatus("Couldn't load citation assets: " + (e && e.message ? e.message : String(e)), true);
    });
  }

  // Scan the document for AshaarCite:/AshaarBib: tagged content controls and
  // re-render each in place at the pane's CURRENT style/locale, rewriting the
  // tag so the new style sticks. Reuses renderCitationInto/renderBibliographyInto
  // (same body-building logic insertCitation/insertBibliography use) so a
  // refreshed citation looks identical to a freshly-inserted one.
  //
  // document.contentControls only reaches the MAIN BODY story — footnote and
  // endnote content lives in separate stories, so footnote/endnote citations
  // are only reachable via document.body.footnotes / document.body.endnotes
  // (Word.NoteItemCollection, WordApi 1.5) → each note's .body.contentControls.
  // Confirmed against Microsoft Learn (Word.Body class): "footnotes"/"endnotes"
  // are both WordApi 1.5 properties of Word.Body returning NoteItemCollection;
  // each Word.NoteItem exposes a .body (Word.Body) that in turn has its own
  // .contentControls. Wrapped in try/catch so an older Word build (< 1.5, or
  // any host that doesn't expose these) still refreshes the main-body set.
  function refreshCitations() {
    if (typeof Word === "undefined" || !Word.run) { setStatus("Refresh needs Word.", true); return; }
    var styleFile = currentStyleFile();
    var lang = currentLang();
    var counts = { refreshed: 0, bibs: 0, unresolved: 0, failed: 0 };
    var notesReached = true;
    var footnoteCcSeen = 0;
    setStatus("Refreshing citations…");
    ensureAssets(styleFile).then(function () {
      return Word.run(function (ctx) {
        var main = ctx.document.contentControls;
        main.load("items/tag");

        var footnotes = null, endnotes = null;
        try {
          footnotes = ctx.document.body.footnotes;
          endnotes = ctx.document.body.endnotes;
          footnotes.load("items");
          endnotes.load("items");
        } catch (e) {
          notesReached = false;
        }

        return ctx.sync().then(function () {
          var noteBodies = [];
          if (notesReached) {
            try {
              var i;
              for (i = 0; i < footnotes.items.length; i++) { noteBodies.push(footnotes.items[i].body); }
              for (i = 0; i < endnotes.items.length; i++) { noteBodies.push(endnotes.items[i].body); }
            } catch (e) {
              notesReached = false;
              noteBodies = [];
            }
          }
          noteBodies.forEach(function (b) { b.contentControls.load("items/tag"); });

          return ctx.sync().then(function () {
            var ccs = main.items.slice();
            noteBodies.forEach(function (b) {
              // Only count OUR content controls (AshaarCite:/AshaarBib:) —
              // an unrelated CC in a footnote/endnote body shouldn't suppress
              // the "No footnote citations reached" hint below.
              var items = b.contentControls.items;
              for (var j = 0; j < items.length; j++) {
                var t = String(items[j].tag || "");
                if (t.indexOf("AshaarCite:") === 0 || t.indexOf("AshaarBib:") === 0) { footnoteCcSeen++; }
              }
              ccs = ccs.concat(items);
            });

            // The complex-script font is a document-level property (Ashaar
            // Normal style / doc-body bidi font), not per-citation — read it
            // ONCE here instead of once per CC (readDocCsFont issues its own
            // ctx.sync(), so per-CC reads would mean one extra sync per Arabic
            // CC). Only needed at all for rtl locales.
            var rtl = isRtlLang(lang);
            var csFontPromise = rtl
              ? readDocCsFont(ctx, ctx.document.body.getRange())
              : Promise.resolve(undefined);

            return csFontPromise.then(function (csFont) {
              var ops = [];
              ccs.forEach(function (cc) {
                var tagStr = String(cc.tag || "");
                try {
                  if (tagStr.indexOf("AshaarCite:") === 0) {
                    var parsed = CiteWord.parseCitationTag(cc.tag);
                    if (!parsed) { return; } // malformed despite the prefix — not really ours
                    var items = CiteWord.citationItemsFromTag(parsed);
                    var allResolved = items.every(function (it) {
                      return cache.items && Object.prototype.hasOwnProperty.call(cache.items, it.id);
                    });
                    if (!allResolved) { counts.unresolved++; return; }
                    // Content (not whole-control) range: insertOoxml/insertHtml
                    // "Replace" on a content control's WHOLE range throws
                    // GeneralException and can orphan the control. Set the tag
                    // BEFORE the body replace so it sticks regardless of how
                    // the range op reshapes the control.
                    var range = cc.getRange("Content");
                    cc.tag = CiteWord.buildCitationTag({ style: styleFile, locale: lang, variant: currentVariant(), items: items });
                    ops.push(renderCitationInto(ctx, range, items, styleFile, lang, csFont).then(function () {
                      counts.refreshed++;
                    }).catch(function () { counts.failed++; }));
                  } else if (tagStr.indexOf("AshaarBib:") === 0) {
                    var bibRange = cc.getRange("Content");
                    cc.tag = CiteWord.buildBibliographyTag({ style: styleFile, locale: lang, variant: currentVariant() });
                    ops.push(renderBibliographyInto(ctx, bibRange, styleFile, lang, csFont).then(function () {
                      counts.bibs++;
                    }).catch(function () { counts.failed++; }));
                  }
                  // else: not one of ours — leave untouched
                } catch (e) { counts.failed++; }
              });

              // Office.js only QUEUES the insertOoxml/insertHtml calls above —
              // nothing actually executes until the ctx.sync() below. So each
              // op's .catch(counts.failed++) only catches synchronous/queueing
              // errors; if a queued op fails at EXECUTION time, this ctx.sync()
              // rejects for the whole batch (caught by the outer .catch, and
              // Word rolls back the entire run) rather than surfacing on that
              // op alone. refreshed/bibs/failed are therefore optimistic —
              // they reflect intent, and are only trustworthy once this sync
              // (and the run's outer promise) resolves without error.
              return Promise.all(ops).then(function () { return ctx.sync(); });
            });
          });
        });
      }).then(function () {
        var msg = "Refreshed " + counts.refreshed + " citation(s), " + counts.bibs + " bibliography.";
        if (counts.unresolved) { msg += " " + counts.unresolved + " unresolved (re-add from Zotero)."; }
        if (counts.failed) { msg += " " + counts.failed + " failed."; }
        if (!notesReached || footnoteCcSeen === 0) { msg += " No footnote citations reached."; }
        setStatus(msg, !!(counts.unresolved || counts.failed));
      });
    }).catch(function (e) {
      setStatus("Refresh failed: " + (e && e.message ? e.message : String(e)), true);
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
        var enriched = enrich(items); // bake cne-* variants into the multi model
        Object.keys(enriched).forEach(function (id) {
          cache.items[id] = enriched[id];
        });
        persistRefs();
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
    var variantSel = byId("cite-variant");
    if (variantSel) { variantSel.addEventListener("change", renderPreview); }
    insertBtn.addEventListener("click", insertCitation);
    byId("cite-insert-bib").addEventListener("click", insertBibliography);
    var refreshBtn = byId("cite-refresh");
    if (refreshBtn) { refreshBtn.addEventListener("click", refreshCitations); }
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

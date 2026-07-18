# Ashaar for Word — Citations Quick How-To

Insert CSL-formatted citations (footnote, endnote, or inline) and a bibliography, drawn
from your Zotero library or the built-in sample. Full reference: [USER_GUIDE.md](USER_GUIDE.md).
Everyday poetry workflow: [HOW-TO.md](HOW-TO.md). Deeper walkthroughs: [HOW-TO-WORKFLOWS.md](HOW-TO-WORKFLOWS.md).

## Setup: Zotero, Better BibTeX & CNE

**None of this is needed to try the Cite tab** — it ships with a built-in sample library, and you
can now also **add citations manually** (the "Add manually" button) with no Zotero at all. You only
need the tools below to cite from **your own** Zotero library.

1. **Zotero** — the reference manager. Install it from [zotero.org/download](https://www.zotero.org/download/).
   (The add-in is built and tested against **Zotero 9**.) Keep Zotero running while you cite.

2. **Better BibTeX (BBT)** — required for "Add from Zotero." It gives each item a stable *citekey*
   and provides the two things the add-in uses: the **CAYW** ("cite as you write") picker and the
   **Better CSL JSON** export.
   - Download the latest `.xpi` from the
     [Better BibTeX releases](https://github.com/retorquere/zotero-better-bibtex/releases).
   - In Zotero: **Tools → Plugins** (called *Add-ons* in older versions) → the gear ⚙ →
     **Install Plugin From File…** → choose the `.xpi` → **restart Zotero**.
   - Without BBT, the picker and export won't resolve and "Add from Zotero" will report Zotero as
     unavailable.

3. **CNE — Cite Non-English** *(optional; only for Arabic ↔ romanized variants)* — lets you author an
   item's Arabic original plus its romanization/translation. CNE stores these as `cne-*` lines in the
   item's **Extra** field, which the add-in reads to power the **Variant** dropdown (Original /
   Romanized / Both).
   - Download the `.xpi` from [cite-non-english](https://github.com/boan-anbo/cite-non-english/releases).
   - Install it the same way (**Install Plugin From File…** → restart Zotero).
   - Without CNE, an item renders in whatever single language its fields hold and the Variant dropdown
     has nothing to switch to. *(Legacy items carrying Juris-M "mlzsync" data can be converted once via
     `npm run migrate:cne`, which uses Zotero's local API — enable it under
     **Settings → Advanced → "Allow other applications on this computer to communicate with Zotero."**)*

**How the connection works:** the add-in reaches Zotero through a small local bridge on
`localhost:23119` (Better BibTeX's JSON-RPC + CAYW) — everything stays on your machine. The live
Zotero connection runs through the add-in's local server, so use the `npm start` / dev-server setup
when citing from your library.

## 1. Open the Cite tab

Open the task pane (**Home → Ashaar Poetry**) and switch to the **Cite** tab. You can try
everything against the built-in sample library right away — Zotero is only needed to cite
from your own references.

## 2. Add references

- **From Zotero** — click **Add from Zotero**. Zotero's citation picker opens; choose one or
  more items and confirm. They appear (checked) in the **Items** list and in the live preview.
  Requires **Zotero 9** running with the **Better BibTeX** plugin installed.
- **From the sample library** — the Items list is pre-populated so you can explore styles and
  output without Zotero.

Tick the items you want to cite. Use the **×** next to an item to drop it from the list.

## 3. Choose style, locale, and form

Four controls at the top drive every citation and the bibliography (the preview updates live):

- **Style** — *Chicago (notes & bibliography)*, *APA (author-date)*, or the two Fatemi variants
  (see §6).
- **Locale** — *English (en-US)* or *Arabic (ar)*. Arabic renders right-to-left with the
  document's complex-script font.
- **Variant** — for multilingual items: *Original (ar)*, *Romanized*, or *Both (orig + romanized)*.
- **Output form** — *Footnote*, *Endnote*, or *Inline*.

## 4. Add locators (optional)

Each item in the list has a locator field — set a **page**, **chapter**, **section**, or **verse**
value to cite a specific place (e.g. "p. 42"). Locators are per-insertion: they apply to the next
citation you insert, then clear, so you can cite the same source again at a different spot.

## 5. Insert

- **Insert citation** — places the citation at the cursor in the chosen form. Footnotes/endnotes
  are inserted as real Word notes; Arabic notes are set right-to-left with italic suppressed (Arabic
  type doesn't use italics).
- **Insert bibliography** — inserts a formatted bibliography of the items at the cursor.
- **Refresh citations** — after you change the **Style** or **Locale**, click this to re-format
  every citation and bibliography already in the document (in the main body *and* inside
  footnotes/endnotes) to the new style, in place. Locators are preserved.

## 6. Fatemi styles — source-classified bibliography

The two **— Fatemi** styles (*Chicago — Fatemi*, *APA — Fatemi*) split the bibliography into
headed sections by two independent axes read from your **Zotero tags**:

- **Corpus** — tag an item `corpus:fatemi` to mark it Fatemi (untagged = non-Fatemi).
- **Class** — tag an item `class:secondary` to mark it secondary (untagged = primary).

With a Fatemi style selected, **Insert bibliography** groups entries into up to four sections in
this order — *Primary Sources — Fatemi*, *Primary Sources — Other*, *Secondary Sources — Fatemi*,
*Secondary Sources — Other* — skipping any empty section, with Arabic headings under the `ar`
locale. If only one section would appear (e.g. everything untagged), the list stays flat with no
heading, exactly like a stock style. Individual entries are formatted identically to the stock
Chicago/APA styles — the Fatemi variants only add the sectioning.

**Tagging tip:** you only tag the exceptions. Untagged references default to *Primary / non-Fatemi*.
If Zotero isn't reachable when you insert, the bibliography simply falls back to a single flat list.

## 7. Notes

- **Arabic / RTL** — for the best right-to-left result, run the Styles tab's RTL document setup so
  the footnote numbers and paragraph direction follow the document.
- **Better BibTeX** — citekeys come from Better BibTeX; the picker returns them and the add-in
  fetches each item's CSL-JSON (and tags) through a local Zotero connection.

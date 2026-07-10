#!/usr/bin/env python3
"""Generate a Word .docx that tests the kashida rendering hypotheses directly in
Word (no Office.js). It only references font *names* (no font bytes), so the
fonts must be installed on the machine opening it. Output: test-documents/kashida-word-test.docx

Tests:
  A. Jameel Kasheeda cs-name resolution — does <w:rFonts w:cs="Jameel Noori Nastaleeq Kasheeda">
     render the WIDER face vs the base "Jameel Noori Nastaleeq"?  (the Jameel font-swap gate)
  B. Mehr tatweel rendering — does a trailing U+0640 after a whitelisted word-final
     letter render as a clean kashida in Word (unlike the browser)?
  C. Word-native kashida justification — do Mehr/Jameel stretch under w:jc="highKashida"?
  D. font-swap misra in an RTL table cell — per-run base/Kasheeda cs fonts (mirrors the add-in output).
"""
import html, zipfile, os

BASE = "Jameel Noori Nastaleeq"
KASH = "Jameel Noori Nastaleeq Kasheeda"
MEHR = "Mehr Nastaliq Web"
T = "ـ"      # tatweel
SZ = "80"         # 40pt (half-points)

def esc(s): return html.escape(s, quote=False)

def arun(text, font, sz=SZ):
    return ('<w:r><w:rPr><w:rtl/>'
            f'<w:rFonts w:cs="{font}" w:ascii="{font}" w:hAnsi="{font}"/>'
            f'<w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/></w:rPr>'
            f'<w:t xml:space="preserve">{esc(text)}</w:t></w:r>')

def lrun(text, bold=False, color=None):
    rpr = ('<w:b/>' if bold else '') + (f'<w:color w:val="{color}"/>' if color else '') + '<w:sz w:val="20"/>'
    return f'<w:r><w:rPr>{rpr}</w:rPr><w:t xml:space="preserve">{esc(text)}</w:t></w:r>'

def apara(runs, jc="right"):
    return f'<w:p><w:pPr><w:bidi/><w:jc w:val="{jc}"/><w:spacing w:after="60"/></w:pPr>{runs}</w:p>'

def lpara(text, bold=False, color=None, before="160"):
    return (f'<w:p><w:pPr><w:spacing w:before="{before}" w:after="40"/></w:pPr>'
            f'{lrun(text, bold, color)}</w:p>')

def heading(text):
    return (f'<w:p><w:pPr><w:spacing w:before="260" w:after="60"/>'
            f'<w:pBdr><w:bottom w:val="single" w:sz="6" w:space="2" w:color="999999"/></w:pBdr></w:pPr>'
            f'<w:r><w:rPr><w:b/><w:sz w:val="26"/></w:rPr><w:t xml:space="preserve">{esc(text)}</w:t></w:r></w:p>')

def cell_table(runs_xml, width_dxa=5760):
    return (f'<w:tbl><w:tblPr><w:tblW w:w="{width_dxa}" w:type="dxa"/><w:bidiVisual/>'
            '<w:tblBorders>'
            '<w:top w:val="single" w:sz="4" w:color="BBBBBB"/><w:left w:val="single" w:sz="4" w:color="BBBBBB"/>'
            '<w:bottom w:val="single" w:sz="4" w:color="BBBBBB"/><w:right w:val="single" w:sz="4" w:color="BBBBBB"/>'
            '</w:tblBorders></w:tblPr>'
            f'<w:tblGrid><w:gridCol w:w="{width_dxa}"/></w:tblGrid>'
            f'<w:tr><w:tc><w:tcPr><w:tcW w:w="{width_dxa}" w:type="dxa"/></w:tcPr>'
            f'<w:p><w:pPr><w:bidi/><w:jc w:val="right"/></w:pPr>{runs_xml}</w:p>'
            '</w:tc></w:tr></w:tbl>')

body = []
body.append('<w:p><w:pPr><w:spacing w:after="40"/></w:pPr>'
            '<w:r><w:rPr><w:b/><w:sz w:val="34"/></w:rPr>'
            '<w:t>Kashida rendering tests — open in Word (fonts must be installed)</w:t></w:r></w:p>')
body.append(lpara("For each test, compare the FITTED/variant line to its BASE line above it. "
                  "Notes below say what a PASS looks like. This file references font names only; "
                  "install Jameel Noori Nastaleeq (Regular + Kasheeda) and Mehr Nastaliq Web first.",
                  before="0"))

# ---- Test A: Jameel Kasheeda cs-name resolution ----
body.append(heading("A. Jameel — does the “Kasheeda” cs name resolve the WIDER face?"))
body.append(lpara("PASS = line 2 is visibly WIDER (elongated) than line 1. FAIL = they look identical "
                   "(Word ignored the style name and used Regular — we'd then need a different cs string).", color="7A3E00"))
body.append(lpara("1) base  w:cs=\"Jameel Noori Nastaleeq\""))
body.append(apara(arun("کہہ رہے تھے", BASE)))
body.append(lpara("2) Kasheeda  w:cs=\"Jameel Noori Nastaleeq Kasheeda\"  (should be wider)"))
body.append(apara(arun("کہہ رہے تھے", KASH)))
body.append(lpara("3) same word, isolated, base vs Kasheeda side by side:"))
body.append(apara(arun("خواب", BASE) + arun("      ", BASE) + arun("خواب", KASH)))

# ---- Test B: Mehr tatweel rendering ----
body.append(heading("B. Mehr — does an inserted tatweel (U+0640) render as a clean kashida?"))
body.append(lpara("PASS = the tatweel variants show a smooth elongated stroke. FAIL = broken/stacked/"
                   "disconnected (as it renders in the browser). Whitelisted word-final letters: ب پ ت ٹ ث ف ک گ.", color="7A3E00"))
body.append(lpara("base:"))
body.append(apara(arun("خواب   کتاب   رات   صاحب", MEHR)))
body.append(lpara("trailing tatweel after final letter (خوابـ کتابـ راتـ صاحبـ):"))
body.append(apara(arun("خواب"+T+"   کتاب"+T+"   رات"+T+"   صاحب"+T, MEHR)))
body.append(lpara("medial tatweel between letters (کـتاب — expected inert/ugly):"))
body.append(apara(arun("ک"+T+"تاب", MEHR)))

# ---- Test C: Word-native kashida justification ----
body.append(heading("C. Word-native kashida justification (w:jc=\"highKashida\")"))
body.append(lpara("Independent of our engine: does WORD's own kashida stretch these fonts? "
                   "PASS = the first (wrapped) line elongates to the margin. Kashida never stretches the LAST line.", color="7A3E00"))
long_ur = "ہم نے مانا کہ تغافل نہ کرو گے لیکن خاک ہو جائیں گے ہم تم کو خبر ہونے تک"
body.append(lpara("Mehr, highKashida:"))
body.append(f'<w:p><w:pPr><w:bidi/><w:jc w:val="highKashida"/></w:pPr>{arun(long_ur, MEHR, sz="52")}</w:p>')
body.append(lpara("Jameel (base), highKashida:"))
body.append(f'<w:p><w:pPr><w:bidi/><w:jc w:val="highKashida"/></w:pPr>{arun(long_ur, BASE, sz="52")}</w:p>')

# ---- Test D: font-swap misra in an RTL table cell (mirrors add-in output) ----
body.append(heading("D. Font-swap output in an RTL table cell (as the add-in emits)"))
body.append(lpara("PASS = the swapped words (خواب, عذاب) render wider in the Kasheeda face within the cell, "
                   "the rest in base — a filled, mixed-face misra. This is the shape the add-in inserts.", color="7A3E00"))
# رات کے خواب دن کے عذاب  — swap the خواب (خو,ا,ب) and عذاب (عذ,ا,ب) fasls to Kasheeda
fasls = ["ر","ا","ت"," ","کے"," ","خو","ا","ب"," ","د","ن"," ","کے"," ","عذ","ا","ب"]
swap_idx = {6,7,8,15,16,17}
runs = "".join(arun(f, KASH if i in swap_idx else BASE) for i, f in enumerate(fasls))
body.append(cell_table(runs))
body.append(lpara("For contrast, the same misra entirely in base (no swap):"))
body.append(cell_table("".join(arun(f, BASE) for f in fasls)))

document = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            f'<w:body>{"".join(body)}'
            '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>'
            '<w:pgMar w:top="1080" w:right="1080" w:bottom="1080" w:left="1080"/></w:sectPr>'
            '</w:body></w:document>')

content_types = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                 '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                 '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                 '<Default Extension="xml" ContentType="application/xml"/>'
                 '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
                 '</Types>')

rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
        '</Relationships>')

out = os.path.join(os.path.dirname(__file__), "..", "test-documents", "kashida-word-test.docx")
out = os.path.abspath(out)
os.makedirs(os.path.dirname(out), exist_ok=True)
with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml", content_types)
    z.writestr("_rels/.rels", rels)
    z.writestr("word/document.xml", document)
print("wrote", out)

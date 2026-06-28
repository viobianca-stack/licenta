# -*- coding: utf-8 -*-
"""
Prezentare licenta - design ea5b5019, text minimal + figuri din lucrare.
Accent pe rezultate si comparatia cu state-of-the-art.
Slide size: 20" x 11.25".
"""

import copy
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
import lxml.etree as etree
from PIL import Image

TEMPLATE_PATH = "/root/.claude/uploads/03ba0bac-fff2-5a16-859d-96a3dba91012/ea5b5019-Design_f_r__titlu.pptx"
OUTPUT_PATH = "/home/user/licenta/Prezentare_Licenta_Template.pptx"
FIGS = "/home/user/licenta/.figs"

BLUE   = RGBColor(0x00, 0x4A, 0xAD)
DBLUE  = RGBColor(0x00, 0x2D, 0x7A)
LBLUE  = RGBColor(0x72, 0x9B, 0xCB)
DARK   = RGBColor(0x30, 0x36, 0x42)
WHITE  = RGBColor(0xFF, 0xFF, 0xFF)
LGRAY  = RGBColor(0x60, 0x60, 0x60)
PANEL  = RGBColor(0xE8, 0xF0, 0xFB)
GREEN  = RGBColor(0x1E, 0x8A, 0x4E)
REDX   = RGBColor(0xC0, 0x39, 0x39)

SW, SH = Inches(20), Inches(11.25)
FONT = "Calibri"
TOTAL = 12


# ─── infra ────────────────────────────────────────────────────────────────────

def clone_slide(src_prs, idx, dst_prs, keep_decor=False):
    src = src_prs.slides[idx]
    slide = dst_prs.slides.add_slide(dst_prs.slide_layouts[6])
    tree = slide.shapes._spTree
    for ch in list(tree):
        tree.remove(ch)
    for ch in src.shapes._spTree:
        tree.append(copy.deepcopy(ch))
    for rel in src.part.rels.values():
        if not rel.is_external and "/image" in rel.reltype:
            slide.part.relate_to(rel.target_part, rel.reltype)
    # strip text boxes always; on content slides also strip stock photos &
    # decorative groups, keeping only the subtle full-bleed background.
    shapes = list(slide.shapes)
    for i, sh in enumerate(shapes):
        remove = False
        if sh.has_text_frame:
            remove = True
        elif not keep_decor and i > 0 and sh.shape_type in (6, 13):
            # keep i==0 (full-bleed background freeform), drop groups/pictures
            remove = True
        if remove:
            sh._element.getparent().remove(sh._element)
    return slide


def txb(slide, l, t, w, h, text, size=18, bold=False, color=DARK,
        align=PP_ALIGN.LEFT, italic=False, anchor=None):
    box = slide.shapes.add_textbox(l, t, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    if anchor:
        tf.vertical_anchor = anchor
    for j, line in enumerate(text.split("\n")):
        p = tf.paragraphs[0] if j == 0 else tf.add_paragraph()
        p.alignment = align
        r = p.add_run(); r.text = line
        r.font.size = Pt(size); r.font.bold = bold
        r.font.italic = italic; r.font.color.rgb = color; r.font.name = FONT
    return box


def bullets(slide, l, t, w, h, items, size=16, color=DARK, spacing=1.3, gap=6):
    box = slide.shapes.add_textbox(l, t, w, h)
    tf = box.text_frame; tf.word_wrap = True
    for j, it in enumerate(items):
        p = tf.paragraphs[0] if j == 0 else tf.add_paragraph()
        pPr = p._pPr if p._pPr is not None else etree.SubElement(p._p, qn('a:pPr'))
        ln = etree.SubElement(pPr, qn('a:lnSpc')); etree.SubElement(ln, qn('a:spcPct')).set('val', str(int(spacing*100000)))
        sb = etree.SubElement(pPr, qn('a:spcBef')); etree.SubElement(sb, qn('a:spcPts')).set('val', str(gap*100))
        r = p.add_run(); r.text = it
        r.font.size = Pt(size); r.font.color.rgb = color; r.font.name = FONT
    return box


def rect(slide, l, t, w, h, color, line=None):
    s = slide.shapes.add_shape(1, l, t, w, h)
    s.fill.solid(); s.fill.fore_color.rgb = color
    if line:
        s.line.color.rgb = line; s.line.width = Pt(1)
    else:
        s.line.fill.background()
    s.shadow.inherit = False
    return s


def hline(slide, l, t, w, color=BLUE, thick=4):
    return rect(slide, l, t, w, Pt(thick), color)


def pic(slide, path, l, t, max_w, max_h, card=True, pad=Inches(0.18)):
    """Place image scaled into max box, optionally on a white card. Centered."""
    iw, ih = Image.open(path).size
    ar = iw / ih
    bw, bh = max_w, max_h
    if bw / bh > ar:
        h = bh; w = int(bh * ar)
    else:
        w = bw; h = int(bw / ar)
    cx = l + (max_w - w) // 2
    cy = t + (max_h - h) // 2
    if card:
        rect(slide, cx - pad, cy - pad, w + 2*pad, h + 2*pad, WHITE)
    slide.shapes.add_picture(path, cx, cy, w, h)
    return cx, cy, w, h


def header(slide, title, num, kicker=None):
    txb(slide, Inches(1), Inches(0.5), Inches(15), Inches(0.9),
        title, size=30, bold=True, color=BLUE)
    hline(slide, Inches(1), Inches(1.45), Inches(3.5), BLUE, 4)
    if kicker:
        txb(slide, Inches(1), Inches(1.6), Inches(15), Inches(0.5),
            kicker, size=15, italic=True, color=LGRAY)
    txb(slide, SW - Inches(2.4), SH - Inches(0.75), Inches(1.9), Inches(0.5),
        f"{num} / {TOTAL}", size=12, color=LGRAY, align=PP_ALIGN.RIGHT)


def table(slide, l, t, w, rows, col_w, header_fill=BLUE, row_h=Inches(0.62),
          fsize=14, hsize=14, highlight_row=None):
    """rows: list of list[str]; first row = header."""
    n = len(rows); ncol = len(rows[0])
    x = l
    xs = [l]
    for cw in col_w[:-1]:
        x += cw; xs.append(x)
    y = t
    for ri, row in enumerate(rows):
        rh = row_h
        cy = y
        for ci, cell in enumerate(row):
            cw = col_w[ci]
            if ri == 0:
                fill = header_fill; fc = WHITE; bold = True; fs = hsize
            elif highlight_row is not None and ri == highlight_row:
                fill = PANEL; fc = DBLUE; bold = True; fs = fsize
            else:
                fill = WHITE if ri % 2 else RGBColor(0xF3, 0xF6, 0xFB)
                fc = DARK; bold = (ci == 0); fs = fsize
            c = rect(slide, xs[ci], cy, cw, rh, fill, line=RGBColor(0xD5,0xDE,0xEA))
            tf = c.text_frame; tf.word_wrap = True
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            tf.margin_left = Inches(0.1); tf.margin_right = Inches(0.05)
            tf.margin_top = 0; tf.margin_bottom = 0
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.LEFT if ci == 0 else PP_ALIGN.CENTER
            r = p.add_run(); r.text = cell
            r.font.size = Pt(fs); r.font.bold = bold
            r.font.color.rgb = fc; r.font.name = FONT
        y += rh
    return y


# ─── slides ───────────────────────────────────────────────────────────────────

def s_title(prs, src):
    sl = clone_slide(src, 0, prs, keep_decor=True)
    cx, cy = SW//2, SH//2
    txb(sl, cx-Inches(6), cy-Inches(2.5), Inches(12), Inches(0.6),
        "LUCRARE DE DIPLOMĂ · 2026", size=16, bold=True, color=LGRAY, align=PP_ALIGN.CENTER)
    txb(sl, cx-Inches(7), cy-Inches(1.7), Inches(14), Inches(1.9),
        "Detectarea mesajelor generate de AI\npe rețelele sociale",
        size=40, bold=True, color=BLUE, align=PP_ALIGN.CENTER)
    txb(sl, cx-Inches(5), cy+Inches(0.7), Inches(10), Inches(1.2),
        "Absolvent: Bianca-Ștefania GHEORGHE\nCoordonator: Conf.dr.ing. Daniel-Marian MEREZEANU",
        size=18, color=DARK, align=PP_ALIGN.CENTER)
    txb(sl, cx-Inches(7), cy+Inches(2.4), Inches(14), Inches(0.6),
        "Universitatea Națională de Știință și Tehnologie POLITEHNICA București · Facultatea Automatică și Calculatoare",
        size=13, color=LGRAY, align=PP_ALIGN.CENTER)


def s_cuprins(prs, src):
    sl = clone_slide(src, 3, prs, keep_decor=True)
    txb(sl, Inches(1), Inches(0.7), Inches(18), Inches(1.0),
        "Cuprins", size=32, bold=True, color=BLUE, align=PP_ALIGN.CENTER)
    hline(sl, Inches(9), Inches(1.7), Inches(2), BLUE, 4)
    items = [
        ("01", "Problema și stadiul actual"),
        ("02", "Obiective și tehnologii"),
        ("03", "Arhitectura sistemului"),
        ("04", "Rezultate obținute"),
        ("05", "Comparație cu state-of-the-art"),
        ("06", "Contribuții și demo aplicație"),
    ]
    for i, (n, t) in enumerate(items):
        col = i % 2; row = i // 2
        x = Inches(2.2) + col * Inches(8.5)
        y = Inches(2.8) + row * Inches(2.0)
        rect(sl, x, y, Inches(1.3), Inches(1.3), BLUE)
        txb(sl, x, y+Inches(0.18), Inches(1.3), Inches(1.0), n, size=30, bold=True,
            color=WHITE, align=PP_ALIGN.CENTER)
        txb(sl, x+Inches(1.7), y+Inches(0.25), Inches(6.3), Inches(0.9), t,
            size=20, bold=True, color=DARK)


def s_problema(prs, src):
    sl = clone_slide(src, 1, prs)
    header(sl, "Problema și stadiul actual", 3)
    bullets(sl, Inches(1), Inches(2.3), Inches(6.6), Inches(6),
        [
            "LLM-urile (ChatGPT, GPT-4, Gemini) produc texte\nindistinctibile de cele umane.",
            "Rețelele sociale = publicare instant, fără filtrare\n→ dezinformare rapidă.",
            "Texte scurte și informale: argou, abrevieri, greșeli\n→ metodele clasice dau greș.",
            "Instrumentele existente sunt evaluate pe texte\nlungi și formale.",
            "Parafrazarea scade acuratețea de la >90% la <50%.",
        ], size=17, spacing=1.15, gap=14)
    rect(sl, Inches(1), Inches(8.7), Inches(6.6), Inches(1.4), PANEL)
    txb(sl, Inches(1.25), Inches(8.9), Inches(6.1), Inches(1.0),
        "Niciun instrument comercial nu este optimizat\npentru textele scurte de social media.",
        size=16, bold=True, color=DBLUE)
    # pipeline figure
    txb(sl, Inches(8.2), Inches(2.2), Inches(11), Inches(0.5),
        "Peisajul metodelor de detecție (pipeline studiat)", size=15, italic=True, color=LGRAY)
    pic(sl, f"{FIGS}/fig_2_1_pipeline_0.png", Inches(8), Inches(2.8), Inches(11.2), Inches(7.2))


def s_obiective(prs, src):
    sl = clone_slide(src, 2, prs)
    header(sl, "Obiective", 4)
    items = [
        ("1", "Analiză critică", "Metodele existente și limitările lor pe social media."),
        ("2", "Soluție specializată", "Detector ML antrenat pe texte scurte (50-500 caractere)."),
        ("3", "Comparație de modele", "3 clasificatori complementari, condiții identice."),
        ("4", "Aplicație accesibilă", "Interfață web cu justificare transparentă a deciziei."),
    ]
    for i, (n, t, d) in enumerate(items):
        col = i % 2; row = i // 2
        x = Inches(1.2) + col * Inches(9.3)
        y = Inches(2.6) + row * Inches(3.6)
        rect(sl, x, y, Inches(8.6), Inches(3.0), PANEL)
        rect(sl, x, y, Inches(0.25), Inches(3.0), BLUE)
        rect(sl, x+Inches(0.6), y+Inches(0.55), Inches(1.3), Inches(1.3), BLUE)
        txb(sl, x+Inches(0.6), y+Inches(0.72), Inches(1.3), Inches(1.0), n, size=34,
            bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        txb(sl, x+Inches(2.2), y+Inches(0.55), Inches(6), Inches(0.7), t, size=22, bold=True, color=BLUE)
        txb(sl, x+Inches(2.2), y+Inches(1.35), Inches(6), Inches(1.3), d, size=17, color=DARK)


def s_tehnologii(prs, src):
    sl = clone_slide(src, 7, prs)
    header(sl, "Tehnologii utilizate", 5)
    cats = [
        ("Date", ["Python 3.10", "pandas", "datasets (HuggingFace)"]),
        ("ML clasic", ["scikit-learn", "TF-IDF (10k, n-grame 1-2)", "NLTK"]),
        ("Deep Learning", ["PyTorch 2.x (GPU)", "Transformers — RoBERTa", "SpaCy NER"]),
        ("Aplicație web", ["Flask 3.x (REST)", "HTML / CSS / JS", "Chart.js"]),
    ]
    for i, (cat, items) in enumerate(cats):
        x = Inches(1) + i * Inches(4.65)
        y = Inches(2.7)
        rect(sl, x, y, Inches(4.3), Inches(4.6), PANEL)
        rect(sl, x, y, Inches(4.3), Inches(0.9), BLUE)
        txb(sl, x+Inches(0.3), y+Inches(0.15), Inches(3.8), Inches(0.6), cat,
            size=20, bold=True, color=WHITE)
        bullets(sl, x+Inches(0.35), y+Inches(1.2), Inches(3.7), Inches(3.2),
            ["• "+it for it in items], size=17, gap=12)
    rect(sl, Inches(1), Inches(8.0), Inches(18.3), Inches(2.0), RGBColor(0xF3,0xF6,0xFB))
    rect(sl, Inches(1), Inches(8.0), Inches(0.25), Inches(2.0), BLUE)
    txb(sl, Inches(1.4), Inches(8.2), Inches(5), Inches(0.6), "Mediu de antrenare",
        size=18, bold=True, color=BLUE)
    txb(sl, Inches(1.4), Inches(8.8), Inches(17.5), Inches(1.1),
        "Google Colab · GPU NVIDIA Tesla T4 (15.6 GB) · FP16 · RoBERTa ~85 min (3 epoci) · modele clasice <2 min · inferență CPU 0.5-2 s/text",
        size=16, color=DARK)


def s_arhitectura(prs, src):
    sl = clone_slide(src, 6, prs)
    header(sl, "Arhitectura sistemului", 6, "Figura 3.2 — pipeline complet: surse → preprocesare → split → antrenare → evaluare")
    pic(sl, f"{FIGS}/fig_3_2_arhitectura_0.png", Inches(0.6), Inches(2.4), Inches(10.5), Inches(8.5))
    # right side: dataset summary
    x = Inches(11.6)
    rect(sl, x, Inches(2.6), Inches(7.7), Inches(7.7), PANEL)
    txb(sl, x+Inches(0.5), Inches(2.9), Inches(6.7), Inches(0.9), "45.834 exemple",
        size=40, bold=True, color=BLUE, align=PP_ALIGN.CENTER)
    txb(sl, x+Inches(0.5), Inches(3.85), Inches(6.7), Inches(0.5),
        "echilibrate 50/50  uman / AI", size=19, color=DARK, align=PP_ALIGN.CENTER)
    hline(sl, x+Inches(1.5), Inches(4.6), Inches(4.7), BLUE, 2)
    bullets(sl, x+Inches(0.6), Inches(4.85), Inches(6.6), Inches(2.6),
        ["• Filtrate 50-500 caractere",
         "• Split stratificat 70 / 15 / 15",
         "• 6 surse: HC3, TweetEval, Reddit, RAID,\n  AI_Human, AI Detection",
         "• 3 modele: Naive Bayes · Reg. Logistică · RoBERTa"], size=17, gap=12)
    txb(sl, x+Inches(0.5), Inches(9.5), Inches(6.7), Inches(0.7),
        "Backend Flask (/predict, /batch) + frontend single-page",
        size=15, italic=True, color=LGRAY, align=PP_ALIGN.CENTER)


def s_rezultate(prs, src):
    sl = clone_slide(src, 3, prs)
    header(sl, "Rezultate obținute", 7, "Set de test: 6.876 exemple, echilibrat 50/50")
    # Tabel 4.5
    txb(sl, Inches(1), Inches(2.3), Inches(9), Inches(0.5),
        "Tabel 4.5 — Performanță comparativă", size=16, bold=True, color=BLUE)
    rows = [
        ["Model", "Acuratețe", "F1", "Precision", "Recall"],
        ["Naive Bayes", "94.24%", "94.17%", "95.35%", "93.02%"],
        ["Reg. Logistică", "96.71%", "96.66%", "98.11%", "95.26%"],
        ["RoBERTa", "99.17%", "99.18%", "98.56%", "99.80%"],
    ]
    cw = [Inches(2.6), Inches(1.55), Inches(1.5), Inches(1.55), Inches(1.5)]
    table(sl, Inches(1), Inches(2.9), Inches(8.7), rows, cw, row_h=Inches(0.75),
          fsize=15, hsize=15, highlight_row=3)
    txb(sl, Inches(1), Inches(6.2), Inches(8.7), Inches(2.2),
        "RoBERTa depășește semnificativ modelele clasice:\n+2.5 puncte față de Reg. Logistică, +5 puncte față de Naive Bayes.",
        size=17, color=DARK)
    rect(sl, Inches(1), Inches(7.6), Inches(8.7), Inches(2.4), PANEL)
    bullets(sl, Inches(1.3), Inches(7.85), Inches(8.1), Inches(2.0),
        ["Pe sursele de social media: acuratețe > 99.5%",
         "Doar 57 erori / 6.876 (50 fals-pozitive, 7 fals-negative)",
         "Aproape toate erorile vin din texte formale, nu social media"],
        size=16, color=DBLUE, gap=10)
    # Figura 4.1 confusion matrix
    txb(sl, Inches(10.3), Inches(2.3), Inches(9), Inches(0.5),
        "Figura 4.1 — Confusion matrix + curba de antrenare", size=16, bold=True, color=BLUE)
    pic(sl, f"{FIGS}/fig_4_1_confusion_clean.png", Inches(10.2), Inches(3.0), Inches(9.2), Inches(7.0))


def s_sota(prs, src):
    sl = clone_slide(src, 8, prs)
    header(sl, "Mai bine decât state-of-the-art", 8,
           "Tabel 5.1 — soluția propusă vs. instrumentele comerciale")
    rows = [
        ["Criteriu", "Soluția propusă", "GPTZero", "ZeroGPT", "OpenAI", "Originality", "Sapling"],
        ["Acuratețe (social media)", "99.17%", "n/a", "n/a", "26%", "76%", "68-87%"],
        ["Texte scurte <500 car.", "Optimizat", "Slab", "Slab", "F. slab", "Moderat", "Moderat"],
        ["Justificare per cuvânt", "Da (LOO)", "Nu", "Nu", "Nu", "Nu", "Nu"],
        ["Comparație 3 modele", "Da", "Nu", "Nu", "Nu", "Nu", "Nu"],
        ["Scor agregat multi-semnal", "Da", "Nu", "Nu", "Nu", "Nu", "Nu"],
        ["Cost", "Gratuit, open", "Freemium", "Gratuit", "Retras", "Plată", "Gratuit"],
        ["Transparență metodă", "Cod public", "Parțial", "Nu", "Parțial", "Nu", "Nu"],
    ]
    cw = [Inches(4.1), Inches(2.7), Inches(2.0), Inches(2.0), Inches(2.0), Inches(2.3), Inches(2.0)]
    table(sl, Inches(0.7), Inches(2.5), Inches(18.6), rows, cw, row_h=Inches(0.78),
          fsize=14, hsize=14)
    # highlight the "solutia propusa" column visually with a frame
    col_x = Inches(0.7) + Inches(4.1)
    frame = sl.shapes.add_shape(1, col_x, Inches(2.5), Inches(2.7), Inches(0.78)*8)
    frame.fill.background(); frame.line.color.rgb = BLUE; frame.line.width = Pt(3)
    frame.shadow.inherit = False


def s_contributii(prs, src):
    sl = clone_slide(src, 7, prs)
    header(sl, "Contribuții originale", 9)
    items = [
        ("Set de date dedicat", "6 surse de social media, focalizat pe texte scurte (50-500 car.)."),
        ("Comparație 3 modele", "NB, Reg. Logistică și RoBERTa în condiții identice."),
        ("Justificare Leave-One-Out", "Explicare la nivel de cuvânt — absentă în tool-urile comerciale."),
        ("Scor agregat de risc", "0-100, multi-semnal: RoBERTa 60% + fraze AI 25% + vocabular 15%."),
        ("Aplicație web integrată", "Analiză, comparație, NER, per-propoziție, batch CSV — open-source."),
        ("Validare pe domeniu", ">99.5% pe social media real, peste toate tool-urile comerciale."),
    ]
    for i, (t, d) in enumerate(items):
        col = i % 2; row = i // 3 if False else i % 3
        # 2 columns x 3 rows
        col = i // 3; row = i % 3
        x = Inches(1) + col * Inches(9.3)
        y = Inches(2.5) + row * Inches(2.6)
        rect(sl, x, y, Inches(8.6), Inches(2.2), PANEL)
        rect(sl, x, y, Inches(8.6), Inches(0.12), BLUE)
        rect(sl, x+Inches(0.3), y+Inches(0.45), Inches(0.9), Inches(0.9), BLUE)
        txb(sl, x+Inches(0.3), y+Inches(0.55), Inches(0.9), Inches(0.7), str(i+1),
            size=24, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        txb(sl, x+Inches(1.5), y+Inches(0.4), Inches(6.9), Inches(0.7), t, size=19, bold=True, color=BLUE)
        txb(sl, x+Inches(1.5), y+Inches(1.1), Inches(6.9), Inches(1.0), d, size=15, color=DARK)


def s_demo(prs, src):
    sl = clone_slide(src, 4, prs)
    header(sl, "Demo aplicație", 10,
           "Verdict + scor de risc · evidențiere cuvinte (Leave-One-Out) · comparație 3 modele · NER · batch CSV")
    pic(sl, f"{FIGS}/app_main_1.png", Inches(0.6), Inches(2.5), Inches(11.3), Inches(7.8))
    pic(sl, f"{FIGS}/app_compare_0.png", Inches(12.0), Inches(2.5), Inches(7.4), Inches(7.8))


def s_dezvoltari(prs, src):
    sl = clone_slide(src, 8, prs)
    header(sl, "Dezvoltări viitoare", 11)
    items = [
        ("Extindere multilingvă", "XLM-RoBERTa / seturi dedicate, inclusiv limba română."),
        ("Reantrenare periodică", "Pe modele noi (GPT-4o, Claude, Gemini) — concept drift."),
        ("Robustețe la evitare", "Testare împotriva parafrazării și „humanizării”."),
        ("Semnale comportamentale", "Metadate de rețea: frecvență postare, profil cont."),
        ("Infrastructură", "Extensie browser, API public REST, Docker."),
        ("Calibrare scor", "Optimizare automată a ponderilor scorului de risc."),
    ]
    for i, (t, d) in enumerate(items):
        col = i % 3; row = i // 3
        x = Inches(1) + col * Inches(6.2)
        y = Inches(2.7) + row * Inches(3.7)
        rect(sl, x, y, Inches(5.8), Inches(3.2), PANEL)
        rect(sl, x, y, Inches(5.8), Inches(0.12), BLUE)
        txb(sl, x+Inches(0.35), y+Inches(0.35), Inches(5.2), Inches(0.7), t, size=19, bold=True, color=BLUE)
        txb(sl, x+Inches(0.35), y+Inches(1.2), Inches(5.2), Inches(1.8), d, size=16, color=DARK)


def s_final(prs, src):
    sl = clone_slide(src, 14, prs, keep_decor=True)
    txb(sl, Inches(1), SH//2-Inches(2.0), Inches(18), Inches(1.5),
        "Vă mulțumesc pentru atenție!", size=46, bold=True, color=BLUE, align=PP_ALIGN.CENTER)
    txb(sl, Inches(1), SH//2-Inches(0.3), Inches(18), Inches(0.7),
        "Detectarea mesajelor generate de AI pe rețelele sociale",
        size=22, color=DARK, align=PP_ALIGN.CENTER)
    txb(sl, Inches(1), SH//2+Inches(0.9), Inches(18), Inches(0.6),
        "Bianca-Ștefania GHEORGHE · Automatică și Calculatoare · POLITEHNICA București · 2026",
        size=16, color=LGRAY, align=PP_ALIGN.CENTER)


def build():
    src = Presentation(TEMPLATE_PATH)
    prs = Presentation(); prs.slide_width = SW; prs.slide_height = SH
    steps = [s_title, s_cuprins, s_problema, s_obiective, s_tehnologii,
             s_arhitectura, s_rezultate, s_sota, s_contributii, s_demo,
             s_dezvoltari, s_final]
    for i, fn in enumerate(steps):
        fn(prs, src)
        print(f"  {i+1}. {fn.__name__} done")
    prs.save(OUTPUT_PATH)
    print(f"\nSaved: {OUTPUT_PATH}  ({len(prs.slides.__iter__.__self__._sldIdLst)} slides)")


if __name__ == "__main__":
    build()

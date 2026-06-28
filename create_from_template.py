# -*- coding: utf-8 -*-
"""
Apply content from existing presentation to ea5b5019 template design.
Adds a cuprins slide. Slide size: 20" x 11.25" (template native).
"""

import copy
import lxml.etree as etree
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn, nsmap

TEMPLATE_PATH = "/root/.claude/uploads/03ba0bac-fff2-5a16-859d-96a3dba91012/ea5b5019-Design_f_r__titlu.pptx"
OUTPUT_PATH = "/home/user/licenta/Prezentare_Licenta_Template.pptx"

BLUE = RGBColor(0x00, 0x4A, 0xAD)
DARK = RGBColor(0x30, 0x36, 0x42)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
LGRAY = RGBColor(0x60, 0x60, 0x60)
CYAN  = RGBColor(0x00, 0x7B, 0xCC)

# Slide dimensions (20" x 11.25")
SW = Inches(20)
SH = Inches(11.25)

FONT = "Calibri"
TOTAL = 12  # 11 content + cuprins


# ─── Slide clone helper ────────────────────────────────────────────────────────

def clone_slide(src_prs, slide_index, dst_prs):
    """Clone a slide from src into dst. Returns the new slide."""
    src_slide = src_prs.slides[slide_index]
    layout = dst_prs.slide_layouts[6]  # blank
    new_slide = dst_prs.slides.add_slide(layout)

    # Copy all shapes XML
    sp_tree = new_slide.shapes._spTree
    # Remove default shapes
    for child in list(sp_tree):
        sp_tree.remove(child)
    # Add back nvGrpSpPr and grpSpPr (required)
    src_tree = src_slide.shapes._spTree
    for child in src_tree:
        sp_tree.append(copy.deepcopy(child))

    # Copy relationships (images)
    from pptx.opc.packuri import PackURI
    src_part = src_slide.part
    dst_part = new_slide.part

    for rel in src_part.rels.values():
        if rel.is_external:
            continue
        if "/image" in rel.reltype:
            target_part = rel.target_part
            dst_part.relate_to(target_part, rel.reltype)

    return new_slide


def clear_text(slide):
    """Remove all text runs from text boxes on slide (keep shapes/images)."""
    for shape in slide.shapes:
        if shape.has_text_frame:
            for para in shape.text_frame.paragraphs:
                for run in para.runs:
                    run.text = ""


# ─── Text helper ──────────────────────────────────────────────────────────────

def txb(slide, left, top, width, height, text, size=18, bold=False,
        color=DARK, align=PP_ALIGN.LEFT, wrap=True, italic=False):
    """Add a text box."""
    tf_box = slide.shapes.add_textbox(left, top, width, height)
    tf = tf_box.text_frame
    tf.word_wrap = wrap
    tf.auto_size = None

    from pptx.util import Pt
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    run.font.name = FONT
    return tf_box


def txb_lines(slide, left, top, width, height, lines, size=14, color=DARK,
              bold=False, spacing=1.15, align=PP_ALIGN.LEFT):
    """Add multi-line text box from list of (text, bold, color, size) tuples or plain strings."""
    from pptx.util import Pt
    from pptx.oxml.ns import qn
    import lxml.etree as etree

    tf_box = slide.shapes.add_textbox(left, top, width, height)
    tf = tf_box.text_frame
    tf.word_wrap = True

    first = True
    for item in lines:
        if isinstance(item, str):
            t, b, c, s = item, bold, color, size
        else:
            t = item.get('text', '')
            b = item.get('bold', bold)
            c = item.get('color', color)
            s = item.get('size', size)

        if first:
            p = tf.paragraphs[0]
            first = False
        else:
            p = tf.add_paragraph()

        p.alignment = align
        # line spacing
        pPr = p._pPr
        if pPr is None:
            pPr = etree.SubElement(p._p, qn('a:pPr'))
        lnSpc = etree.SubElement(pPr, qn('a:lnSpc'))
        spcPct = etree.SubElement(lnSpc, qn('a:spcPct'))
        spcPct.set('val', str(int(spacing * 100000)))

        if not t:
            continue
        run = p.add_run()
        run.text = t
        run.font.size = Pt(s)
        run.font.bold = b
        run.font.color.rgb = c
        run.font.name = FONT

    return tf_box


def rect(slide, left, top, width, height, fill_color, alpha=None):
    """Add filled rectangle."""
    from pptx.util import Emu
    shape = slide.shapes.add_shape(
        1,  # MSO_SHAPE_TYPE.RECTANGLE
        left, top, width, height
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    shape.line.fill.background()
    return shape


def hline(slide, left, top, width, color=BLUE, thickness=3):
    """Add a horizontal line as thin rectangle."""
    return rect(slide, left, top, width, Pt(thickness), color)


# ─── Slide builders ───────────────────────────────────────────────────────────

def build_title(prs, src_prs):
    """Slide 1: Title page. Clone template slide 1."""
    slide = clone_slide(src_prs, 0, prs)

    # Remove all existing text boxes (keep background freeform + groups)
    to_remove = []
    for shape in slide.shapes:
        if shape.has_text_frame:
            to_remove.append(shape._element)
    for el in to_remove:
        el.getparent().remove(el)

    cx, cy = SW / 2, SH / 2

    # Subtitle line above title
    txb(slide, cx - Inches(5), cy - Inches(1.6), Inches(10), Inches(0.7),
        "Lucrare de licenta · 2024-2025",
        size=16, color=DARK, align=PP_ALIGN.CENTER)

    # Main title
    txb(slide, cx - Inches(6), cy - Inches(1.0), Inches(12), Inches(1.8),
        "Detectarea mesajelor generate de AI\npe rețelele sociale",
        size=36, bold=True, color=BLUE, align=PP_ALIGN.CENTER)

    # Author / coordinator
    txb_lines(slide, cx - Inches(5), cy + Inches(0.9), Inches(10), Inches(1.5),
        [
            {'text': 'Absolvent: Bianca-Stefania GHEORGHE', 'bold': False, 'size': 18, 'color': DARK},
            {'text': 'Coordonator: Conf.dr.ing. Daniel-Marian MEREZEANU', 'bold': False, 'size': 18, 'color': DARK},
        ], align=PP_ALIGN.CENTER)

    # University
    txb(slide, cx - Inches(6), cy + Inches(2.1), Inches(12), Inches(0.6),
        "UNIVERSITATEA POLITEHNICA BUCURESTI · FACULTATEA AUTOMATICA SI CALCULATOARE",
        size=13, color=LGRAY, align=PP_ALIGN.CENTER)

    return slide


def build_cuprins(prs, src_prs):
    """Slide 2: Table of contents. Clone template slide 4."""
    slide = clone_slide(src_prs, 3, prs)

    # Remove existing text boxes
    to_remove = []
    for shape in slide.shapes:
        if shape.has_text_frame:
            to_remove.append(shape._element)
    for el in to_remove:
        el.getparent().remove(el)

    # Title
    txb(slide, Inches(1), Inches(0.8), Inches(18), Inches(1.0),
        "CUPRINS", size=36, bold=True, color=BLUE, align=PP_ALIGN.CENTER)
    hline(slide, Inches(7), Inches(1.85), Inches(6), BLUE, 4)

    # Two columns of topics
    items_col1 = [
        "01.  Context si stadiul actual",
        "02.  Scopul lucrarii / Obiective",
        "03.  Tehnologii utilizate",
        "04.  Analiza cerintelor",
        "05.  Decizii de arhitectura",
        "06.  Rezultate obtinute",
    ]
    items_col2 = [
        "07.  Contributii originale",
        "08.  Dezvoltari viitoare",
        "09.  Bibliografie",
        "",
        "",
        "",
    ]

    for i, item in enumerate(items_col1):
        y = Inches(2.5) + i * Inches(1.2)
        rect(slide, Inches(1.5), y, Inches(0.12), Inches(0.7), BLUE)
        txb(slide, Inches(1.8), y + Inches(0.05), Inches(7.5), Inches(0.9),
            item, size=17, color=DARK)

    for i, item in enumerate(items_col2):
        if not item:
            continue
        y = Inches(2.5) + i * Inches(1.2)
        rect(slide, Inches(10.5), y, Inches(0.12), Inches(0.7), BLUE)
        txb(slide, Inches(10.8), y + Inches(0.05), Inches(7.5), Inches(0.9),
            item, size=17, color=DARK)

    return slide


def _slide_header(slide, title, slide_num):
    """Add standard header: title + slide number."""
    hline(slide, Inches(1), Inches(1.45), Inches(4), BLUE, 4)
    txb(slide, Inches(1), Inches(0.55), Inches(13), Inches(0.9),
        title, size=28, bold=True, color=BLUE)
    txb(slide, SW - Inches(2.5), SH - Inches(0.8), Inches(2), Inches(0.5),
        f"{slide_num}/{TOTAL}", size=13, color=LGRAY, align=PP_ALIGN.RIGHT)


def build_context(prs, src_prs):
    """Slide 3: Context si stadiul actual. Clone template slide 2 (left panel)."""
    slide = clone_slide(src_prs, 1, prs)

    to_remove = [s._element for s in slide.shapes if s.has_text_frame]
    for el in to_remove:
        el.getparent().remove(el)

    _slide_header(slide, "Context si stadiul actual", 3)

    # Left column: Problema
    txb(slide, Inches(1), Inches(1.9), Inches(8), Inches(0.6),
        "Problema", size=18, bold=True, color=BLUE)

    problema = [
        "Modelele LLM (ChatGPT, GPT-4, Gemini) genereaza texte aproape identice cu cele scrise de oameni [2].",
        "Retelele sociale: publicare instantanee, fara filtrare editoriala -> dezinformarea se propage rapid.",
        "Texte scurte si informale (abrevieri, argou, greseli intentionate) fac clasificarea dificila.",
        "Conform KPMG 2025 [6]: 65% din organizatii au intampinat riscuri din cauza AI.",
        "NewsGuard 2023 [12]: >1.000 surse de stiri false generate de AI.",
    ]
    txb_lines(slide, Inches(1), Inches(2.5), Inches(8.5), Inches(3.5),
        [{'text': f"• {t}", 'size': 15, 'color': DARK} for t in problema], spacing=1.4)

    # Right column: Stadiul actual
    txb(slide, Inches(10.5), Inches(1.9), Inches(8.5), Inches(0.6),
        "Stadiul actual al cercetarii (din 2019)", size=18, bold=True, color=BLUE)

    stadiu = [
        ("Metode statistice", "Incertitudine, entropie — GLTR [5]", BLUE),
        ("Watermarking", "Semnal ascuns la generare [8]", BLUE),
        ("Clasificatori supervizati", "RoBERTa fine-tuned [10]", BLUE),
        ("Metode zero-shot", "DetectGPT [11]", BLUE),
    ]
    for i, (title_s, desc, _) in enumerate(stadiu):
        y = Inches(2.5) + i * Inches(1.6)
        rect(slide, Inches(10.5), y, Inches(0.08), Inches(1.2), BLUE)
        txb(slide, Inches(10.8), y, Inches(8), Inches(0.55),
            title_s, size=16, bold=True, color=BLUE)
        txb(slide, Inches(10.8), y + Inches(0.55), Inches(8), Inches(0.65),
            desc, size=14, color=DARK)

    # Limitare
    rect(slide, Inches(1), Inches(6.2), Inches(18), Inches(1.4), RGBColor(0xE8, 0xF0, 0xFB))
    txb(slide, Inches(1.3), Inches(6.4), Inches(17.4), Inches(1.0),
        "Limitare comuna: metodele existente sunt evaluate pe texte lungi si formale; parafrazarea reduce acuratetea de la >90% la <50% [9].",
        size=15, italic=True, color=DARK)

    return slide


def build_obiective(prs, src_prs):
    """Slide 4: Scopul lucrarii / Obiective. Clone template slide 3."""
    slide = clone_slide(src_prs, 2, prs)

    to_remove = [s._element for s in slide.shapes if s.has_text_frame]
    for el in to_remove:
        el.getparent().remove(el)

    _slide_header(slide, "Scopul lucrarii / Obiective", 4)

    items = [
        ("1", "Analiza critica",
         "Studierea metodelor existente de detectare si a limitarilor lor pe retelele sociale."),
        ("2", "Solutie specializata",
         "Detector bazat pe machine learning, antrenat specific pe texte scurte de social media (50-500 caractere)."),
        ("3", "Comparatie de modele",
         "Evaluarea sistematica a 3 clasificatori complementari pe acelasi set de date."),
        ("4", "Aplicatie accesibila",
         "Interfata web cu justificarea transparenta a deciziilor, nu doar un scor final."),
    ]

    cols = [(Inches(1), Inches(9.5)), (Inches(10.5), Inches(9.5))]
    for i, (num, title_s, desc) in enumerate(items):
        col_x, col_w = cols[i % 2]
        row = i // 2
        y = Inches(2.1) + row * Inches(3.5)

        # Number circle
        rect(slide, col_x, y, Inches(1.0), Inches(1.0), BLUE)
        txb(slide, col_x, y, Inches(1.0), Inches(1.0),
            num, size=28, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

        txb(slide, col_x + Inches(1.2), y + Inches(0.1), col_w - Inches(1.4), Inches(0.65),
            title_s, size=19, bold=True, color=BLUE)
        txb(slide, col_x + Inches(1.2), y + Inches(0.75), col_w - Inches(1.4), Inches(2.4),
            desc, size=15, color=DARK)

    return slide


def build_tehnologii(prs, src_prs):
    """Slide 5: Tehnologii utilizate. Clone template slide 8."""
    slide = clone_slide(src_prs, 7, prs)

    to_remove = [s._element for s in slide.shapes if s.has_text_frame]
    for el in to_remove:
        el.getparent().remove(el)

    _slide_header(slide, "Tehnologii utilizate", 5)

    categories = [
        ("Date & Prelucrare",
         "Python 3.10 · pandas (filtrare, echilibrare) · datasets HuggingFace (streaming)"),
        ("ML Clasic",
         "scikit-learn — Naive Bayes, Reg. Logistica · TF-IDF (10.000 vocab, n-grame 1-2) · NLTK stop-words"),
        ("Deep Learning / NLP",
         "PyTorch 2.x (GPU) · Transformers — RoBERTa-base [10] [15] · SpaCy NER"),
        ("Aplicatie web",
         "Flask 3.x — API REST · HTML/CSS/JS · Chart.js vizualizari"),
    ]

    xs = [Inches(1), Inches(6), Inches(11), Inches(16)]
    yw = Inches(2.2)
    for i, (cat, desc) in enumerate(categories):
        x = xs[i]
        rect(slide, x, yw - Inches(0.1), Inches(4.5), Inches(0.08), BLUE)
        txb(slide, x, yw, Inches(4.5), Inches(0.7),
            cat, size=17, bold=True, color=BLUE)
        txb(slide, x, yw + Inches(0.75), Inches(4.5), Inches(2.5),
            desc, size=14, color=DARK)

    # Mediu de antrenare box
    rect(slide, Inches(1), Inches(5.8), Inches(18), Inches(1.8), RGBColor(0xE8, 0xF0, 0xFB))
    txb(slide, Inches(1.2), Inches(5.95), Inches(3), Inches(0.7),
        "Mediu de antrenare", size=16, bold=True, color=BLUE)
    txb(slide, Inches(1.2), Inches(6.65), Inches(17.6), Inches(0.7),
        "Google Colaboratory · GPU NVIDIA Tesla T4 (15.6 GB VRAM) · precizie mixta FP16 · antrenare RoBERTa ~85 min (3 epoci) · modele clasice <2 min · inferenta locala CPU (0.5-2 s/text)",
        size=14, color=DARK)

    return slide


def build_analiza(prs, src_prs):
    """Slide 6: Analiza cerintelor. Clone template slide 2."""
    slide = clone_slide(src_prs, 1, prs)

    to_remove = [s._element for s in slide.shapes if s.has_text_frame]
    for el in to_remove:
        el.getparent().remove(el)

    _slide_header(slide, "Analiza cerintelor", 6)

    items = [
        ("Instrumente existente",
         "5 aplicatii comerciale studiate (GPTZero [14], ZeroGPT, OpenAI Classifier, Originality.ai, Sapling) - toate slabe pe texte scurte."),
        ("Metode din literatura",
         "4 categorii identificate: statistice [5], watermarking [8], supervizate, zero-shot [11] si limitarile lor [13]."),
        ("Specificul social media",
         "Lungime mica, limbaj informal, abrevieri si greseli intentionate ingreuneaza clasificarea."),
        ("Concluzii -> specificatii",
         "Set de date 50-500 caractere · comparatie multi-model · justificare per cuvant · procesare in lot."),
    ]

    for i, (title_s, desc) in enumerate(items):
        x = Inches(1) if i % 2 == 0 else Inches(10.5)
        y = Inches(2.0) + (i // 2) * Inches(3.8)
        rect(slide, x, y, Inches(0.12), Inches(3.0), BLUE)
        txb(slide, x + Inches(0.3), y, Inches(8.5), Inches(0.7),
            title_s, size=18, bold=True, color=BLUE)
        txb(slide, x + Inches(0.3), y + Inches(0.75), Inches(8.5), Inches(2.0),
            desc, size=15, color=DARK)

    return slide


def build_arhitectura(prs, src_prs):
    """Slide 7: Decizii de arhitectura si model de date. Clone template slide 7."""
    slide = clone_slide(src_prs, 6, prs)

    to_remove = [s._element for s in slide.shapes if s.has_text_frame]
    for el in to_remove:
        el.getparent().remove(el)

    _slide_header(slide, "Decizii de arhitectura si model de date", 7)

    # Left: arhitectura
    txb(slide, Inches(1), Inches(1.9), Inches(8.5), Inches(0.6),
        "Arhitectura software: client-server", size=18, bold=True, color=BLUE)
    txb_lines(slide, Inches(1), Inches(2.55), Inches(8.5), Inches(2.5),
        [{'text': t, 'size': 15, 'color': DARK} for t in [
            "• Backend Flask: incarca cele 3 modele la pornire; endpoint-uri /predict si /batch.",
            "• Frontend HTML/CSS/JS: vizualizari, comutare model in browser fara re-apel server.",
            "• Trei modele complementare: Naive Bayes & Reg. Logistica (TF-IDF) + RoBERTa fine-tuned (125M parametri) [10] bazat pe Transformer [15].",
        ]], spacing=1.5)

    # Left: 6 surse
    txb(slide, Inches(1), Inches(5.1), Inches(8.5), Inches(0.6),
        "Model de date: 6 surse combinate", size=18, bold=True, color=BLUE)

    surse = [
        ("HC3 Reddit ELI5 [7]", "perechi uman/AI"),
        ("TweetEval [1]", "tweets reale"),
        ("Reddit comentarii", "limbaj informal"),
        ("RAID 2024 [4]", "AI multi-model"),
        ("AI_Human.csv", "volum/diversitate"),
        ("AI Detection", "texte AI extra"),
    ]
    for i, (src, desc) in enumerate(surse):
        col = i % 3
        row = i // 3
        x = Inches(1) + col * Inches(2.9)
        y = Inches(5.8) + row * Inches(1.1)
        rect(slide, x, y, Inches(0.08), Inches(0.9), BLUE)
        txb(slide, x + Inches(0.2), y, Inches(2.5), Inches(0.5),
            src, size=13, bold=True, color=DARK)
        txb(slide, x + Inches(0.2), y + Inches(0.45), Inches(2.5), Inches(0.45),
            desc, size=12, color=LGRAY)

    # Right: stat box
    rect(slide, Inches(10.5), Inches(1.9), Inches(8.5), Inches(5.0), RGBColor(0xE8, 0xF0, 0xFB))
    txb(slide, Inches(10.8), Inches(2.1), Inches(8), Inches(0.7),
        "45.834 exemple", size=36, bold=True, color=BLUE, align=PP_ALIGN.CENTER)
    txb(slide, Inches(10.8), Inches(2.9), Inches(8), Inches(0.5),
        "echilibrate 50/50 (uman/AI)", size=18, color=DARK, align=PP_ALIGN.CENTER)
    hline(slide, Inches(11.5), Inches(3.55), Inches(6.5), BLUE, 2)
    txb_lines(slide, Inches(10.8), Inches(3.7), Inches(8), Inches(2.5),
        [{'text': t, 'size': 15, 'color': DARK, 'bold': False} for t in [
            "• Filtrate: 50-500 caractere",
            "• Impartire stratificata 70 / 15 / 15",
            "• Surse: Twitter, Reddit, HC3, RAID, AI_Human, AI Detection",
        ]], spacing=1.5)

    return slide


def build_rezultate(prs, src_prs):
    """Slide 8: Rezultate obtinute. Clone template slide 4."""
    slide = clone_slide(src_prs, 3, prs)

    to_remove = [s._element for s in slide.shapes if s.has_text_frame]
    for el in to_remove:
        el.getparent().remove(el)

    _slide_header(slide, "Rezultate obtinute", 8)

    # 3 model result cards
    models = [
        ("Naive Bayes", "94.24%", "94.17%", "F1"),
        ("Reg. Logistica", "96.71%", "96.66%", "F1"),
        ("RoBERTa", "99.17%", "99.18%", "F1"),
    ]
    colors_m = [RGBColor(0x72, 0x9B, 0xCB), BLUE, RGBColor(0x00, 0x2D, 0x7A)]

    for i, (name, acc, f1, _) in enumerate(models):
        x = Inches(1) + i * Inches(6.3)
        y = Inches(2.0)
        rect(slide, x, y, Inches(5.8), Inches(3.8), colors_m[i])
        txb(slide, x + Inches(0.3), y + Inches(0.3), Inches(5.2), Inches(0.7),
            name, size=20, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        txb(slide, x + Inches(0.3), y + Inches(1.1), Inches(5.2), Inches(1.2),
            acc, size=40, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        txb(slide, x + Inches(0.3), y + Inches(2.5), Inches(5.2), Inches(0.7),
            f"F1-score: {f1}", size=17, color=RGBColor(0xDD, 0xEE, 0xFF),
            align=PP_ALIGN.CENTER)

    # Details
    txb(slide, Inches(1), Inches(6.0), Inches(18), Inches(0.7),
        "Testat pe 6.876 exemple (50/50 echilibrat). Pe surse propriu-zise de social media: acuratete >99.5% - confirma teoria detectabilitatii [3].",
        size=15, color=DARK)

    txb(slide, Inches(1), Inches(6.9), Inches(8.5), Inches(0.5),
        "Demo aplicatie:", size=16, bold=True, color=BLUE)
    txb_lines(slide, Inches(1), Inches(7.5), Inches(8.5), Inches(2.0),
        [{'text': t, 'size': 14, 'color': DARK} for t in [
            "• Verdict + scor agregat de risc (0-100)",
            "• Evidentiere cuvinte (Leave-One-Out)",
            "• Comparatie 3 modele simultan · Procesare lot CSV",
        ]], spacing=1.4)

    txb(slide, Inches(10.5), Inches(6.9), Inches(8.5), Inches(0.5),
        "Matricea de confuzie RoBERTa:", size=16, bold=True, color=BLUE)
    txb_lines(slide, Inches(10.5), Inches(7.5), Inches(8.5), Inches(2.0),
        [{'text': t, 'size': 14, 'color': DARK} for t in [
            "• 3.388 uman corect · 3.431 AI corect",
            "• 50 fals-pozitive · 7 fals-negative",
            "• Aproape toate erorile vin din texte formale, nu social media.",
        ]], spacing=1.4)

    return slide


def build_contributii(prs, src_prs):
    """Slide 9: Contributii originale. Clone template slide 9."""
    slide = clone_slide(src_prs, 8, prs)

    to_remove = [s._element for s in slide.shapes if s.has_text_frame]
    for el in to_remove:
        el.getparent().remove(el)

    _slide_header(slide, "Contributii originale", 9)

    items = [
        ("1", "Set de date dedicat",
         "6 surse de social media combinate, focalizat pe texte scurte (50-500 car.)."),
        ("2", "Comparatie 3 modele",
         "Naive Bayes, Reg. Logistica si RoBERTa evaluate in conditii identice."),
        ("3", "Justificare Leave-One-Out",
         "Explicare la nivel de cuvant - absenta in instrumentele comerciale."),
        ("4", "Scor agregat de risc",
         "Indicator 0-100 multi-semnal (RoBERTa 60% + fraze AI 25% + statistici vocabular 15%)."),
        ("5", "Aplicatie web integrata",
         "Analiza, comparatie, per-propozitie, NER, batch CSV - open-source."),
    ]

    for i, (num, title_s, desc) in enumerate(items):
        col = i % 2
        row = i // 2
        x = Inches(1) if col == 0 else Inches(10.5)
        y = Inches(2.0) + row * Inches(2.9)
        if i == 4:  # last one centered
            x = Inches(5.75)

        rect(slide, x, y, Inches(0.9), Inches(0.9), BLUE)
        txb(slide, x, y, Inches(0.9), Inches(0.9),
            num, size=22, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        txb(slide, x + Inches(1.1), y, Inches(8), Inches(0.6),
            title_s, size=17, bold=True, color=BLUE)
        txb(slide, x + Inches(1.1), y + Inches(0.65), Inches(8), Inches(1.8),
            desc, size=14, color=DARK)

    return slide


def build_dezvoltari(prs, src_prs):
    """Slide 10: Dezvoltari viitoare. Clone template slide 8."""
    slide = clone_slide(src_prs, 7, prs)

    to_remove = [s._element for s in slide.shapes if s.has_text_frame]
    for el in to_remove:
        el.getparent().remove(el)

    _slide_header(slide, "Dezvoltari viitoare", 10)

    items = [
        ("Extindere multilingva", "XLM-RoBERTa sau seturi de date dedicate, inclusiv pentru limba romana."),
        ("Reantrenare periodica", "Pe modele noi (GPT-4o, Claude, Gemini) - combaterea 'concept drift'."),
        ("Robustete la evitare", "Testare sistematica impotriva parafrazarii si 'humanizarii' textului."),
        ("Semnale comportamentale", "Integrarea metadatelor de retea (frecventa postare, profil cont)."),
        ("Infrastructura", "Extensie de browser, API public REST, containerizare Docker."),
        ("Calibrare scor", "Optimizarea automata a ponderilor scorului agregat de risc."),
    ]

    for i, (title_s, desc) in enumerate(items):
        col = i % 3
        row = i // 3
        x = Inches(1) + col * Inches(6.3)
        y = Inches(2.2) + row * Inches(3.5)
        rect(slide, x, y, Inches(5.8), Inches(3.0), RGBColor(0xE8, 0xF0, 0xFB))
        rect(slide, x, y, Inches(5.8), Inches(0.12), BLUE)
        txb(slide, x + Inches(0.2), y + Inches(0.25), Inches(5.4), Inches(0.65),
            title_s, size=17, bold=True, color=BLUE)
        txb(slide, x + Inches(0.2), y + Inches(0.9), Inches(5.4), Inches(1.9),
            desc, size=14, color=DARK)

    return slide


def build_bibliografie(prs, src_prs):
    """Slide 11: Bibliografie. Clone template slide 3."""
    slide = clone_slide(src_prs, 2, prs)

    to_remove = [s._element for s in slide.shapes if s.has_text_frame]
    for el in to_remove:
        el.getparent().remove(el)

    txb(slide, Inches(1), Inches(0.4), Inches(18), Inches(0.9),
        "Bibliografie", size=28, bold=True, color=BLUE)
    hline(slide, Inches(1), Inches(1.35), Inches(5), BLUE, 4)
    txb(slide, SW - Inches(2.5), SH - Inches(0.8), Inches(2), Inches(0.5),
        f"11/{TOTAL}", size=13, color=LGRAY, align=PP_ALIGN.RIGHT)

    refs_col1 = [
        "[1] Barbieri, F., Camacho-Collados, J., Espinosa-Anke, L., & Neves, L. (2020). TweetEval. arXiv:2010.12421.",
        "[2] Brown, T., Mann, B., Ryder, N., Subbiah, M., Kaplan, J., Dhariwal, P., et al. (2020). GPT-3. arXiv:2005.14165.",
        "[3] Chakraborty, S., Bedi, A.S., Zhu, S., An, B., Manocha, D., & Huang, F. (2023). arXiv:2304.04736.",
        "[4] Dugan, L., Hwang, A., Trhlik, F., Zhu, Z., Ippolito, D., & Callison-Burch, C. (2024). RAID. arXiv:2405.07940.",
        "[5] Gehrmann, S., Strobelt, H., & Rush, A.M. (2019). GLTR. arXiv:1906.04043.",
        "[6] Gillespie, N., Lockey, S., Curtis, C., Pool, J., & Akbari, A. (2025). KPMG & Univ. Melbourne.",
        "[7] Guo, W., Shen, W., Lei, J., Chow, K., & Shi, E. (2023). HC3. arXiv:2301.07597.",
        "[8] Kirchenbauer, J., Geiping, J., Wen, Y., Katz, J., Miers, I., & Goldstein, T. (2023). arXiv:2301.10226.",
    ]
    refs_col2 = [
        "[9] Krishna, K., Song, Y., Karpinska, M., Wieting, J., & Iyyer, M. (2023). arXiv:2303.13408.",
        "[10] Liu, Y., Ott, M., Goyal, N., Du, J., Joshi, M., Chen, D., et al. (2019). RoBERTa. arXiv:1907.11692.",
        "[11] Mitchell, E., Lee, Y., Khazatsky, A., Manning, C.D., & Finn, C. (2023). DetectGPT. arXiv:2301.11305.",
        "[12] NewsGuard (2023). The Year AI Supercharged Misinformation.",
        "[13] Sadasivan, V.S., Kumar, A., Balasubramanian, S., Wang, W., & Feizi, S. (2023). arXiv:2303.11156.",
        "[14] Tian, E. (2023). GPTZero. gptzero.me.",
        "[15] Vaswani, A., Shazeer, N., Parmar, N., Uszkoreit, J., Jones, L., Gomez, A.N., et al. (2017). Transformer. arXiv:1706.03762.",
    ]

    for i, ref in enumerate(refs_col1):
        txb(slide, Inches(1), Inches(1.6) + i * Inches(1.15), Inches(9), Inches(1.0),
            ref, size=12, color=DARK)

    for i, ref in enumerate(refs_col2):
        txb(slide, Inches(10.5), Inches(1.6) + i * Inches(1.15), Inches(9), Inches(1.0),
            ref, size=12, color=DARK)

    return slide


def build_final(prs, src_prs):
    """Slide 12: Final / Thank you. Clone template slide 15."""
    slide = clone_slide(src_prs, 14, prs)

    to_remove = [s._element for s in slide.shapes if s.has_text_frame]
    for el in to_remove:
        el.getparent().remove(el)

    txb(slide, Inches(1), SH / 2 - Inches(1.8), Inches(18), Inches(1.4),
        "Va multumesc pentru atentie!", size=42, bold=True, color=BLUE,
        align=PP_ALIGN.CENTER)

    txb(slide, Inches(1), SH / 2 - Inches(0.2), Inches(18), Inches(0.7),
        "Detectarea mesajelor generate de AI pe retelele sociale",
        size=20, color=DARK, align=PP_ALIGN.CENTER)

    txb(slide, Inches(1), SH / 2 + Inches(0.8), Inches(18), Inches(0.6),
        "Bianca-Stefania GHEORGHE · Facultatea Automatica si Calculatoare · UPB · 2025",
        size=16, color=LGRAY, align=PP_ALIGN.CENTER)

    return slide


# ─── Main build ───────────────────────────────────────────────────────────────

def build():
    src_prs = Presentation(TEMPLATE_PATH)

    dst_prs = Presentation()
    dst_prs.slide_width = SW
    dst_prs.slide_height = SH

    print("Building slides...")
    build_title(dst_prs, src_prs)
    print("  1. Title done")
    build_cuprins(dst_prs, src_prs)
    print("  2. Cuprins done")
    build_context(dst_prs, src_prs)
    print("  3. Context done")
    build_obiective(dst_prs, src_prs)
    print("  4. Obiective done")
    build_tehnologii(dst_prs, src_prs)
    print("  5. Tehnologii done")
    build_analiza(dst_prs, src_prs)
    print("  6. Analiza done")
    build_arhitectura(dst_prs, src_prs)
    print("  7. Arhitectura done")
    build_rezultate(dst_prs, src_prs)
    print("  8. Rezultate done")
    build_contributii(dst_prs, src_prs)
    print("  9. Contributii done")
    build_dezvoltari(dst_prs, src_prs)
    print("  10. Dezvoltari done")
    build_bibliografie(dst_prs, src_prs)
    print("  11. Bibliografie done")
    build_final(dst_prs, src_prs)
    print("  12. Final done")

    dst_prs.save(OUTPUT_PATH)
    print(f"\nSaved: {OUTPUT_PATH}")


if __name__ == "__main__":
    build()

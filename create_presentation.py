#!/usr/bin/env python3
"""
Creates a complete license defense presentation for Bianca-Stefania Gheorghe
based on the AI Technology template and thesis content.
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu, Cm
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from pptx.oxml import parse_xml
import lxml.etree as etree
import copy

# ─── Color palette (inspired by the AI template dark theme) ───────────────────
BG_DARK     = RGBColor(0x0A, 0x0A, 0x1A)   # very dark navy (slide bg)
BG_CARD     = RGBColor(0x12, 0x12, 0x2E)   # slightly lighter navy (cards)
BG_HEADER   = RGBColor(0x0D, 0x0D, 0x26)   # dark header bar
ACCENT_BLUE = RGBColor(0x00, 0x8B, 0xFF)   # bright blue (headings / borders)
ACCENT_CYAN = RGBColor(0x00, 0xD4, 0xFF)   # cyan (sub-headings / highlight)
ACCENT_PURPLE = RGBColor(0x8B, 0x5C, 0xF6) # purple accent
TEXT_WHITE  = RGBColor(0xFF, 0xFF, 0xFF)   # white (body text)
TEXT_LIGHT  = RGBColor(0xCC, 0xDD, 0xFF)   # light blue-white (secondary text)
TEXT_GREY   = RGBColor(0x88, 0x99, 0xBB)   # grey (footnotes / labels)
ACCENT_GREEN = RGBColor(0x00, 0xE6, 0x76)  # green (good results)
ACCENT_ORANGE = RGBColor(0xFF, 0x8C, 0x00) # orange (warnings)
ACCENT_RED  = RGBColor(0xFF, 0x3B, 0x30)   # red (attention)

# ─── Slide dimensions (widescreen 16:9) ───────────────────────────────────────
W = Inches(13.33)
H = Inches(7.50)


def new_prs():
    prs = Presentation()
    prs.slide_width  = W
    prs.slide_height = H
    return prs


# ─── Helpers ──────────────────────────────────────────────────────────────────

def add_slide(prs, layout_idx=6):
    """Add blank slide (layout 6 = Blank)."""
    layout = prs.slide_layouts[layout_idx]
    return prs.slides.add_slide(layout)


def set_bg(slide, color: RGBColor):
    """Set solid background color on slide."""
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = color


def add_rect(slide, x, y, w, h, fill_color: RGBColor = None,
             line_color: RGBColor = None, line_width_pt: float = 0):
    """Add a rectangle shape."""
    shape = slide.shapes.add_shape(
        1,  # MSO_SHAPE_TYPE.RECTANGLE
        Inches(x), Inches(y), Inches(w), Inches(h)
    )
    if fill_color:
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill_color
    else:
        shape.fill.background()
    if line_color:
        shape.line.color.rgb = line_color
        shape.line.width = Pt(line_width_pt)
    else:
        shape.line.fill.background()
    return shape


def add_textbox(slide, x, y, w, h, text: str,
                font_size=18, bold=False, italic=False,
                color: RGBColor = None, align=PP_ALIGN.LEFT,
                word_wrap=True, font_name="Calibri"):
    """Add a textbox."""
    txBox = slide.shapes.add_textbox(
        Inches(x), Inches(y), Inches(w), Inches(h)
    )
    tf = txBox.text_frame
    tf.word_wrap = word_wrap
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.name = font_name
    if color:
        run.font.color.rgb = color
    return txBox


def add_textbox_multi(slide, x, y, w, h, paragraphs: list,
                      default_font_size=16, default_color=None,
                      word_wrap=True, font_name="Calibri", line_spacing=1.15):
    """
    paragraphs: list of dicts with keys:
        text, font_size, bold, italic, color, align, space_before
    """
    txBox = slide.shapes.add_textbox(
        Inches(x), Inches(y), Inches(w), Inches(h)
    )
    tf = txBox.text_frame
    tf.word_wrap = word_wrap

    for i, para in enumerate(paragraphs):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        p.alignment = para.get('align', PP_ALIGN.LEFT)

        # Space before
        from pptx.util import Pt as _Pt
        from pptx.oxml.ns import qn as _qn
        if para.get('space_before', 0):
            pPr = p._p.get_or_add_pPr()
            spcBef = etree.SubElement(pPr, _qn('a:spcBef'))
            spcPts = etree.SubElement(spcBef, _qn('a:spcPts'))
            spcPts.set('val', str(int(para['space_before'] * 100)))

        text = para.get('text', '')
        if not text:
            # Empty paragraph for spacing
            run = p.add_run()
            run.text = ''
            continue
        run = p.add_run()
        run.text = text
        run.font.size = Pt(para.get('font_size', default_font_size))
        run.font.bold = para.get('bold', False)
        run.font.italic = para.get('italic', False)
        run.font.name = para.get('font_name', font_name)
        c = para.get('color', default_color)
        if c:
            run.font.color.rgb = c

    return txBox


def add_slide_number(slide, num, total, color=TEXT_GREY):
    """Add slide number bottom right."""
    add_textbox(slide, 12.3, 7.1, 0.9, 0.3, f"{num}/{total}",
                font_size=11, color=color, align=PP_ALIGN.RIGHT)


def add_footer(slide, title_text, color=TEXT_GREY):
    """Add footer with title bottom left."""
    add_textbox(slide, 0.2, 7.1, 8.0, 0.3, title_text,
                font_size=10, color=color, italic=True)


def add_top_bar(slide, title: str, subtitle: str = ""):
    """Dark top bar with title."""
    add_rect(slide, 0, 0, 13.33, 1.25, fill_color=BG_HEADER)
    # Blue accent line
    add_rect(slide, 0, 1.20, 13.33, 0.05, fill_color=ACCENT_BLUE)
    # Title text
    add_textbox(slide, 0.4, 0.08, 10, 0.65, title,
                font_size=28, bold=True, color=TEXT_WHITE,
                align=PP_ALIGN.LEFT, font_name="Calibri")
    if subtitle:
        add_textbox(slide, 0.4, 0.72, 10, 0.45, subtitle,
                    font_size=15, color=ACCENT_CYAN, italic=True,
                    align=PP_ALIGN.LEFT, font_name="Calibri")


def add_bullet_card(slide, x, y, w, h, title: str, bullets: list,
                    title_color=ACCENT_CYAN, bullet_color=TEXT_WHITE,
                    bullet_size=15, title_size=17):
    """Card with title and bullet points."""
    add_rect(slide, x, y, w, h, fill_color=BG_CARD,
             line_color=ACCENT_BLUE, line_width_pt=1.0)
    # Title
    add_textbox(slide, x+0.15, y+0.1, w-0.3, 0.35, title,
                font_size=title_size, bold=True, color=title_color,
                font_name="Calibri")
    # Separator
    add_rect(slide, x+0.15, y+0.48, w-0.3, 0.02, fill_color=ACCENT_BLUE)
    # Bullets
    if bullets:
        bullet_text = "\n".join([f"• {b}" for b in bullets])
        add_textbox(slide, x+0.15, y+0.55, w-0.3, h-0.65, bullet_text,
                    font_size=bullet_size, color=bullet_color,
                    word_wrap=True, font_name="Calibri")


def add_metric_box(slide, x, y, w, h, value: str, label: str,
                   value_color=ACCENT_GREEN):
    """Big number metric box."""
    add_rect(slide, x, y, w, h, fill_color=BG_CARD,
             line_color=ACCENT_BLUE, line_width_pt=1.0)
    add_textbox(slide, x+0.1, y+0.12, w-0.2, h*0.55, value,
                font_size=36, bold=True, color=value_color,
                align=PP_ALIGN.CENTER, font_name="Calibri")
    add_textbox(slide, x+0.1, y+h*0.58, w-0.2, h*0.35, label,
                font_size=13, color=TEXT_LIGHT,
                align=PP_ALIGN.CENTER, word_wrap=True, font_name="Calibri")


# ─── SLIDE BUILDERS ───────────────────────────────────────────────────────────

PRESENTATION_TITLE = "Detectarea mesajelor generate de AI pe rețelele sociale"
TOTAL_SLIDES = 18


def slide_title(prs):
    """Slide 1 – Title."""
    slide = add_slide(prs)
    set_bg(slide, BG_DARK)

    # Big gradient-like rectangle at top
    add_rect(slide, 0, 0, 13.33, 3.8, fill_color=BG_HEADER)
    # Decorative blue accent bar
    add_rect(slide, 0, 3.75, 13.33, 0.06, fill_color=ACCENT_BLUE)
    # Decorative cyan thin line
    add_rect(slide, 0, 3.81, 13.33, 0.02, fill_color=ACCENT_CYAN)

    # Institution
    add_textbox(slide, 0.5, 0.18, 12.3, 0.45,
                "Universitatea Națională de Știință și Tehnologie POLITEHNICA București",
                font_size=13, color=TEXT_GREY, align=PP_ALIGN.CENTER,
                font_name="Calibri")
    add_textbox(slide, 0.5, 0.55, 12.3, 0.40,
                "Facultatea Automatică și Calculatoare · Departamentul Automatică și Informatică Industrială",
                font_size=12, color=TEXT_GREY, align=PP_ALIGN.CENTER,
                font_name="Calibri")

    # Diploma label
    add_textbox(slide, 0.5, 1.05, 12.3, 0.55,
                "LUCRARE DE DIPLOMĂ",
                font_size=17, bold=True, color=ACCENT_CYAN,
                align=PP_ALIGN.CENTER, font_name="Calibri")

    # Main title
    add_textbox(slide, 0.5, 1.55, 12.3, 1.8,
                PRESENTATION_TITLE,
                font_size=40, bold=True, color=TEXT_WHITE,
                align=PP_ALIGN.CENTER, font_name="Calibri")

    # Bottom info area
    add_textbox(slide, 1.5, 4.3, 4.5, 0.5,
                "Absolvent:",
                font_size=13, color=TEXT_GREY, font_name="Calibri")
    add_textbox(slide, 1.5, 4.72, 4.5, 0.5,
                "Bianca-Ștefania GHEORGHE",
                font_size=17, bold=True, color=TEXT_WHITE, font_name="Calibri")

    add_textbox(slide, 7.3, 4.3, 5.5, 0.5,
                "Coordonator:",
                font_size=13, color=TEXT_GREY, font_name="Calibri")
    add_textbox(slide, 7.3, 4.72, 5.5, 0.5,
                "Conf.dr.ing. Daniel-Marian MEREZEANU",
                font_size=17, bold=True, color=TEXT_WHITE, font_name="Calibri")

    # Separator vertical
    add_rect(slide, 6.8, 4.2, 0.03, 1.3, fill_color=ACCENT_BLUE)

    add_textbox(slide, 0.5, 5.55, 12.3, 0.5,
                "Sesiunea Iunie 2026",
                font_size=14, color=ACCENT_BLUE,
                align=PP_ALIGN.CENTER, font_name="Calibri")

    # Decorative corner boxes
    add_rect(slide, 0, 0, 0.08, 0.5, fill_color=ACCENT_BLUE)
    add_rect(slide, 0, 0, 0.5, 0.08, fill_color=ACCENT_BLUE)
    add_rect(slide, 12.87, 0, 0.08, 0.5, fill_color=ACCENT_CYAN)
    add_rect(slide, 12.85, 0, 0.5, 0.08, fill_color=ACCENT_CYAN)


def slide_cuprins(prs):
    """Slide 2 – Cuprins."""
    slide = add_slide(prs)
    set_bg(slide, BG_DARK)
    add_top_bar(slide, "CUPRINS", "Structura prezentării")
    add_footer(slide, PRESENTATION_TITLE)
    add_slide_number(slide, 2, TOTAL_SLIDES)

    items = [
        ("01", "Contextul și motivația temei"),
        ("02", "Procesarea limbajului natural și modelele LLM"),
        ("03", "Metode de detectare a textului generat de AI"),
        ("04", "Aplicații similare existente"),
        ("05", "Metoda propusă – arhitectura soluției"),
        ("06", "Setul de date utilizat"),
        ("07", "Implementare tehnică"),
        ("08", "Rezultate obținute"),
        ("09", "Comparație cu instrumente comerciale"),
        ("10", "Contribuții originale"),
        ("11", "Domenii de utilizare & Dezvoltări ulterioare"),
        ("12", "Concluzii"),
        ("13", "Bibliografie"),
    ]

    col1_items = items[:7]
    col2_items = items[7:]

    for i, (num, label) in enumerate(col1_items):
        y = 1.45 + i * 0.73
        add_rect(slide, 0.5, y, 0.55, 0.55, fill_color=ACCENT_BLUE)
        add_textbox(slide, 0.52, y+0.05, 0.52, 0.45, num,
                    font_size=18, bold=True, color=TEXT_WHITE,
                    align=PP_ALIGN.CENTER)
        add_textbox(slide, 1.15, y+0.08, 5.6, 0.42, label,
                    font_size=15, color=TEXT_LIGHT)

    for i, (num, label) in enumerate(col2_items):
        y = 1.45 + i * 0.73
        add_rect(slide, 6.9, y, 0.55, 0.55, fill_color=ACCENT_PURPLE)
        add_textbox(slide, 6.92, y+0.05, 0.52, 0.45, num,
                    font_size=18, bold=True, color=TEXT_WHITE,
                    align=PP_ALIGN.CENTER)
        add_textbox(slide, 7.55, y+0.08, 5.5, 0.42, label,
                    font_size=15, color=TEXT_LIGHT)

    # Vertical separator
    add_rect(slide, 6.6, 1.4, 0.03, 5.5, fill_color=ACCENT_BLUE)


def slide_context(prs):
    """Slide 3 – Context și motivație."""
    slide = add_slide(prs)
    set_bg(slide, BG_DARK)
    add_top_bar(slide, "CONTEXTUL ȘI MOTIVAȚIA TEMEI",
                "De ce este importantă detectarea textului generat de AI?")
    add_footer(slide, PRESENTATION_TITLE)
    add_slide_number(slide, 3, TOTAL_SLIDES)

    # Left column - problem description
    add_bullet_card(slide, 0.35, 1.35, 6.1, 2.8,
        "Problema actuală",
        [
            "Modelele LLM (ChatGPT, GPT-4, Gemini) generează texte aproape\n  indistinguibile de cele scrise de oameni",
            "Platformele sociale permit publicarea instantanee, fără filtrare\n  editorială prealabilă",
            "70% dintre utilizatori consideră necesară reglementarea AI\n  (KPMG & Univ. Melbourne, sondaj global 2025, 48.000+ respondenți)",
            "64% sunt îngrijorați că AI poate influența alegerile",
        ],
        title_color=ACCENT_CYAN, bullet_size=14
    )

    # Right column - impact
    add_bullet_card(slide, 6.8, 1.35, 6.15, 2.8,
        "Impact real documentat",
        [
            "Oct. 2023: articol fals despre aprobarea unui ETF Bitcoin a produs\n  volatilitate pe piețele financiare",
            "Numărul site-urilor de știri false susținute de AI a crescut de 10×\n  în 2023 (49 → 600+ domenii, fără supervizare umană)",
            "Factorul uman contribuie la 68% din incidentele de securitate\n  cibernetică (Verizon DBIR 2024)",
            ">1/3 din conținutul de pe Medium, Quora și Reddit (2022-2024)\n  este generat de AI (Sun et al., 2024)",
        ],
        title_color=ACCENT_ORANGE, bullet_size=14
    )

    # Bottom - why social media
    add_rect(slide, 0.35, 4.35, 12.6, 1.05, fill_color=BG_CARD,
             line_color=ACCENT_GREEN, line_width_pt=1.0)
    add_textbox(slide, 0.55, 4.42, 12.2, 0.3,
                "De ce rețelele sociale reprezintă mediul cel mai vulnerabil?",
                font_size=14, bold=True, color=ACCENT_GREEN)
    add_textbox(slide, 0.55, 4.72, 12.2, 0.6,
                "Texte scurte & informale  ·  Publicare instantanee, fără filtrare  ·  Viteza propagării dezinformării  ·  "
                "Limbaj colocvial, abrevieri, greșeli gramaticale intenționate\n"
                "→ metodele clasice de detectare, dezvoltate pentru texte formale și lungi, FUNCȚIONEAZĂ SLAB pe social media",
                font_size=13, color=TEXT_LIGHT)

    add_textbox(slide, 0.35, 5.55, 12.6, 0.45,
                "Scopul lucrării: analiza metodelor existente și propunerea unei soluții bazate pe machine learning,"
                " adaptată specific pentru texte scurte de social media.",
                font_size=13, bold=True, color=ACCENT_BLUE)


def slide_nlp_llm(prs):
    """Slide 4 – NLP și LLM."""
    slide = add_slide(prs)
    set_bg(slide, BG_DARK)
    add_top_bar(slide, "PROCESAREA LIMBAJULUI NATURAL ȘI MODELELE LLM",
                "Fundamentele teoretice")
    add_footer(slide, PRESENTATION_TITLE)
    add_slide_number(slide, 4, TOTAL_SLIDES)

    # NLP definition
    add_rect(slide, 0.35, 1.35, 12.6, 0.7, fill_color=BG_CARD,
             line_color=ACCENT_BLUE, line_width_pt=1.0)
    add_textbox(slide, 0.55, 1.42, 12.2, 0.25,
                "NLP (Natural Language Processing)",
                font_size=15, bold=True, color=ACCENT_CYAN)
    add_textbox(slide, 0.55, 1.68, 12.2, 0.32,
                "Subdomeniu AI care permite sistemelor informatice să înțeleagă, interpreteze și genereze text"
                " similar oamenilor. Sarcini: traducere automată, rezumare, analiză sentimente, clasificare.",
                font_size=13, color=TEXT_LIGHT)

    # Three columns: Transformer, BERT/RoBERTa, GPT
    add_bullet_card(slide, 0.35, 2.2, 3.95, 3.2,
        "Arhitectura Transformer",
        [
            "Vaswani et al. (2017) — revoluționează NLP",
            "Înlocuiește arhitecturile recurente",
            "Mecanism de self-attention: fiecare cuvânt\n  interacționează direct cu oricare altul,\n  indiferent de distanță",
            "Baza tuturor modelelor LLM moderne",
        ],
        title_color=ACCENT_BLUE, bullet_size=13
    )

    add_bullet_card(slide, 4.55, 2.2, 4.0, 3.2,
        "BERT / RoBERTa (Encoder)",
        [
            "BERT: procesare bidirecțională simultană\n  (vezi întreg contextul înainte de predicție)",
            "Antrenare: Masked Language Modeling\n  + Next Sentence Prediction",
            "Tokenizare: Byte-Pair Encoding (BPE)",
            "RoBERTa (Liu et al., 2019):\n  BERT optimizat, performanțe superioare\n  pe clasificare text",
            "125 milioane parametri (RoBERTa-base)",
        ],
        title_color=ACCENT_CYAN, bullet_size=13
    )

    add_bullet_card(slide, 8.8, 2.2, 4.15, 3.2,
        "GPT-3 și modelele generative",
        [
            "Brown et al. (2020) — 175B parametri",
            "Model autoregresiv: generează token cu token",
            "Few-shot learning: rezolvă sarcini noi\n  fără date de antrenament specifice",
            "Generează articole de știri pe care\n  evaluatorii umani NU le pot distinge\n  de cele scrise de oameni",
            "→ necesitatea detectării automate!",
        ],
        title_color=ACCENT_PURPLE, bullet_size=13
    )

    # Foundation models note
    add_textbox(slide, 0.35, 5.55, 12.6, 0.45,
                'Bommasani et al. (2022) - Foundation Models: riscul omogenizarii; orice bias dintr-un model de baza'
                " este mostenit de TOATE aplicatiile construite pe el.",
                font_size=12, italic=True, color=TEXT_GREY)


def slide_methods(prs):
    """Slide 5 – Metode de detectare."""
    slide = add_slide(prs)
    set_bg(slide, BG_DARK)
    add_top_bar(slide, "METODE DE DETECTARE A TEXTULUI GENERAT DE AI",
                "Categorii din literatura de specialitate")
    add_footer(slide, PRESENTATION_TITLE)
    add_slide_number(slide, 5, TOTAL_SLIDES)

    # 4 categories in 2x2 grid
    add_bullet_card(slide, 0.35, 1.35, 6.1, 2.55,
        "1. Metode statistice",
        [
            "Analizează perplexitatea, entropia, rangul log-probabilistic",
            "Text AI = perplexitate scăzută (modelul alege\n  cuvintele cele mai probabile)",
            "GLTR (Gehrmann et al., 2019): colorează cuvintele\n  după rangul în distribuția modelului",
            "⚠ Nu necesită antrenare, DAR performează slab\n  pe texte scurte și modele noi",
        ],
        title_color=ACCENT_BLUE, bullet_size=13
    )

    add_bullet_card(slide, 6.85, 1.35, 6.1, 2.55,
        "2. Metode bazate pe Watermarking",
        [
            "Kirchenbauer et al. (2023): semnal ascuns\n  introdus LA GENERARE",
            "Vocabularul impartit in lista verde (favorizata)\n  si lista rosie (penalizata)",
            "Detecție cu rată de eroare foarte mică",
            "⚠ Necesită cooperarea furnizorului modelului\n⚠ Parafrazarea poate elimina semnalul",
        ],
        title_color=ACCENT_CYAN, bullet_size=13
    )

    add_bullet_card(slide, 0.35, 4.05, 6.1, 2.55,
        "3. Clasificatori supervizați",
        [
            "Învață să distingă text uman vs. AI din exemple etichetate",
            "Chakraborty et al. (2023): detectarea este aproape\n  ÎNTOTDEAUNA posibilă cât timp există diferențe statistice",
            "RoBERTa fine-tuned: cea mai bună performanță",
            "⚠ Depinde de calitatea datelor de antrenare\n⚠ Necesită reantrenare pe modele noi",
        ],
        title_color=ACCENT_PURPLE, bullet_size=13
    )

    add_bullet_card(slide, 6.85, 4.05, 6.1, 2.55,
        "4. Metode zero-shot",
        [
            "DetectGPT (Mitchell et al., 2023): textul AI ocupa\n  zone de curbura negativa in spatiul log-probabilitatii",
            "Perturbarea minoră a textului AI → scădere mai mare\n  a log-probabilității decât pentru textul uman",
            "Avantaj: nu necesită date de antrenament",
            "⚠ Necesită acces la un LLM similar celui generator\n⚠ Performanță slabă pe texte scurte",
        ],
        title_color=ACCENT_ORANGE, bullet_size=13
    )


def slide_existing_apps(prs):
    """Slide 6 – Aplicații similare existente."""
    slide = add_slide(prs)
    set_bg(slide, BG_DARK)
    add_top_bar(slide, "APLICAȚII SIMILARE EXISTENTE",
                "Instrumente comerciale și limitările lor")
    add_footer(slide, PRESENTATION_TITLE)
    add_slide_number(slide, 6, TOTAL_SLIDES)

    apps = [
        ("GPTZero\n(Edward Tian, 2023)",
         "Perplexitate + Burstiness\n• Proiectat pentru eseuri academice\n• Slab pe texte scurte & informale\n• Parafrazarea reduce acuratețea",
         ACCENT_BLUE),
        ("ZeroGPT\n(2023)",
         "Metode statistice combinate\n• Metodologie nepublicată\n• Rată ridicată fals-pozitive pe texte\n  tehnice formale\n• Gratuit, fără transparență",
         ACCENT_CYAN),
        ("OpenAI Classifier\n(Retras 2023)",
         "Fine-tuned pe perechi uman/AI\n• Identifică corect doar 26% din\n  textele AI\n• Clasifică eronat 9% texte umane\n• Retras în iulie 2023",
         ACCENT_ORANGE),
        ("Originality.ai",
         "AI + plagiat + lizibilitate\n• Acuratețe: 76% în teste standardizate\n• Detectează text AI parafrazat\n• ⚠ Doar engleză, DOAR CU PLATĂ",
         ACCENT_PURPLE),
        ("Sapling AI Detector",
         "Model neuronal + API REST\n• Acuratețe: 68–87%\n• 0 fals-pozitive în unele evaluări\n• Disponibil pentru integrări\n• Metodologie neclară",
         ACCENT_GREEN),
    ]

    for i, (name, desc, color) in enumerate(apps):
        x = 0.35 + i * 2.58
        add_rect(slide, x, 1.35, 2.42, 4.1, fill_color=BG_CARD,
                 line_color=color, line_width_pt=1.5)
        add_rect(slide, x, 1.35, 2.42, 0.45, fill_color=color)
        add_textbox(slide, x+0.1, 1.37, 2.22, 0.42, name,
                    font_size=12, bold=True, color=TEXT_WHITE,
                    align=PP_ALIGN.CENTER)
        add_textbox(slide, x+0.1, 1.88, 2.22, 3.4, desc,
                    font_size=12, color=TEXT_LIGHT, word_wrap=True)

    # Common limitation
    add_rect(slide, 0.35, 5.6, 12.6, 0.75, fill_color=BG_CARD,
             line_color=ACCENT_RED, line_width_pt=1.5)
    add_textbox(slide, 0.55, 5.65, 12.2, 0.3,
                "Limitare comună TUTUROR instrumentelor:",
                font_size=14, bold=True, color=ACCENT_RED)
    add_textbox(slide, 0.55, 5.92, 12.2, 0.36,
                "Au fost proiectate și evaluate pe texte lungi și formale (eseuri, articole)."
                " Performanța pe texte scurte și informale, tipice rețelelor sociale, este semnificativ mai slabă."
                " → Aceasta este motivația directă a soluției propuse în această lucrare.",
                font_size=13, color=TEXT_LIGHT)


def slide_limitations(prs):
    """Slide 7 – Limitări ale metodelor existente."""
    slide = add_slide(prs)
    set_bg(slide, BG_DARK)
    add_top_bar(slide, "LIMITĂRI ALE METODELOR EXISTENTE",
                "De ce metodele actuale nu sunt suficiente?")
    add_footer(slide, PRESENTATION_TITLE)
    add_slide_number(slide, 7, TOTAL_SLIDES)

    limitations = [
        (ACCENT_RED, "Parafrazarea",
         "Krishna et al. (2023) demonstrează că un simplu pas de parafrazare automată"
         " poate reduce rata de detecție de la peste 90% la sub 50%,"
         " făcând detectorul practic inutil."),
        (ACCENT_ORANGE, "Limita fundamentală",
         "Sadasivan et al. (2023): pe măsură ce modelele LLM devin mai performante,"
         " detectarea fiabilă devine teoretic tot mai dificilă."
         " Există o limită determinată de distanța totală de variație între distribuțiile textului uman și AI."),
        (ACCENT_CYAN, "Texte scurte",
         "Majoritatea metodelor existente performează semnificativ mai bine pe texte lungi."
         " Pe textele scurte tipice social media (50-500 caractere),"
         " diferențele statistice dintre textul uman și AI devin mai greu de detectat."),
        (ACCENT_PURPLE, "Problema domeniului de antrenare",
         "Detectori antrenați pe texte formale (articole academice, știri, eseuri)"
         " generalizează slab pe texte informale de social media:"
         " abrevieri, emoticoane, greșeli gramaticale voite, limbaj colocvial."),
        (ACCENT_BLUE, "Transferabilitate",
         "Un detector antrenat pe texte GPT-3 poate funcționa slab pe texte produse de LLaMA sau Gemini."
         " Fiecare model generativ are propriile tipare stilistice."),
    ]

    for i, (color, title, text) in enumerate(limitations):
        y = 1.38 + i * 1.15
        add_rect(slide, 0.35, y, 0.08, 0.9, fill_color=color)
        add_rect(slide, 0.55, y, 12.4, 0.9, fill_color=BG_CARD)
        add_textbox(slide, 0.7, y+0.05, 2.8, 0.35, title,
                    font_size=15, bold=True, color=color)
        add_textbox(slide, 0.7, y+0.42, 12.1, 0.42, text,
                    font_size=13, color=TEXT_LIGHT, word_wrap=True)


def slide_proposed_method(prs):
    """Slide 8 – Metoda propusă."""
    slide = add_slide(prs)
    set_bg(slide, BG_DARK)
    add_top_bar(slide, "METODA PROPUSĂ",
                "Arhitectura sistemului de detectare")
    add_footer(slide, PRESENTATION_TITLE)
    add_slide_number(slide, 8, TOTAL_SLIDES)

    # Pipeline boxes
    pipeline = [
        ("Colectare &\nFiltrare date", ACCENT_BLUE),
        ("Preprocesare\nDiferențiată", ACCENT_CYAN),
        ("Împărțire\nStratificată\n70/15/15", ACCENT_PURPLE),
        ("Antrenare\nParalelă\n3 modele", ACCENT_ORANGE),
        ("Evaluare\nComparativă", ACCENT_GREEN),
        ("Salvare modele\n+ Aplicație web", ACCENT_BLUE),
    ]
    for i, (label, color) in enumerate(pipeline):
        x = 0.4 + i * 2.12
        add_rect(slide, x, 1.4, 1.8, 1.1, fill_color=BG_CARD,
                 line_color=color, line_width_pt=2.0)
        add_textbox(slide, x+0.08, 1.48, 1.65, 0.95, label,
                    font_size=12, bold=True, color=color,
                    align=PP_ALIGN.CENTER, word_wrap=True)
        if i < len(pipeline) - 1:
            add_textbox(slide, x+1.82, 1.85, 0.28, 0.4, "→",
                        font_size=22, bold=True, color=TEXT_GREY,
                        align=PP_ALIGN.CENTER)

    # Three models detail
    add_bullet_card(slide, 0.35, 2.75, 3.95, 2.95,
        "Naive Bayes (clasic)",
        [
            "TF-IDF: vocabular 10.000 termeni,\n  n-grame 1–2, normalizare sublineară",
            "Rapid și interpretabil",
            "Serveste ca baseline (linie de bază)",
            "Acuratețe: 94.24% pe setul de test",
            "Antrenare: < 2 minute",
        ],
        title_color=ACCENT_CYAN, bullet_size=13
    )

    add_bullet_card(slide, 4.55, 2.75, 4.22, 2.95,
        "Regresie Logistică (clasic)",
        [
            "TF-IDF + regularizare",
            "Echilibru bun performanță/interpretabilitate",
            "Precizie ridicată (98.11%) — puține\n  fals-pozitive",
            "Acuratețe: 96.71%",
            "Alternativă viabilă fără GPU",
        ],
        title_color=ACCENT_PURPLE, bullet_size=13
    )

    add_bullet_card(slide, 9.02, 2.75, 3.95, 2.95,
        "RoBERTa Fine-tuned (DL)",
        [
            "125 milioane parametri (RoBERTa-base)",
            "Fine-tuning: 3 epoci, GPU Tesla T4",
            "Tokenizare BPE, max 128 tokeni",
            "Early stopping pe F1-score",
            "Acuratețe: 99.17% ✓ (model principal)",
        ],
        title_color=ACCENT_BLUE, bullet_size=13
    )

    # Note
    add_textbox(slide, 0.35, 5.85, 12.6, 0.45,
                "Motivație: Naive Bayes și Regresia Logistică servesc ca baseline. "
                "Compararea celor 3 modele pe același set de date permite alegerea justificată "
                "a modelului potrivit pentru contextul de utilizare.",
                font_size=12, italic=True, color=TEXT_GREY)


def slide_dataset(prs):
    """Slide 9 – Setul de date."""
    slide = add_slide(prs)
    set_bg(slide, BG_DARK)
    add_top_bar(slide, "SETUL DE DATE",
                "Construcție și caracteristici")
    add_footer(slide, PRESENTATION_TITLE)
    add_slide_number(slide, 9, TOTAL_SLIDES)

    # Sources table
    sources = [
        ("HC3 Reddit ELI5", "Mixt", "4.040 + 629",
         "Perechi uman/AI pe aceleași întrebări → elimină variabilele de conținut"),
        ("TweetEval", "Uman", "8.176",
         "Benchmark academic, tweets reale, 7 categorii distincte"),
        ("Reddit comentarii", "Uman", "4.058",
         "Limbaj conversațional, argou, ironii, referințe culturale"),
        ("RAID (2024)", "AI", "7.771",
         "Texte generate de GPT-4, LLaMA, Mistral — diversitate stilistică AI"),
        ("AI_Human.csv", "Mixt", "20.548",
         "Texte mai formale — volum și diversitate tematică"),
        ("AI Detection", "AI", "612",
         "Texte AI suplimentare pentru diversitate"),
    ]

    headers = ["Sursă", "Clasă", "Exemple", "Motivație"]
    col_widths = [2.2, 0.9, 1.2, 7.4]
    x_starts = [0.35, 2.62, 3.56, 4.82]

    # Header row
    add_rect(slide, 0.35, 1.35, 12.6, 0.42, fill_color=ACCENT_BLUE)
    for j, (header, cw, xs) in enumerate(zip(headers, col_widths, x_starts)):
        add_textbox(slide, xs+0.08, 1.38, cw-0.1, 0.38, header,
                    font_size=13, bold=True, color=TEXT_WHITE)

    for i, (src, cls, ex, motiv) in enumerate(sources):
        y = 1.82 + i * 0.67
        row_color = BG_CARD if i % 2 == 0 else BG_DARK
        add_rect(slide, 0.35, y, 12.6, 0.65, fill_color=row_color)
        add_rect(slide, 0.35, y, 12.6, 0.65, fill_color=None,
                 line_color=RGBColor(0x22, 0x22, 0x44), line_width_pt=0.5)
        texts = [src, cls, ex, motiv]
        for j, (text, cw, xs) in enumerate(zip(texts, col_widths, x_starts)):
            color = ACCENT_CYAN if j == 0 else TEXT_LIGHT
            add_textbox(slide, xs+0.06, y+0.08, cw-0.08, 0.52, text,
                        font_size=12, color=color, word_wrap=True)

    # Summary stats
    add_rect(slide, 0.35, 5.88, 12.6, 0.75, fill_color=BG_CARD,
             line_color=ACCENT_GREEN, line_width_pt=1.0)
    stats = [
        ("78.961", "exemple înainte de echilibrare"),
        ("45.834", "exemple finale (echilibrate)"),
        ("50/50", "uman / AI"),
        ("50–500\ncaractere", "lungime filtrată"),
        ("70/15/15%", "train / val / test"),
    ]
    for i, (val, label) in enumerate(stats):
        x = 0.7 + i * 2.5
        add_textbox(slide, x, 5.9, 2.2, 0.38, val,
                    font_size=20, bold=True, color=ACCENT_GREEN,
                    align=PP_ALIGN.CENTER)
        add_textbox(slide, x, 6.28, 2.2, 0.32, label,
                    font_size=11, color=TEXT_GREY,
                    align=PP_ALIGN.CENTER, word_wrap=True)


def slide_implementation(prs):
    """Slide 10 – Implementare tehnică."""
    slide = add_slide(prs)
    set_bg(slide, BG_DARK)
    add_top_bar(slide, "IMPLEMENTARE TEHNICĂ",
                "Stack tehnologic și arhitectura aplicației")
    add_footer(slide, PRESENTATION_TITLE)
    add_slide_number(slide, 10, TOTAL_SLIDES)

    # Tech stack - 4 columns
    stacks = [
        ("Date", ACCENT_BLUE,
         ["Python 3.10",
          "pandas — filtrare, deduplicare, echilibrare",
          "datasets (Hugging Face) — streaming\n  (RAID > RAM disponibil)"]),
        ("ML Clasic", ACCENT_CYAN,
         ["scikit-learn — NB, LR",
          "TfidfVectorizer: vocabular 10.000,\n  ngram_range=(1,2), sublinear_tf=True",
          "Evaluare: sklearn.metrics"]),
        ("Deep Learning", ACCENT_PURPLE,
         ["PyTorch 2.x — GPU operations",
          "Transformers HuggingFace 4.x\n  (RoBERTa, Trainer, Tokenizer)",
          "FP16 mixed precision → ↓ memorie GPU",
          "NLTK — eliminare cuvinte stop",
          "SpaCy 3.x — Named Entity Recognition"]),
        ("Aplicație web", ACCENT_GREEN,
         ["Flask 3.x — API REST:\n  /predict și /batch",
          "HTML / CSS / JavaScript",
          "Chart.js 4.x — grafice cerc și bară",
          "Flask-CORS pentru cross-origin"]),
    ]

    for i, (title, color, items) in enumerate(stacks):
        x = 0.35 + i * 3.22
        add_bullet_card(slide, x, 1.35, 3.05, 3.85,
            title, items, title_color=color, bullet_size=12)

    # Hardware environment
    add_rect(slide, 0.35, 5.35, 12.6, 0.95, fill_color=BG_CARD,
             line_color=ACCENT_ORANGE, line_width_pt=1.0)
    add_textbox(slide, 0.55, 5.4, 12.2, 0.3,
                "Mediu hardware",
                font_size=14, bold=True, color=ACCENT_ORANGE)
    hw_items = [
        "ANTRENARE: Google Colaboratory — GPU NVIDIA Tesla T4, 15.6 GB VRAM GDDR6,"
        " precizie mixtă FP16",
        "Timp antrenare RoBERTa: ~85 min (3 epoci, 3.009 pași).  "
        "Modele clasice: < 2 min.   RAM sistem: 12.7 GB.",
        "INFERENȚĂ: CPU local (laptop).  RoBERTa: 0.5–2 sec/text.  "
        "Modele salvate: RoBERTa 502 MB, NB 313 KB, LR 78 KB, TF-IDF 394 KB.",
    ]
    for j, item in enumerate(hw_items):
        add_textbox(slide, 0.55, 5.7 + j * 0.0, 12.2, 0.22, item,
                    font_size=11, color=TEXT_LIGHT, word_wrap=True)


def slide_preprocessing(prs):
    """Slide 11 – Preprocesare și componente avansate."""
    slide = add_slide(prs)
    set_bg(slide, BG_DARK)
    add_top_bar(slide, "PREPROCESARE ȘI COMPONENTE AVANSATE",
                "Detalii de implementare")
    add_footer(slide, PRESENTATION_TITLE)
    add_slide_number(slide, 11, TOTAL_SLIDES)

    # Preprocessing
    add_bullet_card(slide, 0.35, 1.38, 5.8, 2.5,
        "Preprocesare diferențiată",
        [
            "clean_for_nb (Naive Bayes / LR):\n  → litere mici → elimină URL, @user, #hashtag\n  → doar caractere alfabetice\n  → elimină cuvinte stop (NLTK)",
            "clean_for_roberta (RoBERTa):\n  → URL → [URL], @user → [USER]\n  → păstrează structura textuală\n  → max 128 tokeni BPE",
            "Inconsistența între antrenare și inferență\n  produce degradare silențioasă → identic!",
        ],
        title_color=ACCENT_BLUE, bullet_size=13
    )

    add_bullet_card(slide, 6.4, 1.38, 6.55, 2.5,
        "Leave-One-Out — justificarea deciziei",
        [
            "Elimină pe rând fiecare cuvânt din text",
            "Măsoară modificarea probabilității predicției",
            "Cuvinte cu impact pozitiv → indicator spre AI",
            "Cuvinte cu impact negativ → indicator spre uman",
            "Implementat pentru TOATE cele 3 modele\n  (NB: feature_log_prob_; LR: coef_)",
            "Avantaj față de toate instrumentele comerciale:\n  justificare la nivel de cuvânt",
        ],
        title_color=ACCENT_CYAN, bullet_size=13
    )

    add_bullet_card(slide, 0.35, 4.05, 5.8, 2.55,
        "Scorul agregat de risc (0–100)",
        [
            "RoBERTa score: 60% din total (max 60 pct)",
            "Fraze tipice AI detectate: max 25 pct\n  (5 pct/frază): \"it is worth noting\",\n  \"furthermore\", \"în concluzie\" etc.",
            "Profil statistic vocabular: max 15 pct\n  (TTR, % cuvinte comune, lungime medie propoziții)",
            "Vizualizare: indicator circular animat\n  Verde(0-39) · Galben(40-69) · Roșu(70-100)",
        ],
        title_color=ACCENT_PURPLE, bullet_size=13
    )

    add_bullet_card(slide, 6.4, 4.05, 6.55, 2.55,
        "Alte componente ale aplicației",
        [
            "Analiză per propoziție: colorare pe 5 niveluri\n  (AI puternic → uman puternic)",
            "NER (SpaCy en_core_web_sm): persoane,\n  organizații, locații, date",
            "Comparație simultană 3 modele:\n  switchView fără apel suplimentar la server",
            "Batch CSV: procesare până la 500 texte,\n  export rezultate CSV",
            "Istoric sesiune: ultimele 10 analize",
        ],
        title_color=ACCENT_GREEN, bullet_size=13
    )


def slide_results(prs):
    """Slide 12 – Rezultate."""
    slide = add_slide(prs)
    set_bg(slide, BG_DARK)
    add_top_bar(slide, "REZULTATELE OBȚINUTE",
                "Evaluare pe setul de test (6.876 exemple, 50/50 uman/AI)")
    add_footer(slide, PRESENTATION_TITLE)
    add_slide_number(slide, 12, TOTAL_SLIDES)

    # Metrics for 3 models
    models = [
        ("Naive Bayes", "94.24%", "94.17%", "95.35%", "93.02%", ACCENT_CYAN),
        ("Regresie Logistică", "96.71%", "96.66%", "98.11%", "95.26%", ACCENT_PURPLE),
        ("RoBERTa Fine-tuned", "99.17%", "99.18%", "98.56%", "99.80%", ACCENT_GREEN),
    ]
    headers = ["Model", "Acuratețe", "F1-Score", "Precision", "Recall"]
    col_w = [3.0, 2.2, 2.2, 2.2, 2.2]
    col_x = [0.35, 3.42, 5.68, 7.95, 10.22]

    # Table header
    add_rect(slide, 0.35, 1.35, 12.1, 0.45, fill_color=ACCENT_BLUE)
    for j, (h, cw, cx) in enumerate(zip(headers, col_w, col_x)):
        add_textbox(slide, cx+0.1, 1.38, cw-0.15, 0.4, h,
                    font_size=14, bold=True, color=TEXT_WHITE,
                    align=PP_ALIGN.CENTER)

    for i, (name, acc, f1, prec, rec, color) in enumerate(models):
        y = 1.84 + i * 0.75
        row_color = BG_CARD if i % 2 == 0 else BG_DARK
        add_rect(slide, 0.35, y, 12.1, 0.72, fill_color=row_color)
        add_rect(slide, 0.35, y, 0.06, 0.72, fill_color=color)
        vals = [name, acc, f1, prec, rec]
        for j, (val, cw, cx) in enumerate(zip(vals, col_w, col_x)):
            is_best = (i == 2)
            c = ACCENT_GREEN if is_best and j > 0 else (ACCENT_CYAN if j == 0 else TEXT_WHITE)
            add_textbox(slide, cx+0.1, y+0.14, cw-0.15, 0.45, val,
                        font_size=16 if is_best and j > 0 else 14,
                        bold=(is_best and j > 0),
                        color=c, align=PP_ALIGN.CENTER)

    # Confusion matrix info
    add_rect(slide, 0.35, 4.1, 5.8, 1.65, fill_color=BG_CARD,
             line_color=ACCENT_GREEN, line_width_pt=1.0)
    add_textbox(slide, 0.55, 4.15, 5.4, 0.32,
                "Confusion Matrix RoBERTa",
                font_size=14, bold=True, color=ACCENT_GREEN)
    add_textbox(slide, 0.55, 4.48, 5.4, 1.2,
                "✓  3.388 texte umane clasificate CORECT\n"
                "✓  3.431 texte AI clasificate CORECT\n"
                "⚠  50 fals-pozitive (texte umane clasificate ca AI)\n"
                "⚠  7 fals-negative (texte AI clasificate ca umane)\n"
                "→ Toate erorile provin din texte formale (ai_human_csv)",
                font_size=13, color=TEXT_LIGHT, word_wrap=True)

    # Per source accuracy
    add_rect(slide, 6.4, 4.1, 6.05, 1.65, fill_color=BG_CARD,
             line_color=ACCENT_CYAN, line_width_pt=1.0)
    add_textbox(slide, 6.6, 4.15, 5.7, 0.32,
                "Acuratețe RoBERTa pe surse social media",
                font_size=14, bold=True, color=ACCENT_CYAN)
    src_data = [
        ("TweetEval extra (human)", "100%"),
        ("Reddit ELI5 (human + AI)", "100%"),
        ("Twitter human", "99.84%"),
        ("RAID AI", "99.57%"),
        ("Reddit human", "98.10%"),
    ]
    for j, (src, acc_val) in enumerate(src_data):
        y_off = 4.5 + j * 0.25
        add_textbox(slide, 6.6, y_off, 4.5, 0.24, f"• {src}",
                    font_size=12, color=TEXT_LIGHT)
        add_textbox(slide, 11.0, y_off, 1.2, 0.24, acc_val,
                    font_size=12, bold=True, color=ACCENT_GREEN,
                    align=PP_ALIGN.RIGHT)

    # Training curve info
    add_rect(slide, 0.35, 5.9, 12.1, 0.72, fill_color=BG_CARD,
             line_color=ACCENT_PURPLE, line_width_pt=1.0)
    epochs_data = [
        ("Epocă 1", "0.0405", "0.0988", "97.76%", "97.80%"),
        ("Epocă 2 ★", "0.0136", "0.0520", "99.16%", "99.16%"),
        ("Epocă 3", "0.0034", "0.0633", "99.03%", "99.03%"),
    ]
    add_textbox(slide, 0.55, 5.93, 4.0, 0.28,
                "Curba de antrenare RoBERTa:",
                font_size=13, bold=True, color=ACCENT_PURPLE)
    for j, (ep, tl, vl, ac, f1_v) in enumerate(epochs_data):
        x_off = 4.7 + j * 2.62
        c = ACCENT_GREEN if j == 1 else TEXT_LIGHT
        add_textbox(slide, x_off, 5.92, 2.5, 0.65,
                    f"{ep}\nTrain loss: {tl} | Val loss: {vl}\nAcc: {ac} | F1: {f1_v}",
                    font_size=11, bold=(j == 1), color=c, word_wrap=True)


def slide_comparison(prs):
    """Slide 13 – Comparație cu instrumente comerciale."""
    slide = add_slide(prs)
    set_bg(slide, BG_DARK)
    add_top_bar(slide, "COMPARAȚIE CU INSTRUMENTE COMERCIALE",
                "Avantaje și limitări ale soluției propuse")
    add_footer(slide, PRESENTATION_TITLE)
    add_slide_number(slide, 13, TOTAL_SLIDES)

    # Comparison table
    criteria = [
        "Acuratețe (texte sociale)",
        "Optimizat pentru texte scurte\n(<500 car.)",
        "Justificare decizie\n(per cuvânt)",
        "Comparație 3 modele simultan",
        "Analiză per propoziție",
        "Scor agregat multi-semnal",
        "Procesare batch CSV",
        "Cost",
        "Transparență metodologie",
        "Suport multilingv",
    ]

    col_headers = ["Criteriu", "Soluția noastră", "GPTZero", "ZeroGPT", "Originality.ai", "Sapling"]
    col_data = [
        ["99.17%\n(social media)", "Nepublicată", "Nepublicată", "76%", "68–87%"],
        ["✓ Optimizat", "✗ Slab", "✗ Slab", "~ Moderat", "~ Moderat"],
        ["✓ Leave-One-Out", "✗ Nu", "✗ Nu", "✗ Nu", "✗ Nu"],
        ["✓ 3 modele", "✗ Nu", "✗ Nu", "✗ Nu", "✗ Nu"],
        ["✓ 5 niveluri", "✗ Nu", "~ Parțial", "✗ Nu", "✗ Nu"],
        ["✓ Da", "✗ Nu", "✗ Nu", "✗ Nu", "✗ Nu"],
        ["✓ Până la 500", "✗ Doar API", "✗ Nu", "✓ API", "✓ API"],
        ["✓ Gratuit\nopen-source", "~ Freemium", "✓ Gratuit", "✗ Plată", "✓ Gratuit"],
        ["✓ Cod public", "~ Parțial", "✗ Nu", "✗ Nu", "✗ Nu"],
        ["✗ Doar engleză", "~ Parțial", "~ Parțial", "✗ Doar engl.", "✗ Doar engl."],
    ]

    col_widths = [3.1, 2.0, 1.8, 1.8, 2.0, 1.8]
    col_x_starts = [0.35, 3.52, 5.58, 7.44, 9.3, 11.25]

    # Header
    add_rect(slide, 0.35, 1.35, 12.7, 0.42, fill_color=ACCENT_BLUE)
    for j, (h, cw, cx) in enumerate(zip(col_headers, col_widths, col_x_starts)):
        add_textbox(slide, cx+0.05, 1.37, cw-0.08, 0.4, h,
                    font_size=11, bold=True, color=TEXT_WHITE,
                    align=PP_ALIGN.CENTER)

    for i, (crit, row) in enumerate(zip(criteria, col_data)):
        y = 1.82 + i * 0.5
        row_color = BG_CARD if i % 2 == 0 else BG_DARK
        add_rect(slide, 0.35, y, 12.7, 0.48, fill_color=row_color)
        add_textbox(slide, 0.42, y+0.04, 3.0, 0.42, crit,
                    font_size=11, color=TEXT_LIGHT, word_wrap=True)
        for j, (val, cw, cx) in enumerate(zip(row, col_widths[1:], col_x_starts[1:])):
            if val.startswith("✓"):
                c = ACCENT_GREEN
            elif val.startswith("✗"):
                c = ACCENT_RED
            elif val.startswith("~"):
                c = ACCENT_ORANGE
            else:
                c = TEXT_WHITE
            # Highlight our column
            if j == 0:
                add_rect(slide, cx, y, cw, 0.48, fill_color=RGBColor(0x00, 0x1A, 0x30))
            add_textbox(slide, cx+0.05, y+0.04, cw-0.08, 0.42, val,
                        font_size=11, bold=(j == 0), color=c,
                        align=PP_ALIGN.CENTER, word_wrap=True)


def slide_contributions(prs):
    """Slide 14 – Contribuții originale."""
    slide = add_slide(prs)
    set_bg(slide, BG_DARK)
    add_top_bar(slide, "CONTRIBUȚII ORIGINALE",
                "Ce aduce nou această lucrare?")
    add_footer(slide, PRESENTATION_TITLE)
    add_slide_number(slide, 14, TOTAL_SLIDES)

    contributions = [
        (ACCENT_BLUE, "1",
         "Set de date dedicat social media",
         "6 surse combinate, 45.834 exemple filtrate la 50-500 caractere. "
         "Structura HC3 (perechi uman/AI pe aceleași întrebări) permite modelului să "
         "înveţe diferențele de stil IZOLAT de diferențele de conținut. "
         "RAID 2024 previne supra-adaptarea pe un singur model generator."),
        (ACCENT_CYAN, "2",
         "Comparare sistematică 3 clasificatori",
         "Naive Bayes, Regresie Logistică și RoBERTa pe același set de date, în "
         "condiții identice. Diferența de 2.46 pp între LR și RoBERTa la cost "
         "computațional de milisecunde vs. 0.5-2 secunde oferă perspectivă practică."),
        (ACCENT_PURPLE, "3",
         "Componenta Leave-One-Out (justificare decizie)",
         "Justificare la nivel de cuvânt: care cuvinte contribuie spre AI și care spre "
         "uman. Implementată pentru toate 3 modelele. Niciun instrument comercial "
         "(GPTZero, ZeroGPT, Originality.ai, Sapling) nu oferă această transparență."),
        (ACCENT_ORANGE, "4",
         "Scor agregat de risc multi-semnal (0-100)",
         "Combină: RoBERTa (60%), fraze tipice AI (max 25%), statistici vocabular (max 15%). "
         "Vizualizat ca indicator circular animat cu 3 niveluri cromatice (verde/galben/roșu)."),
        (ACCENT_GREEN, "5",
         "Aplicație web integrată open-source",
         "Analiză text individual, comparație 3 modele, analiză per propoziție (5 niveluri), "
         "NER (SpaCy), batch CSV până la 500 texte cu export, istoric 10 analize. "
         "Cod sursă public — verificabil, extensibil."),
    ]

    for i, (color, num, title, text) in enumerate(contributions):
        y = 1.38 + i * 1.18
        add_rect(slide, 0.35, y, 0.65, 1.0, fill_color=color)
        add_textbox(slide, 0.35, y+0.24, 0.65, 0.52, num,
                    font_size=26, bold=True, color=TEXT_WHITE,
                    align=PP_ALIGN.CENTER)
        add_rect(slide, 1.08, y, 11.6, 1.0, fill_color=BG_CARD)
        add_textbox(slide, 1.2, y+0.06, 11.3, 0.34, title,
                    font_size=15, bold=True, color=color)
        add_textbox(slide, 1.2, y+0.42, 11.3, 0.52, text,
                    font_size=12.5, color=TEXT_LIGHT, word_wrap=True)


def slide_use_cases(prs):
    """Slide 15 – Domenii de utilizare și Dezvoltări ulterioare."""
    slide = add_slide(prs)
    set_bg(slide, BG_DARK)
    add_top_bar(slide, "DOMENII DE UTILIZARE & DEZVOLTĂRI ULTERIOARE",
                "Aplicabilitate practică și direcții de cercetare")
    add_footer(slide, PRESENTATION_TITLE)
    add_slide_number(slide, 15, TOTAL_SLIDES)

    # Use cases - left column
    add_textbox(slide, 0.35, 1.38, 6.2, 0.38,
                "Domenii de utilizare",
                font_size=18, bold=True, color=ACCENT_CYAN)
    add_rect(slide, 0.35, 1.75, 6.2, 0.03, fill_color=ACCENT_CYAN)

    use_cases = [
        (ACCENT_BLUE, "Verificare conținut online",
         "Utilizatori individuali care doresc să verifice autenticitatea\npostărilor înainte de a le redistribui."),
        (ACCENT_PURPLE, "Moderare platforme sociale",
         "Procesare în lot: admin-ii analizează sistematic până la\n500 texte cu raport CSV exportabil."),
        (ACCENT_ORANGE, "Jurnalism & fact-checking",
         "Prim filtru în fluxul de lucru al jurnaliștilor,\nsemnalând sursele suspecte (600+ domenii AI în 2023)."),
        (ACCENT_CYAN, "Securitate cibernetică",
         "Detectare phishing personalizat cu LLM-uri.\nFactorul uman = 68% din incidentele de securitate."),
        (ACCENT_GREEN, "Cercetare academică",
         "Arhitectură modulară — modele noi adăugabile\nfără modificarea interfeței."),
    ]

    for i, (color, title, text) in enumerate(use_cases):
        y = 1.88 + i * 1.02
        add_rect(slide, 0.35, y, 0.06, 0.85, fill_color=color)
        add_textbox(slide, 0.5, y+0.04, 6.0, 0.32, title,
                    font_size=13, bold=True, color=color)
        add_textbox(slide, 0.5, y+0.38, 6.0, 0.42, text,
                    font_size=12, color=TEXT_LIGHT, word_wrap=True)

    # Vertical separator
    add_rect(slide, 6.82, 1.35, 0.03, 5.3, fill_color=ACCENT_BLUE)

    # Future development - right column
    add_textbox(slide, 7.0, 1.38, 6.0, 0.38,
                "Dezvoltări ulterioare",
                font_size=18, bold=True, color=ACCENT_PURPLE)
    add_rect(slide, 7.0, 1.75, 6.0, 0.03, fill_color=ACCENT_PURPLE)

    future = [
        (ACCENT_BLUE, "Extindere multilingvă",
         "XLM-RoBERTa (100+ limbi) sau seturi de date\nspecifice per limbă (incl. română — Facebook, Reddit r/România)"),
        (ACCENT_CYAN, "Reantrenare periodică",
         "Pipeline semi-automat pe GPT-4o, Claude, Gemini,\nLLaMA 3 — combatere concept drift"),
        (ACCENT_PURPLE, "Robustețe la atacuri de evitare",
         "Testare sistematica: parafrazare, humanizare,\ntexte mixte, post-procesare comerciala"),
        (ACCENT_ORANGE, "Semnale comportamentale",
         "Frecvență postare, redistribuiri, profil cont —\nanaliză multi-modală (text + comportament rețea)"),
        (ACCENT_GREEN, "Infrastructură",
         "Extensie browser, API public REST,\ncontainerizare Docker pentru deployment cloud"),
        (ACCENT_BLUE, "Calibrare scor agregat",
         "Optimizare automată a ponderilor (60/25/15)\npe set de validare dedicat"),
    ]

    for i, (color, title, text) in enumerate(future):
        y = 1.88 + i * 0.85
        add_rect(slide, 7.0, y, 0.38, 0.38, fill_color=color)
        add_textbox(slide, 7.0, y+0.06, 0.38, 0.28,
                    str(i+1), font_size=15, bold=True, color=TEXT_WHITE,
                    align=PP_ALIGN.CENTER)
        add_textbox(slide, 7.48, y+0.02, 5.55, 0.32, title,
                    font_size=13, bold=True, color=color)
        add_textbox(slide, 7.48, y+0.36, 5.55, 0.45, text,
                    font_size=12, color=TEXT_LIGHT, word_wrap=True)


def slide_conclusions(prs):
    """Slide 16 – Concluzii."""
    slide = add_slide(prs)
    set_bg(slide, BG_DARK)
    add_top_bar(slide, "CONCLUZII",
                "Ce a demonstrat această lucrare?")
    add_footer(slide, PRESENTATION_TITLE)
    add_slide_number(slide, 16, TOTAL_SLIDES)

    # Main conclusion boxes
    conclusions = [
        (ACCENT_GREEN,
         "Detectarea este posibilă chiar pe texte scurte de social media",
         "Contra intuiției, detectarea textului AI funcționează bine chiar și pe texte de 50-500 caractere"
         " — acuratețe 99.17% pe setul de test și >99.5% pe sursele propriu-zise de social media."
         " Confirmă teorema lui Chakraborty et al. (2023): cât timp există diferențe statistice,"
         " un clasificator cu suficiente date le poate exploata."),
        (ACCENT_BLUE,
         "RoBERTa depășește semnificativ metodele clasice",
         "Diferența de 5 pp față de Naive Bayes și 2.5 pp față de Regresia Logistică demonstrează"
         " valoarea modelării contextuale Transformer față de reprezentările statistice TF-IDF."
         " Totuși, Regresia Logistică (96.71%) rămâne o alternativă viabilă în scenarii fără GPU."),
        (ACCENT_CYAN,
         "Setul de date este crucial",
         "Aproape toate erorile provin din texte formale (ai_human_csv)."
         " Pe toate sursele de social media propriu-zise, acuratețea depășește 99.5%,"
         " validând decizia de a antrena pe date specifice acestui domeniu."),
        (ACCENT_PURPLE,
         "Limitări rămase și direcții de viitor",
         "Dependența de limba engleză, lipsa testării sistematice la parafrazare"
         " și riscul de concept drift pe modele noi sunt limitările principale."
         " Extinderea cu XLM-RoBERTa și reantrenarea periodică sunt direcțiile prioritare."),
    ]

    for i, (color, title, text) in enumerate(conclusions):
        y = 1.4 + i * 1.38
        add_rect(slide, 0.35, y, 12.6, 1.28, fill_color=BG_CARD,
                 line_color=color, line_width_pt=1.5)
        add_rect(slide, 0.35, y, 0.08, 1.28, fill_color=color)
        add_textbox(slide, 0.52, y+0.07, 12.2, 0.38, title,
                    font_size=15, bold=True, color=color)
        add_textbox(slide, 0.52, y+0.48, 12.2, 0.72, text,
                    font_size=13, color=TEXT_LIGHT, word_wrap=True)


def slide_bibliography(prs):
    """Slide 17 – Bibliografie."""
    slide = add_slide(prs)
    set_bg(slide, BG_DARK)
    add_top_bar(slide, "BIBLIOGRAFIE",
                "Referințe utilizate în prezentare")
    add_footer(slide, PRESENTATION_TITLE)
    add_slide_number(slide, 17, TOTAL_SLIDES)

    bibliography = [
        "[1] Barbieri, F., Camacho-Collados, J., Espinosa-Anke, L., & Neves, L. (2020). TweetEval: Unified Benchmark and Comparative Evaluation for Tweet Classification. arXiv:2010.12421.",
        "[4] Bommasani, R. et al. (2022). On the Opportunities and Risks of Foundation Models. arXiv:2108.07258.",
        "[5] Brown, T., Mann, B., Ryder, N., Subbiah, M., Kaplan, J., Dhariwal, P., ... & Amodei, D. (2020). Language Models are Few-Shot Learners. arXiv:2005.14165.",
        "[6] Chakraborty, S., Bedi, A. S., Zhu, S., An, B., Manocha, D., & Huang, F. (2023). On the Possibilities of AI-Generated Text Detection. arXiv:2304.04736.",
        "[9] Dugan, L., Hwang, A., Trhlík, F., Zhu, Z., Ippolito, D., & Callison-Burch, C. (2024). RAID: A Shared Benchmark for Robust Evaluation of Machine-Generated Text Detectors. Proceedings of ACL 2024. arXiv:2405.07940.",
        "[10] Gehrmann, S., Strobelt, H., & Rush, A. M. (2019). GLTR: Statistical Detection and Visualization of Generated Text. arXiv:1906.04043.",
        "[11] Gillespie, N., Lockey, S., Curtis, C., Pool, J., & Akbari, A. (2025). Trust, Attitudes and Use of Artificial Intelligence: A Global Study 2025. The University of Melbourne & KPMG.",
        "[13] Guo, W., Shen, W., Lei, J., Chow, K., & Shi, E. (2023). How Close is ChatGPT to Human Experts? Comparison Corpus, Evaluation, and Detection (HC3). arXiv:2301.07597.",
        "[15] Kirchenbauer, J., Geiping, J., Wen, Y., Katz, J., Miers, I., & Goldstein, T. (2023). A Watermark for Large Language Models. arXiv:2301.10226.",
        "[16] Krishna, K., Song, Y., Karpinska, M., Wieting, J., & Iyyer, M. (2023). Paraphrasing Evades Detectors of AI-Generated Text, but Retrieval is an Effective Defense. arXiv:2303.13408.",
        "[19] Liu, Y., Ott, M., Goyal, N., Du, J., Joshi, M., Chen, D., ... & Stoyanov, V. (2019). RoBERTa: A Robustly Optimized BERT Pretraining Approach. arXiv:1907.11692.",
        "[20] Mitchell, E., Lee, Y., Khazatsky, A., Manning, C. D., & Finn, C. (2023). DetectGPT: Zero-Shot Machine-Generated Text Detection Using Probability Curvature. arXiv:2301.11305.",
        "[21] NewsGuard (2023). The Year AI Supercharged Misinformation: NewsGuard's 2023 in Review.",
        "[27] Sadasivan, V. S., Kumar, A., Balasubramanian, S., Wang, W., & Feizi, S. (2023). Can AI-Generated Text be Reliably Detected? arXiv:2303.11156.",
        "[30] Sun, Z., Shi, T., Guo, X., Liu, D., & Chen, C. (2024). Are We in the AI-Generated Text World Already? Quantifying and Monitoring AIGT on Social Media. arXiv:2412.18148.",
        "[32] Vaswani, A., Shazeer, N., Parmar, N., Uszkoreit, J., Jones, L., Gomez, A. N., ... & Polosukhin, I. (2017). Attention is All You Need. arXiv:1706.03762.",
        "[33] Verizon (2024). 2024 Data Breach Investigations Report.",
    ]

    # Two columns
    col1 = bibliography[:9]
    col2 = bibliography[9:]

    for i, ref in enumerate(col1):
        y = 1.38 + i * 0.56
        # Bracket number in color
        bracket_end = ref.index("]") + 1
        num_part = ref[:bracket_end]
        rest_part = ref[bracket_end:]

        add_textbox(slide, 0.35, y, 0.5, 0.52, num_part,
                    font_size=10, bold=True, color=ACCENT_CYAN, word_wrap=True)
        add_textbox(slide, 0.88, y, 5.6, 0.52, rest_part.strip(),
                    font_size=10, color=TEXT_LIGHT, word_wrap=True)

    for i, ref in enumerate(col2):
        y = 1.38 + i * 0.56
        bracket_end = ref.index("]") + 1
        num_part = ref[:bracket_end]
        rest_part = ref[bracket_end:]

        add_textbox(slide, 6.75, y, 0.5, 0.52, num_part,
                    font_size=10, bold=True, color=ACCENT_CYAN, word_wrap=True)
        add_textbox(slide, 7.28, y, 5.7, 0.52, rest_part.strip(),
                    font_size=10, color=TEXT_LIGHT, word_wrap=True)

    # Vertical separator
    add_rect(slide, 6.55, 1.35, 0.03, 5.6, fill_color=ACCENT_BLUE)


def slide_final(prs):
    """Slide 18 – Final / Mulțumiri."""
    slide = add_slide(prs)
    set_bg(slide, BG_DARK)

    # Top decorative rectangle
    add_rect(slide, 0, 0, 13.33, 2.5, fill_color=BG_HEADER)
    add_rect(slide, 0, 2.48, 13.33, 0.06, fill_color=ACCENT_BLUE)
    add_rect(slide, 0, 2.54, 13.33, 0.02, fill_color=ACCENT_CYAN)

    # Title
    add_textbox(slide, 0.5, 0.5, 12.3, 1.0,
                PRESENTATION_TITLE,
                font_size=28, bold=True, color=TEXT_WHITE,
                align=PP_ALIGN.CENTER, font_name="Calibri")
    add_textbox(slide, 0.5, 1.5, 12.3, 0.55,
                "Bianca-Ștefania GHEORGHE  ·  Coordonator: Conf.dr.ing. Daniel-Marian MEREZEANU",
                font_size=14, color=TEXT_GREY, align=PP_ALIGN.CENTER)

    # Main message
    add_textbox(slide, 1.5, 3.1, 10.3, 1.2,
                "Vă mulțumesc pentru atenție!",
                font_size=46, bold=True, color=TEXT_WHITE,
                align=PP_ALIGN.CENTER, font_name="Calibri")

    add_textbox(slide, 2.0, 4.4, 9.3, 0.65,
                "Sunt disponibilă pentru întrebări.",
                font_size=24, color=ACCENT_CYAN,
                align=PP_ALIGN.CENTER, font_name="Calibri")

    # Key stats reminder
    stats = [
        ("99.17%", "Acuratețe\nRoBERTa"),
        ("45.834", "Exemple\nantrenare"),
        ("6 surse", "Date\nsocial media"),
        ("5", "Contribuții\noriginale"),
    ]
    for i, (val, lbl) in enumerate(stats):
        x = 1.5 + i * 2.65
        add_rect(slide, x, 5.35, 2.3, 1.45, fill_color=BG_CARD,
                 line_color=ACCENT_BLUE, line_width_pt=1.0)
        add_textbox(slide, x+0.1, 5.42, 2.1, 0.72, val,
                    font_size=30, bold=True, color=ACCENT_GREEN,
                    align=PP_ALIGN.CENTER)
        add_textbox(slide, x+0.1, 6.14, 2.1, 0.58, lbl,
                    font_size=12, color=TEXT_LIGHT,
                    align=PP_ALIGN.CENTER, word_wrap=True)

    # Corner decorations
    add_rect(slide, 0, 0, 0.08, 0.5, fill_color=ACCENT_BLUE)
    add_rect(slide, 0, 0, 0.5, 0.08, fill_color=ACCENT_BLUE)
    add_rect(slide, 12.87, 0, 0.08, 0.5, fill_color=ACCENT_CYAN)
    add_rect(slide, 12.85, 0, 0.5, 0.08, fill_color=ACCENT_CYAN)
    add_rect(slide, 0, 7.42, 0.08, 0.5, fill_color=ACCENT_CYAN)
    add_rect(slide, 0, 7.42, 0.5, 0.08, fill_color=ACCENT_CYAN)
    add_rect(slide, 12.87, 7.42, 0.08, 0.5, fill_color=ACCENT_BLUE)
    add_rect(slide, 12.85, 7.42, 0.5, 0.08, fill_color=ACCENT_BLUE)


# ─── BUILD PRESENTATION ───────────────────────────────────────────────────────

def build():
    prs = new_prs()

    slide_title(prs)          # 1
    slide_cuprins(prs)        # 2
    slide_context(prs)        # 3
    slide_nlp_llm(prs)        # 4
    slide_methods(prs)        # 5
    slide_existing_apps(prs)  # 6
    slide_limitations(prs)    # 7
    slide_proposed_method(prs) # 8
    slide_dataset(prs)        # 9
    slide_implementation(prs) # 10
    slide_preprocessing(prs)  # 11
    slide_results(prs)        # 12
    slide_comparison(prs)     # 13
    slide_contributions(prs)  # 14
    slide_use_cases(prs)      # 15
    slide_conclusions(prs)    # 16
    slide_bibliography(prs)   # 17
    slide_final(prs)          # 18

    out_path = "/home/user/licenta/Prezentare_Licenta_Gheorghe_Bianca.pptx"
    prs.save(out_path)
    print(f"Saved: {out_path}")
    print(f"Total slides: {len(prs.slides)}")
    return out_path


if __name__ == "__main__":
    build()

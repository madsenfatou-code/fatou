"""
Génère tous les books FIM SHY-Performance en PPTX
Charte graphique V3 : Architecture premium Teal/Jaune, Calibri
Logo : shy_logo_v2.png — présent sur TOUTES les slides sans exception
Texte : CENTRÉ sur toutes les slides
"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Cm
import os

# ── Couleurs charte ──────────────────────────────────────────────────────────
TEAL       = RGBColor(0x00, 0x80, 0x80)   # #008080
TEAL_DARK  = RGBColor(0x00, 0x4D, 0x4D)   # #004D4D
TEAL_MID   = RGBColor(0x00, 0x66, 0x66)   # #006666
YELLOW     = RGBColor(0xD4, 0xA0, 0x17)   # #D4A017
AMBER      = RGBColor(0xFF, 0x8C, 0x00)   # #FF8C00
GOLD       = RGBColor(0xB8, 0x86, 0x00)
RED        = RGBColor(0xFF, 0x00, 0x00)
WHITE      = RGBColor(0xFF, 0xFF, 0xFF)
DARK       = RGBColor(0x1A, 0x1A, 0x1A)
LGRAY      = RGBColor(0xF0, 0xF7, 0xF7)
MGRAY      = RGBColor(0xCC, 0xE5, 0xE5)
CREAM      = RGBColor(0xFF, 0xFB, 0xF0)

W = Inches(13.33)
H = Inches(7.5)

LOGO_PATH = "/home/user/fatou/shy_logo_v2.png"

INTERDICTION_TEXT = (
    "Toute modification, rectification ou ajout est strictement interdit(e). "
    "Élaboré par Tamou Eljerrari — Ingénieure en Formation — sur ordre de la Directrice Générale, "
    "Mme Salima Negrao. Seules la DG SHY-Performance ou une personne mandatée sont autorisées."
)

# ── Helpers ──────────────────────────────────────────────────────────────────
def new_prs():
    prs = Presentation()
    prs.slide_width  = W
    prs.slide_height = H
    return prs

def blank_slide(prs):
    layout = prs.slide_layouts[6]  # blank
    return prs.slides.add_slide(layout)

def rect(slide, x, y, w, h, fill_rgb):
    shape = slide.shapes.add_shape(1, x, y, w, h)
    shape.line.fill.background()
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_rgb
    return shape

def txbox(slide, text, x, y, w, h,
          size=18, bold=False, color=DARK,
          align=PP_ALIGN.CENTER, italic=False, font="Calibri"):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.name = font
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    return tb

def add_lines(slide, items, x, y, w, size=16, color=DARK, bullet="▸ "):
    tb = slide.shapes.add_textbox(x, y, w, Inches(6))
    tf = tb.text_frame
    tf.word_wrap = True
    first = True
    for item in items:
        if first:
            p = tf.paragraphs[0]
            first = False
        else:
            p = tf.add_paragraph()
        p.alignment = PP_ALIGN.CENTER
        p.space_before = Pt(5)
        run = p.add_run()
        run.text = bullet + item
        run.font.name = "Calibri"
        run.font.size = Pt(size)
        run.font.color.rgb = color
    return tb

def add_lines_centered(slide, items, x, y, w, size=18, color=DARK, bullet="★ "):
    tb = slide.shapes.add_textbox(x, y, w, Inches(5.5))
    tf = tb.text_frame
    tf.word_wrap = True
    first = True
    for item in items:
        if first:
            p = tf.paragraphs[0]
            first = False
        else:
            p = tf.add_paragraph()
        p.alignment = PP_ALIGN.CENTER
        p.space_before = Pt(8)
        run = p.add_run()
        run.text = bullet + item
        run.font.name = "Calibri"
        run.font.size = Pt(size)
        run.font.color.rgb = color
        run.font.bold = True

def add_logo(sl, x=Inches(11.0), y=Inches(0.08), h=Inches(1.0)):
    if os.path.exists(LOGO_PATH):
        aspect = 1081 / 1055
        w = h * aspect
        sl.shapes.add_picture(LOGO_PATH, x, y, width=w, height=h)

def add_interdiction(sl):
    band_y = H - Inches(0.75)
    band_h = Inches(0.38)
    rect(sl, 0, band_y, W, band_h, TEAL_DARK)
    rect(sl, 0, band_y, W, Inches(0.03), RED)
    txbox(sl, INTERDICTION_TEXT,
          Inches(0.3), band_y + Inches(0.03), W - Inches(0.6), band_h - Inches(0.04),
          size=7, color=WHITE, align=PP_ALIGN.CENTER, italic=True)

# ── Slides types ─────────────────────────────────────────────────────────────

def slide_cover(prs, title, subtitle, day_label=""):
    sl = blank_slide(prs)

    # Right panel background (full slide)
    rect(sl, 0, 0, W, H, LGRAY)

    # Left panel — deep teal column
    rect(sl, 0, 0, Inches(5.2), H, TEAL_DARK)

    # Yellow horizontal stripe at top of left panel
    rect(sl, 0, 0, Inches(5.2), Inches(0.5), YELLOW)

    # Logo — large in left panel
    add_logo(sl, x=Inches(0.3), y=Inches(0.6), h=Inches(1.2))

    # Day badge in left panel
    if day_label:
        rect(sl, Inches(0.5), Inches(2.1), Inches(4.2), Inches(0.55), YELLOW)
        txbox(sl, day_label, Inches(0.5), Inches(2.1), Inches(4.2), Inches(0.55),
              size=22, bold=True, color=TEAL_DARK, align=PP_ALIGN.CENTER)

    # Vertical white decorative line in left panel
    rect(sl, Inches(2.4), Inches(2.8), Inches(0.04), Inches(3.0), WHITE)

    # SHY-Performance text in left panel
    txbox(sl, "SHY-Performance", Inches(0.3), Inches(6.2), Inches(4.6), Inches(0.45),
          size=13, color=YELLOW, italic=True, align=PP_ALIGN.CENTER)

    # FIDELIS × UNICEF France in left panel
    txbox(sl, "FIDELIS × UNICEF France", Inches(0.3), Inches(6.55), Inches(4.6), Inches(0.35),
          size=10, color=MGRAY, align=PP_ALIGN.CENTER)

    # Right panel content — small yellow accent bar
    rect(sl, Inches(5.5), Inches(1.0), Inches(7.5), Inches(0.08), YELLOW)

    # Title — right panel
    txbox(sl, title, Inches(5.5), Inches(1.2), Inches(7.5), Inches(2.8),
          size=40, bold=True, color=TEAL_DARK, align=PP_ALIGN.CENTER)

    # Separator line
    rect(sl, Inches(5.5), Inches(4.2), Inches(5.0), Inches(0.05), TEAL)

    # Subtitle
    txbox(sl, subtitle, Inches(5.5), Inches(4.4), Inches(7.5), Inches(1.0),
          size=16, color=DARK, align=PP_ALIGN.CENTER)

    # Formation label badge
    rect(sl, Inches(5.5), Inches(5.55), Inches(4.0), Inches(0.42), TEAL)
    txbox(sl, "Formation Initiale Module", Inches(5.5), Inches(5.55), Inches(4.0), Inches(0.42),
          size=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

    add_interdiction(sl)
    return sl


def slide_section(prs, number, title, items, note="", is_regle_dor=False):
    sl = blank_slide(prs)

    if is_regle_dor:
        # Background warm cream
        rect(sl, 0, 0, W, H, CREAM)

        # Left amber accent bar
        rect(sl, 0, 0, Inches(0.25), H, AMBER)

        # Full-width AMBER header band
        rect(sl, 0, 0, W, Inches(1.8), AMBER)

        # "★ RÈGLE D'OR ★" in header
        txbox(sl, "★  RÈGLE D'OR  ★", Inches(0.3), Inches(0.05), W - Inches(0.6), Inches(0.55),
              size=30, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

        # Title below in header
        txbox(sl, title, Inches(0.5), Inches(0.65), W - Inches(1.0), Inches(0.55),
              size=20, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

        # Logo in header
        add_logo(sl, x=Inches(11.0), y=Inches(0.08), h=Inches(1.0))

        # Content items — centered, bold, teal_dark with star bullet
        add_lines_centered(sl, items, Inches(0.5), Inches(1.9), W - Inches(1.0),
                           size=19, color=TEAL_DARK, bullet="⭐ ")

        if note:
            note_y = H - Inches(1.35)
            rect(sl, Inches(0.5), note_y, Inches(0.08), Inches(0.55), AMBER)
            rect(sl, Inches(0.6), note_y, W - Inches(1.1), Inches(0.55), RGBColor(0xFF, 0xF0, 0xCC))
            txbox(sl, "⭐  " + note, Inches(0.7), note_y + Inches(0.04), W - Inches(1.2), Inches(0.48),
                  size=13, italic=True, color=TEAL_DARK, align=PP_ALIGN.CENTER)

        # Amber footer band
        rect(sl, 0, H - Inches(0.75), W, Inches(0.37), AMBER)
        txbox(sl, "SHY-Performance  ×  FIDELIS  ×  UNICEF France",
              Inches(0.3), H - Inches(0.74), Inches(10), Inches(0.35),
              size=11, color=WHITE, align=PP_ALIGN.CENTER)

    else:
        # White background
        rect(sl, 0, 0, W, H, WHITE)

        # Left accent bar — thin dark teal
        rect(sl, 0, 0, Inches(0.18), H, TEAL_DARK)

        # Header band
        rect(sl, 0, 0, W, Inches(1.55), TEAL)

        # Number badge square
        rect(sl, Inches(0.3), Inches(0.15), Inches(1.2), Inches(1.2), TEAL_DARK)
        txbox(sl, f"{number:02d}", Inches(0.3), Inches(0.15), Inches(1.2), Inches(1.2),
              size=44, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

        # Title in header
        txbox(sl, title, Inches(1.7), Inches(0.25), Inches(9.0), Inches(1.0),
              size=26, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

        # Logo in header
        add_logo(sl, x=Inches(11.0), y=Inches(0.08), h=Inches(1.0))

        # Content area
        add_lines(sl, items, Inches(0.6), Inches(1.65), W - Inches(1.2), size=16, color=DARK)

        if note:
            note_y = H - Inches(1.38)
            # Left teal border accent
            rect(sl, Inches(0.4), note_y, Inches(0.08), Inches(0.55), YELLOW)
            rect(sl, Inches(0.5), note_y, W - Inches(1.0), Inches(0.55), LGRAY)
            txbox(sl, "💡  " + note, Inches(0.6), note_y + Inches(0.04), W - Inches(1.2), Inches(0.48),
                  size=13, italic=True, color=TEAL, align=PP_ALIGN.CENTER)

        # Footer band
        rect(sl, 0, H - Inches(0.75), W, Inches(0.75), TEAL_DARK)
        txbox(sl, "SHY-Performance  ×  FIDELIS  ×  UNICEF France",
              Inches(0.3), H - Inches(0.72), Inches(10), Inches(0.35),
              size=11, color=WHITE, align=PP_ALIGN.CENTER)

    add_interdiction(sl)
    return sl


def slide_two_col(prs, number, title, left_title, left_items, right_title, right_items):
    sl = blank_slide(prs)

    # White background
    rect(sl, 0, 0, W, H, WHITE)

    # Left accent bar
    rect(sl, 0, 0, Inches(0.18), H, TEAL_DARK)

    # Header band
    rect(sl, 0, 0, W, Inches(1.55), TEAL)

    # Number badge
    rect(sl, Inches(0.3), Inches(0.15), Inches(1.2), Inches(1.2), TEAL_DARK)
    txbox(sl, f"{number:02d}", Inches(0.3), Inches(0.15), Inches(1.2), Inches(1.2),
          size=44, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

    # Title in header
    txbox(sl, title, Inches(1.7), Inches(0.25), Inches(9.0), Inches(1.0),
          size=26, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

    # Logo in header
    add_logo(sl, x=Inches(11.0), y=Inches(0.08), h=Inches(1.0))

    # Card left
    card_y = Inches(1.65)
    card_h = Inches(5.1)
    # Shadow
    rect(sl, Inches(0.57), card_y + Inches(0.07), Inches(5.9), card_h, MGRAY)
    # Card bg
    rect(sl, Inches(0.5), card_y, Inches(5.9), card_h, WHITE)
    # Card header TEAL_DARK
    rect(sl, Inches(0.5), card_y, Inches(5.9), Inches(0.55), TEAL_DARK)
    txbox(sl, left_title, Inches(0.5), card_y, Inches(5.9), Inches(0.55),
          size=15, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_lines(sl, left_items, Inches(0.6), card_y + Inches(0.6), Inches(5.7), size=13, color=DARK)

    # Card right
    # Shadow
    rect(sl, Inches(7.07), card_y + Inches(0.07), Inches(5.9), card_h, MGRAY)
    # Card bg
    rect(sl, Inches(7.0), card_y, Inches(5.9), card_h, WHITE)
    # Card header TEAL_MID
    rect(sl, Inches(7.0), card_y, Inches(5.9), Inches(0.55), TEAL_MID)
    txbox(sl, right_title, Inches(7.0), card_y, Inches(5.9), Inches(0.55),
          size=15, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_lines(sl, right_items, Inches(7.1), card_y + Inches(0.6), Inches(5.7), size=13, color=DARK)

    # Footer band
    rect(sl, 0, H - Inches(0.75), W, Inches(0.75), TEAL_DARK)
    txbox(sl, "SHY-Performance  ×  FIDELIS  ×  UNICEF France",
          Inches(0.3), H - Inches(0.72), Inches(10), Inches(0.35),
          size=11, color=WHITE, align=PP_ALIGN.CENTER)

    add_interdiction(sl)
    return sl


def slide_quote(prs, quote, author=""):
    sl = blank_slide(prs)

    # Full TEAL_DARK background
    rect(sl, 0, 0, W, H, TEAL_DARK)

    # Large watermark quote mark behind
    txbox(sl, "❝", Inches(1), -Inches(0.5), Inches(5), Inches(4),
          size=200, color=RGBColor(0x00, 0x60, 0x60), align=PP_ALIGN.CENTER)

    # Left gold vertical stripe
    rect(sl, Inches(0.5), Inches(0.8), Inches(0.12), H - Inches(1.6), YELLOW)

    # Decorative line above author
    rect(sl, Inches(4), Inches(5.2), Inches(5.3), Inches(0.05), YELLOW)

    # Quote text
    txbox(sl, quote, Inches(1.2), Inches(1.5), Inches(11.5), Inches(3.5),
          size=32, bold=True, color=WHITE, italic=True, align=PP_ALIGN.CENTER)

    if author:
        txbox(sl, "— " + author, Inches(1.2), Inches(5.4), Inches(11.5), Inches(0.6),
              size=16, color=YELLOW, align=PP_ALIGN.CENTER)

    # Logo
    add_logo(sl, x=Inches(11.0), y=Inches(0.15), h=Inches(1.0))

    # Footer
    rect(sl, 0, H - Inches(0.75), W, Inches(0.37), TEAL_DARK)
    txbox(sl, "SHY-Performance  ×  FIDELIS  ×  UNICEF France",
          Inches(0.3), H - Inches(0.72), Inches(10), Inches(0.35),
          size=11, color=MGRAY, align=PP_ALIGN.CENTER)

    add_interdiction(sl)
    return sl


def slide_merci(prs, extra_msg=""):
    sl = blank_slide(prs)

    # Background TEAL_DARK
    rect(sl, 0, 0, W, H, TEAL_DARK)

    # Top yellow stripe
    rect(sl, 0, 0, W, Inches(0.45), YELLOW)

    # Bottom yellow stripe
    rect(sl, 0, H - Inches(0.45), W, Inches(0.45), YELLOW)

    # "MERCI"
    txbox(sl, "MERCI", Inches(0.5), Inches(1.8), Inches(12.3), Inches(1.5),
          size=72, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

    # "POUR VOTRE PARTICIPATION"
    txbox(sl, "POUR VOTRE PARTICIPATION", Inches(0.5), Inches(3.2), Inches(12.3), Inches(0.7),
          size=26, color=YELLOW, align=PP_ALIGN.CENTER)

    # Thin yellow line
    rect(sl, Inches(3.5), Inches(3.0), Inches(6.3), Inches(0.06), YELLOW)

    # Sub text
    txbox(sl, "SHY-Performance  ×  FIDELIS  ×  UNICEF France",
          Inches(0.5), Inches(4.0), Inches(12.3), Inches(0.6),
          size=18, color=MGRAY, align=PP_ALIGN.CENTER)

    if extra_msg:
        txbox(sl, extra_msg, Inches(0.5), Inches(4.7), Inches(12.3), Inches(1.5),
              size=16, color=WHITE, align=PP_ALIGN.CENTER, italic=True)

    # Credits
    txbox(sl, "Élaboré par Tamou Eljerrari  |  Sur ordre de Mme Salima Negrao, DG",
          Inches(0.3), Inches(5.2), Inches(12.7), Inches(0.4),
          size=13, color=MGRAY, align=PP_ALIGN.CENTER, italic=True)

    # Logo centered bottom
    add_logo(sl, x=Inches(5.8), y=Inches(5.9), h=Inches(1.1))

    add_interdiction(sl)
    return sl


def slide_felicitations(prs):
    sl = blank_slide(prs)

    # Background TEAL_DARK
    rect(sl, 0, 0, W, H, TEAL_DARK)

    # Yellow stripes top and bottom
    rect(sl, 0, 0, W, Inches(0.5), YELLOW)
    rect(sl, 0, H - Inches(0.5), W, Inches(0.5), YELLOW)

    # Gold stars top row
    txbox(sl, "★  ★  ★  ★  ★", Inches(0.5), Inches(0.6), Inches(12.3), Inches(0.65),
          size=28, color=YELLOW, align=PP_ALIGN.CENTER)

    # "FÉLICITATIONS !"
    txbox(sl, "FÉLICITATIONS !", Inches(0.5), Inches(1.4), Inches(12.3), Inches(1.3),
          size=60, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

    # Yellow separator
    rect(sl, Inches(2.0), Inches(2.85), Inches(9.3), Inches(0.07), YELLOW)

    msg = (
        "Vous avez complété avec succès la Formation Initiale Module (FIM)\n"
        "Fundraiser UNICEF — SHY-Performance × FIDELIS\n\n"
        "Votre engagement, votre sérieux et votre conviction\n"
        "font de vous des fundraisers prêts à changer des vies.\n\n"
        "Chaque appel que vous passerez contribue directement\n"
        "à protéger des enfants à travers le monde.\n\n"
        "L'équipe SHY-Performance est fière de vous.\n"
        "Bonne continuation et beau terrain à toutes et à tous !"
    )
    txbox(sl, msg, Inches(1.0), Inches(3.0), Inches(11.3), Inches(3.0),
          size=16, color=WHITE, align=PP_ALIGN.CENTER, italic=True)

    # Bottom stars
    txbox(sl, "★  ★  ★  ★  ★", Inches(0.5), H - Inches(1.15), Inches(12.3), Inches(0.5),
          size=22, color=YELLOW, align=PP_ALIGN.CENTER)

    # Logo centered bottom
    add_logo(sl, x=Inches(5.8), y=Inches(5.9), h=Inches(1.1))

    add_interdiction(sl)
    return sl


def slide_kpi_table(prs):
    sl = blank_slide(prs)

    # White background
    rect(sl, 0, 0, W, H, WHITE)

    # Left accent bar
    rect(sl, 0, 0, Inches(0.18), H, TEAL_DARK)

    # Header band
    rect(sl, 0, 0, W, Inches(1.55), TEAL)

    # Number badge
    rect(sl, Inches(0.3), Inches(0.15), Inches(1.2), Inches(1.2), TEAL_DARK)
    txbox(sl, "04", Inches(0.3), Inches(0.15), Inches(1.2), Inches(1.2),
          size=44, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

    # Title
    txbox(sl, "Indicateurs de Performance — KPI", Inches(1.7), Inches(0.25),
          Inches(9.0), Inches(1.0), size=26, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

    # Logo in header
    add_logo(sl, x=Inches(11.0), y=Inches(0.08), h=Inches(1.0))

    rows = [
        ("KPI", "Libellé", "Objectif"),
        ("CU/H", "Contacts Utiles par Heure", "9 minimum"),
        ("TX Transfo", "% de PEL/PA parmi les CU", "Cible selon association"),
        ("PDC", "Plan de Charge mensuel", "Volume défini par Fidelis"),
        ("Qualification", "Conformité des qualifications", "50% de la qualité Fidelis"),
        ("PEL", "Prélèvement En Ligne", "Mode prioritaire"),
        ("PA", "Prélèvement Automatique Régulier", "Produit unique de la campagne"),
    ]
    col_w = [Inches(2.0), Inches(5.5), Inches(4.5)]
    col_x = [Inches(0.4), Inches(2.5), Inches(8.1)]
    y0 = Inches(1.58)
    row_h = Inches(0.72)
    for r, row in enumerate(rows):
        y = y0 + r * row_h
        bg = TEAL if r == 0 else (LGRAY if r % 2 == 0 else WHITE)
        fc = WHITE if r == 0 else DARK
        for c, (cell, cw, cx) in enumerate(zip(row, col_w, col_x)):
            rect(sl, cx, y, cw - Inches(0.05), row_h - Inches(0.04), bg)
            txbox(sl, cell, cx + Inches(0.1), y + Inches(0.1),
                  cw - Inches(0.2), row_h - Inches(0.1),
                  size=14 if r > 0 else 15, bold=(r == 0), color=fc, align=PP_ALIGN.CENTER)

    # Footer band
    rect(sl, 0, H - Inches(0.75), W, Inches(0.75), TEAL_DARK)
    txbox(sl, "SHY-Performance  ×  FIDELIS  ×  UNICEF France",
          Inches(0.3), H - Inches(0.72), Inches(10), Inches(0.35),
          size=11, color=WHITE, align=PP_ALIGN.CENTER)

    add_interdiction(sl)
    return sl


def slide_cards(prs, number, title, cards, kpi_bar=None, note=""):
    sl = blank_slide(prs)

    # White background
    rect(sl, 0, 0, W, H, WHITE)

    # Left accent bar
    rect(sl, 0, 0, Inches(0.18), H, TEAL_DARK)

    # Header band
    rect(sl, 0, 0, W, Inches(1.55), TEAL)

    # Number badge
    rect(sl, Inches(0.3), Inches(0.15), Inches(1.2), Inches(1.2), TEAL_DARK)
    txbox(sl, f"{number:02d}", Inches(0.3), Inches(0.15), Inches(1.2), Inches(1.2),
          size=44, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

    # Title
    txbox(sl, title, Inches(1.7), Inches(0.25), Inches(9.0), Inches(1.0),
          size=26, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

    # Logo in header
    add_logo(sl, x=Inches(11.0), y=Inches(0.08), h=Inches(1.0))

    n = len(cards)
    card_w = Inches((13.33 - 1.2) / n - 0.2)
    card_gap = Inches(0.2)
    card_y = Inches(1.65)
    kpi_h = Inches(0.9) if kpi_bar else 0
    note_h = Inches(0.6) if note else 0
    card_h = H - card_y - Inches(0.75) - kpi_h - note_h - Inches(0.1)

    CARD_COLORS = [TEAL, RGBColor(0x00, 0x60, 0x60), RGBColor(0x00, 0x70, 0x70),
                   RGBColor(0x00, 0x50, 0x50), TEAL_DARK]

    for i, card in enumerate(cards):
        x = Inches(0.6) + i * (card_w + card_gap)
        cc = card.get("color", CARD_COLORS[i % len(CARD_COLORS)])
        # Shadow
        rect(sl, x + Inches(0.07), card_y + Inches(0.07), card_w, card_h, MGRAY)
        # Card bg WHITE
        rect(sl, x, card_y, card_w, card_h, WHITE)
        # Card header TEAL_DARK
        rect(sl, x, card_y, card_w, Inches(0.55), cc)
        # Icon in header
        txbox(sl, card.get("icon", "●"), x, card_y + Inches(0.04), card_w, Inches(0.47),
              size=22, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        # Label
        txbox(sl, card["label"], x, card_y + Inches(0.58), card_w, Inches(0.65),
              size=16, bold=True, color=cc, align=PP_ALIGN.CENTER)
        # Separator
        rect(sl, x + Inches(0.3), card_y + Inches(1.25), card_w - Inches(0.6), Inches(0.03), MGRAY)
        # Sub content
        sub = card.get("sub", "")
        if isinstance(sub, list):
            tb = sl.shapes.add_textbox(x + Inches(0.1), card_y + Inches(1.3),
                                       card_w - Inches(0.2), card_h - Inches(1.35))
            tf = tb.text_frame
            tf.word_wrap = True
            first = True
            for item in sub:
                p = tf.paragraphs[0] if first else tf.add_paragraph()
                first = False
                p.alignment = PP_ALIGN.CENTER
                p.space_before = Pt(4)
                run = p.add_run()
                run.text = "▸ " + item
                run.font.name = "Calibri"
                run.font.size = Pt(13)
                run.font.color.rgb = DARK
        else:
            txbox(sl, sub, x + Inches(0.1), card_y + Inches(1.3),
                  card_w - Inches(0.2), card_h - Inches(1.35),
                  size=13, color=DARK, align=PP_ALIGN.CENTER)

    # Bandeau KPI
    if kpi_bar:
        bar_y = H - Inches(0.75) - kpi_h - (note_h if note else 0) - Inches(0.05)
        rect(sl, Inches(0.5), bar_y, W - Inches(1.0), kpi_h, LGRAY)
        rect(sl, Inches(0.5), bar_y, W - Inches(1.0), Inches(0.04), TEAL)
        kpi_w = (W - Inches(1.0)) / len(kpi_bar)
        for j, kpi in enumerate(kpi_bar):
            kx = Inches(0.5) + j * kpi_w
            if j > 0:
                rect(sl, kx, bar_y + Inches(0.1), Inches(0.03), kpi_h - Inches(0.2), MGRAY)
            txbox(sl, kpi.get("icon","") + " " + kpi["value"],
                  kx, bar_y + Inches(0.04), kpi_w, Inches(0.45),
                  size=20, bold=True, color=TEAL, align=PP_ALIGN.CENTER)
            txbox(sl, kpi["label"], kx, bar_y + Inches(0.48), kpi_w, Inches(0.38),
                  size=11, color=TEAL_DARK, align=PP_ALIGN.CENTER)

    if note:
        note_y = H - Inches(0.75) - note_h
        rect(sl, Inches(0.4), note_y, Inches(0.08), note_h - Inches(0.05), YELLOW)
        rect(sl, Inches(0.5), note_y, W - Inches(1.0), note_h - Inches(0.05), LGRAY)
        txbox(sl, "💡  " + note, Inches(0.6), note_y + Inches(0.04),
              W - Inches(1.2), note_h - Inches(0.1),
              size=12, italic=True, color=TEAL, align=PP_ALIGN.CENTER)

    # Footer band
    rect(sl, 0, H - Inches(0.75), W, Inches(0.75), TEAL_DARK)
    txbox(sl, "SHY-Performance  ×  FIDELIS  ×  UNICEF France",
          Inches(0.3), H - Inches(0.72), Inches(10), Inches(0.35),
          size=11, color=WHITE, align=PP_ALIGN.CENTER)

    add_interdiction(sl)
    return sl


# ════════════════════════════════════════════════════════════════════════════════
# BOOK 0 — FILE CONDUCTEUR
# ════════════════════════════════════════════════════════════════════════════════
def build_file_conducteur():
    prs = new_prs()
    slide_cover(prs, "FILE CONDUCTEUR", "Guide d'utilisation du kit de formation FIM\nFormation Initiale Module — Fundraiser UNICEF", "KIT FIM")
    slide_section(prs, 1, "Présentation du Kit de Formation", [
        "7 documents composent ce kit de formation complet",
        "Book 0 — File Conducteur (ce document) : guide d'utilisation globale",
        "Book 1 — Synopsis & Objectifs Pédagogiques : cadre et vision de la formation",
        "Book J1 — Fondations & Mission : KPIs, nomenclature, script accroche",
        "Book J2 — Monde Associatif & Humanitaire : contexte, loi 1901, UNICEF",
        "Book J3 — Analyse et Lecture du Script : décryptage phrase par phrase",
        "Book J4 — Traitement des Objections : 15 objections, méthode AAR",
        "Book J5 — Synthèse Générale & Simulations : évaluation finale, quiz 40Q",
    ])
    slide_cards(prs, 2, "Qui forme ? Qui est formé ?", [
        {
            "icon": "🎯",
            "label": "DONNEUR D'ORDRE",
            "color": TEAL_DARK,
            "sub": ["FIDELIS", "Mandataire officiel", "Campagne UNICEF France", "Pilotage & conformité"]
        },
        {
            "icon": "🏢",
            "label": "OPÉRATEUR TERRAIN",
            "color": TEAL,
            "sub": ["SHY-Performance", "Centre de collecte téléphonique", "Formation & encadrement", "Supervision qualité"]
        },
        {
            "icon": "👥",
            "label": "PUBLIC CIBLE",
            "color": RGBColor(0x00, 0x70, 0x70),
            "sub": ["Fundraisers juniors", "Nouveaux entrants", "Minimum niveau Bac, excellent niveau de langue", "Aucun prérequis"]
        },
    ],
    kpi_bar=[
        {"icon": "📅", "value": "5 JOURS", "label": "Formation intensive"},
        {"icon": "👤", "value": "8 – 14", "label": "Fundraisers par session"},
        {"icon": "⏱", "value": "6h40", "label": "Production effective/jour"},
        {"icon": "🎓", "value": "0", "label": "Prérequis requis"},
    ])
    slide_section(prs, 3, "Comment utiliser ce kit ?", [
        "Chaque book jour est autonome — il peut être utilisé seul pour un rappel ciblé",
        "Le Book 1 Synopsis sert de document de référence permanent",
        "Les slides de quiz (Book J5) sont prêtes à l'emploi pour tout nouvel apprenant",
        "Les jeux de rôle sont signalés par le picto 🎭 — prévoir binômes à l'avance",
        "Les slides 'MANTRA' sont à laisser affichées pendant les exercices",
        "La grille d'observation formateur est à imprimer avant chaque journée J3–J5",
    ])
    slide_section(prs, 4, "Code Couleur & Légende", [
        "🟦 Bandeau TEAL #008080 — slide de contenu standard",
        "🟥 Encart ROUGE — point d'attention critique, règle obligatoire",
        "⬜ Fond blanc / teal clair — exercice pratique ou définition",
        "★ Slide MANTRA — citation à mémoriser, laisser affichée",
        "🎭 Jeu de rôle — exercice en binôme ou groupe",
        "📊 Grille d'évaluation — outil formateur",
        "❓ Quiz — évaluation des connaissances",
    ])
    slide_cards(prs, 5, "Planning Recommandé — 5 Jours de Formation", [
        {
            "icon": "J1",
            "label": "FONDATIONS\n& MISSION",
            "color": TEAL_DARK,
            "sub": ["KPIs & Nomenclature", "Script accroche", "Types d'opérations", "9h30 → 17h30"]
        },
        {
            "icon": "J2",
            "label": "MONDE\nASSOCIATIF",
            "color": TEAL,
            "sub": ["Loi 1901", "Histoire humanitaire", "UNICEF & Charte", "9h30 → 17h30"]
        },
        {
            "icon": "J3",
            "label": "LECTURE\nDU SCRIPT",
            "color": RGBColor(0x00, 0x70, 0x70),
            "sub": ["7 étapes décryptées", "Exercices intensifs", "Jeux de rôle", "9h30 → 17h30"]
        },
        {
            "icon": "J4",
            "label": "TRAITEMENT\nOBJECTIONS",
            "color": RGBColor(0x00, 0x65, 0x65),
            "sub": ["15 objections / 5 cat.", "Méthode AAR", "Simulations terrain", "9h30 → 17h30"]
        },
        {
            "icon": "J5",
            "label": "SYNTHÈSE &\nSIMULATIONS",
            "color": RGBColor(0x00, 0x55, 0x55),
            "sub": ["Quiz 40 questions", "Certifications", "Félicitations", "9h30 → 17h30"]
        },
    ],
    kpi_bar=[
        {"icon": "🌅", "value": "9h30", "label": "Prise de poste + briefing"},
        {"icon": "☕", "value": "11h00 / 15h30", "label": "Pauses matin & après-midi"},
        {"icon": "🍽", "value": "13h00–14h00", "label": "Pause déjeuner"},
        {"icon": "🏁", "value": "17h30", "label": "Fin de journée officielle"},
    ])
    slide_merci(prs)
    prs.save("/home/user/fatou/FIM_Book0_FileConducteur.pptx")
    print("✓ Book 0 — File Conducteur")

# ════════════════════════════════════════════════════════════════════════════════
# BOOK 1 — SYNOPSIS & OBJECTIFS PÉDAGOGIQUES
# ════════════════════════════════════════════════════════════════════════════════
def build_synopsis():
    prs = new_prs()
    slide_cover(prs, "SYNOPSIS &\nOBJECTIFS PÉDAGOGIQUES", "Formation Initiale Module — Fundraiser UNICEF\nSHY-Performance × FIDELIS", "BOOK 1")

    slide_section(prs, 1, "Synopsis de la Formation", [
        "Formation Initiale Module (FIM) de 5 jours destinée aux fundraisers SHY-Performance",
        "Mandatés par FIDELIS pour le compte de l'UNICEF France",
        "Campagne : collecte de dons téléphoniques — Prélèvement Automatique Régulier (PPA)",
        "Cause : lutte contre la malnutrition infantile aigüe sévère",
        "Couvre l'intégralité de la chaîne de compétences : cause → script → objections → closing",
        "Pédagogie active : apport théorique + exercice pratique toutes les 20 minutes maximum",
        "Documents officiels de campagne intégrés : script UNICEF, nomenclature Fidelis, grilles KPI",
    ])

    slide_section(prs, 2, "Programme — Vue d'Ensemble 5 Jours", [
        "J1 — Fondations & Mission : KPIs, nomenclature, types d'opérations, script accroche",
        "J2 — Monde Associatif & Humanitaire : Loi 1901, histoire, UNICEF, marché de la collecte",
        "J3 — Analyse et Lecture du Script : décryptage des 7 étapes, exercices voix haute",
        "J4 — Traitement des Objections : 15 objections, 5 catégories, méthode AAR, jeux de rôle",
        "J5 — Synthèse Générale & Simulations : révision, simulations terrain, quiz 40Q, certifications",
    ])

    slide_two_col(prs, 3, "Objectifs Pédagogiques — SAVOIR & SAVOIR-FAIRE",
        "SAVOIR — Connaissance", [
            "Présenter l'UNICEF, sa mission, 75 ans, 190 pays",
            "Expliquer la malnutrition & les sachets RUTF (92g)",
            "Situer le secteur associatif (Loi 1901, Comité Charte, RGPD, Bloctel)",
            "Citer les arguments de légitimité UNICEF (93% terrain, Nobel 1965)",
            "Définir la nomenclature métier (CU/H, TX Transfo, PDC, PEL, PA)",
        ],
        "SAVOIR-FAIRE — Technique d'appel", [
            "Appliquer les 5 règles officielles de l'accroche",
            "Conduire le script complet en 7 étapes chronométrées",
            "Qualifier chaque appel conformément à la nomenclature Fidelis",
            "Gérer les 4 modes de validation du don (IBAN / PA ligne / promesse / courrier)",
            "Traiter les 15 objections types par la méthode AAR",
        ])

    slide_section(prs, 4, "Objectifs Pédagogiques — SAVOIR-ÊTRE", [
        "Maintenir un ton chaleureux, souriant et naturel tout au long de l'appel",
        "Gérer les objections sensibles (accent, source numéro, accusations d'arnaque) avec calme",
        "Transmettre une conviction sincère et personnelle pour la mission UNICEF",
        "Rebondir après un refus sans perdre énergie ni conviction",
        "Respecter le prospect en toutes circonstances — jamais d'insistance agressive",
        "Intégrer et vivre le mantra : 'Le refus d'aujourd'hui peut être le don de demain'",
    ])

    slide_kpi_table(prs)

    slide_section(prs, 5, "Population Cible & Architecture Pédagogique", [
        "Public : fundraisers juniors ou en reconversion, minimum niveau Bac, excellent niveau de langue",
        "Effectif : 8 à 14 fundraisers par session (qualité des jeux de rôle garantie)",
        "Prérequis : aucune expérience en collecte de dons requise",
        "Dimension Opérationnelle (50%) : script, grille AAR, quiz, mises en situation",
        "Dimension Relationnelle (35%) : jeux de rôle binômes, grille observation, débriefs",
        "Dimension Motivationnelle (15%) : chiffres-chocs, storytelling terrain, personas",
    ])

    slide_merci(prs)
    prs.save("/home/user/fatou/FIM_Book1_Synopsis_Objectifs.pptx")
    print("✓ Book 1 — Synopsis & Objectifs")

# ════════════════════════════════════════════════════════════════════════════════
# BOOK J1 — FONDATIONS & MISSION
# ════════════════════════════════════════════════════════════════════════════════
def build_j1():
    prs = new_prs()
    slide_cover(prs, "JOUR 1\nFONDATIONS & MISSION", "Mission, KPIs, Nomenclature & Script d'Accroche\nSHY-Performance × FIDELIS × UNICEF France", "JOUR 1")

    slide_section(prs, 1, "Mission & Finalité de la Campagne", [
        "Finalité officielle : obtenir des prélèvements automatiques réguliers (PA)",
        "Permettre à l'UNICEF de CONTINUER son action humanitaire sur le terrain",
        "AUGMENTER durablement le taux de dons pour la lutte contre la malnutrition",
        "L'UNICEF intervient dans 190 pays — 75 ans d'action humanitaire depuis 1946",
        "Chaque don finance directement : sachets RUTF, vaccins, eau potable, protection",
        "Campagne opérée par SHY-Performance, mandatée par FIDELIS pour l'UNICEF France",
    ])

    slide_section(prs, 2, "Nomenclature Métier — Définitions (1/2)", [
        "CU — Contact Utile : appel de plus de 60 secondes correctement qualifié",
        "CU/H — Contacts Utiles par Heure : mesure de la productivité horaire",
        "TX Transfo — Taux de Transformation : % de PEL ou PA parmi les CU",
        "PDC — Plan de Charge : volume mensuel de CU à produire",
        "PA — Prélèvement Automatique : don mensuel régulier du donateur",
        "PEL — Prélèvement En Ligne : validation PA via formulaire en ligne",
        "DON — qualification : le contact accepte le prélèvement automatique régulier",
    ], note="Maîtriser ces termes avant le premier appel — évalués au Quiz J5")

    slide_section(prs, 3, "Nomenclature Métier — Définitions (2/2)", [
        "INDÉCIS — qualification : contact intéressé mais n'a pas validé son don",
        "REFUS — qualification : contact ne souhaite pas soutenir l'UNICEF",
        "NON CONTACTÉ — qualification : appel < 60 secondes ou pas d'identification réelle",
        "ABSENT — qualification : répondeur, NRP, raccroché sans identification",
        "CONQUÊTE — appeler pour la 1ère fois des prospects non encore donateurs",
        "FIDÉLISATION / RÉACTIVATION — relancer d'anciens donateurs ou contacts indécis",
    ], note="50% de la qualité Fidelis repose sur une qualification CONFORME")

    slide_section(prs, 4, "KPI 1 — Contacts Utiles par Heure (CU/H)", [
        "Définition : nombre de contacts utiles réalisés en une heure de production",
        "Objectif minimum : 9 CU/H — seuil de performance attendu par Fidelis",
        "Mesure la productivité horaire du fundraiser",
        "Un appel est Utile (CU) s'il dure plus de 60 secondes et est correctement qualifié",
        "En dessous de 9 CU/H → identifier les blocages : tempo, transitions, décrochage",
        "Suivi en temps réel sur l'outil CRM — visible par le superviseur",
    ], note="9 CU/H minimum — sous ce seuil, un accompagnement superviseur est déclenché")

    slide_section(prs, 5, "KPI 2 & 3 — Taux de Transformation & Plan de Charge", [
        "TX Transfo : % de PEL ou PA obtenus parmi les Contacts Utiles",
        "TX Transfo cible : défini par l'association selon la campagne en cours",
        "Un TX Transfo élevé = argumentation efficace + bon traitement des objections",
        "PDC — Plan de Charge : volume mensuel global de CU à produire par l'équipe",
        "Le PDC est calculé par Fidelis en fonction des objectifs campagne",
        "Pilotage collectif : chaque fundraiser contribue au PDC de l'équipe",
    ])

    slide_cards(prs, 6, "Les 2 Types d'Opérations", [
        {
            "icon": "🚀",
            "label": "LA CONQUÊTE",
            "color": TEAL_DARK,
            "sub": [
                "Appeler pour la 1ère fois des prospects",
                "Présenter les missions de l'UNICEF",
                "Convaincre d'adhérer à la cause",
                "Inciter au soutien : PA en ligne ou DON",
                "Dossier : conquête PA ou conquête DON",
            ]
        },
        {
            "icon": "🤝",
            "label": "FIDÉLISATION / RÉACTIVATION",
            "color": RGBColor(0x00, 0x65, 0x65),
            "sub": [
                "Relancer d'anciens donateurs ou contacts indécis",
                "Maintenir et renforcer l'engagement existant",
                "Réactiver les donateurs inactifs",
                "Proposer une augmentation du montant du don",
                "Dossier : fidélisation ou réactivation",
            ]
        },
    ],
    kpi_bar=[
        {"icon": "🎯", "value": "CONQUÊTE", "label": "Nouveaux donateurs à convertir"},
        {"icon": "♻️", "value": "RÉACTIVATION", "label": "Anciens donateurs à relancer"},
        {"icon": "💡", "value": "OBJECTIF COMMUN", "label": "Prélèvement Automatique Régulier"},
    ])

    # RÈGLE D'OR slide
    slide_section(prs, 7, "Le Produit — RÈGLE D'OR DE LA CAMPAGNE", [
        "SEUL le prélèvement automatique régulier (PA) est proposé",
        "Aucun don ponctuel — uniquement le PA mensuel régulier",
        "MODE A — Prélèvement À CHAUD : IBAN recueilli en direct lors de l'appel",
        "MODE B — PA en ligne en direct : guider le donateur sur unicef.fr",
        "MODE C — Promesse PA en ligne : engagement + rappel prévu",
        "MODE D — PA Courrier (exceptionnel) : autorisation par courrier",
        "Paliers proposés : 10€ / 15€ / 20€ par mois",
    ], note="Rappel déduction fiscale : 15€/mois = 3,75€ réels après déduction à 75%",
    is_regle_dor=True)

    # RÈGLE D'OR — Les 5 règles de l'accroche
    slide_section(prs, 8, "Les 5 Règles Officielles de l'Accroche — RÈGLE D'OR", [
        "RÈGLE 1 — Dire 'Allô ?' et attendre la réponse : ne jamais parler en premier",
        "RÈGLE 2 — Identifier le genre : voix masculine → M. / voix féminine → Mme",
        "RÈGLE 3 — Annoncer l'enregistrement de l'appel à des fins de qualité et formation",
        "RÈGLE 4 — Se présenter : prénom + pour l'UNICEF",
        "RÈGLE 5 — Accroche cause : 'Je suppose que vous connaissez bien l'UNICEF ?'",
    ], note="Ces 5 règles sont OBLIGATOIRES — leur non-respect constitue une faute qualité",
    is_regle_dor=True)

    slide_section(prs, 9, "Organisation de la Journée de Production", [
        "9h30 — Prise de poste + briefing équipe",
        "9h45 — Début de la production téléphonique",
        "11h00 — Pause matin (15 min)",
        "13h00 — Pause déjeuner (1 heure)",
        "14h00 — Reprise de la production",
        "15h30 — Pause après-midi (15 min)",
        "17h15 — Fin de production + débriefing superviseur",
        "17h30 — Fin de journée officielle | Durée production effective : 6h40",
    ])

    slide_quote(prs,
        "Le refus d'aujourd'hui\npeut être le don de demain.",
        "Mantra Fundraiser SHY-Performance")

    slide_merci(prs)
    prs.save("/home/user/fatou/FIM_BookJ1_Fondations_Mission.pptx")
    print("✓ Book J1 — Fondations & Mission")

# ════════════════════════════════════════════════════════════════════════════════
# BOOK J2 — MONDE ASSOCIATIF & HUMANITAIRE
# ════════════════════════════════════════════════════════════════════════════════
def build_j2():
    prs = new_prs()
    slide_cover(prs, "JOUR 2\nMONDE ASSOCIATIF\n& HUMANITAIRE", "Comprendre pour Convaincre — Loi 1901, Histoire, UNICEF\nSHY-Performance × FIDELIS × UNICEF France", "JOUR 2")

    slide_section(prs, 1, "La France Associative", [
        "1,5 million d'associations actives en France",
        "1,8 million de salariés employés par le secteur associatif",
        "22 millions de bénévoles mobilisés chaque année",
        "Le secteur associatif représente 3,5% du PIB français",
        "Chaque citoyen a bénéficié directement ou indirectement d'une action associative",
        "Santé, éducation, sport, culture, humanitaire : tous les domaines sont couverts",
    ], note="Pensez-y : les Restos du Cœur servent 170 millions de repas/an")

    slide_section(prs, 2, "Définition Juridique — Loi du 1er Juillet 1901", [
        "Article 1 : 'Convention par laquelle deux ou plusieurs personnes mettent en commun"
        " leurs connaissances ou activité dans un but autre que de partager des bénéfices'",
        "Caractère désintéressé : aucun profit distribué aux membres",
        "But très divers : sportif, humanitaire, défense d'intérêts, promotion d'idées",
        "Association de FAIT : non déclarée, n'existe pas aux yeux de l'administration",
        "Association DÉCLARÉE : enregistrée officiellement — droits complets (subventions, dons)",
        "Association RECONNUE D'UTILITÉ PUBLIQUE : agrément État, collecte de dons autorisée",
    ], note="L'UNICEF France est reconnue d'utilité publique — légitimité maximale")

    slide_section(prs, 3, "Histoire de l'Humanitaire — Des Origines à l'UNICEF", [
        "Moyen Âge : premières associations caritatives d'émanation religieuse",
        "24 juin 1859 — Bataille de Solférino : Henry Dunant vient en aide aux blessés",
        "1863 — Naissance de la Croix-Rouge et du droit humanitaire international",
        "1901 — Loi française sur les associations : cadre légal moderne",
        "1945 — Fin de la 2ème Guerre Mondiale : 13 millions d'enfants déplacés en Europe",
        "11 décembre 1946 — Création de l'UNICEF par l'Assemblée Générale de l'ONU",
        "1965 — L'UNICEF reçoit le Prix Nobel de la Paix",
    ])

    slide_section(prs, 4, "Les Principes Fondamentaux de l'Humanitaire", [
        "HUMANITÉ : soulager la souffrance humaine sans discrimination",
        "IMPARTIALITÉ : aide accordée selon les besoins, sans distinction de nationalité, race ou religion",
        "NEUTRALITÉ : ne pas prendre parti dans les conflits politiques ou militaires",
        "INDÉPENDANCE : autonomie vis-à-vis des autorités politiques, économiques ou militaires",
        "UNIVERSALITÉ : mouvement mondial — chaque organisation est égale en droits et devoirs",
        "Ces principes guident l'action de l'UNICEF dans les 190 pays où il intervient",
    ])

    slide_section(prs, 5, "Évolution du Marché de la Collecte", [
        "AVANT 1990 — Plein boom : ONG reconnues, collecte facile, confiance totale, mailing dominant",
        "1989 — Création du Comité de la Charte du Don en Confiance",
        "1996 — Affaire ARC : crise de confiance majeure dans le secteur",
        "1996 — Renforcement des contrôles : audits obligatoires, transparence totale",
        "2000s — Montée du télémarketing et du marketing digital",
        "2010s — RGPD, Bloctel : nouvelles réglementations encadrant la prospection",
        "Aujourd'hui : label 'Don en Confiance' = gage de sérieux pour les donateurs",
    ])

    slide_section(prs, 6, "Impact Mondial — L'UNICEF en Action", [
        "VACCINATION : l'UNICEF vaccine 45% des enfants dans le monde",
        "2022 : 2,2 milliards de vaccins distribués — la polio quasi éradiquée",
        "EAU POTABLE : accès à l'eau saine pour des millions d'enfants en zones de crise",
        "NUTRITION : distribution de sachets RUTF pour la malnutrition aigüe sévère",
        "ÉDUCATION : programmes scolaires dans les zones de conflit et de catastrophe",
        "PROTECTION : lutte contre le travail des enfants, les mariages forcés, le recrutement armé",
        "93% des fonds collectés vont directement aux programmes terrain",
    ], note="Chaque 3 secondes, un enfant reçoit un vaccin grâce à l'UNICEF")

    slide_section(prs, 7, "Le Comité de la Charte & Le Don en Confiance", [
        "Créé en 1989, renforcé après l'Affaire ARC en 1996",
        "Mission : renforcer le contrôle et la transparence des associations collectant des dons",
        "Audit obligatoire des comptes + publication annuelle des rapports financiers",
        "Label 'Don en Confiance' : garantie pour le donateur que l'association est sérieuse",
        "UNICEF France en est membre — argument de légitimité clé face aux objections",
        "Commissaire aux comptes indépendant certifie les comptes d'UNICEF France",
        "Utiliser ce label comme réponse à l'objection 'Je ne fais pas confiance aux associations'",
    ])

    # RÈGLE D'OR — Comprendre pour convaincre
    slide_section(prs, 8, "RÈGLE D'OR — Pourquoi le Don Régulier est Vital", [
        "COMPRENDRE → S'ENGAGER → CONVAINCRE : le fundraiser qui comprend convainc mieux",
        "Le don régulier permet de planifier les actions humanitaires sur le long terme",
        "La faim, la maladie, le manque de soins ne s'arrêtent pas après une aide ponctuelle",
        "Un don de 10€/mois = accompagner un enfant durant toutes les étapes de sa guérison",
        "Flexible : le donateur peut modifier ou suspendre son don à tout moment",
        "Impact cumulatif : 10€ × 12 mois = 120€/an soit 4 mois complets de nutrition",
        "Argument : 'Sans votre soutien régulier, nos équipes ne peuvent pas planifier'",
    ], is_regle_dor=True)

    slide_quote(prs,
        "Comprendre pour Convaincre.\nChaque citoyen est concerné.",
        "Mantra J2 — SHY-Performance")

    slide_merci(prs)
    prs.save("/home/user/fatou/FIM_BookJ2_MondeAssociatif.pptx")
    print("✓ Book J2 — Monde Associatif & Humanitaire")

# ════════════════════════════════════════════════════════════════════════════════
# BOOK J3 — ANALYSE ET LECTURE DU SCRIPT
# ════════════════════════════════════════════════════════════════════════════════
def build_j3():
    prs = new_prs()
    slide_cover(prs, "JOUR 3\nANALYSE ET LECTURE\nDU SCRIPT", "Décryptage phrase par phrase — Les 7 étapes officielles UNICEF\nSHY-Performance × FIDELIS × UNICEF France", "JOUR 3")

    slide_section(prs, 1, "Structure du Script — Les 7 Étapes", [
        "ÉTAPE 1 — Accroche & Présentation : identification, annonce enregistrement, présentation UNICEF",
        "ÉTAPE 2 — La Malnutrition Infantile : contextualiser la cause, créer l'émotion",
        "ÉTAPE 3 — L'Aliment Thérapeutique (RUTF) : les sachets de 92g, traitement de référence",
        "ÉTAPE 4 — Appel au Soutien : proposition 15€/mois = 50 centimes/jour",
        "ÉTAPE 5 — Si préfère don ponctuel : reconversion vers le don régulier, flexibilité",
        "ÉTAPE 6 — Indécis / Souhaite réfléchir : valoriser la bonne volonté, reprendre",
        "ÉTAPE 7 — Validation des coordonnées : IBAN / PA ligne / Promesse / Courrier",
    ])

    slide_section(prs, 2, "Étape 1 — Accroche & Présentation (Analyse)", [
        "'Allô ?' → attendre la réponse — identifier le genre avant de continuer",
        "'Je vous appelle dans le cadre d'une campagne d'information humanitaire'",
        "'Cet appel est enregistré à des fins de qualité et de formation'",
        "Question miroir : 'Je suppose que vous connaissez bien sûr l'UNICEF ?'",
        "SI OUI : 'Très bien ! L'UNICEF intervient dans plus de 190 pays...'",
        "SI NON : définir l'UNICEF en 2 phrases — soins, éducation, protection en période de crise",
        "Ton : chaleureux, souriant, jamais récité — le prospect entend si on lit",
    ], note="🎭 Exercice : lire l'étape 1 à voix haute × 3 en binôme — feedback sur le naturel")

    slide_section(prs, 3, "Étape 2 — La Malnutrition Infantile (Analyse)", [
        "'Plus d'un million d'enfants de 3 mois à 5 ans confrontés à une mort évitable'",
        "Mettre en avant les 3 causes de propagation : catastrophes, conflits, guerres",
        "'Les enfants sont malheureusement les premières victimes'",
        "Question de connivence : 'En avez-vous déjà entendu parler Mr/Mme XX ?'",
        "But : créer un pont émotionnel avant la solution (RUTF)",
        "Pause après la question — laisser le prospect répondre, ne pas interrompre",
        "Adapter le ton : grave mais porteur d'espoir — jamais alarmiste au point de bloquer",
    ], note="🎭 Exercice : lire l'étape 2 — identifier les mots-clés émotionnels à accentuer")

    slide_section(prs, 4, "Étape 3 — Les Sachets RUTF (Analyse)", [
        "'Ces petits sachets de 92 grammes peuvent littéralement sauver des vies'",
        "Composition : pâte d'arachide, matière grasse végétale, lait écrémé en poudre, vitamines",
        "Avantage clé : sans électricité, sans cuisson, sans eau — idéal en zones de crise",
        "'Ce formidable traitement a amélioré spectaculairement la prise en charge'",
        "'Notre engagement est permanent... mais il reste du chemin à parcourir'",
        "Transition vers la proposition : créer le sentiment de continuité nécessaire",
        "Ton : conviction, fierté de l'action, pas de complaisance — montrer que le besoin persiste",
    ], note="🎭 Exercice : présenter les RUTF sans regarder le script — 60 secondes chrono")

    slide_section(prs, 5, "Étape 4 — Appel au Soutien (Analyse)", [
        "'J'aimerais vous inviter à apporter votre pierre à l'édifice'",
        "Ancrage prix : '50 centimes d'euros par jour soit 15€/mois'",
        "Déduction fiscale : 'cela ne vous reviendra qu'à 3,75€ après déduction fiscale'",
        "Question de closing : 'Alors peut-on compter sur votre contribution M/Mme ?'",
        "Si OUI ferme → verrouillage : 'Je peux enregistrer un don régulier de X€ par mois ?'",
        "Si OUI hésitant → budget prévisionnel : 'Ceci nous permet de planifier nos actions'",
        "Paliers : 10€ / 15€ / 20€ — jamais descendre sans accord explicite du donateur",
    ], note="🎭 Exercice en binôme : jouer l'étape 4 — un fundraiser, un prospect hésitant")

    slide_section(prs, 6, "Étapes 5 & 6 — Ponctuel & Indécis (Analyse)", [
        "DON PONCTUEL — ne pas refuser : 'C'est déjà formidable !'",
        "Reconversion : 'Ce qui est génial avec le don mensuel, c'est qu'il est totalement flexible'",
        "Argument durée : 'accompagner un enfant durant toutes les étapes de sa guérison'",
        "INDÉCIS — valoriser : 'Vous ne nous fermez pas la porte, cela témoigne de votre sensibilité'",
        "Reproposer avec un montant plus accessible : 10€ → 2,50€ après déduction fiscale",
        "Question de relance : 'Seriez-vous prêt(e) à poser ce petit geste régulier ?'",
        "Si toujours indécis → classer INDÉCIS 'doit réfléchir' — respecter la décision",
    ], note="🎭 Exercice : jeu de rôle étape 5 — prospect qui préfère le don ponctuel")

    slide_section(prs, 7, "Étape 7 — Validation des Coordonnées (Analyse)", [
        "Annonce prélèvement : '1er prélèvement le 10 du mois prochain, puis le 10 de chaque mois'",
        "MODE 1 — PA À CHAUD : 'J'ai besoin de votre IBAN — série de 27 chiffres commençant par FR'",
        "Blindage : 'M'autorisez-vous à enregistrer un prélèvement mensuel de X€ pour l'UNICEF ?'",
        "MODE 2 — PA LIGNE EN DIRECT : envoyer le lien, guider, rester en ligne, valider ensemble",
        "MODE 3 — PROMESSE PA LIGNE : fixer un moment précis + rappel planifié",
        "MODE 4 — PA COURRIER (exceptionnel) : engagement ferme de renvoyer l'autorisation",
        "Clôture : 'Merci infiniment — votre soutien va changer concrètement la vie d'enfants'",
    ], note="🎭 Exercice : simulation complète étape 7 — traiter la méfiance IBAN")

    # RÈGLE D'OR — Script
    slide_section(prs, 8, "RÈGLE D'OR — Exercice de Lecture du Script", [
        "🎭 EXERCICE 1 : Lecture silencieuse du script complet (10 minutes)",
        "🎭 EXERCICE 2 : Lecture à voix haute individuelle — 1 étape par fundraiser",
        "🎭 EXERCICE 3 : Jeu de rôle complet en binôme — script de bout en bout",
        "Grille d'observation : accroche / naturel / écoute / proposition / closing",
        "🎭 EXERCICE 4 : Appel chronométré — objectif : script complet en moins de 5 minutes",
        "Débrief collectif : points forts identifiés + 1 axe d'amélioration par fundraiser",
        "Auto-évaluation : chaque fundraiser note ses propres points de vigilance",
    ], is_regle_dor=True)

    slide_quote(prs,
        "Votre voix, votre conviction,\nvotre argument — c'est ce qui transforme\nune hésitation en don concret.",
        "Script UNICEF — SHY-Performance")

    slide_merci(prs)
    prs.save("/home/user/fatou/FIM_BookJ3_Lecture_Script.pptx")
    print("✓ Book J3 — Analyse et Lecture du Script")

# ════════════════════════════════════════════════════════════════════════════════
# BOOK J4 — TRAITEMENT DES OBJECTIONS
# ════════════════════════════════════════════════════════════════════════════════
def build_j4():
    prs = new_prs()
    slide_cover(prs, "JOUR 4\nTRAITEMENT\nDES OBJECTIONS", "15 Objections — 5 Catégories — Méthode AAR\nSHY-Performance × FIDELIS × UNICEF France", "JOUR 4")

    # RÈGLE D'OR — Méthode AAR
    slide_section(prs, 1, "RÈGLE D'OR — La Méthode AAR", [
        "AAR = Accuser réception / Argumenter / Relancer",
        "A — ACCUSER RÉCEPTION : montrer que l'on a entendu et compris l'objection",
        "Exemples : 'Je comprends tout à fait...' / 'C'est tout à fait légitime...'",
        "A — ARGUMENTER : répondre avec un argument factuel ET émotionnel",
        "R — RELANCER : reposer une question orientée vers l'accord",
        "Exemples de relance : 'Alors, est-ce que 10€ par mois vous semble accessible ?'",
        "Règle d'or : ne jamais contredire frontalement — toujours amortir avant d'argumenter",
    ], note="La méthode AAR s'applique à TOUTES les objections sans exception",
    is_regle_dor=True)

    slide_section(prs, 2, "Catégorie 1 — Objections Initiales (Phase d'Accroche)", [
        "OBJ 1 : 'Pas intéressé / Je n'ai pas le temps' → 'Permettez-moi de vous expliquer rapidement...'",
        "OBJ 2 : 'Faux numéro' → 'Mon appel n'est pas nominatif — campagne nationale d'information'",
        "OBJ 3 : 'Vous m'appelez pour un don ?' → 'Avant tout, je vous contacte pour avoir votre avis'",
        "OBJ 4 : 'Je donne déjà ailleurs' → 'Je vous remercie — il ne s'agit pas seulement d'un don'",
        "OBJ 5 : 'C'est une arnaque !' → 'Je vous appelle bien de la part d'UNICEF France — unicef.fr'",
        "OBJ 6 : 'Je ne donne pas par téléphone' → 'Il n'est pas question de donner de l'argent par téléphone'",
        "OBJ 7 : 'Pas confiance aux associations' → citer Comité de la Charte, 75 ans, Nobel 1965",
        "OBJ 8 : 'Je n'aime pas être contacté' → 'Message important à diffuser — moyen le plus efficace'",
    ])

    slide_section(prs, 3, "Catégorie 2 — Objections Financières", [
        "OBJ 9 : 'Je donne déjà à d'autres associations' → ne pas modifier habitudes + proposer 10€",
        "Argument : '300 millions de personnes en insécurité alimentaire — chaque don compte'",
        "OBJ 10 : 'Je n'ai pas les moyens' → '15€ était un exemple — que pensez-vous de 10€ ?'",
        "Argument : '10€ = 2,50€ après déduction fiscale — pour 4 enfants pris en charge'",
        "Dernier positionnement : 'Une personne meurt de faim toutes les 13 secondes'",
        "Descendre à 8€ ou 6€ si nécessaire — '6€ = 1,50€ après déduction fiscale'",
        "Règle : ne jamais proposer moins de 6€ sans accord explicite du superviseur",
    ], note="Toujours ramener le montant à son coût réel après déduction fiscale de 75%")

    slide_section(prs, 4, "Catégorie 3 — Contre le Don Régulier (PA)", [
        "OBJ 11 : 'Je préfère un don ponctuel' → 'C'est formidable ! Mais le régulier sauve sur le long terme'",
        "Argument : 'La faim ne s'arrête pas après une aide ponctuelle'",
        "Argument flexibilité : 'Vous pouvez stopper ou modifier le montant à tout moment'",
        "OBJ 12 : 'Je n'aime pas l'engagement mensuel' → 'Ce n'est pas un engagement rigide'",
        "Argument : 'Vous gardez un contrôle total — un appel suffit pour suspendre'",
        "Dernier positionnement PA : 'Des milliers d'enfants dont la plupart ont moins de 5 ans'",
        "Question de relance finale : '8€ ou même 6€ par mois vous semblent envisageables ?'",
    ])

    slide_section(prs, 5, "Catégorie 4 — Méfiance IBAN", [
        "OBJ 13 : 'Je ne veux pas donner mon IBAN' → rassurer avec précision",
        "'Votre IBAN ne permet PAS d'accéder à votre compte — uniquement un virement SEPA'",
        "'Sans accord verbal clair de votre part, rien n'est possible'",
        "'Votre IBAN figure déjà sur vos factures d'énergie, internet, mutuelle, CAF...'",
        "'Aucun prélèvement aujourd'hui — 1er prélèvement le 10 du mois prochain'",
        "'Procédure encadrée accordée exclusivement aux organisations françaises reconnues'",
        "Si toujours non → orienter vers PA en ligne en direct sur unicef.fr",
    ], note="Ne jamais recueillir l'IBAN sans la présence du superviseur + coupure enregistrement")

    slide_section(prs, 6, "Catégorie 5 — Objections Conflictuelles", [
        "OBJ 14a : 'D'où m'appelez-vous ?' → 'De Paris — 6ème arrondissement, siège UNICEF France'",
        "OBJ 14b : 'D'où avez-vous mon numéro ?' → citer la source (partenaire / annuaire)",
        "OBJ 14c : 'Vous n'avez pas le droit de m'appeler' → informer sur l'inscription annuaire",
        "OBJ 14d : 'Je suis sur liste Bloctel' → 'Les associations à but non lucratif sont autorisées'",
        "OBJ 14e : 'Je suis sur liste rouge' → s'excuser, remonter l'info, prendre congé",
        "OBJ 14f : 'Vous avez un accent !' → assumer avec humour, enchaîner sur le script",
        "OBJ 14g : 'Je connais un membre de l'association' → valoriser + poursuivre ou prendre congé",
        "OBJ 14h/i : 'Qui êtes-vous ?' → 'Mandaté par l'UNICEF — vérifiable sur unicef.fr'",
    ], note="Face à une situation conflictuelle : SOURIRE + CALME + BLOC — jamais d'impatience")

    slide_section(prs, 7, "Grille d'Observation — Jeux de Rôle J4", [
        "CRITÈRE 1 — Méthode AAR appliquée : accuser réception / argumenter / relancer",
        "CRITÈRE 2 — Ton : chaleur maintenue même face à une objection agressive",
        "CRITÈRE 3 — Naturel : pas de récitation mécanique de la réponse",
        "CRITÈRE 4 — Argumentation : fait + émotion + chiffre concret",
        "CRITÈRE 5 — Relance : question orientée vers l'accord après l'argument",
        "CRITÈRE 6 — Gestion du refus : respect de la décision finale du prospect",
        "Score : 1 (à travailler) / 2 (en progrès) / 3 (maîtrisé) par critère",
    ])

    # RÈGLE D'OR — Programme jeux de rôle
    slide_section(prs, 8, "RÈGLE D'OR — Programme Jeux de Rôle Journée J4", [
        "🎭 MATIN : Jeux de rôle catégories 1 & 2 (objections initiales + financières)",
        "Binômes tournants : chaque fundraiser joue les 2 rôles (fundraiser + prospect)",
        "Formateur observe 2 binômes simultanément avec la grille",
        "Débrief à mi-journée : retours individuels + points de vigilance collectifs",
        "🎭 APRÈS-MIDI : Jeux de rôle catégories 3, 4 & 5 (PA + IBAN + conflictuelles)",
        "Simulation d'appels complets : script entier + objection tirée au sort",
        "Débrief final : 3 points forts collectifs + 2 axes d'amélioration prioritaires",
    ], is_regle_dor=True)

    slide_quote(prs,
        "Chaque objection est une invitation\nà convaincre davantage.",
        "Méthode AAR — SHY-Performance")

    slide_merci(prs)
    prs.save("/home/user/fatou/FIM_BookJ4_Objections.pptx")
    print("✓ Book J4 — Traitement des Objections")

# ════════════════════════════════════════════════════════════════════════════════
# BOOK J5 — SYNTHÈSE & SIMULATIONS
# ════════════════════════════════════════════════════════════════════════════════
def build_j5():
    prs = new_prs()
    slide_cover(prs, "JOUR 5\nSYNTHÈSE GÉNÉRALE\n& SIMULATIONS", "Révision, Simulations Terrain, Quiz 40Q & Certifications\nSHY-Performance × FIDELIS × UNICEF France", "JOUR 5")

    slide_section(prs, 1, "Programme de la Journée J5", [
        "9h30 — Révision express des 5 modules (J1 → J4) : questions flash en groupe",
        "10h00 — Simulations d'appels complets chronométrés en conditions réelles",
        "11h00 — Pause + débriefing des simulations",
        "11h15 — Jeux de rôle finaux : script + objection tirée au sort",
        "13h00 — Pause déjeuner",
        "14h00 — Quiz d'évaluation finale des connaissances (40 questions)",
        "15h30 — Correction collective + débriefing pédagogique",
        "16h30 — Remise des certifications de formation",
        "17h00 — Mot de clôture + félicitations + vœux de bonne continuation",
    ])

    slide_section(prs, 2, "Révision Flash — J1 à J4", [
        "J1 — Fondations : KPI objectif CU/H ? → 9 minimum | Produit unique ? → PA Régulier",
        "J1 — Nomenclature : CU / TX Transfo / PDC / PEL / PA — définitions express",
        "J2 — Monde associatif : Loi 1901 ? | UNICEF fondé quand ? → 1946 | Nobel ? → 1965",
        "J2 — Histoire : Bataille de Solférino → quelle année ? → 1859 | Qui ? → Henry Dunant",
        "J3 — Script : combien d'étapes ? → 7 | Les 5 règles de l'accroche ?",
        "J4 — AAR : que signifient les 3 lettres ? | Combien d'objections couvertes ? → 15",
        "J4 — Objections : 5 catégories — lesquelles ?",
    ], note="Questions flash : 30 secondes par question — lever la main pour répondre")

    slide_section(prs, 3, "Simulations d'Appels — Règles du Jeu", [
        "Simulation NIVEAU 1 : appel standard, prospect coopératif → valider le script complet",
        "Simulation NIVEAU 2 : appel avec 2 objections tirées au sort parmi les 15",
        "Simulation NIVEAU 3 : appel difficile — prospect conflictuel, méfiance IBAN, dernier positionnement",
        "Durée maximale par simulation : 5 minutes chrono",
        "Grille d'évaluation 24 critères appliquée par le formateur sur chaque simulation",
        "Feedback immédiat après chaque simulation : 2 points forts + 1 axe prioritaire",
        "Chaque fundraiser passe au minimum 2 simulations complètes",
    ])

    slide_section(prs, 4, "Grille d'Évaluation Finale — 24 Critères", [
        "ACCROCHE (4 pts) : naturel / règles 1-5 appliquées / identification genre / ton souriant",
        "SCRIPT (6 pts) : malnutrition / RUTF / proposition paliers / chiffres-clés / transition",
        "ÉCOUTE (3 pts) : silence respecté / reformulation / adaptation au profil prospect",
        "OBJECTIONS (5 pts) : méthode AAR / argumentation / relance / calme / conviction",
        "CLOSING (4 pts) : verrouillage / blindage / demande IBAN / prise de congé",
        "TON GÉNÉRAL (2 pts) : chaleur / conviction / naturel tout au long de l'appel",
        "Expert : 21-24 pts | Intermédiaire : 15-20 pts | Débutant : 8-14 pts",
    ])

    # RÈGLE D'OR — Quiz
    slide_section(prs, 5, "RÈGLE D'OR — Instructions Quiz 40 Questions", [
        "Le quiz comporte 40 questions réparties en 6 parties thématiques",
        "PARTIE A — UNICEF & Cause (Q1–Q8)",
        "PARTIE B — Monde Associatif (Q9–Q15)",
        "PARTIE C — Script & Accroche (Q16–Q22)",
        "PARTIE D — KPIs & Nomenclature (Q23–Q28)",
        "PARTIE E — Objections & AAR (Q29–Q35)",
        "PARTIE F — Closing & Validation (Q36–Q40)",
        "Durée : 45 minutes maximum | Réponse unique par question",
    ], is_regle_dor=True)

    # ── Quiz — 40 questions réparties sur 10 slides de 4Q ──
    quiz_data = [
        ("Q1", "En quelle année l'UNICEF a-t-il été créé ?", ["1939","1946","1956","1963"], 1),
        ("Q2", "Dans combien de pays l'UNICEF intervient-il ?", ["150 pays","175 pays","190 pays","210 pays"], 2),
        ("Q3", "Quel est le poids d'un sachet RUTF ?", ["50 grammes","72 grammes","92 grammes","110 grammes"], 2),
        ("Q4", "Quel % des fonds collectés va directement aux programmes terrain ?", ["75%","85%","93%","100%"], 2),
        ("Q5", "En quelle année l'UNICEF a-t-il reçu le Prix Nobel de la Paix ?", ["1955","1965","1975","1985"], 1),
        ("Q6", "Quel % des enfants du monde l'UNICEF vaccine-t-il ?", ["25%","35%","45%","55%"], 2),
        ("Q7", "RUTF signifie Ready-to-Use Therapeutic Food. Quelle est sa composition de base ?", [
            "Riz, eau, sel, vitamines",
            "Pâte d'arachide, matière grasse végétale, lait écrémé, vitamines",
            "Farine de maïs, huile, sucre, minéraux",
            "Lait entier, beurre de cacao, protéines végétales"], 1),
        ("Q8", "De quoi n'a PAS besoin le sachet RUTF pour être consommé ?", [
            "D'être ouvert avant consommation",
            "D'eau, d'électricité ni de cuisson",
            "D'un enfant de moins de 5 ans",
            "D'un stockage à température ambiante"], 1),
        ("Q9", "Combien d'associations actives compte la France ?", ["500 000","1 million","1,5 million","2 millions"], 2),
        ("Q10", "La Loi 1901 a été adoptée le :", ["1er juillet 1901","14 juillet 1901","1er janvier 1901","1er mars 1901"], 0),
        ("Q11", "La bataille de Solférino a eu lieu en :", ["1845","1859","1871","1889"], 1),
        ("Q12", "Qui est à l'origine de la création de la Croix-Rouge ?", ["Florence Nightingale","Jean-Henri Dunant","Louis Pasteur","Albert Schweitzer"], 1),
        ("Q13", "Qu'est-ce que le label 'Don en Confiance' ?", [
            "Un label commercial pour les boutiques caritatives",
            "Une garantie que l'association est sérieuse et contrôlée",
            "Un statut fiscal accordé par l'État",
            "Un label délivré par l'ONU"], 1),
        ("Q14", "Combien de bénévoles mobilise le secteur associatif français ?", ["5 millions","12 millions","22 millions","35 millions"], 2),
        ("Q15", "Quel organisme a été créé en 1989 pour renforcer le contrôle des associations ?", [
            "Le Comité de la Charte du Don en Confiance",
            "L'Autorité des Marchés Financiers",
            "La Direction Générale des Finances Publiques",
            "Le Haut Commissariat aux Associations"], 0),
        ("Q16", "Combien d'étapes comporte le script officiel UNICEF ?", ["5 étapes","6 étapes","7 étapes","8 étapes"], 2),
        ("Q17", "Quelle est la 1ère règle officielle de l'accroche ?", [
            "Se présenter immédiatement avec son prénom",
            "Dire 'Allô ?' et attendre la réponse",
            "Annoncer l'enregistrement de l'appel",
            "Demander si la personne connaît l'UNICEF"], 1),
        ("Q18", "À quelle étape du script propose-t-on les paliers de don ?", [
            "Étape 2 — La Malnutrition",
            "Étape 3 — Les Sachets RUTF",
            "Étape 4 — L'Appel au Soutien",
            "Étape 7 — Validation des coordonnées"], 2),
        ("Q19", "Quel montant correspond à 50 centimes d'euros par jour ?", ["10€/mois","12€/mois","15€/mois","20€/mois"], 2),
        ("Q20", "Quel est le coût réel de 15€/mois après déduction fiscale à 75% ?", ["2,50€","3,75€","5€","7,50€"], 1),
        ("Q21", "En mode PA À CHAUD, que recueille-t-on auprès du donateur ?", [
            "Son numéro de carte bancaire",
            "Son numéro de sécurité sociale",
            "Son IBAN (série de 27 chiffres commençant par FR)",
            "Son numéro de compte client"], 2),
        ("Q22", "Combien de modes de validation du don existent dans le script ?", ["2 modes","3 modes","4 modes","5 modes"], 2),
        ("Q23", "Que signifie CU/H ?", [
            "Chiffre d'Utilité Horaire",
            "Contacts Utiles par Heure",
            "Coût Unitaire par Heure",
            "Collecte Utile Hebdomadaire"], 1),
        ("Q24", "Quel est l'objectif minimum de CU/H fixé par Fidelis ?", ["6 CU/H","7 CU/H","9 CU/H","12 CU/H"], 2),
        ("Q25", "Un appel est qualifié Contact Utile (CU) s'il dure au minimum :", [
            "30 secondes","45 secondes","60 secondes","90 secondes"], 2),
        ("Q26", "Que signifie TX Transfo ?", [
            "Taux de Transfert des fonds",
            "Taux de Transformation des CU en dons",
            "Taxe de Transaction en ligne",
            "Temps de Traitement des formulaires"], 1),
        ("Q27", "PDC signifie :", [
            "Prélèvement Direct Campagne",
            "Projet De Collecte",
            "Plan de Charge mensuel",
            "Paramètre De Contact"], 2),
        ("Q28", "Quel % de la qualité Fidelis repose sur une qualification conforme ?", ["25%","35%","50%","75%"], 2),
        ("Q29", "Que signifie la méthode AAR ?", [
            "Argumenter, Accepter, Relancer",
            "Accuser réception, Argumenter, Relancer",
            "Analyser, Adapter, Répondre",
            "Accrocher, Aborder, Résoudre"], 1),
        ("Q30", "Face à 'C'est une arnaque !', quelle est la réponse recommandée ?", [
            "Raccrocher immédiatement",
            "Confirmer l'identité UNICEF France + proposer unicef.fr pour vérification",
            "Transférer l'appel au superviseur",
            "Proposer d'envoyer un courrier"], 1),
        ("Q31", "Combien d'objections le module J4 couvre-t-il ?", ["8","10","12","15"], 3),
        ("Q32", "En combien de catégories les objections sont-elles regroupées ?", ["3","4","5","6"], 2),
        ("Q33", "Face à 'Je suis sur liste Bloctel', que répondre ?", [
            "S'excuser et raccrocher immédiatement",
            "Les associations à but non lucratif sont autorisées par Bloctel",
            "Demander de se désincrire de Bloctel",
            "Transférer au service juridique"], 1),
        ("Q34", "Pour la méfiance IBAN, quel est l'argument principal ?", [
            "L'IBAN permet d'accéder au compte bancaire",
            "L'IBAN ne permet pas d'accéder au compte — il autorise uniquement un virement SEPA",
            "L'IBAN est facultatif pour valider un don",
            "L'IBAN sera supprimé après le premier prélèvement"], 1),
        ("Q35", "Face à 'Je n'ai pas les moyens', quel montant peut-on proposer en dernier recours ?", [
            "5€/mois","6€/mois","8€/mois","10€/mois"], 1),
        ("Q36", "Le 1er prélèvement a lieu :", [
            "Immédiatement après l'accord",
            "Le 1er du mois suivant",
            "Le 10 du mois suivant",
            "Le dernier jour du mois suivant"], 2),
        ("Q37", "Qu'est-ce que le 'Blindage' dans le script ?", [
            "Une technique d'enregistrement audio",
            "La demande d'autorisation explicite avant de valider le prélèvement",
            "Un outil de vérification anti-fraude",
            "La mise en attente du prospect"], 1),
        ("Q38", "En mode PA en ligne en direct, le fundraiser doit :", [
            "Demander les coordonnées bancaires oralement",
            "Guider le donateur sur le site sans demander ses coordonnées bancaires",
            "Raccrocher et attendre la confirmation par mail",
            "Transférer l'appel au service en ligne"], 1),
        ("Q39", "La qualification 'INDÉCIS' correspond à :", [
            "Un prospect qui a raccroché sans répondre",
            "Un prospect intéressé mais qui n'a pas validé son don",
            "Un prospect qui a refusé catégoriquement",
            "Un prospect absent lors du rappel"], 1),
        ("Q40", "Que doit faire le fundraiser si un prospect dit 'Je suis sur liste rouge' ?", [
            "Proposer quand même le don en continuant le script",
            "S'excuser, remonter l'info, prendre congé sans insister",
            "Demander confirmation écrite",
            "Transférer à un superviseur"], 1),
    ]

    # Grouper par 4 questions par slide — design premium
    for i in range(0, 40, 4):
        group = quiz_data[i:i+4]
        sl = blank_slide(prs)

        # White background
        rect(sl, 0, 0, W, H, WHITE)

        # Left accent bar
        rect(sl, 0, 0, Inches(0.18), H, TEAL_DARK)

        # TEAL header band
        rect(sl, 0, 0, W, Inches(1.0), TEAL)

        part_num = i // 4 + 1
        parts = ["A — UNICEF & Cause", "B — Monde Associatif", "C — Script & Accroche",
                 "D — KPIs & Nomenclature", "E — Objections & AAR", "F — Closing & Validation",
                 "G","H","I","J"]
        part_label = parts[part_num - 1] if part_num <= len(parts) else f"Partie {part_num}"

        txbox(sl, f"QUIZ FIM — Partie {part_label}", Inches(0.3), Inches(0.1),
              Inches(10), Inches(0.8), size=20, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        txbox(sl, f"Q{i+1}–Q{i+4}  |  Nom : ________________________",
              Inches(10.3), Inches(0.1), Inches(2.8), Inches(0.8), size=11, color=WHITE, align=PP_ALIGN.CENTER)

        # Logo in header
        add_logo(sl, x=Inches(11.0), y=Inches(0.0), h=Inches(1.0))

        letters = ["A", "B", "C", "D"]
        LETTER_COLORS = [TEAL_DARK, TEAL, TEAL_MID, RGBColor(0x00, 0x55, 0x55)]
        for idx, (qcode, question, options, _correct) in enumerate(group):
            col = idx % 2
            row = idx // 2
            x0 = Inches(0.3) + col * Inches(6.55)
            y0 = Inches(1.05) + row * Inches(3.1)
            # Question header — TEAL_MID
            rect(sl, x0, y0, Inches(6.4), Inches(0.52), TEAL_MID)
            qnum = i + idx + 1
            txbox(sl, f"Q{qnum} — {question}", x0 + Inches(0.1), y0 + Inches(0.04),
                  Inches(6.2), Inches(0.46), size=12, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
            # Options
            for oi, (letter, opt) in enumerate(zip(letters, options)):
                oy = y0 + Inches(0.58) + oi * Inches(0.61)
                bg_opt = LGRAY if oi % 2 == 0 else WHITE
                rect(sl, x0 + Inches(0.1), oy + Inches(0.02), Inches(6.2), Inches(0.54), bg_opt)
                # Letter badge
                rect(sl, x0 + Inches(0.12), oy + Inches(0.05), Inches(0.42), Inches(0.42),
                     LETTER_COLORS[oi])
                txbox(sl, letter, x0 + Inches(0.12), oy + Inches(0.05), Inches(0.42), Inches(0.42),
                      size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
                txbox(sl, opt, x0 + Inches(0.62), oy + Inches(0.06), Inches(5.7), Inches(0.44),
                      size=11, color=DARK, align=PP_ALIGN.CENTER)

        # Footer band
        rect(sl, 0, H - Inches(0.75), W, Inches(0.75), TEAL_DARK)
        txbox(sl, "SHY-Performance  ×  FIDELIS  ×  UNICEF France — Quiz FIM 40 Questions",
              Inches(0.3), H - Inches(0.72), Inches(12.7), Inches(0.35), size=10, color=WHITE,
              align=PP_ALIGN.CENTER)
        add_interdiction(sl)

    # Grille formateur — design premium
    sl = blank_slide(prs)

    # White background
    rect(sl, 0, 0, W, H, WHITE)

    # Left accent bar
    rect(sl, 0, 0, Inches(0.18), H, TEAL_DARK)

    # Header band
    rect(sl, 0, 0, W, Inches(1.1), TEAL)

    txbox(sl, "GRILLE FORMATEUR — Correction Quiz 40 Questions", Inches(0.3), Inches(0.15),
          Inches(10.5), Inches(0.8), size=24, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

    # Logo in header
    add_logo(sl, x=Inches(11.0), y=Inches(0.08), h=Inches(1.0))

    corr_letters = ["C","C","C","C","B","C","B","B","C","A","B","B","B","C","A","C","B","C","C","B","C","C","B","C","C","B","C","C","B","B","D","C","B","B","B","C","B","B","B","B"]
    y0 = Inches(1.2)
    for row in range(5):
        for col in range(8):
            idx = row * 8 + col
            if idx >= 40: break
            x = Inches(0.3) + col * Inches(1.6)
            y = y0 + row * Inches(1.05)
            # Card shadow
            rect(sl, x + Inches(0.04), y + Inches(0.04), Inches(1.5), Inches(0.88), MGRAY)
            rect(sl, x, y, Inches(1.5), Inches(0.88), WHITE)
            rect(sl, x, y, Inches(1.5), Inches(0.28), LGRAY)
            txbox(sl, f"Q{idx+1}", x + Inches(0.05), y + Inches(0.02), Inches(1.4), Inches(0.25),
                  size=11, bold=True, color=TEAL_DARK, align=PP_ALIGN.CENTER)
            rect(sl, x + Inches(0.45), y + Inches(0.33), Inches(0.6), Inches(0.48), TEAL)
            txbox(sl, corr_letters[idx], x + Inches(0.45), y + Inches(0.33), Inches(0.6), Inches(0.48),
                  size=18, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

    # Footer band
    rect(sl, 0, H - Inches(0.75), W, Inches(0.75), TEAL_DARK)
    txbox(sl, "SHY-Performance  ×  FIDELIS  ×  UNICEF France",
          Inches(0.3), H - Inches(0.72), Inches(12.7), Inches(0.35), size=10, color=WHITE,
          align=PP_ALIGN.CENTER)
    add_interdiction(sl)

    slide_merci(prs)
    slide_felicitations(prs)
    prs.save("/home/user/fatou/FIM_BookJ5_Synthese_Simulations.pptx")
    print("✓ Book J5 — Synthèse & Simulations")

# ════════════════════════════════════════════════════════════════════════════════
# RUN ALL
# ════════════════════════════════════════════════════════════════════════════════
build_file_conducteur()
build_synopsis()
build_j1()
build_j2()
build_j3()
build_j4()
build_j5()
print("\n✅ Tous les books FIM générés avec succès !")

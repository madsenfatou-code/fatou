"""
J2 MATINÉE — Formation Initiale Module J2
Fidelis × SHY Performance
Le Monde Associatif & Humanitaire en France
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

C_BLUE_DARK   = RGBColor(0x00, 0x30, 0x87)
C_BLUE_MID    = RGBColor(0x00, 0x57, 0xB8)
C_BLUE_LIGHT  = RGBColor(0xE8, 0xF0, 0xFB)
C_CYAN        = RGBColor(0x00, 0xAE, 0xEF)
C_ORANGE      = RGBColor(0xF7, 0x94, 0x1D)
C_ORANGE_DARK = RGBColor(0xF7, 0x61, 0x1D)
C_GREEN       = RGBColor(0x27, 0xAE, 0x60)
C_GREEN_LIGHT = RGBColor(0xEA, 0xFA, 0xF1)
C_RED         = RGBColor(0xE7, 0x4C, 0x3C)
C_RED_LIGHT   = RGBColor(0xFD, 0xF0, 0xEF)
C_YELLOW      = RGBColor(0xF1, 0xC4, 0x0F)
C_YELLOW_LIGHT= RGBColor(0xFE, 0xFD, 0xE8)
C_PURPLE      = RGBColor(0x7C, 0x3A, 0xED)
C_TEAL        = RGBColor(0x0D, 0x94, 0x88)
C_TEAL_LIGHT  = RGBColor(0xF0, 0xFD, 0xFA)
C_WHITE       = RGBColor(0xFF, 0xFF, 0xFF)
C_DARK        = RGBColor(0x0A, 0x16, 0x28)
C_GREY        = RGBColor(0xF0, 0xF4, 0xF9)
C_GREY_LINE   = RGBColor(0xDC, 0xE3, 0xF0)
C_TEXT_LIGHT  = RGBColor(0x55, 0x55, 0x55)

W = Inches(13.33)
H = Inches(7.5)

prs = Presentation()
prs.slide_width  = W
prs.slide_height = H
BLANK = prs.slide_layouts[6]

FOOTER_TXT = "J2 FIM — Matinée  |  Fidelis × SHY Performance  |  Le Monde Associatif & Humanitaire"


def add_slide():
    return prs.slides.add_slide(BLANK)


def rect(slide, x, y, w, h, fill=None, line=None, line_w=None):
    shape = slide.shapes.add_shape(1, x, y, w, h)
    shape.line.fill.background()
    if fill:
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill
    else:
        shape.fill.background()
    if line:
        shape.line.color.rgb = line
        if line_w:
            shape.line.width = line_w
    else:
        shape.line.fill.background()
    return shape


def tb(slide, x, y, w, h, text, size=16, bold=False, color=C_WHITE,
       align=PP_ALIGN.LEFT, italic=False, font="Segoe UI", wrap=True):
    box = slide.shapes.add_textbox(x, y, w, h)
    tf = box.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    run.font.name = font
    return box


def tb_multi(slide, x, y, w, h, lines, wrap=True):
    box = slide.shapes.add_textbox(x, y, w, h)
    tf = box.text_frame
    tf.word_wrap = wrap
    first = True
    for ln in lines:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.alignment = ln.get("align", PP_ALIGN.LEFT)
        if ln.get("space_before"):
            p.space_before = Pt(ln["space_before"])
        run = p.add_run()
        run.text = ln.get("text", "")
        run.font.size = Pt(ln.get("size", 14))
        run.font.bold = ln.get("bold", False)
        run.font.italic = ln.get("italic", False)
        run.font.color.rgb = ln.get("color", C_WHITE)
        run.font.name = ln.get("font", "Segoe UI")
    return box


def footer(slide, txt=FOOTER_TXT):
    rect(slide, 0, H - Inches(.4), W, Inches(.4), fill=C_BLUE_DARK)
    tb(slide, Inches(.3), H - Inches(.38), W - Inches(.6), Inches(.34),
       txt, size=9, color=RGBColor(0xAA, 0xBB, 0xDD), align=PP_ALIGN.CENTER)


def section_top(slide, title, subtitle, accent=C_ORANGE):
    rect(slide, 0, 0, W, Inches(1.35), fill=C_BLUE_DARK)
    rect(slide, 0, 0, W, Inches(.06), fill=accent)
    tb(slide, Inches(.4), Inches(.1), Inches(12.5), Inches(.58),
       title, size=22, bold=True, color=C_WHITE)
    tb(slide, Inches(.4), Inches(.72), Inches(12.5), Inches(.44),
       subtitle, size=12, italic=True, color=C_CYAN)


def formula_dark(slide, x, y, w, h, label, lines, note=""):
    rect(slide, x, y, w, h, fill=C_DARK, line=C_BLUE_MID, line_w=Pt(1))
    tb(slide, x + Inches(.2), y + Inches(.08), w - Inches(.4), Inches(.24),
       label, size=8, bold=True, color=C_CYAN, font="Segoe UI")
    fy = y + Inches(.35)
    for fl in lines:
        tb(slide, x + Inches(.2), fy, w - Inches(.4), Inches(.38),
           fl["text"], size=fl.get("size", 12), bold=fl.get("bold", False),
           color=fl.get("color", C_WHITE), font="Courier New")
        fy += Inches(.36)
    if note:
        tb(slide, x + Inches(.2), y + h - Inches(.32), w - Inches(.4), Inches(.28),
           note, size=8, italic=True, color=RGBColor(0x94, 0xA3, 0xB8))


# ════════════════════════════════════════════════════════════════════
# SLIDE 1 — COVER
# ════════════════════════════════════════════════════════════════════
sl = add_slide()
rect(sl, 0, 0, W, H, fill=C_BLUE_DARK)
rect(sl, 0, 0, W, Inches(.08), fill=C_ORANGE)
rect(sl, 0, H - Inches(.08), W, Inches(.08), fill=C_CYAN)
rect(sl, Inches(7.2), 0, Inches(.06), H, fill=C_CYAN)

# Left panel
tb(sl, Inches(.55), Inches(.85), Inches(6.4), Inches(.42),
   "FORMATION INITIALE MODULE  ·  J2 MATINÉE", size=11, bold=True, color=C_ORANGE)

tb_multi(sl, Inches(.55), Inches(1.35), Inches(6.4), Inches(2.2), [
    {"text": "MODULE J2 — MATINÉE", "size": 32, "bold": True, "color": C_WHITE},
    {"text": "Le Monde Associatif & Humanitaire en France", "size": 16, "bold": False, "color": C_CYAN},
])

tb(sl, Inches(.55), Inches(3.62), Inches(6.4), Inches(.36),
   "9h30 – 13h30  |  4 heures  |  Fidelis mandaté par l'UNICEF",
   size=12, italic=True, color=RGBColor(0xBB, 0xCC, 0xEE))

# Mantra banner
rect(sl, Inches(.55), Inches(4.25), Inches(6.4), Inches(.06), fill=C_ORANGE)
rect(sl, Inches(.55), Inches(4.42), Inches(6.4), Inches(.56), fill=C_ORANGE_DARK)
tb(sl, Inches(.65), Inches(4.47), Inches(6.2), Inches(.46),
   "COMPRENDRE POUR CONVAINCRE — CHAQUE CITOYEN EST CONCERNÉ",
   size=11, bold=True, color=C_WHITE, wrap=True)

# Right panel — programme
rect(sl, Inches(7.5), Inches(.62), Inches(5.52), Inches(6.48), fill=RGBColor(0x00, 0x3A, 0x8A))
tb(sl, Inches(7.7), Inches(.76), Inches(5.1), Inches(.36),
   "AU PROGRAMME", size=11, bold=True, color=C_CYAN)

programme = [
    ("01", "Le Monde Associatif"),
    ("02", "Définition Loi 1901"),
    ("03", "Types d'Associations"),
    ("04", "Ressources des Associations"),
    ("05", "Histoire & Humanitaire"),
    ("06", "Marché de la Collecte"),
    ("07", "Tout Citoyen est Concerné"),
    ("08", "Impact Concret — Exemples Vérifiables"),
]
ty = Inches(1.22)
for num, theme in programme:
    rect(sl, Inches(7.68), ty, Inches(.52), Inches(.46), fill=C_BLUE_MID)
    tb(sl, Inches(7.68), ty, Inches(.52), Inches(.46),
       num, size=9, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)
    tb(sl, Inches(8.28), ty + Inches(.06), Inches(4.65), Inches(.34),
       theme, size=11, color=C_WHITE)
    ty += Inches(.54)

footer(sl)


# ════════════════════════════════════════════════════════════════════
# SLIDE 2 — LE MONDE ASSOCIATIF
# ════════════════════════════════════════════════════════════════════
sl = add_slide()
rect(sl, 0, 0, W, H, fill=C_GREY)
section_top(sl, "01 — LE MONDE ASSOCIATIF",
            "La France associative : 1,5 million d'associations actives")

# Info banner
rect(sl, Inches(.35), Inches(1.45), Inches(12.6), Inches(.58), fill=C_CYAN)
tb(sl, Inches(.55), Inches(1.52), Inches(12.2), Inches(.44),
   "En France, 1,5 million d'associations emploient 1,8 million de salariés et mobilisent 22 millions de bénévoles.",
   size=12, bold=True, color=C_BLUE_DARK)

# Grid of 12 association boxes
assocs = [
    ("CARE", C_BLUE_MID),
    ("SOS SAHEL", C_GREEN),
    ("KTO Télévision Catholique", C_PURPLE),
    ("Institut de Cancérologie de Lorraine", C_RED),
    ("ICM", C_TEAL),
    ("UNADeV", C_ORANGE),
    ("Terre Solidaire", C_GREEN),
    ("UNICEF", C_CYAN),
    ("Douleurs sans frontières", C_RED),
    ("Break Poverty Foundation", C_ORANGE_DARK),
    ("Unapei", C_BLUE_MID),
    ("Action contre la Faim", C_ORANGE),
]

cols_count = 4
rows_count = 3
box_w = Inches(3.05)
box_h = Inches(.72)
gap_x = Inches(.12)
gap_y = Inches(.1)
start_x = Inches(.35)
start_y = Inches(2.18)

for i, (name, color) in enumerate(assocs):
    col = i % cols_count
    row = i // cols_count
    bx = start_x + col * (box_w + gap_x)
    by = start_y + row * (box_h + gap_y)
    rect(sl, bx, by, box_w, box_h, fill=C_WHITE, line=color, line_w=Pt(1.5))
    rect(sl, bx, by, Inches(.06), box_h, fill=color)
    tb(sl, bx + Inches(.15), by + Inches(.18), box_w - Inches(.25), Inches(.36),
       name, size=11, bold=True, color=C_BLUE_DARK)

# Bottom orange note
rect(sl, Inches(.35), Inches(5.72), Inches(12.6), Inches(.6), fill=C_ORANGE_DARK)
tb(sl, Inches(.55), Inches(5.79), Inches(12.2), Inches(.46),
   "L'UNICEF — partenaire de Fidelis — fait partie de ce monde associatif que vous représentez lors de chaque appel.",
   size=12, bold=True, color=C_WHITE)

footer(sl)


# ════════════════════════════════════════════════════════════════════
# SLIDE 3 — DÉFINITION LOI 1901
# ════════════════════════════════════════════════════════════════════
sl = add_slide()
rect(sl, 0, 0, W, H, fill=C_GREY)
section_top(sl, "02 — DÉFINITION D'UNE ASSOCIATION",
            "Loi du 1er juillet 1901 — Texte fondateur du droit associatif français",
            accent=C_TEAL)

# Large dark blue quote box
rect(sl, Inches(.35), Inches(1.5), Inches(12.6), Inches(1.5), fill=C_BLUE_DARK)
rect(sl, Inches(.35), Inches(1.5), Inches(12.6), Inches(.055), fill=C_TEAL)
tb(sl, Inches(.55), Inches(1.58), Inches(12.2), Inches(1.32),
   ("L'article 1er de la loi du 1er juillet 1901 définit une association comme : "
    "« la convention par laquelle deux ou plusieurs personnes mettent en commun, "
    "d'une façon permanente, leurs connaissances ou leur activité dans un but autre "
    "que de partager des bénéfices... »"),
   size=13, italic=True, color=C_WHITE, wrap=True)

# Two columns below
col_w = Inches(6.1)
col_h = Inches(3.1)
col_y = Inches(3.15)

# Left column — blue light
rect(sl, Inches(.35), col_y, col_w, col_h, fill=C_BLUE_LIGHT, line=C_BLUE_MID, line_w=Pt(1.5))
rect(sl, Inches(.35), col_y, col_w, Inches(.055), fill=C_BLUE_MID)
tb(sl, Inches(.5), col_y + Inches(.1), col_w - Inches(.25), Inches(.32),
   "QU'EST-CE QU'UNE ASSOCIATION ?", size=12, bold=True, color=C_BLUE_DARK)
tb(sl, Inches(.5), col_y + Inches(.52), col_w - Inches(.25), col_h - Inches(.65),
   ("Une association est un regroupement de deux personnes au minimum réunies autour "
    "d'un projet commun ou partageant des activités, mais sans chercher à réaliser de "
    "bénéfices. Elle peut avoir des buts très divers : sportif, défense des intérêts "
    "des membres, humanitaire, promotion d'idées ou d'œuvres..."),
   size=11, color=C_TEXT_LIGHT, wrap=True)

# Right column — green light
rect(sl, Inches(6.6), col_y, col_w + Inches(.1), col_h, fill=C_GREEN_LIGHT, line=C_GREEN, line_w=Pt(1.5))
rect(sl, Inches(6.6), col_y, col_w + Inches(.1), Inches(.055), fill=C_GREEN)
tb(sl, Inches(6.75), col_y + Inches(.1), col_w - Inches(.15), Inches(.32),
   "LE CARACTÈRE DÉSINTÉRESSÉ", size=12, bold=True, color=C_GREEN)
tb(sl, Inches(6.75), col_y + Inches(.52), col_w - Inches(.15), col_h - Inches(.65),
   ("Le caractère désintéressé interdit la distribution d'un bénéfice aux associés "
    "MAIS n'implique pas que l'activité soit non-commerciale ou déficitaire. "
    "L'objet peut être commercial (ex : commerce équitable) et le bénéfice peut "
    "servir à la développer."),
   size=11, color=C_TEXT_LIGHT, wrap=True)

footer(sl)


# ════════════════════════════════════════════════════════════════════
# SLIDE 4 — TYPES D'ASSOCIATIONS
# ════════════════════════════════════════════════════════════════════
sl = add_slide()
rect(sl, 0, 0, W, H, fill=C_GREY)
section_top(sl, "03 — LES TYPES D'ASSOCIATIONS",
            "Trois statuts — des droits et obligations différents",
            accent=C_PURPLE)

card_data = [
    (C_ORANGE, RGBColor(0xFF, 0xF4, 0xE6),
     "L'ASSOCIATION DE FAIT",
     ("N'a pas fait l'objet d'une déclaration officielle → n'existe pas aux yeux de "
      "l'administration. Tout ce que ses membres entreprennent est réalisé en leur nom. "
      "Aucun accès aux avantages d'une association déclarée (subventions, agrément, dons...).")),
    (C_BLUE_MID, C_BLUE_LIGHT,
     "L'ASSOCIATION DÉCLARÉE",
     ("Association classique par excellence. Officiellement enregistrée par l'administration "
      "française → existe aux yeux de l'État. "
      "Exemple : association culturelle, club de cinéma, club sportif.")),
    (C_GREEN, C_GREEN_LIGHT,
     "L'ASSOCIATION RECONNUE D'UTILITÉ PUBLIQUE (RUP)",
     ("Statut accordé par l'État. Confère des avantages importants : réception de donations "
      "et legs. Toute association loi 1901 peut obtenir l'agrément. "
      "→ L'UNICEF dispose de ce statut en France.")),
]

cw = Inches(4.1)
card_h = Inches(4.4)
card_y = Inches(1.5)
cx = Inches(.35)
for border, bg, title, body in card_data:
    rect(sl, cx, card_y, cw, card_h, fill=bg, line=border, line_w=Pt(2))
    rect(sl, cx, card_y, cw, Inches(.055), fill=border)
    tb(sl, cx + Inches(.15), card_y + Inches(.1), cw - Inches(.3), Inches(.44),
       title, size=13, bold=True, color=border, wrap=True)
    tb(sl, cx + Inches(.15), card_y + Inches(.65), cw - Inches(.3), card_h - Inches(.8),
       body, size=11, color=C_TEXT_LIGHT, wrap=True)
    cx += cw + Inches(.13)

# Bottom teal banner
rect(sl, Inches(.35), Inches(6.05), Inches(12.6), Inches(.56), fill=C_TEAL)
tb(sl, Inches(.55), Inches(6.12), Inches(12.2), Inches(.42),
   "LES RESSOURCES DES ASSOCIATIONS : Cotisations  ·  Subventions  ·  Dons  ·  Legs  ·  Mécénat  ·  Ventes de produits/services",
   size=12, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)

footer(sl)


# ════════════════════════════════════════════════════════════════════
# SLIDE 5 — HISTOIRE DES ASSOCIATIONS
# ════════════════════════════════════════════════════════════════════
sl = add_slide()
rect(sl, 0, 0, W, H, fill=C_GREY)
section_top(sl, "04 — L'HISTOIRE DES ASSOCIATIONS CARITATIVES",
            "Du Moyen Âge à l'humanitaire moderne — une longue histoire de solidarité",
            accent=C_ORANGE)

era_data = [
    (C_BLUE_DARK, C_WHITE,
     "MOYEN ÂGE",
     ("Les associations caritatives et humanitaires trouvent leur origine au Moyen Âge "
      "en Europe. D'émanation religieuse, elles combattent pour les populations indigentes "
      "et les soldats blessés.")),
    (C_TEAL, C_WHITE,
     "1859",
     ("Henry Dunant, homme d'affaires suisse, vient en aide aux soldats blessés à la "
      "bataille de Solférino. Naissance de la Croix-Rouge et du droit humanitaire "
      "international.")),
    (C_GREEN, C_WHITE,
     "1901",
     ("Loi du 1er juillet 1901 : cadre légal des associations en France. "
      "Explosion du monde associatif. Apparition des grandes ONG modernes.")),
    (C_ORANGE, C_WHITE,
     "AUJOURD'HUI",
     ("1,5 million d'associations en France. L'UNICEF présent dans 190 pays. "
      "Les dons réguliers financent des actions continues et vitales.")),
]

ew = Inches(3.05)
eh = Inches(3.0)
ey = Inches(1.5)
ex = Inches(.35)
for bg, tc, era_title, era_body in era_data:
    rect(sl, ex, ey, ew, eh, fill=bg)
    tb(sl, ex + Inches(.15), ey + Inches(.1), ew - Inches(.3), Inches(.42),
       era_title, size=16, bold=True, color=tc)
    rect(sl, ex + Inches(.15), ey + Inches(.6), ew - Inches(.3), Inches(.03), fill=tc)
    tb(sl, ex + Inches(.15), ey + Inches(.72), ew - Inches(.3), eh - Inches(.85),
       era_body, size=10, color=tc, wrap=True)
    ex += ew + Inches(.12)

# Definition box
rect(sl, Inches(.35), Inches(4.65), Inches(12.6), Inches(.92), fill=C_DARK)
tb(sl, Inches(.55), Inches(4.72), Inches(12.2), Inches(.26),
   "DÉFINITION", size=9, bold=True, color=C_CYAN)
tb(sl, Inches(.55), Inches(5.02), Inches(12.2), Inches(.46),
   "L'HUMANITAIRE se résume à l'ASSISTANCE et aux SECOURS dans les situations de conflits armés ou catastrophes naturelles.",
   size=12, bold=True, color=C_WHITE, wrap=True)

footer(sl)


# ════════════════════════════════════════════════════════════════════
# SLIDE 6 — LES PRINCIPES HUMANITAIRES
# ════════════════════════════════════════════════════════════════════
sl = add_slide()
rect(sl, 0, 0, W, H, fill=C_GREY)
section_top(sl, "04 — LES PRINCIPES DE L'HUMANITAIRE",
            "Les valeurs fondamentales qui guident l'action associative",
            accent=C_RED)

left_x = Inches(.35)
right_x = Inches(6.75)
col_w_l = Inches(6.1)
col_w_r = Inches(6.2)
block_y1 = Inches(1.5)
block_h1 = Inches(2.4)
block_y2 = Inches(4.02)
block_h2 = Inches(2.1)

# Block 1 — dark blue
rect(sl, left_x, block_y1, col_w_l, block_h1, fill=C_BLUE_DARK)
tb(sl, left_x + Inches(.2), block_y1 + Inches(.1), col_w_l - Inches(.4), Inches(.36),
   "LES BASES DE L'ASSISTANCE HUMANITAIRE", size=12, bold=True, color=C_CYAN)
bases = ["L'Humanité", "L'Impartialité", "La Neutralité", "L'Indépendance", "L'Assistance aux plus vulnérables"]
by = block_y1 + Inches(.56)
for b in bases:
    tb(sl, left_x + Inches(.2), by, col_w_l - Inches(.4), Inches(.34),
       f"  ▸  {b}", size=11, color=C_WHITE)
    by += Inches(.34)

# Block 2 — blue light
rect(sl, left_x, block_y2, col_w_l, block_h2, fill=C_BLUE_LIGHT, line=C_BLUE_MID, line_w=Pt(1))
tb(sl, left_x + Inches(.2), block_y2 + Inches(.1), col_w_l - Inches(.4), Inches(.34),
   "LES PRINCIPES", size=12, bold=True, color=C_BLUE_DARK)
principes = [
    "Interdiction que l'aide soit distribuée pour des motifs religieux ou politiques",
    "Devoir de transparence envers les bénéficiaires et les donateurs",
    "Contrôle par un commissaire aux comptes",
]
py = block_y2 + Inches(.54)
for p_item in principes:
    tb(sl, left_x + Inches(.2), py, col_w_l - Inches(.4), Inches(.4),
       f"  ▸  {p_item}", size=10, color=C_TEXT_LIGHT, wrap=True)
    py += Inches(.44)

# Right column — 3 cards stacked
acteurs_title_h = Inches(.38)
rect(sl, right_x, block_y1, col_w_r, acteurs_title_h, fill=C_BLUE_MID)
tb(sl, right_x + Inches(.15), block_y1 + Inches(.05), col_w_r - Inches(.3), Inches(.3),
   "LES ACTEURS DE L'HUMANITAIRE", size=12, bold=True, color=C_WHITE)

acteurs = [
    (C_GREEN, C_GREEN_LIGHT, "L'ÉTAT", "Financements publics, cadre légal, diplomatie internationale"),
    (C_ORANGE, RGBColor(0xFF, 0xF4, 0xE6), "LES ONG GÉNÉRALISTES", "UNICEF, Croix-Rouge, Médecins Sans Frontières..."),
    (C_PURPLE, RGBColor(0xF3, 0xF0, 0xFF), "LES ONG SPÉCIALISÉES", "Action contre la Faim, Care, SOS Sahel, Break Poverty..."),
]
ay = block_y1 + acteurs_title_h + Inches(.06)
acart_h = Inches(1.22)
for border, bg, a_title, a_body in acteurs:
    rect(sl, right_x, ay, col_w_r, acart_h, fill=bg, line=border, line_w=Pt(1.5))
    rect(sl, right_x, ay, col_w_r, Inches(.045), fill=border)
    tb(sl, right_x + Inches(.15), ay + Inches(.08), col_w_r - Inches(.3), Inches(.32),
       a_title, size=11, bold=True, color=border)
    tb(sl, right_x + Inches(.15), ay + Inches(.44), col_w_r - Inches(.3), Inches(.66),
       a_body, size=10, color=C_TEXT_LIGHT, wrap=True)
    ay += acart_h + Inches(.06)

# Bottom causes banner
rect(sl, Inches(.35), Inches(6.25), Inches(12.6), Inches(.48), fill=C_DARK)
tb(sl, Inches(.55), Inches(6.31), Inches(12.2), Inches(.36),
   "LES PRINCIPALES CAUSES : Sociale & Enfance  ·  Développement du tiers monde  ·  Santé & Recherche  ·  Cause animale  ·  Environnement",
   size=11, color=C_CYAN, align=PP_ALIGN.CENTER)

footer(sl)


# ════════════════════════════════════════════════════════════════════
# SLIDE 7 — MARCHÉ DE LA COLLECTE (AVANT 2000)
# ════════════════════════════════════════════════════════════════════
sl = add_slide()
rect(sl, 0, 0, W, H, fill=C_GREY)
section_top(sl, "05 — LE MARCHÉ DE LA COLLECTE",
            "Évolution historique — du boom à la crise de confiance",
            accent=C_YELLOW)

era3 = [
    (C_GREEN, C_WHITE,
     "AVANT 1990 — PLEIN BOOM",
     ("Les ONG reconnues pour leurs actions. Collecte facile. "
      "Outil quasi unique : le courrier postal (mailing). "
      "Confiance totale des donateurs.")),
    (C_ORANGE, C_WHITE,
     "1989–1995 — RENFORCEMENT",
     ("Création du Comité de la Charte (1989) et de l'AFF (Association Française des "
      "Fundraisers). Pleine expansion. Mailing dominant. "
      "Certaines causes s'offrent la TV : Le Téléthon.")),
    (C_RED, C_WHITE,
     "1996 — TOURNANT : L'AFFAIRE ARC",
     ("L'affaire de l'ARC (Association pour la Recherche contre le Cancer) éclabousse "
      "tout le secteur. Toutes les associations caritatives sont pointées du doigt. "
      "Crise de confiance majeure.")),
]

ew3 = Inches(4.1)
eh3 = Inches(3.3)
ey3 = Inches(1.5)
ex3 = Inches(.35)
for bg, tc, e_title, e_body in era3:
    rect(sl, ex3, ey3, ew3, eh3, fill=bg)
    tb(sl, ex3 + Inches(.2), ey3 + Inches(.12), ew3 - Inches(.4), Inches(.42),
       e_title, size=14, bold=True, color=tc)
    rect(sl, ex3 + Inches(.2), ey3 + Inches(.65), ew3 - Inches(.4), Inches(.03), fill=tc)
    tb(sl, ex3 + Inches(.2), ey3 + Inches(.78), ew3 - Inches(.4), eh3 - Inches(.95),
       e_body, size=11, color=tc, wrap=True)
    ex3 += ew3 + Inches(.13)

# Dark formula box
formula_dark(sl, Inches(.35), Inches(4.96), Inches(12.6), Inches(1.62),
             "DEPUIS 2000",
             [{"text": ("DEPUIS 2000 : Saturation du marché · Concurrence accrue entre ONG · "
                        "Sollicitations multiples · Contexte socio-économique difficile"),
               "color": RGBColor(0xE8, 0xF4, 0xFF), "size": 12},
              {"text": "→ Le FUNDRAISER professionnel et éthique devient INDISPENSABLE.",
               "color": C_CYAN, "size": 13, "bold": True}])

footer(sl)


# ════════════════════════════════════════════════════════════════════
# SLIDE 8 — TOUT CITOYEN EST CONCERNÉ (FRANCE)
# ════════════════════════════════════════════════════════════════════
sl = add_slide()
rect(sl, 0, 0, W, H, fill=C_GREY)
section_top(sl, "06 — TOUT CITOYEN EST CONCERNÉ",
            "Vous, vos proches, votre communauté — la preuve par les faits en France",
            accent=C_GREEN)

tb(sl, Inches(.35), Inches(1.48), Inches(12.6), Inches(.38),
   "Pensez-y : chacun d'entre nous a bénéficié directement ou indirectement d'une action associative ou humanitaire au cours de sa vie.",
   size=11, italic=True, color=C_TEXT_LIGHT, wrap=True)

examples = [
    (C_BLUE_MID, C_BLUE_LIGHT,
     "SANTE",
     ("Les Restos du Cœur : 170 millions de repas/an. La Croix-Rouge : aide aux "
      "sans-abri. Les Hôpitaux : financés en partie par des fondations. "
      "→ Si vous ou un proche avez été hospitalisé, des associations ont probablement contribué.")),
    (C_ORANGE, RGBColor(0xFF, 0xF4, 0xE6),
     "ENFANCE",
     ("L'UNICEF agit aussi en France pour les enfants défavorisés. Des milliers "
      "d'associations protègent les enfants en danger. "
      "→ Si vous connaissez un enfant en difficulté scolaire ou sociale, une association l'a peut-être aidé.")),
    (C_GREEN, C_GREEN_LIGHT,
     "CATASTROPHES",
     ("Inondations dans le Var, tremblements de terre en Outre-mer. "
      "Croix-Rouge, Secours Populaire. "
      "→ En cas de catastrophe naturelle en France, ce sont les bénévoles associatifs qui arrivent en premiers.")),
    (C_PURPLE, RGBColor(0xF3, 0xF0, 0xFF),
     "HANDICAP",
     ("Unapei : 3 000 établissements pour personnes handicapées. APF France Handicap. "
      "→ Si vous ou un proche êtes en situation de handicap, des associations permettent leur inclusion.")),
    (C_TEAL, C_TEAL_LIGHT,
     "EDUCATION",
     ("Des milliers d'associations de soutien scolaire. ATD Quart Monde lutte contre "
      "l'illettrisme. "
      "→ Des millions d'élèves ont reçu une aide associative pour réussir leur scolarité.")),
    (C_RED, C_RED_LIGHT,
     "PERSONNES AGEES",
     ("Les Petits Frères des Pauvres combattent l'isolement. ADMR : aide à domicile pour "
      "700 000 personnes/an. "
      "→ Si un de vos proches âgés a besoin d'aide, des associations sont là.")),
]

box_w2 = Inches(4.1)
box_h2 = Inches(1.56)
gap2 = Inches(.1)
base_y = Inches(1.95)
for i, (border, bg, ex_title, ex_body) in enumerate(examples):
    col = i % 3
    row = i // 3
    bx2 = Inches(.35) + col * (box_w2 + gap2)
    by2 = base_y + row * (box_h2 + gap2)
    rect(sl, bx2, by2, box_w2, box_h2, fill=bg, line=border, line_w=Pt(1.5))
    rect(sl, bx2, by2, box_w2, Inches(.05), fill=border)
    tb(sl, bx2 + Inches(.15), by2 + Inches(.08), box_w2 - Inches(.3), Inches(.28),
       ex_title, size=11, bold=True, color=border)
    tb(sl, bx2 + Inches(.15), by2 + Inches(.4), box_w2 - Inches(.3), box_h2 - Inches(.5),
       ex_body, size=9, color=C_TEXT_LIGHT, wrap=True)

footer(sl)


# ════════════════════════════════════════════════════════════════════
# SLIDE 9 — IMPACT MONDIAL & UNICEF
# ════════════════════════════════════════════════════════════════════
sl = add_slide()
rect(sl, 0, 0, W, H, fill=C_GREY)
section_top(sl, "07 — IMPACT MONDIAL — L'UNICEF EN ACTION",
            "Des exemples concrets et vérifiables à travers le monde",
            accent=C_BLUE_MID)

# Header banner
rect(sl, Inches(.35), Inches(1.45), Inches(12.6), Inches(.5), fill=C_BLUE_MID)
tb(sl, Inches(.55), Inches(1.52), Inches(12.2), Inches(.36),
   "L'UNICEF : 75 ans d'action dans 190 pays — quelques résultats concrets et vérifiables",
   size=12, bold=True, color=C_WHITE)

impacts = [
    (C_BLUE_MID, C_BLUE_LIGHT,
     "VACCINATION",
     ("L'UNICEF vaccine 45% des enfants du monde. En 2022 : 2,2 milliards de vaccins "
      "distribués. → La polio, quasi éradiquée grâce à l'UNICEF, aurait paralysé des "
      "millions d'enfants.")),
    (C_GREEN, C_GREEN_LIGHT,
     "EAU POTABLE",
     ("6,4 millions de personnes ont accès à l'eau potable grâce à l'UNICEF. Des milliers "
      "de puits et pompes installés en Afrique subsaharienne. "
      "→ Sans eau propre : diarrhées mortelles, choléra.")),
    (C_ORANGE, RGBColor(0xFF, 0xF4, 0xE6),
     "NUTRITION",
     ("Chaque année : 3 millions d'enfants souffrent de malnutrition sévère. L'UNICEF "
      "fournit des sachets de Plumpy'Nut (aliment thérapeutique). "
      "→ Un sachet de 5€ peut sauver un enfant de la mort.")),
    (C_TEAL, C_TEAL_LIGHT,
     "EDUCATION",
     ("L'UNICEF scolarise des millions d'enfants dans les zones de conflit. Écoles mobiles, "
      "kits scolaires. → 244 millions d'enfants ne vont pas à l'école sans aide humanitaire.")),
    (C_PURPLE, RGBColor(0xF3, 0xF0, 0xFF),
     "PROTECTION",
     ("Protection des enfants soldats, lutte contre les mariages forcés. "
      "→ 300 000 enfants soldats dans le monde. L'UNICEF travaille à leur démobilisation.")),
    (C_RED, C_RED_LIGHT,
     "URGENCES",
     ("Haïti, Syrie, Ukraine, Gaza, Yémen. L'UNICEF intervient dans les 72h après une "
      "catastrophe. → Sans intervention rapide : épidémies, famines, mortalité infantile explosive.")),
]

imp_w = Inches(4.1)
imp_h = Inches(1.56)
imp_gap = Inches(.1)
imp_base_y = Inches(2.08)
for i, (border, bg, imp_title, imp_body) in enumerate(impacts):
    col = i % 3
    row = i // 3
    ix = Inches(.35) + col * (imp_w + imp_gap)
    iy = imp_base_y + row * (imp_h + imp_gap)
    rect(sl, ix, iy, imp_w, imp_h, fill=bg, line=border, line_w=Pt(1.5))
    rect(sl, ix, iy, imp_w, Inches(.05), fill=border)
    tb(sl, ix + Inches(.15), iy + Inches(.08), imp_w - Inches(.3), Inches(.28),
       imp_title, size=11, bold=True, color=border)
    tb(sl, ix + Inches(.15), iy + Inches(.4), imp_w - Inches(.3), imp_h - Inches(.5),
       imp_body, size=9, color=C_TEXT_LIGHT, wrap=True)

footer(sl)


# ════════════════════════════════════════════════════════════════════
# SLIDE 10 — POURQUOI LE DON RÉGULIER EST VITAL
# ════════════════════════════════════════════════════════════════════
sl = add_slide()
rect(sl, 0, 0, W, H, fill=C_GREY)
section_top(sl, "08 — POURQUOI LE DON RÉGULIER EST VITAL",
            "La légitimité de notre demande — transformer la compréhension en conviction",
            accent=C_ORANGE)

# Big dark formula box
formula_dark(sl, Inches(.35), Inches(1.5), Inches(12.6), Inches(1.18),
             "COMPRENDRE → S'ENGAGER → CONVAINCRE",
             [{"text": ("COMPRENDRE → S'ENGAGER → CONVAINCRE  |  "
                        "Un donateur qui comprend l'impact de son don donne davantage, "
                        "plus longtemps, et convainc son entourage."),
               "color": C_WHITE, "size": 12}])

# Two columns
col2_w = Inches(6.1)
col2_h = Inches(3.3)
col2_y = Inches(2.82)

# Left column — blue light
rect(sl, Inches(.35), col2_y, col2_w, col2_h, fill=C_BLUE_LIGHT, line=C_BLUE_MID, line_w=Pt(1.5))
rect(sl, Inches(.35), col2_y, col2_w, Inches(.055), fill=C_BLUE_MID)
tb(sl, Inches(.5), col2_y + Inches(.1), col2_w - Inches(.25), Inches(.32),
   "POURQUOI LE PRÉLÈVEMENT RÉGULIER ?", size=12, bold=True, color=C_BLUE_DARK)
left_pts = [
    "Permet de planifier les actions sur le long terme",
    "Finance des programmes pluriannuels (ex : construction d'écoles)",
    "Un don de 10€/mois = 120€/an = 1 enfant vacciné et nourri pendant 4 mois",
    "La régularité est plus efficace que les dons ponctuels",
    "L'IBAN garantit la continuité du soutien",
    "Chaque donateur régulier est un partenaire de l'UNICEF.",
]
lpy = col2_y + Inches(.52)
for pt in left_pts:
    tb(sl, Inches(.5), lpy, col2_w - Inches(.25), Inches(.42),
       f"  ▸  {pt}", size=10, color=C_TEXT_LIGHT, wrap=True)
    lpy += Inches(.43)

# Right column — orange
rect(sl, Inches(6.6), col2_y, col2_w + Inches(.08), col2_h, fill=RGBColor(0xFF, 0xF4, 0xE6), line=C_ORANGE, line_w=Pt(1.5))
rect(sl, Inches(6.6), col2_y, col2_w + Inches(.08), Inches(.055), fill=C_ORANGE)
tb(sl, Inches(6.75), col2_y + Inches(.1), col2_w - Inches(.15), Inches(.32),
   "VOTRE RÔLE DE FUNDRAISER", size=12, bold=True, color=C_ORANGE)
right_pts = [
    "Vous n'êtes pas un vendeur — vous êtes un messager",
    "Vous relayez une cause qui touche chaque citoyen",
    "Vous permettez à l'UNICEF de continuer son action",
    "Un refus aujourd'hui peut être un don demain",
    "Votre conviction et votre culture de la cause font la différence.",
]
rpy = col2_y + Inches(.52)
for pt in right_pts:
    tb(sl, Inches(6.75), rpy, col2_w - Inches(.15), Inches(.42),
       f"  ▸  {pt}", size=10, color=C_TEXT_LIGHT, wrap=True)
    rpy += Inches(.43)

# Bottom mantra
rect(sl, Inches(.35), Inches(6.22), Inches(12.6), Inches(.52), fill=C_ORANGE_DARK)
tb(sl, Inches(.55), Inches(6.28), Inches(12.2), Inches(.4),
   '"Sans vos appels, sans votre conviction — des programmes vitaux s\'arrêtent. Vous êtes un maillon essentiel de la chaîne humanitaire."',
   size=11, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER, wrap=True)

footer(sl)


# ════════════════════════════════════════════════════════════════════
# SLIDE 11 — RÉCAPITULATIF MATINÉE
# ════════════════════════════════════════════════════════════════════
sl = add_slide()
rect(sl, 0, 0, W, H, fill=C_BLUE_DARK)
rect(sl, 0, 0, W, Inches(.08), fill=C_ORANGE)
rect(sl, 0, H - Inches(.08), W, Inches(.08), fill=C_CYAN)
rect(sl, Inches(6.62), 0, Inches(.055), H, fill=C_CYAN)

# Left panel — mantra + key messages
tb(sl, Inches(.5), Inches(.95), Inches(5.8), Inches(.4),
   "MANTRA DE LA MATINÉE", size=11, bold=True, color=C_ORANGE)
rect(sl, Inches(.5), Inches(1.42), Inches(5.8), Inches(.04), fill=C_ORANGE)
tb(sl, Inches(.5), Inches(1.56), Inches(5.8), Inches(.9),
   "COMPRENDRE\nPOUR CONVAINCRE",
   size=26, bold=True, color=C_WHITE, wrap=True)

tb(sl, Inches(.5), Inches(2.62), Inches(5.8), Inches(.38),
   "5 MESSAGES CLÉS À RETENIR", size=11, bold=True, color=C_CYAN)

key_msgs = [
    "Le monde associatif représente 1,5 M d'associations & 22 M de bénévoles",
    "Une association = 2 personnes min. + but non lucratif (loi 1901)",
    "L'UNICEF est RUP — statut le plus solide accordé par l'État",
    "1996 : crise ARC — la confiance doit se reconquérir par l'éthique",
    "Chaque citoyen est concerné — vous êtes un maillon vital",
]
kmy = Inches(3.08)
for km in key_msgs:
    rect(sl, Inches(.5), kmy, Inches(5.8), Inches(.52), fill=RGBColor(0x00, 0x3A, 0x8A))
    rect(sl, Inches(.5), kmy, Inches(.42), Inches(.52), fill=C_CYAN)
    tb(sl, Inches(.5), kmy, Inches(.42), Inches(.52),
       "✓", size=11, bold=True, color=C_BLUE_DARK, align=PP_ALIGN.CENTER)
    tb(sl, Inches(1.02), kmy + Inches(.09), Inches(5.2), Inches(.34),
       km, size=10, color=C_WHITE)
    kmy += Inches(.56)

# Right panel — 10 chiffres
rect(sl, Inches(7.0), Inches(.9), Inches(5.95), Inches(.46), fill=C_ORANGE)
tb(sl, Inches(7.15), Inches(.96), Inches(5.65), Inches(.34),
   "10 CHIFFRES À RETENIR", size=13, bold=True, color=C_WHITE)

chiffres = [
    "1,5 million d'associations en France",
    "22 millions de bénévoles en France",
    "Loi du 1er juillet 1901",
    "2 personnes minimum pour créer une association",
    "1859 : Henry Dunant fonde l'humanitaire moderne",
    "1996 : affaire ARC — crise de confiance",
    "190 pays : présence de l'UNICEF",
    "45% des enfants du monde vaccinés par l'UNICEF",
    "3 millions d'enfants souffrant de malnutrition sévère/an",
    "10€/mois = impact concret et mesurable sur le terrain",
]
cy = Inches(1.45)
for i, chiffre in enumerate(chiffres, 1):
    rect(sl, Inches(7.0), cy, Inches(5.95), Inches(.5), fill=RGBColor(0x00, 0x3A, 0x8A))
    rect(sl, Inches(7.0), cy, Inches(.42), Inches(.5), fill=C_BLUE_MID)
    tb(sl, Inches(7.0), cy, Inches(.42), Inches(.5),
       str(i), size=10, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)
    tb(sl, Inches(7.52), cy + Inches(.09), Inches(5.38), Inches(.32),
       chiffre, size=10, color=C_WHITE)
    cy += Inches(.52)

footer(sl, "J2 FIM — Récapitulatif Matinée  |  Fidelis × SHY Performance  |  Prêt(e) pour l'après-midi")


# ════════════════════════════════════════════════════════════════════
# SAUVEGARDE
# ════════════════════════════════════════════════════════════════════
out = "/home/user/fatou/J2_Matin_Formation_Fidelis.pptx"
prs.save(out)
print(f"Enregistre : {out}")
print(f"Slides      : {len(prs.slides)}")

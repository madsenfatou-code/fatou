"""
SYNTHÈSE J1 FIM v2 — Réveil pédagogique corrigé
Corrections :
- Suppression ERR / traitement objections (non animé en J1)
- Suppression horaires / durée / objectifs de production (formation, pas production)
- PEL : peut être fait par le client OU par le fundraiser
- Ajout définition FUNDRAISER dans nomenclature
- Fidélisation = Fidelis uniquement / Conquête + PA = SHY Performance
- Fidelis mandate SHY Performance — Fidelis coordonne tout
- Ajout paiement par CB (possible mais non recommandé — CB expire, IBAN non)
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
C_PURPLE_LIGHT= RGBColor(0xF3, 0xF0, 0xFF)
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

def footer(slide, txt="SYNTHÈSE J1 FIM  |  Réveil Pédagogique  |  Fidelis × SHY Performance"):
    rect(slide, 0, H - Inches(.4), W, Inches(.4), fill=C_BLUE_DARK)
    tb(slide, Inches(.3), H - Inches(.38), W - Inches(.6), Inches(.34),
       txt, size=9, color=RGBColor(0xAA,0xBB,0xDD), align=PP_ALIGN.CENTER)

def formula_dark(slide, x, y, w, h, label, lines, note=""):
    rect(slide, x, y, w, h, fill=C_DARK, line=C_BLUE_MID, line_w=Pt(1))
    tb(slide, x+Inches(.2), y+Inches(.08), w-Inches(.4), Inches(.24),
       label, size=8, bold=True, color=C_CYAN, font="Segoe UI")
    fy = y + Inches(.35)
    for fl in lines:
        tb(slide, x+Inches(.2), fy, w-Inches(.4), Inches(.38),
           fl["text"], size=fl.get("size", 12), bold=fl.get("bold", False),
           color=fl.get("color", C_WHITE), font="Courier New")
        fy += Inches(.36)
    if note:
        tb(slide, x+Inches(.2), y+h-Inches(.32), w-Inches(.4), Inches(.28),
           note, size=8, italic=True, color=RGBColor(0x94,0xA3,0xB8))

def qr_bloc(slide, x, y, w, h, question, reponse, q_color=C_ORANGE, r_color=C_GREEN):
    half = h / 2 - Inches(.03)
    rect(slide, x, y, w, half, fill=RGBColor(0xFF,0xF4,0xE6), line=q_color, line_w=Pt(1.5))
    tb(slide, x+Inches(.12), y+Inches(.05), w-Inches(.24), Inches(.2),
       "❓  QUESTION", size=8, bold=True, color=q_color)
    tb(slide, x+Inches(.12), y+Inches(.27), w-Inches(.24), half-Inches(.32),
       question, size=11, bold=True, color=C_BLUE_DARK, wrap=True)
    rect(slide, x, y+half+Inches(.06), w, half, fill=C_GREEN_LIGHT, line=r_color, line_w=Pt(1.5))
    tb(slide, x+Inches(.12), y+half+Inches(.1), w-Inches(.24), Inches(.2),
       "✅  RÉPONSE", size=8, bold=True, color=r_color)
    tb(slide, x+Inches(.12), y+half+Inches(.32), w-Inches(.24), half-Inches(.38),
       reponse, size=11, color=C_BLUE_DARK, wrap=True)

def section_top(slide, num, title, subtitle, accent=C_ORANGE):
    rect(slide, 0, 0, W, Inches(1.35), fill=C_BLUE_DARK)
    rect(slide, 0, 0, W, Inches(.06), fill=accent)
    tb(slide, Inches(.4), Inches(.1), Inches(12.5), Inches(.58),
       f"{num}  —  {title}", size=22, bold=True, color=C_WHITE)
    tb(slide, Inches(.4), Inches(.72), Inches(12.5), Inches(.44),
       subtitle, size=12, italic=True, color=C_CYAN)


# ════════════════════════════════════════════════════════════════════
# SLIDE 1 — COUVERTURE
# ════════════════════════════════════════════════════════════════════
sl = add_slide()
rect(sl, 0, 0, W, H, fill=C_BLUE_DARK)
rect(sl, 0, 0, W, Inches(.08), fill=C_ORANGE)
rect(sl, 0, H-Inches(.08), W, Inches(.08), fill=C_CYAN)
rect(sl, Inches(7.2), 0, Inches(.06), H, fill=C_CYAN)

tb(sl, Inches(.55), Inches(.85), Inches(6.4), Inches(.42),
   "RÉVEIL PÉDAGOGIQUE  ·  PRÉ-SESSION", size=11, bold=True, color=C_ORANGE)
tb_multi(sl, Inches(.55), Inches(1.35), Inches(6.4), Inches(2.2), [
    {"text": "SYNTHÈSE", "size": 52, "bold": True, "color": C_WHITE},
    {"text": "MODULE J1 FIM", "size": 28, "bold": False, "color": C_CYAN},
])
tb(sl, Inches(.55), Inches(3.72), Inches(6.4), Inches(.55),
   "Rappel des points clés du J1 avant de\ndémarrer la suite du programme.",
   size=13, italic=True, color=RGBColor(0xBB,0xCC,0xEE), wrap=True)

rect(sl, Inches(.55), Inches(4.45), Inches(6.4), Inches(.05), fill=C_CYAN)
tb(sl, Inches(.55), Inches(4.62), Inches(6.4), Inches(.34),
   "Objectif de cette synthèse :", size=11, bold=True, color=C_CYAN)
for i, obj in enumerate([
    "Revoir les définitions vues en J1",
    "Ancrer les formules et qualifications",
    "Identifier les notions à consolider",
]):
    tb(sl, Inches(.55), Inches(5.0) + Inches(i*0.35), Inches(6.4), Inches(.32),
       f"  ▸  {obj}", size=12, color=C_WHITE)

# Sommaire droite
rect(sl, Inches(7.5), Inches(.62), Inches(5.52), Inches(6.42), fill=RGBColor(0x00,0x3A,0x8A))
tb(sl, Inches(7.7), Inches(.76), Inches(5.1), Inches(.36),
   "AU PROGRAMME", size=11, bold=True, color=C_CYAN)
themes = [
    ("01", "Mission, Mandat & Acteurs"),
    ("02", "La Nomenclature Clé"),
    ("03", "Les Formules KPI"),
    ("04", "Les 3 Qualifications"),
    ("05", "Modes de Paiement & Produit"),
    ("06", "Conquête & Fidélisation"),
    ("07", "La Phrase d'Accroche"),
    ("08", "Quiz Flash — Testez-vous !"),
]
ty = Inches(1.2)
for num, theme in themes:
    rect(sl, Inches(7.68), ty, Inches(.52), Inches(.46), fill=C_BLUE_MID)
    tb(sl, Inches(7.68), ty, Inches(.52), Inches(.46),
       num, size=9, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)
    tb(sl, Inches(8.28), ty+Inches(.06), Inches(4.65), Inches(.34),
       theme, size=11, color=C_WHITE)
    ty += Inches(.54)

footer(sl)


# ════════════════════════════════════════════════════════════════════
# SLIDE 2 — MISSION, MANDAT & ACTEURS
# ════════════════════════════════════════════════════════════════════
sl = add_slide()
rect(sl, 0, 0, W, H, fill=C_GREY)
section_top(sl, "01", "MISSION, MANDAT & ACTEURS",
            "Qui fait quoi — la chaîne de responsabilité", C_ORANGE)

# Bloc FIDELIS (haut)
rect(sl, Inches(.35), Inches(1.5), Inches(12.6), Inches(1.4), fill=C_BLUE_DARK)
rect(sl, Inches(.35), Inches(1.5), Inches(12.6), Inches(.055), fill=C_CYAN)
tb(sl, Inches(.55), Inches(1.6), Inches(4.0), Inches(.38),
   "🏛  FIDELIS — LE DONNEUR D'ORDRE", size=12, bold=True, color=C_CYAN)
tb(sl, Inches(.55), Inches(1.98), Inches(12.1), Inches(.82),
   ("Fidelis est le partenaire central. Il mandate SHY Performance pour la Conquête et les PA. "
    "Il gère lui-même la Fidélisation / Réactivation. "
    "Fidelis est le coordinateur de toutes les opérations — il fait tout."),
   size=11, color=C_WHITE, wrap=True)

# Flèche mandate
rect(sl, Inches(6.42), Inches(3.05), Inches(.3), Inches(.5), fill=C_ORANGE)
tb(sl, Inches(6.2), Inches(3.08), Inches(.75), Inches(.44),
   "▼", size=20, bold=True, color=C_ORANGE, align=PP_ALIGN.CENTER)
tb(sl, Inches(5.5), Inches(3.22), Inches(2.3), Inches(.3),
   "MANDATE", size=11, bold=True, color=C_ORANGE, align=PP_ALIGN.CENTER)

# 3 blocs : SHY / UNICEF / FIDELIS FIDÉLISATION
blocs = [
    (C_BLUE_MID, C_BLUE_LIGHT,
     "🎯  SHY PERFORMANCE",
     "Mandatée par Fidelis",
     ["Réalise la CONQUÊTE (1ère fois)",
      "Collecte les PA (Prélèvements Automatiques)",
      "Représente l'UNICEF lors de chaque appel",
      "Qualifie chaque CU dans l'outil"]),
    (C_GREEN, C_GREEN_LIGHT,
     "🌍  L'UNICEF",
     "Association partenaire représentée",
     ["Présent dans 190 pays",
      "Défend les droits et la survie des enfants",
      "Nutrition · Santé · Eau · Éducation",
      "Les dons financent ses programmes terrain"]),
    (C_TEAL, C_TEAL_LIGHT,
     "🔄  FIDELIS — FIDÉLISATION",
     "Géré directement par Fidelis (pas SHY)",
     ["Rappelle les anciens donateurs inactifs",
      "Les informe de l'actualité de l'association",
      "Réactive leur soutien si accord",
      "⚠ Ne doit pas ressembler à un rappel de cotisation"]),
]
cx = Inches(.35)
for border, bg, title, sub, pts in blocs:
    cw = Inches(4.1)
    rect(sl, cx, Inches(3.58), cw, Inches(3.42), fill=bg, line=border, line_w=Pt(2))
    rect(sl, cx, Inches(3.58), cw, Inches(.055), fill=border)
    tb(sl, cx+Inches(.15), Inches(3.68), cw-Inches(.3), Inches(.38),
       title, size=12, bold=True, color=border)
    tb(sl, cx+Inches(.15), Inches(4.05), cw-Inches(.3), Inches(.28),
       sub, size=9, italic=True, color=C_TEXT_LIGHT)
    py = Inches(4.4)
    for pt in pts:
        tb(sl, cx+Inches(.15), py, cw-Inches(.3), Inches(.36),
           f"✓  {pt}", size=10, color=C_TEXT_LIGHT)
        py += Inches(.38)
    cx += cw + Inches(.13)

rect(sl, Inches(.35), Inches(7.06), Inches(12.6), Inches(.03), fill=C_ORANGE)
tb(sl, Inches(.35), Inches(7.1), Inches(12.6), Inches(.3),
   '"LE REFUS D\'AUJOURD\'HUI PEUT ÊTRE LE DON DE DEMAIN."',
   size=12, bold=True, color=C_ORANGE_DARK, align=PP_ALIGN.CENTER)
footer(sl)


# ════════════════════════════════════════════════════════════════════
# SLIDE 3 — NOMENCLATURE CLÉ (avec FUNDRAISER, sans ERR)
# ════════════════════════════════════════════════════════════════════
sl = add_slide()
rect(sl, 0, 0, W, H, fill=C_GREY)
section_top(sl, "02", "LA NOMENCLATURE CLÉ",
            "Les termes que vous devez savoir définir instantanément", C_CYAN)

termes = [
    ("FUNDRAISER", "Collecteur de fonds",
     "Agent SHY Performance mandaté par Fidelis pour collecter des dons par téléphone au nom de l'UNICEF.",
     C_ORANGE),
    ("CU",         "Contact Utile / Contact Argumenté",
     "Appel dont la durée de communication dépasse STRICTEMENT 60 secondes. En dessous → non comptabilisé.",
     C_BLUE_MID),
    ("CU/H",       "Contacts Utiles par Heure",
     "Nombre de CU réalisés ÷ Heures de production.",
     C_BLUE_MID),
    ("PDC",        "Plan de Charge",
     "Volume de CU à produire sur le mois, fixé par la campagne.",
     C_PURPLE),
    ("PA",         "Prélèvement Automatique",
     "Don régulier prélevé automatiquement. SEUL produit proposé par SHY sur cette campagne.",
     C_TEAL),
    ("PEL",        "Prélèvement En Ligne",
     "PA finalisé en ligne. Peut être effectué par le FUNDRAISER (à chaud, avec l'IBAN) OU par le client lui-même via le lien.",
     C_TEAL),
    ("TX",         "Taux de Transformation",
     "(Nombre de PEL × 100) ÷ Nombre de CU  —  % de dons parmi les CU.",
     C_GREEN),
    ("DON",        "Qualification : Don",
     "CU dont l'issue est un PA confirmé — à chaud (IBAN collecté) ou par promesse (lien envoyé).",
     C_GREEN),
    ("REF ARG.",   "Refus Argumenté",
     "CU conclu par refus, après présentation d'au moins 1 contre-argument par le fundraiser.",
     C_RED),
    ("INDÉCIS",    "Indécis",
     "CU conclu par hésitation du contact. Lien ou rappel possible. Potentiel de transformation futur.",
     C_ORANGE),
]

cols = [termes[:5], termes[5:]]
xs = [Inches(.35), Inches(6.75)]
for col, cx in zip(cols, xs):
    ry = Inches(1.5)
    for sigle, nom, defin, color in col:
        rh = Inches(.98)
        rect(sl, cx, ry, Inches(6.2), rh, fill=C_WHITE, line=C_GREY_LINE, line_w=Pt(.5))
        rect(sl, cx, ry, Inches(.06), rh, fill=color)
        tb(sl, cx+Inches(.18), ry+Inches(.05), Inches(1.15), Inches(.35),
           sigle, size=11, bold=True, color=color)
        tb(sl, cx+Inches(1.35), ry+Inches(.05), Inches(4.68), Inches(.3),
           nom, size=10, bold=True, color=C_BLUE_DARK)
        tb(sl, cx+Inches(1.35), ry+Inches(.38), Inches(4.68), Inches(.52),
           defin, size=9, color=C_TEXT_LIGHT, italic=True, wrap=True)
        ry += rh + Inches(.03)

footer(sl)


# ════════════════════════════════════════════════════════════════════
# SLIDE 4 — FORMULES KPI
# ════════════════════════════════════════════════════════════════════
sl = add_slide()
rect(sl, 0, 0, W, H, fill=C_GREY)
section_top(sl, "03", "LES FORMULES KPI",
            "Savoir les énoncer et les calculer", C_YELLOW)

formula_dark(sl, Inches(.35), Inches(1.5), Inches(12.6), Inches(1.45),
             "KPI 1 — CU/H",
             [{"text": "CU/H  =  CU réalisés  ÷  Heures de production",
               "color": RGBColor(0xE8,0xF4,0xFF), "size": 14, "bold": True},
              {"text": "Exemple Philippe :  68 CU  ÷  7 h  =  9,71 CU/H",
               "color": RGBColor(0x34,0xD3,0x99), "size": 12}])

formula_dark(sl, Inches(.35), Inches(3.08), Inches(12.6), Inches(1.58),
             "KPI 2 — TAUX DE TRANSFORMATION",
             [{"text": "TX (%)  =  ( Nombre de PEL ou PA  ×  100 )  ÷  Nombre de CU",
               "color": RGBColor(0xE8,0xF4,0xFF), "size": 14, "bold": True},
              {"text": "Exemple Michel  :  ( 1 PEL  ×  100 )  ÷  72 CU  =  1,39 %",
               "color": RGBColor(0x34,0xD3,0x99), "size": 12},
              {"text": "   PEL = fait par le fundraiser (IBAN) OU par le client (lien)",
               "color": RGBColor(0x94,0xA3,0xB8), "size": 10}])

formula_dark(sl, Inches(.35), Inches(4.8), Inches(12.6), Inches(1.28),
             "KPI 3 — PDC (Plan de Charge)",
             [{"text": "PDC atteint  =  CU/H moyen équipe  ×  Heures produites totales équipe",
               "color": RGBColor(0xE8,0xF4,0xFF), "size": 14, "bold": True}],
             note="La valeur PDC est communiquée par le superviseur en début de campagne.")

# Grille d'interprétation
rect(sl, Inches(.35), Inches(6.22), Inches(12.6), Inches(.36),
     fill=C_BLUE_LIGHT, line=C_BLUE_MID, line_w=Pt(1))
tb(sl, Inches(.55), Inches(6.28), Inches(12.2), Inches(.26),
   "INTERPRÉTATION CU/H :    ✅  ≥ 9  →  Objectif atteint     "
   "⚠  7–9  →  Zone d'alerte     ✗  < 7  →  Sous le seuil",
   size=11, bold=True, color=C_BLUE_DARK)

rect(sl, Inches(.35), Inches(6.65), Inches(12.6), Inches(.48),
     fill=C_YELLOW_LIGHT, line=C_YELLOW, line_w=Pt(1))
tb(sl, Inches(.55), Inches(6.7), Inches(12.2), Inches(.38),
   "RAPPEL :  Le CU/H et le TX sont des KPI de production. Durant cette formation initiale, ils servent à comprendre le pilotage, pas à mesurer une performance.",
   size=10, italic=True, color=RGBColor(0x78,0x60,0x00))
footer(sl)


# ════════════════════════════════════════════════════════════════════
# SLIDE 5 — LES 3 QUALIFICATIONS
# ════════════════════════════════════════════════════════════════════
sl = add_slide()
rect(sl, 0, 0, W, H, fill=C_GREY)
section_top(sl, "04", "LES 3 QUALIFICATIONS DE CONTACT UTILE",
            "Tout appel > 60 secondes DOIT être qualifié immédiatement", C_GREEN)

rect(sl, Inches(.35), Inches(1.48), Inches(12.6), Inches(.5),
     fill=C_BLUE_LIGHT, line=C_BLUE_MID, line_w=Pt(1.5))
tb(sl, Inches(.55), Inches(1.54), Inches(12.2), Inches(.38),
   "RAPPEL :  Un appel n'est un CU que si sa durée dépasse STRICTEMENT 60 secondes. En dessous → non comptabilisé.",
   size=12, bold=True, color=C_BLUE_DARK)

qw = Inches(4.1)
qdata = [
    (C_GREEN, C_GREEN_LIGHT, "💚", "DON",
     ["PA confirmé à chaud (IBAN collecté par le fundraiser)",
      "OU lien envoyé — client finalise lui-même en ligne",
      "Entre dans le calcul du TX de transformation",
      "Ne qualifier DON que si PA réellement initié"]),
    (C_RED, C_RED_LIGHT, "❌", "REFUS ARGUMENTÉ",
     ["Contact a refusé après un échange argumenté",
      "Le fundraiser a présenté au moins 1 contre-argument",
      "Ne compte pas dans le TX de transformation",
      "Ne jamais qualifier sans avoir vraiment argumenté"]),
    (C_ORANGE, RGBColor(0xFF,0xF4,0xE6), "⏳", "INDÉCIS",
     ["Contact hésite ou souhaite réfléchir",
      "Un lien de don peut lui être envoyé",
      "Un rappel ultérieur peut être planifié",
      "Potentiel de transformation futur"]),
]
cx = Inches(.35)
for border, bg, icon, name, pts in qdata:
    rect(sl, cx, Inches(2.1), qw, Inches(5.0), fill=bg, line=border, line_w=Pt(2))
    rect(sl, cx, Inches(2.1), qw, Inches(.055), fill=border)
    tb(sl, cx, Inches(2.2), qw, Inches(.55), icon, size=28,
       align=PP_ALIGN.CENTER, color=border)
    tb(sl, cx, Inches(2.72), qw, Inches(.4),
       name, size=14, bold=True, color=border, align=PP_ALIGN.CENTER)
    py = Inches(3.22)
    for pt in pts:
        tb(sl, cx+Inches(.18), py, qw-Inches(.36), Inches(.38),
           f"  {pt}", size=10, color=C_TEXT_LIGHT)
        py += Inches(.43)
    cx += qw + Inches(.13)

rect(sl, Inches(.35), Inches(7.1), Inches(12.6), Inches(.3),
     fill=C_YELLOW_LIGHT, line=C_YELLOW, line_w=Pt(1))
tb(sl, Inches(.55), Inches(7.13), Inches(12.2), Inches(.25),
   "RÈGLE ABSOLUE : Tout doute sur la qualification = INDÉCIS. Ne jamais forcer un DON sans PA réellement initié.",
   size=10, bold=True, color=RGBColor(0x78,0x35,0x00))
footer(sl)


# ════════════════════════════════════════════════════════════════════
# SLIDE 6 — MODES DE PAIEMENT & PRODUIT
# ════════════════════════════════════════════════════════════════════
sl = add_slide()
rect(sl, 0, 0, W, H, fill=C_GREY)
section_top(sl, "05", "MODES DE PAIEMENT & LE PRODUIT PA RÉGULIER",
            "Comment le don est collecté — l'essentiel à retenir", C_TEAL)

# Règle d'or
rect(sl, Inches(.35), Inches(1.48), Inches(12.6), Inches(.5), fill=C_BLUE_MID)
tb(sl, Inches(.55), Inches(1.54), Inches(12.2), Inches(.38),
   "RÈGLE D'OR : Seul le Prélèvement Automatique Régulier est proposé sur cette campagne. Aucun don ponctuel.",
   size=12, bold=True, color=C_WHITE)

# 3 modes
modes = [
    (C_GREEN, C_GREEN_LIGHT,
     "✅  IBAN — Prélèvement automatique",
     "MODE RECOMMANDÉ",
     ("Le fundraiser collecte l'IBAN du client pendant l'appel "
      "et valide le prélèvement en ligne lui-même.\n\n"
      "OU le client reçoit un lien et saisit son IBAN seul.\n\n"
      "✓  Don sécurisé et pérenne\n"
      "✓  L'IBAN ne périme jamais\n"
      "✓  Mode prioritaire sur cette campagne")),
    (C_ORANGE, RGBColor(0xFF,0xF4,0xE6),
     "⚠  CB — Carte Bancaire",
     "POSSIBLE MAIS NON RECOMMANDÉ",
     ("Le paiement peut techniquement être réalisé par carte bancaire.\n\n"
      "❌  La CB a une date d'expiration — le prélèvement s'arrête\n"
      "     automatiquement à l'expiration de la carte.\n\n"
      "❌  Risque de perte du donateur sans action de sa part.\n\n"
      "→  Préférer systématiquement l'IBAN.")),
    (C_TEAL, C_TEAL_LIGHT,
     "🔗  Lien de don",
     "PROMESSE — CLIENT FINALISE SEUL",
     ("Le fundraiser envoie un lien sécurisé au client.\n"
      "Le client finalise lui-même en ligne.\n\n"
      "✓  Lien PERMANENT — aucune date d'expiration\n"
      "✓  Redirige vers le site officiel de l'association\n"
      "✓  E-mail de confirmation automatique au donateur\n"
      "→  Qualification : INDÉCIS jusqu'à confirmation")),
]
cx = Inches(.35)
for border, bg, title, badge, desc in modes:
    cw = Inches(4.1)
    rect(sl, cx, Inches(2.1), cw, Inches(5.0), fill=bg, line=border, line_w=Pt(2))
    rect(sl, cx, Inches(2.1), cw, Inches(.055), fill=border)
    tb(sl, cx+Inches(.15), Inches(2.2), cw-Inches(.3), Inches(.38),
       title, size=11, bold=True, color=border)
    rect(sl, cx+Inches(.15), Inches(2.62), cw-Inches(.3), Inches(.3), fill=border)
    tb(sl, cx+Inches(.2), Inches(2.64), cw-Inches(.4), Inches(.26),
       badge, size=9, bold=True, color=C_WHITE)
    tb(sl, cx+Inches(.15), Inches(3.0), cw-Inches(.3), Inches(4.0),
       desc, size=10, color=C_TEXT_LIGHT, wrap=True)
    cx += cw + Inches(.13)

footer(sl)


# ════════════════════════════════════════════════════════════════════
# SLIDE 7 — CONQUÊTE & FIDÉLISATION
# ════════════════════════════════════════════════════════════════════
sl = add_slide()
rect(sl, 0, 0, W, H, fill=C_GREY)
section_top(sl, "06", "TYPES D'OPÉRATIONS : CONQUÊTE & FIDÉLISATION",
            "Deux missions distinctes — deux acteurs différents", C_PURPLE)

# Schéma chaîne
rect(sl, Inches(.35), Inches(1.48), Inches(12.6), Inches(.62), fill=C_BLUE_DARK)
tb(sl, Inches(.55), Inches(1.56), Inches(4.0), Inches(.44),
   "🏛  FIDELIS", size=13, bold=True, color=C_CYAN)
tb(sl, Inches(.55), Inches(1.86), Inches(4.0), Inches(.18),
   "Coordinateur — mandate SHY Performance — gère lui-même la Fidélisation",
   size=9, color=RGBColor(0xBB,0xCC,0xEE))
tb(sl, Inches(5.8), Inches(1.58), Inches(1.6), Inches(.42),
   "→ mandate →", size=12, bold=True, color=C_ORANGE, align=PP_ALIGN.CENTER)
tb(sl, Inches(7.5), Inches(1.56), Inches(3.5), Inches(.44),
   "🎯  SHY PERFORMANCE", size=13, bold=True, color=C_ORANGE)
tb(sl, Inches(7.5), Inches(1.86), Inches(5.3), Inches(.18),
   "Réalise la Conquête et collecte les PA",
   size=9, color=RGBColor(0xBB,0xCC,0xEE))

# Conquête
rect(sl, Inches(.35), Inches(2.26), Inches(6.0), Inches(4.62),
     fill=C_BLUE_LIGHT, line=C_BLUE_MID, line_w=Pt(2))
rect(sl, Inches(.35), Inches(2.26), Inches(6.0), Inches(.055), fill=C_BLUE_MID)
tb(sl, Inches(.5), Inches(2.38), Inches(5.7), Inches(.42),
   "🎯  LA CONQUÊTE  —  SHY Performance", size=14, bold=True, color=C_BLUE_DARK)
tb(sl, Inches(.5), Inches(2.86), Inches(5.7), Inches(.3),
   "Appeler pour la 1ère fois un prospect.", size=11, bold=True, color=C_BLUE_MID)
tb(sl, Inches(.5), Inches(3.22), Inches(5.7), Inches(1.4),
   ("Présenter les missions de l'UNICEF représenté par Fidelis.\n"
    "Convaincre d'adhérer à la cause.\n"
    "Inciter à effectuer un PA en ligne (PEL)."),
   size=11, color=C_TEXT_LIGHT, wrap=True)
dossiers = [
    (C_GREEN,    "Dossier conquête PA",   "Prélèvement automatique en ligne"),
    (C_BLUE_MID, "Dossier conquête DON",  "Don en ligne"),
]
dy = Inches(4.72)
for border, titre, desc in dossiers:
    rect(sl, Inches(.5), dy, Inches(5.55), Inches(.72),
         fill=C_WHITE, line=border, line_w=Pt(1.5))
    tb(sl, Inches(.65), dy+Inches(.08), Inches(5.2), Inches(.26),
       f"▶  {titre}", size=11, bold=True, color=border)
    tb(sl, Inches(.65), dy+Inches(.36), Inches(5.2), Inches(.26),
       desc, size=10, color=C_TEXT_LIGHT)
    dy += Inches(.82)

# Fidélisation
rect(sl, Inches(6.7), Inches(2.26), Inches(6.25), Inches(4.62),
     fill=RGBColor(0xFF,0xF4,0xE6), line=C_ORANGE, line_w=Pt(2))
rect(sl, Inches(6.7), Inches(2.26), Inches(6.25), Inches(.055), fill=C_ORANGE)
tb(sl, Inches(6.88), Inches(2.38), Inches(5.9), Inches(.42),
   "🔄  LA FIDÉLISATION  —  FIDELIS uniquement", size=14, bold=True, color=C_ORANGE)
tb(sl, Inches(6.88), Inches(2.86), Inches(5.9), Inches(.3),
   "Rappeler un donateur existant mais inactif.", size=11, bold=True, color=C_ORANGE)
tb(sl, Inches(6.88), Inches(3.22), Inches(5.9), Inches(1.6),
   ("L'informer de l'importance de son soutien à la cause.\n"
    "Le tenir informé de l'actualité de l'association.\n"
    "L'inciter à soutenir de nouveau.\n"
    "S'il accepte → réactiver son soutien."),
   size=11, color=C_TEXT_LIGHT, wrap=True)
rect(sl, Inches(6.72), Inches(4.92), Inches(6.21), Inches(1.28),
     fill=C_RED_LIGHT, line=C_RED, line_w=Pt(1.5))
tb(sl, Inches(6.88), Inches(5.02), Inches(5.9), Inches(.3),
   "⚠  POINT DE VIGILANCE ABSOLU", size=11, bold=True, color=C_RED)
tb(sl, Inches(6.88), Inches(5.36), Inches(5.9), Inches(.78),
   "Cet appel ne doit JAMAIS être vécu par le donateur\ncomme un rappel de paiement de cotisation.",
   size=11, color=C_RED, bold=True, wrap=True)

footer(sl)


# ════════════════════════════════════════════════════════════════════
# SLIDE 8 — PHRASE D'ACCROCHE
# ════════════════════════════════════════════════════════════════════
sl = add_slide()
rect(sl, 0, 0, W, H, fill=C_GREY)
section_top(sl, "07", "LA PHRASE D'ACCROCHE",
            "5 étapes obligatoires — dans cet ordre exact — sans exception", C_ORANGE)

steps = [
    (C_BLUE_MID, "1", "ALLÔ ?",
     "Prononcer uniquement « Allô ? » puis attendre la réponse.",
     "La voix permet d'identifier le genre.  Masculine → M.   Féminine → Mme"),
    (C_TEAL, "2", "IDENTIFIER",
     "Nommer le contact : PRÉNOM + NOM — jamais l'inverse.",
     '✓  "Manuel Macron"  ·  "Christiano Ronaldo"  ·  "Donald Trump"\n✗  JAMAIS "Macron Manuel" — ce n\'est pas français'),
    (C_GREEN, "3", "BONJOUR",
     "Formule de politesse avec le genre identifié à l'étape 1.",
     '"Bonjour M. [Prénom Nom]"   ou   "Bonjour Mme [Prénom Nom]"'),
    (C_ORANGE, "4", "SE PRÉSENTER",
     "Son prénom + Fidelis mandaté par l'UNICEF.",
     '"Je suis [Prénom], j\'appelle de la part de Fidelis, mandaté par l\'UNICEF."'),
    (C_PURPLE, "5", "CRÉER LE LIEN",
     "Demander si le contact connaît les actions humanitaires de l'UNICEF.",
     "Quelle que soit la réponse → expliquer la mission UNICEF et créer l'engagement émotionnel."),
]
ry = Inches(1.5)
for color, num, titre, desc, note in steps:
    rh = Inches(.96)
    rect(sl, Inches(.35), ry, Inches(12.6), rh, fill=C_WHITE, line=C_GREY_LINE, line_w=Pt(.4))
    rect(sl, Inches(.35), ry, Inches(.48), rh, fill=color)
    tb(sl, Inches(.35), ry, Inches(.48), rh,
       num, size=20, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)
    rect(sl, Inches(.9), ry, Inches(1.3), rh, fill=color)
    tb(sl, Inches(.9), ry, Inches(1.3), rh,
       titre, size=11, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)
    tb(sl, Inches(2.34), ry+Inches(.07), Inches(4.9), Inches(.36),
       desc, size=10, bold=True, color=C_BLUE_DARK, wrap=True)
    rect(sl, Inches(7.38), ry+Inches(.06), Inches(.02), rh-Inches(.12), fill=C_GREY_LINE)
    tb(sl, Inches(7.5), ry+Inches(.06), Inches(5.3), rh-Inches(.12),
       note, size=10, italic=True, color=C_TEXT_LIGHT, wrap=True)
    ry += rh + Inches(.04)

footer(sl)


# ════════════════════════════════════════════════════════════════════
# SLIDE 9 — QUIZ FLASH
# ════════════════════════════════════════════════════════════════════
sl = add_slide()
rect(sl, 0, 0, W, H, fill=C_GREY)
section_top(sl, "08", "QUIZ FLASH — TESTEZ-VOUS !",
            "8 questions clés — répondez mentalement avant de regarder la réponse", C_PURPLE)

qr_data = [
    ("Qu'est-ce qu'un CU ?",
     "Appel dont la durée dépasse STRICTEMENT 60 secondes."),
    ("Quelle est la formule du CU/H ?",
     "CU/H = Nombre de CU réalisés ÷ Heures de production"),
    ("Formule du Taux de Transformation ?",
     "TX (%) = (Nombre de PEL × 100) ÷ Nombre de CU"),
    ("Qui peut réaliser un PEL ?",
     "Le FUNDRAISER (IBAN à chaud) OU le CLIENT lui-même (via le lien)."),
    ("Pourquoi l'IBAN est préféré à la CB ?",
     "L'IBAN ne périme jamais. La CB a une date d'expiration → le don s'arrête."),
    ("Qui gère la Fidélisation ?",
     "FIDELIS uniquement — pas SHY Performance."),
    ("Dans quel ordre nommer un contact ?",
     "PRÉNOM puis NOM. Ex : « Manuel Macron ». Jamais l'inverse."),
    ("Quel est le seul produit proposé par SHY sur cette campagne ?",
     "Le Prélèvement Automatique Régulier (PA). Aucun don ponctuel."),
]

cols = [qr_data[:4], qr_data[4:]]
xs = [Inches(.35), Inches(6.75)]
for col, cx in zip(cols, xs):
    qy = Inches(1.5)
    for question, reponse in col:
        qr_bloc(sl, cx, qy, Inches(6.22), Inches(1.36), question, reponse)
        qy += Inches(1.44)

footer(sl)


# ════════════════════════════════════════════════════════════════════
# SLIDE 10 — RÉCAPITULATIF FINAL
# ════════════════════════════════════════════════════════════════════
sl = add_slide()
rect(sl, 0, 0, W, H, fill=C_BLUE_DARK)
rect(sl, 0, 0, W, Inches(.08), fill=C_ORANGE)
rect(sl, 0, H-Inches(.08), W, Inches(.08), fill=C_CYAN)
rect(sl, Inches(6.62), 0, Inches(.055), H, fill=C_CYAN)

tb(sl, Inches(.5), Inches(1.1), Inches(5.8), Inches(.4),
   "LE MANTRA DE LA CAMPAGNE", size=11, bold=True, color=C_ORANGE)
rect(sl, Inches(.5), Inches(1.56), Inches(5.8), Inches(.04), fill=C_ORANGE)
tb(sl, Inches(.5), Inches(1.7), Inches(5.8), Inches(1.3),
   '"LE REFUS\nD\'AUJOURD\'HUI\nPEUT ÊTRE\nLE DON DE DEMAIN."',
   size=22, bold=True, color=C_WHITE, wrap=True)
tb(sl, Inches(.5), Inches(3.12), Inches(5.8), Inches(1.2),
   "Chaque appel compte.\nChaque argument compte.\nChaque donateur compte.",
   size=13, italic=True, color=RGBColor(0xBB,0xCC,0xEE), wrap=True)

tb(sl, Inches(7.0), Inches(1.1), Inches(5.9), Inches(.4),
   "LES RÉFLEXES FONDAMENTAUX", size=11, bold=True, color=C_CYAN)
reflexes = [
    "CU = appel dépassant STRICTEMENT 60 secondes",
    "CU/H = CU ÷ heures de production",
    "TX = (PEL × 100) ÷ CU",
    "PEL : fait par le fundraiser OU par le client",
    "IBAN recommandé — CB non (expiration)",
    "Seul le PA régulier est proposé par SHY",
    "Conquête + PA → SHY Performance",
    "Fidélisation → Fidelis uniquement",
    "Allô ? → genre → Prénom Nom → Bonjour → Fidelis/UNICEF → UNICEF ?",
    "DON / REFUS ARG. / INDÉCIS : qualifier chaque CU immédiatement",
]
ry = Inches(1.6)
for i, ref in enumerate(reflexes, 1):
    rect(sl, Inches(7.0), ry, Inches(5.9), Inches(.5), fill=RGBColor(0x00,0x3A,0x8A))
    rect(sl, Inches(7.0), ry, Inches(.42), Inches(.5), fill=C_CYAN)
    tb(sl, Inches(7.0), ry, Inches(.42), Inches(.5),
       str(i), size=10, bold=True, color=C_BLUE_DARK, align=PP_ALIGN.CENTER)
    tb(sl, Inches(7.5), ry+Inches(.09), Inches(5.35), Inches(.32),
       ref, size=10, color=C_WHITE)
    ry += Inches(.52)

footer(sl, "SYNTHÈSE J1 FIM v2  |  Fidelis × SHY Performance  |  Prêt(e) pour la suite ✅")


# ════════════════════════════════════════════════════════════════════
# SAUVEGARDE
# ════════════════════════════════════════════════════════════════════
out = "/home/user/fatou/Synthese_J1_FIM_Fidelis.pptx"
prs.save(out)
print(f"✅  Enregistré : {out}")
print(f"   Slides      : {len(prs.slides)}")

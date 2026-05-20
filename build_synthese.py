"""
SYNTHÈSE J1 FIM — Rappel & Réveil pédagogique
Objectif : récapituler les points clés avant de démarrer la suite du programme.
Format : slides épurés, visuels forts, questions-réponses flash, formules.
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

def question_reponse(slide, x, y, w, h, question, reponse, q_color=C_ORANGE, r_color=C_GREEN):
    """Bloc question / réponse façon flash card."""
    rect(slide, x, y, w, h/2 - Inches(.04), fill=RGBColor(0xFF,0xF4,0xE6), line=q_color, line_w=Pt(1.5))
    tb(slide, x+Inches(.12), y+Inches(.05), w-Inches(.24), Inches(.22),
       "❓  QUESTION", size=8, bold=True, color=q_color)
    tb(slide, x+Inches(.12), y+Inches(.28), w-Inches(.24), h/2-Inches(.36),
       question, size=11, bold=True, color=C_BLUE_DARK, wrap=True)

    rect(slide, x, y+h/2+Inches(.04), w, h/2 - Inches(.04), fill=C_GREEN_LIGHT, line=r_color, line_w=Pt(1.5))
    tb(slide, x+Inches(.12), y+h/2+Inches(.08), w-Inches(.24), Inches(.22),
       "✅  RÉPONSE", size=8, bold=True, color=r_color)
    tb(slide, x+Inches(.12), y+h/2+Inches(.32), w-Inches(.24), h/2-Inches(.4),
       reponse, size=11, color=C_BLUE_DARK, wrap=True)


# ════════════════════════════════════════════════════════════════════
# SLIDE 1 — COUVERTURE SYNTHÈSE
# ════════════════════════════════════════════════════════════════════
sl = add_slide()
rect(sl, 0, 0, W, H, fill=C_BLUE_DARK)
# Bandes décoratives
rect(sl, 0, 0, W, Inches(.08), fill=C_ORANGE)
rect(sl, 0, H - Inches(.08), W, Inches(.08), fill=C_CYAN)
rect(sl, Inches(7.2), 0, Inches(.06), H, fill=C_CYAN)

# Gauche — titre
tb(sl, Inches(.55), Inches(.9), Inches(6.4), Inches(.42),
   "RÉVEIL PÉDAGOGIQUE  ·  PRÉ-SESSION", size=11, bold=True, color=C_ORANGE)
tb_multi(sl, Inches(.55), Inches(1.38), Inches(6.4), Inches(2.2), [
    {"text": "SYNTHÈSE", "size": 52, "bold": True, "color": C_WHITE},
    {"text": "MODULE J1 FIM", "size": 28, "bold": False, "color": C_CYAN},
])
tb(sl, Inches(.55), Inches(3.75), Inches(6.4), Inches(.55),
   "Rappel des points clés avant de démarrer\nla suite du programme de formation.",
   size=13, color=RGBColor(0xBB,0xCC,0xEE), italic=True, wrap=True)

rect(sl, Inches(.55), Inches(4.5), Inches(6.4), Inches(.06), fill=C_CYAN)

tb(sl, Inches(.55), Inches(4.7), Inches(6.4), Inches(.35),
   "Objectif de cette synthèse :", size=11, bold=True, color=C_CYAN)
objectifs_synt = [
    "Revoir les définitions et formules clés",
    "Valider les acquis avant la suite",
    "Identifier les notions à consolider",
]
oy = Inches(5.1)
for obj in objectifs_synt:
    tb(sl, Inches(.55), oy, Inches(6.4), Inches(.34), f"  ▸  {obj}",
       size=12, color=C_WHITE)
    oy += Inches(.36)

# Droite — sommaire rapide
rect(sl, Inches(7.5), Inches(.65), Inches(5.5), Inches(6.42),
     fill=RGBColor(0x00,0x3A,0x8A))
tb(sl, Inches(7.7), Inches(.78), Inches(5.1), Inches(.38),
   "AU PROGRAMME", size=11, bold=True, color=C_CYAN)
themes = [
    ("01", "Mission & Mandat UNICEF"),
    ("02", "La Nomenclature Clé"),
    ("03", "Les Formules KPI"),
    ("04", "Les 3 Qualifications"),
    ("05", "Carte Bancaire — L'essentiel"),
    ("06", "Conquête & Fidélisation"),
    ("07", "La Phrase d'Accroche"),
    ("08", "Les Horaires J1"),
    ("09", "Quiz Flash — Testez-vous !"),
]
ty = Inches(1.22)
for num, theme in themes:
    rect(sl, Inches(7.68), ty, Inches(.52), Inches(.44), fill=C_BLUE_MID)
    tb(sl, Inches(7.68), ty, Inches(.52), Inches(.44),
       num, size=9, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)
    tb(sl, Inches(8.28), ty+Inches(.06), Inches(4.6), Inches(.34),
       theme, size=11, color=C_WHITE)
    ty += Inches(.52)

footer(sl)

# ════════════════════════════════════════════════════════════════════
# SLIDE 2 — MISSION & MANDAT
# ════════════════════════════════════════════════════════════════════
sl = add_slide()
rect(sl, 0, 0, W, H, fill=C_GREY)
rect(sl, 0, 0, W, Inches(1.35), fill=C_BLUE_DARK)
rect(sl, 0, 0, W, Inches(.06), fill=C_ORANGE)
tb(sl, Inches(.4), Inches(.12), Inches(12.5), Inches(.55),
   "01  —  MISSION & MANDAT UNICEF", size=22, bold=True, color=C_WHITE)
tb(sl, Inches(.4), Inches(.72), Inches(12.5), Inches(.42),
   "Rappel  |  Pourquoi nous appelons et qui nous représentons",
   size=12, italic=True, color=C_CYAN)

# 3 blocs
blocs = [
    (C_BLUE_MID, C_BLUE_LIGHT, "🌍  QUI ?",
     ("Fidelis, mandaté par l'UNICEF.\n\n"
      "L'UNICEF est présent dans 190 pays pour défendre les droits "
      "et la survie des enfants : nutrition, santé, eau potable, éducation, "
      "protection en zones de conflit.")),
    (C_ORANGE, RGBColor(0xFF,0xF4,0xE6), "🎯  QUOI ?",
     ("Obtenir des dons sous forme de prélèvements automatiques réguliers.\n\n"
      "Permettre à l'association de CONTINUER son action humanitaire "
      "et d'AUGMENTER durablement le taux de dons.")),
    (C_GREEN, C_GREEN_LIGHT, "💡  POURQUOI ?",
     ("Chaque prélèvement régulier = des vies impactées sur le long terme.\n\n"
      "Vous êtes le seul lien humain direct entre la cause et le donateur. "
      "Votre conviction fait la différence.")),
]
cx = Inches(.35)
for border, bg, title, desc in blocs:
    rect(sl, cx, Inches(1.52), Inches(4.1), Inches(5.42),
         fill=bg, line=border, line_w=Pt(2))
    rect(sl, cx, Inches(1.52), Inches(4.1), Inches(.055), fill=border)
    tb(sl, cx+Inches(.18), Inches(1.65), Inches(3.74), Inches(.45),
       title, size=14, bold=True, color=border)
    tb(sl, cx+Inches(.18), Inches(2.15), Inches(3.74), Inches(4.65),
       desc, size=11, color=C_TEXT_LIGHT, wrap=True)
    cx += Inches(4.27)

rect(sl, Inches(.35), Inches(7.05), Inches(12.6), Inches(.02), fill=C_ORANGE)
tb(sl, Inches(.35), Inches(7.1), Inches(12.6), Inches(.32),
   '"LE REFUS D\'AUJOURD\'HUI PEUT ÊTRE LE DON DE DEMAIN."',
   size=12, bold=True, color=C_ORANGE_DARK, align=PP_ALIGN.CENTER)
footer(sl)

# ════════════════════════════════════════════════════════════════════
# SLIDE 3 — NOMENCLATURE CLÉ
# ════════════════════════════════════════════════════════════════════
sl = add_slide()
rect(sl, 0, 0, W, H, fill=C_GREY)
rect(sl, 0, 0, W, Inches(1.35), fill=C_BLUE_DARK)
rect(sl, 0, 0, W, Inches(.06), fill=C_CYAN)
tb(sl, Inches(.4), Inches(.12), Inches(12.5), Inches(.55),
   "02  —  LA NOMENCLATURE CLÉ", size=22, bold=True, color=C_WHITE)
tb(sl, Inches(.4), Inches(.72), Inches(12.5), Inches(.42),
   "Les termes que vous devez savoir définir instantanément",
   size=12, italic=True, color=C_CYAN)

termes = [
    ("CU",       "Contact Utile",
     "Appel dont la durée dépasse strictement 60 secondes",
     C_BLUE_MID),
    ("CU/H",     "Contacts Utiles / Heure",
     "Nombre de CU ÷ Heures de production  →  Objectif : 9",
     C_BLUE_MID),
    ("PDC",      "Plan de Charge",
     "Volume de CU à produire sur le mois (fixé par la campagne)",
     C_PURPLE),
    ("PA",       "Prélèvement Automatique",
     "Mode de don régulier  —  seul produit proposé sur cette campagne",
     C_TEAL),
    ("PEL",      "Prélèvement En Ligne",
     "Synonyme opérationnel de PA  —  entre dans le calcul du TX",
     C_TEAL),
    ("TX",       "Taux de Transformation",
     "(Nombre de PEL × 100) ÷ Nombre de CU  —  Objectif selon association",
     C_GREEN),
    ("DON",      "Qualification : Don",
     "CU dont l'issue est un PA confirmé (à chaud ou promesse)",
     C_GREEN),
    ("REF ARG.", "Refus Argumenté",
     "CU conclu par refus après qu'au moins 1 contre-argument a été présenté",
     C_RED),
    ("INDÉCIS",  "Indécis",
     "CU conclu par hésitation  —  lien ou rappel à prévoir",
     C_ORANGE),
    ("ERR",      "Écouter → Reformuler → Rebondir",
     "Méthode officielle de traitement des objections (book dédié)",
     C_PURPLE),
]

# Affichage en 2 colonnes de 5
cols = [termes[:5], termes[5:]]
xs = [Inches(.35), Inches(6.75)]
for col, cx in zip(cols, xs):
    ry = Inches(1.5)
    for sigle, nom, defin, color in col:
        rect(sl, cx, ry, Inches(6.2), Inches(.98), fill=C_WHITE,
             line=C_GREY_LINE, line_w=Pt(.5))
        rect(sl, cx, ry, Inches(.06), Inches(.98), fill=color)
        tb(sl, cx+Inches(.18), ry+Inches(.06), Inches(1.0), Inches(.34),
           sigle, size=12, bold=True, color=color)
        tb(sl, cx+Inches(1.22), ry+Inches(.06), Inches(4.82), Inches(.3),
           nom, size=11, bold=True, color=C_BLUE_DARK)
        tb(sl, cx+Inches(1.22), ry+Inches(.42), Inches(4.82), Inches(.46),
           defin, size=10, color=C_TEXT_LIGHT, italic=True, wrap=True)
        ry += Inches(1.02)

footer(sl)

# ════════════════════════════════════════════════════════════════════
# SLIDE 4 — FORMULES KPI
# ════════════════════════════════════════════════════════════════════
sl = add_slide()
rect(sl, 0, 0, W, H, fill=C_GREY)
rect(sl, 0, 0, W, Inches(1.35), fill=C_BLUE_DARK)
rect(sl, 0, 0, W, Inches(.06), fill=C_YELLOW)
tb(sl, Inches(.4), Inches(.12), Inches(12.5), Inches(.55),
   "03  —  LES FORMULES KPI", size=22, bold=True, color=C_WHITE)
tb(sl, Inches(.4), Inches(.72), Inches(12.5), Inches(.42),
   "Savoir les énoncer, les calculer et les interpréter",
   size=12, italic=True, color=C_CYAN)

# KPI 1
formula_dark(sl, Inches(.35), Inches(1.5), Inches(12.6), Inches(1.45),
             "KPI 1 — CU/H  (Objectif : 9 minimum)",
             [{"text": "CU/H  =  CU réalisés  ÷  Heures de production",
               "color": RGBColor(0xE8,0xF4,0xFF), "size": 14, "bold": True},
              {"text": "Exemple Philippe :  68  ÷  7h  =  9,71 CU/H  ✅",
               "color": RGBColor(0x34,0xD3,0x99), "size": 12}])

# KPI 2
formula_dark(sl, Inches(.35), Inches(3.08), Inches(12.6), Inches(1.55),
             "KPI 2 — TAUX DE TRANSFORMATION  (Objectif selon association)",
             [{"text": "TX (%)  =  ( Nombre de PEL ou PA  ×  100 )  ÷  Nombre de CU",
               "color": RGBColor(0xE8,0xF4,0xFF), "size": 14, "bold": True},
              {"text": "Exemple Michel  :  ( 1 PEL  ×  100 )  ÷  72 CU  =  1,39 %",
               "color": RGBColor(0x34,0xD3,0x99), "size": 12}])

# KPI 3
formula_dark(sl, Inches(.35), Inches(4.76), Inches(12.6), Inches(1.45),
             "KPI 3 — PDC  (Plan de Charge mensuel — valeur communiquée par superviseur)",
             [{"text": "PDC atteint  =  CU/H moyen équipe  ×  Heures produites totales",
               "color": RGBColor(0xE8,0xF4,0xFF), "size": 14, "bold": True},
              {"text": "Objectif journée  =  9 CU/H  ×  6,67 h production  =  60 CU/jour",
               "color": RGBColor(0x34,0xD3,0x99), "size": 12}])

# Grille interprétation CU/H
rect(sl, Inches(.35), Inches(6.32), Inches(12.6), Inches(.08), fill=C_GREY_LINE)
grades = [
    (C_GREEN,  "✅  CU/H ≥ 9",    "Objectif atteint"),
    (C_YELLOW, "⚠  CU/H 7 – 9",  "Zone d'alerte"),
    (C_RED,    "✗  CU/H < 7",    "Sous le seuil"),
]
gx = Inches(.35)
for color, label, desc in grades:
    rect(sl, gx, Inches(6.45), Inches(4.1), Inches(.58), fill=C_WHITE,
         line=color, line_w=Pt(1.5))
    tb(sl, gx+Inches(.12), Inches(6.5), Inches(2.0), Inches(.26),
       label, size=11, bold=True, color=color)
    tb(sl, gx+Inches(.12), Inches(6.76), Inches(3.8), Inches(.22),
       desc, size=10, color=C_TEXT_LIGHT)
    gx += Inches(4.27)

footer(sl)

# ════════════════════════════════════════════════════════════════════
# SLIDE 5 — LES 3 QUALIFICATIONS
# ════════════════════════════════════════════════════════════════════
sl = add_slide()
rect(sl, 0, 0, W, H, fill=C_GREY)
rect(sl, 0, 0, W, Inches(1.35), fill=C_BLUE_DARK)
rect(sl, 0, 0, W, Inches(.06), fill=C_GREEN)
tb(sl, Inches(.4), Inches(.12), Inches(12.5), Inches(.55),
   "04  —  LES 3 QUALIFICATIONS DE CONTACT UTILE", size=22, bold=True, color=C_WHITE)
tb(sl, Inches(.4), Inches(.72), Inches(12.5), Inches(.42),
   "Tout appel > 60 secondes DOIT être qualifié immédiatement",
   size=12, italic=True, color=C_CYAN)

# Règle du 60s
rect(sl, Inches(.35), Inches(1.48), Inches(12.6), Inches(.52),
     fill=C_BLUE_LIGHT, line=C_BLUE_MID, line_w=Pt(1.5))
tb(sl, Inches(.55), Inches(1.54), Inches(12.2), Inches(.4),
   "RAPPEL :  Un appel n'est un CU que si sa durée dépasse strictement 60 secondes. En dessous → non comptabilisé.",
   size=12, bold=True, color=C_BLUE_DARK)

# 3 grandes cartes
qw = Inches(4.1)
qdata = [
    (C_GREEN, C_GREEN_LIGHT, "💚", "DON",
     ["PA confirmé à chaud (IBAN collecté)", "OU lien envoyé avec engagement oral",
      "→ Entre dans le calcul du TX", "→ Qualifier DON seulement si PA initié"]),
    (C_RED, C_RED_LIGHT, "❌", "REFUS ARGUMENTÉ",
     ["Contact a refusé après échange argumenté", "Agent a présenté ≥ 1 contre-argument",
      "→ Ne compte pas dans le TX", "→ Ne jamais qualifier sans avoir argumenté"]),
    (C_ORANGE, RGBColor(0xFF,0xF4,0xE6), "⏳", "INDÉCIS",
     ["Contact hésite ou veut réfléchir", "Lien de don peut être envoyé",
      "→ Rappel possible à planifier", "→ Potentiel de transformation futur"]),
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
        py += Inches(.42)
    cx += qw + Inches(.13)

rect(sl, Inches(.35), Inches(7.1), Inches(12.6), Inches(.3),
     fill=C_YELLOW_LIGHT, line=C_YELLOW, line_w=Pt(1))
tb(sl, Inches(.55), Inches(7.12), Inches(12.2), Inches(.26),
   "RÈGLE ABSOLUE : Tout doute sur la qualification = INDÉCIS. Ne jamais forcer un DON sans PA réellement initié.",
   size=10, bold=True, color=RGBColor(0x78,0x35,0x00))

footer(sl)

# ════════════════════════════════════════════════════════════════════
# SLIDE 6 — CARTE BANCAIRE + PRODUIT
# ════════════════════════════════════════════════════════════════════
sl = add_slide()
rect(sl, 0, 0, W, H, fill=C_GREY)
rect(sl, 0, 0, W, Inches(1.35), fill=C_BLUE_DARK)
rect(sl, 0, 0, W, Inches(.06), fill=C_TEAL)
tb(sl, Inches(.4), Inches(.12), Inches(12.5), Inches(.55),
   "05  —  CARTE BANCAIRE & LE PRODUIT PA RÉGULIER", size=22, bold=True, color=C_WHITE)
tb(sl, Inches(.4), Inches(.72), Inches(12.5), Inches(.42),
   "L'essentiel à retenir pour collecter avec confiance",
   size=12, italic=True, color=C_CYAN)

# Carte bancaire — composants essentiels
rect(sl, Inches(.35), Inches(1.48), Inches(6.0), Inches(5.5),
     fill=C_BLUE_LIGHT, line=C_BLUE_MID, line_w=Pt(1.5))
tb(sl, Inches(.5), Inches(1.6), Inches(5.7), Inches(.35),
   "CARTE BANCAIRE — Composants clés", size=12, bold=True, color=C_BLUE_DARK)
composants = [
    ("Puce (depuis 1992)",      "Élément de sécurité du système CB"),
    ("Numéro de carte",         "16 chiffres — identifiant unique"),
    ("Nom du titulaire",        "Propriétaire de la carte"),
    ("Date d'expiration",       "MM/AA — validité de la carte"),
    ("Cryptogramme",            "3 derniers chiffres — sécurise le paiement à distance"),
    ("Logo Visa / Mastercard",  "Réseau de paiement"),
    ("Piste magnétique",        "Données de paiement en bande noire"),
    ("Hologramme",              "Élément anti-contrefaçon"),
    ("IBAN",                    "Identifiant bancaire pour le prélèvement automatique"),
]
cy = Inches(2.02)
for comp, desc in composants:
    rect(sl, Inches(.38), cy, Inches(5.94), Inches(.5),
         fill=C_WHITE, line=C_GREY_LINE, line_w=Pt(.3))
    tb(sl, Inches(.52), cy+Inches(.07), Inches(2.1), Inches(.3),
       comp, size=10, bold=True, color=C_BLUE_DARK)
    tb(sl, Inches(2.7), cy+Inches(.08), Inches(3.5), Inches(.3),
       desc, size=10, color=C_TEXT_LIGHT, italic=True)
    cy += Inches(.52)

# Produit droite
rect(sl, Inches(6.6), Inches(1.48), Inches(6.35), Inches(5.5),
     fill=C_WHITE, line=C_GREY_LINE, line_w=Pt(1))
tb(sl, Inches(6.78), Inches(1.6), Inches(6.0), Inches(.35),
   "LE PRODUIT — PA RÉGULIER", size=12, bold=True, color=C_BLUE_DARK)

rect(sl, Inches(6.6), Inches(2.02), Inches(6.35), Inches(.48), fill=C_BLUE_MID)
tb(sl, Inches(6.78), Inches(2.08), Inches(6.1), Inches(.36),
   "SEUL produit proposé sur cette campagne — Aucun don ponctuel",
   size=11, bold=True, color=C_WHITE, wrap=True)

modes = [
    (C_GREEN,  "MODE A — À CHAUD",
     "L'agent recueille l'IBAN pendant l'appel et valide lui-même le prélèvement en ligne.\n→ Don sécurisé immédiatement  ·  Mode prioritaire"),
    (C_ORANGE, "MODE B — PROMESSE (LIEN)",
     "L'agent envoie un lien sécurisé. Le donateur finalise lui-même.\n→ Lien permanent (pas de date d'expiration)\n→ E-mail de confirmation automatique au donateur"),
]
my = Inches(2.62)
for border, title, desc in modes:
    rect(sl, Inches(6.62), my, Inches(6.31), Inches(1.9),
         fill=C_GREEN_LIGHT if border == C_GREEN else RGBColor(0xFF,0xF4,0xE6),
         line=border, line_w=Pt(1.5))
    tb(sl, Inches(6.78), my+Inches(.1), Inches(6.1), Inches(.32),
       title, size=11, bold=True, color=border)
    tb(sl, Inches(6.78), my+Inches(.44), Inches(6.1), Inches(1.35),
       desc, size=10, color=C_TEXT_LIGHT, wrap=True)
    my += Inches(2.0)

footer(sl)

# ════════════════════════════════════════════════════════════════════
# SLIDE 7 — CONQUÊTE & FIDÉLISATION
# ════════════════════════════════════════════════════════════════════
sl = add_slide()
rect(sl, 0, 0, W, H, fill=C_GREY)
rect(sl, 0, 0, W, Inches(1.35), fill=C_BLUE_DARK)
rect(sl, 0, 0, W, Inches(.06), fill=C_PURPLE)
tb(sl, Inches(.4), Inches(.12), Inches(12.5), Inches(.55),
   "06  —  TYPES D'OPÉRATIONS  :  CONQUÊTE & FIDÉLISATION", size=22, bold=True, color=C_WHITE)
tb(sl, Inches(.4), Inches(.72), Inches(12.5), Inches(.42),
   "Deux missions distinctes — adapter son discours en conséquence",
   size=12, italic=True, color=C_CYAN)

# Conquête gauche
rect(sl, Inches(.35), Inches(1.48), Inches(6.0), Inches(5.62),
     fill=C_BLUE_LIGHT, line=C_BLUE_MID, line_w=Pt(2))
rect(sl, Inches(.35), Inches(1.48), Inches(6.0), Inches(.055), fill=C_BLUE_MID)
tb(sl, Inches(.5), Inches(1.6), Inches(5.7), Inches(.46),
   "🎯  LA CONQUÊTE", size=18, bold=True, color=C_BLUE_DARK)

tb(sl, Inches(.5), Inches(2.14), Inches(5.7), Inches(.32),
   "Appeler pour la 1ère fois un prospect.", size=11, bold=True, color=C_BLUE_MID)
tb(sl, Inches(.5), Inches(2.5), Inches(5.7), Inches(1.6),
   ("Présenter les missions de l'association.\n"
    "Convaincre d'adhérer à la cause.\n"
    "Inciter à effectuer un PA en ligne."),
   size=11, color=C_TEXT_LIGHT, wrap=True)

# Dossiers Conquête
dossiers = [
    (C_GREEN,  "Dossier conquête PA",  "Prélèvement automatique en ligne"),
    (C_BLUE_MID,"Dossier conquête DON","Don en ligne"),
]
dy = Inches(4.2)
for border, titre, desc in dossiers:
    rect(sl, Inches(.5), dy, Inches(5.6), Inches(.75),
         fill=C_WHITE, line=border, line_w=Pt(1.5))
    tb(sl, Inches(.65), dy+Inches(.08), Inches(5.2), Inches(.28),
       f"▶  {titre}", size=11, bold=True, color=border)
    tb(sl, Inches(.65), dy+Inches(.36), Inches(5.2), Inches(.28),
       desc, size=10, color=C_TEXT_LIGHT)
    dy += Inches(.84)

# Fidélisation droite
rect(sl, Inches(6.7), Inches(1.48), Inches(6.25), Inches(5.62),
     fill=RGBColor(0xFF,0xF4,0xE6), line=C_ORANGE, line_w=Pt(2))
rect(sl, Inches(6.7), Inches(1.48), Inches(6.25), Inches(.055), fill=C_ORANGE)
tb(sl, Inches(6.88), Inches(1.6), Inches(5.9), Inches(.46),
   "🔄  LA FIDÉLISATION (Réactivation)", size=16, bold=True, color=C_ORANGE)

tb(sl, Inches(6.88), Inches(2.14), Inches(5.9), Inches(.32),
   "Rappeler un donateur existant mais inactif.", size=11, bold=True, color=C_ORANGE)
tb(sl, Inches(6.88), Inches(2.5), Inches(5.9), Inches(2.0),
   ("Lui indiquer l'importance de son soutien.\n"
    "Le tenir informé de l'actualité de l'association.\n"
    "L'inciter à soutenir de nouveau.\n"
    "S'il accepte → réactiver son soutien."),
   size=11, color=C_TEXT_LIGHT, wrap=True)

rect(sl, Inches(6.7), Inches(4.62), Inches(6.25), Inches(1.38),
     fill=C_RED_LIGHT, line=C_RED, line_w=Pt(1.5))
tb(sl, Inches(6.88), Inches(4.72), Inches(5.9), Inches(.3),
   "⚠  POINT DE VIGILANCE", size=11, bold=True, color=C_RED)
tb(sl, Inches(6.88), Inches(5.06), Inches(5.9), Inches(.88),
   "Veiller à ce que cet appel ne soit JAMAIS vécu par le donateur comme un rappel de paiement de cotisation.",
   size=11, color=C_RED, bold=True, wrap=True)

footer(sl)

# ════════════════════════════════════════════════════════════════════
# SLIDE 8 — PHRASE D'ACCROCHE
# ════════════════════════════════════════════════════════════════════
sl = add_slide()
rect(sl, 0, 0, W, H, fill=C_GREY)
rect(sl, 0, 0, W, Inches(1.35), fill=C_BLUE_DARK)
rect(sl, 0, 0, W, Inches(.06), fill=C_ORANGE)
tb(sl, Inches(.4), Inches(.12), Inches(12.5), Inches(.55),
   "07  —  LA PHRASE D'ACCROCHE", size=22, bold=True, color=C_WHITE)
tb(sl, Inches(.4), Inches(.72), Inches(12.5), Inches(.42),
   "5 étapes obligatoires — dans cet ordre exact",
   size=12, italic=True, color=C_CYAN)

steps = [
    (C_BLUE_MID,  "1", "ALLÔ ?",
     "Décrocher et prononcer uniquement « Allô ? ». Attendre la réponse.",
     "La voix identifie le genre : masculine → M.  /  féminine → Mme"),
    (C_TEAL,      "2", "IDENTIFIER",
     "Nommer le contact : PRÉNOM + NOM (jamais l'inverse).",
     "✓ « Manuel Macron »  ·  « Christiano Ronaldo »  ·  « Donald Trump »\n✗ JAMAIS « Macron Manuel » — ce n'est pas français"),
    (C_GREEN,     "3", "BONJOUR",
     "Formule de politesse avec le genre identifié.",
     "« Bonjour M. [Prénom Nom] »  ou  « Bonjour Mme [Prénom Nom] »"),
    (C_ORANGE,    "4", "SE PRÉSENTER",
     "Donner son prénom + présenter Fidelis mandaté par l'UNICEF.",
     "« Je suis [Prénom], j'appelle de la part de Fidelis, mandaté par l'UNICEF. »"),
    (C_PURPLE,    "5", "CRÉER LE LIEN",
     "Demander si le contact connaît les actions humanitaires de l'UNICEF.",
     "Quelle que soit la réponse → expliquer la mission de l'UNICEF et créer l'engagement émotionnel."),
]

ry = Inches(1.5)
for color, num, titre, desc, note in steps:
    rh = Inches(.98)
    rect(sl, Inches(.35), ry, Inches(12.6), rh, fill=C_WHITE,
         line=C_GREY_LINE, line_w=Pt(.5))
    rect(sl, Inches(.35), ry, Inches(.48), rh, fill=color)
    tb(sl, Inches(.35), ry, Inches(.48), rh,
       num, size=20, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)
    rect(sl, Inches(.9), ry, Inches(1.35), rh, fill=color)
    tb(sl, Inches(.9), ry, Inches(1.35), rh,
       titre, size=11, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)
    tb(sl, Inches(2.38), ry+Inches(.08), Inches(4.8), Inches(.36),
       desc, size=10, bold=True, color=C_BLUE_DARK, wrap=True)
    tb(sl, Inches(7.3), ry+Inches(.06), Inches(5.5), rh-Inches(.12),
       note, size=10, italic=True, color=C_TEXT_LIGHT, wrap=True)
    ry += rh + Inches(.04)

footer(sl)

# ════════════════════════════════════════════════════════════════════
# SLIDE 9 — HORAIRES J1
# ════════════════════════════════════════════════════════════════════
sl = add_slide()
rect(sl, 0, 0, W, H, fill=C_GREY)
rect(sl, 0, 0, W, Inches(1.35), fill=C_BLUE_DARK)
rect(sl, 0, 0, W, Inches(.06), fill=C_YELLOW)
tb(sl, Inches(.4), Inches(.12), Inches(12.5), Inches(.55),
   "08  —  ORGANISATION DU TEMPS — HORAIRES J1", size=22, bold=True, color=C_WHITE)
tb(sl, Inches(.4), Inches(.72), Inches(12.5), Inches(.42),
   "9h30 → 17h30  |  Production effective : 6h40  |  Objectif : 60 CU / journée",
   size=12, italic=True, color=C_CYAN)

# Timeline
timeline = [
    (Inches(.4),   "9h30",  "PRISE DE POSTE\nBriefing",   C_GREEN),
    (Inches(2.5),  "10h00", "PRODUCTION\nAppels",          C_BLUE_MID),
    (Inches(5.5),  "12h30", "DÉJEUNER\n1 heure",           C_ORANGE),
    (Inches(7.75), "13h30", "REPRISE\nProduction",          C_BLUE_MID),
    (Inches(10.1), "15h30", "PAUSETTES\n20 minutes",        C_TEAL),
    (Inches(11.5), "17h30", "FIN\nDébriefing",              C_RED),
]
rect(sl, Inches(.35), Inches(1.98), Inches(12.6), Inches(.06), fill=C_BLUE_MID)
for cx, heure, label, color in timeline:
    rect(sl, cx, Inches(1.7), Inches(.06), Inches(.55), fill=color)
    tb(sl, cx-Inches(.32), Inches(1.4), Inches(.74), Inches(.28),
       heure, size=9, bold=True, color=C_BLUE_DARK, align=PP_ALIGN.CENTER)
    rect(sl, cx, Inches(2.12), Inches(1.82), Inches(1.22),
         fill=C_GREEN_LIGHT if color==C_GREEN else (C_BLUE_LIGHT if color==C_BLUE_MID else
         (RGBColor(0xFF,0xF4,0xE6) if color==C_ORANGE else
         (C_TEAL_LIGHT if color==C_TEAL else C_RED_LIGHT))),
         line=color, line_w=Pt(1.5))
    tb(sl, cx+Inches(.1), Inches(2.22), Inches(1.62), Inches(1.02),
       label, size=10, bold=True, color=color, wrap=True)

# Récap 4 cases
recap = [
    (C_BLUE_DARK, C_BLUE_LIGHT, "⏰  Amplitude", "8 heures\n9h30 → 17h30"),
    (C_RED,       C_RED_LIGHT,  "🍽  Déjeuner",  "– 1h00\nnon produite"),
    (C_ORANGE,    RGBColor(0xFF,0xF4,0xE6), "☕  Pausettes", "– 0h20\nnon produites"),
    (C_GREEN,     C_GREEN_LIGHT,"✅  Production", "= 6h40\nbase CU/H"),
]
rcx = Inches(.35)
ry_recap = Inches(3.5)
for border, bg, label, val in recap:
    rect(sl, rcx, ry_recap, Inches(3.1), Inches(1.5), fill=bg, line=border, line_w=Pt(2))
    tb(sl, rcx+Inches(.15), ry_recap+Inches(.12), Inches(2.8), Inches(.34),
       label, size=12, bold=True, color=border)
    tb(sl, rcx+Inches(.15), ry_recap+Inches(.5), Inches(2.8), Inches(.85),
       val, size=20, bold=True, color=border, align=PP_ALIGN.CENTER)
    rcx += Inches(3.27)

# Formule objectif journée
formula_dark(sl, Inches(.35), Inches(5.15), Inches(12.6), Inches(1.42),
             "CALCUL DE L'OBJECTIF CU JOURNALIER",
             [{"text": "Objectif CU / journée  =  9 CU/H  ×  6,67 h  =  60 CU",
               "color": RGBColor(0x34,0xD3,0x99), "size": 14, "bold": True},
              {"text": "Soit environ 10 CU par heure de production effective",
               "color": RGBColor(0x7D,0xD3,0xFC), "size": 11}],
             note="⚠  6h40 = 6,67 h décimales. Cet objectif s'adapte si les horaires varient.")

footer(sl)

# ════════════════════════════════════════════════════════════════════
# SLIDE 10 — QUIZ FLASH (Q&R)
# ════════════════════════════════════════════════════════════════════
sl = add_slide()
rect(sl, 0, 0, W, H, fill=C_GREY)
rect(sl, 0, 0, W, Inches(1.35), fill=C_BLUE_DARK)
rect(sl, 0, 0, W, Inches(.06), fill=C_PURPLE)
tb(sl, Inches(.4), Inches(.12), Inches(12.5), Inches(.55),
   "09  —  QUIZ FLASH — TESTEZ-VOUS !", size=22, bold=True, color=C_WHITE)
tb(sl, Inches(.4), Inches(.72), Inches(12.5), Inches(.42),
   "6 questions clés — répondez avant de regarder la réponse",
   size=12, italic=True, color=C_CYAN)

qr_data = [
    ("Qu'est-ce qu'un CU ?",
     "Un appel dont la durée dépasse STRICTEMENT 60 secondes."),
    ("Quelle est la formule du CU/H ?",
     "CU/H = Nombre de CU réalisés ÷ Heures de production"),
    ("Quelle est la formule du TX de transformation ?",
     "TX (%) = (Nombre de PEL ou PA × 100) ÷ Nombre de CU"),
    ("Quel est l'objectif CU/H de la campagne ?",
     "9 CU/H minimum"),
    ("Quel est le seul produit proposé sur cette campagne ?",
     "Le prélèvement automatique régulier (PA). Aucun don ponctuel."),
    ("Dans quel ordre doit-on nommer un contact ?",
     "Toujours PRÉNOM puis NOM. Exemple : « Manuel Macron ». Jamais l'inverse."),
]

cols = [qr_data[:3], qr_data[3:]]
xs = [Inches(.35), Inches(6.75)]
for col, cx in zip(cols, xs):
    qy = Inches(1.48)
    for question, reponse in col:
        qh = Inches(1.85)
        question_reponse(sl, cx, qy, Inches(6.22), qh, question, reponse)
        qy += qh + Inches(.08)

footer(sl)

# ════════════════════════════════════════════════════════════════════
# SLIDE 11 — RÉCAPITULATIF FINAL
# ════════════════════════════════════════════════════════════════════
sl = add_slide()
rect(sl, 0, 0, W, H, fill=C_BLUE_DARK)
rect(sl, 0, 0, W, Inches(.08), fill=C_ORANGE)
rect(sl, 0, H-Inches(.08), W, Inches(.08), fill=C_CYAN)
rect(sl, Inches(6.62), 0, Inches(.055), H, fill=C_CYAN)

# Gauche — mantra
tb(sl, Inches(.5), Inches(1.2), Inches(5.8), Inches(.42),
   "LE MANTRA DE LA CAMPAGNE", size=11, bold=True, color=C_ORANGE)
rect(sl, Inches(.5), Inches(1.68), Inches(5.8), Inches(.04), fill=C_ORANGE)
tb(sl, Inches(.5), Inches(1.82), Inches(5.8), Inches(1.3),
   '"LE REFUS\nD\'AUJOURD\'HUI\nPEUT ÊTRE\nLE DON DE DEMAIN."',
   size=22, bold=True, color=C_WHITE, wrap=True)
tb(sl, Inches(.5), Inches(3.25), Inches(5.8), Inches(1.2),
   "Chaque appel compte.\nChaque argument compte.\nChaque donateur compte.",
   size=13, italic=True, color=RGBColor(0xBB,0xCC,0xEE), wrap=True)

# Droite — les 10 réflexes à ancrer
tb(sl, Inches(7.0), Inches(1.2), Inches(5.9), Inches(.42),
   "LES 10 RÉFLEXES À ANCRER", size=11, bold=True, color=C_CYAN)
reflexes = [
    "Un CU = appel > 60 secondes",
    "CU/H = CU ÷ heures  →  objectif 9",
    "TX = (PEL × 100) ÷ CU",
    "Seul le PA régulier est proposé",
    "Allô ? → identifier genre → Prénom Nom",
    "Présenter Fidelis mandaté par l'UNICEF",
    "Qualifier chaque CU : DON / REFUS ARG. / INDÉCIS",
    "Conquête = 1ère fois  |  Fidélisation = réactivation",
    "Lien de don = permanent, pas de date d'expiration",
    "Production effective J1 : 6h40 → objectif 60 CU",
]
ry = Inches(1.72)
for i, ref in enumerate(reflexes, 1):
    rect(sl, Inches(7.0), ry, Inches(5.9), Inches(.5), fill=RGBColor(0x00,0x3A,0x8A))
    rect(sl, Inches(7.0), ry, Inches(.42), Inches(.5), fill=C_CYAN)
    tb(sl, Inches(7.0), ry, Inches(.42), Inches(.5),
       str(i), size=10, bold=True, color=C_BLUE_DARK, align=PP_ALIGN.CENTER)
    tb(sl, Inches(7.5), ry+Inches(.08), Inches(5.35), Inches(.34),
       ref, size=11, color=C_WHITE)
    ry += Inches(.54)

footer(sl, "SYNTHÈSE J1 FIM  |  Fidelis × SHY Performance  |  Prêt(e) pour la suite ✅")


# ════════════════════════════════════════════════════════════════════
# SAUVEGARDE
# ════════════════════════════════════════════════════════════════════
out = "/home/user/fatou/Synthese_J1_FIM_Fidelis.pptx"
prs.save(out)
print(f"✅  Fichier enregistré : {out}")
print(f"   Nombre de slides   : {len(prs.slides)}")

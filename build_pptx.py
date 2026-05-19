"""
Book J1 FIM Fidelis × SHY Performance — VERSION 2
Contenu strictement limité aux informations transmises dans les 2 commandes.
Slides supprimés : structure appel 7 étapes, objections, posture, auto-suivi, checklist.
Ajout : phrase d'accroche avec règles précises + correction DEL = Don en Ligne.
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

# ── Palette couleurs ─────────────────────────────────────────────────
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

# ════════════════════════════════════════════════════════════════════
# HELPERS
# ════════════════════════════════════════════════════════════════════

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

def tb(slide, x, y, w, h, text, size=18, bold=False, color=C_WHITE,
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

def header(slide, num, title, subtitle=""):
    rect(slide, 0, 0, W, Inches(1.52), fill=C_BLUE_DARK)
    rect(slide, Inches(.35), Inches(.26), Inches(.9), Inches(.44), fill=C_CYAN)
    tb(slide, Inches(.35), Inches(.26), Inches(.9), Inches(.44),
       num, size=13, bold=True, color=C_BLUE_DARK, align=PP_ALIGN.CENTER)
    tb(slide, Inches(1.38), Inches(.18), Inches(11.5), Inches(.62),
       title, size=24, bold=True, color=C_WHITE)
    if subtitle:
        tb(slide, Inches(1.38), Inches(.8), Inches(11.5), Inches(.5),
           subtitle, size=12, italic=True, color=C_CYAN)

def footer(slide, txt="J1 FIM — Fidelis × SHY Performance  |  Formation Initiale Module"):
    rect(slide, 0, H - Inches(.44), W, Inches(.44), fill=C_BLUE_DARK)
    tb(slide, Inches(.3), H - Inches(.41), W - Inches(.6), Inches(.38),
       txt, size=9, color=RGBColor(0xAA,0xBB,0xDD), align=PP_ALIGN.CENTER)

def formula_box(slide, x, y, w, h, label, lines, note=""):
    rect(slide, x, y, w, h, fill=C_DARK, line=C_BLUE_MID, line_w=Pt(1))
    tb(slide, x+Inches(.2), y+Inches(.1), w-Inches(.4), Inches(.26),
       label, size=8, bold=True, color=C_CYAN, font="Segoe UI")
    fy = y + Inches(.4)
    for fl in lines:
        tb(slide, x+Inches(.2), fy, w-Inches(.4), Inches(.38),
           fl["text"], size=fl.get("size",13), bold=fl.get("bold",False),
           color=fl.get("color", C_WHITE), font="Courier New")
        fy += Inches(.36)
    if note:
        tb(slide, x+Inches(.2), y+h-Inches(.35), w-Inches(.4), Inches(.3),
           note, size=8, italic=True, color=RGBColor(0x94,0xA3,0xB8))

# ════════════════════════════════════════════════════════════════════
# SLIDE 1 — COUVERTURE
# ════════════════════════════════════════════════════════════════════
sl = add_slide()
rect(sl, 0, 0, W*0.62, H, fill=C_BLUE_DARK)
rect(sl, W*0.62, 0, W*0.38, H, fill=C_BLUE_MID)
rect(sl, W*0.605, 0, Inches(.055), H, fill=C_CYAN)
rect(sl, 0, H*0.72, W*0.62, Inches(.048), fill=C_ORANGE)

rect(sl, Inches(.5), Inches(.52), Inches(2.85), Inches(.42), fill=C_ORANGE)
tb(sl, Inches(.5), Inches(.52), Inches(2.85), Inches(.42),
   "FORMATION INITIALE MODULE  ·  J1",
   size=10, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)

tb_multi(sl, Inches(.5), Inches(1.15), Inches(7.5), Inches(2.5), [
    {"text":"MODULE J1 FIM", "size":42, "bold":True, "color":C_WHITE},
    {"text":"Collecte de Dons Humanitaires", "size":28, "color":C_CYAN},
])
tb(sl, Inches(.5), Inches(3.75), Inches(7.4), Inches(.8),
   "Formation des agents SHY Performance mandatés par Fidelis\npour la collecte de dons en faveur de l'UNICEF.",
   size=13, color=RGBColor(0xBB,0xCC,0xEE), wrap=True)

infos = [
    ("🎯","Objectif CU/H","9 minimum"),
    ("⏱","Seuil CU","> 60 secondes de communication"),
    ("💳","Produit","Prélèvement automatique régulier"),
    ("🌍","Mandat","Fidelis mandaté par l'UNICEF"),
    ("📊","KPI","CU/H  |  Taux de Transformation  |  PDC"),
]
iy = Inches(1.0)
for icon, lab, val in infos:
    rect(sl, W*0.645, iy, Inches(4.6), Inches(.82), fill=RGBColor(0x00,0x40,0xA0))
    tb(sl, W*0.66, iy+Inches(.06), Inches(4.3), Inches(.3),
       f"{icon}  {lab}", size=10, bold=True, color=C_CYAN)
    tb(sl, W*0.66, iy+Inches(.36), Inches(4.3), Inches(.32),
       f"     {val}", size=12, color=C_WHITE)
    iy += Inches(.88)

rect(sl, 0, H-Inches(1.15), W*0.62, Inches(.72), fill=C_ORANGE_DARK)
tb(sl, Inches(.5), H-Inches(1.13), W*0.58, Inches(.68),
   '"LE REFUS D\'AUJOURD\'HUI PEUT ÊTRE LE DON DE DEMAIN."',
   size=13, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)
footer(sl)

# ════════════════════════════════════════════════════════════════════
# SLIDE 2 — SOMMAIRE
# ════════════════════════════════════════════════════════════════════
sl = add_slide()
rect(sl, 0, 0, W, H, fill=C_GREY)
header(sl, "PLAN", "Sommaire du Module J1",
       "Contenu issu strictement des informations de campagne transmises")

sections = [
    ("01","Mission & Finalité",         "Comprendre l'enjeu humanitaire de la campagne"),
    ("02","Nomenclature",               "Définitions officielles — CU, PDC, TX, DEL, PA…"),
    ("03","KPI — CU/H",                 "Formule, objectif 9 CU/H, exemple agent Philippe"),
    ("04","KPI — Taux de Transformation","Formule, objectif association, exemple agent Michel"),
    ("05","KPI — Plan de Charge (PDC)", "Volume mensuel, formule de vérification"),
    ("06","Les 3 Qualifications",       "DON / REFUS ARGUMENTÉ / INDÉCIS"),
    ("07","Le Produit — Prélèvement Auto","PA régulier uniquement, 2 modes, lien de don"),
    ("08","La Phrase d'Accroche",       "Script d'ouverture d'appel — règles officielles"),
]

cols = [sections[:4], sections[4:]]
xs = [Inches(.4), Inches(6.85)]
for col, cx in zip(cols, xs):
    y = Inches(1.72)
    for num, title, sub in col:
        rect(sl, cx, y, Inches(6.2), Inches(.9), fill=C_WHITE,
             line=C_BLUE_LIGHT, line_w=Pt(1))
        rect(sl, cx, y, Inches(.55), Inches(.9), fill=C_BLUE_MID)
        tb(sl, cx, y, Inches(.55), Inches(.9),
           num, size=10, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)
        tb(sl, cx+Inches(.62), y+Inches(.06), Inches(5.5), Inches(.34),
           title, size=13, bold=True, color=C_BLUE_DARK)
        tb(sl, cx+Inches(.62), y+Inches(.44), Inches(5.5), Inches(.3),
           sub, size=10, italic=True, color=C_TEXT_LIGHT)
        y += Inches(.98)

rect(sl, Inches(.4), H-Inches(1.05), Inches(12.5), Inches(.48),
     fill=C_YELLOW_LIGHT, line=C_YELLOW, line_w=Pt(1.5))
tb(sl, Inches(.6), H-Inches(1.02), Inches(12.1), Inches(.42),
   "NOTE : Le traitement des objections, le script complet et les techniques avancées feront l'objet d'un book dédié séparé.",
   size=11, bold=True, color=RGBColor(0x78,0x35,0x00))
footer(sl)

# ════════════════════════════════════════════════════════════════════
# SLIDE 3 — MISSION & FINALITÉ
# ════════════════════════════════════════════════════════════════════
sl = add_slide()
rect(sl, 0, 0, W, H, fill=C_GREY)
header(sl, "01", "Mission & Finalité de la Campagne",
       "Pourquoi nous appelons — ce que chaque don finance concrètement")

rect(sl, Inches(.35), Inches(1.72), Inches(12.6), Inches(.88), fill=C_BLUE_MID)
tb(sl, Inches(.55), Inches(1.78), Inches(12.2), Inches(.38),
   "FINALITÉ OFFICIELLE", size=9, bold=True, color=C_CYAN)
tb(sl, Inches(.55), Inches(2.1), Inches(12.2), Inches(.42),
   "Obtenir des dons (prélèvements automatiques réguliers) pour permettre à l'association de CONTINUER son action humanitaire et d'AUGMENTER durablement le taux de dons.",
   size=12, bold=True, color=C_WHITE, wrap=True)

# 3 colonnes
cols3 = [
    (C_GREEN, C_GREEN_LIGHT, "🌍  L'UNICEF & SA MISSION",
     ["Organisation de l'ONU créée en 1946",
      "Présente dans 190 pays et territoires",
      "Défend les droits et la survie des enfants",
      "Programmes : nutrition, santé, eau, éducation",
      "Protection des enfants en zones de conflit"]),
    (C_BLUE_MID, C_BLUE_LIGHT, "💳  FIDELIS — NOTRE MANDAT",
     ["Fidelis est mandaté par l'UNICEF",
      "Mission : collecter des dons par téléphone",
      "Produit unique : prélèvement automatique régulier",
      "Nous représentons l'UNICEF lors de chaque appel",
      "Qualité du contact = image de l'UNICEF"]),
    (C_ORANGE, RGBColor(0xFF,0xF4,0xE6), "🎯  NOTRE OBJECTIF",
     ["Obtenir des dons réguliers et pérennes",
      "Augmenter le taux de dons de l'association",
      "Le refus d'aujourd'hui = le don de demain",
      "Chaque argument planté est une graine",
      "Chaque donateur = des vies impactées"]),
]
cx = Inches(.35)
for border, bg, title, pts in cols3:
    cw = Inches(4.12)
    rect(sl, cx, Inches(2.75), cw, Inches(3.62), fill=bg, line=border, line_w=Pt(2))
    rect(sl, cx, Inches(2.75), cw, Inches(.05), fill=border)
    tb(sl, cx+Inches(.15), Inches(2.84), cw-Inches(.3), Inches(.38),
       title, size=11, bold=True, color=border)
    py = Inches(3.3)
    for pt in pts:
        tb(sl, cx+Inches(.15), py, cw-Inches(.3), Inches(.36),
           f"✓  {pt}", size=10, color=C_TEXT_LIGHT)
        py += Inches(.4)
    cx += cw + Inches(.1)

rect(sl, Inches(.35), Inches(6.48), Inches(12.6), Inches(.5), fill=C_ORANGE_DARK)
tb(sl, Inches(.55), Inches(6.51), Inches(12.2), Inches(.44),
   '"LE REFUS D\'AUJOURD\'HUI PEUT ÊTRE LE DON DE DEMAIN."  —  Chaque appel compte.',
   size=12, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)
footer(sl)

# ════════════════════════════════════════════════════════════════════
# SLIDE 4 — NOMENCLATURE (1/2)
# ════════════════════════════════════════════════════════════════════
sl = add_slide()
rect(sl, 0, 0, W, H, fill=C_GREY)
header(sl, "02", "Nomenclature — Définitions Officielles (1/2)",
       "Termes fondamentaux à maîtriser avant le premier appel")

terms1 = [
    ("FIM",   "Formation Initiale Module",
     "Programme de formation dispensé le premier jour d'intégration d'un agent sur la campagne.", "J1"),
    ("CU",    "Contact Utile / Contact Argumenté",
     "Appel dont la durée de communication dépasse strictement 60 secondes. En-dessous de ce seuil, l'appel n'est pas comptabilisé.", "> 60 sec."),
    ("CU/H",  "Contacts Utiles par Heure",
     "Nombre de CU réalisés divisé par le nombre d'heures de production. Indicateur de productivité horaire.", "Objectif : 9"),
    ("PDC",   "Plan de Charge",
     "Volume de CU à produire sur le mois, fixé au niveau de la campagne ou de l'équipe par la direction.", "XXXX CU/mois"),
    ("TX",    "Taux de Transformation",
     "Pourcentage de CU débouchant sur un don (DEL ou PA). Mesure l'efficacité commerciale de l'agent.", "Selon assoc."),
]
hy = Inches(1.68)
for txt, ww, cx in [("SIGLE",Inches(1.0),Inches(.35)),
                     ("NOM COMPLET",Inches(3.2),Inches(1.4)),
                     ("DÉFINITION",Inches(6.5),Inches(4.65)),
                     ("VALEUR / OBJ.",Inches(1.9),Inches(11.2))]:
    rect(sl, cx, hy, ww, Inches(.38), fill=C_BLUE_DARK)
    tb(sl, cx+Inches(.06), hy, ww-Inches(.12), Inches(.38),
       txt, size=9, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)
ry = hy + Inches(.4)
for i,(sigle,nom,defin,val) in enumerate(terms1):
    bg = C_BLUE_LIGHT if i%2==0 else C_WHITE
    rh = Inches(.64)
    rect(sl, Inches(.35), ry, Inches(12.6), rh, fill=bg, line=C_GREY_LINE, line_w=Pt(.5))
    tb(sl, Inches(.4), ry+Inches(.1), Inches(.9), rh-Inches(.2),
       sigle, size=13, bold=True, color=C_BLUE_MID)
    tb(sl, Inches(1.42), ry+Inches(.1), Inches(3.1), rh-Inches(.2),
       nom, size=10, bold=True, color=C_BLUE_DARK)
    tb(sl, Inches(4.67), ry+Inches(.06), Inches(6.4), rh-Inches(.12),
       defin, size=10, color=C_TEXT_LIGHT, wrap=True)
    tb(sl, Inches(11.22), ry+Inches(.1), Inches(1.65), rh-Inches(.2),
       val, size=10, bold=True, color=C_ORANGE, align=PP_ALIGN.CENTER)
    ry += rh + Inches(.04)
footer(sl)

# ════════════════════════════════════════════════════════════════════
# SLIDE 5 — NOMENCLATURE (2/2)
# ════════════════════════════════════════════════════════════════════
sl = add_slide()
rect(sl, 0, 0, W, H, fill=C_GREY)
header(sl, "02", "Nomenclature — Définitions Officielles (2/2)",
       "Modes de prélèvement et qualifications d'appel")

# DEL corrigé : Don en Ligne
terms2 = [
    ("DEL",     "Don en Ligne",
     "Don validé par prélèvement automatique, finalisé en ligne (par l'agent ou par le donateur lui-même via le lien).", "Compte dans TX"),
    ("PA",      "Prélèvement Automatique",
     "Mode de don régulier prélevé automatiquement. Sur cette campagne, SEUL le PA régulier est proposé — aucun don ponctuel.", "Mode unique"),
    ("IBAN",    "Coordonnées bancaires",
     "Collectées à chaud pendant l'appel pour valider le PA immédiatement. L'agent valide lui-même le prélèvement en ligne.", "Mode A — À chaud"),
    ("DON",     "Qualification : Don",
     "CU dont l'issue est un PA confirmé, à chaud (IBAN) ou par promesse (lien envoyé). Entre dans le calcul du TX.", "Compte dans TX"),
    ("REF ARG.","Qualification : Refus Argumenté",
     "CU conclu par refus mais après présentation d'au moins un contre-argument. Qualifier uniquement si l'agent a vraiment argumenté.", "—"),
    ("INDÉCIS", "Qualification : Indécis",
     "CU conclu par hésitation du contact. Un lien ou rappel peut être prévu. Potentiel de transformation ultérieure.", "Suivi à prévoir"),
    ("ERR",     "Écouter → Reformuler → Rebondir",
     "Méthode officielle de traitement des objections (objet d'un book dédié). Ne jamais interrompre, toujours reformuler avant de répondre.", "Book dédié"),
]
hy = Inches(1.68)
for txt, ww, cx in [("SIGLE",Inches(1.1),Inches(.35)),
                     ("NOM COMPLET",Inches(3.0),Inches(1.5)),
                     ("DÉFINITION",Inches(6.5),Inches(4.55)),
                     ("VALEUR / OBJ.",Inches(1.9),Inches(11.1))]:
    rect(sl, cx, hy, ww, Inches(.34), fill=C_BLUE_DARK)
    tb(sl, cx+Inches(.06), hy, ww-Inches(.12), Inches(.34),
       txt, size=9, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)
ry = hy + Inches(.36)
for i,(sigle,nom,defin,val) in enumerate(terms2):
    bg = C_BLUE_LIGHT if i%2==0 else C_WHITE
    rh = Inches(.52)
    rect(sl, Inches(.35), ry, Inches(12.6), rh, fill=bg, line=C_GREY_LINE, line_w=Pt(.5))
    tb(sl, Inches(.4), ry+Inches(.06), Inches(1.0), rh-Inches(.12),
       sigle, size=11, bold=True, color=C_BLUE_MID)
    tb(sl, Inches(1.52), ry+Inches(.06), Inches(2.9), rh-Inches(.12),
       nom, size=10, bold=True, color=C_BLUE_DARK)
    tb(sl, Inches(4.57), ry+Inches(.04), Inches(6.4), rh-Inches(.08),
       defin, size=10, color=C_TEXT_LIGHT, wrap=True)
    tb(sl, Inches(11.12), ry+Inches(.06), Inches(1.7), rh-Inches(.12),
       val, size=10, bold=True, color=C_ORANGE, align=PP_ALIGN.CENTER)
    ry += rh + Inches(.025)
footer(sl)

# ════════════════════════════════════════════════════════════════════
# SLIDE 6 — KPI CU/H
# ════════════════════════════════════════════════════════════════════
sl = add_slide()
rect(sl, 0, 0, W, H, fill=C_GREY)
header(sl, "03", "KPI 1 — Contacts Utiles par Heure (CU/H)",
       "Mesure de la productivité horaire — objectif : 9 CU/H minimum")

rect(sl, Inches(.35), Inches(1.72), Inches(3.5), Inches(2.5), fill=C_BLUE_DARK)
tb(sl, Inches(.45), Inches(1.95), Inches(3.3), Inches(.5),
   "CU / H", size=18, bold=True, color=C_CYAN, align=PP_ALIGN.CENTER)
tb(sl, Inches(.45), Inches(2.45), Inches(3.3), Inches(.95),
   "9", size=62, bold=True, color=C_CYAN, align=PP_ALIGN.CENTER)
tb(sl, Inches(.45), Inches(3.48), Inches(3.3), Inches(.5),
   "OBJECTIF MINIMUM", size=11, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)

rect(sl, Inches(4.05), Inches(1.72), Inches(8.9), Inches(1.02), fill=C_BLUE_LIGHT)
tb(sl, Inches(4.22), Inches(1.78), Inches(8.5), Inches(.35),
   "Contact Utile (CU) — définition officielle", size=12, bold=True, color=C_BLUE_DARK)
tb(sl, Inches(4.22), Inches(2.1), Inches(8.5), Inches(.44),
   "Appel dont la durée de communication dépasse strictement 60 secondes. En-dessous, l'appel n'est PAS comptabilisé comme CU.",
   size=11, color=C_TEXT_LIGHT, wrap=True)

formula_box(sl, Inches(4.05), Inches(2.88), Inches(8.9), Inches(1.35),
            "FORMULE MATHÉMATIQUE — CU/H",
            [{"text":"CU/H  =  Nombre de CU réalisés  ÷  Heures de production",
              "color":RGBColor(0xE8,0xF4,0xFF),"size":13,"bold":True}])

formula_box(sl, Inches(.35), Inches(4.38), Inches(12.6), Inches(1.75),
            "EXEMPLE NUMÉRIQUE — Agent Philippe",
            [{"text":"CU réalisés      =  68",
              "color":RGBColor(0x7D,0xD3,0xFC),"size":12},
             {"text":"Heures produites =   7 h",
              "color":RGBColor(0x7D,0xD3,0xFC),"size":12},
             {"text":"CU/H  =  68  ÷  7  =  9,71  ✅  Objectif dépassé (≥ 9)",
              "color":RGBColor(0x34,0xD3,0x99),"size":13,"bold":True}],
            note="⚠  Heure de production = temps effectivement passé en ligne (hors pauses et formation)")
footer(sl)

# ════════════════════════════════════════════════════════════════════
# SLIDE 7 — KPI TX TRANSFORMATION
# ════════════════════════════════════════════════════════════════════
sl = add_slide()
rect(sl, 0, 0, W, H, fill=C_GREY)
header(sl, "04", "KPI 2 — Taux de Transformation",
       "% de dons (DEL / PA) parmi les CU — mesure de l'efficacité commerciale")

rect(sl, Inches(.35), Inches(1.72), Inches(3.5), Inches(2.5), fill=RGBColor(0x1A,0x3A,0x6B))
tb(sl, Inches(.45), Inches(1.95), Inches(3.3), Inches(.45),
   "TX TRANSFO.", size=14, bold=True, color=C_CYAN, align=PP_ALIGN.CENTER)
tb(sl, Inches(.45), Inches(2.42), Inches(3.3), Inches(.95),
   "%", size=62, bold=True, color=C_CYAN, align=PP_ALIGN.CENTER)
tb(sl, Inches(.45), Inches(3.48), Inches(3.3), Inches(.5),
   "CIBLE SELON ASSOCIATION", size=10, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)

formula_box(sl, Inches(4.05), Inches(1.72), Inches(8.9), Inches(2.55),
            "FORMULE MATHÉMATIQUE — TAUX DE TRANSFORMATION",
            [{"text":"TX (%)  =  ( Nombre de DEL ou PA  ×  100 )  ÷  Nombre de CU",
              "color":RGBColor(0xE8,0xF4,0xFF),"size":13,"bold":True},
             {"text":" ","color":C_WHITE,"size":6},
             {"text":"   DEL = Don en Ligne  (validé à chaud ou via le lien)",
              "color":RGBColor(0x94,0xA3,0xB8),"size":10},
             {"text":"   Seuls les DEL / PA confirmés entrent dans le numérateur",
              "color":RGBColor(0x94,0xA3,0xB8),"size":10}])

formula_box(sl, Inches(.35), Inches(4.4), Inches(12.6), Inches(1.75),
            "EXEMPLE NUMÉRIQUE — Agent Michel",
            [{"text":"CU réalisés      =  72",
              "color":RGBColor(0x7D,0xD3,0xFC),"size":12},
             {"text":"DEL (PA) obtenus =   3",
              "color":RGBColor(0x7D,0xD3,0xFC),"size":12},
             {"text":"TX  =  ( 3  ×  100 )  ÷  72  =  300  ÷  72  =  4,16 %",
              "color":RGBColor(0x34,0xD3,0x99),"size":13,"bold":True}],
            note="ℹ  Une promesse non finalisée (lien non cliqué) ne compte pas encore dans le numérateur.")
footer(sl)

# ════════════════════════════════════════════════════════════════════
# SLIDE 8 — KPI PDC
# ════════════════════════════════════════════════════════════════════
sl = add_slide()
rect(sl, 0, 0, W, H, fill=C_GREY)
header(sl, "05", "KPI 3 — Plan de Charge (PDC)",
       "Volume mensuel de CU à produire — pilotage collectif de l'équipe")

rect(sl, Inches(.35), Inches(1.72), Inches(3.5), Inches(2.5), fill=C_PURPLE)
tb(sl, Inches(.45), Inches(1.95), Inches(3.3), Inches(.5),
   "PDC", size=22, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)
tb(sl, Inches(.45), Inches(2.48), Inches(3.3), Inches(.9),
   "XXXX", size=38, bold=True, color=RGBColor(0xD8,0xB4,0xFE), align=PP_ALIGN.CENTER)
tb(sl, Inches(.45), Inches(3.46), Inches(3.3), Inches(.48),
   "CU / MOIS", size=12, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)

formula_box(sl, Inches(4.05), Inches(1.72), Inches(8.9), Inches(2.55),
            "FORMULE DE VÉRIFICATION — PDC",
            [{"text":"PDC atteint  =  CU/H moyen équipe  ×  Heures produites totales",
              "color":RGBColor(0xE8,0xF4,0xFF),"size":13,"bold":True},
             {"text":" ","color":C_WHITE,"size":6},
             {"text":"CU/H requis  =  PDC mensuel cible  ÷  Heures totales équipe",
              "color":RGBColor(0xC4,0xB5,0xFD),"size":11}])

formula_box(sl, Inches(.35), Inches(4.4), Inches(12.6), Inches(1.5),
            "EXEMPLE — PDC cible 1 800 CU / équipe 200 h de production",
            [{"text":"CU/H requis  =  1 800  ÷  200  =  9 CU/H  =  objectif fixé ✅",
              "color":RGBColor(0x34,0xD3,0x99),"size":13,"bold":True}],
            note="Le PDC exact de la campagne sera communiqué par votre superviseur en début de mission.")

rect(sl, Inches(.35), Inches(6.08), Inches(12.6), Inches(.5),
     fill=C_BLUE_LIGHT, line=C_BLUE_MID, line_w=Pt(1))
tb(sl, Inches(.55), Inches(6.12), Inches(12.2), Inches(.42),
   "RAPPEL : Chaque CU/H à 9 sur 200h d'équipe = 1 800 CU/mois. C'est pourquoi l'objectif individuel de 9 CU/H est non négociable.",
   size=11, color=C_BLUE_DARK, bold=True)
footer(sl)

# ════════════════════════════════════════════════════════════════════
# SLIDE 9 — LES 3 QUALIFICATIONS
# ════════════════════════════════════════════════════════════════════
sl = add_slide()
rect(sl, 0, 0, W, H, fill=C_GREY)
header(sl, "06", "Les 3 Qualifications de Contact Utile",
       "Tout appel > 60 secondes doit être qualifié immédiatement dans l'outil")

qw = Inches(4.1)
qdata = [
    (C_GREEN, C_GREEN_LIGHT, "💚", "DON",
     "Le contact accepte le prélèvement automatique régulier.",
     ["PA confirmé à chaud (IBAN collecté pendant l'appel)",
      "OU lien envoyé avec engagement verbal du donateur",
      "Entre dans le calcul du Taux de Transformation",
      "Qualifier DON uniquement si PA effectivement initié"],
     C_GREEN),
    (C_RED, C_RED_LIGHT, "❌", "REFUS ARGUMENTÉ",
     "Le contact refuse après avoir entendu l'argumentaire.",
     ["L'agent a présenté au moins 1 contre-argument",
      "Le contact maintient son refus après échange",
      "Ne PAS qualifier sans avoir vraiment argumenté",
      "Ne compte pas dans le Taux de Transformation"],
     C_RED),
    (C_ORANGE, RGBColor(0xFF,0xF4,0xE6), "⏳", "INDÉCIS",
     "Le contact hésite ou souhaite réfléchir.",
     ["Le contact est réceptif mais ne décide pas encore",
      "Un lien de don peut lui être envoyé",
      "Un rappel ultérieur peut être planifié",
      "Potentiel de transformation — le don de demain"],
     C_ORANGE),
]
cx = Inches(.35)
for border, bg, icon, name, desc, pts, nc in qdata:
    rect(sl, cx, Inches(1.72), qw, Inches(5.1), fill=bg, line=border, line_w=Pt(2))
    rect(sl, cx, Inches(1.72), qw, Inches(.055), fill=border)
    tb(sl, cx, Inches(1.82), qw, Inches(.65), icon, size=30, align=PP_ALIGN.CENTER, color=border)
    tb(sl, cx, Inches(2.48), qw, Inches(.42),
       name, size=13, bold=True, color=border, align=PP_ALIGN.CENTER)
    tb(sl, cx+Inches(.18), Inches(2.92), qw-Inches(.36), Inches(.55),
       desc, size=10, color=C_TEXT_LIGHT, align=PP_ALIGN.CENTER, wrap=True)
    py = Inches(3.52)
    for pt in pts:
        tb(sl, cx+Inches(.18), py, qw-Inches(.36), Inches(.36),
           f"✓  {pt}", size=10, color=RGBColor(0x1A,0x1A,0x2E))
        py += Inches(.38)
    cx += qw + Inches(.13)

rect(sl, Inches(.35), Inches(6.94), Inches(12.6), Inches(.46),
     fill=C_YELLOW_LIGHT, line=C_YELLOW, line_w=Pt(1.5))
tb(sl, Inches(.55), Inches(6.97), Inches(12.2), Inches(.4),
   "RÈGLE ABSOLUE : Qualifier DON uniquement si un PA est effectivement initié (IBAN ou lien avec engagement). Tout doute = INDÉCIS.",
   size=11, bold=True, color=RGBColor(0x78,0x35,0x00))
footer(sl)

# ════════════════════════════════════════════════════════════════════
# SLIDE 10 — LE PRODUIT : PA RÉGULIER
# ════════════════════════════════════════════════════════════════════
sl = add_slide()
rect(sl, 0, 0, W, H, fill=C_GREY)
header(sl, "07", "Le Produit — Prélèvement Automatique Régulier",
       "Position officielle SHY Performance sur cette campagne Fidelis × UNICEF")

rect(sl, Inches(.35), Inches(1.72), Inches(12.6), Inches(.84), fill=C_BLUE_MID)
tb(sl, Inches(.55), Inches(1.78), Inches(12.2), Inches(.35),
   "RÈGLE D'OR DE LA CAMPAGNE", size=9, bold=True, color=C_CYAN)
tb(sl, Inches(.55), Inches(2.08), Inches(12.2), Inches(.38),
   "Fidelis propose plusieurs types de prélèvement. Sur NOTRE campagne : SEUL le prélèvement automatique régulier est proposé. Aucun don ponctuel.",
   size=12, bold=True, color=C_WHITE, wrap=True)

# Mode A
rect(sl, Inches(.35), Inches(2.72), Inches(6.0), Inches(2.75),
     fill=C_GREEN_LIGHT, line=C_GREEN, line_w=Pt(2))
rect(sl, Inches(.35), Inches(2.72), Inches(6.0), Inches(.055), fill=C_GREEN)
tb(sl, Inches(.5), Inches(2.82), Inches(5.7), Inches(.38),
   "MODE A — Prélèvement À CHAUD  (prioritaire)", size=12, bold=True, color=C_GREEN)
tb(sl, Inches(.5), Inches(3.24), Inches(5.7), Inches(2.14),
   ("L'agent recueille l'IBAN du donateur pendant l'appel.\n"
    "L'agent valide lui-même le prélèvement en ligne.\n\n"
    "✓  Don sécurisé immédiatement\n"
    "✓  Taux de concrétisation maximal\n"
    "✓  Qualification : DON"),
   size=11, color=C_TEXT_LIGHT, wrap=True)

# Mode B
rect(sl, Inches(6.7), Inches(2.72), Inches(6.25), Inches(2.75),
     fill=RGBColor(0xFF,0xF4,0xE6), line=C_ORANGE, line_w=Pt(2))
rect(sl, Inches(6.7), Inches(2.72), Inches(6.25), Inches(.055), fill=C_ORANGE)
tb(sl, Inches(6.88), Inches(2.82), Inches(5.9), Inches(.38),
   "MODE B — Promesse par LIEN", size=12, bold=True, color=C_ORANGE)
tb(sl, Inches(6.88), Inches(3.24), Inches(5.9), Inches(2.14),
   ("L'agent envoie un lien sécurisé au donateur.\n"
    "Le donateur finalise lui-même en ligne après l'appel.\n\n"
    "✓  Le lien n'a PAS de date de validité — il est permanent\n"
    "✓  Redirige vers le site de l'association (UNICEF)\n"
    "✓  Un e-mail de confirmation est envoyé au donateur"),
   size=11, color=C_TEXT_LIGHT, wrap=True)

# Récap lien
rect(sl, Inches(.35), Inches(5.6), Inches(12.6), Inches(1.0),
     fill=C_BLUE_LIGHT, line=C_BLUE_MID, line_w=Pt(1.5))
tb(sl, Inches(.55), Inches(5.68), Inches(12.2), Inches(.35),
   "CARACTÉRISTIQUES DU LIEN DE DON", size=11, bold=True, color=C_BLUE_DARK)
tb(sl, Inches(.55), Inches(6.0), Inches(12.2), Inches(.52),
   "Aucune date d'expiration  ·  Redirige vers le site officiel de l'association  ·  E-mail de confirmation automatique au donateur  ·  Ne jamais mentionner de date limite.",
   size=11, color=C_TEXT_LIGHT, wrap=True)
footer(sl)

# ════════════════════════════════════════════════════════════════════
# SLIDE 11 — PHRASE D'ACCROCHE (page 1 : règles)
# ════════════════════════════════════════════════════════════════════
sl = add_slide()
rect(sl, 0, 0, W, H, fill=C_GREY)
header(sl, "08", "La Phrase d'Accroche — Règles Officielles",
       "Script d'ouverture d'appel — chaque règle est obligatoire")

# Règles numérotées
rules = [
    ("1", C_BLUE_MID,
     'Dire "Allô ?" et attendre la réponse',
     ("Ne pas parler en premier. Prononcer uniquement « Allô ? » puis attendre. "
      "Le premier mot du contact permet d'identifier le genre (voix masculine → M. / voix féminine → Mme).")),
    ("2", C_TEAL,
     "Identifier le donateur : Prénom + Nom (jamais l'inverse)",
     ("En français, on dit toujours le prénom en premier, suivi du nom. "
      "Exemples corrects : « Manuel Macron », « Christiano Ronaldo », « Donald Trump ».\n"
      "❌ JAMAIS : « Macron Manuel » — ce n'est pas français.")),
    ("3", C_GREEN,
     'Formule de politesse : "Bonjour M. [Prénom Nom]"',
     ("Après identification, saluer chaleureusement. Utiliser le genre identifié à l'étape 1. "
      "La politesse dès les premières secondes empêche le raccrochage immédiat.")),
    ("4", C_ORANGE,
     "Se présenter et présenter Fidelis mandaté par l'UNICEF",
     ("Donner son prénom, préciser que vous appelez de la part de Fidelis, "
      "mandaté par l'UNICEF. Légitimer immédiatement l'appel par la notoriété de l'UNICEF.")),
    ("5", C_PURPLE,
     "Créer le lien — question sur la connaissance de l'UNICEF",
     ("Demander au donateur s'il connaît les actions humanitaires de l'UNICEF. "
      "Quelle que soit la réponse, expliquer la mission. Cela crée l'engagement et dépasse les 60 secondes.")),
]

ry = Inches(1.72)
for num, color, title, desc in rules:
    rh = Inches(1.0)
    rect(sl, Inches(.35), ry, Inches(12.6), rh, fill=C_WHITE,
         line=C_GREY_LINE, line_w=Pt(.5))
    rect(sl, Inches(.35), ry, Inches(.5), rh, fill=color)
    tb(sl, Inches(.35), ry, Inches(.5), rh,
       num, size=18, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)
    tb(sl, Inches(.96), ry+Inches(.08), Inches(11.8), Inches(.36),
       title, size=12, bold=True, color=color)
    tb(sl, Inches(.96), ry+Inches(.46), Inches(11.8), Inches(.48),
       desc, size=10, color=C_TEXT_LIGHT, wrap=True)
    ry += rh + Inches(.04)
footer(sl)

# ════════════════════════════════════════════════════════════════════
# SLIDE 12 — PHRASE D'ACCROCHE (page 2 : script type)
# ════════════════════════════════════════════════════════════════════
sl = add_slide()
rect(sl, 0, 0, W, H, fill=C_GREY)
header(sl, "08", "La Phrase d'Accroche — Script Type",
       "Exemple complet d'ouverture d'appel respectant les 5 règles officielles")

# Script visuel séquentiel
rect(sl, Inches(.35), Inches(1.72), Inches(12.6), Inches(.6), fill=C_BLUE_LIGHT)
tb(sl, Inches(.55), Inches(1.76), Inches(12.2), Inches(.52),
   "SCRIPT TYPE — à adapter selon le genre identifié (M. / Mme) et le prénom/nom du fichier",
   size=12, bold=True, color=C_BLUE_DARK)

script_steps = [
    (C_BLUE_MID,   "ÉTAPE 1 — Décrochage",
     '"Allô ?"',
     "→ Attendre la réponse. Identifier la voix : masculine = M.  /  féminine = Mme."),
    (C_TEAL,       "ÉTAPE 2 — Identification",
     '"Bonjour, M. [Prénom] [Nom] ?"',
     '→ Toujours Prénom puis Nom. Exemple : "Bonjour, M. Manuel Macron ?"  |  Jamais "M. Macron Manuel".'),
    (C_GREEN,      "ÉTAPE 3 — Présentation",
     '"Je suis [Prénom], j\'appelle de la part de Fidelis, mandaté par l\'UNICEF."',
     "→ Ton chaleureux et posé. La mention UNICEF légitime l'appel immédiatement."),
    (C_ORANGE,     "ÉTAPE 4 — Création du lien",
     '"Est-ce que vous avez des connaissances sur les actions humanitaires de l\'UNICEF ?"',
     "→ Question ouverte. Quelle que soit la réponse, continuer avec la présentation de la mission."),
    (C_PURPLE,     "ÉTAPE 5 — Présentation de la mission",
     '"L\'UNICEF est présent dans 190 pays pour protéger les enfants : nutrition, santé, eau potable, éducation…"',
     "→ Créer l'émotion. Personnaliser si le donateur connaît déjà. Préparer la proposition de don."),
]

ry = Inches(2.46)
for color, etape, phrase, note in script_steps:
    rh = Inches(.88)
    rect(sl, Inches(.35), ry, Inches(12.6), rh, fill=C_WHITE,
         line=color, line_w=Pt(1.5))
    rect(sl, Inches(.35), ry, Inches(.06), rh, fill=color)
    tb(sl, Inches(.5), ry+Inches(.05), Inches(2.0), Inches(.28),
       etape, size=9, bold=True, color=color)
    tb(sl, Inches(.5), ry+Inches(.32), Inches(5.5), Inches(.38),
       phrase, size=11, bold=True, color=C_BLUE_DARK, italic=True, wrap=True)
    rect(sl, Inches(6.2), ry+Inches(.08), Inches(.02), rh-Inches(.16),
         fill=C_GREY_LINE)
    tb(sl, Inches(6.35), ry+Inches(.08), Inches(6.45), rh-Inches(.16),
       note, size=10, color=C_TEXT_LIGHT, italic=True, wrap=True)
    ry += rh + Inches(.04)
footer(sl)

# ════════════════════════════════════════════════════════════════════
# SLIDE 13 — CLÔTURE
# ════════════════════════════════════════════════════════════════════
sl = add_slide()
rect(sl, 0, 0, W, H, fill=C_BLUE_DARK)
rect(sl, 0, 0, W, Inches(.07), fill=C_ORANGE)
rect(sl, 0, H-Inches(.07), W, Inches(.07), fill=C_CYAN)
rect(sl, Inches(6.6), 0, Inches(.055), H, fill=C_CYAN)

# Gauche — mantra + récap
tb(sl, Inches(.5), Inches(1.5), Inches(5.8), Inches(.42),
   "POUR RETENIR", size=11, bold=True, color=C_ORANGE)
tb(sl, Inches(.5), Inches(2.0), Inches(5.8), Inches(1.15),
   '"LE REFUS D\'AUJOURD\'HUI\nPEUT ÊTRE\nLE DON DE DEMAIN."',
   size=22, bold=True, color=C_WHITE, wrap=True)
tb(sl, Inches(.5), Inches(3.25), Inches(5.8), Inches(1.6),
   ("Chaque appel est une opportunité de changer une vie.\n\n"
    "Votre voix, votre conviction, votre argument —\n"
    "c'est ce qui transforme une hésitation\nen don concret pour l'UNICEF."),
   size=12, italic=True, color=RGBColor(0xBB,0xCC,0xEE), wrap=True)

# Droite — chiffres clés
tb(sl, Inches(7.0), Inches(1.45), Inches(5.9), Inches(.42),
   "LES CHIFFRES CLÉS DU J1", size=11, bold=True, color=C_CYAN)
kpis = [
    ("9",      "CU/H — objectif minimum"),
    ("> 60s",  "Durée seuil d'un CU"),
    ("3",      "Qualifications : DON / REFUS ARG. / INDÉCIS"),
    ("2",      "Modes de PA : à chaud (IBAN) et promesse (lien)"),
    ("5",      "Étapes de la phrase d'accroche"),
    ("∞",      "Validité du lien de don — aucune expiration"),
]
ky = Inches(2.0)
for val, lbl in kpis:
    rect(sl, Inches(7.0), ky, Inches(5.9), Inches(.62), fill=RGBColor(0x00,0x40,0xA0))
    tb(sl, Inches(7.08), ky+Inches(.08), Inches(1.15), Inches(.46),
       val, size=18, bold=True, color=C_CYAN, align=PP_ALIGN.CENTER)
    tb(sl, Inches(8.32), ky+Inches(.16), Inches(4.5), Inches(.3),
       lbl, size=12, color=C_WHITE)
    ky += Inches(.7)

rect(sl, Inches(.35), H-Inches(1.22), Inches(12.6), Inches(.62), fill=C_ORANGE_DARK)
tb(sl, Inches(.55), H-Inches(1.2), Inches(12.2), Inches(.58),
   "Le traitement des objections, le script complet et les techniques avancées feront l'objet d'un book dédié séparé.",
   size=12, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)

footer(sl, "Fidelis × SHY Performance  |  Module J1 FIM  |  Formation Initiale Module  |  Confidentiel")

# ════════════════════════════════════════════════════════════════════
# SAUVEGARDE
# ════════════════════════════════════════════════════════════════════
out = "/home/user/fatou/J1_FIM_Fidelis_Formation.pptx"
prs.save(out)
print(f"✅  Fichier enregistré : {out}")
print(f"   Nombre de slides   : {len(prs.slides)}")

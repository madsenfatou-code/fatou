"""
Génère le book de formation J1 FIM Fidelis × SHY Performance au format PowerPoint.
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.oxml.ns import qn
from pptx.enum.dml import MSO_THEME_COLOR
import copy
from lxml import etree

# ── Palette couleurs ────────────────────────────────────────────────
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

# ── Dimensions slide 16:9 ────────────────────────────────────────────
W = Inches(13.33)
H = Inches(7.5)

prs = Presentation()
prs.slide_width  = W
prs.slide_height = H

BLANK = prs.slide_layouts[6]   # layout totalement vide


# ════════════════════════════════════════════════════════════════════
#  HELPERS
# ════════════════════════════════════════════════════════════════════

def add_slide():
    return prs.slides.add_slide(BLANK)


def rect(slide, x, y, w, h, fill=None, line=None, line_w=None):
    """Ajoute un rectangle avec couleur de fond et/ou contour."""
    shape = slide.shapes.add_shape(1, x, y, w, h)  # MSO_SHAPE_TYPE.RECTANGLE = 1
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


def txbox(slide, x, y, w, h, text, size=18, bold=False, color=C_WHITE,
          align=PP_ALIGN.LEFT, wrap=True, italic=False, font="Segoe UI"):
    """Ajoute une zone de texte simple."""
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
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
    return tb


def txbox_multi(slide, x, y, w, h, lines, wrap=True):
    """
    lines = list of dict:
      { text, size, bold, color, align, italic, space_before, bullet }
    """
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = wrap
    first = True
    for ln in lines:
        if first:
            p = tf.paragraphs[0]
            first = False
        else:
            p = tf.add_paragraph()
        p.alignment = ln.get("align", PP_ALIGN.LEFT)
        if ln.get("space_before"):
            p.space_before = Pt(ln["space_before"])
        run = p.add_run()
        run.text = ln.get("text", "")
        run.font.size = Pt(ln.get("size", 18))
        run.font.bold = ln.get("bold", False)
        run.font.italic = ln.get("italic", False)
        run.font.color.rgb = ln.get("color", C_WHITE)
        run.font.name = ln.get("font", "Segoe UI")
    return tb


def gradient_rect(slide, x, y, w, h, c1, c2, angle=None):
    """Rectangle avec gradient (approximé par 2 rectangles superposés)."""
    r1 = rect(slide, x, y, w, h, fill=c1)
    return r1


def bullet_box(slide, x, y, w, h, title, items,
               title_color=C_BLUE_DARK, item_color=C_TEXT_LIGHT,
               check_color=C_GREEN, bg=None, border=None,
               title_size=13, item_size=11):
    """Bloc avec titre gras + liste à puces."""
    if bg:
        rect(slide, x, y, w, h, fill=bg, line=border, line_w=Pt(1.5) if border else None)
    lines = [{"text": title, "size": title_size, "bold": True, "color": title_color, "font": "Segoe UI"}]
    for item in items:
        lines.append({
            "text": f"  ✓  {item}",
            "size": item_size, "bold": False,
            "color": item_color, "font": "Segoe UI",
            "space_before": 2
        })
    txbox_multi(slide, x + Inches(.18), y + Inches(.15),
                w - Inches(.36), h - Inches(.3), lines)


def stop_box(slide, x, y, w, h, title, items,
             title_color=C_RED, item_color=C_TEXT_LIGHT, bg=C_RED_LIGHT, border=C_RED,
             title_size=13, item_size=11):
    rect(slide, x, y, w, h, fill=bg, line=border, line_w=Pt(1.5))
    lines = [{"text": title, "size": title_size, "bold": True, "color": title_color, "font": "Segoe UI"}]
    for item in items:
        lines.append({
            "text": f"  ✗  {item}",
            "size": item_size, "bold": False,
            "color": item_color, "font": "Segoe UI",
            "space_before": 2
        })
    txbox_multi(slide, x + Inches(.18), y + Inches(.15),
                w - Inches(.36), h - Inches(.3), lines)


def section_header(slide, num, title, subtitle=""):
    """Bandeau titre de section en haut."""
    rect(slide, 0, 0, W, Inches(1.55), fill=C_BLUE_DARK)
    # numéro badge
    rect(slide, Inches(.35), Inches(.28), Inches(.85), Inches(.42), fill=C_CYAN)
    txbox(slide, Inches(.35), Inches(.28), Inches(.85), Inches(.42),
          num, size=14, bold=True, color=C_BLUE_DARK, align=PP_ALIGN.CENTER)
    txbox(slide, Inches(1.35), Inches(.2), Inches(11.5), Inches(.6),
          title, size=24, bold=True, color=C_WHITE)
    if subtitle:
        txbox(slide, Inches(1.35), Inches(.78), Inches(11.5), Inches(.5),
              subtitle, size=13, color=C_CYAN, italic=True)


def footer_bar(slide, text="J1 FIM — Fidelis × SHY Performance  |  Formation Initiale Module"):
    rect(slide, 0, H - Inches(.45), W, Inches(.45), fill=C_BLUE_DARK)
    txbox(slide, Inches(.3), H - Inches(.42), W - Inches(.6), Inches(.38),
          text, size=9, color=RGBColor(0xAA, 0xBB, 0xDD), align=PP_ALIGN.CENTER)


def formula_box(slide, x, y, w, h, label, formula_lines, note=""):
    """Bloc formule fond sombre."""
    rect(slide, x, y, w, h, fill=C_DARK, line=C_BLUE_MID, line_w=Pt(1))
    # label
    txbox(slide, x + Inches(.2), y + Inches(.12), w - Inches(.4), Inches(.28),
          label, size=8, bold=True, color=C_CYAN, font="Segoe UI")
    # formule (lignes)
    fy = y + Inches(.42)
    for fl in formula_lines:
        txbox(slide, x + Inches(.2), fy, w - Inches(.4), Inches(.4),
              fl["text"], size=fl.get("size", 13), bold=fl.get("bold", False),
              color=fl.get("color", C_WHITE), font="Courier New")
        fy += Inches(.38)
    if note:
        txbox(slide, x + Inches(.2), y + h - Inches(.38), w - Inches(.4), Inches(.34),
              note, size=8, italic=True, color=RGBColor(0x94, 0xA3, 0xB8))


# ════════════════════════════════════════════════════════════════════
#  SLIDE 1 — COUVERTURE
# ════════════════════════════════════════════════════════════════════

sl = add_slide()

# Fond bleu dégradé (bloc gauche + bloc droit)
rect(sl, 0, 0, W * 0.62, H, fill=C_BLUE_DARK)
rect(sl, W * 0.62, 0, W * 0.38, H, fill=C_BLUE_MID)

# Formes décoratives
rect(sl, W * 0.58, 0, Inches(.06), H, fill=C_CYAN)
rect(sl, 0, H * 0.72, W * 0.62, Inches(.05), fill=C_ORANGE)

# Badge
rect(sl, Inches(.5), Inches(.55), Inches(2.6), Inches(.42), fill=C_ORANGE)
txbox(sl, Inches(.5), Inches(.55), Inches(2.6), Inches(.42),
      "FORMATION INITIALE MODULE  ·  J1",
      size=10, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)

# Titre principal
txbox_multi(sl, Inches(.5), Inches(1.2), Inches(7.5), Inches(2.5), [
    {"text": "MODULE J1 FIM", "size": 42, "bold": True, "color": C_WHITE, "font": "Segoe UI"},
    {"text": "Collecte de Dons", "size": 30, "bold": False, "color": C_CYAN, "font": "Segoe UI"},
    {"text": "Humanitaires", "size": 30, "bold": False, "color": C_CYAN, "font": "Segoe UI"},
])

# Sous-titre
txbox(sl, Inches(.5), Inches(3.9), Inches(7.5), Inches(.85),
      "Formation complète à destination des agents SHY Performance.\nMaîtriser les KPI, les formules, la structure d'appel et l'argumentation.",
      size=13, color=RGBColor(0xBB, 0xCC, 0xEE), wrap=True)

# Infos droite
rect(sl, W * 0.64, Inches(.6), Inches(4.4), Inches(5.8), fill=RGBColor(0x00, 0x40, 0xA0))
infos = [
    ("🎯", "Objectif CU/H", "9 minimum"),
    ("⏱", "Durée seuil CU", "> 60 secondes"),
    ("💳", "Produit unique", "Prélèvement auto régulier"),
    ("📊", "KPI principaux", "CU/H & TX Transformation"),
    ("🌍", "Mission", "Action humanitaire"),
]
iy = Inches(1.0)
for icon, lab, val in infos:
    txbox(sl, W * 0.66, iy, Inches(4.0), Inches(.3), f"{icon}  {lab}", size=10, color=C_CYAN, bold=True)
    txbox(sl, W * 0.66, iy + Inches(.28), Inches(4.0), Inches(.32), f"     {val}", size=12, color=C_WHITE)
    iy += Inches(.9)

# Mantra bas
rect(sl, 0, H - Inches(1.2), W * 0.62, Inches(.75), fill=C_ORANGE_DARK)
txbox(sl, Inches(.5), H - Inches(1.18), W * 0.58, Inches(.72),
      '"LE REFUS D\'AUJOURD\'HUI PEUT ÊTRE LE DON DE DEMAIN."',
      size=13, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)

# Footer
footer_bar(sl)


# ════════════════════════════════════════════════════════════════════
#  SLIDE 2 — SOMMAIRE
# ════════════════════════════════════════════════════════════════════

sl = add_slide()
rect(sl, 0, 0, W, H, fill=C_GREY)
section_header(sl, "PLAN", "Sommaire du Module J1", "10 sections progressives — de la mission à la performance")

sections = [
    ("01", "Mission & Finalité",          "Comprendre l'enjeu humanitaire"),
    ("02", "Nomenclature",                "Définitions officielles des termes"),
    ("03", "KPI & Formules",              "CU/H, Taux de Transformation, PDC"),
    ("04", "Les 3 Qualifications",        "DON / REFUS ARGUMENTÉ / INDÉCIS"),
    ("05", "Le Produit",                  "Prélèvement automatique régulier"),
    ("06", "Structure d'un Appel",        "7 étapes séquentielles"),
    ("07", "Traitement des Objections",   "Méthode ERR"),
    ("08", "Posture & Communication",     "Triangle Émotion / Raison / Facilité"),
    ("09", "Auto-Suivi de Performance",   "Tableau de bord horaire"),
    ("10", "Checklist J1",                "Suis-je prêt à décrocher ?"),
]

cols = [sections[:5], sections[5:]]
xs = [Inches(.4), Inches(6.85)]
for col, cx in zip(cols, xs):
    y = Inches(1.75)
    for num, title, sub in col:
        rect(sl, cx, y, Inches(6.2), Inches(.88), fill=C_WHITE,
             line=C_BLUE_LIGHT, line_w=Pt(1))
        rect(sl, cx, y, Inches(.55), Inches(.88), fill=C_BLUE_MID)
        txbox(sl, cx, y, Inches(.55), Inches(.88),
              num, size=10, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)
        txbox(sl, cx + Inches(.62), y + Inches(.06), Inches(5.5), Inches(.34),
              title, size=13, bold=True, color=C_BLUE_DARK)
        txbox(sl, cx + Inches(.62), y + Inches(.42), Inches(5.5), Inches(.3),
              sub, size=10, italic=True, color=C_TEXT_LIGHT)
        y += Inches(.98)

footer_bar(sl)


# ════════════════════════════════════════════════════════════════════
#  SLIDE 3 — SYNOPSIS & OBJECTIFS
# ════════════════════════════════════════════════════════════════════

sl = add_slide()
rect(sl, 0, 0, W, H, fill=C_GREY)
section_header(sl, "INTRO", "Synopsis & Objectifs Pédagogiques J1",
               "Ce que vous maîtriserez à l'issue de cette journée")

# Synopsis bloc gauche
rect(sl, Inches(.35), Inches(1.75), Inches(6.1), Inches(5.2), fill=C_BLUE_DARK)
txbox(sl, Inches(.55), Inches(1.9), Inches(5.7), Inches(.35),
      "SYNOPSIS", size=10, bold=True, color=C_CYAN)
txbox(sl, Inches(.55), Inches(2.28), Inches(5.7), Inches(4.4),
      ("Ce module de formation initiale prépare chaque agent à représenter "
       "dignement l'association partenaire lors de campagnes de collecte "
       "téléphonique.\n\n"
       "À l'issue du J1, l'agent sera capable de :\n"
       "  • Comprendre les enjeux humanitaires\n"
       "  • Maîtriser les KPI et les calculer en temps réel\n"
       "  • Conduire un appel structuré et argumenté\n"
       "  • Traiter les objections avec méthode\n"
       "  • Suivre sa propre performance heure par heure"),
      size=12, color=C_WHITE, wrap=True)

# Objectifs droite
rect(sl, Inches(6.75), Inches(1.75), Inches(6.2), Inches(5.2), fill=C_WHITE,
     line=C_BLUE_LIGHT, line_w=Pt(1.5))
txbox(sl, Inches(6.95), Inches(1.9), Inches(5.8), Inches(.35),
      "OBJECTIFS PÉDAGOGIQUES", size=10, bold=True, color=C_ORANGE)
objectifs = [
    "Connaître la finalité humanitaire de la mission",
    "Maîtriser la nomenclature (CU, PDC, TX, DEL, PA…)",
    "Calculer CU/H et Taux de Transformation",
    "Distinguer les 3 qualifications de CU",
    "Présenter le produit PA régulier correctement",
    "Conduire un appel en 7 étapes",
    "Traiter 7 objections clés avec la méthode ERR",
    "Suivre sa performance autonomement",
]
oy = Inches(2.3)
for obj in objectifs:
    txbox(sl, Inches(6.95), oy, Inches(5.8), Inches(.36),
          f"✓  {obj}", size=11, color=C_BLUE_DARK)
    oy += Inches(.38)

footer_bar(sl)


# ════════════════════════════════════════════════════════════════════
#  SLIDE 4 — SECTION 1 : MISSION & FINALITÉ
# ════════════════════════════════════════════════════════════════════

sl = add_slide()
rect(sl, 0, 0, W, H, fill=C_GREY)
section_header(sl, "01", "Mission & Finalité de la Campagne",
               "Comprendre l'enjeu humanitaire pour argumenter avec conviction")

# Bloc finalité centre
rect(sl, Inches(.35), Inches(1.75), Inches(12.6), Inches(1.05), fill=C_BLUE_MID)
txbox(sl, Inches(.55), Inches(1.82), Inches(12.2), Inches(.45),
      "FINALITÉ OFFICIELLE", size=9, bold=True, color=C_CYAN)
txbox(sl, Inches(.55), Inches(2.12), Inches(12.2), Inches(.5),
      "Obtenir des dons (prélèvements automatiques réguliers) pour permettre à l'association de CONTINUER son action humanitaire et d'AUGMENTER durablement le taux de dons.",
      size=12, bold=True, color=C_WHITE, wrap=True)

# 2 colonnes
bullet_box(sl, Inches(.35), Inches(2.98), Inches(6.0), Inches(2.1),
           "Ce que votre appel finance concrètement",
           ["Programmes d'aide humanitaire d'urgence",
            "Nutrition et santé pour les enfants",
            "Accès à l'eau potable et à l'éducation",
            "Protection des populations vulnérables"],
           bg=C_GREEN_LIGHT, border=C_GREEN,
           title_color=C_GREEN, item_color=C_TEXT_LIGHT)

bullet_box(sl, Inches(6.6), Inches(2.98), Inches(6.35), Inches(2.1),
           "Pourquoi l'argumentation compte à long terme",
           ["Un donateur convaincu s'engage sur plusieurs années",
            "Un refus bien géré peut revenir donateur demain",
            "L'image de l'association dépend de chaque échange",
            "Votre posture = ambassadeur de la cause humanitaire"],
           bg=RGBColor(0xFF, 0xF4, 0xE6), border=C_ORANGE,
           title_color=C_ORANGE, item_color=C_TEXT_LIGHT)

# Mantra bas
rect(sl, Inches(.35), Inches(5.25), Inches(12.6), Inches(.65), fill=C_ORANGE_DARK)
txbox(sl, Inches(.55), Inches(5.28), Inches(12.2), Inches(.58),
      '"LE REFUS D\'AUJOURD\'HUI PEUT ÊTRE LE DON DE DEMAIN."  —  Chaque argument planté est une graine pour l\'avenir.',
      size=12, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)

footer_bar(sl)


# ════════════════════════════════════════════════════════════════════
#  SLIDE 5 — SECTION 2 : NOMENCLATURE (page 1/2)
# ════════════════════════════════════════════════════════════════════

sl = add_slide()
rect(sl, 0, 0, W, H, fill=C_GREY)
section_header(sl, "02", "Nomenclature — Définitions Officielles (1/2)",
               "Termes fondamentaux à maîtriser avant le premier appel")

terms1 = [
    ("FIM",  "Formation Initiale Module",
     "Programme de formation dispensé le premier jour d'intégration d'un agent sur la campagne.", "J1"),
    ("CU",   "Contact Utile / Contact Argumenté",
     "Appel dont la durée dépasse strictement 60 secondes. En-dessous, l'appel n'est pas comptabilisé.", "> 60 secondes"),
    ("CU/H", "Contacts Utiles par Heure",
     "Nombre de CU réalisés ÷ nombre d'heures de production. Mesure la productivité horaire.", "Objectif : 9"),
    ("PDC",  "Plan de Charge",
     "Volume de CU à produire sur le mois, fixé au niveau de la campagne ou de l'équipe.", "XXXX CU/mois"),
    ("TX",   "Taux de Transformation",
     "Pourcentage de CU débouchant sur un don (DEL ou PA). Mesure l'efficacité commerciale.", "Selon association"),
]

hy = Inches(1.68)
# en-têtes
for txt, ww, cx in [("SIGLE", Inches(1.0), Inches(.35)),
                     ("NOM COMPLET", Inches(3.2), Inches(1.4)),
                     ("DÉFINITION", Inches(6.5), Inches(4.65)),
                     ("VALEUR / OBJ.", Inches(1.9), Inches(11.2))]:
    rect(sl, cx, hy, ww, Inches(.38), fill=C_BLUE_DARK)
    txbox(sl, cx + Inches(.08), hy, ww - Inches(.16), Inches(.38),
          txt, size=9, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)

ry = hy + Inches(.4)
for i, (sigle, nom, defin, val) in enumerate(terms1):
    bg = C_BLUE_LIGHT if i % 2 == 0 else C_WHITE
    rh = Inches(.62)
    rect(sl, Inches(.35), ry, Inches(12.6), rh, fill=bg, line=C_GREY_LINE, line_w=Pt(.5))
    txbox(sl, Inches(.4), ry + Inches(.08), Inches(.9), rh - Inches(.16),
          sigle, size=13, bold=True, color=C_BLUE_MID)
    txbox(sl, Inches(1.42), ry + Inches(.08), Inches(3.1), rh - Inches(.16),
          nom, size=10, bold=True, color=C_BLUE_DARK)
    txbox(sl, Inches(4.67), ry + Inches(.06), Inches(6.4), rh - Inches(.12),
          defin, size=10, color=C_TEXT_LIGHT, wrap=True)
    txbox(sl, Inches(11.22), ry + Inches(.08), Inches(1.65), rh - Inches(.16),
          val, size=10, bold=True, color=C_ORANGE, align=PP_ALIGN.CENTER)
    ry += rh + Inches(.04)

footer_bar(sl)


# ════════════════════════════════════════════════════════════════════
#  SLIDE 6 — SECTION 2 : NOMENCLATURE (page 2/2)
# ════════════════════════════════════════════════════════════════════

sl = add_slide()
rect(sl, 0, 0, W, H, fill=C_GREY)
section_header(sl, "02", "Nomenclature — Définitions Officielles (2/2)",
               "Qualifications, méthodes et modes de prélèvement")

terms2 = [
    ("DEL",     "Dossier En Ligne",
     "Don validé par prélèvement automatique, finalisé en ligne. Synonyme opérationnel de PA.", "Compte dans TX"),
    ("PA",      "Prélèvement Automatique",
     "Mode de don régulier. Sur cette campagne, SEUL le PA régulier est proposé — jamais de don ponctuel.", "Mode unique"),
    ("IBAN",    "Coordonnées bancaires",
     "Collectées lors d'un prélèvement à chaud pour valider le PA immédiatement pendant l'appel.", "Mode A — À chaud"),
    ("ERR",     "Écouter → Reformuler → Rebondir",
     "Méthode en 3 temps pour traiter les objections. Ne jamais interrompre, toujours reformuler avant de répondre.", "Méthode officielle"),
    ("DON",     "Qualification : Don",
     "CU dont l'issue est un PA confirmé (à chaud ou promesse validée). Entre dans le calcul du TX.", "Compte dans TX"),
    ("REF ARG.","Qualification : Refus Argumenté",
     "CU conclu par refus mais après présentation d'au moins un contre-argument par l'agent.", "—"),
    ("INDÉCIS", "Qualification : Indécis",
     "CU conclu par hésitation. Un lien ou rappel peut être planifié. Potentiel de transformation futur.", "Suivi à prévoir"),
]

hy = Inches(1.68)
for txt, ww, cx in [("SIGLE", Inches(1.1), Inches(.35)),
                     ("NOM COMPLET", Inches(3.0), Inches(1.5)),
                     ("DÉFINITION", Inches(6.5), Inches(4.55)),
                     ("VALEUR / OBJ.", Inches(1.9), Inches(11.1))]:
    rect(sl, cx, hy, ww, Inches(.34), fill=C_BLUE_DARK)
    txbox(sl, cx + Inches(.06), hy, ww - Inches(.12), Inches(.34),
          txt, size=9, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)

ry = hy + Inches(.36)
for i, (sigle, nom, defin, val) in enumerate(terms2):
    bg = C_BLUE_LIGHT if i % 2 == 0 else C_WHITE
    rh = Inches(.56)
    rect(sl, Inches(.35), ry, Inches(12.6), rh, fill=bg, line=C_GREY_LINE, line_w=Pt(.5))
    txbox(sl, Inches(.4), ry + Inches(.06), Inches(1.0), rh - Inches(.12),
          sigle, size=12, bold=True, color=C_BLUE_MID)
    txbox(sl, Inches(1.52), ry + Inches(.06), Inches(2.9), rh - Inches(.12),
          nom, size=10, bold=True, color=C_BLUE_DARK)
    txbox(sl, Inches(4.57), ry + Inches(.04), Inches(6.4), rh - Inches(.08),
          defin, size=10, color=C_TEXT_LIGHT, wrap=True)
    txbox(sl, Inches(11.12), ry + Inches(.06), Inches(1.7), rh - Inches(.12),
          val, size=10, bold=True, color=C_ORANGE, align=PP_ALIGN.CENTER)
    ry += rh + Inches(.03)

footer_bar(sl)


# ════════════════════════════════════════════════════════════════════
#  SLIDE 7 — SECTION 3 : KPI — CU/H
# ════════════════════════════════════════════════════════════════════

sl = add_slide()
rect(sl, 0, 0, W, H, fill=C_GREY)
section_header(sl, "03", "KPI 1 — Contacts Utiles par Heure (CU/H)",
               "Mesure de la productivité horaire de l'agent")

# KPI card grand format
rect(sl, Inches(.35), Inches(1.72), Inches(3.5), Inches(2.5), fill=C_BLUE_DARK)
txbox(sl, Inches(.45), Inches(2.0), Inches(3.3), Inches(.5),
      "CU / H", size=18, bold=True, color=C_CYAN, align=PP_ALIGN.CENTER)
txbox(sl, Inches(.45), Inches(2.55), Inches(3.3), Inches(.9),
      "9", size=58, bold=True, color=C_CYAN, align=PP_ALIGN.CENTER)
txbox(sl, Inches(.45), Inches(3.5), Inches(3.3), Inches(.5),
      "OBJECTIF MINIMUM", size=11, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)

# Définition
rect(sl, Inches(4.05), Inches(1.72), Inches(8.9), Inches(1.1), fill=C_BLUE_LIGHT)
txbox(sl, Inches(4.25), Inches(1.78), Inches(8.5), Inches(.38),
      "Définition du Contact Utile (CU)", size=12, bold=True, color=C_BLUE_DARK)
txbox(sl, Inches(4.25), Inches(2.12), Inches(8.5), Inches(.5),
      "Appel dont la durée de communication dépasse strictement 60 secondes. En-dessous de ce seuil, l'appel n'est pas comptabilisé comme CU.",
      size=11, color=C_TEXT_LIGHT, wrap=True)

# Formule
formula_box(sl, Inches(4.05), Inches(2.95), Inches(8.9), Inches(2.0),
            "FORMULE MATHÉMATIQUE — CU/H",
            [
                {"text": "CU/H  =  Nombre de CU réalisés  ÷  Heures de production",
                 "color": RGBColor(0xE8, 0xF4, 0xFF), "size": 13, "bold": True},
            ])

# Exemple
formula_box(sl, Inches(.35), Inches(4.35), Inches(12.6), Inches(1.68),
            "EXEMPLE NUMÉRIQUE — Agent Philippe",
            [
                {"text": "CU réalisés     =  68",
                 "color": RGBColor(0x7D, 0xD3, 0xFC), "size": 12},
                {"text": "Heures produites=   7 h",
                 "color": RGBColor(0x7D, 0xD3, 0xFC), "size": 12},
                {"text": "CU/H            =  68  ÷  7  =  9,71  ✅  Objectif dépassé (≥ 9)",
                 "color": RGBColor(0x34, 0xD3, 0x99), "size": 13, "bold": True},
            ],
            note="⚠  Heure de production = temps effectivement passé en ligne (hors pauses et formation)")

footer_bar(sl)


# ════════════════════════════════════════════════════════════════════
#  SLIDE 8 — SECTION 3 : KPI — TAUX DE TRANSFORMATION
# ════════════════════════════════════════════════════════════════════

sl = add_slide()
rect(sl, 0, 0, W, H, fill=C_GREY)
section_header(sl, "03", "KPI 2 — Taux de Transformation",
               "Mesure de l'efficacité commerciale de l'agent (% de dons parmi les CU)")

# KPI card
rect(sl, Inches(.35), Inches(1.72), Inches(3.5), Inches(2.5), fill=RGBColor(0x1A, 0x3A, 0x6B))
txbox(sl, Inches(.45), Inches(2.0), Inches(3.3), Inches(.5),
      "TX TRANSFORMATION", size=13, bold=True, color=C_CYAN, align=PP_ALIGN.CENTER)
txbox(sl, Inches(.45), Inches(2.55), Inches(3.3), Inches(.9),
      "%", size=58, bold=True, color=C_CYAN, align=PP_ALIGN.CENTER)
txbox(sl, Inches(.45), Inches(3.5), Inches(3.3), Inches(.5),
      "CIBLE SELON ASSOCIATION", size=10, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)

# Formule principale
formula_box(sl, Inches(4.05), Inches(1.72), Inches(8.9), Inches(2.55),
            "FORMULE MATHÉMATIQUE — TAUX DE TRANSFORMATION",
            [
                {"text": "TX (%)  =  ( Nombre de DEL ou PA  ×  100 )  ÷  Nombre de CU",
                 "color": RGBColor(0xE8, 0xF4, 0xFF), "size": 13, "bold": True},
                {"text": " ", "color": C_WHITE, "size": 8},
                {"text": "   Numérateur   :  DEL / PA confirmés uniquement",
                 "color": RGBColor(0x94, 0xA3, 0xB8), "size": 10},
                {"text": "   Dénominateur :  Tous les CU (> 60 sec.)",
                 "color": RGBColor(0x94, 0xA3, 0xB8), "size": 10},
            ])

# Exemple
formula_box(sl, Inches(.35), Inches(4.4), Inches(12.6), Inches(1.75),
            "EXEMPLE NUMÉRIQUE — Agent Michel",
            [
                {"text": "CU réalisés      =  72",
                 "color": RGBColor(0x7D, 0xD3, 0xFC), "size": 12},
                {"text": "DEL (PA) obtenus =   3",
                 "color": RGBColor(0x7D, 0xD3, 0xFC), "size": 12},
                {"text": "TX  =  ( 3  ×  100 )  ÷  72  =  300  ÷  72  =  4,16 %",
                 "color": RGBColor(0x34, 0xD3, 0x99), "size": 13, "bold": True},
            ],
            note="ℹ  Une promesse non finalisée (lien non cliqué) ne compte pas encore dans le numérateur.")

footer_bar(sl)


# ════════════════════════════════════════════════════════════════════
#  SLIDE 9 — SECTION 3 : KPI — PDC
# ════════════════════════════════════════════════════════════════════

sl = add_slide()
rect(sl, 0, 0, W, H, fill=C_GREY)
section_header(sl, "03", "KPI 3 — Plan de Charge (PDC)",
               "Volume mensuel de CU à produire — pilotage collectif")

rect(sl, Inches(.35), Inches(1.72), Inches(3.5), Inches(2.5), fill=C_PURPLE)
txbox(sl, Inches(.45), Inches(2.0), Inches(3.3), Inches(.5),
      "PDC", size=22, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)
txbox(sl, Inches(.45), Inches(2.5), Inches(3.3), Inches(.9),
      "XXXX", size=36, bold=True, color=RGBColor(0xD8, 0xB4, 0xFE), align=PP_ALIGN.CENTER)
txbox(sl, Inches(.45), Inches(3.5), Inches(3.3), Inches(.5),
      "CU / MOIS  (valeur campagne)", size=10, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)

formula_box(sl, Inches(4.05), Inches(1.72), Inches(8.9), Inches(2.55),
            "FORMULE DE VÉRIFICATION — PDC",
            [
                {"text": "PDC atteint  =  CU/H moyen équipe  ×  Heures produites totales",
                 "color": RGBColor(0xE8, 0xF4, 0xFF), "size": 13, "bold": True},
                {"text": " ", "color": C_WHITE, "size": 8},
                {"text": "CU/H requis  =  PDC mensuel cible  ÷  Heures totales équipe",
                 "color": RGBColor(0xC4, 0xB5, 0xFD), "size": 11},
            ])

formula_box(sl, Inches(.35), Inches(4.4), Inches(12.6), Inches(1.5),
            "EXEMPLE — PDC cible 1 800 CU / équipe 200 h de production",
            [
                {"text": "CU/H requis  =  1 800  ÷  200  =  9 CU/H  =  objectif fixé ✅",
                 "color": RGBColor(0x34, 0xD3, 0x99), "size": 13, "bold": True},
            ],
            note="Le PDC exact de la campagne sera communiqué par votre superviseur en début de mission.")

footer_bar(sl)


# ════════════════════════════════════════════════════════════════════
#  SLIDE 10 — SECTION 3 : INTERPRÉTATION KPI
# ════════════════════════════════════════════════════════════════════

sl = add_slide()
rect(sl, 0, 0, W, H, fill=C_GREY)
section_header(sl, "03", "Interpréter ses KPI — Grille de Lecture",
               "Comment réagir selon son niveau de performance")

rows = [
    (C_GREEN,  C_GREEN_LIGHT,  "✅  CU/H ≥ 9",      "Performance conforme",
     "Maintenir le rythme. Travailler le TX transformation pour maximiser les dons."),
    (C_YELLOW, C_YELLOW_LIGHT, "⚠  CU/H 7 – 9",     "Zone d'alerte",
     "Identifier les freins : durée d'appel trop longue, trop de silences, objections non traitées."),
    (C_RED,    C_RED_LIGHT,    "✗  CU/H < 7",        "En-dessous du seuil",
     "Plan d'action individuel obligatoire. Écoute d'appels avec superviseur. Retour formation."),
    (C_BLUE_MID, C_BLUE_LIGHT, "📈  TX élevé",       "Excellente efficacité",
     "Partager les techniques d'argumentation avec l'équipe. Bonus de performance possible."),
    (C_ORANGE, RGBColor(0xFF,0xF4,0xE6), "⏳  TX faible + INDÉCIS élevé",
     "Argumentation insuffisante",
     "Retravailler les étapes 4 et 5 de l'appel : ancrage montant et traitement des objections."),
]

ry = Inches(1.72)
for border, bg, kpi, label, action in rows:
    rh = Inches(.9)
    rect(sl, Inches(.35), ry, Inches(12.6), rh, fill=bg, line=border, line_w=Pt(1.5))
    rect(sl, Inches(.35), ry, Inches(.08), rh, fill=border)
    txbox(sl, Inches(.55), ry + Inches(.1), Inches(2.5), Inches(.35),
          kpi, size=12, bold=True, color=border)
    txbox(sl, Inches(3.1), ry + Inches(.1), Inches(2.0), Inches(.35),
          label, size=11, bold=True, color=RGBColor(0x1A,0x1A,0x2E))
    txbox(sl, Inches(5.2), ry + Inches(.1), Inches(7.5), Inches(.65),
          f"→  {action}", size=11, color=C_TEXT_LIGHT, wrap=True)
    ry += rh + Inches(.08)

footer_bar(sl)


# ════════════════════════════════════════════════════════════════════
#  SLIDE 11 — SECTION 4 : LES 3 QUALIFICATIONS
# ════════════════════════════════════════════════════════════════════

sl = add_slide()
rect(sl, 0, 0, W, H, fill=C_GREY)
section_header(sl, "04", "Les 3 Qualifications de Contact Utile",
               "Tout CU doit être qualifié immédiatement après chaque appel")

qw = Inches(4.0)
qdata = [
    (C_GREEN, C_GREEN_LIGHT, "💚", "DON",
     "Le contact accepte le prélèvement automatique régulier.",
     ["PA confirmé à chaud (IBAN collecté)", "OU promesse par lien envoyé", "Entre dans le calcul du TX"],
     C_GREEN),
    (C_RED, C_RED_LIGHT, "❌", "REFUS ARGUMENTÉ",
     "Le contact refuse après avoir entendu l'argumentaire.",
     ["L'agent a présenté ≥ 1 contre-argument", "Ne pas qualifier sans avoir vraiment argumenté", "Ne compte pas dans le TX"],
     C_RED),
    (C_ORANGE, RGBColor(0xFF,0xF4,0xE6), "⏳", "INDÉCIS",
     "Le contact hésite ou souhaite réfléchir.",
     ["Un lien peut être envoyé", "Un rappel peut être planifié", "Potentiel de transformation futur"],
     C_ORANGE),
]

cx = Inches(.35)
for border, bg, icon, name, desc, pts, nc in qdata:
    rect(sl, cx, Inches(1.72), qw, Inches(4.8), fill=bg, line=border, line_w=Pt(2))
    rect(sl, cx, Inches(1.72), qw, Inches(.06), fill=border)
    txbox(sl, cx, Inches(1.82), qw, Inches(.7), icon, size=32, align=PP_ALIGN.CENTER, color=border)
    txbox(sl, cx, Inches(2.5), qw, Inches(.45),
          name, size=14, bold=True, color=border, align=PP_ALIGN.CENTER)
    txbox(sl, cx + Inches(.2), Inches(2.98), qw - Inches(.4), Inches(.6),
          desc, size=11, color=C_TEXT_LIGHT, align=PP_ALIGN.CENTER, wrap=True)
    py = Inches(3.62)
    for pt in pts:
        txbox(sl, cx + Inches(.2), py, qw - Inches(.4), Inches(.36),
              f"✓  {pt}", size=10, color=RGBColor(0x1A,0x1A,0x2E))
        py += Inches(.38)
    cx += qw + Inches(.17)

rect(sl, Inches(.35), Inches(6.65), Inches(12.6), Inches(.5), fill=C_YELLOW_LIGHT,
     line=C_YELLOW, line_w=Pt(1.5))
txbox(sl, Inches(.55), Inches(6.68), Inches(12.2), Inches(.44),
      "RÈGLE : Qualifier DON uniquement si un PA est effectivement initié. Tout doute = INDÉCIS.",
      size=11, bold=True, color=RGBColor(0x78, 0x35, 0x00))

footer_bar(sl)


# ════════════════════════════════════════════════════════════════════
#  SLIDE 12 — SECTION 5 : LE PRODUIT
# ════════════════════════════════════════════════════════════════════

sl = add_slide()
rect(sl, 0, 0, W, H, fill=C_GREY)
section_header(sl, "05", "Le Produit — Prélèvement Automatique Régulier",
               "Position officielle SHY Performance sur cette campagne")

# Règle d'or
rect(sl, Inches(.35), Inches(1.72), Inches(12.6), Inches(.9), fill=C_BLUE_MID)
txbox(sl, Inches(.55), Inches(1.8), Inches(12.2), Inches(.74),
      "RÈGLE D'OR :  Sur cette campagne, SEUL le prélèvement automatique régulier est proposé.  Aucun don ponctuel.",
      size=14, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)

# Mode A
rect(sl, Inches(.35), Inches(2.78), Inches(6.0), Inches(2.45), fill=C_GREEN_LIGHT,
     line=C_GREEN, line_w=Pt(2))
rect(sl, Inches(.35), Inches(2.78), Inches(6.0), Inches(.06), fill=C_GREEN)
txbox(sl, Inches(.45), Inches(2.88), Inches(5.8), Inches(.38),
      "MODE A — Prélèvement À CHAUD (IBAN)", size=13, bold=True, color=C_GREEN)
txbox(sl, Inches(.45), Inches(3.28), Inches(5.8), Inches(1.85),
      ("L'agent recueille l'IBAN du donateur pendant l'appel\n"
       "et valide lui-même le prélèvement en ligne.\n\n"
       "✓  Don sécurisé immédiatement\n"
       "✓  Taux de concrétisation maximal\n"
       "✓  Mode PRIORITAIRE"),
      size=11, color=C_TEXT_LIGHT, wrap=True)

# Mode B
rect(sl, Inches(6.7), Inches(2.78), Inches(6.25), Inches(2.45),
     fill=RGBColor(0xFF,0xF4,0xE6), line=C_ORANGE, line_w=Pt(2))
rect(sl, Inches(6.7), Inches(2.78), Inches(6.25), Inches(.06), fill=C_ORANGE)
txbox(sl, Inches(6.9), Inches(2.88), Inches(6.0), Inches(.38),
      "MODE B — Promesse par LIEN", size=13, bold=True, color=C_ORANGE)
txbox(sl, Inches(6.9), Inches(3.28), Inches(6.0), Inches(1.85),
      ("L'agent envoie un lien sécurisé au donateur.\n"
       "Ce dernier finalise seul après l'appel.\n\n"
       "✓  Lien PERMANENT — aucune date d'expiration\n"
       "✓  Redirige vers le site officiel de l'association\n"
       "✓  E-mail de confirmation envoyé automatiquement"),
      size=11, color=C_TEXT_LIGHT, wrap=True)

# À ne jamais dire
stop_box(sl, Inches(.35), Inches(5.38), Inches(12.6), Inches(1.12),
         "À NE JAMAIS DIRE OU FAIRE sur cette campagne",
         ["Proposer un don ponctuel — ce n'est pas l'offre de la campagne",
          "Indiquer une date d'expiration sur le lien — il est permanent",
          "Créer une pression artificielle ou culpabiliser le donateur"],
         title_size=12, item_size=11)

footer_bar(sl)


# ════════════════════════════════════════════════════════════════════
#  SLIDE 13 — SECTION 6 : STRUCTURE D'UN APPEL (étapes 1–4)
# ════════════════════════════════════════════════════════════════════

sl = add_slide()
rect(sl, 0, 0, W, H, fill=C_GREY)
section_header(sl, "06", "Structure d'un Appel Réussi — Étapes 1 à 4",
               "Progression logique et irréversible — ne jamais brûler une étape")

steps_a = [
    ("1", "ACCROCHE & IDENTIFICATION",
     "Se présenter clairement, nommer l'association, vérifier l'identité du contact.",
     '"Bonjour, je suis [Prénom], j\'appelle de la part de [l\'Association]. Ai-je bien M./Mme [Nom] ?"'),
    ("2", "CONTEXTE & LÉGITIMITÉ",
     "Expliquer pourquoi ce contact est appelé. Créer un lien émotionnel avec la cause.",
     '"Nous appelons des personnes qui, comme vous, ont témoigné leur sensibilité à [la cause]…"'),
    ("3", "PRÉSENTATION DE L'IMPACT",
     "Chiffrer l'impact d'un don mensuel. Rendre tangible chaque euro.",
     '"Un prélèvement de [X]€/mois, c\'est [impact précis] — moins de [Y]€/jour pour changer une vie."'),
    ("4", "PROPOSITION & ANCRAGE MONTANT",
     "Proposer un montant de référence (ancrage haut). Laisser le contact négocier à la baisse.",
     '"Beaucoup de nos donateurs commencent à [X]€/mois. Est-ce une somme accessible pour vous ?"'),
]

ry = Inches(1.72)
for num, title, desc, ex in steps_a:
    rh = Inches(1.22)
    rect(sl, Inches(.35), ry, Inches(12.6), rh, fill=C_WHITE,
         line=C_BLUE_LIGHT, line_w=Pt(1))
    rect(sl, Inches(.35), ry, Inches(.55), rh, fill=C_BLUE_MID)
    txbox(sl, Inches(.35), ry, Inches(.55), rh,
          num, size=20, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)
    txbox(sl, Inches(1.0), ry + Inches(.08), Inches(11.7), Inches(.36),
          title, size=12, bold=True, color=C_BLUE_DARK)
    txbox(sl, Inches(1.0), ry + Inches(.44), Inches(5.5), Inches(.56),
          desc, size=10, color=C_TEXT_LIGHT, wrap=True)
    rect(sl, Inches(6.6), ry + Inches(.1), Inches(5.95), rh - Inches(.2),
         fill=C_BLUE_LIGHT)
    txbox(sl, Inches(6.72), ry + Inches(.12), Inches(5.7), rh - Inches(.24),
          ex, size=10, italic=True, color=C_BLUE_DARK, wrap=True)
    ry += rh + Inches(.06)

footer_bar(sl)


# ════════════════════════════════════════════════════════════════════
#  SLIDE 14 — SECTION 6 : STRUCTURE D'UN APPEL (étapes 5–7)
# ════════════════════════════════════════════════════════════════════

sl = add_slide()
rect(sl, 0, 0, W, H, fill=C_GREY)
section_header(sl, "06", "Structure d'un Appel Réussi — Étapes 5 à 7",
               "Closing, sécurisation du don et qualification dans l'outil")

steps_b = [
    ("5", "TRAITEMENT DES OBJECTIONS",
     "Méthode ERR : Écouter → Reformuler → Rebondir. Ne jamais abandonner à la 1re résistance.",
     "Voir Section 07 — Tableau complet des 7 objections clés et réponses recommandées."),
    ("6", "CLOSING & SÉCURISATION DU DON",
     "Confirmer l'accord, collecter l'IBAN (mode à chaud) ou envoyer le lien. Récapituler montant et fréquence.",
     '"Parfait ! Je note votre IBAN pour finaliser le prélèvement de [X]€/mois. Vous recevrez un e-mail. Merci !"'),
    ("7", "QUALIFICATION & SAISIE",
     "Immédiatement après l'appel : qualifier en DON / REFUS ARGUMENTÉ / INDÉCIS dans l'outil.",
     "Une saisie rapide et précise garantit des statistiques fiables et un suivi de qualité pour l'équipe."),
]

ry = Inches(1.72)
for num, title, desc, ex in steps_b:
    rh = Inches(1.45)
    rect(sl, Inches(.35), ry, Inches(12.6), rh, fill=C_WHITE,
         line=C_BLUE_LIGHT, line_w=Pt(1))
    rect(sl, Inches(.35), ry, Inches(.55), rh, fill=C_BLUE_MID)
    txbox(sl, Inches(.35), ry, Inches(.55), rh,
          num, size=20, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)
    txbox(sl, Inches(1.0), ry + Inches(.1), Inches(11.7), Inches(.4),
          title, size=12, bold=True, color=C_BLUE_DARK)
    txbox(sl, Inches(1.0), ry + Inches(.52), Inches(5.5), Inches(.78),
          desc, size=10, color=C_TEXT_LIGHT, wrap=True)
    rect(sl, Inches(6.6), ry + Inches(.1), Inches(5.95), rh - Inches(.2),
         fill=C_BLUE_LIGHT)
    txbox(sl, Inches(6.72), ry + Inches(.15), Inches(5.7), rh - Inches(.3),
          ex, size=10, italic=True, color=C_BLUE_DARK, wrap=True)
    ry += rh + Inches(.12)

# Règle 60 sec
rect(sl, Inches(.35), H - Inches(1.0), Inches(12.6), Inches(.48), fill=C_ORANGE)
txbox(sl, Inches(.55), H - Inches(1.0), Inches(12.2), Inches(.44),
      "OBJECTIF DE CHAQUE APPEL : dépasser les 60 secondes — c'est le seuil qui transforme l'appel en CU.",
      size=12, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)

footer_bar(sl)


# ════════════════════════════════════════════════════════════════════
#  SLIDE 15 — SECTION 7 : OBJECTIONS (1/2)
# ════════════════════════════════════════════════════════════════════

sl = add_slide()
rect(sl, 0, 0, W, H, fill=C_GREY)
section_header(sl, "07", "Traitement des Objections — Méthode ERR (1/2)",
               "Écouter → Reformuler → Rebondir  |  Une objection = une opportunité de dialogue")

rect(sl, Inches(.35), Inches(1.72), Inches(12.6), Inches(.6), fill=C_PURPLE_LIGHT,
     line=C_PURPLE, line_w=Pt(1.5))
txbox(sl, Inches(.55), Inches(1.76), Inches(12.2), Inches(.52),
      "PRINCIPE : Une objection n'est pas un refus. C'est une demande d'information déguisée. Maximum 2 à 3 tentatives avant de qualifier REFUS ARGUMENTÉ.",
      size=11, bold=True, color=C_PURPLE)

objections_a = [
    ('"Je n\'ai pas les moyens."',
     "Peur d'un engagement trop lourd",
     '"Je comprends. C\'est pourquoi nous permettons de commencer avec un montant symbolique comme [X]€/mois — moins de [Y]€/jour. Même un geste modeste change des vies concrètement."'),
    ('"Je vais y réfléchir."',
     "Manque de conviction ou peur de décider",
     '"Absolument. Puis-je vous envoyer un lien sécurisé maintenant ? Vous finalisez à votre rythme — le lien n\'expire jamais. Je vous l\'envoie par email ?"'),
    ('"Je donne déjà à d\'autres associations."',
     "Sentiment de saturation",
     '"C\'est formidable. Les besoins humanitaires sont immenses. Un soutien complémentaire, même modeste, peut faire une vraie différence dans des zones que d\'autres n\'atteignent pas."'),
    ('"Je ne fais pas confiance aux associations."',
     "Scepticisme sur l'usage des fonds",
     '"C\'est légitime. [L\'Association] publie chaque année un rapport financier certifié. Une grande part des dons va directement aux programmes terrain. Je vous donne des exemples ?"'),
]

ry = Inches(2.48)
for objec, sens, rep in objections_a:
    rh = Inches(.88)
    rect(sl, Inches(.35), ry, Inches(12.6), rh, fill=C_WHITE,
         line=C_GREY_LINE, line_w=Pt(.5))
    rect(sl, Inches(.35), ry, Inches(.06), rh, fill=C_RED)
    txbox(sl, Inches(.48), ry + Inches(.08), Inches(2.8), Inches(.36),
          objec, size=11, bold=True, color=C_RED)
    txbox(sl, Inches(.48), ry + Inches(.46), Inches(2.8), Inches(.35),
          f"→ {sens}", size=9, italic=True, color=C_TEXT_LIGHT)
    rect(sl, Inches(3.4), ry + Inches(.08), Inches(.02), rh - Inches(.16), fill=C_GREY_LINE)
    txbox(sl, Inches(3.52), ry + Inches(.08), Inches(9.3), rh - Inches(.16),
          rep, size=10, color=C_BLUE_DARK, italic=True, wrap=True)
    ry += rh + Inches(.05)

footer_bar(sl)


# ════════════════════════════════════════════════════════════════════
#  SLIDE 16 — SECTION 7 : OBJECTIONS (2/2)
# ════════════════════════════════════════════════════════════════════

sl = add_slide()
rect(sl, 0, 0, W, H, fill=C_GREY)
section_header(sl, "07", "Traitement des Objections — Méthode ERR (2/2)",
               "Suite des objections clés + clôture d'un refus définitif")

objections_b = [
    ('"Je ne veux pas donner mon IBAN."',
     "Peur de la fraude",
     '"Pas de problème. Je vous envoie un lien sécurisé sur le site officiel — c\'est vous qui saisissez directement, sans intermédiaire. Votre sécurité est notre priorité absolue."'),
    ('"Ce n\'est pas le bon moment."',
     "Envie de raccrocher",
     '"Je comprends. Pour les enfants dans les zones de crise, malheureusement le temps est précieux. Le lien n\'a pas de date limite — vous agissez quand vous vous sentez prêt(e). Votre email ?"'),
    ('"Envoyez-moi des informations par courrier."',
     "Stratégie pour terminer l'appel",
     '"Bien sûr. Cela dit, le plus simple reste le lien email que je peux vous envoyer maintenant. En 2 minutes, toutes les informations sont sur votre écran. Votre adresse email ?"'),
]

ry = Inches(1.72)
for objec, sens, rep in objections_b:
    rh = Inches(.95)
    rect(sl, Inches(.35), ry, Inches(12.6), rh, fill=C_WHITE,
         line=C_GREY_LINE, line_w=Pt(.5))
    rect(sl, Inches(.35), ry, Inches(.06), rh, fill=C_RED)
    txbox(sl, Inches(.48), ry + Inches(.1), Inches(2.8), Inches(.36),
          objec, size=11, bold=True, color=C_RED)
    txbox(sl, Inches(.48), ry + Inches(.5), Inches(2.8), Inches(.36),
          f"→ {sens}", size=9, italic=True, color=C_TEXT_LIGHT)
    rect(sl, Inches(3.4), ry + Inches(.1), Inches(.02), rh - Inches(.2), fill=C_GREY_LINE)
    txbox(sl, Inches(3.52), ry + Inches(.1), Inches(9.3), rh - Inches(.2),
          rep, size=10, color=C_BLUE_DARK, italic=True, wrap=True)
    ry += rh + Inches(.08)

# Clôture refus
rect(sl, Inches(.35), ry + Inches(.1), Inches(12.6), Inches(1.55), fill=C_RED_LIGHT,
     line=C_RED, line_w=Pt(2))
txbox(sl, Inches(.55), ry + Inches(.2), Inches(12.2), Inches(.38),
      "CLORE AVEC DIGNITÉ — Le refus définitif", size=13, bold=True, color=C_RED)
txbox(sl, Inches(.55), ry + Inches(.6), Inches(12.2), Inches(.95),
      ("Si après 2 à 3 argumentations le contact maintient son refus : ne pas insister. "
       "Remercier chaleureusement et conclure positivement.\n"
       '"Merci pour votre temps. N\'hésitez pas à nous contacter si votre situation évolue. Bonne journée !"  →  Qualifier : REFUS ARGUMENTÉ'),
      size=11, color=C_TEXT_LIGHT, wrap=True)

footer_bar(sl)


# ════════════════════════════════════════════════════════════════════
#  SLIDE 17 — SECTION 8 : POSTURE & COMMUNICATION
# ════════════════════════════════════════════════════════════════════

sl = add_slide()
rect(sl, 0, 0, W, H, fill=C_GREY)
section_header(sl, "08", "Posture & Fondamentaux de Communication",
               "Le triangle Émotion / Raison / Facilité — la base de l'argumentation efficace")

# Triangle
rect(sl, Inches(.35), Inches(1.72), Inches(3.9), Inches(4.85), fill=C_TEAL_LIGHT,
     line=C_TEAL, line_w=Pt(2))
txbox(sl, Inches(.45), Inches(1.82), Inches(3.7), Inches(.38),
      "LE TRIANGLE DE L'ARGUMENTATION", size=10, bold=True, color=C_TEAL)
for icon, lbl, desc, cy in [
    ("❤️", "ÉMOTION", "Toucher le cœur\n(enfants, vies sauvées,\nurgence humanitaire)", Inches(2.28)),
    ("🧠", "RAISON",  "Rassurer l'intellect\n(transparence des fonds,\nimpact chiffré, label)", Inches(3.52)),
    ("✋", "FACILITÉ","Lever les freins\n(lien simple, IBAN\nsécurisé, pas de limite)", Inches(4.76)),
]:
    rect(sl, Inches(.5), cy, Inches(3.6), Inches(.95), fill=C_WHITE, line=C_TEAL, line_w=Pt(1))
    txbox(sl, Inches(.55), cy + Inches(.06), Inches(.6), Inches(.38), icon, size=16)
    txbox(sl, Inches(1.15), cy + Inches(.05), Inches(.85), Inches(.3),
          lbl, size=10, bold=True, color=C_TEAL)
    txbox(sl, Inches(1.15), cy + Inches(.32), Inches(2.8), Inches(.58),
          desc, size=9, color=C_TEXT_LIGHT, wrap=True)

# À FAIRE
bullet_box(sl, Inches(4.4), Inches(1.72), Inches(4.15), Inches(4.85),
           "À FAIRE SYSTÉMATIQUEMENT",
           ["Adopter un sourire vocal — il s'entend",
            "Parler à un rythme posé et articulé",
            "Écoute active — ne jamais interrompre",
            "Reformuler les objections avant de répondre",
            "Personnaliser en utilisant le nom du donateur",
            "Rester positif face à l'agressivité",
            "Dépasser les 60 secondes sur chaque appel",
            "Qualifier immédiatement après chaque appel"],
           bg=C_GREEN_LIGHT, border=C_GREEN,
           title_color=C_GREEN, item_color=C_TEXT_LIGHT,
           title_size=12, item_size=10)

# À ÉVITER
stop_box(sl, Inches(8.7), Inches(1.72), Inches(4.25), Inches(4.85),
         "À ÉVITER ABSOLUMENT",
         ["Raccrocher sans avoir argumenté",
          "Proposer un don ponctuel (hors offre)",
          "Donner une date limite au lien",
          "Promettre des résultats non garantis",
          "Interrompre le donateur",
          "Qualifier DON sans PA réel",
          "Oublier la saisie dans l'outil",
          "Créer une pression ou culpabilité excessive"],
         title_size=12, item_size=10)

footer_bar(sl)


# ════════════════════════════════════════════════════════════════════
#  SLIDE 18 — SECTION 9 : AUTO-SUIVI
# ════════════════════════════════════════════════════════════════════

sl = add_slide()
rect(sl, 0, 0, W, H, fill=C_GREY)
section_header(sl, "09", "Auto-Suivi de Performance — Tableau de Bord Horaire",
               "Un agent performant calcule ses KPI à chaque fin de tranche horaire")

# Rappel formules compact
formula_box(sl, Inches(.35), Inches(1.72), Inches(12.6), Inches(1.05),
            "RAPPEL FORMULES — À utiliser à chaque fin d'heure",
            [
                {"text":
                 "CU/H (%)  =  CU cumulés  ÷  Heures produites cumulées          "
                 "TX (%)  =  ( DON cumulés × 100 )  ÷  CU cumulés",
                 "color": RGBColor(0xE8,0xF4,0xFF), "size": 11, "bold": True},
            ])

# Tableau suivi
headers = ["Tranche", "CU réalisés", "H produites\n(cumulé)", "CU/H\ncumulé",
           "DON", "Refus Arg.", "Indécis", "TX Transfo"]
col_widths = [Inches(1.55), Inches(1.3), Inches(1.3), Inches(1.3),
              Inches(1.1), Inches(1.3), Inches(1.2), Inches(1.6)]
rows_data = [
    ["9h – 10h",  "_____", "1 h", "_____", "_____", "_____", "_____", "_____%"],
    ["10h – 11h", "_____", "2 h", "_____", "_____", "_____", "_____", "_____%"],
    ["11h – 12h", "_____", "3 h", "_____", "_____", "_____", "_____", "_____%"],
    ["14h – 15h", "_____", "4 h", "_____", "_____", "_____", "_____", "_____%"],
    ["15h – 16h", "_____", "5 h", "_____", "_____", "_____", "_____", "_____%"],
    ["16h – 17h", "_____", "6 h", "_____", "_____", "_____", "_____", "_____%"],
    ["TOTAL J.",  "_____", "__ h","_____", "_____", "_____", "_____", "_____%"],
]
tx = Inches(.35)
ty = Inches(2.9)
ch = Inches(.4)

# Header row
cx = tx
for i, (h, cw) in enumerate(zip(headers, col_widths)):
    rect(sl, cx, ty, cw, ch, fill=C_BLUE_DARK)
    txbox(sl, cx + Inches(.04), ty, cw - Inches(.08), ch,
          h, size=8, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)
    cx += cw

ry = ty + ch
for ri, row in enumerate(rows_data):
    is_total = ri == len(rows_data) - 1
    bg = C_BLUE_MID if is_total else (C_BLUE_LIGHT if ri % 2 == 0 else C_WHITE)
    cx = tx
    for ci, (cell, cw) in enumerate(zip(row, col_widths)):
        rect(sl, cx, ry, cw, ch, fill=bg, line=C_GREY_LINE, line_w=Pt(.3))
        col = C_WHITE if is_total else (C_BLUE_DARK if ci == 0 else C_TEXT_LIGHT)
        if is_total and ci == 0:
            col = C_CYAN
        txbox(sl, cx + Inches(.04), ry, cw - Inches(.08), ch,
              cell, size=9 if not is_total else 10, bold=is_total,
              color=col, align=PP_ALIGN.CENTER)
        cx += cw
    ry += ch

# Exemple
rect(sl, Inches(.35), ry + Inches(.12), Inches(12.6), Inches(.52), fill=C_GREEN_LIGHT,
     line=C_GREEN, line_w=Pt(1))
txbox(sl, Inches(.55), ry + Inches(.15), Inches(12.2), Inches(.44),
      "Exemple : après 3h de production, 28 CU dont 2 DON  →  CU/H = 28 ÷ 3 = 9,33 ✅  |  TX = (2 × 100) ÷ 28 = 7,14 %",
      size=11, bold=True, color=C_GREEN)

footer_bar(sl)


# ════════════════════════════════════════════════════════════════════
#  SLIDE 19 — SECTION 10 : CHECKLIST J1
# ════════════════════════════════════════════════════════════════════

sl = add_slide()
rect(sl, 0, 0, W, H, fill=C_GREY)
section_header(sl, "10", "Checklist J1 — Suis-je Prêt à Décrocher ?",
               "Valider chaque point avant de prendre son premier appel en autonomie")

checks = [
    ("Mission & Nomenclature", C_BLUE_MID, C_BLUE_LIGHT, [
        "Je connais la finalité humanitaire de la campagne",
        "Je sais définir : CU, CU/H, PDC, TX, DEL, PA",
        "Je connais les 3 qualifications : DON / REFUS ARG. / INDÉCIS",
        "Je maîtrise la méthode ERR",
    ]),
    ("KPI & Formules", C_ORANGE, RGBColor(0xFF,0xF4,0xE6), [
        "Je sais calculer mon CU/H en temps réel",
        "Mon objectif CU/H est 9 minimum",
        "Je sais calculer le Taux de Transformation",
        "Je connais la formule du PDC équipe",
    ]),
    ("Produit", C_PURPLE, C_PURPLE_LIGHT, [
        "Seuls les PA réguliers sont proposés (jamais ponctuel)",
        "Je connais les 2 modes : à chaud et promesse",
        "Le lien est permanent — pas de date d'expiration",
        "Un email de confirmation est envoyé automatiquement",
    ]),
    ("Appel & Argumentation", C_GREEN, C_GREEN_LIGHT, [
        "Je connais les 7 étapes de l'appel",
        "Je sais traiter au moins 4 objections principales",
        "Je sais conclure dignement un refus définitif",
        "Je sais suivre ma performance heure par heure",
    ]),
]

cx = Inches(.35)
for title, border, bg, pts in checks:
    cw = Inches(3.1)
    ch_box = Inches(4.85)
    rect(sl, cx, Inches(1.72), cw, ch_box, fill=bg, line=border, line_w=Pt(2))
    rect(sl, cx, Inches(1.72), cw, Inches(.06), fill=border)
    txbox(sl, cx + Inches(.12), Inches(1.82), cw - Inches(.24), Inches(.4),
          title, size=11, bold=True, color=border)
    py = Inches(2.28)
    for pt in pts:
        rect(sl, cx + Inches(.15), py, Inches(.28), Inches(.28), fill=border)
        txbox(sl, cx + Inches(.15), py, Inches(.28), Inches(.28),
              "✓", size=9, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)
        txbox(sl, cx + Inches(.5), py, cw - Inches(.65), Inches(.34),
              pt, size=10, color=RGBColor(0x1A,0x1A,0x2E))
        py += Inches(.45)
    cx += cw + Inches(.12)

footer_bar(sl)


# ════════════════════════════════════════════════════════════════════
#  SLIDE 20 — SLIDE DE CLÔTURE
# ════════════════════════════════════════════════════════════════════

sl = add_slide()
rect(sl, 0, 0, W, H, fill=C_BLUE_DARK)
rect(sl, 0, 0, W, Inches(.08), fill=C_ORANGE)
rect(sl, 0, H - Inches(.08), W, Inches(.08), fill=C_CYAN)
rect(sl, Inches(6.55), 0, Inches(.06), H, fill=C_CYAN)

# Gauche — mantra
txbox(sl, Inches(.5), Inches(1.6), Inches(5.8), Inches(.6),
      "POUR RETENIR", size=11, bold=True, color=C_ORANGE)
txbox(sl, Inches(.5), Inches(2.2), Inches(5.8), Inches(1.2),
      '"LE REFUS D\'AUJOURD\'HUI\nPEUT ÊTRE\nLE DON DE DEMAIN."',
      size=22, bold=True, color=C_WHITE, wrap=True)
txbox(sl, Inches(.5), Inches(3.6), Inches(5.8), Inches(1.8),
      ("Chaque appel est une opportunité de changer une vie.\n\n"
       "Votre voix, votre conviction, votre argument —\n"
       "c'est ce qui transforme une hésitation\nen espoir concret."),
      size=13, color=RGBColor(0xBB,0xCC,0xEE), italic=True, wrap=True)

# Droite — récap chiffres
txbox(sl, Inches(7.0), Inches(1.5), Inches(5.9), Inches(.45),
      "LES CHIFFRES À RETENIR", size=11, bold=True, color=C_CYAN)

kpis_final = [
    ("9",      "CU/H — objectif minimum"),
    ("> 60s",  "Durée seuil d'un CU"),
    ("3",      "Qualifications possibles"),
    ("7",      "Étapes d'un appel réussi"),
    ("ERR",    "Méthode anti-objection"),
    ("∞",      "Validité du lien de don"),
]
ky = Inches(2.05)
for val, lbl in kpis_final:
    rect(sl, Inches(7.0), ky, Inches(5.9), Inches(.62), fill=RGBColor(0x00,0x40,0xA0))
    txbox(sl, Inches(7.08), ky + Inches(.08), Inches(1.2), Inches(.46),
          val, size=20, bold=True, color=C_CYAN, align=PP_ALIGN.CENTER)
    txbox(sl, Inches(8.38), ky + Inches(.16), Inches(4.4), Inches(.32),
          lbl, size=12, color=C_WHITE)
    ky += Inches(.7)

footer_bar(sl, "Fidelis × SHY Performance  |  Module J1 FIM  |  Formation Initiale Module  |  Confidentiel")


# ════════════════════════════════════════════════════════════════════
#  SAUVEGARDE
# ════════════════════════════════════════════════════════════════════

output_path = "/home/user/fatou/J1_FIM_Fidelis_Formation.pptx"
prs.save(output_path)
print(f"Fichier enregistré : {output_path}")
print(f"Nombre de slides   : {len(prs.slides)}")

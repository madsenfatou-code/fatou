"""
SYNTHÈSE J2 MATINÉE — Réveil pédagogique
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

FOOTER_TXT = "SYNTHÈSE J2 Matinée  |  Réveil Pédagogique  |  Fidelis × SHY Performance"

def footer(slide, txt=FOOTER_TXT):
    rect(slide, 0, H - Inches(.4), W, Inches(.4), fill=C_BLUE_DARK)
    tb(slide, Inches(.3), H - Inches(.38), W - Inches(.6), Inches(.34),
       txt, size=9, color=RGBColor(0xAA, 0xBB, 0xDD), align=PP_ALIGN.CENTER)

def section_top(slide, num, title, subtitle, accent=C_ORANGE):
    rect(slide, 0, 0, W, Inches(1.35), fill=C_BLUE_DARK)
    rect(slide, 0, 0, W, Inches(.06), fill=accent)
    tb(slide, Inches(.4), Inches(.1), Inches(12.5), Inches(.58),
       f"{num}  —  {title}", size=22, bold=True, color=C_WHITE)
    tb(slide, Inches(.4), Inches(.72), Inches(12.5), Inches(.44),
       subtitle, size=12, italic=True, color=C_CYAN)

def qr_bloc(slide, x, y, w, h, question, reponse, q_color=C_ORANGE, r_color=C_GREEN):
    half = h / 2 - Inches(.03)
    rect(slide, x, y, w, half, fill=RGBColor(0xFF, 0xF4, 0xE6), line=q_color, line_w=Pt(1.5))
    tb(slide, x + Inches(.12), y + Inches(.05), w - Inches(.24), Inches(.2),
       "❓  QUESTION", size=8, bold=True, color=q_color)
    tb(slide, x + Inches(.12), y + Inches(.27), w - Inches(.24), half - Inches(.32),
       question, size=11, bold=True, color=C_BLUE_DARK, wrap=True)
    rect(slide, x, y + half + Inches(.06), w, half, fill=C_GREEN_LIGHT, line=r_color, line_w=Pt(1.5))
    tb(slide, x + Inches(.12), y + half + Inches(.1), w - Inches(.24), Inches(.2),
       "✅  RÉPONSE", size=8, bold=True, color=r_color)
    tb(slide, x + Inches(.12), y + half + Inches(.32), w - Inches(.24), half - Inches(.38),
       reponse, size=11, color=C_BLUE_DARK, wrap=True)


# ════════════════════════════════════════════════════════════════════
# SLIDE 1 — COUVERTURE
# ════════════════════════════════════════════════════════════════════
sl = add_slide()
rect(sl, 0, 0, W, H, fill=C_BLUE_DARK)
rect(sl, 0, 0, W, Inches(.08), fill=C_ORANGE)
rect(sl, 0, H - Inches(.08), W, Inches(.08), fill=C_CYAN)
rect(sl, Inches(7.2), 0, Inches(.06), H, fill=C_CYAN)

tb(sl, Inches(.55), Inches(.85), Inches(6.4), Inches(.42),
   "RÉVEIL PÉDAGOGIQUE  ·  PRÉ-SESSION J2 APRÈS-MIDI", size=10, bold=True, color=C_ORANGE)
tb_multi(sl, Inches(.55), Inches(1.35), Inches(6.4), Inches(2.2), [
    {"text": "SYNTHÈSE", "size": 52, "bold": True, "color": C_WHITE},
    {"text": "MODULE J2 — MATINÉE", "size": 24, "bold": False, "color": C_CYAN},
    {"text": "Le Monde Associatif & Humanitaire en France", "size": 14,
     "italic": True, "color": RGBColor(0xBB, 0xCC, 0xEE), "space_before": 8},
])
tb(sl, Inches(.55), Inches(3.72), Inches(6.4), Inches(.55),
   "Rappel des axes clés de la matinée avant de\ndémarrer la suite du programme.",
   size=13, italic=True, color=RGBColor(0xBB, 0xCC, 0xEE), wrap=True)

rect(sl, Inches(.55), Inches(4.45), Inches(6.4), Inches(.05), fill=C_CYAN)
tb(sl, Inches(.55), Inches(4.62), Inches(6.4), Inches(.34),
   "Objectif de cette synthèse :", size=11, bold=True, color=C_CYAN)
for i, obj in enumerate([
    "Ancrer les chiffres et définitions clés",
    "Comprendre le contexte historique et juridique",
    "Renforcer la légitimité du fundraiser",
]):
    tb(sl, Inches(.7), Inches(5.02) + i * Inches(.4), Inches(6.1), Inches(.36),
       f"→  {obj}", size=11, color=C_WHITE)

# Right panel
rect(sl, Inches(7.5), Inches(.4), Inches(5.5), Inches(6.6), fill=C_DARK)
rect(sl, Inches(7.5), Inches(.4), Inches(5.5), Inches(.04), fill=C_ORANGE)
tb(sl, Inches(7.75), Inches(.52), Inches(5.0), Inches(.3),
   "AU PROGRAMME — SYNTHÈSE", size=11, bold=True, color=C_ORANGE)
rect(sl, Inches(7.75), Inches(.88), Inches(5.0), Inches(.03), fill=C_CYAN)

chapters = [
    "01  Le Monde Associatif — 1,5 M associations",
    "02  Loi 1901 — Définition & principes",
    "03  Types d'associations — RUP / déclarée / fait",
    "04  Ressources — Sources de financement",
    "05  Histoire & Humanitaire — Chronologie",
    "06  Marché de la Collecte — Affaire ARC",
    "07  Tout Citoyen est Concerné — Exemples FR",
    "08  Impact UNICEF — Chiffres vérifiables",
]
for i, ch in enumerate(chapters):
    rect(sl, Inches(7.5), Inches(1.02) + i * Inches(.68), Inches(.04), Inches(.5), fill=C_CYAN)
    tb(sl, Inches(7.68), Inches(1.06) + i * Inches(.68), Inches(5.1), Inches(.42),
       ch, size=11, color=C_WHITE)


# ════════════════════════════════════════════════════════════════════
# SLIDE 2 — CHIFFRES CLÉS & DÉFINITIONS
# ════════════════════════════════════════════════════════════════════
sl = add_slide()
section_top(sl, "01", "CHIFFRES CLÉS — LE MONDE ASSOCIATIF", "Les données essentielles à retenir sur la vie associative française")
footer(sl)

stats = [
    (C_ORANGE, "1,5 M", "associations en France", "2e pays le plus associatif d'Europe après la Suède"),
    (C_BLUE_MID, "23 M", "bénévoles engagés", "Soit 1 Français sur 3 actif dans une asso"),
    (C_GREEN, "1,8 M", "salariés du secteur", "Devant l'industrie automobile en France"),
    (C_TEAL, "113 Mds€", "contribution au PIB", "5 % du PIB national — secteur économique majeur"),
    (C_PURPLE, "165 000", "RUP reconnues", "dont UNICEF, Croix-Rouge, MSF, Restos du Cœur"),
    (C_RED, "36 %", "des ressources = dons", "1er poste de financement pour les grandes ONG"),
]
for i, (accent, number, label, sub) in enumerate(stats):
    col = i % 3
    row = i // 3
    x = Inches(.4) + col * Inches(4.3)
    y = Inches(1.55) + row * Inches(2.5)
    rect(sl, x, y, Inches(4.05), Inches(2.25), fill=C_DARK, line=accent, line_w=Pt(1.5))
    rect(sl, x, y, Inches(4.05), Inches(.05), fill=accent)
    tb(sl, x + Inches(.15), y + Inches(.15), Inches(3.75), Inches(.7),
       number, size=38, bold=True, color=accent, align=PP_ALIGN.CENTER)
    tb(sl, x + Inches(.15), y + Inches(.88), Inches(3.75), Inches(.32),
       label, size=12, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER, wrap=True)
    rect(sl, x + Inches(.15), y + Inches(1.28), Inches(3.75), Inches(.03), fill=accent)
    tb(sl, x + Inches(.15), y + Inches(1.38), Inches(3.75), Inches(.72),
       sub, size=9, italic=True, color=RGBColor(0xBB, 0xCC, 0xDD), align=PP_ALIGN.CENTER, wrap=True)


# ════════════════════════════════════════════════════════════════════
# SLIDE 3 — LOI 1901 & TYPES D'ASSOCIATIONS
# ════════════════════════════════════════════════════════════════════
sl = add_slide()
section_top(sl, "02-03", "LOI 1901 & TYPES D'ASSOCIATIONS", "Définition légale et classifications — ce que vous devez savoir")
footer(sl)

rect(sl, Inches(.4), Inches(1.5), Inches(12.5), Inches(1.2), fill=C_DARK, line=C_CYAN, line_w=Pt(1.2))
rect(sl, Inches(.4), Inches(1.5), Inches(12.5), Inches(.04), fill=C_CYAN)
tb(sl, Inches(.65), Inches(1.58), Inches(12.0), Inches(.24),
   "LOI DU 1er JUILLET 1901 — ARTICLE 1er", size=9, bold=True, color=C_CYAN)
tb(sl, Inches(.65), Inches(1.86), Inches(12.0), Inches(.72),
   "« Une association est la convention par laquelle deux ou plusieurs personnes mettent en commun, d'une façon permanente, "
   "leurs connaissances ou leur activité dans un but autre que de partager des bénéfices. »",
   size=12, italic=True, color=C_WHITE, wrap=True)

types_data = [
    (C_GREY_LINE, "ASSOCIATION DE FAIT",
     ["Non déclarée", "Pas de personnalité juridique", "Pas de compte bancaire", "Responsabilité personnelle des membres"],
     "Groupe de voisins,\ncollectif informel"),
    (C_BLUE_MID, "ASSOCIATION DÉCLARÉE",
     ["Déclarée en préfecture (JO)", "Personnalité juridique", "Compte bancaire possible", "Peut recevoir des dons — forme la + courante"],
     "Club de foot, asso parents\nd'élèves, association locale"),
    (C_ORANGE, "RUP — UTILITÉ PUBLIQUE",
     ["Décret en Conseil d'État", "Critères stricts : 3 ans+ / 200+ membres", "Legs, donations avantages fiscaux", "~2 800 en France — UNICEF France"],
     "UNICEF France, Croix-Rouge,\nRestos du Cœur, MSF"),
]
for i, (accent, title, bullets, example) in enumerate(types_data):
    x = Inches(.4) + i * Inches(4.3)
    y = Inches(2.85)
    h = Inches(3.7)
    rect(sl, x, y, Inches(4.05), h, fill=C_WHITE, line=accent, line_w=Pt(1.5))
    rect(sl, x, y, Inches(4.05), Inches(.05), fill=accent)
    rect(sl, x, y + Inches(.05), Inches(4.05), Inches(.52), fill=accent)
    tb(sl, x + Inches(.12), y + Inches(.14), Inches(3.8), Inches(.3),
       title, size=11, bold=True, color=C_WHITE)
    for j, b in enumerate(bullets):
        tb(sl, x + Inches(.15), y + Inches(.72) + j * Inches(.52),
           Inches(3.75), Inches(.46), f"• {b}", size=9.5, color=C_TEXT_LIGHT, wrap=True)
    rect(sl, x + Inches(.12), y + h - Inches(.78), Inches(3.8), Inches(.65),
         fill=C_GREY, line=accent, line_w=Pt(.8))
    tb(sl, x + Inches(.22), y + h - Inches(.72), Inches(3.6), Inches(.58),
       f"Ex : {example}", size=9, italic=True, color=C_BLUE_DARK, wrap=True)


# ════════════════════════════════════════════════════════════════════
# SLIDE 4 — CHRONOLOGIE & PRINCIPES HUMANITAIRES
# ════════════════════════════════════════════════════════════════════
sl = add_slide()
section_top(sl, "04-05", "CHRONOLOGIE & PRINCIPES HUMANITAIRES", "Histoire du secteur et valeurs fondatrices de l'action humanitaire")
footer(sl)

# Timeline
rect(sl, Inches(.4), Inches(1.5), Inches(5.8), Inches(5.3), fill=C_WHITE, line=C_GREY_LINE, line_w=Pt(1))
rect(sl, Inches(.4), Inches(1.5), Inches(5.8), Inches(.04), fill=C_TEAL)
tb(sl, Inches(.65), Inches(1.6), Inches(5.3), Inches(.3),
   "CHRONOLOGIE CLÉS", size=12, bold=True, color=C_TEAL)
rect(sl, Inches(1.22), Inches(2.05), Inches(.04), Inches(4.5), fill=C_TEAL)

timeline = [
    ("XIIe s.", "Confréries médiévales — premières formes de solidarité collective organisée en Europe."),
    ("1859", "Solférino — Henry Dunant → Croix-Rouge. Naissance du droit humanitaire international."),
    ("1901", "Loi du 1er juillet 1901 — liberté d'association garantie en France par la loi."),
    ("1946", "Création de l'UNICEF par l'ONU. Présent aujourd'hui dans 190 pays."),
    ("1971", "Fondation de MSF — Médecins Sans Frontières. Modèle de l'urgence médicale."),
    ("1996", "Affaire ARC — crise de confiance. Réglementation et transparence imposées."),
]
for i, (year, text) in enumerate(timeline):
    y = Inches(2.05) + i * Inches(.78)
    rect(sl, Inches(1.08), y + Inches(.1), Inches(.3), Inches(.28), fill=C_TEAL)
    tb(sl, Inches(.5), y, Inches(.56), Inches(.28), year, size=8, bold=True, color=C_TEAL)
    tb(sl, Inches(1.55), y, Inches(4.4), Inches(.68), text, size=9, color=C_TEXT_LIGHT, wrap=True)

# Principles
rect(sl, Inches(6.6), Inches(1.5), Inches(6.3), Inches(5.3), fill=C_GREY)
rect(sl, Inches(6.6), Inches(1.5), Inches(6.3), Inches(.04), fill=C_ORANGE)
tb(sl, Inches(6.85), Inches(1.6), Inches(5.8), Inches(.3),
   "4 PRINCIPES HUMANITAIRES FONDAMENTAUX", size=11, bold=True, color=C_ORANGE)

principles = [
    (C_RED, "HUMANITÉ", "Protéger la vie et la santé. Secourir sans distinction. Prévenir et alléger la souffrance."),
    (C_BLUE_MID, "IMPARTIALITÉ", "Aide basée uniquement sur les besoins. Aucune discrimination par nationalité, race, religion."),
    (C_GREEN, "NEUTRALITÉ", "Ne pas prendre parti dans les hostilités ni dans les controverses politiques ou idéologiques."),
    (C_ORANGE, "INDÉPENDANCE", "Autonomie par rapport aux gouvernements, partis, groupes économiques ou religieux."),
]
for i, (accent, title, body) in enumerate(principles):
    y = Inches(2.08) + i * Inches(1.12)
    rect(sl, Inches(6.7), y, Inches(.06), Inches(.9), fill=accent)
    tb(sl, Inches(6.95), y + Inches(.04), Inches(5.6), Inches(.3),
       title, size=11, bold=True, color=accent)
    tb(sl, Inches(6.95), y + Inches(.35), Inches(5.6), Inches(.58),
       body, size=9.5, color=C_TEXT_LIGHT, wrap=True)


# ════════════════════════════════════════════════════════════════════
# SLIDE 5 — MARCHÉ & AFFAIRE ARC
# ════════════════════════════════════════════════════════════════════
sl = add_slide()
section_top(sl, "06", "MARCHÉ DE LA COLLECTE", "Contexte, chiffres clés et leçon de l'Affaire ARC")
footer(sl)

rect(sl, Inches(.4), Inches(1.5), Inches(5.8), Inches(5.3), fill=C_RED_LIGHT, line=C_RED, line_w=Pt(1.5))
rect(sl, Inches(.4), Inches(1.5), Inches(5.8), Inches(.05), fill=C_RED)
tb(sl, Inches(.65), Inches(1.62), Inches(5.3), Inches(.3),
   "⚠️  L'AFFAIRE ARC — CE QU'IL FAUT RETENIR", size=12, bold=True, color=C_RED)
tb(sl, Inches(.65), Inches(2.0), Inches(5.3), Inches(2.1),
   "1996 — L'Association de Recherche sur le Cancer (fondée 1962) est au cœur d'un scandale "
   "sans précédent : son directeur Jacques Crozemarie détourne des millions de dons "
   "à des fins personnelles.\n\n"
   "Conséquences : effondrement de la confiance, création du Comité de la Charte du Don "
   "en Confiance, audits obligatoires, publication des comptes.",
   size=10, color=C_TEXT_LIGHT, wrap=True)
rect(sl, Inches(.55), Inches(4.3), Inches(5.5), Inches(.55), fill=C_RED)
tb(sl, Inches(.7), Inches(4.38), Inches(5.2), Inches(.4),
   "LEÇON : Transparence + éthique = fondement de toute collecte.",
   size=11, bold=True, color=C_WHITE, wrap=True)
tb(sl, Inches(.65), Inches(5.05), Inches(5.3), Inches(.65),
   "Pour le fundraiser : votre légitimité repose sur l'intégrité de l'UNICEF. "
   "Le donateur doit SAVOIR que son argent arrive à destination. Vous êtes garant de cette confiance.",
   size=10, italic=True, color=C_TEXT_LIGHT, wrap=True)

rect(sl, Inches(6.6), Inches(1.5), Inches(6.3), Inches(5.3), fill=C_BLUE_LIGHT, line=C_BLUE_MID, line_w=Pt(1))
rect(sl, Inches(6.6), Inches(1.5), Inches(6.3), Inches(.05), fill=C_BLUE_MID)
tb(sl, Inches(6.85), Inches(1.62), Inches(5.8), Inches(.3),
   "LE MARCHÉ AUJOURD'HUI", size=12, bold=True, color=C_BLUE_DARK)

market = [
    ("5,5 Mds €", "collectés par an en France"),
    ("3 Mds €", "via le phoning professionnel (55 %)"),
    ("7,5 M", "donateurs réguliers actifs"),
    ("24 %", "des Français donnent régulièrement"),
    ("164 €", "don moyen annuel par donateur"),
    ("Don Loi 1901", "66 % à 75 % de réduction fiscale"),
    ("UNICEF = RUP", "Audité, transparent, label Don en Confiance"),
    ("70 %+", "des fonds vont directement aux programmes terrain"),
]
for i, (stat, label) in enumerate(market):
    y = Inches(2.05) + i * Inches(.58)
    rect(sl, Inches(6.7), y, Inches(.05), Inches(.42), fill=C_CYAN)
    tb(sl, Inches(6.9), y + Inches(.04), Inches(1.6), Inches(.3),
       stat, size=11, bold=True, color=C_BLUE_DARK)
    tb(sl, Inches(8.6), y + Inches(.06), Inches(4.1), Inches(.28),
       label, size=10, color=C_TEXT_LIGHT)


# ════════════════════════════════════════════════════════════════════
# SLIDE 6 — TOUT CITOYEN EST CONCERNÉ
# ════════════════════════════════════════════════════════════════════
sl = add_slide()
section_top(sl, "07", "TOUT CITOYEN EST CONCERNÉ", "Des exemples concrets vérifiables — France & Monde")
footer(sl)

tb(sl, Inches(.4), Inches(1.45), Inches(12.5), Inches(.32),
   "Chaque Français a bénéficié ou bénéficiera d'une action associative :", size=11, italic=True, color=C_TEXT_LIGHT)

examples = [
    (C_RED, "🏥 SANTÉ", "Ligue Cancer, AIDES, APF — 4M bénévoles soins / soutien"),
    (C_BLUE_MID, "📚 ENFANCE", "UNICEF, SOS Villages Enfants, Croix-Rouge aide devoirs"),
    (C_ORANGE, "🌊 CATASTROPHES", "Séisme Turquie 18M€ mobilisés en 72h. COVID aide isolés"),
    (C_GREEN, "👴 PERSONNES ÂGÉES", "Petits Frères des Pauvres — 45 000 aîné(e)s accompagné(e)s"),
    (C_PURPLE, "🏠 LOGEMENT", "Fondation Abbé Pierre — 4M de mal logés en France"),
    (C_TEAL, "🌍 INTERNATIONAL", "UNICEF — 190 pays, 1er acteur mondial de l'enfance"),
]
for i, (accent, title, body) in enumerate(examples):
    col = i % 3
    row = i // 3
    x = Inches(.35) + col * Inches(4.32)
    y = Inches(1.88) + row * Inches(2.5)
    rect(sl, x, y, Inches(4.1), Inches(2.3), fill=C_WHITE, line=accent, line_w=Pt(1.5))
    rect(sl, x, y, Inches(4.1), Inches(.55), fill=accent)
    rect(sl, x, y, Inches(4.1), Inches(.05), fill=C_DARK)
    tb(sl, x + Inches(.15), y + Inches(.12), Inches(3.8), Inches(.35),
       title, size=13, bold=True, color=C_WHITE)
    tb(sl, x + Inches(.15), y + Inches(.7), Inches(3.8), Inches(1.4),
       body, size=10.5, color=C_TEXT_LIGHT, wrap=True)


# ════════════════════════════════════════════════════════════════════
# SLIDE 7 — IMPACT UNICEF
# ════════════════════════════════════════════════════════════════════
sl = add_slide()
section_top(sl, "08", "IMPACT UNICEF — CHIFFRES VÉRIFIABLES", "Des résultats mesurables pour argumenter avec confiance")
footer(sl)

tb(sl, Inches(.4), Inches(1.45), Inches(12.5), Inches(.28),
   "Chiffres officiels UNICEF — vérifiables sur unicef.fr et unicef.org :", size=11, italic=True, color=C_TEXT_LIGHT)

impacts = [
    (C_BLUE_MID, "💉", "560 M enfants\nvaccinés / an", "Rougeole, polio, tétanos, méningite"),
    (C_TEAL, "💧", "820 M personnes\naccès eau potable", "3 € = eau pour 1 famille / 1 mois"),
    (C_ORANGE, "🌾", "45 €\nsauve 1 enfant", "De la malnutrition sévère — RUTF thérapeutique"),
    (C_GREEN, "📚", "73 M enfants\nnon scolarisés", "Kits scolaires d'urgence déployés"),
    (C_PURPLE, "🛡️", "120 pays\nsystèmes protection", "300M enfants en situation de risque couverts"),
    (C_RED, "🚨", "247 opérations\nd'urgence / an", "Déploiement en 72h — Turquie, Gaza, Haïti..."),
]
for i, (accent, icon, stat, detail) in enumerate(impacts):
    col = i % 3
    row = i // 3
    x = Inches(.35) + col * Inches(4.32)
    y = Inches(1.88) + row * Inches(2.5)
    rect(sl, x, y, Inches(4.1), Inches(2.3), fill=C_DARK, line=accent, line_w=Pt(1.5))
    rect(sl, x, y, Inches(4.1), Inches(.05), fill=accent)
    tb(sl, x + Inches(.15), y + Inches(.12), Inches(.5), Inches(.5),
       icon, size=24)
    tb(sl, x + Inches(.75), y + Inches(.12), Inches(3.2), Inches(.22),
       "UNICEF", size=8, bold=True, color=accent)
    tb(sl, x + Inches(.15), y + Inches(.62), Inches(3.8), Inches(.7),
       stat, size=18, bold=True, color=C_WHITE, wrap=True)
    rect(sl, x + Inches(.15), y + Inches(1.38), Inches(3.8), Inches(.04), fill=accent)
    tb(sl, x + Inches(.15), y + Inches(1.48), Inches(3.8), Inches(.72),
       detail, size=9.5, color=RGBColor(0xBB, 0xCC, 0xDD), wrap=True)


# ════════════════════════════════════════════════════════════════════
# SLIDE 8 — QUIZ RÉVEIL PÉDAGOGIQUE
# ════════════════════════════════════════════════════════════════════
sl = add_slide()
section_top(sl, "QUIZ", "RÉVEIL PÉDAGOGIQUE — QUIZ FLASH", "Testez vos connaissances avant de démarrer l'après-midi")
footer(sl)

quiz = [
    ("Combien d'associations existe-t-il en France ?",
     "1,5 million d'associations — dont 165 000 reconnues d'utilité publique (RUP)"),
    ("Que dit l'Article 1er de la Loi 1901 ?",
     "Convention de 2+ personnes mettant en commun connaissances/activité dans un but NON lucratif"),
    ("Quelle est la différence entre une asso déclarée et une RUP ?",
     "RUP = décret Conseil d'État, critères stricts, peut recevoir legs/donations avec avantages fiscaux"),
    ("Nommez les 4 principes humanitaires fondamentaux.",
     "Humanité — Impartialité — Neutralité — Indépendance"),
    ("Que s'est-il passé en 1996 avec l'Affaire ARC ?",
     "Détournements massifs. Crise de confiance → Charte du Don en Confiance + audits obligatoires"),
    ("Quel est le montant collecté chaque année en France ?",
     "5,5 milliards d'euros dont 3 milliards via le phoning professionnel (55 %)"),
]
for i, (q, r) in enumerate(quiz):
    col = i % 2
    row = i // 2
    x = Inches(.35) + col * Inches(6.5)
    y = Inches(1.55) + row * Inches(1.88)
    qr_bloc(sl, x, y, Inches(6.2), Inches(1.72), q, r)


# ════════════════════════════════════════════════════════════════════
# SLIDE 9 — LES 10 POINTS À RETENIR
# ════════════════════════════════════════════════════════════════════
sl = add_slide()
rect(sl, 0, 0, W, H, fill=C_DARK)
rect(sl, 0, 0, W, Inches(.06), fill=C_ORANGE)
rect(sl, 0, H - Inches(.06), W, Inches(.06), fill=C_CYAN)
rect(sl, Inches(6.55), 0, Inches(.04), H, fill=C_CYAN)

tb(sl, Inches(.4), Inches(.25), Inches(5.8), Inches(.42),
   "LES 10 POINTS ESSENTIELS", size=18, bold=True, color=C_ORANGE)
tb(sl, Inches(.4), Inches(.72), Inches(5.8), Inches(.3),
   "À retenir de la matinée J2", size=12, italic=True, color=C_CYAN)

points_l = [
    ("1,5 M", "associations en France — 23 M bénévoles"),
    ("Loi 1901", "but non lucratif, convention, utilité sociale"),
    ("3 statuts", "de fait / déclarée / RUP — UNICEF = RUP"),
    ("36 %", "des ressources asso viennent des dons"),
    ("1859", "Henry Dunant, Solférino, Croix-Rouge"),
]
points_r = [
    ("4 principes", "Humanité / Impartialité / Neutralité / Indépendance"),
    ("1996", "Affaire ARC — transparence obligatoire"),
    ("5,5 Mds €", "collectés/an — 55 % par phoning"),
    ("190 pays", "UNICEF présent — 1er acteur mondial de l'enfance"),
    ("560 M", "enfants vaccinés / an grâce aux dons réguliers"),
]
for i, (stat, label) in enumerate(points_l):
    y = Inches(1.2) + i * Inches(.57)
    rect(sl, Inches(.4), y, Inches(1.1), Inches(.42), fill=C_BLUE_MID)
    tb(sl, Inches(.4), y + Inches(.06), Inches(1.1), Inches(.3),
       stat, size=10, bold=True, color=C_ORANGE, align=PP_ALIGN.CENTER)
    tb(sl, Inches(1.62), y + Inches(.08), Inches(4.4), Inches(.3),
       label, size=10, color=C_WHITE)

for i, (stat, label) in enumerate(points_r):
    y = Inches(1.2) + i * Inches(.57)
    rect(sl, Inches(6.8), y, Inches(1.3), Inches(.42), fill=C_BLUE_MID)
    tb(sl, Inches(6.8), y + Inches(.06), Inches(1.3), Inches(.3),
       stat, size=9, bold=True, color=C_ORANGE, align=PP_ALIGN.CENTER)
    tb(sl, Inches(8.22), y + Inches(.08), Inches(4.8), Inches(.3),
       label, size=10, color=C_WHITE)

rect(sl, Inches(.4), Inches(4.25), Inches(12.5), Inches(.04), fill=C_CYAN)
rect(sl, Inches(.4), Inches(4.38), Inches(12.5), Inches(1.25), fill=C_BLUE_DARK)
tb(sl, Inches(.7), Inches(4.5), Inches(12.0), Inches(.32),
   "MANTRA J2 MATINÉE", size=10, bold=True, color=C_CYAN)
tb(sl, Inches(.7), Inches(4.88), Inches(12.0), Inches(.62),
   "« COMPRENDRE POUR CONVAINCRE — CHAQUE CITOYEN EST DIRECTEMENT OU INDIRECTEMENT CONCERNÉ »",
   size=16, bold=True, italic=True, color=C_WHITE, align=PP_ALIGN.CENTER, wrap=True)

rect(sl, Inches(.4), Inches(5.82), Inches(12.5), Inches(1.2), fill=RGBColor(0x0A, 0x1F, 0x3A))
tb(sl, Inches(.7), Inches(5.95), Inches(12.0), Inches(.28),
   "PROCHAINE ÉTAPE — APRÈS-MIDI J2", size=11, bold=True, color=C_ORANGE)
tb(sl, Inches(.7), Inches(6.28), Inches(12.0), Inches(.62),
   "Traitement des objections  •  Discours de conviction  •  Simulations téléphoniques  •  Plan d'action personnel",
   size=11, color=C_WHITE, align=PP_ALIGN.CENTER)

footer(sl)


# ════════════════════════════════════════════════════════════════════
# SLIDE 10 — QUALIFICATION & NOMENCLATURE COMPLÉMENTAIRE
# ════════════════════════════════════════════════════════════════════
sl = add_slide()
section_top(sl, "→", "QUALIFICATION & NOMENCLATURE — POINTS CLÉS",
            "50 % de la qualité Fidelis repose sur une qualification CONFORME", accent=C_RED)
footer(sl)

# Top warning
rect(sl, Inches(.35), Inches(1.5), Inches(12.6), Inches(.88), fill=C_DARK, line=C_RED, line_w=Pt(1.5))
rect(sl, Inches(.35), Inches(1.5), Inches(12.6), Inches(.05), fill=C_RED)
tb(sl, Inches(.55), Inches(1.62), Inches(12.0), Inches(.24),
   "IMPACT D'UNE FAUSSE QUALIFICATION", size=10, bold=True, color=C_RED)
tb(sl, Inches(.55), Inches(1.9), Inches(12.0), Inches(.42),
   "Refus Argumenté / Refus de Répondre = non contacté 4 mois UNICEF.  "
   "SHY Performance & Nescall : 300 000 – 400 000 fiches/mois chacun. Chaque fiche mal qualifiée = contact perdu.",
   size=10, color=C_WHITE, wrap=True)

# Qualification table
qual_rows = [
    (C_GREEN, "✅  PEL", "Prélèvement En Ligne — mission accomplie"),
    (C_GREEN, "✅  RIM", "Rappel Imminent — intéressé, à rappeler vite"),
    (C_TEAL, "🕐  ABSENT", "Répondeur / NRP / Raccroché SANS identification — PAS un refus"),
    (C_TEAL, "🕐  DEL", "Don En Ligne — préfère donner via lien web"),
    (C_RED, "❌  RA", "Refus Argumenté — raison valide — 4 mois sans contact UNICEF"),
    (C_RED, "❌  RR", "Refus de Répondre — raccroché / agressif — idem 4 mois"),
]
qy = Inches(2.52)
for accent, code, desc in qual_rows:
    clr = C_GREEN_LIGHT if accent == C_GREEN else (C_TEAL_LIGHT if accent == C_TEAL else C_RED_LIGHT)
    rect(sl, Inches(.35), qy, Inches(12.6), Inches(.4), fill=clr, line=C_GREY_LINE, line_w=Pt(.5))
    rect(sl, Inches(.35), qy, Inches(.06), Inches(.4), fill=accent)
    tb(sl, Inches(.55), qy + Inches(.07), Inches(2.2), Inches(.26),
       code, size=10, bold=True, color=C_BLUE_DARK)
    tb(sl, Inches(2.85), qy + Inches(.07), Inches(9.8), Inches(.26),
       desc, size=10, color=C_TEXT_LIGHT)
    qy += Inches(.41)

# Error example
rect(sl, Inches(.35), qy + Inches(.08), Inches(12.6), Inches(.48), fill=C_ORANGE)
tb(sl, Inches(.55), qy + Inches(.14), Inches(12.0), Inches(.36),
   "EXEMPLE ERREUR : Répondeur → qualifié RA (FAUX) — le bon code est ABSENT. La personne n'a pas refusé, elle était absente.",
   size=10, bold=True, color=C_WHITE, wrap=True)
qy += Inches(.62)

# Nomenclature complémentaire
new_terms = [
    (C_TEAL, "ABSENT", "Répondeur, NRP, raccroché sans ID — recontacter, pas un refus"),
    (C_RED, "CICR", "Croix-Rouge internationale — zones de conflit, échanges d'otages, prisonniers de guerre"),
    (C_BLUE_MID, "COMITÉ DE LA CHARTE", "Créé 1989, renforcé 1996 — audits, label Don en Confiance, contrôle État"),
    (C_ORANGE, "PARTICULIER / PRO", "Dons = particuliers uniquement. Pro atteint → ABSENT ou demander numéro perso"),
    (C_PURPLE, "FIDELIS = RUP", "Fidelis travaille EXCLUSIVEMENT avec associations Reconnues d'Utilité Publique"),
]
ny = qy + Inches(.1)
for accent, term, desc in new_terms:
    clr = C_BLUE_LIGHT
    rect(sl, Inches(.35), ny, Inches(12.6), Inches(.38), fill=clr, line=accent, line_w=Pt(.8))
    rect(sl, Inches(.35), ny, Inches(.06), Inches(.38), fill=accent)
    tb(sl, Inches(.55), ny + Inches(.06), Inches(2.5), Inches(.26),
       term, size=10, bold=True, color=accent)
    tb(sl, Inches(3.1), ny + Inches(.06), Inches(9.6), Inches(.26),
       desc, size=10, color=C_TEXT_LIGHT)
    ny += Inches(.4)


OUT = "/home/user/fatou/Synthese_J2_Matin_Fidelis.pptx"
prs.save(OUT)
print(f"Saved: {OUT}")
print(f"Slides: {len(prs.slides)}")

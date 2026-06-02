"""
FIM V6-FINAL — Générateur PPTX — 7 modules
Charte SHY-Performance (teal #008080 / blanc — sans rouge)
Logo SHY-Performance sur chaque slide
Prospection TÉLÉPHONIQUE (FIDELIS × UNICEF France)
Données exactes semaine 19–23/05/2026
Slides adressées aux apprenants · Méthodes pédagogiques non affichées
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
import os

# ─── CHARTE GRAPHIQUE SHY-PERFORMANCE ──────────────────────────────────────
# Couleurs d'origine (synopsis_formation_fundraiser_FIM.html)
TEAL      = RGBColor(0x00, 0x80, 0x80)   # #008080  — couleur principale
TEAL_D    = RGBColor(0x00, 0x55, 0x55)   # foncé
TEAL_L    = RGBColor(0xF0, 0xF7, 0xF7)   # clair fond
TEAL_MID  = RGBColor(0x00, 0x6A, 0x6A)   # intermédiaire
WHITE     = RGBColor(0xFF, 0xFF, 0xFF)
DARK      = RGBColor(0x1A, 0x1A, 0x1A)
GRAY      = RGBColor(0x55, 0x55, 0x55)
LGRAY     = RGBColor(0xDD, 0xDD, 0xDD)
TEAL_ACC  = RGBColor(0x00, 0xAA, 0xAA)   # accent clair

LOGO_PATH = "/home/user/fatou/shy_logo_v2.png"
OUT_DIR   = "/home/user/fatou"
VERSION   = "V6-01062026SHY-TE"
COMPANY   = "SHY-Performance"

# ─── PRIMITIVES ─────────────────────────────────────────────────────────────

def new_prs():
    prs = Presentation()
    prs.slide_width  = Inches(13.33)
    prs.slide_height = Inches(7.5)
    return prs

def SL(prs): return prs.slide_layouts[6]

def r(slide, x, y, w, h, fill=None, line=None, lw=Pt(0)):
    sh = slide.shapes.add_shape(1, Inches(x), Inches(y), Inches(w), Inches(h))
    sh.line.width = lw
    if fill:
        sh.fill.solid(); sh.fill.fore_color.rgb = fill
    else:
        sh.fill.background()
    if line:
        sh.line.color.rgb = line
    else:
        sh.line.fill.background()
    return sh

def t(slide, text, x, y, w, h, sz=13, bold=False, color=None,
      align=PP_ALIGN.LEFT, italic=False, wrap=True):
    color = color or DARK
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    box.word_wrap = wrap
    tf = box.text_frame; tf.word_wrap = wrap
    p = tf.paragraphs[0]; p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(sz); run.font.bold = bold
    run.font.italic = italic; run.font.color.rgb = color
    return box

def logo(slide, x=12.3, y=0.05, w=0.9, h=0.88):
    """Insère le logo SHY-Performance en haut à droite"""
    try:
        slide.shapes.add_picture(LOGO_PATH, Inches(x), Inches(y), Inches(w), Inches(h))
    except Exception:
        pass

# ─── ENTÊTE SHY-PERFORMANCE ─────────────────────────────────────────────────

def header_title(slide, title, subtitle="", tag=""):
    """Slide de titre : bandeau teal plein + logo + texte"""
    r(slide, 0, 0, 13.33, 7.5, fill=TEAL_L)
    r(slide, 0, 0, 13.33, 2.55, fill=TEAL_D)
    r(slide, 0, 2.53, 13.33, 0.06, fill=TEAL_ACC)
    # Nom société
    t(slide, COMPANY, 0.38, 0.1, 3.5, 0.48,
      sz=22, bold=True, color=WHITE)
    t(slide, "Formation Initiale Module — Fundraising Téléphonique UNICEF",
      0.38, 0.58, 8.5, 0.38, sz=10, color=RGBColor(0xBB,0xEE,0xEE), italic=True)
    if tag:
        t(slide, tag, 0.38, 1.02, 8.5, 0.32, sz=9, bold=True,
          color=RGBColor(0x88,0xDD,0xDD))
    t(slide, title, 0.38, 1.32, 11.5, 1.0, sz=30, bold=True, color=WHITE)
    if subtitle:
        t(slide, subtitle, 0.38, 2.08, 11.5, 0.44, sz=12,
          color=RGBColor(0xCC,0xEE,0xEE), italic=True)
    logo(slide)
    _footer(slide)

def header_content(slide, title):
    """Bandeau compact pour slides de contenu"""
    r(slide, 0, 0, 13.33, 7.5, fill=TEAL_L)
    r(slide, 0, 0, 13.33, 0.75, fill=TEAL_D)
    r(slide, 0, 0.73, 13.33, 0.05, fill=TEAL_ACC)
    t(slide, COMPANY, 0.18, 0.08, 2.2, 0.38, sz=14, bold=True, color=WHITE)
    r(slide, 2.5, 0.1, 0.04, 0.55, fill=TEAL_ACC)
    t(slide, title, 2.65, 0.1, 10.1, 0.55, sz=16, bold=True, color=WHITE)
    logo(slide, x=12.3, y=0.0, w=0.9, h=0.74)
    _footer(slide)

def _footer(slide):
    r(slide, 0, 7.1, 13.33, 0.4, fill=TEAL_D)
    t(slide, f"{COMPANY}  ·  Formation FIM  ·  UNICEF France  ·  {VERSION}",
      0, 7.13, 13.33, 0.3, sz=8, color=WHITE, align=PP_ALIGN.CENTER)

# ─── COMPOSANTS ─────────────────────────────────────────────────────────────

def two_col(slide, lt, li, rt, ri, y=0.88):
    cw = 6.18
    for (head, items, ox) in [(lt, li, 0.3), (rt, ri, 6.83)]:
        r(slide, ox, y, cw, 0.36, fill=TEAL)
        t(slide, head, ox+0.12, y+0.05, cw-0.2, 0.28,
          sz=10, bold=True, color=WHITE)
        iy = y + 0.38
        for i, item in enumerate(items):
            bg = WHITE if i % 2 == 0 else TEAL_L
            r(slide, ox, iy, cw, 0.44, fill=bg, line=TEAL, lw=Pt(0.4))
            t(slide, "• " + item, ox+0.14, iy+0.06, cw-0.24, 0.34,
              sz=10, color=DARK)
            iy += 0.46

def seq4(slide, steps, labels, y=0.88, h=5.1):
    """Séquence 4 phases sans nom de méthode"""
    colors = [TEAL_D, TEAL, TEAL_MID, TEAL_ACC]
    sw = (13.33 - 0.6) / 4
    for i, (step, label, col) in enumerate(zip(steps, labels, colors)):
        bx = 0.3 + i * (sw + 0.02)
        r(slide, bx, y, sw, 0.32, fill=col)
        t(slide, label, bx+0.1, y+0.04, sw-0.18, 0.26,
          sz=9, bold=True, color=WHITE)
        r(slide, bx, y+0.32, sw, h-0.32,
          fill=WHITE, line=col, lw=Pt(1.2))
        t(slide, step, bx+0.12, y+0.38, sw-0.22, h-0.55,
          sz=9.5, color=DARK, wrap=True)

def highlight(slide, text, x, y, w, h, fill=None, tc=None):
    fill = fill or TEAL
    tc   = tc   or WHITE
    r(slide, x, y, w, h, fill=fill)
    t(slide, text, x+0.15, y+0.08, w-0.3, h-0.16,
      sz=10, color=tc, wrap=True)

def kpi_row(slide, items, y=6.08, label="À retenir"):
    r(slide, 0.3, y, 12.73, 0.28, fill=TEAL)
    t(slide, label, 0.42, y+0.04, 12.5, 0.22,
      sz=8, bold=True, color=WHITE)
    w = 12.73 / len(items)
    for i, item in enumerate(items):
        bg = WHITE if i % 2 == 0 else TEAL_L
        r(slide, 0.3+i*w, y+0.3, w-0.04, 0.55,
          fill=bg, line=TEAL, lw=Pt(0.6))
        t(slide, "✔ "+item, 0.42+i*w, y+0.34, w-0.18, 0.46,
          sz=9, color=DARK, wrap=True)

def table_header(slide, cols, col_x, col_w, y):
    for h_txt, x, w in zip(cols, col_x, col_w):
        r(slide, x, y, w, 0.3, fill=TEAL_D)
        t(slide, h_txt, x+0.06, y+0.04, w-0.1, 0.24,
          sz=9, bold=True, color=WHITE,
          align=PP_ALIGN.CENTER if len(h_txt) <= 3 else PP_ALIGN.LEFT)

def save(prs, name):
    path = os.path.join(OUT_DIR, name)
    prs.save(path)
    print(f"  ✓ {name}  ({os.path.getsize(path)//1024} KB)")
    return path


# ══════════════════════════════════════════════════════════════════════════
#  BOOK 0 — FIL CONDUCTEUR
# ══════════════════════════════════════════════════════════════════════════

def build_book0():
    prs = new_prs()

    # Slide 1 — Titre
    s = prs.slides.add_slide(SL(prs))
    header_title(s,
        "Bienvenue dans la Formation FIM",
        "Fundraiser Impact Mission — UNICEF France × SHY-Performance × FIDELIS",
        tag="BOOK 0  ·  Fil Conducteur  ·  Document de parcours")
    t(s,
      "Ce livret est votre guide pour les 5 jours de formation.\n"
      "Il contient votre planning, vos engagements et votre tableau de bord personnel.",
      0.5, 2.75, 11.8, 0.9, sz=14, color=DARK)
    t(s,
      "Vous allez maîtriser le script officiel UNICEF en 7 étapes, qualifier vos appels\n"
      "selon la nomenclature FIDELIS et atteindre l'objectif de 9 CU/H minimum.",
      0.5, 3.75, 11.8, 0.9, sz=13, bold=True, color=TEAL_D)

    # Slide 2 — Les 5 journées
    s = prs.slides.add_slide(SL(prs))
    header_content(s, "Vos 5 Journées — Ce qui vous attend")
    days = [
        ("JOUR 1",  "Fondations\n& Mission",
         "Mission UNICEF · KPI (CU/H, TX Transfo, PDC) · Nomenclature FIDELIS · Organisation de la production"),
        ("JOUR 2",  "Le Monde\nAssociatif",
         "Loi 1901 · Histoire humanitaire · Principes humanitaires · Don régulier · Label Don en Confiance"),
        ("JOUR 3",  "Analyse &\nScript",
         "Script officiel UNICEF 7 étapes · 5 règles d'accroche · 4 modes de validation du don · Lecture à voix haute"),
        ("JOUR 4",  "Traitement\ndes Objections",
         "15 objections · 5 catégories · Méthode AAR · Jeux de rôle · Grille formateur"),
        ("JOUR 5",  "Synthèse &\nCertification",
         "Simulations d'appels complets · Grille 24 critères · Quiz 40 questions · Remise des certifications"),
    ]
    x = 0.3
    for tag, title, desc in days:
        r(s, x, 0.88, 2.5, 0.36, fill=TEAL_D)
        t(s, tag, x+0.08, 0.91, 2.34, 0.28, sz=10, bold=True, color=WHITE)
        r(s, x, 1.24, 2.5, 1.05, fill=TEAL)
        t(s, title, x+0.12, 1.27, 2.26, 0.98, sz=13, bold=True,
          color=WHITE, align=PP_ALIGN.CENTER)
        r(s, x, 2.29, 2.5, 3.6, fill=WHITE, line=TEAL, lw=Pt(1.5))
        t(s, desc, x+0.14, 2.37, 2.22, 3.42, sz=10, color=DARK, wrap=True)
        x += 2.58
    highlight(s,
        "35 heures de formation intensive · Organisation journée production : 9h30–17h30 · 6h40 de production effective",
        0.3, 6.05, 12.73, 0.38, fill=TEAL)

    # Slide 3 — Votre contrat d'engagement
    s = prs.slides.add_slide(SL(prs))
    header_content(s, "Votre Contrat d'Engagement")
    two_col(s,
        "Vos engagements en tant qu'apprenant",
        ["Participer activement à toutes les mises en situation",
         "Compléter votre livret personnel chaque soir (5 min)",
         "Demander un feedback précis après chaque exercice",
         "Respecter la confidentialité des échanges du groupe",
         "Arriver à l'heure — le groupe démarre ensemble"],
        "Ce que SHY-Performance s'engage à vous offrir",
        ["Un feedback individuel personnalisé chaque jour",
         "Des exercices pratiques progressifs et bienveillants",
         "Un formateur qui adapte le rythme à votre profil",
         "Un suivi après la formation à J+15 et J+30",
         "Une certification reconnue FIDELIS × UNICEF France"])
    highlight(s,
        "Droit à l'erreur garanti — Chaque simulation ratée est un pas vers la maîtrise. "
        "Mantra de la formation : « Le refus d'aujourd'hui peut être le don de demain »",
        0.3, 6.02, 12.73, 0.44)

    # Slide 4 — Votre livret personnel
    s = prs.slides.add_slide(SL(prs))
    header_content(s, "Votre Livret Personnel — Journal de Bord Quotidien")
    t(s, "Chaque soir, 5 minutes pour compléter ces 4 points :",
      0.35, 0.88, 12.5, 0.35, sz=12, bold=True, color=TEAL_D)
    entries = [
        ("①", "Ce que j'ai réussi aujourd'hui",
         "Notez 1 ou 2 moments où vous avez senti la progression.\n"
         "Ex : 'J'ai appliqué les 5 règles d'accroche sans hésitation.'"),
        ("②", "Ce qui m'a challengé",
         "Une situation difficile et pourquoi elle vous a bloqué.\n"
         "Ex : 'Quand le prospect a dit \"arnaque\", j'ai manqué la reformulation.'"),
        ("③", "Mon insight du jour",
         "L'apprentissage que vous n'auriez pas pu anticiper ce matin.\n"
         "Ex : 'La pause après la proposition vaut mieux qu'une relance immédiate.'"),
        ("④", "Mon engagement pour demain",
         "1 comportement précis à changer ou renforcer dès demain.\n"
         "Ex : 'Je vais sourire à voix haute dès le mot \"Bonjour\".'"),
    ]
    y = 1.32
    for num, title, body in entries:
        r(s, 0.3, y, 0.58, 0.98, fill=TEAL)
        t(s, num, 0.32, y+0.2, 0.54, 0.58, sz=20, bold=True,
          color=WHITE, align=PP_ALIGN.CENTER)
        r(s, 0.9, y, 12.1, 0.98, fill=WHITE, line=TEAL, lw=Pt(0.5))
        t(s, title, 1.02, y+0.06, 11.9, 0.3, sz=11, bold=True, color=TEAL_D)
        t(s, body, 1.02, y+0.37, 11.9, 0.54, sz=10, color=DARK, italic=True)
        y += 1.02
    highlight(s,
        "Votre livret reste strictement personnel — il ne sera pas noté. C'est votre outil de progression.",
        0.3, 5.42, 12.73, 0.38)

    return save(prs, f"FIM_Book0_FilConducteur_{VERSION}.pptx")


# ══════════════════════════════════════════════════════════════════════════
#  BOOK 1 — SYNOPSIS & OBJECTIFS
# ══════════════════════════════════════════════════════════════════════════

def build_book1():
    prs = new_prs()

    # Slide 1 — Titre
    s = prs.slides.add_slide(SL(prs))
    header_title(s,
        "Vos Objectifs de Formation",
        "Ce que vous saurez faire à l'issue des 5 jours FIM",
        tag="BOOK 1  ·  Synopsis & Objectifs  ·  SHY-Performance × FIDELIS × UNICEF France")
    t(s,
      "À l'issue de ces 5 jours, vous serez capable de conduire un appel de collecte de dons\n"
      "au nom de l'UNICEF, de la phrase d'accroche jusqu'à la validation du Prélèvement Automatique.",
      0.5, 2.75, 11.8, 0.9, sz=14, color=DARK)

    # Slide 2 — SAVOIR / SAVOIR-FAIRE / SAVOIR-ÊTRE
    s = prs.slides.add_slide(SL(prs))
    header_content(s, "Vos 3 Dimensions de Compétence")
    cols = [
        ("SAVOIR", "Connaissance de la cause",
         TEAL_D,
         ["Présenter l'UNICEF, sa mission, son histoire (75 ans, 190 pays)",
          "Expliquer la malnutrition aigüe sévère et le rôle des sachets RUTF (92g)",
          "Situer le secteur associatif (Loi 1901, Comité de la Charte, Bloctel, RGPD)",
          "Citer les arguments de légitimité UNICEF",
          "Connaître les 3 principes humanitaires fondamentaux",
          "Comprendre Conquête vs Fidélisation/Réactivation"]),
        ("SAVOIR-FAIRE", "Technique d'appel & qualification",
         TEAL,
         ["Appliquer les 5 règles officielles de l'accroche téléphonique",
          "Conduire le script officiel UNICEF en 7 étapes",
          "Qualifier chaque appel selon la nomenclature FIDELIS (DON/INDÉCIS/REFUS)",
          "Gérer les 4 modes de validation du don (IBAN à chaud, PA en ligne, PEL, PA courrier)",
          "Traiter les 15 objections types par la méthode AAR",
          "Atteindre l'objectif de 9 CU/H minimum en production"]),
        ("SAVOIR-ÊTRE", "Posture & conviction",
         TEAL_MID,
         ["Maintenir un ton chaleureux, souriant et naturel tout au long de l'appel",
          "Gérer les objections sensibles avec calme et professionnalisme",
          "Transmettre une conviction sincère pour la mission UNICEF",
          "Intégrer le mantra : «Le refus d'aujourd'hui peut être le don de demain»",
          "Adopter une posture de représentant digne de la cause humanitaire",
          "Respecter systématiquement les règles éthiques du secteur associatif"]),
    ]
    x = 0.3
    for label, sub, col, items in cols:
        r(s, x, 0.88, 4.18, 0.52, fill=col)
        t(s, label, x+0.12, 0.9, 4.0, 0.3, sz=14, bold=True, color=WHITE)
        t(s, sub, x+0.12, 1.2, 4.0, 0.22, sz=9, color=RGBColor(0xCC,0xEE,0xEE))
        iy = 1.44
        for i, item in enumerate(items):
            bg = WHITE if i % 2 == 0 else TEAL_L
            r(s, x, iy, 4.18, 0.5, fill=bg, line=col, lw=Pt(0.4))
            t(s, "• "+item, x+0.12, iy+0.06, 3.96, 0.4, sz=9, color=DARK)
            iy += 0.52
        x += 4.35

    # Slide 3 — KPI & nomenclature FIDELIS
    s = prs.slides.add_slide(SL(prs))
    header_content(s, "Vos Indicateurs de Performance — Nomenclature FIDELIS")
    kpis = [
        ("CU/H",        "Contacts Utiles / Heure",
         "Nombre de contacts qualifiés (DON/INDÉCIS/REFUS) par heure de production",
         "≥ 9 CU/H"),
        ("TX Transfo",  "Taux de Transformation",
         "% de PEL/PA obtenus parmi les Contacts Utiles — mesure l'efficacité au don régulier",
         "Cible FIDELIS"),
        ("PDC",         "Plan de Charge",
         "Volume mensuel de dons réguliers à atteindre selon FIDELIS × UNICEF France",
         "Volume défini"),
        ("PEL",         "Promesse En Ligne",
         "Prospect ayant accepté de valider son don en ligne (lien envoyé en direct)",
         "Comptabilisé TX"),
        ("PA",          "Prélèvement Automatique",
         "Don régulier sécurisé par mandat SEPA (IBAN à chaud, PA en ligne, PA courrier)",
         "Produit unique"),
        ("Qualification","Conformité des qualifications",
         "Taux de qualifications d'appels correctement renseignées (DON/INDÉCIS/REFUS)",
         "≥ 50% qualité"),
    ]
    col_x = [0.3, 1.35, 4.65, 10.38]
    col_w = [1.02, 3.27, 5.7,  2.65]
    table_header(s, ["Code", "Libellé complet", "Définition opérationnelle", "Objectif"],
                 col_x, col_w, y=0.88)
    y = 1.2
    for i, (code, lib, defin, obj) in enumerate(kpis):
        bg = WHITE if i % 2 == 0 else TEAL_L
        rh = 0.65
        r(s, col_x[0], y, col_w[0], rh, fill=TEAL_D)
        t(s, code, col_x[0]+0.06, y+0.1, col_w[0]-0.1, rh-0.15,
          sz=9, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        r(s, col_x[1], y, col_w[1], rh, fill=bg)
        t(s, lib, col_x[1]+0.1, y+0.1, col_w[1]-0.15, rh-0.15,
          sz=10, bold=True, color=TEAL_D)
        r(s, col_x[2], y, col_w[2], rh, fill=bg)
        t(s, defin, col_x[2]+0.1, y+0.08, col_w[2]-0.15, rh-0.1,
          sz=9, color=DARK, italic=True)
        r(s, col_x[3], y, col_w[3], rh, fill=TEAL)
        t(s, obj, col_x[3]+0.1, y+0.15, col_w[3]-0.15, rh-0.22,
          sz=10, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        y += rh + 0.02

    # Slide 4 — 4 modes de validation
    s = prs.slides.add_slide(SL(prs))
    header_content(s, "Les 4 Modes de Validation du Don — Votre Arsenal de Closing")
    modes = [
        ("MODE 1\nIBAN à Chaud",
         TEAL_D,
         "Le prospect dicte son IBAN directement lors de l'appel.\n\n"
         "✔ Validation immédiate\n✔ Sécurisation mandat SEPA\n✔ Confirmation par SMS\n\n"
         "→ Mode le plus efficace\nPrioritaire en conquête"),
        ("MODE 2\nPA en Ligne — Direct",
         TEAL,
         "Vous envoyez un lien de paiement sécurisé pendant l'appel.\n\n"
         "✔ Prospect valide seul en temps réel\n✔ Confirmation email automatique\n✔ Zéro saisie IBAN par téléphone\n\n"
         "→ Pour les prospects méfiants sur l'IBAN"),
        ("MODE 3\nPromesse PA en Ligne (PEL)",
         TEAL_MID,
         "Le prospect s'engage à valider le lien après l'appel.\n\n"
         "✔ Lien envoyé par SMS / email\n✔ Relance automatique J+1\n✔ Comptabilisé dans le TX Transfo\n\n"
         "→ Pour les indécis à fort potentiel"),
        ("MODE 4\nPA Courrier",
         TEAL_ACC,
         "Envoi d'un formulaire papier avec enveloppe T pré-affranchie.\n\n"
         "✔ Pour les seniors moins digitaux\n✔ Mandat SEPA envoyé par voie postale\n✔ Délai de retour : 5–10 jours\n\n"
         "→ Alternative quand les autres modes échouent"),
    ]
    x = 0.3
    for title, col, body in modes:
        r(s, x, 0.88, 3.08, 0.65, fill=col)
        t(s, title, x+0.12, 0.9, 2.84, 0.61, sz=12, bold=True,
          color=WHITE, align=PP_ALIGN.CENTER)
        r(s, x, 1.53, 3.08, 4.5, fill=WHITE, line=col, lw=Pt(1.5))
        t(s, body, x+0.15, 1.61, 2.78, 4.24, sz=10, color=DARK, wrap=True)
        x += 3.2
    highlight(s,
        "Votre priorité : proposer le MODE 1 (IBAN à chaud) en premier. "
        "Si refus, enchaîner Mode 2 → Mode 3 → Mode 4. Ne jamais sauter directement au Mode 4.",
        0.3, 6.2, 12.73, 0.44)

    return save(prs, f"FIM_Book1_Synopsis_Objectifs_{VERSION}.pptx")


# ══════════════════════════════════════════════════════════════════════════
#  BOOK J1 — FONDATIONS & MISSION
# ══════════════════════════════════════════════════════════════════════════

def build_bookJ1():
    prs = new_prs()

    # Slide 1
    s = prs.slides.add_slide(SL(prs))
    header_title(s,
        "Fondations & Mission",
        "Votre identité de fundraiser · La cause · Votre raison d'agir",
        tag="BOOK J1  ·  Journée 1  ·  UNICEF France × SHY-Performance")
    t(s,
      "Aujourd'hui vous ancrez votre identité de fundraiser téléphonique.\n"
      "Vous comprenez la mission UNICEF, les KPI de production et la nomenclature\n"
      "qui structurera chacune de vos journées de travail.",
      0.5, 2.75, 11.8, 1.1, sz=14, color=DARK)

    # Slide 2 — Séquence J1
    s = prs.slides.add_slide(SL(prs))
    header_content(s, "Journée 1 — Ce que vous allez vivre aujourd'hui")
    seq4(s,
        ["Vous entendez le témoignage\nd'un enfant bénéficiaire UNICEF.\nNotez ce que vous ressentez.",
         "En groupe : 'Qu'avez-vous\nressenti ? Qu'est-ce qui vous\na le plus touché ?'",
         "Mission UNICEF · KPI FIDELIS\nNomenclature DON/INDÉCIS/REFUS\nOrganisation de la journée",
         "Vous expliquez la mission\nUNICEF en 60 secondes\nà votre binôme — sans notes."],
        ["IMMERSION", "ÉCHANGE", "APPORT", "MISE EN PRATIQUE"],
        y=0.88, h=5.1)
    kpi_row(s,
        ["Pitch mission 60s validé",
         "Quiz 5 questions UNICEF (seuil 60%)",
         "3 lignes dans votre livret personnel"],
        y=6.12, label="Vos 3 réussites à valider aujourd'hui")

    # Slide 3 — UNICEF : la cause
    s = prs.slides.add_slide(SL(prs))
    header_content(s, "L'UNICEF — La Cause que Vous Défendez par Téléphone")
    facts = [
        ("1946",    "Fondation de\nl'UNICEF (ONU)"),
        ("190",     "pays où l'UNICEF\nest présent"),
        ("45%",     "des enfants du monde\nvaccinés par l'UNICEF"),
        ("92g",     "poids d'1 sachet RUTF\ntraitement malnutrition"),
        ("3 princ.","Humanité\nImpartialité\nNeutralité"),
        ("Loi 1901","Cadre juridique\nassociatif français"),
    ]
    x = 0.3
    for val, label in facts:
        r(s, x, 0.88, 2.1, 1.05, fill=TEAL_D)
        t(s, val, x+0.1, 0.9, 1.9, 0.62, sz=22, bold=True,
          color=WHITE, align=PP_ALIGN.CENTER)
        t(s, label, x+0.1, 1.5, 1.9, 0.4, sz=9,
          color=RGBColor(0xCC,0xEE,0xEE), align=PP_ALIGN.CENTER, wrap=True)
        x += 2.16
    t(s, "La cause centrale de la campagne :",
      0.3, 2.12, 12.73, 0.3, sz=12, bold=True, color=TEAL_D)
    r(s, 0.3, 2.48, 0.12, 0.82, fill=TEAL)
    highlight(s,
        "La malnutrition aigüe sévère — 1 enfant meurt de faim toutes les 11 secondes dans le monde.\n"
        "Les sachets RUTF (92g) permettent à un enfant sévèrement malnutri de guérir en 6 à 8 semaines.",
        0.44, 2.48, 12.59, 0.82, fill=TEAL_L, tc=DARK)
    t(s, "Les KPI qui structurent votre journée de production :",
      0.3, 3.5, 12.73, 0.3, sz=12, bold=True, color=TEAL_D)
    kpis_j1 = [
        ("CU/H ≥ 9", "Contacts Utiles\npar heure minimum"),
        ("TX Transfo", "% PEL+PA\nparmi les CU"),
        ("PDC", "Plan de Charge\nmensuel FIDELIS"),
        ("DON", "Contact qui accepte\nle don régulier"),
        ("INDÉCIS", "Contact à relancer\naprès réflexion"),
        ("REFUS", "Contact qui refuse\ndéfinitivement"),
    ]
    x = 0.3
    for val, label in kpis_j1:
        r(s, x, 3.88, 2.1, 1.08, fill=TEAL)
        t(s, val, x+0.1, 3.9, 1.9, 0.55, sz=14, bold=True,
          color=WHITE, align=PP_ALIGN.CENTER)
        t(s, label, x+0.1, 4.44, 1.9, 0.48, sz=9,
          color=RGBColor(0xCC,0xEE,0xEE), align=PP_ALIGN.CENTER, wrap=True)
        x += 2.16
    highlight(s,
        "Organisation journée de production : 9h30–17h30  ·  6h40 de production effective\n"
        "Opérations Conquête (nouveaux prospects) vs Fidélisation / Réactivation (anciens donateurs)",
        0.3, 5.12, 12.73, 0.62)

    # Slide 4 — Posture vocale
    s = prs.slides.add_slide(SL(prs))
    header_content(s, "Votre Posture Vocale — Ce que le Donateur Perçoit en 3 Secondes")
    two_col(s,
        "Ce qui crée la confiance dès l'accroche",
        ["Sourire audible dès le mot 'Bonjour' — le prospect l'entend vraiment",
         "Ton chaleureux et naturel — ni trop commercial, ni trop distant",
         "Débit posé : 130–150 mots/minute — pas de précipitation",
         "Articulation claire — chaque mot compte au téléphone",
         "Conviction sincère — vous croyez en la mission UNICEF"],
        "Ce qui fait raccrocher le donateur",
        ["Voix monocorde ou mécanique — lecture évidente du script",
         "Débit trop rapide — le prospect décroche mentalement",
         "Hésitations répétées : 'euh…', 'donc…', 'voilà…'",
         "Ton défensif face aux premières objections",
         "Manque d'énergie — s'entend même à travers le téléphone"])
    highlight(s,
        "Exercice : Lisez votre script à voix haute en souriant physiquement. "
        "Enregistrez-vous 60 secondes. Réécoutez. Qu'est-ce que vous changeriez ?",
        0.3, 6.02, 12.73, 0.44)

    return save(prs, f"FIM_BookJ1_Fondations_Mission_{VERSION}.pptx")


# ══════════════════════════════════════════════════════════════════════════
#  BOOK J2 — MONDE ASSOCIATIF & HUMANITAIRE
# ══════════════════════════════════════════════════════════════════════════

def build_bookJ2():
    prs = new_prs()

    # Slide 1
    s = prs.slides.add_slide(SL(prs))
    header_title(s,
        "Le Monde Associatif & Humanitaire",
        "Comprendre le secteur pour représenter la cause avec légitimité",
        tag="BOOK J2  ·  Journée 2  ·  France associative & Histoire humanitaire")
    t(s,
      "Aujourd'hui vous comprenez le secteur associatif français, l'histoire de l'humanitaire\n"
      "et pourquoi le don régulier par prélèvement automatique est le modèle qui change tout\n"
      "pour la pérennité des programmes UNICEF.",
      0.5, 2.75, 11.8, 1.1, sz=14, color=DARK)

    # Slide 2 — Séquence J2
    s = prs.slides.add_slide(SL(prs))
    header_content(s, "Journée 2 — Ce que vous allez vivre aujourd'hui")
    seq4(s,
        ["Vous analysez 2 campagnes\nde collecte téléphonique :\nune réussie, une en difficulté.",
         "'Qu'est-ce qui a fait\nla différence ?\nPourquoi l'une a converti ?'",
         "France associative · Loi 1901\nHistoire humanitaire\nDon régulier · Don en Confiance",
         "Simulation : 'Expliquez\nle don régulier à un prospect\nsceptique en 90 secondes.'"],
        ["DÉCOUVERTE", "ANALYSE", "APPORT", "MISE EN PRATIQUE"],
        y=0.88, h=5.1)
    kpi_row(s,
        ["Restitution orale étude de cas (5 min)",
         "Fiche 'Secteur associatif' complétée",
         "Réflexion LFI : légitimité du PA"],
        y=6.12, label="Vos 3 réussites à valider aujourd'hui")

    # Slide 3 — La France associative
    s = prs.slides.add_slide(SL(prs))
    header_content(s, "La France Associative — Votre Cadre de Légitimité")
    facts = [
        ("1,5M",    "associations en\nFrance"),
        ("22M",     "bénévoles en\nFrance"),
        ("1er juil.\n1901", "Loi fondatrice\ndu droit associatif"),
        ("3 types", "De fait / Déclarée /\nReconnue utilité pub."),
        ("Don en\nConfiance", "Label indépendant\nde contrôle et garantie"),
        ("RGPD\nBloctel", "Cadre légal données\nprospects à respecter"),
    ]
    x = 0.3
    for val, label in facts:
        r(s, x, 0.88, 2.1, 1.05, fill=TEAL_D)
        t(s, val, x+0.1, 0.9, 1.9, 0.6, sz=16, bold=True,
          color=WHITE, align=PP_ALIGN.CENTER)
        t(s, label, x+0.1, 1.5, 1.9, 0.4, sz=9,
          color=RGBColor(0xCC,0xEE,0xEE), align=PP_ALIGN.CENTER, wrap=True)
        x += 2.16
    t(s, "Histoire de l'humanitaire — Les jalons que vous devez connaître :",
      0.3, 2.1, 12.73, 0.3, sz=12, bold=True, color=TEAL_D)
    timeline = [
        ("1859", "Bataille de Solférino\n→ Naissance Croix-Rouge"),
        ("1863", "Fondation Croix-Rouge\npar Henry Dunant"),
        ("1946", "Fondation UNICEF\n(ONU - après WWII)"),
        ("1989", "Convention Droits\nde l'Enfant (ONU)"),
        ("1989", "Création Comité\nDon en Confiance"),
        ("2000s","RGPD & Bloctel\nencadrement légal"),
    ]
    x = 0.3
    for year, event in timeline:
        r(s, x, 2.5, 2.1, 0.38, fill=TEAL)
        t(s, year, x+0.1, 2.53, 1.9, 0.28, sz=12, bold=True,
          color=WHITE, align=PP_ALIGN.CENTER)
        r(s, x, 2.88, 2.1, 1.25, fill=TEAL_L, line=TEAL, lw=Pt(0.8))
        t(s, event, x+0.12, 2.94, 1.86, 1.1, sz=10, color=DARK, wrap=True)
        x += 2.16
    highlight(s,
        "Les 3 principes humanitaires fondamentaux que vous incarnez chaque jour :\n"
        "HUMANITÉ — Soulager la souffrance humaine sans discrimination  ·  "
        "IMPARTIALITÉ — Sans distinction de nationalité, race, religion  ·  "
        "NEUTRALITÉ — Aucun parti pris politique",
        0.3, 4.3, 12.73, 0.72)
    highlight(s,
        "Le don régulier par PA : Valeur à vie du donateur régulier = 8× celle d'un donateur ponctuel — "
        "c'est la colonne vertébrale des programmes UNICEF.",
        0.3, 5.15, 12.73, 0.44)

    return save(prs, f"FIM_BookJ2_MondeAssociatif_{VERSION}.pptx")


# ══════════════════════════════════════════════════════════════════════════
#  BOOK J3 — ANALYSE & SCRIPT OFFICIEL
# ══════════════════════════════════════════════════════════════════════════

def build_bookJ3():
    prs = new_prs()

    # Slide 1
    s = prs.slides.add_slide(SL(prs))
    header_title(s,
        "Analyse & Script Officiel UNICEF",
        "Les 7 étapes · Les 5 règles d'accroche · Les 4 modes de validation",
        tag="BOOK J3  ·  Journée 3  ·  Script UNICEF × FIDELIS")
    t(s,
      "Aujourd'hui vous allez lire, analyser et incarner le script officiel UNICEF.\n"
      "Vous ne le lirez pas mécaniquement — vous allez comprendre l'intention\n"
      "de chaque phrase pour le délivrer avec conviction et naturel.",
      0.5, 2.75, 11.8, 1.1, sz=14, color=DARK)

    # Slide 2 — Séquence J3
    s = prs.slides.add_slide(SL(prs))
    header_content(s, "Journée 3 — Ce que vous allez vivre aujourd'hui")
    seq4(s,
        ["Lecture intégrale du script\nofficiel UNICEF à voix haute.\nPremière impression.",
         "Analyse phrase par phrase :\n'Quelle est l'intention ici ?\nPourquoi ce mot précis ?'",
         "Les 7 étapes · Les 5 règles\nd'accroche · Les 4 modes\nde validation du don",
         "Exercices de lecture en binômes.\nDébriefs collectifs.\nPoints forts & axes d'amélioration."],
        ["DÉCOUVERTE", "ANALYSE", "APPORT", "ENTRAÎNEMENT"],
        y=0.88, h=5.1)
    kpi_row(s,
        ["Script lu 3 fois sans hésitation majeure",
         "5 règles d'accroche récitées sans support",
         "Score seuil : 12/20 pour passage J4"],
        y=6.12, label="Vos 3 réussites à valider aujourd'hui")

    # Slide 3 — Les 7 étapes
    s = prs.slides.add_slide(SL(prs))
    header_content(s, "Le Script Officiel UNICEF — Les 7 Étapes que Vous Allez Maîtriser")
    steps = [
        ("①\nACCROCHE",
         TEAL_D,
         "Ouverture téléphonique conforme\naux 5 règles officielles\n\n"
         "Ton chaleureux, prénom + nom,\norganisation clairement nommée,\nraison de l'appel annoncée"),
        ("②\nMALNUTRITION",
         TEAL,
         "Présentation de la cause :\nla malnutrition aigüe sévère\n\n"
         "Chiffre-clé : 1 enfant toutes\nles 11 secondes · Dimension\nhumaine et émotionnelle"),
        ("③\nSACHETS RUTF",
         TEAL_MID,
         "La solution concrète de l'UNICEF :\nles sachets RUTF (92g)\n\n"
         "6 à 8 semaines pour guérir\nun enfant sévèrement malnutri\nCoût : quelques euros/traitement"),
        ("④\nAPPEL AU SOUTIEN",
         TEAL_ACC,
         "Proposition du don mensuel\npar Prélèvement Automatique\n\n"
         "Montant suggéré adapté\nImpact chiffré du don\nSimplicité du geste"),
        ("⑤\nSI DON PONCTUEL",
         RGBColor(0x33,0x77,0x77),
         "Si le prospect propose\nun don unique, réorienter\nvers le don régulier\n\n"
         "'Je comprends — et si on\nenvisageait quelque chose\nde régulier ?'"),
        ("⑥\nINDÉCIS",
         RGBColor(0x22,0x66,0x66),
         "Gestion de l'hésitation :\nrester dans l'échange,\nne pas forcer la conclusion\n\n"
         "Relancer avec une question\nouverture sur le délai\nde réflexion"),
        ("⑦\nVALIDATION\nCOORDONNÉES",
         RGBColor(0x11,0x55,0x55),
         "Sécurisation du don :\nIBAN à chaud / PA en ligne /\nPEL / PA courrier\n\n"
         "Confirmer toutes\nles coordonnées · SMS\nde confirmation envoyé"),
    ]
    x = 0.3
    sw = (13.33 - 0.6) / 7
    for title, col, body in steps:
        r(s, x, 0.88, sw, 0.5, fill=col)
        t(s, title, x+0.06, 0.9, sw-0.1, 0.46, sz=8, bold=True,
          color=WHITE, align=PP_ALIGN.CENTER)
        r(s, x, 1.38, sw, 4.85, fill=WHITE, line=col, lw=Pt(1.2))
        t(s, body, x+0.08, 1.44, sw-0.14, 4.65, sz=8.5, color=DARK, wrap=True)
        x += sw + 0.02
    highlight(s,
        "Durée standard d'un appel complet : 3 à 5 minutes  ·  "
        "Chaque étape a une fonction précise — ne pas sauter, ne pas inverser.",
        0.3, 6.34, 12.73, 0.38)

    # Slide 4 — Les 5 règles d'accroche
    s = prs.slides.add_slide(SL(prs))
    header_content(s, "Les 5 Règles Officielles de la Phrase d'Accroche")
    t(s, "Ces 5 règles sont non négociables — elles définissent un appel conforme :",
      0.35, 0.88, 12.5, 0.32, sz=12, bold=True, color=TEAL_D)
    rules = [
        ("RÈGLE 1", "Se présenter avec prénom et nom",
         "Votre identité est donnée clairement dès les premières secondes.\n"
         "'Bonjour, je m'appelle [Prénom NOM]…'"),
        ("RÈGLE 2", "Nommer l'organisation représentée",
         "'…je vous appelle au nom de l'UNICEF France…'\n"
         "Le prospect sait immédiatement qui appelle."),
        ("RÈGLE 3", "Annoncer la raison de l'appel",
         "'…pour vous parler d'une initiative importante pour les enfants…'\n"
         "Clarté totale — pas d'ambiguïté sur l'objet de l'appel."),
        ("RÈGLE 4", "Vérifier la disponibilité du prospect",
         "'Est-ce que vous avez quelques minutes ?'\n"
         "Respect de l'interlocuteur — condition d'un échange de qualité."),
        ("RÈGLE 5", "Adopter un ton chaleureux et naturel",
         "Sourire audible · Débit posé · Articulation claire\n"
         "Le ton est aussi important que les mots — souvent plus."),
    ]
    y = 1.28
    for i, (tag, title, body) in enumerate(rules):
        bg = WHITE if i % 2 == 0 else TEAL_L
        r(s, 0.3, y, 1.1, 0.9, fill=TEAL_D)
        t(s, tag, 0.32, y+0.18, 1.06, 0.54, sz=9, bold=True,
          color=WHITE, align=PP_ALIGN.CENTER)
        r(s, 1.42, y, 3.4, 0.9, fill=TEAL_L, line=TEAL_D, lw=Pt(0.8))
        t(s, title, 1.52, y+0.2, 3.22, 0.54, sz=11, bold=True, color=TEAL_D)
        r(s, 4.84, y, 8.19, 0.9, fill=bg)
        t(s, body, 4.96, y+0.1, 7.98, 0.72, sz=10, color=DARK, italic=True)
        y += 0.94
    highlight(s,
        "Ces 5 règles sont vérifiées par la grille d'évaluation FIDELIS. "
        "Un appel qui viole l'une d'elles est disqualifié dans la nomenclature.",
        0.3, 5.98, 12.73, 0.44)

    return save(prs, f"FIM_BookJ3_Lecture_Script_{VERSION}.pptx")


# ══════════════════════════════════════════════════════════════════════════
#  BOOK J4 — TRAITEMENT DES OBJECTIONS
# ══════════════════════════════════════════════════════════════════════════

def build_bookJ4():
    prs = new_prs()

    # Slide 1
    s = prs.slides.add_slide(SL(prs))
    header_title(s,
        "Traitement des Objections",
        "15 objections · 5 catégories · Méthode AAR · Jeux de rôle téléphonique",
        tag="BOOK J4  ·  Journée 4  ·  Objections UNICEF France × FIDELIS")
    t(s,
      "Aujourd'hui vous apprenez à transformer chaque objection en opportunité de dialogue.\n"
      "Un prospect qui objecte est un prospect qui est encore là — et ça compte.\n"
      "Mantra : « Le refus d'aujourd'hui peut être le don de demain »",
      0.5, 2.75, 11.8, 1.1, sz=14, color=DARK)

    # Slide 2 — Séquence J4
    s = prs.slides.add_slide(SL(prs))
    header_content(s, "Journée 4 — Ce que vous allez vivre aujourd'hui")
    seq4(s,
        ["Vous écoutez 5 extraits d'appels :\nbonne et mauvaise gestion\nd'une objection.",
         "'Qu'avez-vous entendu ?\nQu'est-ce qui a tout changé\ndans la réponse ?'",
         "15 objections · 5 catégories\nMéthode AAR : Accuser réception\n/ Argumenter / Relancer",
         "Jeux de rôle en binômes\nsur chaque catégorie.\nGrille d'observation formateur."],
        ["ÉCOUTE", "ANALYSE", "APPORT", "ENTRAÎNEMENT"],
        y=0.88, h=5.1)
    kpi_row(s,
        ["Grille écoute active validée par formateur",
         "Réflexion LFI : 'Quelle objection me challenge le plus ?'",
         "Score seuil : 14/20 pour validation J4"],
        y=6.12, label="Vos 3 réussites à valider aujourd'hui")

    # Slide 3 — Méthode AAR
    s = prs.slides.add_slide(SL(prs))
    header_content(s, "Votre Méthode en 3 Temps — AAR")
    aar = [
        ("A\nACCUSER\nRÉCEPTION",
         TEAL_D, "4.05 in",
         "Montrez que vous avez entendu.\nNe justifiez pas immédiatement.\nNe contre-attaquez pas.\n\n"
         "Exemples :\n'Je comprends tout à fait votre réserve…'\n"
         "'C'est une question que beaucoup de personnes\nse posent et c'est tout à fait légitime…'\n"
         "'Je vous entends…'\n\n"
         "⚠ Ce que vous ne faites pas :\n'Oui mais…' ou 'Non, vous avez tort…'"),
        ("A\nARGUMENTER",
         TEAL, "4.05 in",
         "1 argument précis + 1 preuve concrète\n+ 1 chiffre si possible.\n\n"
         "Règles d'or :\n→ 1 seul argument par réponse (pas de liste)\n"
         "→ Preuve vérifiable (rapport annuel, audit)\n→ Chiffre simple et mémorisable\n\n"
         "Exemple objection financière :\n'10€/mois c'est 33 centimes par jour — "
         "moins\nqu'un café. Et cela traite 1 enfant sévèrement\nmalnutri en 6 semaines.'"),
        ("R\nRELANCER",
         TEAL_MID, "4.05 in",
         "Revenez à la proposition avec naturel.\nPas de pression. Pas d'insistance agressive.\n\n"
         "Exemples :\n'Est-ce que cela répond à votre question ?'\n"
         "'On peut aller plus loin ensemble ?'\n'Qu'est-ce qui vous permettrait de vous sentir\n"
         "à l'aise avec cet engagement ?'\n\n"
         "⚠ Ce que vous ne faites pas :\nRelancer avant que le prospect ait eu\nle temps de traiter votre réponse."),
    ]
    x = 0.3
    for title, col, w_str, body in aar:
        w = 4.17
        r(s, x, 0.88, w, 0.7, fill=col)
        t(s, title, x+0.14, 0.9, w-0.22, 0.66, sz=16, bold=True,
          color=WHITE, align=PP_ALIGN.CENTER)
        r(s, x, 1.58, w, 4.65, fill=WHITE, line=col, lw=Pt(1.8))
        t(s, body, x+0.16, 1.66, w-0.28, 4.45, sz=10, color=DARK, wrap=True)
        x += w + 0.15
    highlight(s,
        "Le silence après votre argument est votre allié. Comptez mentalement 1–2–3 avant de relancer. "
        "Ce silence donne au prospect le temps de traiter — et souvent de se convaincre lui-même.",
        0.3, 6.35, 12.73, 0.44)

    # Slide 4 — Les 5 catégories d'objections
    s = prs.slides.add_slide(SL(prs))
    header_content(s, "Les 5 Catégories d'Objections — 15 Situations à Maîtriser")
    cats = [
        ("CAT. 1\nOBJECTIONS INITIALES",
         TEAL_D,
         "Phase d'accroche — 8 objections :\n"
         "• Pas intéressé\n• Faux numéro\n• Appel pour un don\n• Je donne déjà\n"
         "• Arnaque\n• Je ne donne pas par téléphone\n• Pas confiance aux associations\n"
         "• N'aime pas être contacté"),
        ("CAT. 2\nOBJECTIONS FINANCIÈRES",
         TEAL,
         "Après l'appel au don — 3 objections :\n"
         "• Je donne déjà ailleurs\n• Je n'ai pas les moyens\n• Dernier positionnement financier"),
        ("CAT. 3\nCONTRE LE DON RÉGULIER (PA)",
         TEAL_MID,
         "Sur le prélèvement automatique — 3 objections :\n"
         "• Préfère le don ponctuel\n• N'aime pas l'engagement mensuel\n"
         "• Dernier positionnement PA"),
        ("CAT. 4\nMÉFIANCE IBAN",
         TEAL_ACC,
         "Sur la sécurité bancaire :\n"
         "• Réassurance sur le protocole SEPA\n• Sécurité des données bancaires\n"
         "• Mandat SEPA : droits du donateur\n• Arrêt / modification possible à tout moment"),
        ("CAT. 5\nOBJECTIONS CONFLICTUELLES",
         RGBColor(0x33,0x77,0x77),
         "Objections sensibles :\n"
         "• Origine de l'appel\n• Source du numéro\n• Droits légaux RGPD\n• Bloctel / liste rouge\n"
         "• Accent de l'appelant\n• Connaissance d'un membre\n• Identité de l'appelant"),
    ]
    x = 0.3
    cw = (13.33 - 0.6) / 5
    for i, (title, col, body) in enumerate(cats):
        r(s, x, 0.88, cw, 0.55, fill=col)
        t(s, title, x+0.1, 0.9, cw-0.16, 0.51, sz=8, bold=True,
          color=WHITE, align=PP_ALIGN.CENTER)
        r(s, x, 1.43, cw, 4.78, fill=WHITE, line=col, lw=Pt(1.2))
        t(s, body, x+0.1, 1.5, cw-0.16, 4.58, sz=9, color=DARK, wrap=True)
        x += cw + 0.02
    highlight(s,
        "En production, la majorité des appels se terminent à la Catégorie 1. "
        "Maîtriser ces 8 objections initiales est la clé pour atteindre vos 9 CU/H.",
        0.3, 6.34, 12.73, 0.44)

    return save(prs, f"FIM_BookJ4_Objections_{VERSION}.pptx")


# ══════════════════════════════════════════════════════════════════════════
#  BOOK J5 — SYNTHÈSE & SIMULATIONS
# ══════════════════════════════════════════════════════════════════════════

def build_bookJ5():
    prs = new_prs()

    # Slide 1
    s = prs.slides.add_slide(SL(prs))
    header_title(s,
        "Synthèse Générale & Simulations",
        "Certification FIM · Grille 24 critères · Quiz 40 questions",
        tag="BOOK J5  ·  Journée 5  ·  Certification FIDELIS × UNICEF France × SHY-Performance")
    t(s,
      "C'est votre journée de certification.\n"
      "Vous allez démontrer que vous maîtrisez l'intégralité de la chaîne de compétences FIM :\n"
      "du script officiel UNICEF en 7 étapes au traitement des 15 objections, jusqu'au closing.",
      0.5, 2.75, 11.8, 1.1, sz=14, color=DARK)

    # Slide 2 — Programme J5
    s = prs.slides.add_slide(SL(prs))
    header_content(s, "Votre Journée 5 — Programme de Certification")
    programme = [
        ("08h30", "09h00", "Révision consolidée",
         "Révision rapide des 5 modules FIM · Quiz flash · Météo émotionnelle du groupe"),
        ("09h00", "11h00", "Simulations d'appels complets",
         "Appels simulés chronométrés : script complet + objections + closing en conditions réelles\n"
         "Jeux de rôle fundraiser vs prospect · Rotation des rôles · Grille d'observation"),
        ("11h00", "11h30", "Débriefing collectif",
         "Retour sur les simulations · Célébration des progrès · 3 apprentissages collectifs de la semaine"),
        ("11h30", "12h30", "Grille d'évaluation finale",
         "Évaluation individuelle sur 24 critères · Entretien formateur 5 min · Feedback personnalisé"),
        ("13h30", "14h30", "Quiz 40 questions",
         "Évaluation finale des connaissances : UNICEF, secteur associatif, script, objections, KPI FIDELIS"),
        ("14h30", "15h00", "Remise des certifications",
         "Remise officielle des certifications FIM · Votre plan d'action terrain · Premier appel demain"),
    ]
    y = 0.88
    for start, end, title, detail in programme:
        r(s, 0.3, y, 1.45, 0.7, fill=TEAL_D)
        t(s, f"{start}\n{end}", 0.32, y+0.08, 1.41, 0.56,
          sz=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        r(s, 1.77, y, 3.0, 0.7, fill=TEAL)
        t(s, title, 1.88, y+0.15, 2.82, 0.42, sz=12, bold=True, color=WHITE)
        r(s, 4.79, y, 8.24, 0.7, fill=WHITE, line=TEAL, lw=Pt(0.8))
        t(s, detail, 4.92, y+0.06, 8.02, 0.6, sz=9, color=DARK, wrap=True)
        y += 0.73

    # Slide 3 — Grille 24 critères
    s = prs.slides.add_slide(SL(prs))
    header_content(s, "Votre Grille d'Évaluation Finale — 24 Critères")
    t(s, "Voici exactement ce sur quoi vous serez évalué(e) — aucune surprise :",
      0.35, 0.85, 12.5, 0.3, sz=11, bold=True, color=TEAL_D)
    criteria = [
        ("A1","Respect des 5 règles d'accroche",                "Accroche"),
        ("A2","Présentation claire de l'organisation",           "Accroche"),
        ("A3","Ton chaleureux et naturel dès l'ouverture",       "Accroche"),
        ("A4","Présentation de la cause malnutrition",           "Script"),
        ("A5","Explication des sachets RUTF avec précision",     "Script"),
        ("A6","Appel au soutien clair avec montant suggéré",     "Script"),
        ("A7","Gestion du don ponctuel → réorientation PA",      "Script"),
        ("A8","Gestion de l'indécis — relance constructive",     "Script"),
        ("A9","Validation coordonnées — mode de paiement choisi","Closing"),
        ("B1","Accusé réception sincère des objections Cat.1",   "Objections"),
        ("B2","Argument financier précis + preuve concrète",     "Objections"),
        ("B3","Réorientation PA après objection ponctuel",       "Objections"),
        ("B4","Réassurance IBAN / SEPA convaincante",            "Objections"),
        ("B5","Gestion calme des objections conflictuelles",     "Objections"),
        ("C1","Ton maintenu chaleureux sur toute la durée",      "Posture"),
        ("C2","Débit adapté — ni trop rapide, ni trop lent",     "Posture"),
        ("C3","Absence d'hésitations et de tics verbaux",        "Posture"),
        ("C4","Conviction sincère pour la mission UNICEF",       "Posture"),
        ("C5","Résilience après un refus direct",                "Posture"),
        ("C6","Respect des règles éthiques et RGPD",             "Éthique"),
        ("D1","Qualification conforme FIDELIS (DON/IND/REFUS)",  "FIDELIS"),
        ("D2","Nomination correcte du mode de validation",       "FIDELIS"),
        ("D3","Durée d'appel dans la norme (3–5 min)",           "FIDELIS"),
        ("D4","Saisie coordonnées complète et exacte",           "FIDELIS"),
    ]
    col_x = [0.3, 0.76, 7.0, 9.9,  10.7, 11.5, 12.3]
    col_w = [0.43, 6.22, 2.87, 0.78, 0.78, 0.78, 0.75]
    headers = ["#", "Ce que vous devez démontrer", "Catégorie", "1", "2", "3", "4"]
    y = 1.2
    table_header(s, headers, col_x, col_w, y)
    y += 0.32
    rh = 0.36
    cat_colors = {
        "Accroche":   TEAL_D,
        "Script":     TEAL,
        "Closing":    TEAL_MID,
        "Objections": TEAL_ACC,
        "Posture":    RGBColor(0x33,0x77,0x77),
        "Éthique":    RGBColor(0x22,0x66,0x66),
        "FIDELIS":    RGBColor(0x11,0x55,0x55),
    }
    for i, (code, crit, cat) in enumerate(criteria):
        bg = WHITE if i % 2 == 0 else TEAL_L
        r(s, col_x[0], y, col_w[0], rh, fill=TEAL_D)
        t(s, code, col_x[0]+0.04, y+0.06, col_w[0]-0.06, rh-0.1,
          sz=8, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        r(s, col_x[1], y, col_w[1], rh, fill=bg)
        t(s, crit, col_x[1]+0.08, y+0.06, col_w[1]-0.12, rh-0.08, sz=9, color=DARK)
        cc = cat_colors.get(cat, TEAL)
        r(s, col_x[2], y, col_w[2], rh, fill=cc)
        t(s, cat, col_x[2]+0.06, y+0.06, col_w[2]-0.1, rh-0.08,
          sz=8, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        for k, (x, w) in enumerate(zip(col_x[3:], col_w[3:])):
            r(s, x, y, w, rh, fill=WHITE, line=TEAL, lw=Pt(0.4))
        y += rh + 0.01
    # Total
    r(s, col_x[0], y, sum(col_w[:3])+0.04, 0.34, fill=TEAL_D)
    t(s, "TOTAL  /24", col_x[0]+0.12, y+0.06, 9.2, 0.25,
      sz=10, bold=True, color=WHITE)
    r(s, col_x[3], y, sum(col_w[3:])+0.04, 0.34,
      fill=TEAL_L, line=TEAL_D, lw=Pt(2))
    t(s, "__ / 24", col_x[3]+0.1, y+0.06, 2.4, 0.24,
      sz=11, bold=True, color=TEAL_D, align=PP_ALIGN.CENTER)

    # Slide 4 — Quiz & Plan d'action
    s = prs.slides.add_slide(SL(prs))
    header_content(s, "Quiz 40 Questions & Votre Plan d'Action Terrain")
    two_col(s,
        "Quiz final — 40 questions · 5 parties",
        ["Partie A (Q1–Q8) : Connaissance UNICEF & de la cause",
         "Partie B (Q9–Q15) : Secteur associatif & humanitaire",
         "Partie C (Q16–Q23) : Script officiel & règles d'accroche",
         "Partie D (Q24–Q32) : Traitement des objections & méthode AAR",
         "Partie E (Q33–Q40) : KPI FIDELIS, nomenclature & modes de validation",
         "Seuil de validation : 28/40 (70%)"],
        "Votre plan d'action — 4 semaines terrain",
        ["Semaine 1 (J+7) : Objectif 9 CU/H atteint et stable",
         "Semaine 2 (J+14) : TX Transfo ≥ cible FIDELIS première mesure",
         "Semaine 3 (J+21) : 0 disqualification nomenclature FIDELIS",
         "Semaine 4 (J+30) : Bilan formateur SHY-Performance + objectifs M2",
         "J+15 : Observation supervisée par SHY-Performance",
         "J+90 : Analyse KPI terrain consolidée"])
    highlight(s,
        "Bienvenue dans la famille SHY-Performance × FIDELIS × UNICEF France.\n"
        "Vous êtes maintenant Fundraiser Impact Mission certifié(e). "
        "Votre premier appel de production commence demain.",
        0.3, 6.02, 12.73, 0.62)

    return save(prs, f"FIM_BookJ5_Synthese_Simulations_{VERSION}.pptx")


# ══════════════════════════════════════════════════════════════════════════
#  MAIN
# ══════════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    print(f"\n🎨 FIM {VERSION} — Charte SHY-Performance · Prospection téléphonique UNICEF France\n")
    files = []
    files.append(build_book0())
    files.append(build_book1())
    files.append(build_bookJ1())
    files.append(build_bookJ2())
    files.append(build_bookJ3())
    files.append(build_bookJ4())
    files.append(build_bookJ5())
    print(f"\n✅ {len(files)} modules générés dans {OUT_DIR}/")

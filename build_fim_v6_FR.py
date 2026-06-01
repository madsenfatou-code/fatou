"""
FIM V6 — Générateur PPTX — 7 modules
Charte SHY-Performance · Donateurs français · UNICEF France & Monde
Slides adressées aux apprenants · Méthodes pédagogiques appliquées (non affichées)
Version : V6-01062026SHY-TE
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
import os

# ─── CHARTE GRAPHIQUE SHY-PERFORMANCE ──────────────────────────────────────
SHY_TEAL    = RGBColor(0x00, 0x80, 0x80)   # Vert SHY principal
SHY_TEAL_D  = RGBColor(0x00, 0x55, 0x55)   # Vert foncé
SHY_RED     = RGBColor(0xFF, 0x00, 0x00)   # Rouge SHY
SHY_WHITE   = RGBColor(0xFF, 0xFF, 0xFF)
SHY_DARK    = RGBColor(0x1A, 0x1A, 0x1A)
SHY_LIGHT   = RGBColor(0xF0, 0xF7, 0xF7)
SHY_GRAY    = RGBColor(0x55, 0x55, 0x55)
SHY_LGRAY   = RGBColor(0xEE, 0xEE, 0xEE)

OUT_DIR = "/home/user/fatou"
VERSION = "V6-01062026SHY-TE"
COMPANY = "SHY-Performance"

# ─── PRIMITIVES ─────────────────────────────────────────────────────────────

def new_prs():
    prs = Presentation()
    prs.slide_width  = Inches(13.33)
    prs.slide_height = Inches(7.5)
    return prs

def SL(prs): return prs.slide_layouts[6]

def rect(slide, x, y, w, h, fill=None, line=None, lw=Pt(0)):
    from pptx.util import Emu
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

def tb(slide, text, x, y, w, h,
       sz=14, bold=False, color=None, align=PP_ALIGN.LEFT,
       italic=False, wrap=True):
    color = color or SHY_DARK
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    box.word_wrap = wrap
    tf = box.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = text
    r.font.size = Pt(sz)
    r.font.bold = bold
    r.font.italic = italic
    r.font.color.rgb = color
    return box

def add_p(tf, text, sz=13, bold=False, color=None, align=PP_ALIGN.LEFT,
          bullet=False, italic=False):
    color = color or SHY_DARK
    p = tf.add_paragraph()
    p.alignment = align
    r = p.add_run()
    r.text = ("• " if bullet else "") + text
    r.font.size = Pt(sz)
    r.font.bold = bold
    r.font.italic = italic
    r.font.color.rgb = color

# ─── HEADER SHY-PERFORMANCE ─────────────────────────────────────────────────

def shy_header(slide, title, subtitle="", tag=""):
    """Entête SHY-Performance : bande teal + trait rouge + logo texte"""
    # Fond général
    rect(slide, 0, 0, 13.33, 7.5, fill=SHY_LIGHT)
    # Bandeau supérieur teal
    rect(slide, 0, 0, 13.33, 2.5, fill=SHY_TEAL_D)
    # Trait rouge signature SHY
    rect(slide, 0, 2.48, 13.33, 0.08, fill=SHY_RED)
    # Bloc logo SHY à gauche
    rect(slide, 0.3, 0.12, 3.2, 0.9, fill=SHY_TEAL)
    tb(slide, "SHY", 0.35, 0.13, 1.2, 0.88, sz=36, bold=True, color=SHY_WHITE)
    tb(slide, "Performance", 1.5, 0.38, 2.0, 0.45, sz=14, bold=True,
       color=SHY_WHITE)
    # Trait séparateur vertical
    rect(slide, 3.65, 0.18, 0.04, 0.72, fill=SHY_RED)
    # Tag module
    if tag:
        tb(slide, tag, 3.8, 0.18, 9.3, 0.38, sz=9, bold=True,
           color=RGBColor(0xBB,0xEE,0xEE))
    # Titre principal
    tb(slide, title, 0.35, 1.05, 12.6, 1.25, sz=30, bold=True, color=SHY_WHITE)
    # Sous-titre
    if subtitle:
        tb(slide, subtitle, 0.35, 2.1, 12.6, 0.38, sz=12,
           color=RGBColor(0xCC,0xEE,0xEE), italic=True)
    # Pied de page
    rect(slide, 0, 7.1, 13.33, 0.4, fill=SHY_TEAL_D)
    tb(slide, f"{COMPANY}  ·  Formation FIM  ·  {VERSION}",
       0, 7.13, 13.33, 0.3, sz=9, color=SHY_WHITE, align=PP_ALIGN.CENTER)

def shy_slide_header(slide, title):
    """Entête compact pour slides de contenu"""
    rect(slide, 0, 0, 13.33, 7.5, fill=SHY_LIGHT)
    rect(slide, 0, 0, 13.33, 0.72, fill=SHY_TEAL_D)
    rect(slide, 0, 0.70, 13.33, 0.06, fill=SHY_RED)
    # Mini logo
    tb(slide, "SHY", 0.18, 0.07, 0.75, 0.58, sz=20, bold=True, color=SHY_WHITE)
    rect(slide, 1.0, 0.1, 0.04, 0.52, fill=SHY_RED)
    tb(slide, title, 1.12, 0.1, 11.9, 0.52, sz=17, bold=True, color=SHY_WHITE)
    rect(slide, 0, 7.1, 13.33, 0.4, fill=SHY_TEAL_D)
    tb(slide, f"{COMPANY}  ·  Formation FIM  ·  {VERSION}",
       0, 7.13, 13.33, 0.3, sz=9, color=SHY_WHITE, align=PP_ALIGN.CENTER)

# ─── COMPOSANTS RÉUTILISABLES ────────────────────────────────────────────────

def two_col(slide, lt, left_items, rt, right_items, y=0.9):
    lx, rx, cw = 0.3, 6.95, 6.2
    # Titres colonnes
    rect(slide, lx, y, cw, 0.36, fill=SHY_TEAL)
    tb(slide, lt, lx+0.12, y+0.05, cw-0.2, 0.28, sz=11, bold=True, color=SHY_WHITE)
    rect(slide, rx, y, cw, 0.36, fill=SHY_TEAL)
    tb(slide, rt, rx+0.12, y+0.05, cw-0.2, 0.28, sz=11, bold=True, color=SHY_WHITE)
    ly = ry = y + 0.42
    for item in left_items:
        rect(slide, lx, ly, cw, 0.44,
             fill=SHY_WHITE if int((ly*10)%10) == 0 else SHY_LIGHT,
             line=SHY_TEAL, lw=Pt(0.4))
        tb(slide, "• " + item, lx+0.14, ly+0.06, cw-0.25, 0.36, sz=10, color=SHY_DARK)
        ly += 0.46
    for item in right_items:
        rect(slide, rx, ry, cw, 0.44,
             fill=SHY_WHITE if int((ry*10)%10) == 0 else SHY_LIGHT,
             line=SHY_TEAL, lw=Pt(0.4))
        tb(slide, "• " + item, rx+0.14, ry+0.06, cw-0.25, 0.36, sz=10, color=SHY_DARK)
        ry += 0.46

def seq_box(slide, steps, labels, colors, y=4.5, h=1.85):
    """Séquence pédagogique horizontale sans libellé méthodologique"""
    n = len(steps)
    sw = (13.33 - 0.6) / n
    for i, (step, label, color) in enumerate(zip(steps, labels, colors)):
        bx = 0.3 + i * (sw + 0.02)
        rect(slide, bx, y, sw, 0.3, fill=color)
        tb(slide, label, bx+0.08, y+0.04, sw-0.16, 0.24,
           sz=8, bold=True, color=SHY_WHITE)
        rect(slide, bx, y+0.3, sw, h-0.3,
             fill=SHY_WHITE, line=color, lw=Pt(1))
        tb(slide, step, bx+0.12, y+0.35, sw-0.24, h-0.5,
           sz=9, color=SHY_DARK, wrap=True)

def highlight_box(slide, text, x, y, w, h, fill=None, text_color=None):
    fill = fill or SHY_TEAL
    text_color = text_color or SHY_WHITE
    rect(slide, x, y, w, h, fill=fill)
    tb(slide, text, x+0.15, y+0.08, w-0.3, h-0.16,
       sz=10, color=text_color, wrap=True)

def checklist_bar(slide, items, y=6.1, label="Pour vous — À retenir"):
    rect(slide, 0.3, y, 12.73, 0.3, fill=SHY_RED)
    tb(slide, label, 0.42, y+0.04, 12.5, 0.24,
       sz=9, bold=True, color=SHY_WHITE)
    tx = 0.3
    w = 12.73 / len(items)
    for item in items:
        rect(slide, tx, y+0.32, w-0.04, 0.55,
             fill=SHY_WHITE, line=SHY_RED, lw=Pt(0.8))
        tb(slide, "✔ " + item, tx+0.1, y+0.36, w-0.18, 0.46,
           sz=9, color=SHY_DARK, wrap=True)
        tx += w

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

    # ── Slide 1 : Titre ──────────────────────────────────────────────────
    s = prs.slides.add_slide(SL(prs))
    shy_header(s,
        "Bienvenue dans la Formation FIM",
        "Fundraiser Impact Mission — Votre parcours de 5 jours",
        tag="BOOK 0  ·  Fil Conducteur  ·  Document de parcours")
    tb(s,
        "Ce livret est votre guide pour les 5 jours de formation.\n"
        "Il contient votre planning, vos engagements et votre tableau de bord personnel.",
        0.5, 2.75, 12.3, 1.0, sz=14, color=SHY_DARK)
    rect(s, 0.5, 3.85, 12.3, 0.06, fill=SHY_RED)
    tb(s,
        "Vous allez acquérir les compétences pour collecter des dons réguliers\n"
        "au nom de l'UNICEF auprès de donateurs français engagés.",
        0.5, 4.05, 12.3, 0.9, sz=13, color=SHY_TEAL_D, bold=True)

    # ── Slide 2 : Vos 5 jours en un coup d'œil ───────────────────────────
    s = prs.slides.add_slide(SL(prs))
    shy_slide_header(s, "Vos 5 Journées — Ce qui vous attend")
    days = [
        ("JOUR 1", "Fondations\n& Mission",
         "Vous découvrez votre rôle,\nles valeurs UNICEF\net votre posture."),
        ("JOUR 2", "Le Monde\nAssociatif",
         "Vous comprenez le don\nrégulier et l'impact\nconcret de chaque euro."),
        ("JOUR 3", "Votre Script\nde Présentation",
         "Vous maîtrisez les\n5 étapes pour convaincre\navec authenticité."),
        ("JOUR 4", "Gérer les\nObjections",
         "Vous transformez chaque\nrefus en opportunité\nd'engager le donateur."),
        ("JOUR 5", "Certification\n& Terrain",
         "Vous passez votre\ncertification FIM\net construisez votre plan."),
    ]
    x = 0.3
    for tag, title, desc in days:
        rect(s, x, 0.88, 2.5, 0.38, fill=SHY_RED)
        tb(s, tag, x+0.08, 0.91, 2.34, 0.3, sz=11, bold=True, color=SHY_WHITE)
        rect(s, x, 1.26, 2.5, 1.1, fill=SHY_TEAL_D)
        tb(s, title, x+0.12, 1.3, 2.26, 1.0, sz=13, bold=True, color=SHY_WHITE,
           align=PP_ALIGN.CENTER)
        rect(s, x, 2.36, 2.5, 3.55,
             fill=SHY_WHITE, line=SHY_TEAL, lw=Pt(1.5))
        tb(s, desc, x+0.15, 2.44, 2.2, 3.2, sz=11, color=SHY_DARK, wrap=True)
        x += 2.58
    highlight_box(s,
        "35 heures de formation intensive + une journée terrain encadrée — "
        "Ratio pratique : 65%  ·  Théorie : 35%",
        0.3, 6.05, 12.73, 0.42,
        fill=SHY_TEAL, text_color=SHY_WHITE)

    # ── Slide 3 : Votre contrat d'engagement ─────────────────────────────
    s = prs.slides.add_slide(SL(prs))
    shy_slide_header(s, "Votre Contrat d'Engagement — Ce que vous vous engagez à faire")
    two_col(s,
        "Vos engagements en tant qu'apprenant",
        ["Participer activement à toutes les mises en situation",
         "Compléter votre livret personnel chaque soir (5 min)",
         "Demander du feedback précis après chaque exercice",
         "Respecter la confidentialité des échanges du groupe",
         "Arriver à l'heure — le groupe démarre ensemble"],
        "Ce que SHY-Performance s'engage à vous offrir",
        ["Un feedback individuel personnalisé chaque jour",
         "Des exercices pratiques progressifs et sécurisants",
         "Un formateur qui adapte le rythme à votre profil",
         "Un suivi après la formation à J+15 et J+30",
         "Une certification reconnue par les partenaires FIM"])
    highlight_box(s,
        "Droit à l'erreur garanti — Chaque simulation ratée est un pas vers la maîtrise. "
        "Ici, l'erreur n'est pas une faute, c'est une information.",
        0.3, 6.02, 12.73, 0.44,
        fill=SHY_RED, text_color=SHY_WHITE)

    # ── Slide 4 : Votre livret personnel ─────────────────────────────────
    s = prs.slides.add_slide(SL(prs))
    shy_slide_header(s, "Votre Livret Personnel — Votre journal de bord FIM")
    tb(s, "Chaque soir, prenez 5 minutes pour compléter ces 4 points :",
       0.35, 0.88, 12.6, 0.38, sz=13, bold=True, color=SHY_TEAL_D)
    entries = [
        ("① Ce que j'ai réussi aujourd'hui",
         "Notez 1 ou 2 moments concrets où vous avez senti la progression.\n"
         "Ex : 'J'ai réussi mon accroche en moins de 12 secondes.'"),
        ("② Ce qui m'a challengé",
         "Décrivez une situation difficile et pourquoi elle vous a bloqué.\n"
         "Ex : 'Quand le passant a dit \"je n'ai pas les moyens\", j'ai hésité.'"),
        ("③ Mon insight du jour",
         "L'apprentissage que vous n'auriez pas pu anticiper en arrivant ce matin.\n"
         "Ex : 'Le silence après la proposition est plus puissant qu'une relance.'"),
        ("④ Mon engagement pour demain",
         "1 comportement précis à changer ou renforcer dès demain matin.\n"
         "Ex : 'Je vais maintenir le contact visuel pendant toute la proposition.'"),
    ]
    y = 1.35
    for i, (title, body) in enumerate(entries):
        bg = SHY_WHITE if i % 2 == 0 else SHY_LIGHT
        rect(s, 0.3, y, 0.6, 1.0, fill=SHY_RED)
        tb(s, str(i+1), 0.32, y+0.22, 0.56, 0.56, sz=22, bold=True,
           color=SHY_WHITE, align=PP_ALIGN.CENTER)
        rect(s, 0.92, y, 12.1, 1.0, fill=bg, line=SHY_TEAL, lw=Pt(0.5))
        tb(s, title, 1.04, y+0.06, 11.9, 0.3, sz=11, bold=True, color=SHY_TEAL_D)
        tb(s, body, 1.04, y+0.38, 11.9, 0.56, sz=10, color=SHY_DARK, italic=True)
        y += 1.05
    highlight_box(s,
        "Votre livret reste strictement personnel — il ne sera pas noté. "
        "C'est votre outil de progression, pas une évaluation.",
        0.3, 5.55, 12.73, 0.42, fill=SHY_TEAL, text_color=SHY_WHITE)

    return save(prs, f"FIM_Book0_FilConducteur_{VERSION}.pptx")


# ══════════════════════════════════════════════════════════════════════════
#  BOOK 1 — SYNOPSIS & OBJECTIFS
# ══════════════════════════════════════════════════════════════════════════

def build_book1():
    prs = new_prs()

    # ── Slide 1 : Titre ──────────────────────────────────────────────────
    s = prs.slides.add_slide(SL(prs))
    shy_header(s,
        "Votre Feuille de Route",
        "Ce que vous saurez faire à la fin de cette formation",
        tag="BOOK 1  ·  Synopsis & Objectifs")
    tb(s,
        "À l'issue de ces 5 jours, vous serez capable de convaincre un donateur français\n"
        "de s'engager dans un don mensuel régulier au profit de l'UNICEF.",
        0.5, 2.75, 12.3, 1.0, sz=14, color=SHY_DARK)

    # ── Slide 2 : Vos 7 compétences ───────────────────────────────────────
    s = prs.slides.add_slide(SL(prs))
    shy_slide_header(s, "Les 7 Compétences que Vous Allez Développer")
    comps = [
        ("C1", "Accroche",         "Initier un contact naturel en moins de 15 secondes"),
        ("C2", "Impact Story",     "Raconter une histoire vraie qui touche et convainc"),
        ("C3", "Écoute active",    "Lire les signaux du donateur et ajuster votre discours"),
        ("C4", "Objections",       "Transformer un refus en opportunité d'engagement"),
        ("C5", "Engagement",       "Conclure un don régulier mensuel avec confiance"),
        ("C6", "Résilience",       "Rester positif et efficace après dix refus consécutifs"),
        ("C7", "Éthique",          "Respecter les règles déontologiques en toute situation"),
    ]
    y = 0.88
    alt = [SHY_TEAL_D, SHY_TEAL]
    for i, (code, name, desc) in enumerate(comps):
        bg = alt[i % 2]
        rect(s, 0.3,  y, 0.9,  0.54, fill=bg)
        tb(s, code, 0.32, y+0.1, 0.86, 0.36, sz=14, bold=True,
           color=SHY_WHITE, align=PP_ALIGN.CENTER)
        rect(s, 1.22, y, 3.2,  0.54, fill=SHY_LIGHT, line=bg, lw=Pt(0.8))
        tb(s, name, 1.32, y+0.1, 3.0, 0.36, sz=12, bold=True, color=bg)
        rect(s, 4.44, y, 8.59, 0.54, fill=SHY_WHITE, line=bg, lw=Pt(0.5))
        tb(s, desc, 4.56, y+0.1, 8.36, 0.36, sz=11, color=SHY_DARK, italic=True)
        y += 0.57
    highlight_box(s,
        "Votre objectif : obtenir la certification FIM avec un score ≥ 20/28 en Jour 5.",
        0.3, 5.0, 12.73, 0.38, fill=SHY_RED, text_color=SHY_WHITE)

    # ── Slide 3 : Ce que vous allez accomplir ─────────────────────────────
    s = prs.slides.add_slide(SL(prs))
    shy_slide_header(s, "Ce que Vous Allez Accomplir — Vos Résultats Concrets")
    two_col(s,
        "En fin de formation, vous saurez…",
        ["Accrocher un passant en moins de 15 secondes",
         "Raconter l'histoire d'un enfant bénéficiaire UNICEF",
         "Proposer un montant de don adapté au profil",
         "Répondre aux 12 objections types sans hésiter",
         "Signer un engagement de don mensuel régulier"],
        "Vous serez également capable de…",
        ["Gérer un refus sans perdre votre motivation",
         "Adapter votre script à chaque donateur rencontré",
         "Analyser votre propre performance et vous corriger",
         "Respecter la déontologie fundraiser en toute situation",
         "Construire votre plan d'action personnel pour J+30"])
    checklist_bar(s,
        ["Dès J1 : posture et pitch 60s",
         "Dès J3 : script complet maîtrisé",
         "Dès J4 : objections gérées",
         "J5 : certification obtenue"],
        y=6.0, label="Vos jalons de progression")

    # ── Slide 4 : Votre progression ───────────────────────────────────────
    s = prs.slides.add_slide(SL(prs))
    shy_slide_header(s, "Votre Progression — De Débutant à Fundraiser Certifié")
    stages = [
        ("Avant la\nformation",
         SHY_LGRAY, SHY_GRAY,
         "Vous connaissez peu le\nfundraising de rue.\nVous appréhendez peut-être\nle premier contact."),
        ("Après J1–J2",
         SHY_LIGHT, SHY_TEAL,
         "Vous comprenez la cause\net le modèle du don.\nVous êtes prêt(e) à\nconstruire votre discours."),
        ("Après J3–J4",
         RGBColor(0xCC,0xEE,0xCC), RGBColor(0x1A,0x7A,0x1A),
         "Vous maîtrisez le script\net gérez les objections.\nVous vous sentez à l'aise\nen simulation."),
        ("Après J5\n(certifié)",
         RGBColor(0xFF,0xCC,0xCC), SHY_RED,
         "Vous avez obtenu votre\ncertification FIM.\nVous êtes opérationnel(le)\ndès demain sur le terrain."),
    ]
    x = 0.3
    for stage, bg, col, body in stages:
        rect(s, x, 0.88, 3.1, 0.55, fill=col)
        tb(s, stage, x+0.12, 0.9, 2.86, 0.51, sz=11, bold=True,
           color=SHY_WHITE, align=PP_ALIGN.CENTER)
        rect(s, x, 1.43, 3.1, 4.1, fill=bg, line=col, lw=Pt(1.5))
        tb(s, body, x+0.15, 1.55, 2.8, 3.8, sz=12, color=SHY_DARK, wrap=True)
        if x < 9.9:
            tb(s, "→", x+3.12, 2.7, 0.3, 0.8, sz=22, bold=True, color=SHY_TEAL)
        x += 3.26
    highlight_box(s,
        "La progression est individuelle. Votre formateur adapte son accompagnement à votre rythme.",
        0.3, 5.75, 12.73, 0.42, fill=SHY_TEAL, text_color=SHY_WHITE)

    return save(prs, f"FIM_Book1_Synopsis_Objectifs_{VERSION}.pptx")


# ══════════════════════════════════════════════════════════════════════════
#  BOOK J1 — FONDATIONS & MISSION
# ══════════════════════════════════════════════════════════════════════════

def build_bookJ1():
    prs = new_prs()

    # ── Slide 1 : Titre ──────────────────────────────────────────────────
    s = prs.slides.add_slide(SL(prs))
    shy_header(s,
        "Fondations & Mission",
        "Qui êtes-vous ? Pourquoi êtes-vous là ? Quelle est votre mission ?",
        tag="BOOK J1  ·  Journée 1  ·  Identité & Valeurs")
    tb(s,
        "Aujourd'hui, vous allez ancrer votre identité de fundraiser.\n"
        "Vous comprendrez la cause que vous défendez, les valeurs qui vous guident\n"
        "et la posture qui fait toute la différence sur le terrain.",
        0.5, 2.75, 12.3, 1.2, sz=14, color=SHY_DARK)

    # ── Slide 2 : Séquence de travail J1 ─────────────────────────────────
    s = prs.slides.add_slide(SL(prs))
    shy_slide_header(s, "Journée 1 — Ce que vous allez vivre aujourd'hui")
    seq_box(s,
        ["Vous regardez le témoignage\nd'un enfant bénéficiaire.\nNotez ce que vous ressentez.",
         "En groupe : 'Qu'avez-vous\nressenti ? Qu'auriez-vous\ndit à cet instant ?'",
         "L'UNICEF en chiffres.\nLa cause. Le cadre.\nVos droits et devoirs.",
         "Vous présentez la mission\nUNICEF en 60 secondes\nà votre binôme."],
        ["DÉCOUVERTE", "ÉCHANGE", "APPORT", "MISE EN PRATIQUE"],
        [SHY_TEAL_D, SHY_TEAL, RGBColor(0x2A,0x7A,0x50), SHY_RED],
        y=0.88, h=4.5)
    checklist_bar(s,
        ["Pitch 60s évalué avec votre formateur",
         "Quiz 5 questions UNICEF (seuil 60%)",
         "3 lignes dans votre livret personnel"],
        y=5.55, label="Vos 3 réussites à valider aujourd'hui")

    # ── Slide 3 : UNICEF — La cause que vous défendez ────────────────────
    s = prs.slides.add_slide(SL(prs))
    shy_slide_header(s, "L'UNICEF — La cause que vous défendez")
    facts = [
        ("190+",  "pays où l'UNICEF\nest présent"),
        ("200M+", "enfants aidés\nchaque année"),
        ("830M€", "collectés en France\nen 2024"),
        ("72%",   "des fonds viennent\ndes dons privés"),
        ("3€/j",  "coût d'1 journée\nd'alimentation enfant"),
        ("82%",   "des dons affectés\ndirectement aux programmes"),
    ]
    x = 0.3
    for num, label in facts:
        rect(s, x, 0.88, 2.1, 1.08, fill=SHY_TEAL_D)
        tb(s, num, x+0.1, 0.9, 1.9, 0.65, sz=26, bold=True,
           color=SHY_WHITE, align=PP_ALIGN.CENTER)
        tb(s, label, x+0.1, 1.54, 1.9, 0.4, sz=9,
           color=RGBColor(0xCC,0xEE,0xEE), align=PP_ALIGN.CENTER, wrap=True)
        x += 2.15
    tb(s,
       "Ce que vous dites à chaque donateur :",
       0.3, 2.1, 12.73, 0.35, sz=12, bold=True, color=SHY_TEAL_D)
    highlight_box(s,
        "« Un don de 10€/mois, c'est 30 centimes par jour — moins qu'un café.\n"
        "En 1 an, vous offrez à 6 enfants l'accès à l'eau potable, à la nutrition et aux soins. »",
        0.3, 2.5, 12.73, 0.82, fill=SHY_LIGHT,
        text_color=SHY_DARK)
    rect(s, 0.3, 2.5, 0.12, 0.82, fill=SHY_RED)
    tb(s, "Votre activité : L'ascenseur de mission",
       0.3, 3.52, 12.73, 0.32, sz=12, bold=True, color=SHY_TEAL_D)
    steps5 = [
        "① Lisez seul\nle brief mission\n(3 min)",
        "② Formulez\nvotre pitch 60s\n(5 min)",
        "③ Présentez à\nvotre binôme\n(60 sec)",
        "④ Feedback pair :\n1 force + 1 conseil\n(3 min)",
        "⑤ Améliorez et\nreprésentez\n(5 min)",
    ]
    x = 0.3
    for step in steps5:
        rect(s, x, 3.92, 2.46, 1.05,
             fill=SHY_WHITE, line=SHY_TEAL, lw=Pt(1))
        tb(s, step, x+0.12, 3.98, 2.22, 0.92, sz=9, color=SHY_DARK, wrap=True)
        x += 2.5
    highlight_box(s,
        "Éthique fundraiser : Vous représentez l'UNICEF. Badge visible en permanence. "
        "Aucune pression sur le donateur. Respect total du refus.",
        0.3, 5.15, 12.73, 0.42, fill=SHY_RED, text_color=SHY_WHITE)

    # ── Slide 4 : Votre posture ───────────────────────────────────────────
    s = prs.slides.add_slide(SL(prs))
    shy_slide_header(s, "Votre Posture de Fundraiser — Ce que le donateur perçoit en 7 secondes")
    two_col(s,
        "Ce qui fait une accroche réussie",
        ["Contact visuel direct et souriant — avant de parler",
         "Corps ouvert, pieds parallèles, bras détendus",
         "Ton de voix chaleureux et assuré — ni trop fort, ni trop doux",
         "Question ouverte qui invite à s'arrêter",
         "Énergie positive constante — même après le 10ème refus"],
        "Ce qui fait fuir le donateur",
        ["Regard fuyant ou insistant",
         "Corps en tension, bras croisés ou agités",
         "Formule d'excuse : 'Excusez-moi de vous déranger…'",
         "Question fermée : 'Vous avez 2 minutes ?' (réponse attendue : Non)",
         "Débuter par 'Je vais vous parler d'argent…'"])
    highlight_box(s,
        "Exercice miroir : Entraînez-vous 2 minutes devant votre téléphone en mode selfie. "
        "Regardez votre propre accroche. Qu'est-ce que vous changeriez ?",
        0.3, 6.0, 12.73, 0.44, fill=SHY_TEAL, text_color=SHY_WHITE)

    return save(prs, f"FIM_BookJ1_Fondations_Mission_{VERSION}.pptx")


# ══════════════════════════════════════════════════════════════════════════
#  BOOK J2 — MONDE ASSOCIATIF
# ══════════════════════════════════════════════════════════════════════════

def build_bookJ2():
    prs = new_prs()

    # ── Slide 1 : Titre ──────────────────────────────────────────────────
    s = prs.slides.add_slide(SL(prs))
    shy_header(s,
        "Le Monde du Don en France",
        "Comprendre qui sont vos donateurs et pourquoi ils donnent",
        tag="BOOK J2  ·  Journée 2  ·  Donateurs français & Don régulier")
    tb(s,
        "Aujourd'hui, vous allez comprendre le comportement des donateurs français,\n"
        "le modèle économique du don régulier et pourquoi il change tout\n"
        "pour les programmes humanitaires de l'UNICEF.",
        0.5, 2.75, 12.3, 1.15, sz=14, color=SHY_DARK)

    # ── Slide 2 : Séquence J2 ─────────────────────────────────────────────
    s = prs.slides.add_slide(SL(prs))
    shy_slide_header(s, "Journée 2 — Ce que vous allez vivre aujourd'hui")
    seq_box(s,
        ["Vous analysez 3 campagnes\nde collecte en France :\ndeux succès, un échec.",
         "'Qu'est-ce qui a fait\nla différence ?\nPourquoi ça a marché ?'",
         "Portrait du donateur\nfrançais. Économie\ndu don régulier.",
         "Simulation : 'Si 50\nfrançais donnent 10€/mois,\nquel impact en 1 an ?'"],
        ["DÉCOUVERTE", "ANALYSE", "APPORT", "MISE EN PRATIQUE"],
        [SHY_TEAL_D, SHY_TEAL, RGBColor(0x2A,0x7A,0x50), SHY_RED],
        y=0.88, h=4.5)
    checklist_bar(s,
        ["Restitution orale de l'étude de cas (5 min)",
         "Fiche 'Portrait donateur' complétée",
         "1 réflexion dans votre livret personnel"],
        y=5.55, label="Vos 3 réussites à valider aujourd'hui")

    # ── Slide 3 : Le donateur français ────────────────────────────────────
    s = prs.slides.add_slide(SL(prs))
    shy_slide_header(s, "Le Donateur Français — Qui est-il vraiment ?")
    profiles = [
        ("Profil A\nLe Régulier engagé",
         SHY_TEAL_D,
         "35–55 ans · CSP+ · Sensible aux causes humanitaires\n"
         "Donne déjà à 1 ou 2 associations · Cherche un impact mesurable\n"
         "▶ Votre message : transparence des fonds + chiffres concrets"),
        ("Profil B\nLe Curieux bienveillant",
         RGBColor(0x2A,0x7A,0x50),
         "25–40 ans · Actif · Empathique mais pressé\n"
         "N'a jamais fait de don régulier · Sensible aux histoires vraies\n"
         "▶ Votre message : Impact Story + facilité du prélèvement"),
        ("Profil C\nLe Méfiant converti",
         SHY_RED,
         "40–65 ans · Expérience passée négative avec les ONG\n"
         "Demande des preuves · Teste votre connaissance du sujet\n"
         "▶ Votre message : preuves d'audit, rapport annuel public, traçabilité"),
    ]
    x = 0.3
    for title, color, desc in profiles:
        rect(s, x, 0.88, 4.15, 0.5, fill=color)
        tb(s, title, x+0.12, 0.92, 3.9, 0.42, sz=12, bold=True, color=SHY_WHITE)
        rect(s, x, 1.38, 4.15, 3.2,
             fill=SHY_WHITE, line=color, lw=Pt(1.5))
        tb(s, desc, x+0.15, 1.46, 3.85, 3.0, sz=10, color=SHY_DARK, wrap=True)
        x += 4.27
    tb(s,
       "Données clés France 2024 :",
       0.3, 4.76, 12.73, 0.3, sz=12, bold=True, color=SHY_TEAL_D)
    data = [
        ("5,2M", "donateurs\nréguliers en France"),
        ("24€/mois", "montant moyen\ndon régulier"),
        ("8×", "LTV donateur régulier\nvs ponctuel"),
        ("63%", "des Français font\nun don chaque année"),
    ]
    x = 0.3
    for val, label in data:
        rect(s, x, 5.1, 3.06, 0.92, fill=SHY_TEAL_D)
        tb(s, val, x+0.1, 5.12, 2.86, 0.5, sz=22, bold=True,
           color=SHY_WHITE, align=PP_ALIGN.CENTER)
        tb(s, label, x+0.1, 5.62, 2.86, 0.36, sz=9,
           color=RGBColor(0xCC,0xEE,0xEE), align=PP_ALIGN.CENTER, wrap=True)
        x += 3.18

    # ── Slide 4 : L'impact de votre travail ───────────────────────────────
    s = prs.slides.add_slide(SL(prs))
    shy_slide_header(s, "L'Impact de Votre Travail — Chaque Signature Compte")
    tb(s, "Simulation collective — Ce que représente 1 journée de travail :",
       0.3, 0.88, 12.73, 0.36, sz=13, bold=True, color=SHY_TEAL_D)
    calc = [
        ("Vous signez", "1 don régulier\npar heure", SHY_TEAL_D),
        ("En 1 journée\n(6h actives)", "6 nouveaux\ndonateurs", SHY_TEAL),
        ("À 15€/mois\nmoyenne", "90€ collectés\npar jour", RGBColor(0x2A,0x7A,0x50)),
        ("Sur 1 an\nde fidélité", "1 080€\npar donateur", SHY_RED),
        ("Impact UNICEF\nconcret", "6 enfants vaccinés\nchaque jour", RGBColor(0xAA,0x33,0x00)),
    ]
    x = 0.3
    for label, val, color in calc:
        rect(s, x, 1.35, 2.45, 0.5, fill=color)
        tb(s, label, x+0.1, 1.38, 2.25, 0.42, sz=9, bold=True,
           color=SHY_WHITE, align=PP_ALIGN.CENTER)
        rect(s, x, 1.85, 2.45, 1.2, fill=SHY_WHITE, line=color, lw=Pt(1.5))
        tb(s, val, x+0.1, 1.95, 2.25, 1.0, sz=13, bold=True,
           color=color, align=PP_ALIGN.CENTER)
        if x < 10.1:
            tb(s, "=", x+2.48, 2.25, 0.28, 0.6, sz=20, bold=True, color=SHY_TEAL)
        x += 2.58
    highlight_box(s,
        "Message à intégrer : Vous n'êtes pas en train de 'vendre'. "
        "Vous offrez à des personnes bienveillantes un moyen concret, simple et traçable\n"
        "de transformer leur générosité en impact réel pour des enfants dans le monde entier.",
        0.3, 3.28, 12.73, 1.0, fill=SHY_LIGHT, text_color=SHY_DARK)
    rect(s, 0.3, 3.28, 0.14, 1.0, fill=SHY_RED)
    two_col(s,
        "Don ponctuel (1 fois)",
        ["Impact immédiat mais unique",
         "Valeur moyenne : 30€",
         "Relation = 1 transaction",
         "Coût de collecte élevé"],
        "Don régulier mensuel (votre cible)",
        ["Impact continu et planifiable",
         "Valeur sur 3 ans : 540–900€",
         "Relation = engagement durable",
         "Rentabilité programme × 8"],
        y=4.45)
    highlight_box(s,
        "Votre mission : le don régulier. Toujours. Le don ponctuel est un point de départ, jamais un objectif final.",
        0.3, 6.62, 12.73, 0.38, fill=SHY_RED, text_color=SHY_WHITE)

    return save(prs, f"FIM_BookJ2_MondeAssociatif_{VERSION}.pptx")


# ══════════════════════════════════════════════════════════════════════════
#  BOOK J3 — LECTURE & SCRIPT
# ══════════════════════════════════════════════════════════════════════════

def build_bookJ3():
    prs = new_prs()

    # ── Slide 1 : Titre ──────────────────────────────────────────────────
    s = prs.slides.add_slide(SL(prs))
    shy_header(s,
        "Votre Script de Présentation",
        "Les 5 étapes pour convaincre un donateur français avec authenticité",
        tag="BOOK J3  ·  Journée 3  ·  Discours & Storytelling")
    tb(s,
        "Aujourd'hui vous allez construire votre propre version du script FIM.\n"
        "Vous ne le lirez pas — vous allez le vivre, le ressentir et le personnaliser\n"
        "pour qu'il sonne vrai à chaque rencontre.",
        0.5, 2.75, 12.3, 1.15, sz=14, color=SHY_DARK)

    # ── Slide 2 : Séquence J3 ─────────────────────────────────────────────
    s = prs.slides.add_slide(SL(prs))
    shy_slide_header(s, "Journée 3 — Ce que vous allez vivre aujourd'hui")
    seq_box(s,
        ["Votre formateur fait une\ndémonstration complète\ndevant un passant simulé.",
         "Vous observez et notez :\nQu'est-ce qui a marché ?\nQu'auriez-vous fait autrement ?",
         "Déconstruction des\n5 étapes. Ancrages.\nLanguage corporel.",
         "Trinômes 90 min :\nfundraiser / passant /\nobservateur. Rotation 15 min."],
        ["OBSERVATION", "ANALYSE", "APPORT", "ENTRAÎNEMENT"],
        [SHY_TEAL_D, SHY_TEAL, RGBColor(0x2A,0x7A,0x50), SHY_RED],
        y=0.88, h=4.5)
    checklist_bar(s,
        ["Grille C1+C2+C5 validée par votre formateur",
         "Auto-évaluation après votre 3ème round",
         "Score ≥ 12/20 pour passer à J4"],
        y=5.55, label="Vos 3 réussites à valider aujourd'hui")

    # ── Slide 3 : Les 5 étapes ────────────────────────────────────────────
    s = prs.slides.add_slide(SL(prs))
    shy_slide_header(s, "Les 5 Étapes de Votre Script — Apprenez à les Incarner")
    steps = [
        ("① ACCROCHE", "10–15 sec", SHY_TEAL_D,
         "Contact visuel · Sourire · Question ouverte\n"
         "'Bonjour ! Vous avez 2 petites minutes\npour les enfants ?'"),
        ("② CONTEXTE", "20 sec", SHY_TEAL,
         "Identification claire\n'Je travaille avec l'UNICEF — nous\ncollectons des fonds\npour les enfants en danger.'"),
        ("③ IMPACT STORY", "45–60 sec", RGBColor(0x2A,0x7A,0x50),
         "Histoire vraie d'1 enfant\nPrénom + pays + situation\nFait chiffré + résolution\npar le don"),
        ("④ PROPOSITION", "30 sec", RGBColor(0xC0,0x80,0x10),
         "'10€/mois c'est 30 centimes\npar jour — moins qu'un café.\n"
         "En 1 an vous offrez l'accès\naux soins à 6 enfants.'"),
        ("⑤ ENGAGEMENT", "20 sec", SHY_RED,
         "Formulaire simple · Accompagner\nla signature · Remerciement sincère\n"
         "'Merci ! Vous recevrez votre\nconfirmation par SMS sous 24h.'"),
    ]
    x = 0.3
    for title, dur, color, body in steps:
        rect(s, x, 0.88, 2.5, 0.38, fill=color)
        tb(s, f"{title}  [{dur}]", x+0.08, 0.91, 2.34, 0.3,
           sz=9, bold=True, color=SHY_WHITE)
        rect(s, x, 1.26, 2.5, 4.6, fill=SHY_WHITE, line=color, lw=Pt(1.5))
        tb(s, body, x+0.12, 1.34, 2.26, 4.3, sz=10, color=SHY_DARK, wrap=True)
        x += 2.56
    highlight_box(s,
        "Durée totale : 2 min 15 sec  ·  Version express (passant pressé) : 55 sec  "
        "·  Le silence après la proposition est votre allié — ne le brisez pas.",
        0.3, 6.0, 12.73, 0.44, fill=SHY_TEAL_D, text_color=SHY_WHITE)

    # ── Slide 4 : Impact Story ────────────────────────────────────────────
    s = prs.slides.add_slide(SL(prs))
    shy_slide_header(s, "L'Impact Story — Comment Raconter une Histoire qui Touche")
    tb(s,
       "L'Impact Story est le cœur de votre script. "
       "C'est ce moment précis où le donateur passe\n"
       "de 'intéressé' à 'convaincu'. Voici comment la construire :",
       0.35, 0.88, 12.6, 0.7, sz=12, color=SHY_DARK)
    story_parts = [
        ("PERSONNAGE", SHY_TEAL_D,
         "Donnez un prénom, un âge, un pays.\n"
         "'Amara, 3 ans, Sahel.'\n"
         "Le donateur doit voir un visage, pas une statistique."),
        ("SITUATION", SHY_RED,
         "Décrivez le danger concret avec 1 seul chiffre.\n"
         "'Sans accès à l'eau potable, 1 enfant sur 5\nne survivra pas à sa première année.'"),
        ("BASCULE", RGBColor(0xC0,0x80,0x10),
         "C'est ici que votre don entre en scène.\n"
         "'Grâce aux donateurs comme vous,\nl'UNICEF a livré 50 000 kits eau potable.'"),
        ("RÉSOLUTION", RGBColor(0x2A,0x7A,0x50),
         "La fin heureuse — rendue possible\npar le don.\n'Amara a aujourd'hui 4 ans et elle est\nscolarisée.'"),
    ]
    x = 0.3
    for title, color, body in story_parts:
        rect(s, x, 1.7, 3.08, 0.36, fill=color)
        tb(s, title, x+0.1, 1.73, 2.88, 0.28, sz=10, bold=True, color=SHY_WHITE)
        rect(s, x, 2.06, 3.08, 2.9, fill=SHY_WHITE, line=color, lw=Pt(1.5))
        tb(s, body, x+0.14, 2.14, 2.8, 2.72, sz=10, color=SHY_DARK, wrap=True)
        x += 3.2
    highlight_box(s,
        "Règle d'or : 1 seul enfant · 1 seul problème · 1 seul chiffre · 1 seule solution.\n"
        "La simplicité crée l'émotion. La liste de statistiques crée la distance.",
        0.3, 5.14, 12.73, 0.72, fill=SHY_LIGHT, text_color=SHY_DARK)
    rect(s, 0.3, 5.14, 0.14, 0.72, fill=SHY_RED)
    highlight_box(s,
        "Exercice : Écrivez votre propre Impact Story en 5 lignes — un enfant, une situation, un impact.",
        0.3, 6.0, 12.73, 0.44, fill=SHY_RED, text_color=SHY_WHITE)

    return save(prs, f"FIM_BookJ3_Lecture_Script_{VERSION}.pptx")


# ══════════════════════════════════════════════════════════════════════════
#  BOOK J4 — OBJECTIONS
# ══════════════════════════════════════════════════════════════════════════

def build_bookJ4():
    prs = new_prs()

    # ── Slide 1 : Titre ──────────────────────────────────────────────────
    s = prs.slides.add_slide(SL(prs))
    shy_header(s,
        "Gérer les Objections",
        "Transformer chaque refus en opportunité d'engagement",
        tag="BOOK J4  ·  Journée 4  ·  Écoute active & Réponses")
    tb(s,
        "Aujourd'hui vous allez apprendre à aimer les objections.\n"
        "Un donateur qui objecte est un donateur qui dialogue.\n"
        "Chaque 'non' cache souvent un 'pas encore' — et c'est là que vous entrez en jeu.",
        0.5, 2.75, 12.3, 1.15, sz=14, color=SHY_DARK)

    # ── Slide 2 : Séquence J4 ─────────────────────────────────────────────
    s = prs.slides.add_slide(SL(prs))
    shy_slide_header(s, "Journée 4 — Ce que vous allez vivre aujourd'hui")
    seq_box(s,
        ["Vous analysez 5 extraits\nvidéo : bonne et mauvaise\ngestion d'objection.",
         "'Qu'avez-vous remarqué ?\nQu'est-ce qui a tout changé\ndans la réponse ?'",
         "Les 12 objections types.\nLa réponse en 4 temps.\nLes mots qui ouvrent.",
         "Battle : tournoi d'objections\npar équipes. Tirage au sort.\nLe groupe vote la meilleure réponse."],
        ["OBSERVATION", "ANALYSE", "APPORT", "ENTRAÎNEMENT"],
        [SHY_TEAL_D, SHY_TEAL, RGBColor(0x2A,0x7A,0x50), SHY_RED],
        y=0.88, h=4.5)
    checklist_bar(s,
        ["Grille écoute active + objections validée",
         "Réflexion dans votre livret : 'Quelle objection me challenge le plus ?'",
         "Score ≥ 14/20 pour validation J4"],
        y=5.55, label="Vos 3 réussites à valider aujourd'hui")

    # ── Slide 3 : Répondre en 4 temps ─────────────────────────────────────
    s = prs.slides.add_slide(SL(prs))
    shy_slide_header(s, "Répondre à une Objection — Votre Méthode en 4 Temps")
    steps4 = [
        ("1\nACCUSEZ RÉCEPTION",
         SHY_TEAL_D,
         "Ne justifiez pas immédiatement.\n"
         "Montrez que vous avez entendu.\n\n"
         "'Je comprends tout à fait…'\n'C'est une question que beaucoup\nde personnes se posent…'"),
        ("2\nREFORMULEZ",
         SHY_TEAL,
         "Assurez-vous d'avoir compris\nle vrai frein, pas la surface.\n\n"
         "'Si je comprends bien, vous voulez\ndire que… c'est bien ça ?'\n'Ce qui vous préoccupe c'est…'"),
        ("3\nRÉPONDEZ AVEC PREUVE",
         RGBColor(0x2A,0x7A,0x50),
         "1 preuve concrète +\n1 chiffre + 1 exemple réel.\n\n"
         "Rapport annuel UNICEF public.\nAudit Comité de la Charte.\nTémoignage donateur."),
        ("4\nRELANCEZ",
         SHY_RED,
         "Revenez à la proposition\navec naturel. Pas de pression.\n\n"
         "'Est-ce que cela répond à\nvotre question ?\nOn peut aller plus loin ?'"),
    ]
    x = 0.3
    for title, color, body in steps4:
        rect(s, x, 0.88, 3.08, 0.7, fill=color)
        tb(s, title, x+0.12, 0.92, 2.84, 0.62, sz=12, bold=True,
           color=SHY_WHITE, align=PP_ALIGN.CENTER)
        rect(s, x, 1.58, 3.08, 3.7, fill=SHY_WHITE, line=color, lw=Pt(1.5))
        tb(s, body, x+0.14, 1.66, 2.8, 3.5, sz=10, color=SHY_DARK, wrap=True)
        x += 3.2
    highlight_box(s,
        "Le silence après votre réponse est normal. Laissez le donateur traiter l'information. "
        "Comptez 3 secondes avant de relancer.",
        0.3, 5.45, 12.73, 0.44, fill=SHY_TEAL_D, text_color=SHY_WHITE)

    # ── Slide 4 : Les 12 objections ───────────────────────────────────────
    s = prs.slides.add_slide(SL(prs))
    shy_slide_header(s, "Les 12 Objections Types — Vos Réponses Prêtes à l'Emploi")
    objections = [
        ("O1",  "Je n'ai pas les moyens",           "Financière",  "30 centimes par jour — moins qu'un café. Impact : 6 enfants par an."),
        ("O2",  "Je ne fais pas confiance aux ONG", "Méfiance",    "Rapport public, audit Comité de la Charte, 82% fonds aux programmes."),
        ("O3",  "J'ai déjà fait un don",            "Déjà engagé", "Super — et si on transformait ce geste en impact mensuel continu ?"),
        ("O4",  "Je dois réfléchir",                "Report",      "Qu'est-ce qui vous retient ? (identifier le vrai frein, puis ancrer l'urgence)"),
        ("O5",  "Combien vous gagnez ?",             "Méfiance",    "Transparence totale : emploi déclaré, encadrement légal, mission noble."),
        ("O6",  "Je donne à d'autres causes",        "Concurrent",  "Chaque cause est unique — les droits des enfants ne concurrencent rien."),
        ("O7",  "Ce n'est pas le bon moment",        "Temporel",    "Le premier prélèvement est dans 30 jours — vous avez le temps de confirmer."),
        ("O8",  "L'argent ne parvient pas",          "Méfiance",    "82% des dons aux programmes — rapport annuel en ligne, auditeurs indépendants."),
        ("O9",  "Je suis pressé",                    "Temporel",    "45 secondes suffit : 10€/mois = 6 enfants aidés par an. C'est tout."),
        ("O10", "Mon conjoint(e) décide",            "Autorité",    "Je vous envoie la brochure par SMS. Vous en parlez ce soir et je peux vous rappeler ?"),
        ("O11", "Je n'aime pas les prélèvements",    "Processus",   "Arrêt en 1 clic, sans engagement minimum, modification à tout moment."),
        ("O12", "Politique / religion",              "Valeurs",     "UNICEF est strictement apolitique et aconfessionnel — seuls les droits des enfants comptent."),
    ]
    col_x = [0.3, 0.75, 5.52, 7.32]
    col_w = [0.42, 4.74, 1.77, 5.84]
    headers = ["#", "L'objection que vous entendez", "Type", "Ce que vous répondez"]
    y = 0.88
    for h_text, x, w in zip(headers, col_x, col_w):
        rect(s, x, y, w, 0.3, fill=SHY_TEAL_D)
        tb(s, h_text, x+0.06, y+0.04, w-0.1, 0.24,
           sz=9, bold=True, color=SHY_WHITE)
    y += 0.32
    type_colors = {
        "Financière": RGBColor(0x2A,0x7A,0x50),
        "Méfiance":   SHY_RED,
        "Déjà engagé":SHY_TEAL,
        "Report":     RGBColor(0xC0,0x80,0x10),
        "Concurrent": RGBColor(0x66,0x44,0xAA),
        "Temporel":   SHY_TEAL_D,
        "Autorité":   RGBColor(0x88,0x44,0x44),
        "Processus":  RGBColor(0x44,0x77,0xAA),
        "Valeurs":    RGBColor(0x88,0x44,0x88),
    }
    rh = 0.465
    for i, (num, obj, typ, rep) in enumerate(objections):
        bg = SHY_WHITE if i % 2 == 0 else SHY_LIGHT
        rect(s, col_x[0], y, col_w[0], rh, fill=SHY_TEAL_D)
        tb(s, num, col_x[0]+0.04, y+0.1, col_w[0]-0.06, rh-0.15,
           sz=8, bold=True, color=SHY_WHITE, align=PP_ALIGN.CENTER)
        rect(s, col_x[1], y, col_w[1], rh, fill=bg)
        tb(s, obj, col_x[1]+0.1, y+0.08, col_w[1]-0.15, rh-0.1, sz=9, color=SHY_DARK)
        tc = type_colors.get(typ, SHY_TEAL)
        rect(s, col_x[2], y, col_w[2], rh, fill=tc)
        tb(s, typ, col_x[2]+0.06, y+0.08, col_w[2]-0.1, rh-0.1,
           sz=8, bold=True, color=SHY_WHITE, align=PP_ALIGN.CENTER)
        rect(s, col_x[3], y, col_w[3], rh, fill=bg)
        tb(s, rep, col_x[3]+0.1, y+0.06, col_w[3]-0.15, rh-0.08,
           sz=9, color=SHY_DARK, italic=True)
        y += rh + 0.005

    return save(prs, f"FIM_BookJ4_Objections_{VERSION}.pptx")


# ══════════════════════════════════════════════════════════════════════════
#  BOOK J5 — SYNTHÈSE & SIMULATIONS
# ══════════════════════════════════════════════════════════════════════════

def build_bookJ5():
    prs = new_prs()

    # ── Slide 1 : Titre ──────────────────────────────────────────────────
    s = prs.slides.add_slide(SL(prs))
    shy_header(s,
        "Votre Certification FIM",
        "Synthèse · Simulations · Plan d'action · Vous êtes prêt(e) pour le terrain",
        tag="BOOK J5  ·  Journée 5  ·  Certification FIM")
    tb(s,
        "Aujourd'hui c'est votre jour.\n"
        "Vous allez montrer ce que vous savez faire — pas dans un examen théorique,\n"
        "mais en situation réelle simulée, face à un passant avec toute sa complexité.",
        0.5, 2.75, 12.3, 1.15, sz=14, color=SHY_DARK)

    # ── Slide 2 : Programme J5 ────────────────────────────────────────────
    s = prs.slides.add_slide(SL(prs))
    shy_slide_header(s, "Votre Journée 5 — Étape par Étape")
    programme = [
        ("08h30", "09h00", "Révision flash",
         "Tour de table rapide — Comment vous sentez-vous ce matin ? Rappel des 7 compétences."),
        ("09h00", "11h00", "Simulations × 2 rounds",
         "Round 1 : donateur 'neutre'  ·  Round 2 : donateur 'difficile avec 3 objections'\n"
         "Vous observez vos camarades avec la grille — et ils font de même pour vous."),
        ("11h00", "11h30", "Débriefing collectif",
         "Célébration des progrès de chacun. Les 3 apprentissages collectifs de la semaine."),
        ("11h30", "12h30", "Évaluation individuelle",
         "Votre formateur vous évalue sur la grille 28 points. Entretien individuel 5 min."),
        ("13h30", "14h30", "Votre plan d'action",
         "Vous construisez votre plan personnel : 4 objectifs sur 4 semaines terrain."),
        ("14h30", "15h00", "Remise des certifications",
         "Remise officielle · Engagement public 1 min · Votre premier jour terrain = demain."),
    ]
    y = 0.88
    for start, end, title, detail in programme:
        rect(s, 0.3, y, 1.45, 0.72, fill=SHY_RED)
        tb(s, f"{start}\n{end}", 0.32, y+0.08, 1.41, 0.58,
           sz=11, bold=True, color=SHY_WHITE, align=PP_ALIGN.CENTER)
        rect(s, 1.77, y, 3.0, 0.72, fill=SHY_TEAL_D)
        tb(s, title, 1.88, y+0.16, 2.82, 0.42, sz=12, bold=True, color=SHY_WHITE)
        rect(s, 4.79, y, 8.24, 0.72,
             fill=SHY_WHITE, line=SHY_TEAL, lw=Pt(0.8))
        tb(s, detail, 4.92, y+0.06, 8.02, 0.62, sz=9, color=SHY_DARK, wrap=True)
        y += 0.75

    # ── Slide 3 : Grille de certification ────────────────────────────────
    s = prs.slides.add_slide(SL(prs))
    shy_slide_header(s, "Votre Grille de Certification — 7 Critères / 28 Points")
    tb(s, "Voici exactement ce sur quoi vous serez évalué(e) — aucune surprise :",
       0.35, 0.85, 12.6, 0.32, sz=12, bold=True, color=SHY_TEAL_D)
    criteria = [
        ("C1", "Votre accroche est naturelle en < 15 sec",
         "Contact visuel, sourire, question ouverte, aisance corporelle"),
        ("C2", "Votre script est fluide et personnalisé",
         "Les 5 étapes dans l'ordre, langage adapté au profil du donateur"),
        ("C3", "Votre Impact Story touche et convainc",
         "Récit bénéficiaire spécifique, chiffre concret, résolution par le don"),
        ("C4", "Vous traitez ≥ 2 objections avec la méthode en 4 temps",
         "Accusé réception, reformulation, preuve, relance"),
        ("C5", "Vous obtenez l'engagement et concluez",
         "Proposition montant adapté, formulaire signé, remerciement"),
        ("C6", "Vous respectez les règles éthiques en toute situation",
         "Badge visible, aucune pression, déontologie fundraiser"),
        ("C7", "Vous gérez votre état émotionnel",
         "Posture positive maintenue après refus, zéro découragement visible"),
    ]
    col_x = [0.3, 0.75, 7.6, 10.35, 11.15, 11.95, 12.75]
    col_w = [0.42, 6.82, 2.72, 0.78, 0.78, 0.78, 0.5]
    headers = ["#", "Ce que vous devez démontrer", "Indicateurs observés", "1", "2", "3", "4"]
    y = 1.26
    for h_txt, x, w in zip(headers, col_x, col_w):
        rect(s, x, y, w, 0.3, fill=SHY_TEAL_D)
        tb(s, h_txt, x+0.06, y+0.04, w-0.1, 0.24,
           sz=9, bold=True, color=SHY_WHITE,
           align=PP_ALIGN.CENTER if len(h_txt) <= 2 else PP_ALIGN.LEFT)
    y += 0.32
    rh = 0.52
    for i, (code, criterion, indic) in enumerate(criteria):
        bg = SHY_WHITE if i % 2 == 0 else SHY_LIGHT
        rect(s, col_x[0], y, col_w[0], rh, fill=SHY_RED)
        tb(s, code, col_x[0]+0.04, y+0.12, col_w[0]-0.06, rh-0.18,
           sz=10, bold=True, color=SHY_WHITE, align=PP_ALIGN.CENTER)
        rect(s, col_x[1], y, col_w[1], rh, fill=bg)
        tb(s, criterion, col_x[1]+0.1, y+0.08, col_w[1]-0.15, rh-0.12,
           sz=10, bold=True, color=SHY_TEAL_D)
        rect(s, col_x[2], y, col_w[2], rh, fill=bg)
        tb(s, indic, col_x[2]+0.1, y+0.06, col_w[2]-0.15, rh-0.1,
           sz=9, color=SHY_GRAY, italic=True)
        for k, (x, w) in enumerate(zip(col_x[3:], col_w[3:])):
            rect(s, x, y, w, rh, fill=SHY_WHITE, line=SHY_TEAL, lw=Pt(0.5))
            tb(s, str(k+1), x+0.1, y+0.12, w-0.2, rh-0.2,
               sz=9, color=SHY_LGRAY, align=PP_ALIGN.CENTER)
        y += rh + 0.02
    rect(s, col_x[0], y, sum(col_w[:3])+0.02, 0.38, fill=SHY_TEAL_D)
    tb(s, "TOTAL  /28", col_x[0]+0.12, y+0.06, 9.0, 0.28,
       sz=11, bold=True, color=SHY_WHITE)
    rect(s, col_x[3], y, sum(col_w[3:])+0.04, 0.38,
         fill=SHY_WHITE, line=SHY_RED, lw=Pt(2))
    tb(s, "__ / 28", col_x[3]+0.12, y+0.06, 2.4, 0.28,
       sz=12, bold=True, color=SHY_RED, align=PP_ALIGN.CENTER)
    y += 0.42
    highlight_box(s,
        "✅ Certifié(e) FIM : ≥ 20/28   "
        "⚠ Formation complémentaire ciblée : 14–19/28   "
        "🔄 Évaluation de rattrapage à J+15 : < 14/28",
        0.3, y+0.04, 12.73, 0.38, fill=SHY_RED, text_color=SHY_WHITE)

    # ── Slide 4 : Votre plan d'action ─────────────────────────────────────
    s = prs.slides.add_slide(SL(prs))
    shy_slide_header(s, "Votre Plan d'Action — 4 Semaines pour Devenir Expert")
    tb(s, "Vous remplissez ce plan aujourd'hui en J5. "
       "Il vous appartient — votre formateur vous rappelle à J+15 et J+30.",
       0.35, 0.88, 12.6, 0.55, sz=12, color=SHY_DARK)
    weeks = [
        ("SEMAINE 1\nJ+7",    SHY_TEAL_D,
         "Objectif :\nMaîtriser votre accroche en < 12 secondes",
         "Comment mesurer :\nChronomètre + feedback J+7"),
        ("SEMAINE 2\nJ+14",   SHY_TEAL,
         "Objectif :\nTraiter O1 et O2 sans hésitation",
         "Comment mesurer :\n0 rupture de discours en simulation"),
        ("SEMAINE 3\nJ+21",   RGBColor(0x2A,0x7A,0x50),
         "Objectif :\nAtteindre ≥ 1 don régulier par heure",
         "Comment mesurer :\nFeuille suivi terrain quotidienne"),
        ("SEMAINE 4\nJ+30",   SHY_RED,
         "Objectif :\nRévision PAP + objectifs M2",
         "Comment mesurer :\nEntretien téléphonique formateur"),
    ]
    x = 0.3
    for tag, color, obj, mesure in weeks:
        rect(s, x, 1.55, 3.08, 0.62, fill=color)
        tb(s, tag, x+0.12, 1.58, 2.84, 0.56, sz=11, bold=True,
           color=SHY_WHITE, align=PP_ALIGN.CENTER)
        rect(s, x, 2.17, 3.08, 1.7, fill=SHY_WHITE, line=color, lw=Pt(1.5))
        tb(s, obj, x+0.14, 2.23, 2.8, 1.55, sz=10, color=SHY_DARK, wrap=True)
        rect(s, x, 3.87, 3.08, 1.4, fill=SHY_LIGHT, line=color, lw=Pt(1))
        tb(s, mesure, x+0.14, 3.93, 2.8, 1.28, sz=9, color=SHY_GRAY,
           italic=True, wrap=True)
        x += 3.2
    highlight_box(s,
        "Votre formateur SHY-Performance vous contacte à J+15 et J+30 pour un bilan de progression.\n"
        "À J+90 : analyse de vos KPI terrain — taux de conversion, montant moyen, rétention donateur.",
        0.3, 5.48, 12.73, 0.72, fill=SHY_TEAL, text_color=SHY_WHITE)
    highlight_box(s,
        "Bienvenue dans la famille SHY-Performance. Vous êtes maintenant Fundraiser Impact Mission certifié(e).",
        0.3, 6.32, 12.73, 0.42, fill=SHY_RED, text_color=SHY_WHITE)

    return save(prs, f"FIM_BookJ5_Synthese_Simulations_{VERSION}.pptx")


# ══════════════════════════════════════════════════════════════════════════
#  MAIN
# ══════════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    print(f"\n🎨 Génération FIM {VERSION} — Charte SHY-Performance · France · UNICEF\n")
    files = []
    files.append(build_book0())
    files.append(build_book1())
    files.append(build_bookJ1())
    files.append(build_bookJ2())
    files.append(build_bookJ3())
    files.append(build_bookJ4())
    files.append(build_bookJ5())
    print(f"\n✅ {len(files)} modules générés dans {OUT_DIR}/")

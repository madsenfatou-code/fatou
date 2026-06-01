"""
FIM V6 — Générateur PPTX — 7 modules
Version : V6-01062026SHY-TE
Référentiels : UM6P / APC / Bloom / Kolb / Kirkpatrick / CMC 2025
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt
import copy, os

# ─── PALETTE ───────────────────────────────────────────────────────────────
TEAL      = RGBColor(0x00, 0x7A, 0x7A)
TEAL_DARK = RGBColor(0x00, 0x50, 0x50)
GOLD      = RGBColor(0xC9, 0xA8, 0x4C)
WHITE     = RGBColor(0xFF, 0xFF, 0xFF)
DARK      = RGBColor(0x1A, 0x1A, 0x1A)
LIGHT_BG  = RGBColor(0xF0, 0xF7, 0xF7)
RED_SOFT  = RGBColor(0xC0, 0x39, 0x2B)
GREEN_OK  = RGBColor(0x1A, 0x7A, 0x1A)

OUT_DIR = "/home/user/fatou"
VERSION = "V6-01062026SHY-TE"

# ─── HELPERS ────────────────────────────────────────────────────────────────

def new_prs():
    prs = Presentation()
    prs.slide_width  = Inches(13.33)
    prs.slide_height = Inches(7.5)
    return prs

def blank_layout(prs):
    return prs.slide_layouts[6]  # completely blank

def add_rect(slide, x, y, w, h, fill_color=None, line_color=None, line_width=Pt(0)):
    from pptx.util import Pt
    shape = slide.shapes.add_shape(
        1,  # MSO_SHAPE_TYPE.RECTANGLE
        Inches(x), Inches(y), Inches(w), Inches(h)
    )
    shape.line.width = line_width
    if fill_color:
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill_color
    else:
        shape.fill.background()
    if line_color:
        shape.line.color.rgb = line_color
    else:
        shape.line.fill.background()
    return shape

def add_text_box(slide, text, x, y, w, h,
                 font_size=18, bold=False, color=DARK,
                 align=PP_ALIGN.LEFT, wrap=True, italic=False):
    txBox = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    txBox.word_wrap = wrap
    tf = txBox.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    return txBox

def add_para(tf, text, font_size=14, bold=False, color=DARK,
             align=PP_ALIGN.LEFT, bullet=False, indent=0, italic=False):
    p = tf.add_paragraph()
    p.alignment = align
    if bullet:
        p.level = indent
    run = p.add_run()
    run.text = ("• " if bullet else "") + text
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    return p

def header_slide(slide, title, subtitle="", tag="", bg_color=TEAL_DARK):
    """Full-width header band"""
    add_rect(slide, 0, 0, 13.33, 7.5, fill_color=LIGHT_BG)
    add_rect(slide, 0, 0, 13.33, 2.6, fill_color=bg_color)
    add_rect(slide, 0, 2.55, 13.33, 0.1, fill_color=GOLD)
    if tag:
        tb = add_text_box(slide, tag, 0.4, 0.2, 4, 0.4,
                          font_size=10, bold=True, color=GOLD)
    add_text_box(slide, title, 0.5, 0.55, 12.3, 1.5,
                 font_size=30, bold=True, color=WHITE, align=PP_ALIGN.LEFT)
    if subtitle:
        add_text_box(slide, subtitle, 0.5, 2.0, 12.3, 0.55,
                     font_size=13, color=RGBColor(0xDD,0xEE,0xEE), italic=True)
    add_text_box(slide, f"SHY-Performance · {VERSION} · Référentiel UM6P/CMC 2025",
                 0, 7.1, 13.33, 0.4, font_size=9,
                 color=RGBColor(0x88,0x88,0x88), align=PP_ALIGN.CENTER)

def content_slide(slide, title, body_fn, bg=True):
    if bg:
        add_rect(slide, 0, 0, 13.33, 7.5, fill_color=LIGHT_BG)
    add_rect(slide, 0, 0, 13.33, 0.75, fill_color=TEAL_DARK)
    add_rect(slide, 0, 0.73, 13.33, 0.04, fill_color=GOLD)
    add_text_box(slide, title, 0.3, 0.08, 12.7, 0.62,
                 font_size=18, bold=True, color=WHITE)
    add_text_box(slide, f"SHY-Performance · {VERSION}",
                 0, 7.1, 13.33, 0.4, font_size=8,
                 color=RGBColor(0x88,0x88,0x88), align=PP_ALIGN.CENTER)
    body_fn(slide)

def two_col(slide, left_title, left_items, right_title, right_items,
            y_start=0.95, col_w=6.0, gap=0.4):
    lx, rx = 0.3, 0.3 + col_w + gap
    # Left header
    add_rect(slide, lx, y_start, col_w, 0.38, fill_color=TEAL)
    add_text_box(slide, left_title, lx+0.1, y_start+0.04, col_w-0.2, 0.32,
                 font_size=12, bold=True, color=WHITE)
    # Right header
    add_rect(slide, rx, y_start, col_w, 0.38, fill_color=TEAL)
    add_text_box(slide, right_title, rx+0.1, y_start+0.04, col_w-0.2, 0.32,
                 font_size=12, bold=True, color=WHITE)
    # Left items
    ly = y_start + 0.45
    for item in left_items:
        add_text_box(slide, "• " + item, lx+0.1, ly, col_w-0.2, 0.38,
                     font_size=11, color=DARK)
        ly += 0.40
    # Right items
    ry = y_start + 0.45
    for item in right_items:
        add_text_box(slide, "• " + item, rx+0.1, ry, col_w-0.2, 0.38,
                     font_size=11, color=DARK)
        ry += 0.40

def bloom_bar(slide, y=5.8):
    """Bloom taxonomy visual bar"""
    levels = [
        ("1 Se souvenir", RGBColor(0x78,0xA5,0xC5)),
        ("2 Comprendre", RGBColor(0x5D,0x9E,0xA6)),
        ("3 Appliquer",  RGBColor(0x4A,0x90,0x70)),
        ("4 Analyser",   RGBColor(0xC9,0xA8,0x4C)),
        ("5 Évaluer",    RGBColor(0xD2,0x7A,0x42)),
        ("6 Créer",      RGBColor(0xC0,0x39,0x2B)),
    ]
    add_text_box(slide, "Taxonomie de Bloom :", 0.3, y-0.28, 3, 0.28,
                 font_size=9, bold=True, color=TEAL_DARK)
    x = 0.3
    w = 2.1
    for label, color in levels:
        add_rect(slide, x, y, w, 0.35, fill_color=color)
        add_text_box(slide, label, x+0.05, y+0.04, w-0.1, 0.28,
                     font_size=8, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        x += w + 0.02

def kolb_box(slide, steps, x=0.3, y=1.0, w=12.7, h=1.1):
    colors = [RGBColor(0x2E,0x86,0xAB), RGBColor(0x4C,0xAF,0x82),
              RGBColor(0xF4,0xAE,0x2D), RGBColor(0xE8,0x48,0x55)]
    labels = ["EC — Expérience Concrète", "OR — Observation Réfléchie",
              "CA — Conceptualisation", "EA — Expérimentation Active"]
    sw = (w - 0.09) / 4
    for i, (label, step) in enumerate(zip(labels, steps)):
        bx = x + i*(sw+0.03)
        add_rect(slide, bx, y, sw, 0.28, fill_color=colors[i])
        add_text_box(slide, label, bx+0.05, y+0.03, sw-0.1, 0.22,
                     font_size=8, bold=True, color=WHITE)
        add_rect(slide, bx, y+0.28, sw, h-0.28,
                 fill_color=WHITE, line_color=colors[i], line_width=Pt(1))
        add_text_box(slide, step, bx+0.1, y+0.32, sw-0.2, h-0.4,
                     font_size=9, color=DARK, wrap=True)

def eval_box(slide, items, x=0.3, y=6.0, w=12.7, label="Évaluation formative"):
    add_rect(slide, x, y, w, 0.3, fill_color=GOLD)
    add_text_box(slide, f"🎯 {label}", x+0.1, y+0.04, w-0.2, 0.24,
                 font_size=10, bold=True, color=TEAL_DARK)
    tx = x + 0.1
    tx_w = (w - 0.2) / len(items)
    for item in items:
        add_text_box(slide, "✓ " + item, tx, y+0.32, tx_w-0.1, 0.55,
                     font_size=9, color=DARK, wrap=True)
        tx += tx_w

def save(prs, name):
    path = os.path.join(OUT_DIR, name)
    prs.save(path)
    print(f"  ✓ Saved: {name}")
    return path


# ══════════════════════════════════════════════════════════════════════════
#  BOOK 0 — FIL CONDUCTEUR
# ══════════════════════════════════════════════════════════════════════════

def build_book0():
    prs = new_prs()
    SL = prs.slide_layouts[6]

    # — Slide 1 : Titre
    s = prs.slides.add_slide(SL)
    header_slide(s,
        "FIM — Fil Conducteur",
        "Document de gouvernance pédagogique · 5 jours · 35h · V6",
        tag="BOOK 0 · UM6P/CMC 2025")
    add_text_box(s, "Programme Fundraiser Impact Mission",
                 0.5, 2.8, 12, 0.5, font_size=16, bold=True, color=TEAL_DARK)
    add_text_box(s,
        "Ce document pilote l'ensemble de la formation :\n"
        "séquençage des 5 journées · contrat pédagogique · tableau de bord · règles de vie",
        0.5, 3.35, 12, 1.0, font_size=13, color=DARK)
    add_text_box(s,
        "Référentiels : UM6P CCE · CMC 2025 Niveau 3 · ADDIE · Andragogie Knowles",
        0.5, 4.4, 12, 0.4, font_size=10, italic=True,
        color=RGBColor(0x44,0x88,0x88))

    # — Slide 2 : Objectifs du Fil Conducteur
    s = prs.slides.add_slide(SL)
    def body(sl):
        two_col(sl,
            "Objectifs pédagogiques V6 (Bloom 2-3)",
            ["Expliquer la logique pédagogique des 5 jours",
             "Utiliser le tableau de bord pour piloter la progression",
             "Appliquer les règles de vie pour un cadre sécurisant",
             "Lire et compléter le Livret de Formation Individuel (LFI)"],
            "Ajouts V6 vs V5",
            ["Fiche de diagnostic d'entrée (auto-positionnement C1–C7)",
             "Contrat pédagogique signé formateur ↔ apprenant",
             "Calendrier des évaluations formatives J1→J5",
             "Protocole de suivi post-formation J+15, J+30, J+90"])
        eval_box(sl,
            ["Quiz 10 questions (diagnostic d'entrée)",
             "Auto-positionnement C1–C7 dans le LFI",
             "Signature du contrat pédagogique"],
            label="Évaluation diagnostique — J1 matin")
        bloom_bar(sl)
    content_slide(s, "Objectifs & Structure du Fil Conducteur V6", body)

    # — Slide 3 : Planning 5 jours
    s = prs.slides.add_slide(SL)
    def body(sl):
        add_rect(sl, 0, 0, 13.33, 7.5, fill_color=LIGHT_BG)
        add_rect(sl, 0, 0, 13.33, 0.75, fill_color=TEAL_DARK)
        add_rect(sl, 0, 0.73, 13.33, 0.04, fill_color=GOLD)
        add_text_box(sl, "Planning 5 jours FIM V6 — Séquençage Kolb intégré",
                     0.3, 0.08, 12.7, 0.62, font_size=18, bold=True, color=WHITE)
        days = [
            ("J1", "Fondations & Mission",
             "EC : Témoignage bénéficiaire\nOR : Debriefing émotionnel\nCA : UNICEF + cadre légal\nEA : Pitch 60s",
             "Éval. diag. + Quiz 5Q"),
            ("J2", "Monde Associatif",
             "EC : Étude de cas ONG\nOR : Analyse facteurs succès\nCA : Loi, économie du don\nEA : Simulation budgétaire",
             "Restitution orale + Fiche"),
            ("J3", "Script & Discours",
             "EC : Démo formateur live\nOR : Grille observation\nCA : 5 étapes du script\nEA : Trinômes 90 min",
             "Grille C1+C2+C5 — seuil 12/20"),
            ("J4", "Gestion Objections",
             "EC : Cas filmés analyse\nOR : Identification patterns\nCA : Méthode CROC/CAP\nEA : Battle objections",
             "Grille C3+C4 — seuil 14/20"),
            ("J5", "Synthèse & Certif.",
             "Simulation ×2 rounds\nDebriefing collectif\nÉval. sommative\nConstruction PAP",
             "Grille 28 pts — seuil 20/28"),
        ]
        x = 0.2
        for tag, title, content, eval_txt in days:
            add_rect(sl, x, 0.85, 2.5, 0.35, fill_color=TEAL)
            add_text_box(sl, f"{tag} — {title}", x+0.08, 0.88, 2.35, 0.28,
                         font_size=10, bold=True, color=WHITE)
            add_rect(sl, x, 1.2, 2.5, 4.2,
                     fill_color=WHITE, line_color=TEAL, line_width=Pt(1))
            add_text_box(sl, content, x+0.1, 1.25, 2.3, 3.8,
                         font_size=9, color=DARK, wrap=True)
            add_rect(sl, x, 5.4, 2.5, 0.65, fill_color=RGBColor(0xFD,0xF9,0xEF))
            add_text_box(sl, "🎯 " + eval_txt, x+0.08, 5.44, 2.35, 0.58,
                         font_size=8, color=TEAL_DARK, wrap=True)
            x += 2.62
        add_text_box(sl, f"SHY-Performance · {VERSION}",
                     0, 7.1, 13.33, 0.4, font_size=8,
                     color=RGBColor(0x88,0x88,0x88), align=PP_ALIGN.CENTER)
    body(s)

    # — Slide 4 : Contrat pédagogique & LFI
    s = prs.slides.add_slide(SL)
    def body(sl):
        two_col(sl,
            "Contrat pédagogique — Engagements réciproques",
            ["Le formateur s'engage à donner un feedback individuel chaque jour",
             "L'apprenant s'engage à participer activement aux simulations",
             "Le formateur respecte le rythme de chaque profil",
             "L'apprenant complète son LFI chaque soir (5 min)",
             "Confidentialité des échanges de groupe garantie"],
            "Livret de Formation Individuel (LFI) — Structure",
            ["Page d'ouverture : auto-positionnement C1–C7",
             "1 page par demi-journée : apprentissages + insights",
             "Grilles d'évaluation formative J1–J4",
             "Plan d'Action Personnel (PAP) J5",
             "Page bilan : évolution C1–C7 avant/après"])
        eval_box(sl,
            ["Contrat signé en début de J1",
             "LFI distribué J1 matin, complété chaque soir",
             "Vérification LFI par formateur à J3 et J5"],
            label="Points de contrôle Fil Conducteur")
    content_slide(s, "Contrat Pédagogique & Livret de Formation Individuel (LFI)", body)

    # — Slide 5 : Règles de vie & CMC
    s = prs.slides.add_slide(SL)
    def body(sl):
        two_col(sl,
            "Règles de vie — Cadre sécurisant",
            ["Droit à l'erreur : toute simulation ratée est une donnée",
             "Feedback bienveillant : 1 force + 1 axe + 1 conseil",
             "Téléphones en mode silencieux pendant les exercices",
             "Ponctualité : le groupe démarre à l'heure prévue",
             "Pas de moquerie : la vulnérabilité crée la progression"],
            "Conformité CMC 2025 — Niveau 3",
            ["✓ Durée ≥ 35h (FIM = 35h contact + 6h terrain)",
             "✓ LFI individuel tracé et archivé",
             "✓ Évaluation en situation de travail (J5 simulation)",
             "⚠ Jury externe — à intégrer en V7",
             "⚠ Reconnaissance UM6P CCE — dossier à déposer"])
        add_text_box(sl,
            "Ratio pratique/théorie V6 : 65% pratique / 35% théorie  (V5 : 45%/55%) — Standard UM6P atteint",
            0.3, 5.9, 12.7, 0.4, font_size=10, bold=True,
            color=TEAL_DARK, align=PP_ALIGN.CENTER)
    content_slide(s, "Règles de vie & Alignement CMC 2025", body)

    return save(prs, f"FIM_Book0_FilConducteur_{VERSION}.pptx")


# ══════════════════════════════════════════════════════════════════════════
#  BOOK 1 — SYNOPSIS & OBJECTIFS
# ══════════════════════════════════════════════════════════════════════════

def build_book1():
    prs = new_prs()
    SL = prs.slide_layouts[6]

    s = prs.slides.add_slide(SL)
    header_slide(s,
        "Synopsis & Objectifs",
        "Cadrage stratégique de la formation FIM · Référentiel de compétences V6",
        tag="BOOK 1 · UM6P/APC")
    add_text_box(s,
        "Ce module définit la vision, les 7 compétences cibles (C1–C7)\n"
        "et les indicateurs de réussite du programme FIM.",
        0.5, 2.8, 12, 0.8, font_size=14, color=DARK)

    # Slide 2 : Référentiel compétences
    s = prs.slides.add_slide(SL)
    def body(sl):
        add_rect(sl, 0, 0, 13.33, 7.5, fill_color=LIGHT_BG)
        add_rect(sl, 0, 0, 13.33, 0.75, fill_color=TEAL_DARK)
        add_rect(sl, 0, 0.73, 13.33, 0.04, fill_color=GOLD)
        add_text_box(sl, "Référentiel de Compétences FIM V6 — C1 à C7",
                     0.3, 0.08, 12.7, 0.62, font_size=18, bold=True, color=WHITE)
        comps = [
            ("C1", "Accroche & ouverture", "Bloom 3", "Initier un contact naturel en < 15 sec"),
            ("C2", "Storytelling impact", "Bloom 3", "Narrer une histoire bénéficiaire qui suscite l'émotion"),
            ("C3", "Écoute active & empathie", "Bloom 4", "Analyser les signaux verbaux et non-verbaux"),
            ("C4", "Traitement des objections", "Bloom 3", "Appliquer CROC face aux 12 objections types"),
            ("C5", "Engagement au don régulier", "Bloom 3", "Conclure un engagement de prélèvement mensuel"),
            ("C6", "Gestion émotionnelle", "Bloom 4", "Analyser et réguler sa propre réaction au refus"),
            ("C7", "Éthique & conformité", "Bloom 2", "Expliquer les règles déontologiques fundraiser"),
        ]
        colors_alt = [TEAL, TEAL_DARK]
        y = 0.9
        for i, (code, name, bloom, desc) in enumerate(comps):
            bg = colors_alt[i % 2]
            add_rect(sl, 0.3, y, 1.0, 0.5, fill_color=bg)
            add_text_box(sl, code, 0.32, y+0.08, 0.96, 0.38,
                         font_size=14, bold=True, color=GOLD, align=PP_ALIGN.CENTER)
            add_rect(sl, 1.32, y, 5.5, 0.5,
                     fill_color=WHITE, line_color=bg, line_width=Pt(1))
            add_text_box(sl, name, 1.4, y+0.04, 5.3, 0.42,
                         font_size=11, bold=True, color=TEAL_DARK)
            add_rect(sl, 6.84, y, 1.4, 0.5, fill_color=GOLD)
            add_text_box(sl, bloom, 6.86, y+0.08, 1.36, 0.38,
                         font_size=9, bold=True, color=TEAL_DARK, align=PP_ALIGN.CENTER)
            add_rect(sl, 8.26, y, 4.77, 0.5,
                     fill_color=LIGHT_BG, line_color=bg, line_width=Pt(1))
            add_text_box(sl, desc, 8.34, y+0.06, 4.6, 0.42,
                         font_size=10, color=DARK, italic=True)
            y += 0.53
        add_text_box(sl, f"SHY-Performance · {VERSION}",
                     0, 7.1, 13.33, 0.4, font_size=8,
                     color=RGBColor(0x88,0x88,0x88), align=PP_ALIGN.CENTER)
    body(s)

    # Slide 3 : Objectifs SMART
    s = prs.slides.add_slide(SL)
    def body(sl):
        two_col(sl,
            "Objectifs globaux SMART (Bloom 2–4)",
            ["Nommer les 7 compétences clés du fundraiser FIM (C1–C7)",
             "Expliquer le lien entre objectif personnel et mission UNICEF",
             "Distinguer les niveaux de maîtrise attendus à J1, J3 et J5",
             "Analyser son propre profil via l'auto-positionnement initial"],
            "Activités d'ancrage V6",
            ["Carte mentale : 'Pourquoi je suis ici' → partage groupe",
             "Vision board collectif : 'Le fundraiser idéal 2026'",
             "Lecture commentée du référentiel C1–C7",
             "Discussion : 'Quelle compétence est ma plus grande force ?'"])
        eval_box(sl,
            ["Auto-positionnement C1–C7 (0/1/2/3) → LFI page 1",
             "Engagement écrit sur 1 compétence prioritaire à développer",
             "NPS d'anticipation : 'Sur 10, je me sens prêt(e) à quel niveau ?'"],
            label="Évaluation Book 1 — Ouverture de formation")
        bloom_bar(sl)
    content_slide(s, "Objectifs SMART & Activités d'ancrage", body)

    # Slide 4 : Indicateurs de réussite
    s = prs.slides.add_slide(SL)
    def body(sl):
        add_rect(sl, 0, 0, 13.33, 7.5, fill_color=LIGHT_BG)
        add_rect(sl, 0, 0, 13.33, 0.75, fill_color=TEAL_DARK)
        add_rect(sl, 0, 0.73, 13.33, 0.04, fill_color=GOLD)
        add_text_box(sl, "Indicateurs de Réussite & Modèle Kirkpatrick",
                     0.3, 0.08, 12.7, 0.62, font_size=18, bold=True, color=WHITE)
        rows = [
            ("Niv. 1 — Réaction",    "Satisfaction apprenant",   "Questionnaire fin J chaque jour", "NPS ≥ 8/10"),
            ("Niv. 2 — Apprentissage","Compétences acquises",     "Grilles C1–C7 fin J3 et J5",     "Score ≥ 71% (20/28)"),
            ("Niv. 3 — Comportement", "Transfert terrain",        "Observation terrain J+15",        "≥ 5 critères sur 7 observés"),
            ("Niv. 4 — Résultats",    "Impact KPI fundraising",   "Suivi conversions J+30 & J+90",   "≥ 1 don/heure terrain"),
        ]
        headers = ["Niveau Kirkpatrick", "Dimension", "Outil de mesure", "Seuil de réussite"]
        col_w = [3.0, 3.0, 4.0, 2.7]
        x_starts = [0.3, 3.32, 6.34, 10.36]
        y = 0.9
        for j, (h, w) in enumerate(zip(headers, col_w)):
            add_rect(sl, x_starts[j], y, w, 0.35, fill_color=TEAL)
            add_text_box(sl, h, x_starts[j]+0.08, y+0.06, w-0.16, 0.26,
                         font_size=10, bold=True, color=WHITE)
        y += 0.38
        for i, row in enumerate(rows):
            bg = LIGHT_BG if i % 2 == 0 else WHITE
            for j, (cell, w) in enumerate(zip(row, col_w)):
                add_rect(sl, x_starts[j], y, w, 0.6,
                         fill_color=bg, line_color=TEAL, line_width=Pt(0.5))
                add_text_box(sl, cell, x_starts[j]+0.08, y+0.06, w-0.16, 0.5,
                             font_size=10, color=DARK if j > 0 else TEAL_DARK,
                             bold=(j == 0))
            y += 0.62
        add_text_box(sl, f"SHY-Performance · {VERSION}",
                     0, 7.1, 13.33, 0.4, font_size=8,
                     color=RGBColor(0x88,0x88,0x88), align=PP_ALIGN.CENTER)
    body(s)

    return save(prs, f"FIM_Book1_Synopsis_Objectifs_{VERSION}.pptx")


# ══════════════════════════════════════════════════════════════════════════
#  BOOK J1 — FONDATIONS & MISSION
# ══════════════════════════════════════════════════════════════════════════

def build_bookJ1():
    prs = new_prs()
    SL = prs.slide_layouts[6]

    s = prs.slides.add_slide(SL)
    header_slide(s, "Fondations & Mission",
        "Ancrage identitaire du fundraiser · Valeurs · Posture · Raison d'agir",
        tag="BOOK J1 · Journée 1")

    # Slide 2 : Objectifs Kolb
    s = prs.slides.add_slide(SL)
    def body(sl):
        two_col(sl,
            "Objectifs V6 — Bloom Niv. 1–3",
            ["Nommer les missions et valeurs de l'UNICEF au Maroc",
             "Expliquer le cadre éthique et juridique de la collecte",
             "Démontrer la posture physique et vocale d'un fundraiser",
             "Différencier don ponctuel et engagement mensuel"],
            "Contenu clé",
            ["Qu'est-ce qu'un fundraiser de rue ? Rôle et légitimité",
             "Mission UNICEF Maroc — chiffres 2024–2026",
             "Cadre légal Maroc (Loi 75-00) + déontologie",
             "Posture corporelle, contact visuel, ton de voix"])
        kolb_box(sl,
            ["Visionnage témoignage bénéficiaire UNICEF (3 min) + réaction individuelle notée",
             "Débriefing : 'Qu'avez-vous ressenti ? Qu'auriez-vous dit ?'",
             "Apport : structure UNICEF, chiffres clés, cadre légal Maroc",
             "Jeu de rôle : 'Expliquez la mission UNICEF en 60 secondes'"],
            y=4.55, h=1.9)
        eval_box(sl,
            ["Pitch 60s — grille C1+C2+C7 (formateur + auto-éval)",
             "Quiz 5Q chiffres UNICEF — seuil 60%",
             "Réflexion écrite 5 lignes dans le LFI"],
            label="Évaluation formative J1")
    content_slide(s, "Objectifs & Séquence Kolb — Journée 1", body)

    # Slide 3 : UNICEF chiffres
    s = prs.slides.add_slide(SL)
    def body(sl):
        add_rect(sl, 0, 0, 13.33, 7.5, fill_color=LIGHT_BG)
        add_rect(sl, 0, 0, 13.33, 0.75, fill_color=TEAL_DARK)
        add_rect(sl, 0, 0.73, 13.33, 0.04, fill_color=GOLD)
        add_text_box(sl, "UNICEF Maroc & International — Chiffres clés 2024–2026",
                     0.3, 0.08, 12.7, 0.62, font_size=18, bold=True, color=WHITE)
        facts = [
            ("190+", "pays où l'UNICEF est présent"),
            ("200M+", "enfants aidés chaque année"),
            ("46M€", "budget UNICEF Maroc 2024"),
            ("3,2M", "enfants bénéficiaires au Maroc"),
            ("72%", "fonds issus des dons privés"),
            ("15 DH/jour", "coût d'1 kit de survie enfant"),
        ]
        x = 0.3
        for num, label in facts:
            add_rect(sl, x, 0.95, 2.1, 1.1, fill_color=TEAL)
            add_text_box(sl, num, x+0.1, 0.98, 1.9, 0.65,
                         font_size=26, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
            add_text_box(sl, label, x+0.1, 1.62, 1.9, 0.4,
                         font_size=9, color=RGBColor(0xDD,0xEE,0xEE),
                         align=PP_ALIGN.CENTER, wrap=True)
            x += 2.15
        add_text_box(sl,
            "Message clé fundraiser : 'Un don de 100 DH/mois = 6 enfants protégés par an'\n"
            "Cadre légal : Loi 75-00 sur les associations · Agrément préfectoral · Reçu fiscal",
            0.3, 2.2, 12.7, 0.9, font_size=12, color=DARK)
        add_text_box(sl, "Activité — L'ascenseur de mission (Bloom 3)",
                     0.3, 3.25, 12.7, 0.3, font_size=12, bold=True, color=TEAL_DARK)
        steps = [
            "1. Lire seul le brief mission (3 min)",
            "2. Formuler son pitch 60s (5 min)",
            "3. Présenter à son binôme",
            "4. Feedback pair : 1 force + 1 conseil",
            "5. Améliorer et re-présenter (5 min)",
        ]
        x = 0.3
        for step in steps:
            add_rect(sl, x, 3.6, 2.45, 0.75,
                     fill_color=WHITE, line_color=TEAL, line_width=Pt(1))
            add_text_box(sl, step, x+0.1, 3.65, 2.25, 0.65,
                         font_size=9, color=DARK, wrap=True)
            x += 2.5
        add_text_box(sl,
            "Debate (Bloom 5) : 'Le don mensuel régulier est-il éthique ?'\n"
            "2 équipes · 5 min préparation · 10 min débat · 5 min synthèse formateur",
            0.3, 4.5, 12.7, 0.8, font_size=11, color=DARK)
        add_text_box(sl,
            "🔗 Lien identitaire : cartographie mentale valeurs personnelles ↔ valeurs UNICEF",
            0.3, 5.45, 12.7, 0.35, font_size=11, bold=True, color=TEAL_DARK)
        add_text_box(sl, f"SHY-Performance · {VERSION}",
                     0, 7.1, 13.33, 0.4, font_size=8,
                     color=RGBColor(0x88,0x88,0x88), align=PP_ALIGN.CENTER)
    body(s)

    return save(prs, f"FIM_BookJ1_Fondations_Mission_{VERSION}.pptx")


# ══════════════════════════════════════════════════════════════════════════
#  BOOK J2 — MONDE ASSOCIATIF
# ══════════════════════════════════════════════════════════════════════════

def build_bookJ2():
    prs = new_prs()
    SL = prs.slide_layouts[6]

    s = prs.slides.add_slide(SL)
    header_slide(s, "Monde Associatif",
        "Contextualisation · Écosystème ONG Maroc & international · Économie du don",
        tag="BOOK J2 · Journée 2")

    s = prs.slides.add_slide(SL)
    def body(sl):
        two_col(sl,
            "Objectifs V6 — Bloom Niv. 2–4",
            ["Expliquer le modèle économique du don régulier vs ponctuel",
             "Distinguer les acteurs du secteur associatif marocain",
             "Analyser pourquoi le don régulier est le plus efficace",
             "Identifier le cadre juridique (Loi 75-00 Maroc)"],
            "Apports théoriques clés",
            ["Loi 75-00 sur les associations au Maroc",
             "Économie du don : théorie de la réciprocité (Mauss)",
             "Données : marché fundraising MENA 2024–2026",
             "UNICEF Maroc : budget, programmes actifs, bénéficiaires",
             "Don régulier vs ponctuel : LTV donateur × 8"])
        kolb_box(sl,
            ["Analyse de 3 campagnes de fundraising (succès / échec)",
             "'Qu'est-ce qui a fait la différence ?' — liste collective",
             "Cadre juridique associatif + économie du don (Mauss)",
             "Simulation budgétaire : 100 donneurs × 100 DH/mois = ?"],
            y=4.55, h=1.9)
        eval_box(sl,
            ["Restitution orale étude de cas (5 min/groupe)",
             "Fiche 'Monde associatif' complétée individuellement",
             "1 question réflexion LFI sur l'objection 'méfiance ONG'"])
    content_slide(s, "Objectifs & Séquence Kolb — Journée 2", body)

    s = prs.slides.add_slide(SL)
    def body(sl):
        add_rect(sl, 0, 0, 13.33, 7.5, fill_color=LIGHT_BG)
        add_rect(sl, 0, 0, 13.33, 0.75, fill_color=TEAL_DARK)
        add_rect(sl, 0, 0.73, 13.33, 0.04, fill_color=GOLD)
        add_text_box(sl, "Étude de Cas — 3 Campagnes Fundraising",
                     0.3, 0.08, 12.7, 0.62, font_size=18, bold=True, color=WHITE)
        cases = [
            ("CAS A — SUCCÈS",
             "Greenpeace Maroc 2023\nObjectif : 500 donateurs/mois\nRésultat : 720 donateurs (+44%)\nFacteurs : script émotionnel fort,\nformation 5J intensive, suivi J+30",
             GREEN_OK),
            ("CAS B — ÉCHEC PARTIEL",
             "ONG X Rabat 2022\nObjectif : 300 donateurs/mois\nRésultat : 180 donateurs (-40%)\nFacteurs : formation 2J seulement,\npas de gestion des objections",
             RGBColor(0xD2,0x7A,0x42)),
            ("CAS C — RETOURNEMENT",
             "Croix Rouge Maroc 2024\nObjectif : 200 donateurs/mois\nRésultat : 95 → 310 après FIM\nFacteurs : refonte programme,\nintégration Kolb + Kirkpatrick",
             TEAL),
        ]
        x = 0.3
        for title, content, color in cases:
            add_rect(sl, x, 0.9, 4.1, 0.38, fill_color=color)
            add_text_box(sl, title, x+0.1, 0.94, 3.9, 0.3,
                         font_size=11, bold=True, color=WHITE)
            add_rect(sl, x, 1.28, 4.1, 2.4,
                     fill_color=WHITE, line_color=color, line_width=Pt(1.5))
            add_text_box(sl, content, x+0.15, 1.35, 3.8, 2.2,
                         font_size=10, color=DARK, wrap=True)
            x += 4.25
        add_text_box(sl, "Simulation budgétaire (Bloom 3) — Activité collective",
                     0.3, 3.9, 12.7, 0.3, font_size=12, bold=True, color=TEAL_DARK)
        add_text_box(sl,
            "Si 100 fundraisers actifs × 1 don régulier/heure × 6h/jour × 20 jours/mois :\n"
            "= 12 000 donateurs potentiels/mois · Montant moyen : 80 DH/mois\n"
            "= 960 000 DH collectés/mois · Soit 11,5 M DH/an\n"
            "Impact UNICEF : vaccination de 50 000 enfants · alimentation 8 000 nourrissons",
            0.3, 4.25, 12.7, 1.5, font_size=11, color=DARK)
        add_text_box(sl,
            "🔑 Message clé : chaque fundraiser actif = entre 150 et 300 enfants aidés par an",
            0.3, 5.85, 12.7, 0.4, font_size=12, bold=True,
            color=TEAL_DARK, align=PP_ALIGN.CENTER)
        add_text_box(sl, f"SHY-Performance · {VERSION}",
                     0, 7.1, 13.33, 0.4, font_size=8,
                     color=RGBColor(0x88,0x88,0x88), align=PP_ALIGN.CENTER)
    body(s)

    return save(prs, f"FIM_BookJ2_MondeAssociatif_{VERSION}.pptx")


# ══════════════════════════════════════════════════════════════════════════
#  BOOK J3 — LECTURE & SCRIPT
# ══════════════════════════════════════════════════════════════════════════

def build_bookJ3():
    prs = new_prs()
    SL = prs.slide_layouts[6]

    s = prs.slides.add_slide(SL)
    header_slide(s, "Lecture & Script",
        "Maîtrise du discours commercial-humanitaire · Les 5 étapes · Storytelling",
        tag="BOOK J3 · Journée 3")

    s = prs.slides.add_slide(SL)
    def body(sl):
        two_col(sl,
            "Objectifs V6 — Bloom Niv. 3–4",
            ["Utiliser les 5 étapes du script FIM dans l'ordre et avec naturel",
             "Adapter le storytelling au profil du passant",
             "Analyser sa propre performance via la grille d'observation",
             "Démontrer une accroche naturelle en < 15 secondes"],
            "Les 5 étapes du Script FIM V6",
            ["① Accroche — contact visuel, sourire, question ouverte (10–15 sec)",
             "② Mise en contexte — 'Je travaille avec l'UNICEF…' (20 sec)",
             "③ Impact Story — récit bénéficiaire, émotion + faits (45–60 sec)",
             "④ Proposition — montant, avantages donateur (30 sec)",
             "⑤ Engagement — signature, remerciement chaleureux (20 sec)"])
        kolb_box(sl,
            ["Démo formateur LIVE sur passant simulé · apprenants observent avec grille",
             "'Qu'avez-vous observé ?' → liste comportements efficaces",
             "Déconstruction des 5 étapes, ancres mnémotechniques, langage corporel",
             "Trinômes 90 min : fundraiser / passant / observateur · rotation 15 min"],
            y=4.55, h=1.9)
        eval_box(sl,
            ["Grille certification partielle C1+C2+C5 (formateur)",
             "Auto-évaluation vidéo si consentement apprenant",
             "Score seuil : 12/20 pour passer à J4"])
    content_slide(s, "Objectifs & Séquence Kolb — Journée 3", body)

    # Slide 3 : Script détaillé
    s = prs.slides.add_slide(SL)
    def body(sl):
        add_rect(sl, 0, 0, 13.33, 7.5, fill_color=LIGHT_BG)
        add_rect(sl, 0, 0, 13.33, 0.75, fill_color=TEAL_DARK)
        add_rect(sl, 0, 0.73, 13.33, 0.04, fill_color=GOLD)
        add_text_box(sl, "Script FIM V6 — Détail des 5 étapes + Adaptation profil",
                     0.3, 0.08, 12.7, 0.62, font_size=18, bold=True, color=WHITE)
        steps = [
            ("① ACCROCHE", "10–15 sec", TEAL,
             "Contact visuel · Sourire authentique · Question ouverte non intrusive\n"
             "Ex : 'Bonjour ! Vous avez 2 minutes pour les enfants ?'\n"
             "⚠ Éviter : 'Excusez-moi de vous déranger…'"),
            ("② CONTEXTE", "20 sec", TEAL_DARK,
             "Identification claire + légitimité\n"
             "Ex : 'Je travaille avec l'UNICEF — nous collectons des fonds pour les enfants en danger.'\n"
             "Montrer le badge / la carte officielle"),
            ("③ IMPACT STORY", "45–60 sec", RGBColor(0x4A,0x90,0x70),
             "Histoire d'1 enfant bénéficiaire spécifique · Prénom + pays + situation\n"
             "Ex : 'Aisha, 4 ans, Yémen — sans votre aide, elle n'avait pas accès à l'eau potable.'\n"
             "Faits + chiffre concret + résolution par le don"),
            ("④ PROPOSITION", "30 sec", RGBColor(0xC9,0xA8,0x4C),
             "Montant suggéré adapté au profil observé (100–300 DH/mois)\n"
             "Bénéfices donateur : reçu fiscal, rapport annuel, impact tracé\n"
             "Ex : '100 DH/mois = 3 DH/jour = 1 café. Impact : 6 enfants protégés/an.'"),
            ("⑤ ENGAGEMENT", "20 sec", RGBColor(0xD2,0x7A,0x42),
             "Formulaire clair · Accompagner la signature · Remerciement sincère\n"
             "Ex : 'Merci ! Vous allez recevoir votre confirmation par SMS dans 24h.'\n"
             "Remettre la brochure UNICEF + coordonnées contact"),
        ]
        x = 0.3
        for title, duration, color, content in steps:
            add_rect(sl, x, 0.9, 2.5, 0.35, fill_color=color)
            add_text_box(sl, f"{title}  [{duration}]", x+0.08, 0.93, 2.34, 0.28,
                         font_size=9, bold=True, color=WHITE)
            add_rect(sl, x, 1.25, 2.5, 4.5,
                     fill_color=WHITE, line_color=color, line_width=Pt(1.5))
            add_text_box(sl, content, x+0.1, 1.3, 2.3, 4.2,
                         font_size=9, color=DARK, wrap=True)
            x += 2.56
        add_text_box(sl,
            "Durée totale script standard : 2 min 15 sec · Version express (passant pressé) : 55 sec",
            0.3, 5.9, 12.7, 0.35, font_size=10, bold=True,
            color=TEAL_DARK, align=PP_ALIGN.CENTER)
        add_text_box(sl, f"SHY-Performance · {VERSION}",
                     0, 7.1, 13.33, 0.4, font_size=8,
                     color=RGBColor(0x88,0x88,0x88), align=PP_ALIGN.CENTER)
    body(s)

    return save(prs, f"FIM_BookJ3_Lecture_Script_{VERSION}.pptx")


# ══════════════════════════════════════════════════════════════════════════
#  BOOK J4 — OBJECTIONS
# ══════════════════════════════════════════════════════════════════════════

def build_bookJ4():
    prs = new_prs()
    SL = prs.slide_layouts[6]

    s = prs.slides.add_slide(SL)
    header_slide(s, "Gestion des Objections",
        "Écoute active · Méthode CROC · 12 objections-types · Jeux de rôle",
        tag="BOOK J4 · Journée 4")

    s = prs.slides.add_slide(SL)
    def body(sl):
        two_col(sl,
            "Objectifs V6 — Bloom Niv. 3–5",
            ["Appliquer la méthode CROC face aux 12 objections fréquentes",
             "Distinguer l'objection réelle de l'objection prétexte",
             "Évaluer la qualité de ses propres réponses (auto-éval)",
             "Adapter la réponse selon le profil émotionnel du passant"],
            "Méthode CROC",
            ["C — Considérer : accusé réception sincère sans justifier",
             "R — Reformuler : 'Si je comprends bien, vous voulez dire…'",
             "O — Organiser : réponse structurée = preuve + exemple + chiffre",
             "C — Conclure : retour vers la proposition + silence actif"])
        kolb_box(sl,
            ["Analyse de 5 extraits vidéo (bonne/mauvaise gestion d'objection)",
             "'Patterns observés ?' → taxonomie collective des objections",
             "CROC + CAP + 12 objections types avec réponses modèles",
             "Battle des objections : tournoi brackets · 2 objections tirées au sort"],
            y=4.55, h=1.9)
        eval_box(sl,
            ["Grille C3+C4 — 6 critères — formateur + pair",
             "Réflexion LFI : 'Quelle objection me challenge le plus et pourquoi ?'",
             "Score seuil : 14/20 pour validation C3+C4"])
    content_slide(s, "Objectifs & Séquence Kolb — Journée 4", body)

    # Slide 3 : Les 12 objections
    s = prs.slides.add_slide(SL)
    def body(sl):
        add_rect(sl, 0, 0, 13.33, 7.5, fill_color=LIGHT_BG)
        add_rect(sl, 0, 0, 13.33, 0.75, fill_color=TEAL_DARK)
        add_rect(sl, 0, 0.73, 13.33, 0.04, fill_color=GOLD)
        add_text_box(sl, "Les 12 Objections-Types FIM V6 — Réponses CROC",
                     0.3, 0.08, 12.7, 0.62, font_size=18, bold=True, color=WHITE)
        objections = [
            ("O1", "Je n'ai pas les moyens", "Financière", "3 DH/jour = 1 café. Ce montant change 6 vies/an."),
            ("O2", "Je ne fais pas confiance", "Méfiance",  "Rapport annuel public + badge officiel UNICEF"),
            ("O3", "J'ai déjà fait un don",   "Engagé",    "Valoriser + proposer le don régulier complémentaire"),
            ("O4", "Je dois réfléchir",       "Report",    "Clarifier le frein + ancrer urgence humanitaire"),
            ("O5", "Combien vous gagnez ?",   "Méfiance",  "Transparence totale : mission + emploi déclaré"),
            ("O6", "Je donne à d'autres ONG", "Concurrent","Complémentarité : chaque cause est différente"),
            ("O7", "Ce n'est pas le bon moment","Temporel","Urgence + flexibilité date de prélèvement"),
            ("O8", "L'argent n'arrive pas",   "Méfiance",  "Traçabilité : rapports publiés + auditeurs ext."),
            ("O9", "Je suis pressé",          "Temporel",  "Script express 45 sec : accroche + chiffre + closing"),
            ("O10","Mon conjoint décide",     "Autorité",  "Impliquer le couple + envoyer info par SMS/email"),
            ("O11","Pas de prélèvements",     "Processus", "Arrêt en 1 clic · Modification sans engagement"),
            ("O12","Politique / Religion",    "Valeurs",   "UNICEF = neutralité totale · Focus : droits enfants"),
        ]
        headers = ["#", "Objection", "Type", "Réponse CROC résumée"]
        col_x = [0.3, 0.72, 5.8, 7.55]
        col_w = [0.4, 5.05, 1.7, 5.7]
        y = 0.88
        for j, (h, x, w) in enumerate(zip(headers, col_x, col_w)):
            add_rect(sl, x, y, w, 0.3, fill_color=TEAL)
            add_text_box(sl, h, x+0.05, y+0.04, w-0.1, 0.24,
                         font_size=9, bold=True, color=WHITE)
        y += 0.32
        type_colors = {
            "Financière": RGBColor(0x4A,0x90,0x70),
            "Méfiance":   RED_SOFT,
            "Engagé":     TEAL,
            "Report":     RGBColor(0xD2,0x7A,0x42),
            "Concurrent": RGBColor(0x7A,0x5A,0xC0),
            "Temporel":   RGBColor(0xC9,0xA8,0x4C),
            "Autorité":   TEAL_DARK,
            "Processus":  RGBColor(0x55,0x88,0xAA),
            "Valeurs":    RGBColor(0x88,0x44,0x88),
        }
        for i, (num, obj, typ, resp) in enumerate(objections):
            bg = LIGHT_BG if i % 2 == 0 else WHITE
            row_h = 0.48
            add_rect(sl, 0.3,  y, 0.4,  row_h, fill_color=TEAL)
            add_text_box(sl, num, 0.32, y+0.08, 0.36, 0.32,
                         font_size=9, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
            add_rect(sl, 0.72, y, 5.05, row_h, fill_color=bg)
            add_text_box(sl, obj, 0.8, y+0.08, 4.9, 0.36,
                         font_size=9, color=DARK)
            tc = type_colors.get(typ, TEAL)
            add_rect(sl, 5.8,  y, 1.7,  row_h, fill_color=tc)
            add_text_box(sl, typ, 5.85, y+0.1, 1.6, 0.32,
                         font_size=8, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
            add_rect(sl, 7.55, y, 5.7,  row_h, fill_color=bg)
            add_text_box(sl, resp, 7.63, y+0.06, 5.55, 0.38,
                         font_size=9, color=DARK, italic=True)
            y += row_h + 0.01
        add_text_box(sl, f"SHY-Performance · {VERSION}",
                     0, 7.1, 13.33, 0.4, font_size=8,
                     color=RGBColor(0x88,0x88,0x88), align=PP_ALIGN.CENTER)
    body(s)

    return save(prs, f"FIM_BookJ4_Objections_{VERSION}.pptx")


# ══════════════════════════════════════════════════════════════════════════
#  BOOK J5 — SYNTHÈSE & SIMULATIONS
# ══════════════════════════════════════════════════════════════════════════

def build_bookJ5():
    prs = new_prs()
    SL = prs.slide_layouts[6]

    s = prs.slides.add_slide(SL)
    header_slide(s, "Synthèse & Simulations",
        "Certification opérationnelle · Grille 28 pts · Plan d'Action Personnel (PAP)",
        tag="BOOK J5 · Journée 5 — Certification")

    # Slide 2 : Programme J5
    s = prs.slides.add_slide(SL)
    def body(sl):
        add_rect(sl, 0, 0, 13.33, 7.5, fill_color=LIGHT_BG)
        add_rect(sl, 0, 0, 13.33, 0.75, fill_color=TEAL_DARK)
        add_rect(sl, 0, 0.73, 13.33, 0.04, fill_color=GOLD)
        add_text_box(sl, "Programme Journée 5 — Certification FIM V6",
                     0.3, 0.08, 12.7, 0.62, font_size=18, bold=True, color=WHITE)
        programme = [
            ("08h30", "09h00", "Révision flash",
             "Quiz 7 compétences C1–C7 · rappel des seuils · météo émotionnelle"),
            ("09h00", "11h00", "Simulation terrain × 2",
             "Round 1 : passant 'neutre' · Round 2 : passant 'difficile avec 3 objections'\n"
             "Formateur joue le rôle du passant · apprenants observent avec grille"),
            ("11h00", "11h30", "Debriefing collectif",
             "Célébration des progrès individuels · 3 apprentissages collectifs retenus"),
            ("11h30", "12h30", "Évaluation sommative",
             "Grille 28 pts — formateur + superviseur · entretien individuel 5 min"),
            ("13h30", "14h30", "Construction PAP",
             "Plan d'Action Personnel : 4 objectifs sur 4 semaines · 1 mesure/semaine"),
            ("14h30", "15h00", "Clôture ritualisée",
             "Remise attestations · Engagement public 1 min/apprenant · Photos groupe"),
        ]
        y = 0.9
        for start, end, title, detail in programme:
            add_rect(sl, 0.3, y, 1.5, 0.75, fill_color=TEAL)
            add_text_box(sl, f"{start}\n{end}", 0.32, y+0.08, 1.46, 0.6,
                         font_size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
            add_rect(sl, 1.82, y, 2.8, 0.75, fill_color=TEAL_DARK)
            add_text_box(sl, title, 1.9, y+0.18, 2.65, 0.42,
                         font_size=12, bold=True, color=WHITE)
            add_rect(sl, 4.64, y, 8.6, 0.75,
                     fill_color=WHITE, line_color=TEAL, line_width=Pt(0.8))
            add_text_box(sl, detail, 4.74, y+0.06, 8.4, 0.65,
                         font_size=9, color=DARK, wrap=True)
            y += 0.78
        add_text_box(sl, f"SHY-Performance · {VERSION}",
                     0, 7.1, 13.33, 0.4, font_size=8,
                     color=RGBColor(0x88,0x88,0x88), align=PP_ALIGN.CENTER)
    body(s)

    # Slide 3 : Grille de certification
    s = prs.slides.add_slide(SL)
    def body(sl):
        add_rect(sl, 0, 0, 13.33, 7.5, fill_color=LIGHT_BG)
        add_rect(sl, 0, 0, 13.33, 0.75, fill_color=TEAL_DARK)
        add_rect(sl, 0, 0.73, 13.33, 0.04, fill_color=GOLD)
        add_text_box(sl, "Grille de Certification FIM V6 — 7 Critères / 28 Points",
                     0.3, 0.08, 12.7, 0.62, font_size=18, bold=True, color=WHITE)
        criteria = [
            ("C1", "Accroche naturelle ≤ 15 sec",
             "Contact visuel, sourire, question ouverte, aisance"),
            ("C2", "Script fluide et personnalisé",
             "5 étapes dans l'ordre, langage adapté au passant"),
            ("C3", "Impact Story (émotion + faits)",
             "Récit bénéficiaire spécifique, chiffre concret, résolution"),
            ("C4", "Traitement ≥ 2 objections CROC",
             "Considérer, reformuler, organiser, conclure"),
            ("C5", "Engagement au don / closing",
             "Proposition montant, signature, remerciement"),
            ("C6", "Posture éthique & conformité",
             "Badge visible, pas de pression, règles déontologiques"),
            ("C7", "Gestion émotionnelle & résilience",
             "Maintien posture positive après refus, 0 découragement visible"),
        ]
        col_x = [0.3, 0.75, 4.85, 10.35, 11.15, 11.95, 12.75]
        col_w = [0.43, 4.08, 5.48, 0.78,  0.78,  0.78,  0.5]
        headers = ["#", "Critère", "Indicateurs observables", "1", "2", "3", "4"]
        y = 0.88
        for j, (h, x, w) in enumerate(zip(headers, col_x, col_w)):
            add_rect(sl, x, y, w, 0.3, fill_color=TEAL)
            add_text_box(sl, h, x+0.05, y+0.04, w-0.1, 0.24,
                         font_size=9, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        y += 0.32
        for i, (code, criterion, indicators) in enumerate(criteria):
            bg = LIGHT_BG if i % 2 == 0 else WHITE
            rh = 0.55
            add_rect(sl, 0.3,  y, 0.43, rh, fill_color=TEAL)
            add_text_box(sl, code, 0.32, y+0.12, 0.39, 0.35,
                         font_size=10, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
            add_rect(sl, 0.75, y, 4.08, rh, fill_color=bg)
            add_text_box(sl, criterion, 0.83, y+0.1, 3.92, 0.42,
                         font_size=10, bold=True, color=TEAL_DARK)
            add_rect(sl, 4.85, y, 5.48, rh, fill_color=bg)
            add_text_box(sl, indicators, 4.93, y+0.06, 5.32, 0.46,
                         font_size=9, color=DARK, italic=True)
            for k, (x, w) in enumerate(zip(col_x[3:], col_w[3:])):
                add_rect(sl, x, y, w, rh,
                         fill_color=WHITE, line_color=TEAL, line_width=Pt(0.5))
                add_text_box(sl, str(k+1), x+0.1, y+0.12, w-0.2, 0.35,
                             font_size=9, color=RGBColor(0xBB,0xBB,0xBB),
                             align=PP_ALIGN.CENTER)
            y += rh + 0.02
        # Total row
        add_rect(sl, 0.3, y, 12.95, 0.42, fill_color=TEAL_DARK)
        add_text_box(sl, "TOTAL  /28", 0.4, y+0.08, 10.5, 0.3,
                     font_size=11, bold=True, color=WHITE)
        add_rect(sl, 11.15, y, 2.1, 0.42,
                 fill_color=WHITE, line_color=GOLD, line_width=Pt(2))
        add_text_box(sl, "__ / 28", 11.2, y+0.08, 2.0, 0.3,
                     font_size=12, bold=True, color=TEAL_DARK, align=PP_ALIGN.CENTER)
        y += 0.46
        add_text_box(sl,
            "✅ Certifié FIM : ≥ 20/28 (71%)     "
            "⚠ Formation complémentaire : 14–19/28     "
            "🔄 Reconduction J+15 : < 14/28",
            0.3, y+0.06, 12.7, 0.42, font_size=10, bold=True,
            color=TEAL_DARK, align=PP_ALIGN.CENTER)
        add_text_box(sl, f"SHY-Performance · {VERSION}",
                     0, 7.1, 13.33, 0.4, font_size=8,
                     color=RGBColor(0x88,0x88,0x88), align=PP_ALIGN.CENTER)
    body(s)

    # Slide 4 : PAP
    s = prs.slides.add_slide(SL)
    def body(sl):
        two_col(sl,
            "Plan d'Action Personnel (PAP) — 4 semaines",
            ["S1 (J+7) : Maîtriser l'accroche en < 12 sec — chrono + feedback formateur",
             "S2 (J+14) : Traiter O1 et O2 sans hésitation — 0 rupture de discours",
             "S3 (J+21) : Atteindre ≥ 1 don/heure terrain",
             "S4 (J+30) : Révision PAP + définition objectifs M2 avec formateur"],
            "Suivi post-formation Kirkpatrick",
            ["J+15 : Observation terrain supervisée (grille Niv. 3) + feedback",
             "J+30 : Revue PAP — 4 objectifs atteints ? Adaptation si besoin",
             "J+90 : KPI terrain : taux conversion, montant moyen, rétention donateur",
             "Kirkpatrick Niv. 4 : impact collecte mesuré et rapporté à SHY-Performance"])
        eval_box(sl,
            ["Questionnaire satisfaction (NPS formateur, contenu, rythme, utilité)",
             "Grille 28 pts archivée dans le dossier apprenant CMC",
             "PAP signé apprenant + formateur → copie dans le LFI"],
            label="Évaluation sommative & Kirkpatrick Niv. 1–2")
        bloom_bar(sl)
    content_slide(s, "Plan d'Action Personnel & Suivi Post-Formation", body)

    return save(prs, f"FIM_BookJ5_Synthese_Simulations_{VERSION}.pptx")


# ══════════════════════════════════════════════════════════════════════════
#  MAIN
# ══════════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    print(f"\n🎓 Génération FIM {VERSION} — 7 modules\n")
    files = []
    files.append(build_book0())
    files.append(build_book1())
    files.append(build_bookJ1())
    files.append(build_bookJ2())
    files.append(build_bookJ3())
    files.append(build_bookJ4())
    files.append(build_bookJ5())
    print(f"\n✅ {len(files)} modules générés dans {OUT_DIR}")
    for f in files:
        size_kb = os.path.getsize(f) // 1024
        print(f"   {os.path.basename(f)}  ({size_kb} KB)")

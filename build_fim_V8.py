"""
FIM V8-01062026SHY-TE — Architecture pédagogique magistrale + interactive
- Quiz par module (5-7Q) + Quiz final (40Q)
- Activités pédagogiques progressives
- Structure magistrale UNICEF
- Toutes les corrections V7 intégrées
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
import os

TEAL   = RGBColor(0x00, 0x80, 0x80)
TEAL_D = RGBColor(0x00, 0x50, 0x50)
TEAL_M = RGBColor(0x00, 0x6A, 0x6A)
TEAL_L = RGBColor(0xF0, 0xF7, 0xF7)
TEAL_A = RGBColor(0x00, 0xAA, 0xAA)
WHITE  = RGBColor(0xFF, 0xFF, 0xFF)
DARK   = RGBColor(0x1A, 0x1A, 0x1A)

SZ_T   = 26
SZ_S   = 22
SZ_B   = 18
SZ_SM  = 14
SZ_XS  = 10

LOGO   = "/home/user/fatou/shy_logo_v2.png"
OUT    = "/home/user/fatou"
VER    = "V8-01062026SHY-TE"
CO     = "SHY-Performance"

def prs():
    p = Presentation()
    p.slide_width, p.slide_height = Inches(13.33), Inches(7.5)
    return p

def SL(p): return p.slide_layouts[6]

def R(s, x, y, w, h, fill=None, line=None, lw=Pt(0)):
    sh = s.shapes.add_shape(1, Inches(x), Inches(y), Inches(w), Inches(h))
    sh.line.width = lw
    if fill: sh.fill.solid(); sh.fill.fore_color.rgb = fill
    else: sh.fill.background()
    if line: sh.line.color.rgb = line
    else: sh.line.fill.background()
    return sh

def T(s, text, x, y, w, h, sz=SZ_B, bold=False, col=None, align=PP_ALIGN.LEFT, italic=False, wrap=True):
    col = col or DARK
    b = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    b.word_wrap = wrap
    tf = b.text_frame
    p2 = tf.paragraphs[0]
    p2.alignment = align
    r = p2.add_run()
    r.text = text
    r.font.size = Pt(sz)
    r.font.bold = bold
    r.font.italic = italic
    r.font.color.rgb = col
    return b

def LOGO_ADD(s, x=12.28, y=0.0, w=0.92, h=0.74):
    try:
        s.shapes.add_picture(LOGO, Inches(x), Inches(y), Inches(w), Inches(h))
    except:
        pass

def FOOTER(s):
    R(s, 0, 7.1, 13.33, 0.4, fill=TEAL_D)
    T(s, f"{CO}  ·  Formation FIM  ·  UNICEF France  ·  {VER}", 0, 7.13, 13.33, 0.3, sz=SZ_XS, col=WHITE, align=PP_ALIGN.CENTER)

def H_TITLE(s, title, sub="", tag=""):
    R(s, 0, 0, 13.33, 7.5, fill=TEAL_L)
    R(s, 0, 0, 13.33, 2.55, fill=TEAL_D)
    R(s, 0, 2.53, 13.33, 0.06, fill=TEAL_A)
    T(s, CO, 0.38, 0.1, 5, 0.52, sz=SZ_S, bold=True, col=WHITE)
    T(s, "Formation Initiale Module — Fundraising Téléphonique UNICEF", 0.38, 0.62, 9, 0.36, sz=SZ_SM, col=RGBColor(0xBB,0xEE,0xEE), italic=True)
    if tag: T(s, tag, 0.38, 1.0, 9, 0.3, sz=SZ_XS+2, bold=True, col=RGBColor(0x88,0xDD,0xDD))
    T(s, title, 0.38, 1.35, 11.5, 1.05, sz=SZ_T, bold=True, col=WHITE)
    if sub: T(s, sub, 0.38, 2.12, 11.5, 0.42, sz=SZ_SM, col=RGBColor(0xCC,0xEE,0xEE), italic=True)
    LOGO_ADD(s)
    FOOTER(s)

def H_CONTENT(s, title):
    R(s, 0, 0, 13.33, 7.5, fill=TEAL_L)
    R(s, 0, 0, 13.33, 0.82, fill=TEAL_D)
    R(s, 0, 0.8, 13.33, 0.06, fill=TEAL_A)
    T(s, CO, 0.18, 0.1, 2.2, 0.44, sz=SZ_SM+2, bold=True, col=WHITE)
    R(s, 2.52, 0.1, 0.05, 0.6, fill=TEAL_A)
    T(s, title, 2.68, 0.08, 10.4, 0.65, sz=SZ_T, bold=True, col=WHITE)
    LOGO_ADD(s)
    FOOTER(s)

def MERCI(p, module_tag=""):
    s = p.slides.add_slide(SL(p))
    R(s, 0, 0, 13.33, 7.5, fill=TEAL_D)
    R(s, 0, 0, 13.33, 0.08, fill=TEAL_A)
    R(s, 0, 7.42, 13.33, 0.08, fill=TEAL_A)
    sh = s.shapes.add_shape(9, Inches(4.5), Inches(1.2), Inches(4.3), Inches(4.3))
    sh.fill.solid(); sh.fill.fore_color.rgb = TEAL_M
    sh.line.fill.background()
    T(s, "🙏", 5.5, 2.2, 2.3, 1.5, sz=48, align=PP_ALIGN.CENTER, col=WHITE)
    T(s, "Merci à Tous !", 1.5, 1.4, 10.3, 1.2, sz=46, bold=True, col=WHITE, align=PP_ALIGN.CENTER)
    T(s, "Votre engagement fait la différence pour les enfants du monde entier.", 1.5, 2.85, 10.3, 0.6, sz=SZ_B, col=RGBColor(0xCC,0xEE,0xEE), align=PP_ALIGN.CENTER, italic=True)
    if module_tag:
        T(s, module_tag, 1.5, 3.55, 10.3, 0.38, sz=SZ_SM, col=RGBColor(0x88,0xDD,0xDD), align=PP_ALIGN.CENTER)
    T(s, "« Le refus d'aujourd'hui peut être le don de demain »", 2, 4.1, 9.33, 0.5, sz=SZ_SM+2, italic=True, col=RGBColor(0xAA,0xDD,0xDD), align=PP_ALIGN.CENTER)
    LOGO_ADD(s, x=6.2, y=5.6, w=1.0, h=0.98)

def save(p, name):
    path = os.path.join(OUT, name)
    p.save(path)
    sz = os.path.getsize(path) // 1024
    print(f"  ✓ {name}  ({sz} KB)")
    return path

# ══════════════════════════════════════════════════════════════════════════
#  QUIZ MINI PAR MODULE
# ══════════════════════════════════════════════════════════════════════════

QUIZ_J1 = [
    ("En quelle année l'UNICEF a-t-il été fondé ?", ["a) 1901", "b) 1946 ✓", "c) 1959", "d) 1919"]),
    ("Dans combien de pays l'UNICEF est-il présent ?", ["a) 120", "b) 150", "c) 190 ✓", "d) 210"]),
    ("Quel % des enfants du monde l'UNICEF vaccine-t-il ?", ["a) 25%", "b) 35%", "c) 45% ✓", "d) 60%"]),
    ("Que signifie RUTF ?", ["a) Rapid Urban Treatment", "b) Ready-to-Use Therapeutic Food ✓", "c) Regional Unified Task", "d) Routine Universal"]),
    ("Quel est le poids standard d'un sachet RUTF ?", ["a) 45g", "b) 75g", "c) 92g ✓", "d) 120g"]),
]

QUIZ_J2 = [
    ("Date de la loi associative française ?", ["a) 14 juil. 1789", "b) 1er juil. 1901 ✓", "c) 5 mai 1946", "d) 10 mars 1972"]),
    ("Combien d'associations en France ?", ["a) 250 000", "b) 800 000", "c) 1,5 million ✓", "d) 3 millions"]),
    ("Quel événement de 1859 ?", ["a) Conférence Berlin", "b) Traité Genève", "c) Bataille de Solférino ✓", "d) Révolution industrielle"]),
    ("Type d'association le plus reconnu ?", ["a) De fait", "b) Déclarée", "c) Utilité publique ✓", "d) Internationale"]),
    ("Combien de bénévoles en France ?", ["a) 5M", "b) 12M", "c) 22M ✓", "d) 35M"]),
    ("Lequel N'est PAS un principe humanitaire ?", ["a) Humanité", "b) Impartialité", "c) Solidarité nationale ✓", "d) Neutralité"]),
]

QUIZ_J3 = [
    ("Combien d'étapes comporte le script UNICEF ?", ["a) 4", "b) 5", "c) 7 ✓", "d) 10"]),
    ("Combien de règles d'accroche officielles ?", ["a) 2", "b) 3", "c) 5 ✓", "d) 7"]),
    ("Combien de modes de validation du don ?", ["a) 2", "b) 3", "c) 4 ✓", "d) 6"]),
    ("Trame accroche : combien de points ?", ["a) 2", "b) 3", "c) 4", "d) 5 ✓"]),
    ("Durée standard d'un appel complet ?", ["a) 1–2 min", "b) 3–5 min ✓", "c) 6–8 min", "d) 9–10 min"]),
    ("Date pivot pour prélèvement du mois courant ?", ["a) 01", "b) 04 ✓", "c) 10", "d) 15"]),
    ("TX d'audiolyse UNICEF ?", ["a) 50%", "b) 60%", "c) 70%", "d) 80% ✓"]),
]

QUIZ_J4 = [
    ("Combien de catégories d'objections ?", ["a) 3", "b) 5 ✓", "c) 7", "d) 10"]),
    ("Acronyme AAR ?", ["a) Analyser / Adapter", "b) Accuser / Argumenter / Reformuler", "c) Accuser réception / Argumenter / Relancer ✓", "d) Approcher / Assurer"]),
    ("Combien d'objections au total ?", ["a) 8", "b) 10", "c) 15 ✓", "d) 20"]),
    ("Règle 60 secondes : objectif minimum ?", ["a) 1 PEL", "b) 1 CU ✓", "c) 1 PA", "d) 1 DON"]),
    ("Max d'objections APRÈS l'appel au don ?", ["a) 1", "b) 2 ✓", "c) 3", "d) 4"]),
    ("Verrouillage réussi : confirmation promises ?", ["a) 20%", "b) 30%", "c) 50% ✓", "d) 70%"]),
]

QUIZ_J5 = [
    ("Grille de certification : combien de critères ?", ["a) 12", "b) 18", "c) 24 ✓", "d) 30"]),
    ("Seuil de certification FIM ?", ["a) 12/24", "b) 15/24", "c) 17/24 ✓", "d) 20/24"]),
    ("Seuil de quiz final ?", ["a) 20/40", "b) 24/40", "c) 28/40 ✓", "d) 32/40"]),
    ("Score excellent au quiz ?", ["a) ≥ 30/40", "b) ≥ 33/40", "c) ≥ 35/40 ✓", "d) ≥ 38/40"]),
    ("Plan d'action : J+30 objectif principal ?", ["a) 5 CU/H", "b) 6 CU/H", "c) 9 CU/H ✓", "d) 12 CU/H"]),
]

QUIZ_FINAL_40 = [
    ("Fondation UNICEF ?", ["a) 1901", "b) 1946 ✓", "c) 1959", "d) 1989"]),
    ("Pays UNICEF ?", ["a) 150", "b) 190 ✓", "c) 210", "d) 220"]),
    ("Vaccination UNICEF %age ?", ["a) 35%", "b) 40%", "c) 45% ✓", "d) 50%"]),
    ("RUTF poids ?", ["a) 75g", "b) 92g ✓", "c) 100g", "d) 120g"]),
    ("RUTF guérison enfant ?", ["a) 4 semaines", "b) 6–8 semaines ✓", "c) 10 semaines", "d) 12 semaines"]),
    ("Associations France ?", ["a) 1M", "b) 1,2M", "c) 1,5M ✓", "d) 2M"]),
    ("Bénévoles France ?", ["a) 18M", "b) 20M", "c) 22M ✓", "d) 25M"]),
    ("Loi 1901 date ?", ["a) 1er juil. ✓", "b) 14 juil.", "c) 5 mai", "d) 10 mars"]),
    ("Solférino année ?", ["a) 1859 ✓", "b) 1863", "c) 1889", "d) 1946"]),
    ("Croix-Rouge fondation ?", ["a) 1859", "b) 1863 ✓", "c) 1900", "d) 1946"]),
    ("Étapes script ?", ["a) 5", "b) 6", "c) 7 ✓", "d) 8"]),
    ("Règles accroche ?", ["a) 3", "b) 4", "c) 5 ✓", "d) 6"]),
    ("Modes validation ?", ["a) 2", "b) 3", "c) 4 ✓", "d) 5"]),
    ("Première étape script ?", ["a) Accroche ✓", "b) Malnutrition", "c) RUTF", "d) Appel"]),
    ("CU/H objectif ?", ["a) 6", "b) 7", "c) 9 ✓", "d) 10"]),
    ("TX Transfo définition ?", ["a) Taux temps", "b) Taux transformation PEL/PA ✓", "c) Taux transfert", "d) Taux total"]),
    ("PDC signifie ?", ["a) Plan Détails", "b) Plan Débrief", "c) Plan de Charge ✓", "d) Plan Direct"]),
    ("Production effective /jour ?", ["a) 6h00", "b) 6h40 ✓", "c) 7h00", "d) 8h00"]),
    ("Catégories objections ?", ["a) 3", "b) 4", "c) 5 ✓", "d) 6"]),
    ("Objections total ?", ["a) 10", "b) 12", "c) 15 ✓", "d) 18"]),
    ("AAR signifie ?", ["a) Accuser / Adapter / Répondre", "b) Accuser réception / Arg / Relancer ✓", "c) Analyser / Adapter / Reformuler", "d) Approcher / Assurer / Recadrer"]),
    ("Silence avant relance AAR ?", ["a) 1 sec", "b) 2 sec", "c) 3 sec ✓", "d) 5 sec"]),
    ("Objections initiales (Cat.1) ?", ["a) 6", "b) 8 ✓", "c) 10", "d) 12"]),
    ("Règle 60 secondes : phase ?", ["a) Accroche", "b) Vol", "c) Avant accroche ✓", "d) Atterrissage"]),
    ("Max objections post-appel ?", ["a) 1", "b) 2 ✓", "c) 3", "d) 4"]),
    ("Accord principe : avant ?", ["a) Script", "b) Objections", "c) Atterrissage ✓", "d) Validation"]),
    ("Prélèvement : jour ?", ["a) 05", "b) 08", "c) 10 ✓", "d) 15"]),
    ("Avant 04 : prélèvement ?", ["a) Mois prochain", "b) Même mois ✓", "c) Dans 2 mois", "d) Jamais"]),
    ("Après 04 : prélèvement ?", ["a) Même mois", "b) Mois prochain ✓", "c) Dans 2 mois", "d) Jamais"]),
    ("TX audiolyse ?", ["a) 60%", "b) 70%", "c) 80% ✓", "d) 90%"]),
    ("Verrouillage : confirmation ?", ["a) 30%", "b) 40%", "c) 50% ✓", "d) 60%"]),
    ("Sans verrouillage : confirmation ?", ["a) 1/5", "b) 1/10 ✓", "c) 1/3", "d) 1/2"]),
    ("Closing optimal ?", ["a) C'est OK", "b) C'est formidable ✓", "c) C'est bien", "d) C'est super"]),
    ("Grille certification ?", ["a) 18", "b) 20", "c) 24 ✓", "d) 28"]),
    ("Certification minimum ?", ["a) 15/24", "b) 17/24 ✓", "c) 19/24", "d) 20/24"]),
    ("Quiz final seuil ?", ["a) 24/40", "b) 26/40", "c) 28/40 ✓", "d) 30/40"]),
    ("Excellent quiz ?", ["a) 33/40", "b) 35/40 ✓", "c) 37/40", "d) 40/40"]),
    ("Posture vocale ?", ["a) Rapide", "b) Neutre", "c) Chaleureuse ✓", "d) Autoritaire"]),
    ("Mantra FIM ?", ["a) Refus = refus", "b) Refus d'aujourd'hui = don demain ✓", "c) Insistez toujours", "d) Silence vaut mieux"]),
]

def build_j1_interactive():
    p = prs()
    s = p.slides.add_slide(SL(p))
    H_TITLE(s, "Fondations & Mission", "Votre identité de fundraiser · La cause · Votre raison d'agir",
            tag="BOOK J1  ·  Journée 1 — 5 jours de FORMATION")
    
    s = p.slides.add_slide(SL(p))
    H_CONTENT(s, "Programme Jour 1 — Magistral + Interactif")
    prog = [
        ("08h30–09h00", "Accueil & Immersion", "Témoignage bénéficiaire UNICEF · Notez ce qui vous touche"),
        ("09h00–10h00", "Apport magistral J1", "UNICEF & cause · 190 pays · RUTF 92g · Malnutrition · 3 principes humanitaires"),
        ("10h00–10h30", "Exercice interactif", "En groupe : 'Qu'est-ce qui vous a le plus touché ?'"),
        ("10h30–11h00", "Activité pratique", "Pitch mission 60 secondes à votre binôme — sans notes"),
        ("11h00–11h30", "KPI FIDELIS intro", "Nomenclature DON/INDÉCIS/REFUS · CU/H minimum · Organisation journée"),
        ("11h30–12h30", "Quiz J1 & Débrief", "5 questions · Corrigé collectif · 3 apprentissages clés"),
    ]
    y = 0.9
    for time, titre, desc in prog:
        R(s, 0.3, y, 1.8, 0.6, fill=TEAL_D)
        T(s, time, 0.35, y+0.08, 1.7, 0.48, sz=SZ_XS+2, bold=True, col=WHITE, align=PP_ALIGN.CENTER)
        R(s, 2.12, y, 2.8, 0.6, fill=TEAL)
        T(s, titre, 2.2, y+0.1, 2.64, 0.48, sz=SZ_SM, bold=True, col=WHITE)
        R(s, 4.94, y, 8.09, 0.6, fill=WHITE, line=TEAL, lw=Pt(0.8))
        T(s, desc, 5.06, y+0.08, 7.88, 0.48, sz=SZ_B, col=DARK)
        y += 0.63
    
    s = p.slides.add_slide(SL(p))
    H_CONTENT(s, "Contenu Magistral J1 — L'UNICEF & La Cause")
    facts = [
        ("1946", "Fondation UNICEF (ONU)", TEAL_D),
        ("190", "Pays présents dans le monde", TEAL),
        ("45%", "Enfants du monde vaccinés", TEAL_M),
        ("92g", "Poids sachet RUTF malnutrition", TEAL_A),
        ("6–8 sem.", "Guérison enfant après RUTF", RGBColor(0x33,0x77,0x77)),
        ("3 princ.", "Humanité · Impartialité · Neutralité", RGBColor(0x22,0x66,0x66)),
    ]
    x = 0.3
    for val, label, col in facts:
        R(s, x, 0.9, 2.1, 1.1, fill=col)
        T(s, val, x+0.1, 0.92, 1.9, 0.65, sz=22, bold=True, col=WHITE, align=PP_ALIGN.CENTER)
        T(s, label, x+0.1, 1.55, 1.9, 0.42, sz=SZ_SM, col=RGBColor(0xCC,0xEE,0xEE), align=PP_ALIGN.CENTER, wrap=True)
        x += 2.16
    
    s = p.slides.add_slide(SL(p))
    H_CONTENT(s, "Quiz J1 — Évaluation des Connaissances (5 Questions)")
    y = 0.9
    for i, (q, opts) in enumerate(QUIZ_J1[:3]):
        R(s, 0.3, y, 0.5, 1.0, fill=TEAL_D)
        T(s, f"Q{i+1}", 0.32, y+0.3, 0.46, 0.4, sz=14, bold=True, col=WHITE, align=PP_ALIGN.CENTER)
        R(s, 0.82, y, 12.21, 0.45, fill=TEAL_D)
        T(s, q, 0.94, y+0.05, 11.98, 0.38, sz=SZ_SM, bold=True, col=WHITE, wrap=True)
        oy = y + 0.47
        for opt in opts:
            R(s, 0.82, oy, 12.21, 0.35, fill=WHITE if opts.index(opt) % 2 == 0 else TEAL_L, line=TEAL_D, lw=Pt(0.4))
            T(s, opt, 0.94, oy+0.05, 11.98, 0.26, sz=SZ_B, col=DARK)
            oy += 0.36
        y += 2.3
    
    MERCI(p, "Quiz J1 — Fondations & Mission")
    return save(p, f"FIM_BookJ1_Interactif_{VER}.pptx")

def build_j2_interactive():
    p = prs()
    s = p.slides.add_slide(SL(p))
    H_TITLE(s, "Le Monde Associatif", "Comprendre le secteur pour représenter la cause avec légitimité",
            tag="BOOK J2  ·  Journée 2 — 5 jours de FORMATION")
    
    s = p.slides.add_slide(SL(p))
    H_CONTENT(s, "Programme Jour 2 — Cas d'étude + Quiz")
    s = p.slides.add_slide(SL(p))
    H_CONTENT(s, "Contenu Magistral J2 — Secteur Associatif Français")
    
    s = p.slides.add_slide(SL(p))
    H_CONTENT(s, "Quiz J2 — 6 Questions")
    y = 0.9
    for i, (q, opts) in enumerate(QUIZ_J2[:3]):
        R(s, 0.3, y, 0.5, 1.0, fill=TEAL)
        T(s, f"Q{i+1}", 0.32, y+0.3, 0.46, 0.4, sz=14, bold=True, col=WHITE, align=PP_ALIGN.CENTER)
        R(s, 0.82, y, 12.21, 0.45, fill=TEAL)
        T(s, q, 0.94, y+0.05, 11.98, 0.38, sz=SZ_SM, bold=True, col=WHITE, wrap=True)
        oy = y + 0.47
        for opt in opts:
            R(s, 0.82, oy, 12.21, 0.35, fill=WHITE if opts.index(opt) % 2 == 0 else TEAL_L, line=TEAL, lw=Pt(0.4))
            T(s, opt, 0.94, oy+0.05, 11.98, 0.26, sz=SZ_B, col=DARK)
            oy += 0.36
        y += 2.3
    
    MERCI(p, "Quiz J2 — Monde Associatif")
    return save(p, f"FIM_BookJ2_Interactif_{VER}.pptx")

def build_j3_interactive():
    p = prs()
    s = p.slides.add_slide(SL(p))
    H_TITLE(s, "Script & Accroche", "Les 7 étapes · Accueil correct · Trame 5 points · 3 phases appel",
            tag="BOOK J3  ·  Journée 3 — 5 jours de FORMATION")
    
    s = p.slides.add_slide(SL(p))
    H_CONTENT(s, "ALLO — Accueil Correct (Fondamental)")
    steps = [
        (TEAL_D, "ÉTAPE 1 : ALLO", "Fundraiser dit 'ALLO' puis s'arrête. Écoute active pour identifier Madame ou Monsieur."),
        (TEAL, "ÉTAPE 2 : IDENTIFICATION", "Prénom EN PREMIER, puis NOM. 'Est-ce bien M. François DUPONT ?' (convention française)"),
        (TEAL_M, "ÉTAPE 3 : PRÉSENTATION", "'Bonjour M. Dupont, je m'appelle [Prénom], je vous appelle au nom de l'UNICEF France.'"),
    ]
    x = 0.3
    for col, title, body in steps:
        R(s, x, 0.9, 4.17, 0.5, fill=col)
        T(s, title, x+0.12, 0.92, 3.93, 0.46, sz=SZ_SM, bold=True, col=WHITE)
        R(s, x, 1.4, 4.17, 4.7, fill=WHITE, line=col, lw=Pt(1.5))
        T(s, body, x+0.14, 1.5, 3.89, 4.48, sz=SZ_B, col=DARK, wrap=True)
        x += 4.3
    
    s = p.slides.add_slide(SL(p))
    H_CONTENT(s, "Les 3 Phases de l'Appel Sortant")
    
    s = p.slides.add_slide(SL(p))
    H_CONTENT(s, "Trame Accroche — 5 Points Clés")
    
    s = p.slides.add_slide(SL(p))
    H_CONTENT(s, "Quiz J3 — 7 Questions")
    y = 0.88
    for i, (q, opts) in enumerate(QUIZ_J3[:4]):
        R(s, 0.3, y, 0.48, 0.9, fill=TEAL_M)
        T(s, f"Q{i+1}", 0.33, y+0.25, 0.42, 0.4, sz=12, bold=True, col=WHITE, align=PP_ALIGN.CENTER)
        R(s, 0.8, y, 11.93, 0.42, fill=TEAL_M)
        T(s, q, 0.9, y+0.04, 11.73, 0.36, sz=SZ_SM, bold=True, col=WHITE, wrap=True)
        oy = y + 0.44
        for opt in opts:
            R(s, 0.8, oy, 11.93, 0.33, fill=WHITE if opts.index(opt) % 2 == 0 else TEAL_L, line=TEAL_M, lw=Pt(0.4))
            T(s, opt, 0.9, oy+0.04, 11.73, 0.25, sz=SZ_XS+2, col=DARK)
            oy += 0.35
        y += 2.0
    
    MERCI(p, "Quiz J3 — Script & Accroche")
    return save(p, f"FIM_BookJ3_Interactif_{VER}.pptx")

def build_j4_interactive():
    p = prs()
    s = p.slides.add_slide(SL(p))
    H_TITLE(s, "Objections & Verrouillage", "15 objections · AAR · Règle 60s · Verrouillage 50%",
            tag="BOOK J4  ·  Journée 4 — 5 jours de FORMATION")
    
    s = p.slides.add_slide(SL(p))
    H_CONTENT(s, "Règles Fondamentales — Gestion Objections")
    
    s = p.slides.add_slide(SL(p))
    H_CONTENT(s, "Verrouillage & Closing — Sécuriser la Promesse")
    
    s = p.slides.add_slide(SL(p))
    H_CONTENT(s, "Quiz J4 — 6 Questions")
    y = 0.88
    for i, (q, opts) in enumerate(QUIZ_J4):
        R(s, 0.3, y, 0.48, 0.9, fill=TEAL_A)
        T(s, f"Q{i+1}", 0.33, y+0.25, 0.42, 0.4, sz=12, bold=True, col=WHITE, align=PP_ALIGN.CENTER)
        R(s, 0.8, y, 11.93, 0.42, fill=TEAL_A)
        T(s, q, 0.9, y+0.04, 11.73, 0.36, sz=SZ_SM, bold=True, col=WHITE, wrap=True)
        oy = y + 0.44
        for opt in opts:
            R(s, 0.8, oy, 11.93, 0.33, fill=WHITE if opts.index(opt) % 2 == 0 else TEAL_L, line=TEAL_A, lw=Pt(0.4))
            T(s, opt, 0.9, oy+0.04, 11.73, 0.25, sz=SZ_XS+2, col=DARK)
            oy += 0.35
        y += 2.0
    
    MERCI(p, "Quiz J4 — Objections & Verrouillage")
    return save(p, f"FIM_BookJ4_Interactif_{VER}.pptx")

def build_j5_interactive():
    p = prs()
    s = p.slides.add_slide(SL(p))
    H_TITLE(s, "Synthèse & Certification", "Grille 24 critères · Quiz final · Plan d'action terrain",
            tag="BOOK J5  ·  Journée 5 — 5 jours de FORMATION")
    
    s = p.slides.add_slide(SL(p))
    H_CONTENT(s, "Quiz J5 — 5 Questions d'Évaluation")
    y = 0.88
    for i, (q, opts) in enumerate(QUIZ_J5):
        R(s, 0.3, y, 0.48, 0.9, fill=RGBColor(0x33,0x77,0x77))
        T(s, f"Q{i+1}", 0.33, y+0.25, 0.42, 0.4, sz=12, bold=True, col=WHITE, align=PP_ALIGN.CENTER)
        R(s, 0.8, y, 11.93, 0.42, fill=RGBColor(0x33,0x77,0x77))
        T(s, q, 0.9, y+0.04, 11.73, 0.36, sz=SZ_SM, bold=True, col=WHITE, wrap=True)
        oy = y + 0.44
        for opt in opts:
            R(s, 0.8, oy, 11.93, 0.33, fill=WHITE if opts.index(opt) % 2 == 0 else TEAL_L, line=RGBColor(0x33,0x77,0x77), lw=Pt(0.4))
            T(s, opt, 0.9, oy+0.04, 11.73, 0.25, sz=SZ_XS+2, col=DARK)
            oy += 0.35
        y += 2.0
    
    MERCI(p, "Quiz J5 — Synthèse & Certification")
    return save(p, f"FIM_BookJ5_Interactif_{VER}.pptx")

def build_quiz_final_40():
    p = prs()
    s = p.slides.add_slide(SL(p))
    H_TITLE(s, "Quiz Final — 40 Questions", "Évaluation consolidée de formation FIM", tag="QUIZ FINAL")
    T(s, "Seuil de validation : 28/40 (70%)  ·  Excellent : ≥ 35/40",
      0.5, 2.75, 12.33, 0.5, sz=SZ_SM, bold=True, col=TEAL_D)
    
    q_list = [(i+1, q[0], q[1]) for i, q in enumerate(QUIZ_FINAL_40)]
    for slide_i in range(0, len(q_list), 4):
        batch = q_list[slide_i:slide_i+4]
        s = p.slides.add_slide(SL(p))
        R(s, 0, 0, 13.33, 7.5, fill=TEAL_L)
        R(s, 0, 0, 13.33, 0.82, fill=TEAL_D)
        R(s, 0, 0.8, 13.33, 0.06, fill=TEAL_A)
        T(s, CO, 0.18, 0.1, 2.2, 0.44, sz=SZ_SM+2, bold=True, col=WHITE)
        R(s, 2.52, 0.1, 0.05, 0.6, fill=TEAL_A)
        T(s, f"Quiz Final FIM — Questions {batch[0][0]}–{batch[-1][0]}", 2.68, 0.08, 10.4, 0.65, sz=SZ_T, bold=True, col=WHITE)
        LOGO_ADD(s)
        FOOTER(s)
        
        y = 0.9
        for qnum, qtext, opts in batch:
            R(s, 0.3, y, 0.5, 0.9, fill=TEAL_D)
            T(s, f"Q{qnum}", 0.33, y+0.3, 0.44, 0.35, sz=SZ_XS+2, bold=True, col=WHITE, align=PP_ALIGN.CENTER)
            R(s, 0.82, y, 12.21, 0.4, fill=TEAL_D)
            T(s, qtext, 0.92, y+0.04, 12.0, 0.34, sz=SZ_B, bold=True, col=WHITE, wrap=True)
            oy = y + 0.42
            for opt in opts:
                R(s, 0.82, oy, 12.21, 0.32, fill=WHITE if opts.index(opt)%2==0 else TEAL_L, line=TEAL_D, lw=Pt(0.3))
                T(s, opt, 0.92, oy+0.04, 12.0, 0.24, sz=SZ_XS+1, col=DARK)
                oy += 0.34
            y += 1.72
    
    MERCI(p, "Quiz Final — 40 Questions")
    return save(p, f"FIM_BookQUIZ_Final_40Q_{VER}.pptx")

if __name__ == "__main__":
    print(f"\n🎨 FIM {VER} — INTERACTIF + QUIZ PAR MODULE\n")
    files = []
    files.append(build_j1_interactive())
    files.append(build_j2_interactive())
    files.append(build_j3_interactive())
    files.append(build_j4_interactive())
    files.append(build_j5_interactive())
    files.append(build_quiz_final_40())
    print(f"\n✅ {len(files)} modules générés\n")

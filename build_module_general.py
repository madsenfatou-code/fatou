"""
FIM MODULE GÉNÉRAL — V8-01062026SHY-TE
Un seul module complet consolidant J1→J7 + Quiz Final
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
ACCENT = RGBColor(0x33, 0x77, 0x77)

SZ_T  = 26
SZ_S  = 22
SZ_B  = 18
SZ_SM = 14
SZ_XS = 10

LOGO = "/home/user/fatou/shy_logo_v2.png"
OUT  = "/home/user/fatou"
VER  = "V8-01062026SHY-TE"
CO   = "SHY-Performance"

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

def T(s, text, x, y, w, h, sz=SZ_B, bold=False, col=None,
      align=PP_ALIGN.LEFT, italic=False, wrap=True):
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
    try: s.shapes.add_picture(LOGO, Inches(x), Inches(y), Inches(w), Inches(h))
    except: pass

def FOOTER(s, page=""):
    R(s, 0, 7.1, 13.33, 0.4, fill=TEAL_D)
    txt = f"{CO}  ·  Module Général FIM  ·  UNICEF France  ·  {VER}"
    if page: txt += f"  ·  {page}"
    T(s, txt, 0, 7.13, 13.33, 0.3, sz=SZ_XS, col=WHITE, align=PP_ALIGN.CENTER)

def CHAPTER_DIVIDER(p, num, title, subtitle, color=None):
    """Slide de séparation de chapitre"""
    color = color or TEAL_D
    s = p.slides.add_slide(SL(p))
    R(s, 0, 0, 13.33, 7.5, fill=color)
    R(s, 0, 0, 13.33, 0.08, fill=TEAL_A)
    R(s, 0, 7.42, 13.33, 0.08, fill=TEAL_A)
    R(s, 0.3, 2.0, 0.12, 3.5, fill=TEAL_A)
    T(s, num, 0.6, 1.5, 12.0, 1.2, sz=52, bold=True, col=WHITE)
    T(s, title, 0.6, 2.7, 12.0, 1.0, sz=SZ_T, bold=True, col=WHITE)
    T(s, subtitle, 0.6, 3.75, 12.0, 0.6, sz=SZ_S, col=RGBColor(0xCC,0xEE,0xEE), italic=True)
    LOGO_ADD(s, x=12.1, y=6.4, w=1.0, h=0.88)
    return s

def H_CONTENT(s, title, sub=""):
    R(s, 0, 0, 13.33, 7.5, fill=TEAL_L)
    R(s, 0, 0, 13.33, 0.88, fill=TEAL_D)
    R(s, 0, 0.86, 13.33, 0.06, fill=TEAL_A)
    T(s, CO, 0.18, 0.08, 2.3, 0.46, sz=SZ_SM+2, bold=True, col=WHITE)
    R(s, 2.6, 0.1, 0.05, 0.66, fill=TEAL_A)
    T(s, title, 2.76, 0.06, 10.3, 0.7, sz=SZ_T, bold=True, col=WHITE)
    if sub:
        R(s, 0, 0.88, 13.33, 0.3, fill=TEAL_M)
        T(s, sub, 0.3, 0.9, 12.73, 0.26, sz=SZ_SM, bold=True, col=RGBColor(0xCC,0xEE,0xEE))
    LOGO_ADD(s)
    FOOTER(s)
    return s

def HL(s, text, x, y, w, h, fill=None, tc=None, sz=None, bold=False):
    fill = fill or TEAL; tc = tc or WHITE; sz = sz or SZ_SM
    R(s, x, y, w, h, fill=fill)
    T(s, text, x+0.15, y+0.08, w-0.28, h-0.14, sz=sz, col=tc, wrap=True, bold=bold)

def ITEM_LIST(s, items, x=0.3, y=0.9, w=12.73, col=None):
    """Liste d'items avec puces colorées"""
    col = col or TEAL_D
    for i, item in enumerate(items):
        bg = WHITE if i % 2 == 0 else TEAL_L
        R(s, x, y, 0.36, 0.5, fill=col)
        T(s, "▶", x+0.04, y+0.08, 0.28, 0.34, sz=SZ_SM, col=WHITE, align=PP_ALIGN.CENTER)
        R(s, x+0.38, y, w-0.38, 0.5, fill=bg, line=col, lw=Pt(0.4))
        T(s, item, x+0.52, y+0.07, w-0.56, 0.38, sz=SZ_B, col=DARK)
        y += 0.52

def TWO_COL(s, lt, li, rt, ri, y=0.9):
    cw = 6.18
    for head, items, ox in [(lt, li, 0.3), (rt, ri, 6.83)]:
        R(s, ox, y, cw, 0.44, fill=TEAL)
        T(s, head, ox+0.12, y+0.06, cw-0.2, 0.34, sz=SZ_SM, bold=True, col=WHITE)
        iy = y + 0.46
        for i, item in enumerate(items):
            bg = WHITE if i % 2 == 0 else TEAL_L
            R(s, ox, iy, cw, 0.54, fill=bg, line=TEAL, lw=Pt(0.4))
            T(s, "• " + item, ox+0.14, iy+0.08, cw-0.24, 0.4, sz=SZ_B, col=DARK)
            iy += 0.56

def QUIZ_SLIDE(p, num_module, questions, color=None):
    """Slide de quiz interactif pour chaque module"""
    color = color or TEAL_D
    s = p.slides.add_slide(SL(p))
    R(s, 0, 0, 13.33, 7.5, fill=TEAL_L)
    R(s, 0, 0, 13.33, 0.88, fill=color)
    R(s, 0, 0.86, 13.33, 0.06, fill=TEAL_A)
    T(s, CO, 0.18, 0.08, 2.3, 0.46, sz=SZ_SM+2, bold=True, col=WHITE)
    R(s, 2.6, 0.1, 0.05, 0.66, fill=TEAL_A)
    T(s, f"✏  Quiz — {num_module}", 2.76, 0.06, 10.3, 0.7, sz=SZ_T, bold=True, col=WHITE)
    LOGO_ADD(s)
    FOOTER(s)
    y = 0.98
    for i, (q, opts) in enumerate(questions):
        R(s, 0.3, y, 0.52, 0.38, fill=color)
        T(s, f"Q{i+1}", 0.33, y+0.04, 0.46, 0.3, sz=SZ_XS+2, bold=True, col=WHITE, align=PP_ALIGN.CENTER)
        R(s, 0.84, y, 12.19, 0.38, fill=color)
        T(s, q, 0.96, y+0.04, 12.0, 0.3, sz=SZ_SM, bold=True, col=WHITE, wrap=True)
        oy = y + 0.4
        for j, opt in enumerate(opts):
            bg = WHITE if j % 2 == 0 else TEAL_L
            R(s, 0.84, oy, 12.19, 0.3, fill=bg, line=color, lw=Pt(0.3))
            T(s, opt, 0.96, oy+0.04, 12.0, 0.22, sz=SZ_XS+2, col=DARK)
            oy += 0.31
        y += 1.66

def save(p, name):
    path = os.path.join(OUT, name)
    p.save(path)
    sz = os.path.getsize(path) // 1024
    print(f"  ✓ {name}  ({sz} KB)")
    return path

# ═══════════════════════════════════════════════════════════════════════════════
#  DONNÉES QUIZ MODULES
# ═══════════════════════════════════════════════════════════════════════════════

QJ1 = [
    ("En quelle année l'UNICEF a-t-il été fondé ?",          ["a) 1901","b) 1946 ✓","c) 1959","d) 1919"]),
    ("Dans combien de pays l'UNICEF est-il présent ?",        ["a) 150","b) 190 ✓","c) 210","d) 220"]),
    ("Poids d'un sachet RUTF ?",                              ["a) 75g","b) 92g ✓","c) 100g","d) 120g"]),
    ("Durée de guérison avec RUTF ?",                         ["a) 4 sem.","b) 6–8 sem. ✓","c) 10 sem.","d) 12 sem."]),
    ("Objectif CU/H minimum ?",                               ["a) 6","b) 7","c) 9 ✓","d) 10"]),
]
QJ2 = [
    ("Date loi associative française ?",                      ["a) 14 juil. 1789","b) 1er juil. 1901 ✓","c) 5 mai 1946","d) 10 mars 1972"]),
    ("Nombre d'associations en France ?",                     ["a) 800 000","b) 1,5 million ✓","c) 2 millions","d) 3 millions"]),
    ("Quel événement de 1859 ?",                              ["a) Traité Genève","b) Bataille de Solférino ✓","c) Conférence Berlin","d) Loi 1901"]),
    ("Lequel N'est PAS un principe humanitaire ?",            ["a) Humanité","b) Impartialité","c) Solidarité nationale ✓","d) Neutralité"]),
    ("Bénévoles associatifs en France ?",                     ["a) 12M","b) 18M","c) 22M ✓","d) 30M"]),
    ("Type d'association le plus reconnu par l'État ?",       ["a) De fait","b) Déclarée","c) Utilité publique ✓","d) Internationale"]),
]
QJ3 = [
    ("Combien d'étapes dans le script UNICEF ?",              ["a) 5","b) 6","c) 7 ✓","d) 8"]),
    ("Combien de règles officielles d'accroche ?",            ["a) 3","b) 4","c) 5 ✓","d) 6"]),
    ("1ère action du fundraiser en décrochant ?",             ["a) Bonjour","b) ALLO ✓","c) Se présenter","d) Identifier l'appel"]),
    ("Trame accroche : combien de points ?",                  ["a) 3","b) 4","c) 5 ✓","d) 6"]),
    ("Date pivot pour prélèvement mois courant ?",            ["a) 01","b) 04 ✓","c) 05","d) 10"]),
    ("TX d'audiolyse ?",                                      ["a) 60%","b) 70%","c) 80% ✓","d) 90%"]),
    ("Durée standard d'un appel complet ?",                   ["a) 1–2 min","b) 3–5 min ✓","c) 6–8 min","d) 10 min"]),
]
QJ4 = [
    ("Combien de catégories d'objections ?",                  ["a) 3","b) 4","c) 5 ✓","d) 6"]),
    ("Acronyme AAR ?",                                        ["a) Analyser/Adapter/Répondre","b) Accuser réception/Argumenter/Relancer ✓","c) Accepter/Assurer/Recadrer","d) Approcher/Argumenter/Reformuler"]),
    ("Règle 60 secondes : objectif minimum si objection précoce ?", ["a) 1 PEL","b) 1 CU ✓","c) 1 PA","d) 1 DON"]),
    ("Max d'objections APRÈS l'appel au don ?",               ["a) 1","b) 2 ✓","c) 3","d) 4"]),
    ("Verrouillage réussi : % promesses confirmées ?",        ["a) 20%","b) 30%","c) 50% ✓","d) 70%"]),
    ("Sans verrouillage : ratio promesses confirmées ?",      ["a) 1/5","b) 1/10 ✓","c) 1/3","d) 1/2"]),
]
QJ5 = [
    ("Grille de certification : combien de critères ?",       ["a) 18","b) 20","c) 24 ✓","d) 30"]),
    ("Seuil minimum pour certification FIM ?",                ["a) 15/24","b) 17/24 ✓","c) 19/24","d) 20/24"]),
    ("Seuil quiz final ?",                                    ["a) 24/40","b) 26/40","c) 28/40 ✓","d) 30/40"]),
    ("Score 'Excellent' au quiz final ?",                     ["a) ≥ 33/40","b) ≥ 35/40 ✓","c) ≥ 37/40","d) 40/40"]),
    ("Mantra de la formation FIM ?",                          ["a) Refus = refus, passez","b) Le refus d'aujourd'hui = don de demain ✓","c) Insistez toujours","d) Silence = accord"]),
]
Q40 = [
    ("Fondation UNICEF ?",["a) 1901","b) 1946 ✓","c) 1959","d) 1989"]),
    ("Pays UNICEF ?",["a) 150","b) 190 ✓","c) 210","d) 220"]),
    ("Vaccination UNICEF ?",["a) 35%","b) 40%","c) 45% ✓","d) 50%"]),
    ("RUTF poids ?",["a) 75g","b) 92g ✓","c) 100g","d) 120g"]),
    ("Guérison RUTF ?",["a) 4 sem.","b) 6–8 sem. ✓","c) 10 sem.","d) 12 sem."]),
    ("Associations France ?",["a) 1M","b) 1,2M","c) 1,5M ✓","d) 2M"]),
    ("Bénévoles France ?",["a) 18M","b) 20M","c) 22M ✓","d) 25M"]),
    ("Loi 1901 ?",["a) 1er juil. ✓","b) 14 juil.","c) 5 mai","d) 10 mars"]),
    ("Solférino ?",["a) 1859 ✓","b) 1863","c) 1889","d) 1946"]),
    ("Croix-Rouge fondation ?",["a) 1859","b) 1863 ✓","c) 1900","d) 1946"]),
    ("Étapes script ?",["a) 5","b) 6","c) 7 ✓","d) 8"]),
    ("Règles accroche ?",["a) 3","b) 4","c) 5 ✓","d) 6"]),
    ("Modes validation ?",["a) 2","b) 3","c) 4 ✓","d) 5"]),
    ("1ère action : décroché ?",["a) Bonjour","b) ALLO ✓","c) Présentation","d) Question"]),
    ("CU/H objectif ?",["a) 6","b) 7","c) 9 ✓","d) 10"]),
    ("TX Transfo ?",["a) Taux temps","b) Taux transformation PEL/PA ✓","c) Taux transfert","d) Total"]),
    ("PDC ?",["a) Plan Détails","b) Plan de Charge ✓","c) Plan Direct","d) Plan Débrief"]),
    ("Production /jour ?",["a) 6h","b) 6h40 ✓","c) 7h","d) 8h"]),
    ("Catégories objections ?",["a) 3","b) 4","c) 5 ✓","d) 6"]),
    ("Total objections ?",["a) 10","b) 12","c) 15 ✓","d) 18"]),
    ("AAR ?",["a) Analyser/Adapter","b) Accuser réception/Arg./Relancer ✓","c) Accepter/Assurer","d) Approcher/Reformuler"]),
    ("Silence AAR ?",["a) 1 sec","b) 2 sec","c) 3 sec ✓","d) 5 sec"]),
    ("Objections Cat.1 ?",["a) 6","b) 8 ✓","c) 10","d) 12"]),
    ("Règle 60s phase ?",["a) Accroche","b) Vol","c) Avant accroche ✓","d) Atterrissage"]),
    ("Max objections post-don ?",["a) 1","b) 2 ✓","c) 3","d) 4"]),
    ("Accord principe : avant ?",["a) Script","b) Objections","c) Atterrissage ✓","d) Validation"]),
    ("Prélèvement : jour ?",["a) 05","b) 08","c) 10 ✓","d) 15"]),
    ("Avant 04 : prélèvement ?",["a) Mois prochain","b) Même mois ✓","c) 2 mois","d) Jamais"]),
    ("Après 04 : prélèvement ?",["a) Même mois","b) Mois prochain ✓","c) 2 mois","d) Jamais"]),
    ("TX audiolyse ?",["a) 60%","b) 70%","c) 80% ✓","d) 90%"]),
    ("Verrouillage réussi ?",["a) 30%","b) 40%","c) 50% ✓","d) 60%"]),
    ("Sans verrouillage ?",["a) 1/5","b) 1/10 ✓","c) 1/3","d) 1/2"]),
    ("Closing optimal ?",["a) C'est OK","b) C'est formidable ✓","c) C'est bien","d) Très bien"]),
    ("Critères grille ?",["a) 18","b) 20","c) 24 ✓","d) 28"]),
    ("Certification min ?",["a) 15/24","b) 17/24 ✓","c) 19/24","d) 20/24"]),
    ("Quiz seuil ?",["a) 24/40","b) 26/40","c) 28/40 ✓","d) 30/40"]),
    ("Excellent quiz ?",["a) 33/40","b) 35/40 ✓","c) 37/40","d) 40/40"]),
    ("Posture vocale ?",["a) Rapide","b) Neutre","c) Chaleureuse ✓","d) Autoritaire"]),
    ("Mantra FIM ?",["a) Refus = refus","b) Refus aujourd'hui = don demain ✓","c) Insistez toujours","d) Silence vaut mieux"]),
]

# ═══════════════════════════════════════════════════════════════════════════════
#  CONSTRUCTION DU MODULE GÉNÉRAL
# ═══════════════════════════════════════════════════════════════════════════════

def build_module_general():
    p = prs()

    # ─────────────────────────────────────────────────────────
    #  SLIDE 1 — COUVERTURE PRINCIPALE
    # ─────────────────────────────────────────────────────────
    s = p.slides.add_slide(SL(p))
    R(s, 0, 0, 13.33, 7.5, fill=TEAL_D)
    R(s, 0, 0, 13.33, 0.08, fill=TEAL_A)
    R(s, 0, 7.42, 13.33, 0.08, fill=TEAL_A)
    R(s, 0, 2.5, 13.33, 0.06, fill=TEAL_A)
    sh = s.shapes.add_shape(9, Inches(9.8), Inches(0.5), Inches(3.2), Inches(3.2))
    sh.fill.solid(); sh.fill.fore_color.rgb = TEAL_M; sh.line.fill.background()
    T(s, CO, 0.5, 0.15, 9, 0.56, sz=SZ_S, bold=True, col=WHITE)
    T(s, "FORMATION INITIALE MODULE", 0.5, 0.72, 9, 0.52, sz=SZ_SM, col=RGBColor(0xBB,0xEE,0xEE), italic=True)
    T(s, "Module Général\nde Formation", 0.5, 1.28, 9.5, 1.15, sz=SZ_T+6, bold=True, col=WHITE)
    T(s, "Fundraising Téléphonique UNICEF France", 0.5, 2.52, 9, 0.44, sz=SZ_S, col=RGBColor(0xCC,0xEE,0xEE), italic=True)
    T(s, "Contenu consolidé J1 → J7  ·  7 jours de formation  ·  SHY-Performance × FIDELIS × UNICEF France",
       0.5, 3.1, 12.3, 0.44, sz=SZ_SM, col=RGBColor(0xAA,0xDD,0xDD))
    blocs = [("J1","Fondations & Mission"),("J2","Monde Associatif"),("J3","Script & Accroche"),
             ("J4","Objections & Verrouillage"),("J5","Certification"),("J6","Atelier 1"),("J7","Atelier 2 & Certif.")]
    x = 0.5
    for tag, desc in blocs:
        R(s, x, 3.72, 1.76, 0.42, fill=TEAL_M)
        T(s, tag, x+0.08, 3.75, 0.56, 0.34, sz=SZ_XS+2, bold=True, col=WHITE)
        T(s, desc, x+0.66, 3.75, 1.04, 0.34, sz=SZ_XS, col=RGBColor(0xCC,0xEE,0xEE))
        x += 1.81
    infos = [("40 Questions","Quiz final consolidé"),("7 modules","Quiz par module"),
             ("24 Critères","Grille de certification"),("9 CU/H","Objectif minimum")]
    x = 0.5
    for val, lab in infos:
        R(s, x, 4.3, 3.0, 0.9, fill=TEAL_A)
        T(s, val, x+0.14, 4.34, 2.72, 0.5, sz=SZ_T, bold=True, col=WHITE, align=PP_ALIGN.CENTER)
        T(s, lab, x+0.14, 4.84, 2.72, 0.32, sz=SZ_SM, col=RGBColor(0xCC,0xEE,0xEE), align=PP_ALIGN.CENTER)
        x += 3.14
    T(s, f"Version {VER}  ·  Formation 5 jours + 2 ateliers pratiques", 0.5, 5.38, 12.3, 0.36, sz=SZ_SM, col=RGBColor(0x88,0xDD,0xDD))
    T(s, "« Le refus d'aujourd'hui peut être le don de demain »", 0.5, 5.84, 12.3, 0.4, sz=SZ_SM, italic=True, col=RGBColor(0xAA,0xDD,0xDD))
    LOGO_ADD(s, x=12.1, y=6.4, w=1.0, h=0.88)

    # ─────────────────────────────────────────────────────────
    #  SLIDE 2 — SOMMAIRE GÉNÉRAL
    # ─────────────────────────────────────────────────────────
    s = p.slides.add_slide(SL(p))
    H_CONTENT(s, "Sommaire Général — Organisation des 7 Journées")
    chapitres = [
        ("CHAPITRE 1","J1","Fondations & Mission","UNICEF · 190 pays · RUTF 92g · KPI FIDELIS · Nomenclature", TEAL_D),
        ("CHAPITRE 2","J2","Monde Associatif","Loi 1901 · Secteur associatif · 22M bénévoles · 3 principes humanitaires", TEAL),
        ("CHAPITRE 3","J3","Script Officiel & Accroche","ALLO · 7 étapes · 5 règles · 3 phases appel · Trame 5 points · Règles PA", TEAL_M),
        ("CHAPITRE 4","J4","Objections & Verrouillage","15 objections · 5 catégories · Méthode AAR · Closing · Verrouillage 50%", TEAL_A),
        ("CHAPITRE 5","J5","Certification FIM","Grille 24 critères · Simulations · Plan d'action terrain J+30", ACCENT),
        ("CHAPITRE 6","J6","Atelier Pratique 1","Production simulée · Grille live · Débrief individuel · Cas complexes", RGBColor(0x22,0x66,0x66)),
        ("CHAPITRE 7","J7","Atelier Pratique 2","Perfectionnement · Certification finale · Analyse écarts · Plan poste", RGBColor(0x11,0x55,0x55)),
        ("QUIZ FINAL","","40 Questions","Évaluation consolidée J1→J7 · Seuil 28/40 · Mention Excellent ≥ 35/40", TEAL_D),
    ]
    y = 0.92
    for ch, jour, titre, desc, col in chapitres:
        R(s, 0.3, y, 1.5, 0.52, fill=col)
        T(s, ch, 0.36, y+0.07, 1.38, 0.36, sz=SZ_XS+1, bold=True, col=WHITE, align=PP_ALIGN.CENTER)
        R(s, 1.82, y, 0.7, 0.52, fill=TEAL_A)
        T(s, jour, 1.86, y+0.07, 0.62, 0.36, sz=SZ_XS+2, bold=True, col=WHITE, align=PP_ALIGN.CENTER)
        R(s, 2.54, y, 2.9, 0.52, fill=TEAL_L, line=col, lw=Pt(0.8))
        T(s, titre, 2.66, y+0.08, 2.74, 0.36, sz=SZ_SM, bold=True, col=col)
        R(s, 5.46, y, 7.57, 0.52, fill=WHITE, line=col, lw=Pt(0.4))
        T(s, desc, 5.58, y+0.08, 7.42, 0.38, sz=SZ_B, col=DARK)
        y += 0.555

    # ═══════════════════════════════════════════════════════
    #  CHAPITRE 1 — FONDATIONS & MISSION (J1)
    # ═══════════════════════════════════════════════════════
    CHAPTER_DIVIDER(p, "01", "Fondations & Mission", "L'UNICEF · La cause · Votre identité de fundraiser · KPI FIDELIS", TEAL_D)

    s = p.slides.add_slide(SL(p))
    H_CONTENT(s, "L'UNICEF — Chiffres Clés & Mission")
    facts = [
        ("1946", "Fondation UNICEF\n(ONU après WW2)", TEAL_D),
        ("190", "Pays présents\ndans le monde", TEAL),
        ("45%", "Enfants du monde\nvaccinés par UNICEF", TEAL_M),
        ("92g", "Poids sachet RUTF\n(malnutrition)", TEAL_A),
        ("6–8 sem.", "Guérison enfant\naprès traitement RUTF", ACCENT),
        ("1/11 sec", "1 enfant meurt\nde faim dans le monde", RGBColor(0x22,0x66,0x66)),
    ]
    x = 0.3
    for val, label, col in facts:
        R(s, x, 0.96, 2.1, 1.12, fill=col)
        T(s, val, x+0.1, 0.98, 1.9, 0.65, sz=22, bold=True, col=WHITE, align=PP_ALIGN.CENTER)
        T(s, label, x+0.1, 1.62, 1.9, 0.44, sz=SZ_SM, col=RGBColor(0xCC,0xEE,0xEE), align=PP_ALIGN.CENTER, wrap=True)
        x += 2.16
    HL(s, "La cause centrale : Malnutrition aigüe sévère sévère — Les sachets RUTF (Ready-to-Use Therapeutic Food) de 92g permettent à un enfant malnutri de guérir en 6 à 8 semaines.",
       0.3, 2.22, 12.73, 0.6, sz=SZ_B)
    HL(s, "3 principes humanitaires fondamentaux :\n1. HUMANITÉ — Agir pour préserver la vie et soulager la souffrance\n2. IMPARTIALITÉ — Sans discrimination de nationalité, race, religion\n3. NEUTRALITÉ — Ne prendre parti dans aucun conflit",
       0.3, 2.96, 12.73, 0.88, fill=TEAL_L, tc=DARK, sz=SZ_B)
    HL(s, "Votre mission : Représenter l'UNICEF avec conviction. Votre ton, votre posture, vos mots — tout compte.",
       0.3, 3.98, 12.73, 0.48, fill=TEAL_M, sz=SZ_SM)
    HL(s, "Posture vocale attendue : Sourire audible dès 'Bonjour' · Débit posé · Articulation claire · Ton chaleureux et sincère",
       0.3, 4.6, 12.73, 0.48, fill=TEAL_D, sz=SZ_SM)

    s = p.slides.add_slide(SL(p))
    H_CONTENT(s, "Nomenclature FIDELIS — KPI & Indicateurs")
    kpis = [
        ("CU/H","Contacts Utiles / Heure","Contacts qualifiés par heure de production","≥ 9 CU/H",TEAL_D),
        ("TX Transfo","Taux de Transformation","% PEL + PA obtenus parmi les Contacts Utiles","Cible FIDELIS",TEAL),
        ("PDC","Plan de Charge","Volume mensuel de dons réguliers FIDELIS × UNICEF","Défini mensuel",TEAL_M),
        ("DON","Accord de prélèvement","Contact qui accepte le PA — qualification positive","= CU validé",TEAL_A),
        ("INDÉCIS","À relancer","Contact hésitant — promesse non encore confirmée","Suivi Call 2",ACCENT),
        ("REFUS","Refus définitif","Contact qui refuse — clôture avec courtoisie","Qualification R",RGBColor(0x22,0x66,0x66)),
        ("PEL","Promesse En Ligne","Don validé en ligne après l'appel (différé)","Comptabilisé TX",RGBColor(0x11,0x55,0x55)),
        ("PA","Prélèvement Auto.","Don régulier SEPA — produit unique proposé","Produit unique",TEAL_D),
    ]
    x = 0.3; cw = 3.12; y_kpi = 0.96
    for i, (code, lib, defin, obj, col) in enumerate(kpis):
        bx = 0.3 + (i % 4) * 3.24
        by = y_kpi if i < 4 else y_kpi + 1.5
        R(s, bx, by, 3.12, 0.38, fill=col)
        T(s, f"{code}  —  {lib}", bx+0.1, by+0.06, 2.92, 0.28, sz=SZ_XS+2, bold=True, col=WHITE)
        R(s, bx, by+0.38, 3.12, 0.72, fill=WHITE if i%2==0 else TEAL_L, line=col, lw=Pt(0.6))
        T(s, defin, bx+0.1, by+0.44, 2.92, 0.34, sz=SZ_XS+2, col=DARK)
        T(s, obj, bx+0.1, by+0.82, 2.92, 0.24, sz=SZ_SM, bold=True, col=col)
    HL(s, "Conquête = nouveaux prospects  ·  Fidélisation / Réactivation = anciens donateurs  ·  6h40 production effective par journée",
       0.3, 3.26, 12.73, 0.42, fill=TEAL_D, sz=SZ_SM)
    HL(s, "Organisation journée : 9h30–17h30  ·  Briefing · Production · Débrief  ·  Fichiers FIDELIS segmentés par campagne",
       0.3, 3.82, 12.73, 0.42, fill=TEAL_M, sz=SZ_SM)

    QUIZ_SLIDE(p, "J1 — Fondations & Mission (5 Questions)", QJ1, TEAL_D)

    # ═══════════════════════════════════════════════════════
    #  CHAPITRE 2 — MONDE ASSOCIATIF (J2)
    # ═══════════════════════════════════════════════════════
    CHAPTER_DIVIDER(p, "02", "Le Monde Associatif & Humanitaire", "Loi 1901 · Histoire · Secteur · Légitimité de la cause", TEAL)

    s = p.slides.add_slide(SL(p))
    H_CONTENT(s, "La France Associative — Données Clés")
    stats = [
        ("1,5 Million","Associations en France", TEAL_D),
        ("22 Millions","Bénévoles associatifs", TEAL),
        ("1er juil. 1901","Loi fondatrice\nassociation déclarée", TEAL_M),
        ("3 types","De fait · Déclarée\nUtilité publique", TEAL_A),
        ("Don en Confiance","Label indépendant\nde contrôle et éthique", ACCENT),
        ("RGPD · Bloctel","Cadre légal des données\nde prospection téléphonique", RGBColor(0x22,0x66,0x66)),
    ]
    x = 0.3
    for val, label, col in stats:
        R(s, x, 0.96, 2.1, 1.1, fill=col)
        T(s, val, x+0.1, 0.98, 1.9, 0.64, sz=15, bold=True, col=WHITE, align=PP_ALIGN.CENTER)
        T(s, label, x+0.1, 1.6, 1.9, 0.44, sz=SZ_XS+2, col=RGBColor(0xCC,0xEE,0xEE), align=PP_ALIGN.CENTER, wrap=True)
        x += 2.16
    timeline = [
        ("1859", "Bataille de\nSolférino"), ("1863","Fondation\nCroix-Rouge"), ("1901","Loi 1901\nAssociations"),
        ("1946","Fondation\nUNICEF (ONU)"), ("1989","Convention\nDroits Enfant"), ("2000s","RGPD &\nBloctel"),
    ]
    x = 0.3
    for yr, ev in timeline:
        R(s, x, 2.2, 2.1, 0.38, fill=TEAL)
        T(s, yr, x+0.1, 2.23, 1.9, 0.3, sz=SZ_SM, bold=True, col=WHITE, align=PP_ALIGN.CENTER)
        R(s, x, 2.58, 2.1, 1.2, fill=TEAL_L, line=TEAL, lw=Pt(0.8))
        T(s, ev, x+0.12, 2.66, 1.86, 1.06, sz=SZ_B, col=DARK, wrap=True, align=PP_ALIGN.CENTER)
        x += 2.16
    HL(s, "3 principes humanitaires fondamentaux : HUMANITÉ  ·  IMPARTIALITÉ  ·  NEUTRALITÉ",
       0.3, 3.95, 12.73, 0.42, fill=TEAL_D, sz=SZ_SM, bold=True)
    HL(s, "Don régulier PA : valeur à vie × 8 vs don ponctuel. Colonne vertébrale des programmes UNICEF France.",
       0.3, 4.5, 12.73, 0.42, fill=TEAL_M, sz=SZ_SM)

    QUIZ_SLIDE(p, "J2 — Monde Associatif (6 Questions)", QJ2, TEAL)

    # ═══════════════════════════════════════════════════════
    #  CHAPITRE 3 — SCRIPT OFFICIEL & ACCROCHE (J3)
    # ═══════════════════════════════════════════════════════
    CHAPTER_DIVIDER(p, "03", "Script Officiel & Accroche", "ALLO · 7 étapes · 5 règles · 3 phases · Trame 5 points · Règles PA", TEAL_M)

    s = p.slides.add_slide(SL(p))
    H_CONTENT(s, "L'Accueil Téléphonique Correct — Séquence Fondamentale")
    seq = [
        (TEAL_D, "1  —  ALLO", "Le fundraiser décroche et dit UNIQUEMENT 'ALLO'\nPuis s'arrête et ÉCOUTE ACTIVEMENT\nObjectif : identifier le genre de l'interlocuteur\n→ Est-ce Madame ou Monsieur ?"),
        (TEAL, "2  —  IDENTIFICATION", "Convention française : PRÉNOM EN PREMIER, puis NOM\n\n✅ 'Est-ce bien M. François DUPONT ?'\n❌ 'Est-ce bien M. DUPONT François ?'\n\nJamais le nom de famille en premier — non naturel en français."),
        (TEAL_M, "3  —  PRÉSENTATION & SCRIPT", "Après confirmation :\n'Bonjour M. Dupont, je m'appelle [Prénom],\nje vous appelle au nom de l'UNICEF France.'\n\nPuis enchaîner IMMÉDIATEMENT avec les 5 règles d'accroche."),
    ]
    x = 0.3
    for col, title, body in seq:
        R(s, x, 0.96, 4.17, 0.52, fill=col)
        T(s, title, x+0.14, 0.99, 3.93, 0.44, sz=SZ_SM, bold=True, col=WHITE)
        R(s, x, 1.48, 4.17, 4.65, fill=WHITE, line=col, lw=Pt(1.5))
        T(s, body, x+0.14, 1.58, 3.89, 4.42, sz=SZ_B, col=DARK, wrap=True)
        x += 4.3
    HL(s, "⚠  Jamais 'NOM + Prénom' — Cela sonne froid, commercial et non francophone. Prénom + NOM = chaleur + identification claire.",
       0.3, 6.27, 12.73, 0.52, fill=RGBColor(0x00,0x55,0x55), sz=SZ_SM, bold=True)

    s = p.slides.add_slide(SL(p))
    H_CONTENT(s, "Les 3 Phases de l'Appel Sortant")
    phases = [
        (TEAL_D, "✈  DÉCOLLAGE", "Phase 1 — Accueil & Accroche",
         ["ALLO + écoute active + identification Mme/M.",
          "Prénom EN PREMIER → NOM de famille",
          "'Bonjour M. X, je m'appelle [Prénom], UNICEF France'",
          "Vérifier disponibilité : 'Vous avez 3 minutes ?'",
          "Lancer la trame d'accroche 5 points",
          "Si objection précoce → accepter + démarrer trame",
          "Objectif : DÉPASSER 60 SECONDES → 1 CU minimum"]),
        (TEAL, "🛫  VOL", "Phase 2 — Argumentation & Objections",
         ["Script 7 étapes dans l'ordre",
          "Trame : Présentation → Dramatisation → Solution",
          "Faire IMAGINER le donateur, le faire VOYAGER",
          "Traitement objections par méthode AAR",
          "ALLER JUSQU'À l'appel au don SANS s'arrêter",
          "MAXIMUM 2 objections après l'appel au don",
          "Obtenir un ACCORD DE PRINCIPE clair avant atterrissage"]),
        (TEAL_M, "🛬  ATTERRISSAGE", "Phase 3 — Accord & Validation PA",
         ["Accord de principe : OUI ferme + montant clair",
          "Verrouillage : question fermée de confirmation",
          "Closing : 'C'est formidable, merci !'",
          "Valider mode PA : IBAN / Ligne directe / PEL",
          "Vérifier adresse complète (n° de PORTE !)",
          "Confirmer n° téléphone mobile",
          "Historisation + qualification FIDELIS conforme"]),
    ]
    x = 0.3
    for col, title, sub, items in phases:
        R(s, x, 0.96, 4.17, 0.44, fill=col)
        T(s, title, x+0.12, 0.98, 3.93, 0.38, sz=SZ_S, bold=True, col=WHITE)
        R(s, x, 1.4, 4.17, 0.3, fill=RGBColor(0x00,0x66,0x66))
        T(s, sub, x+0.12, 1.43, 3.93, 0.24, sz=SZ_XS+2, bold=True, col=RGBColor(0xCC,0xEE,0xEE))
        R(s, x, 1.7, 4.17, 4.6, fill=WHITE, line=col, lw=Pt(1.2))
        iy = 1.78
        for item in items:
            T(s, "▶  " + item, x+0.12, iy, 3.9, 0.35, sz=SZ_SM, col=DARK)
            iy += 0.52
        x += 4.3
    HL(s, "Chaque phase a un objectif précis. Le vol ne peut pas s'arrêter avant l'appel au don.",
       0.3, 6.44, 12.73, 0.38, sz=SZ_SM)

    s = p.slides.add_slide(SL(p))
    H_CONTENT(s, "Trame d'Accroche — 5 Points Clés")
    T(s, "Si objection précoce : ACCEPTER sans argumenter, puis démarrer immédiatement cette trame.",
       0.35, 0.96, 12.5, 0.42, sz=SZ_SM, bold=True, col=TEAL_D)
    trame = [
        (TEAL_D, "1\nPRÉSEN-\nTATION", "QUI on parle",
         "Brève présentation de l'organisme.\n\n'Je vous appelle au nom de l'UNICEF France, l'organisation internationale qui protège les enfants dans 190 pays.'"),
        (TEAL, "2\nDRAMA-\nTISATION", "POURQUOI cet appel",
         "Gravité humaine. Prise de conscience sans excès.\n\n'En ce moment même, un enfant meurt de malnutrition toutes les 11 secondes...'"),
        (TEAL_M, "3\nSOLUTION", "COMMENT on aide",
         "Transformer la solution en confiance.\n\n'Un sachet RUTF de 92g suffit à traiter un enfant en 6 à 8 semaines. Votre geste peut tout changer.'"),
        (TEAL_A, "4\nCON-\nFIANCE", "FAIRE VOYAGER",
         "Faites IMAGINER le donateur. Faites-le voyager dans ses pensées.\n\n'Imaginez cet enfant qui retrouve l'énergie de jouer...'"),
        (ACCENT, "5\nSENSI-\nBILISATION", "APPEL À L'ACTION",
         "Émotions → réaction à chaud.\n\n'Avec seulement 10€/mois, vous permettez de traiter un enfant. Souhaitez-vous nous rejoindre ?'"),
    ]
    x = 0.3; cw = (13.33-0.6)/5
    for col, title, sub, body in trame:
        R(s, x, 1.46, cw, 0.56, fill=col)
        T(s, title, x+0.1, 1.48, cw-0.16, 0.52, sz=SZ_XS+2, bold=True, col=WHITE, align=PP_ALIGN.CENTER)
        R(s, x, 2.02, cw, 0.28, fill=RGBColor(0x00,0x62,0x62))
        T(s, sub, x+0.08, 2.05, cw-0.14, 0.22, sz=SZ_XS, bold=True, col=RGBColor(0xBB,0xEE,0xEE), align=PP_ALIGN.CENTER)
        R(s, x, 2.3, cw, 3.98, fill=WHITE, line=col, lw=Pt(1.2))
        T(s, body, x+0.1, 2.38, cw-0.16, 3.8, sz=SZ_SM, col=DARK, wrap=True)
        x += cw + 0.02
    HL(s, "Résumé : QUI → QUOI → COMMENT → CONFIANCE → SENSIBILISATION",
       0.3, 6.38, 12.73, 0.42, fill=TEAL_D, sz=SZ_SM, bold=True)

    s = p.slides.add_slide(SL(p))
    H_CONTENT(s, "Le Script Officiel UNICEF — Les 7 Étapes")
    steps = [
        ("①\nACCROCHE", TEAL_D, "Ouverture 5 règles\nTon chaleureux\nPrénom + NOM\nOrganisation nommée\nDisponibilité vérifiée"),
        ("②\nMALNUTRITION", TEAL, "La cause\n1 enfant / 11 sec\nDimension humaine\nPrise de conscience\nÉmotion sincère"),
        ("③\nSACHETS RUTF", TEAL_M, "La solution UNICEF\n92g par sachet\n6–8 semaines\npour guérir\nun enfant"),
        ("④\nAPPEL AU SOUTIEN", TEAL_A, "Proposition PA\nMontant suggéré\nImpact chiffré\nSimplicité du geste\n'10€ = 1 enfant'"),
        ("⑤\nSI DON PONCTUEL", ACCENT, "Réorientation → PA\n'Et si on envisageait\nquelque chose\nde régulier ?\nEngagement simple.'"),
        ("⑥\nINDÉCIS", RGBColor(0x22,0x66,0x66), "Gestion hésitation\nRester dans l'échange\nRelance douce\nQuestion ouverte\nSilence 3 sec"),
        ("⑦\nVALIDATION", RGBColor(0x11,0x55,0x55), "IBAN à chaud\nPA en ligne direct\nPEL / PA courrier\nCoordonnées complètes\nVerrouillage"),
    ]
    x = 0.3; sw = (13.33-0.6)/7
    for title, col, body in steps:
        R(s, x, 0.96, sw, 0.5, fill=col)
        T(s, title, x+0.06, 0.98, sw-0.1, 0.46, sz=SZ_XS+1, bold=True, col=WHITE, align=PP_ALIGN.CENTER)
        R(s, x, 1.46, sw, 4.92, fill=WHITE, line=col, lw=Pt(1.2))
        T(s, body, x+0.08, 1.54, sw-0.14, 4.72, sz=SZ_SM, col=DARK, wrap=True)
        x += sw + 0.02
    HL(s, "Durée standard : 3 à 5 min · Max 2 objections après l'appel au don · Accord de principe AVANT l'atterrissage.",
       0.3, 6.48, 12.73, 0.38, sz=SZ_SM)

    s = p.slides.add_slide(SL(p))
    H_CONTENT(s, "Les 5 Règles Officielles de la Phrase d'Accroche")
    T(s, "Ces 5 règles définissent un appel conforme FIDELIS — violer l'une d'elles = disqualification.",
       0.35, 0.96, 12.5, 0.38, sz=SZ_SM, bold=True, col=TEAL_D)
    rules = [
        ("RÈGLE 1", "Se présenter : prénom EN PREMIER, puis nom de famille",
         "'Bonjour, je m'appelle [Prénom NOM]...' — Convention française. Jamais NOM + Prénom."),
        ("RÈGLE 2", "Nommer l'organisation représentée",
         "'...je vous appelle au nom de l'UNICEF France...' — Le prospect sait immédiatement qui appelle."),
        ("RÈGLE 3", "Annoncer clairement la raison de l'appel",
         "'...pour vous parler d'une initiative importante pour les enfants...' — Transparence totale."),
        ("RÈGLE 4", "Vérifier la disponibilité du prospect",
         "'Est-ce que vous avez quelques minutes ?' — Respect et écoute active de l'interlocuteur."),
        ("RÈGLE 5", "Ton chaleureux, naturel et souriant — le sourire s'entend",
         "Débit posé · Articulation claire · Conviction sincère — Le ton est aussi important que les mots."),
    ]
    y = 1.44
    for idx, (tag, title, body) in enumerate(rules):
        bg = WHITE if idx % 2 == 0 else TEAL_L
        R(s, 0.3, y, 1.1, 0.88, fill=TEAL_D)
        T(s, tag, 0.32, y+0.18, 1.06, 0.52, sz=SZ_XS+1, bold=True, col=WHITE, align=PP_ALIGN.CENTER)
        R(s, 1.42, y, 3.8, 0.88, fill=TEAL_L, line=TEAL_D, lw=Pt(0.8))
        T(s, title, 1.54, y+0.2, 3.6, 0.52, sz=SZ_SM, bold=True, col=TEAL_D)
        R(s, 5.24, y, 7.79, 0.88, fill=bg)
        T(s, body, 5.36, y+0.12, 7.58, 0.68, sz=SZ_SM, col=DARK, italic=True)
        y += 0.93
    HL(s, "Un appel qui viole une de ces 5 règles est disqualifié dans la nomenclature FIDELIS.",
       0.3, 6.06, 12.73, 0.44, sz=SZ_SM)

    s = p.slides.add_slide(SL(p))
    H_CONTENT(s, "Règles de Prélèvement — Date Pivot du 04")
    T(s, "Le prélèvement PA est effectué le 10 de chaque mois. La date de saisie détermine le mois.",
       0.35, 0.96, 12.5, 0.44, sz=SZ_B, bold=True, col=TEAL_D)
    R(s, 0.3, 1.52, 6.1, 0.44, fill=TEAL_D)
    T(s, "PA SAISI AVANT LE 04 DU MOIS", 0.44, 1.56, 5.9, 0.34, sz=SZ_S, bold=True, col=WHITE)
    R(s, 0.3, 1.96, 6.1, 1.08, fill=WHITE, line=TEAL_D, lw=Pt(1.5))
    T(s, "→ Prélevé le 10 du MÊME MOIS\n\nExemple : PA signé le 04 juin → Prélèvement le 10 juin ✅",
       0.44, 2.04, 5.82, 0.94, sz=SZ_B, col=DARK)
    R(s, 6.73, 1.52, 6.1, 0.44, fill=TEAL)
    T(s, "PA SAISI APRÈS LE 04 DU MOIS", 6.87, 1.56, 5.9, 0.34, sz=SZ_S, bold=True, col=WHITE)
    R(s, 6.73, 1.96, 6.1, 1.08, fill=WHITE, line=TEAL, lw=Pt(1.5))
    T(s, "→ Prélevé le 10 du MOIS SUIVANT\n\nExemple : PA signé le 05 juin → Prélèvement le 10 juillet ✅",
       6.87, 2.04, 5.82, 0.94, sz=SZ_B, col=DARK)
    R(s, 0.3, 3.16, 12.53, 0.42, fill=TEAL_A)
    T(s, "RÈGLE SIMPLE : Date pivot = 04  ·  Avant ou le 04 = mois courant  ·  Après le 04 = mois suivant",
       0.44, 3.2, 12.2, 0.32, sz=SZ_SM, bold=True, col=WHITE)
    HL(s, "TX D'AUDIOLYSE : 80% — 8 donateurs sur 10 qui restent en ligne après la trame d'accroche finalisent leur don.",
       0.3, 3.72, 12.73, 0.48, fill=TEAL_M, sz=SZ_B)

    s = p.slides.add_slide(SL(p))
    H_CONTENT(s, "Les 4 Types de Prélèvement Automatique")
    types = [
        (TEAL_D, "PA EN LIGNE\nAGENT EN DIRECT",
         "L'agent collecte l'IBAN à chaud pendant l'appel.\n\n• Le donateur dicte ses coordonnées bancaires\n• L'agent saisit en temps réel dans le logiciel\n• 🟢 Voyant VERT = continuer la saisie\n• 🔴 Voyant ROUGE = suspendre immédiatement\n• Mode le plus rapide — valider IBAN mot à mot"),
        (TEAL, "PA EN LIGNE\nDONATEUR EN DIRECT",
         "L'agent envoie un lien de paiement sécurisé.\n\n• Le donateur saisit lui-même ses coordonnées\n• Envoi recommandé : MAIL (priorité), SMS en secours\n• C'est une PROMESSE — à confirmer J+5\n• Si erreur mail / accès → envoyer via Action Manager\n• Si pas d'adresse mail → uniceftmk@unicef.fr (à vérifier)"),
        (TEAL_M, "PA EN LIGNE\nDIFFÉRÉ CONFIRMÉ",
         "Le donateur confirme avoir payé en ligne.\n\n• Il lit le MESSAGE DE CONFIRMATION à voix haute\n• Ce message correspond à celui visible sur votre écran\n• → Qualifier en PAIEMENT DIFFÉRÉ CONFIRMÉ ✅\n• Si Call 2 rappelle et pas de réponse → Qualifier PEL\n• Promesse Définitive PEL = qualification finale"),
        (TEAL_A, "PROMESSE\nRELANCE CALL 2",
         "Promesse de paiement non encore confirmée.\n\n• Call 2 rappelle pour relance et confirmation\n• Call 1 peut finaliser OU Call 2 finalise\n• Si rappel sans réponse → PROMESSE DÉFINITIVE PEL\n• Si donateur rappelle : voir l'historique\n• CU autre agent → prendre en charge ou transférer"),
    ]
    x = 0.3; cw = (13.33-0.6)/4
    for col, title, body in types:
        R(s, x, 0.96, cw, 0.52, fill=col)
        T(s, title, x+0.1, 0.98, cw-0.16, 0.48, sz=SZ_XS+2, bold=True, col=WHITE, align=PP_ALIGN.CENTER)
        R(s, x, 1.48, cw, 4.95, fill=WHITE, line=col, lw=Pt(1.2))
        T(s, body, x+0.1, 1.56, cw-0.16, 4.72, sz=SZ_SM, col=DARK, wrap=True)
        x += cw + 0.02
    HL(s, "Objection 'D'où avez-vous mon numéro ?' → Répondre : fichiers partenaires (Amazon, M6 boutique, comparateurs, etc.)",
       0.3, 6.56, 12.73, 0.38, sz=SZ_SM)

    QUIZ_SLIDE(p, "J3 — Script Officiel & Accroche (7 Questions)", QJ3, TEAL_M)

    # ═══════════════════════════════════════════════════════
    #  CHAPITRE 4 — OBJECTIONS & VERROUILLAGE (J4)
    # ═══════════════════════════════════════════════════════
    CHAPTER_DIVIDER(p, "04", "Objections & Verrouillage", "15 objections · 5 catégories · Méthode AAR · Closing 50%", TEAL_A)

    s = p.slides.add_slide(SL(p))
    H_CONTENT(s, "Règles Fondamentales — Gestion des Objections dans l'Appel")
    R(s, 0.3, 0.96, 12.73, 0.5, fill=TEAL_D)
    T(s, "⏱  RÈGLE 60 SECONDES — Objection Précoce (avant l'accroche)",
       0.44, 1.0, 12.4, 0.38, sz=SZ_S, bold=True, col=WHITE)
    R(s, 0.3, 1.46, 12.73, 0.96, fill=WHITE, line=TEAL_D, lw=Pt(1.5))
    T(s, "Si le prospect objecte AVANT la trame d'accroche → Accepter l'objection sans argumenter\n"
         "→ Démarrer immédiatement la trame (Présentation → Dramatisation → Solution → Confiance → Sensibilisation)\n"
         "→ Objectif IMPÉRATIF : DÉPASSER 60 SECONDES → conclure au moins 1 CU",
       0.44, 1.54, 12.4, 0.82, sz=SZ_B, col=DARK)

    R(s, 0.3, 2.52, 12.73, 0.5, fill=TEAL)
    T(s, "🚫  RÈGLE 2 OBJECTIONS — Après l'Appel au Don",
       0.44, 2.56, 12.4, 0.38, sz=SZ_S, bold=True, col=WHITE)
    R(s, 0.3, 3.02, 12.73, 0.96, fill=WHITE, line=TEAL, lw=Pt(1.5))
    T(s, "Après avoir formulé l'appel au don → NE PAS DÉPASSER 2 objections traitées\n"
         "→ ALLER JUSQU'À l'appel au don sans s'arrêter à mi-chemin\n"
         "→ Obtenir un ACCORD DE PRINCIPE ferme AVANT l'atterrissage (Phase 3)",
       0.44, 3.1, 12.4, 0.82, sz=SZ_B, col=DARK)

    R(s, 0.3, 4.08, 12.73, 0.5, fill=TEAL_M)
    T(s, "✅  L'ACCORD DE PRINCIPE — Condition Obligatoire de l'Atterrissage",
       0.44, 4.12, 12.4, 0.38, sz=SZ_S, bold=True, col=WHITE)
    R(s, 0.3, 4.58, 12.73, 0.96, fill=WHITE, line=TEAL_M, lw=Pt(1.5))
    T(s, "Un OUI FERME avec montant clair AVANT de passer à la validation du PA.\n"
         "Exemples : 'Oui, je suis d'accord pour 15€/mois' · 'Oui, comment ça fonctionne ?'\n"
         "Sans accord de principe = pas d'atterrissage → Relancer ou conclure en INDÉCIS.",
       0.44, 4.66, 12.4, 0.82, sz=SZ_B, col=DARK)
    HL(s, "Ne jamais forcer un atterrissage sans accord de principe. Un PA sous pression → annulation rapide.",
       0.3, 5.66, 12.73, 0.46, fill=TEAL_D, sz=SZ_SM, bold=True)

    s = p.slides.add_slide(SL(p))
    H_CONTENT(s, "Méthode AAR — Accuser Réception · Argumenter · Relancer")
    aar = [
        ("A\nACCUSER\nRÉCEPTION", TEAL_D,
         "Montrez que vous avez entendu.\nNE PAS justifier immédiatement.\n\n'Je comprends tout à fait…'\n'C'est une question que beaucoup\nse posent et c'est tout à fait légitime…'\n\n⚠  Jamais 'Oui mais…'\n⚠  Jamais 'Non vous avez tort…'"),
        ("A\nARGUMENTER", TEAL,
         "1 argument + 1 preuve + 1 chiffre.\n\nRègles d'or :\n→ 1 seul argument par réponse\n→ Preuve vérifiable\n→ Chiffre simple et mémorisable\n\nEx : '10€/mois = 33 centimes/jour\n= 1 enfant traité en 6 semaines'"),
        ("R\nRELANCER", TEAL_M,
         "Revenez à la proposition.\nPas de pression.\n\n'Est-ce que cela répond\nà votre question ?'\n'Qu'est-ce qui vous permettrait\nde vous sentir à l'aise ?'\n\n⚠  3 secondes de silence\nAVANT de relancer."),
    ]
    x = 0.3
    for title, col, body in aar:
        R(s, x, 0.96, 4.17, 0.7, fill=col)
        T(s, title, x+0.14, 0.99, 3.93, 0.64, sz=SZ_S, bold=True, col=WHITE, align=PP_ALIGN.CENTER)
        R(s, x, 1.66, 4.17, 4.62, fill=WHITE, line=col, lw=Pt(1.8))
        T(s, body, x+0.16, 1.76, 3.85, 4.4, sz=SZ_B, col=DARK, wrap=True)
        x += 4.3
    HL(s, "Le silence après votre argument est votre allié — comptez 1–2–3 mentalement avant de relancer.",
       0.3, 6.42, 12.73, 0.42, sz=SZ_SM)

    s = p.slides.add_slide(SL(p))
    H_CONTENT(s, "Les 5 Catégories d'Objections — 15 Situations")
    cats = [
        ("CAT. 1\nOBJECTIONS\nINITIALES", TEAL_D,
         "Phase d'accroche — 8 objections :\n• Pas intéressé(e)\n• Faux numéro\n• Appel pour un don\n• Je donne déjà\n• Arnaque / méfiance\n• Pas par téléphone\n• Pas confiance associations\n• N'aime pas être contacté"),
        ("CAT. 2\nFINAN-\nCIÈRES", TEAL,
         "Après l'appel au don — 3 objections :\n\n• Je donne déjà ailleurs\n\n• Je n'ai pas les moyens\n\n• Dernier positionnement\nfinancier — refus final"),
        ("CAT. 3\nCONTRE\nLE PA", TEAL_M,
         "Sur le prélèvement — 3 objections :\n\n• Préfère le don ponctuel\n\n• N'aime pas l'engagement\nmensuel récurrent\n\n• Dernier positionnement PA"),
        ("CAT. 4\nMÉFIANCE\nIBAN", TEAL_A,
         "Sur la sécurité bancaire :\n\n• Protocole SEPA expliqué\n\n• Droits du donateur\n(arrêt en 1 clic)\n\n• Réassurance totale\ndonnées bancaires"),
        ("CAT. 5\nCONFLIC-\nTUELLES", ACCENT,
         "Objections sensibles :\n• Source du numéro\n→ Amazon, M6 boutique, etc.\n• RGPD / Droits légaux\n• Bloctel / liste rouge\n• Accent / identité\n• Ton agressif"),
    ]
    x = 0.3; cw = (13.33-0.6)/5
    for title, col, body in cats:
        R(s, x, 0.96, cw, 0.56, fill=col)
        T(s, title, x+0.1, 0.98, cw-0.16, 0.52, sz=SZ_XS+2, bold=True, col=WHITE, align=PP_ALIGN.CENTER)
        R(s, x, 1.52, cw, 4.85, fill=WHITE, line=col, lw=Pt(1.2))
        T(s, body, x+0.1, 1.6, cw-0.16, 4.62, sz=SZ_SM, col=DARK, wrap=True)
        x += cw + 0.02
    HL(s, "La majorité des appels se termine à Cat.1. Maîtriser ces 8 objections = atteindre vos 9 CU/H.",
       0.3, 6.48, 12.73, 0.42, sz=SZ_SM)

    s = p.slides.add_slide(SL(p))
    H_CONTENT(s, "Verrouillage, Closing & Repêchage — Sécuriser la Promesse")
    R(s, 0.3, 0.96, 6.1, 0.46, fill=TEAL_D)
    T(s, "🔒  VERROUILLAGE — Question Fermée", 0.44, 1.0, 5.9, 0.36, sz=SZ_S, bold=True, col=WHITE)
    R(s, 0.3, 1.42, 6.1, 2.12, fill=WHITE, line=TEAL_D, lw=Pt(1.5))
    T(s, "Posez UNE question fermée qui confirme la promesse :\n\n"
         "'Alors on est bien d'accord pour [montant]€/mois\nà partir du [mois] ?'\n\n"
         "⚠  Impact statistique :\n✅ Verrouillage réussi = 50% des promesses confirmées\n"
         "❌ Sans verrouillage = 1/10 seulement confirmées",
       0.44, 1.5, 5.82, 1.98, sz=SZ_SM, col=DARK)
    R(s, 6.73, 0.96, 6.1, 0.46, fill=TEAL)
    T(s, "🌟  CLOSING — Valoriser la Décision", 6.87, 1.0, 5.9, 0.36, sz=SZ_S, bold=True, col=WHITE)
    R(s, 6.73, 1.42, 6.1, 2.12, fill=WHITE, line=TEAL, lw=Pt(1.5))
    T(s, "Après l'accord de principe confirmé :\n\n"
         "Dites : 'C'est FORMIDABLE ! Merci beaucoup.'\n\n"
         "→ Rappeler l'impact : 'Votre geste va permettre\n   de traiter un enfant.'\n"
         "→ Sécuriser psychologiquement la promesse\n"
         "→ Enchaîner immédiatement sur la validation PA",
       6.87, 1.5, 5.82, 1.98, sz=SZ_SM, col=DARK)
    R(s, 0.3, 3.64, 12.73, 0.46, fill=TEAL_M)
    T(s, "🔁  REPÊCHAGE — Dernier Filet si Hésitation Finale", 0.44, 3.68, 12.4, 0.36, sz=SZ_S, bold=True, col=WHITE)
    R(s, 0.3, 4.1, 12.73, 0.88, fill=WHITE, line=TEAL_M, lw=Pt(1.2))
    T(s, "Il faut un OUI DE DON FERME avec montant clair — pas d'ambiguïté acceptée.\n"
         "Si hésitation → poser la question fermée de repêchage une dernière fois.\n"
         "Si refus définitif → 'Merci de m'avoir écouté. Bonne journée à vous.' Clore avec courtoisie.",
       0.44, 4.18, 12.4, 0.74, sz=SZ_B, col=DARK)
    HL(s, "Le verrouillage est psychologique — il SÉCURISE la promesse. Sans lui, 90% des promesses non confirmées disparaissent.",
       0.3, 5.1, 12.73, 0.48, fill=TEAL_A, sz=SZ_SM, bold=True)

    s = p.slides.add_slide(SL(p))
    H_CONTENT(s, "Informations à Collecter — Fiche Hermès & Données Donateur")
    TWO_COL(s,
        "Fiche Hermès — Données à revérifier en direct",
        ["Nom & prénom (confirmer M. ou Mme au décrochage)",
         "Adresse complète avec NUMÉRO DE PORTE\n(sans n° de porte → NPAI → courrier retourné)",
         "Code postal + ville — vérifier chaque caractère",
         "Email : lire lettre par lettre · Si erreur → Action Manager",
         "Si pas d'adresse mail : uniceftmk@unicef.fr (à vérifier)"],
        "Téléphone & Processus de Confirmation",
        ["Appel sur FIXE → confirmer + prendre le MOBILE",
         "Appel sur MOBILE → enrichir la base en gardant le n°",
         "Lien PA → envoyer par MAIL (priorité), SMS en secours",
         "Promesse PA en ligne → confirmer J+5 après validation",
         "Courrier postal → 3 à 4 jours max pour les donateurs\nqui choisissent cette voie"],
        y=0.98)
    HL(s, "Argumentaire adresse : revues trimestrielles UNICEF · Cadeaux aux adhérents · Reçu fiscal annuel",
       0.3, 5.78, 12.73, 0.46, fill=TEAL_M, sz=SZ_SM)
    HL(s, "Notification entrante (call blinding) : si donateur rappelle → voir historique → CU autre agent : reprendre en son nom OU transférer",
       0.3, 6.36, 12.73, 0.46, fill=TEAL_D, sz=SZ_SM)

    QUIZ_SLIDE(p, "J4 — Objections & Verrouillage (6 Questions)", QJ4, TEAL_A)

    # ═══════════════════════════════════════════════════════
    #  CHAPITRE 5 — CERTIFICATION FIM (J5)
    # ═══════════════════════════════════════════════════════
    CHAPTER_DIVIDER(p, "05", "Synthèse & Certification FIM", "Grille 24 critères · Plan d'action terrain J+30", ACCENT)

    s = p.slides.add_slide(SL(p))
    H_CONTENT(s, "Programme Journée 5 — Certification FIM")
    prg = [
        ("08h30–09h00","Révision consolidée","Révision 5 jours de formation · Quiz flash · Météo émotionnelle"),
        ("09h00–11h00","Simulations d'appels","Script + objections + verrouillage + closing chronométrés"),
        ("11h00–11h30","Débriefing collectif","3 apprentissages collectifs · Célébration des progrès"),
        ("11h30–12h30","Grille 24 critères","Évaluation individuelle · Entretien formateur 5 min · Feedback"),
        ("13h30–14h30","Quiz final 40 questions","Évaluation théorique · Seuil 28/40 · Corrigé collectif"),
        ("14h30–15h00","Certification","Attestations FIM officielles · Plan d'action terrain J+30"),
    ]
    y = 0.96
    for tm, ti, de in prg:
        R(s, 0.3, y, 2.2, 0.65, fill=TEAL_D)
        T(s, tm, 0.36, y+0.12, 2.08, 0.44, sz=SZ_XS+2, bold=True, col=WHITE, align=PP_ALIGN.CENTER)
        R(s, 2.52, y, 3.2, 0.65, fill=TEAL)
        T(s, ti, 2.62, y+0.14, 3.0, 0.44, sz=SZ_SM, bold=True, col=WHITE)
        R(s, 5.74, y, 7.29, 0.65, fill=WHITE, line=TEAL, lw=Pt(0.8))
        T(s, de, 5.86, y+0.1, 7.1, 0.5, sz=SZ_B, col=DARK)
        y += 0.68

    s = p.slides.add_slide(SL(p))
    H_CONTENT(s, "Grille de Certification — 24 Critères")
    T(s, "Voici exactement ce sur quoi vous serez évalué(e) — aucune surprise :",
       0.35, 0.92, 12.5, 0.36, sz=SZ_SM, bold=True, col=TEAL_D)
    criteria = [
        ("A1","ALLO + écoute active + identification genre","Accroche"),
        ("A2","Respect des 5 règles d'accroche officielles","Accroche"),
        ("A3","Ton chaleureux et naturel dès l'ouverture","Accroche"),
        ("A4","Présentation cause malnutrition (impact chiffré)","Script"),
        ("A5","Explication sachets RUTF avec précision","Script"),
        ("A6","Appel au soutien clair avec montant suggéré","Script"),
        ("A7","Gestion don ponctuel → réorientation PA","Script"),
        ("A8","Gestion indécis — relance constructive","Script"),
        ("A9","Validation coordonnées — mode paiement correct","Closing"),
        ("B1","Accusé réception objections Cat.1","Objections"),
        ("B2","Argument financier + preuve concrète","Objections"),
        ("B3","Réorientation PA après refus ponctuel","Objections"),
        ("B4","Réassurance IBAN / SEPA convaincante","Objections"),
        ("B5","Gestion calme des objections Cat.5","Objections"),
        ("C1","Ton chaleureux maintenu sur toute la durée","Posture"),
        ("C2","Débit adapté — ni trop rapide, ni trop lent","Posture"),
        ("C3","Absence d'hésitations et tics verbaux","Posture"),
        ("C4","Conviction sincère pour la mission UNICEF","Posture"),
        ("C5","Résilience après un refus direct","Posture"),
        ("C6","Respect des règles éthiques et RGPD","Éthique"),
        ("D1","Qualification conforme FIDELIS DON/IND/REF","FIDELIS"),
        ("D2","Nomination correcte du mode de validation","FIDELIS"),
        ("D3","Durée d'appel dans la norme (3–5 min)","FIDELIS"),
        ("D4","Saisie coordonnées complète et exacte","FIDELIS"),
    ]
    cxs = [0.3,0.78,6.82,9.72,10.52,11.32,12.12]
    cws = [0.45,6.0,2.86,0.78,0.78,0.78,0.75]
    cat_c = {"Accroche":TEAL_D,"Script":TEAL,"Closing":TEAL_M,"Objections":TEAL_A,
             "Posture":ACCENT,"Éthique":RGBColor(0x22,0x66,0x66),"FIDELIS":RGBColor(0x11,0x55,0x55)}
    # header
    for h,x,w in zip(["#","Ce que vous devez démontrer","Catégorie","1","2","3","4"],cxs,cws):
        R(s,x,1.32,w,0.34,fill=TEAL_D)
        T(s,h,x+0.04,1.36,w-0.06,0.24,sz=SZ_XS+1,bold=True,col=WHITE,align=PP_ALIGN.CENTER if len(h)<=3 else PP_ALIGN.LEFT)
    y = 1.68; rh = 0.32
    for i,(code,crit,cat) in enumerate(criteria):
        bg = WHITE if i%2==0 else TEAL_L
        R(s,cxs[0],y,cws[0],rh,fill=TEAL_D)
        T(s,code,cxs[0]+0.04,y+0.05,cws[0]-0.06,rh-0.08,sz=SZ_XS,bold=True,col=WHITE,align=PP_ALIGN.CENTER)
        R(s,cxs[1],y,cws[1],rh,fill=bg)
        T(s,crit,cxs[1]+0.08,y+0.05,cws[1]-0.12,rh-0.06,sz=SZ_XS+1,col=DARK)
        R(s,cxs[2],y,cws[2],rh,fill=cat_c.get(cat,TEAL))
        T(s,cat,cxs[2]+0.06,y+0.05,cws[2]-0.1,rh-0.06,sz=SZ_XS,bold=True,col=WHITE,align=PP_ALIGN.CENTER)
        for x,w in zip(cxs[3:],cws[3:]): R(s,x,y,w,rh,fill=WHITE,line=TEAL,lw=Pt(0.4))
        y += rh+0.01
    R(s,cxs[0],y,9.4,0.34,fill=TEAL_D)
    T(s,"TOTAL  /24",cxs[0]+0.12,y+0.06,9.0,0.22,sz=SZ_XS+2,bold=True,col=WHITE)
    R(s,cxs[3],y,sum(cws[3:])+0.04,0.34,fill=TEAL_L,line=TEAL_D,lw=Pt(2))
    T(s,"__ / 24",cxs[3]+0.1,y+0.06,2.4,0.22,sz=SZ_XS+2,bold=True,col=TEAL_D,align=PP_ALIGN.CENTER)
    y += 0.38
    HL(s,"✅ Certifié(e) FIM : ≥ 17/24  ·  ⚠ Formation complémentaire : 12–16/24  ·  🔄 Rattrapage J+15 : < 12/24",
       0.3,y+0.04,12.73,0.36,fill=TEAL_D,sz=SZ_XS+2)

    QUIZ_SLIDE(p, "J5 — Certification FIM (5 Questions)", QJ5, ACCENT)

    # ═══════════════════════════════════════════════════════
    #  CHAPITRE 6 — ATELIER PRATIQUE 1 (J6)
    # ═══════════════════════════════════════════════════════
    CHAPTER_DIVIDER(p, "06", "Atelier Pratique 1 — Jour 6", "Production simulée · Grille live · Débrief individuel", RGBColor(0x22,0x66,0x66))

    s = p.slides.add_slide(SL(p))
    H_CONTENT(s, "Programme Atelier J6 — Appels en Conditions Réelles Simulées")
    prg6 = [
        ("08h30–09h00","Briefing production","Rappel KPI · Répartition fichiers FIDELIS · Règles qualification"),
        ("09h00–11h00","Production Round 1","Binômes : 1 appelle, 1 observe avec grille. Rotation 30 min."),
        ("11h00–11h30","Débrief Round 1","3 points forts / 3 axes amélioration · Écoute extraits"),
        ("11h30–12h30","Production Round 2","Appels individuels. Formateur écoute en silence. Grille 24 critères."),
        ("13h30–14h30","Cas complexes Cat.5","Simulations objections conflictuelles · Escalade progressive"),
        ("14h30–15h30","Débrief final & scoring","Grilles individuelles restituées · Feedback · Plan d'action J7"),
    ]
    y = 0.96
    for tm, ti, de in prg6:
        R(s, 0.3, y, 2.2, 0.65, fill=RGBColor(0x22,0x66,0x66))
        T(s, tm, 0.36, y+0.12, 2.08, 0.44, sz=SZ_XS+2, bold=True, col=WHITE, align=PP_ALIGN.CENTER)
        R(s, 2.52, y, 3.0, 0.65, fill=TEAL)
        T(s, ti, 2.62, y+0.14, 2.8, 0.44, sz=SZ_SM, bold=True, col=WHITE)
        R(s, 5.54, y, 7.49, 0.65, fill=WHITE, line=TEAL, lw=Pt(0.8))
        T(s, de, 5.66, y+0.1, 7.3, 0.5, sz=SZ_B, col=DARK)
        y += 0.68
    HL(s, "Après chaque appel : 1 force · 1 axe d'amélioration · 1 conseil actionnable précis.",
       0.3, 5.08, 12.73, 0.46, fill=TEAL_D, sz=SZ_SM)

    # ═══════════════════════════════════════════════════════
    #  CHAPITRE 7 — ATELIER PRATIQUE 2 (J7)
    # ═══════════════════════════════════════════════════════
    CHAPTER_DIVIDER(p, "07", "Atelier Pratique 2 — Jour 7", "Perfectionnement · Certification finale · Plan prise de poste J+30", RGBColor(0x11,0x55,0x55))

    s = p.slides.add_slide(SL(p))
    H_CONTENT(s, "Programme Atelier J7 — Perfectionnement & Certification")
    prg7 = [
        ("08h30–09h30","Révision ciblée","Points faibles J6 · Verrouillage · Objections Cat.5 · Règles PA"),
        ("09h30–11h00","Simulation certifiante","Appel complet · Grille 24 critères · Évaluateur externe"),
        ("11h00–11h30","Résultats","Scores certifiants · Feedback individuel · Attestation FIM remise"),
        ("11h30–12h30","Quiz Final 40Q","Évaluation consolidée · Seuil 28/40 · Corrigé collectif"),
        ("13h30–14h30","Plan de poste","Construction plan J+0 à J+30 : objectifs semaine 1 à 4"),
        ("14h30–15h30","Cérémonie clôture","Remise certifications · Engagement public · Votre 1er appel demain"),
    ]
    y = 0.96
    for tm, ti, de in prg7:
        R(s, 0.3, y, 2.2, 0.65, fill=RGBColor(0x11,0x55,0x55))
        T(s, tm, 0.36, y+0.12, 2.08, 0.44, sz=SZ_XS+2, bold=True, col=WHITE, align=PP_ALIGN.CENTER)
        R(s, 2.52, y, 3.0, 0.65, fill=TEAL)
        T(s, ti, 2.62, y+0.14, 2.8, 0.44, sz=SZ_SM, bold=True, col=WHITE)
        R(s, 5.54, y, 7.49, 0.65, fill=WHITE, line=TEAL, lw=Pt(0.8))
        T(s, de, 5.66, y+0.1, 7.3, 0.5, sz=SZ_B, col=DARK)
        y += 0.68

    s = p.slides.add_slide(SL(p))
    H_CONTENT(s, "Votre Plan de Prise de Poste — J+0 à J+30")
    weeks = [
        ("J+0\nDemain", TEAL_D,
         "Premier appel de production\n9 CU/H dès J1\nQualification FIDELIS\nVerrouillage sur chaque appel"),
        ("SEMAINE 1\nJ+7", TEAL,
         "TX Transfo ≥ cible FIDELIS\n0 disqualification nomenclature\nFeedback formateur J+7\nPremier PA IBAN à chaud"),
        ("SEMAINE 2\nJ+14", TEAL_M,
         "0 objection Cat.5 non gérée\nPDC mensuel en bonne voie\nPoint superviseur FIDELIS\nAutoévaluation grille 24 crit."),
        ("SEMAINE 3\nJ+21", TEAL_A,
         "Autonomie complète script\nVerrouillage ≥ 50% promesses\nRelances Call 2 maîtrisées\nFeedback pair formateur"),
        ("SEMAINE 4\nJ+30", ACCENT,
         "Bilan formateur SHY-Performance\nRévision objectifs M2\nCertification terrain confirmée\nKPI FIDELIS × UNICEF France"),
    ]
    x = 0.3
    for tag, col, body in weeks:
        R(s, x, 0.96, 2.46, 0.58, fill=col)
        T(s, tag, x+0.1, 0.98, 2.26, 0.54, sz=SZ_SM, bold=True, col=WHITE, align=PP_ALIGN.CENTER)
        R(s, x, 1.54, 2.46, 4.4, fill=WHITE, line=col, lw=Pt(1.5))
        T(s, body, x+0.14, 1.64, 2.18, 4.18, sz=SZ_B, col=DARK, wrap=True)
        x += 2.56
    HL(s, "SHY-Performance vous contacte à J+15 et J+30. À J+90 : analyse KPI terrain consolidée FIDELIS × UNICEF France.",
       0.3, 6.1, 12.73, 0.46, sz=SZ_SM)

    # ═══════════════════════════════════════════════════════
    #  QUIZ FINAL — 40 QUESTIONS (10 slides × 4Q)
    # ═══════════════════════════════════════════════════════
    CHAPTER_DIVIDER(p, "QUIZ", "Quiz Final — 40 Questions", "Évaluation consolidée J1→J7 · Seuil 28/40 · Mention Excellent ≥ 35/40", TEAL_D)

    # Slide d'intro quiz
    s = p.slides.add_slide(SL(p))
    H_CONTENT(s, "Quiz Final — Modalités & Organisation")
    T(s, "Nom de l'apprenant : ___________________________________",
       0.5, 0.96, 9, 0.48, sz=SZ_B, bold=True, col=DARK)
    T(s, "Date : _______________________",
       0.5, 1.52, 5, 0.44, sz=SZ_B, bold=True, col=DARK)
    T(s, "Score : ______ / 40",
       0.5, 2.0, 5, 0.48, sz=SZ_B, bold=True, col=TEAL_D)
    parts = [("Q1–Q10","Fondations & Mission + Secteur Associatif",TEAL_D),
             ("Q11–Q20","Script Officiel, Accroche & Types de PA",TEAL),
             ("Q21–Q30","Objections, Verrouillage & Règles opérationnelles",TEAL_M),
             ("Q31–Q40","Certification, Posture & Vocabulaire métier",TEAL_A)]
    y = 2.62
    for qr, desc, col in parts:
        R(s, 0.5, y, 2.5, 0.52, fill=col)
        T(s, qr, 0.62, y+0.08, 2.26, 0.36, sz=SZ_SM, bold=True, col=WHITE, align=PP_ALIGN.CENTER)
        R(s, 3.02, y, 9.8, 0.52, fill=WHITE, line=col, lw=Pt(0.8))
        T(s, desc, 3.14, y+0.1, 9.6, 0.36, sz=SZ_B, col=DARK)
        y += 0.58
    HL(s, "✅ Validation : ≥ 28/40 (70%)  ·  ⭐ Excellent : ≥ 35/40  ·  🔄 Rattrapage J+15 : < 28/40",
       0.5, 5.0, 12.33, 0.52, fill=TEAL_D, sz=SZ_SM, bold=True)
    HL(s, "Durée : 45 minutes · Individuel · Corrigé collectif en fin de séance avec le formateur",
       0.5, 5.66, 12.33, 0.46, fill=TEAL_M, sz=SZ_SM)

    # 40 questions — 4 par slide
    for slide_i in range(0, len(Q40), 4):
        batch = Q40[slide_i:slide_i+4]
        s = p.slides.add_slide(SL(p))
        R(s, 0, 0, 13.33, 7.5, fill=TEAL_L)
        R(s, 0, 0, 13.33, 0.88, fill=TEAL_D)
        R(s, 0, 0.86, 13.33, 0.06, fill=TEAL_A)
        T(s, CO, 0.18, 0.08, 2.3, 0.46, sz=SZ_SM+2, bold=True, col=WHITE)
        R(s, 2.6, 0.1, 0.05, 0.66, fill=TEAL_A)
        T(s, f"Quiz Final — Questions {slide_i+1}–{slide_i+len(batch)}", 2.76, 0.06, 10.3, 0.7, sz=SZ_T, bold=True, col=WHITE)
        LOGO_ADD(s); FOOTER(s)
        y = 0.98
        for i, (q, opts) in enumerate(batch):
            qn = slide_i + i + 1
            R(s, 0.3, y, 0.52, 0.36, fill=TEAL_D)
            T(s, f"Q{qn}", 0.34, y+0.04, 0.44, 0.28, sz=SZ_XS+2, bold=True, col=WHITE, align=PP_ALIGN.CENTER)
            R(s, 0.84, y, 12.19, 0.36, fill=TEAL_D)
            T(s, q, 0.96, y+0.04, 12.0, 0.28, sz=SZ_SM, bold=True, col=WHITE, wrap=True)
            oy = y + 0.38
            for j, opt in enumerate(opts):
                bg = WHITE if j % 2 == 0 else TEAL_L
                R(s, 0.84, oy, 12.19, 0.3, fill=bg, line=TEAL_D, lw=Pt(0.3))
                T(s, opt, 0.96, oy+0.04, 12.0, 0.22, sz=SZ_XS+2, col=DARK)
                oy += 0.32
            y += 1.7

    # ─────────────────────────────────────────────────────────
    #  SLIDE FINALE — MERCI
    # ─────────────────────────────────────────────────────────
    s = p.slides.add_slide(SL(p))
    R(s, 0, 0, 13.33, 7.5, fill=TEAL_D)
    R(s, 0, 0, 13.33, 0.08, fill=TEAL_A)
    R(s, 0, 7.42, 13.33, 0.08, fill=TEAL_A)
    sh = s.shapes.add_shape(9, Inches(4.5), Inches(0.8), Inches(4.3), Inches(4.3))
    sh.fill.solid(); sh.fill.fore_color.rgb = TEAL_M; sh.line.fill.background()
    T(s, "🙏", 5.5, 1.7, 2.3, 1.5, sz=48, align=PP_ALIGN.CENTER, col=WHITE)
    T(s, "Merci à Tous !", 1.5, 1.0, 10.3, 1.2, sz=46, bold=True, col=WHITE, align=PP_ALIGN.CENTER)
    T(s, "Votre engagement fait la différence pour les enfants du monde entier.",
       1.5, 2.58, 10.3, 0.6, sz=SZ_B, col=RGBColor(0xCC,0xEE,0xEE), align=PP_ALIGN.CENTER, italic=True)
    T(s, "Module Général FIM — Formation complète J1 → J7",
       1.5, 3.28, 10.3, 0.38, sz=SZ_SM, col=RGBColor(0x88,0xDD,0xDD), align=PP_ALIGN.CENTER)
    T(s, "« Le refus d'aujourd'hui peut être le don de demain »",
       2, 3.78, 9.33, 0.5, sz=SZ_SM+2, italic=True, col=RGBColor(0xAA,0xDD,0xDD), align=PP_ALIGN.CENTER)
    T(s, f"{CO}  ×  FIDELIS  ×  UNICEF France",
       0, 4.46, 13.33, 0.44, sz=SZ_S, bold=True, col=WHITE, align=PP_ALIGN.CENTER)
    LOGO_ADD(s, x=6.2, y=5.2, w=1.1, h=1.1)
    T(s, f"Version {VER}", 0, 6.6, 13.33, 0.36, sz=SZ_XS+2, col=RGBColor(0x88,0xDD,0xDD), align=PP_ALIGN.CENTER)

    return save(p, f"FIM_MODULE_GENERAL_{VER}.pptx")

if __name__ == "__main__":
    print(f"\n🎨  FIM {VER} — MODULE GÉNÉRAL CONSOLIDÉ\n")
    path = build_module_general()
    print(f"\n✅  Module général : {os.path.basename(path)}")
    print(f"    Taille : {os.path.getsize(path)//1024} KB")

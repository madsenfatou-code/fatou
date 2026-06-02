"""
FIM V6-FINAL-PLUS — Générateur PPTX — 9 modules
- Logo SHY-Performance sur CHAQUE slide (entête)
- Quiz 40 questions (module dédié Book Quiz)
- Slide "Merci à tous" en dernière position de chaque module
- 2 jours d'ateliers pratiques (J6 + J7)
- Analyse écarts UNICEF Paris + recommandations intégrées
Charte : teal #008080 uniquement · Prospection téléphonique · Données 19-23/05/2026
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
import os

# ─── CHARTE ─────────────────────────────────────────────────────────────────
TEAL     = RGBColor(0x00, 0x80, 0x80)
TEAL_D   = RGBColor(0x00, 0x50, 0x50)
TEAL_M   = RGBColor(0x00, 0x6A, 0x6A)
TEAL_L   = RGBColor(0xF0, 0xF7, 0xF7)
TEAL_A   = RGBColor(0x00, 0xAA, 0xAA)
WHITE    = RGBColor(0xFF, 0xFF, 0xFF)
DARK     = RGBColor(0x1A, 0x1A, 0x1A)
GRAY     = RGBColor(0x55, 0x55, 0x55)
LGRAY    = RGBColor(0xDD, 0xDD, 0xDD)

LOGO     = "/home/user/fatou/shy_logo_v2.png"
OUT      = "/home/user/fatou"
VER      = "V6-01062026SHY-TE"
CO       = "SHY-Performance"

# ─── PRIMITIVES ──────────────────────────────────────────────────────────────
def prs():
    p = Presentation()
    p.slide_width  = Inches(13.33)
    p.slide_height = Inches(7.5)
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

def T(s, text, x, y, w, h, sz=13, bold=False, col=None,
      align=PP_ALIGN.LEFT, italic=False, wrap=True):
    col = col or DARK
    b = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    b.word_wrap = wrap
    tf = b.text_frame; tf.word_wrap = wrap
    p2 = tf.paragraphs[0]; p2.alignment = align
    r = p2.add_run()
    r.text = text; r.font.size = Pt(sz)
    r.font.bold = bold; r.font.italic = italic
    r.font.color.rgb = col
    return b

def LOGO_ADD(s, x=12.28, y=0.0, w=0.92, h=0.74):
    """Logo sur chaque slide — coin supérieur droit"""
    try:
        s.shapes.add_picture(LOGO, Inches(x), Inches(y), Inches(w), Inches(h))
    except Exception:
        pass

def FOOTER(s):
    R(s, 0, 7.1, 13.33, 0.4, fill=TEAL_D)
    T(s, f"{CO}  ·  Formation FIM  ·  UNICEF France  ·  {VER}",
      0, 7.13, 13.33, 0.3, sz=8, col=WHITE, align=PP_ALIGN.CENTER)

# ─── ENTÊTES ─────────────────────────────────────────────────────────────────
def H_TITLE(s, title, sub="", tag=""):
    R(s, 0, 0, 13.33, 7.5, fill=TEAL_L)
    R(s, 0, 0, 13.33, 2.55, fill=TEAL_D)
    R(s, 0, 2.53, 13.33, 0.06, fill=TEAL_A)
    T(s, CO, 0.38, 0.1, 5, 0.44, sz=22, bold=True, col=WHITE)
    T(s, "Formation Initiale Module — Fundraising Téléphonique UNICEF",
      0.38, 0.54, 9, 0.36, sz=10, col=RGBColor(0xBB,0xEE,0xEE), italic=True)
    if tag: T(s, tag, 0.38, 0.96, 9, 0.3, sz=9, bold=True, col=RGBColor(0x88,0xDD,0xDD))
    T(s, title, 0.38, 1.3, 11.5, 1.05, sz=29, bold=True, col=WHITE)
    if sub: T(s, sub, 0.38, 2.1, 11.5, 0.42, sz=12, col=RGBColor(0xCC,0xEE,0xEE), italic=True)
    LOGO_ADD(s)
    FOOTER(s)

def H_CONTENT(s, title):
    R(s, 0, 0, 13.33, 7.5, fill=TEAL_L)
    R(s, 0, 0, 13.33, 0.75, fill=TEAL_D)
    R(s, 0, 0.73, 13.33, 0.05, fill=TEAL_A)
    T(s, CO, 0.18, 0.1, 2.1, 0.36, sz=14, bold=True, col=WHITE)
    R(s, 2.42, 0.12, 0.04, 0.52, fill=TEAL_A)
    T(s, title, 2.57, 0.1, 9.65, 0.55, sz=16, bold=True, col=WHITE)
    LOGO_ADD(s)
    FOOTER(s)

# ─── SLIDE MERCI ─────────────────────────────────────────────────────────────
def MERCI(p, module_tag=""):
    s = p.slides.add_slide(SL(p))
    R(s, 0, 0, 13.33, 7.5, fill=TEAL_D)
    R(s, 0, 0, 13.33, 0.08, fill=TEAL_A)
    R(s, 0, 7.42, 13.33, 0.08, fill=TEAL_A)
    # Grand cercle décoratif
    sh = s.shapes.add_shape(9, Inches(4.5), Inches(1.2), Inches(4.3), Inches(4.3))
    sh.fill.solid(); sh.fill.fore_color.rgb = TEAL_M
    sh.line.fill.background()
    T(s, "🙏", 5.5, 2.2, 2.3, 1.5, sz=48, align=PP_ALIGN.CENTER, col=WHITE)
    T(s, "Merci à Tous !", 1.5, 1.4, 10.3, 1.2,
      sz=46, bold=True, col=WHITE, align=PP_ALIGN.CENTER)
    T(s, "Votre engagement fait la différence pour les enfants du monde entier.",
      1.5, 2.85, 10.3, 0.6, sz=18, col=RGBColor(0xCC,0xEE,0xEE),
      align=PP_ALIGN.CENTER, italic=True)
    if module_tag:
        T(s, module_tag, 1.5, 3.55, 10.3, 0.38, sz=12,
          col=RGBColor(0x88,0xDD,0xDD), align=PP_ALIGN.CENTER)
    T(s, "« Le refus d'aujourd'hui peut être le don de demain »",
      2, 4.1, 9.33, 0.5, sz=14, italic=True,
      col=RGBColor(0xAA,0xDD,0xDD), align=PP_ALIGN.CENTER)
    T(s, f"{CO}  ·  Formation FIM  ·  UNICEF France  ·  {VER}",
      0, 5.0, 13.33, 0.36, sz=10, col=RGBColor(0x88,0xDD,0xDD),
      align=PP_ALIGN.CENTER)
    LOGO_ADD(s, x=6.2, y=5.6, w=1.0, h=0.98)
    T(s, "SHY-Performance  ×  FIDELIS  ×  UNICEF France",
      0, 6.6, 13.33, 0.36, sz=10, bold=True,
      col=WHITE, align=PP_ALIGN.CENTER)

# ─── COMPOSANTS ──────────────────────────────────────────────────────────────
def TWO_COL(s, lt, li, rt, ri, y=0.88):
    cw = 6.18
    for head, items, ox in [(lt, li, 0.3), (rt, ri, 6.83)]:
        R(s, ox, y, cw, 0.36, fill=TEAL)
        T(s, head, ox+0.12, y+0.05, cw-0.2, 0.28, sz=10, bold=True, col=WHITE)
        iy = y + 0.38
        for i, item in enumerate(items):
            bg = WHITE if i%2==0 else TEAL_L
            R(s, ox, iy, cw, 0.44, fill=bg, line=TEAL, lw=Pt(0.4))
            T(s, "• "+item, ox+0.14, iy+0.06, cw-0.24, 0.34, sz=10, col=DARK)
            iy += 0.46

def SEQ4(s, steps, labels, y=0.88, h=5.1):
    colors = [TEAL_D, TEAL, TEAL_M, TEAL_A]
    sw = (13.33-0.6)/4
    for i,(step,label,col) in enumerate(zip(steps,labels,colors)):
        bx = 0.3+i*(sw+0.02)
        R(s, bx, y, sw, 0.32, fill=col)
        T(s, label, bx+0.1, y+0.04, sw-0.18, 0.26, sz=9, bold=True, col=WHITE)
        R(s, bx, y+0.32, sw, h-0.32, fill=WHITE, line=col, lw=Pt(1.2))
        T(s, step, bx+0.12, y+0.38, sw-0.22, h-0.55, sz=9.5, col=DARK, wrap=True)

def HL(s, text, x, y, w, h, fill=None, tc=None):
    fill = fill or TEAL; tc = tc or WHITE
    R(s, x, y, w, h, fill=fill)
    T(s, text, x+0.15, y+0.08, w-0.3, h-0.16, sz=10, col=tc, wrap=True)

def KPI_BAR(s, items, y=6.08, label="À retenir"):
    R(s, 0.3, y, 12.73, 0.28, fill=TEAL)
    T(s, label, 0.42, y+0.04, 12.5, 0.22, sz=8, bold=True, col=WHITE)
    w = 12.73/len(items)
    for i,item in enumerate(items):
        bg = WHITE if i%2==0 else TEAL_L
        R(s, 0.3+i*w, y+0.3, w-0.04, 0.55, fill=bg, line=TEAL, lw=Pt(0.6))
        T(s, "✔ "+item, 0.42+i*w, y+0.34, w-0.18, 0.46, sz=9, col=DARK, wrap=True)

def TH(s, cols, cxs, cws, y):
    for h,x,w in zip(cols,cxs,cws):
        R(s, x, y, w, 0.3, fill=TEAL_D)
        T(s, h, x+0.06, y+0.04, w-0.1, 0.24, sz=9, bold=True, col=WHITE,
          align=PP_ALIGN.CENTER if len(h)<=3 else PP_ALIGN.LEFT)

def save(p, name):
    path = os.path.join(OUT, name)
    p.save(path)
    print(f"  ✓ {name}  ({os.path.getsize(path)//1024} KB)")
    return path


# ══════════════════════════════════════════════════════════════════════════
#  QUIZ 40 QUESTIONS — MODULE DÉDIÉ
# ══════════════════════════════════════════════════════════════════════════

QUIZ = [
    # Partie A — UNICEF & Cause (Q1–Q8)
    ("A", "En quelle année l'UNICEF a-t-il été fondé ?",
     ["a) 1901","b) 1946 ✓","c) 1959","d) 1919"]),
    ("A", "Dans combien de pays l'UNICEF est-il présent ?",
     ["a) 120 pays","b) 150 pays","c) 190 pays ✓","d) 210 pays"]),
    ("A", "Quel % des enfants du monde l'UNICEF vaccine-t-il ?",
     ["a) 25%","b) 35%","c) 45% ✓","d) 60%"]),
    ("A", "Que signifie RUTF dans le traitement de la malnutrition ?",
     ["a) Rapid Urban Treatment Fund","b) Ready-to-Use Therapeutic Food ✓","c) Regional Unified Task Force","d) Routine Universal Treatment Formula"]),
    ("A", "Quel est le poids standard d'un sachet RUTF ?",
     ["a) 45g","b) 75g","c) 92g ✓","d) 120g"]),
    ("A", "Quelle est la cause humanitaire centrale de la campagne UNICEF portée par SHY-Performance ?",
     ["a) Accès à l'éducation","b) Lutte contre la malnutrition aigüe sévère ✓","c) Eau potable en Afrique","d) Protection contre le travail des enfants"]),
    ("A", "Quelle organisation a précédé et inspiré la création de l'UNICEF après la Seconde Guerre mondiale ?",
     ["a) La Croix-Rouge","b) La Société des Nations (SDN)","c) L'UNRRA ✓","d) Le Fond des Nations Unies"]),
    ("A", "Lequel NE fait PAS partie des missions principales de l'UNICEF ?",
     ["a) Vaccination et santé","b) Eau potable et assainissement","c) Financement des partis politiques ✓","d) Protection de l'enfance et nutrition"]),
    # Partie B — Secteur associatif (Q9–Q15)
    ("B", "En quelle date a été promulguée la loi fondatrice du droit associatif français ?",
     ["a) 14 juillet 1789","b) 1er juillet 1901 ✓","c) 5 mai 1946","d) 10 mars 1972"]),
    ("B", "Combien d'associations sont recensées en France ?",
     ["a) 250 000","b) 800 000","c) 1,5 million ✓","d) 3 millions"]),
    ("B", "Quel événement de 1859 est l'acte fondateur de l'humanitaire moderne ?",
     ["a) La conférence de Berlin","b) Le Traité de Genève","c) La Bataille de Solférino ✓","d) La Révolution industrielle"]),
    ("B", "Quel type d'association bénéficie du statut le plus reconnu par l'État ?",
     ["a) Association de fait","b) Association déclarée","c) Association reconnue d'utilité publique ✓","d) Association internationale"]),
    ("B", "Quel est le rôle du Comité de la Charte du Don en Confiance ?",
     ["a) Gérer les finances publiques","b) Certifier la transparence et l'éthique des organisations ✓","c) Former les agents de collecte","d) Contrôler le respect du RGPD"]),
    ("B", "Combien de bénévoles compte le secteur associatif français ?",
     ["a) 5 millions","b) 12 millions","c) 22 millions ✓","d) 35 millions"]),
    ("B", "Lequel NE figure PAS dans les 3 principes humanitaires fondamentaux ?",
     ["a) Humanité","b) Impartialité","c) Solidarité nationale ✓","d) Neutralité"]),
    # Partie C — Script & Accroche (Q16–Q22)
    ("C", "Combien d'étapes comporte le script officiel UNICEF ?",
     ["a) 4 étapes","b) 5 étapes","c) 7 étapes ✓","d) 10 étapes"]),
    ("C", "Combien de règles officielles régissent la phrase d'accroche ?",
     ["a) 2 règles","b) 3 règles","c) 5 règles ✓","d) 7 règles"]),
    ("C", "Quelle est la première étape du script après la présentation ?",
     ["a) Demande de coordonnées bancaires","b) Accroche sur la mission UNICEF ✓","c) Présentation du montant minimum","d) Vérification de l'identité"]),
    ("C", "Combien de modes de validation du don sont disponibles ?",
     ["a) 2 modes","b) 3 modes","c) 4 modes ✓","d) 6 modes"]),
    ("C", "Lequel NE fait PAS partie des 4 modes de validation ?",
     ["a) PA à chaud IBAN","b) PA en ligne en direct","c) Virement bancaire unique différé ✓","d) PA courrier"]),
    ("C", "À quelle étape le fundraiser aborde-t-il les sachets RUTF ?",
     ["a) Étape 7 (validation coordonnées)","b) Étapes 2-3 (malnutrition/RUTF) ✓","c) Uniquement en cas d'objection financière","d) Dès la première phrase d'accroche"]),
    ("C", "Quel produit unique est proposé lors de chaque appel ?",
     ["a) Don ponctuel par carte","b) Prélèvement automatique régulier (PA) ✓","c) Abonnement mensuel par chèque","d) Parrainage ponctuel d'enfant"]),
    # Partie D — KPI & Nomenclature (Q23–Q28)
    ("D", "Quel est l'objectif minimum de Contacts Utiles par Heure (CU/H) ?",
     ["a) 5 CU/H","b) 7 CU/H","c) 9 CU/H ✓","d) 12 CU/H"]),
    ("D", "Que représente 'TX Transfo' dans la nomenclature métier ?",
     ["a) Nombre total de transferts de dossiers","b) Taux de transformation (% PEL/PA parmi les CU) ✓","c) Temps moyen de traitement par appel","d) Taux de transfert vers un superviseur"]),
    ("D", "Quelles sont les 3 qualifications d'un Contact Utile (CU) ?",
     ["a) Positif / Neutre / Négatif","b) DON / INDÉCIS / REFUS ✓","c) Chaud / Tiède / Froid","d) PA / PEL / PDC"]),
    ("D", "Quelle est la durée effective de production lors d'une journée standard ?",
     ["a) 5h00","b) 6h00","c) 6h40 ✓","d) 8h00"]),
    ("D", "Que désigne 'PDC' dans la nomenclature de la campagne ?",
     ["a) Protocole de Communication","b) Plan de Charge mensuel ✓","c) Programme de Collecte de Dons","d) Priorité de Campagne Définie"]),
    ("D", "Différence entre 'Conquête' et 'Fidélisation / Réactivation' ?",
     ["a) La Conquête cible des anciens donateurs","b) La Conquête cible de nouveaux prospects, la Fidélisation cible les anciens ✓","c) Les deux ciblent les mêmes fichiers","d) La Conquête est réservée aux appels entrants"]),
    # Partie E — Objections & AAR (Q29–Q35)
    ("E", "Combien de catégories d'objections sont enseignées dans la FIM ?",
     ["a) 3 catégories","b) 5 catégories ✓","c) 7 catégories","d) 10 catégories"]),
    ("E", "Que signifie l'acronyme AAR ?",
     ["a) Analyser / Adapter / Répondre","b) Accepter / Argumenter / Reformuler","c) Accuser réception / Argumenter / Relancer ✓","d) Approcher / Assurer / Recadrer"]),
    ("E", "Combien d'objections au total couvre la FIM ?",
     ["a) 8 objections","b) 10 objections","c) 15 objections ✓","d) 20 objections"]),
    ("E", "Quelle catégorie regroupe 'Vous n'avez pas le droit de m'appeler' ?",
     ["a) Catégorie 2 — Financières","b) Catégorie 3 — Contre le PA","c) Catégorie 4 — Méfiance IBAN","d) Catégorie 5 — Conflictuelles ✓"]),
    ("E", "Combien d'objections initiales (phase d'accroche) sont listées dans la Cat.1 ?",
     ["a) 4 objections","b) 6 objections","c) 8 objections ✓","d) 10 objections"]),
    ("E", "Lors du traitement d'une objection financière, quelle preuve invoquer ?",
     ["a) Le nombre de pays où l'UNICEF est présent","b) L'impact chiffré du PA (traitement 1 enfant en 6 semaines) ✓","c) Le nombre de bénévoles en France","d) La date de fondation de l'UNICEF"]),
    ("E", "Qu'est-ce que Bloctel ?",
     ["a) Un logiciel de CRM pour les fundraisers","b) La liste d'opposition au démarchage téléphonique à consulter obligatoirement ✓","c) Le nom du script officiel UNICEF","d) Un label de qualité pour les associations"]),
    # Partie F — Éthique & Posture (Q36–Q40)
    ("F", "Quelle règle éthique s'applique systématiquement lors d'un appel UNICEF ?",
     ["a) Insister si le prospect hésite après la première réponse","b) Proposer un don ponctuel si le PA est refusé définitivement","c) Respecter immédiatement tout refus définitif et clore l'appel avec courtoisie ✓","d) Enregistrer l'appel sans informer le prospect"]),
    ("F", "Que signifie PEL dans la nomenclature FIDELIS ?",
     ["a) Plan d'Envoi de Lien","b) Promesse En Ligne (don validé en ligne après l'appel) ✓","c) Processus d'Engagement Légal","d) Protocole d'Écoute Longitudinale"]),
    ("F", "Quelle est la posture vocale attendue tout au long de l'appel ?",
     ["a) Autoritaire et directive pour capter l'attention","b) Neutre et distante pour rester professionnel","c) Chaleureuse, souriante et naturelle — le sourire s'entend ✓","d) Rapide et concise pour ne pas lasser le prospect"]),
    ("F", "Quel mantra résume la philosophie de gestion du refus dans la FIM ?",
     ["a) 'Un refus est un refus — passez au suivant'","b) 'Le refus d'aujourd'hui peut être le don de demain' ✓","c) 'Insistez toujours — la persévérance paie'","d) 'Le silence vaut mieux qu'un refus argumenté'"]),
    ("F", "Quelle est la durée standard d'un appel de collecte complet (script + closing) ?",
     ["a) 30 secondes à 1 minute","b) 1 à 2 minutes","c) 3 à 5 minutes ✓","d) 8 à 10 minutes"]),
]

ANSWERS = {
    "A": [2,3,3,2,3,2,3,3],
    "B": [2,3,3,3,2,3,3],
    "C": [3,3,2,3,3,2,2],
    "D": [3,2,2,3,2,2],
    "E": [2,3,3,4,3,2,2],
    "F": [3,2,3,2,3],
}


def build_quiz():
    p = prs()
    parts = {
        "A": ("Partie A", "Connaissance de l'UNICEF & de la Cause", "Q1–Q8",   TEAL_D),
        "B": ("Partie B", "Secteur Associatif & Humanitaire",       "Q9–Q15",  TEAL),
        "C": ("Partie C", "Script officiel & Règles d'Accroche",    "Q16–Q22", TEAL_M),
        "D": ("Partie D", "KPIs, Nomenclature FIDELIS & Production","Q23–Q28", TEAL_A),
        "E": ("Partie E", "Traitement des Objections & Méthode AAR","Q29–Q35", RGBColor(0x33,0x77,0x77)),
        "F": ("Partie F", "Éthique, Posture & Vocabulaire métier",  "Q36–Q40", RGBColor(0x22,0x66,0x66)),
    }

    # Slide 1 — Titre quiz
    s = p.slides.add_slide(SL(p))
    H_TITLE(s,
        "Quiz d'Évaluation Finale — 40 Questions",
        "Formation Initiale Module FIM  ·  SHY-Performance × FIDELIS × UNICEF France",
        tag="BOOK QUIZ  ·  Évaluation finale des connaissances")
    T(s, "Nom de l'apprenant : ___________________________________",
      0.5, 2.75, 8, 0.4, sz=13, bold=True, col=DARK)
    T(s, "Date : _______________________",
      0.5, 3.2, 5, 0.4, sz=13, bold=True, col=DARK)
    T(s, "Score : ______ / 40    Seuil de validation : 28/40 (70%)",
      0.5, 3.65, 8, 0.4, sz=13, bold=True, col=TEAL_D)
    parts_info = [
        ("Partie A", "UNICEF & Cause", "Q1–Q8"),
        ("Partie B", "Secteur associatif", "Q9–Q15"),
        ("Partie C", "Script & Accroche", "Q16–Q22"),
        ("Partie D", "KPI & Nomenclature", "Q23–Q28"),
        ("Partie E", "Objections & AAR", "Q29–Q35"),
        ("Partie F", "Éthique & Posture", "Q36–Q40"),
    ]
    x = 0.5
    for i, (pt, desc, qrange) in enumerate(parts_info):
        col = list(parts.values())[i][3]
        R(s, x, 4.25, 2.1, 0.5, fill=col)
        T(s, f"{pt}\n{desc}\n{qrange}", x+0.08, 4.28, 1.94, 0.44,
          sz=8, bold=True, col=WHITE, align=PP_ALIGN.CENTER)
        x += 2.14

    # Slides Q1–Q40 : 4 questions par slide
    q_list = [(i+1, q[0], q[1], q[2]) for i, q in enumerate(QUIZ)]
    for slide_i in range(0, len(q_list), 4):
        batch = q_list[slide_i:slide_i+4]
        # Determine part label from first q in batch
        part_key = batch[0][1]
        part_label = parts[part_key][0]
        part_title = parts[part_key][1]
        part_range = parts[part_key][2]
        part_col   = parts[part_key][3]

        s = p.slides.add_slide(SL(p))
        R(s, 0, 0, 13.33, 7.5, fill=TEAL_L)
        R(s, 0, 0, 13.33, 0.75, fill=TEAL_D)
        R(s, 0, 0.73, 13.33, 0.05, fill=TEAL_A)
        T(s, CO, 0.18, 0.1, 2.1, 0.36, sz=14, bold=True, col=WHITE)
        R(s, 2.42, 0.12, 0.04, 0.52, fill=TEAL_A)
        T(s, f"Quiz FIM — {part_label} : {part_title} ({part_range})",
          2.57, 0.1, 9.65, 0.55, sz=14, bold=True, col=WHITE)
        LOGO_ADD(s)
        FOOTER(s)

        y = 0.9
        per_row = 2  # 2 questions par rangée, 2 rangées = 4 questions
        cols = [0.3, 6.83]
        rows_done = 0
        for qi, (qnum, qpart, qtext, opts) in enumerate(batch):
            ox = cols[qi % 2]
            if qi % 2 == 0 and qi > 0:
                y += 2.85
            if rows_done >= 2:
                break
            # Question number block
            R(s, ox, y, 0.5, 2.6, fill=part_col)
            T(s, f"Q{qnum}", ox+0.04, y+0.85, 0.42, 0.6,
              sz=13, bold=True, col=WHITE, align=PP_ALIGN.CENTER)
            # Question text
            R(s, ox+0.52, y, 5.66, 0.65, fill=part_col)
            T(s, qtext, ox+0.62, y+0.06, 5.46, 0.55, sz=10, bold=True, col=WHITE, wrap=True)
            # Options
            for oi, opt in enumerate(opts):
                og = WHITE if oi%2==0 else TEAL_L
                R(s, ox+0.52, y+0.67+oi*0.48, 5.66, 0.46, fill=og, line=part_col, lw=Pt(0.4))
                T(s, opt, ox+0.66, y+0.72+oi*0.48, 5.42, 0.36, sz=9.5, col=DARK)
            if qi % 2 == 1:
                rows_done += 1

    # Slide corrigé — récapitulatif des bonnes réponses
    s = p.slides.add_slide(SL(p))
    H_CONTENT(s, "Corrigé — Récapitulatif des Bonnes Réponses")
    T(s, "Les réponses marquées ✓ dans le quiz sont les réponses correctes. "
       "Consultez votre formateur pour tout écart de compréhension.",
      0.35, 0.88, 12.6, 0.5, sz=11, col=DARK, italic=True)
    summary = [
        ("Q1–Q8\nPartie A",   "b/c/c/b/c/b/c/c",  TEAL_D),
        ("Q9–Q15\nPartie B",  "b/c/c/c/b/c/c",    TEAL),
        ("Q16–Q22\nPartie C", "c/c/b/c/c/b/b",    TEAL_M),
        ("Q23–Q28\nPartie D", "c/b/b/c/b/b",      TEAL_A),
        ("Q29–Q35\nPartie E", "b/c/c/d/c/b/b",    RGBColor(0x33,0x77,0x77)),
        ("Q36–Q40\nPartie F", "c/b/c/b/c",        RGBColor(0x22,0x66,0x66)),
    ]
    x = 0.3
    for label, answers, col in summary:
        R(s, x, 1.5, 2.08, 0.55, fill=col)
        T(s, label, x+0.08, 1.52, 1.92, 0.51, sz=10, bold=True,
          col=WHITE, align=PP_ALIGN.CENTER)
        R(s, x, 2.05, 2.08, 1.4, fill=WHITE, line=col, lw=Pt(1))
        T(s, answers, x+0.1, 2.12, 1.88, 1.25, sz=11, col=TEAL_D, wrap=True)
        x += 2.18
    HL(s,
       "Seuil de validation : 28/40 (70%)  ·  En dessous : révision ciblée + rattrapage J+15  "
       "·  Score ≥ 35/40 : mention Excellent",
       0.3, 3.65, 12.73, 0.44, fill=TEAL_D)
    HL(s,
       "Ce quiz évalue vos connaissances théoriques. Votre grille pratique (24 critères)\n"
       "est évaluée séparément lors des simulations d'appels en Journée 5.",
       0.3, 4.22, 12.73, 0.58, fill=TEAL_L, tc=DARK)

    MERCI(p, "Quiz d'Évaluation Finale — Formation FIM")
    return save(p, f"FIM_BookQUIZ_40Questions_{VER}.pptx")


# ══════════════════════════════════════════════════════════════════════════
#  J6 — ATELIER PRATIQUE 1 : APPELS EN CONDITIONS RÉELLES
# ══════════════════════════════════════════════════════════════════════════

def build_J6():
    p = prs()

    s = p.slides.add_slide(SL(p))
    H_TITLE(s,
        "Atelier Pratique — Jour 6",
        "Appels en conditions réelles simulées · Grille live · Débrief individuel",
        tag="BOOK J6  ·  Atelier 1  ·  Mise en situation téléphonique")
    T(s,
      "Aujourd'hui vous êtes en production simulée.\n"
      "Vous appelez, vous gérez, vous closez — en temps réel, avec retour immédiat.\n"
      "Objectif : atteindre 9 CU/H et votre premier PA validé.",
      0.5, 2.75, 11.8, 1.1, sz=14, col=DARK)

    s = p.slides.add_slide(SL(p))
    H_CONTENT(s, "Programme Atelier Pratique J6")
    prg = [
        ("08h30","09h00","Briefing production","Rappel des KPI cibles · Répartition des fichiers FIDELIS · Règles de qualification"),
        ("09h00","11h00","Production simulée — Round 1","Appels en binômes : 1 appelle, 1 observe avec grille. Rotation toutes les 30 min."),
        ("11h00","11h30","Débrief Round 1","Analyse collective des appels. 3 points forts / 3 axes d'amélioration. Écoute d'extraits."),
        ("11h30","12h30","Production simulée — Round 2","Appels individuels. Formateur écoute en silence et note sur grille 24 critères."),
        ("13h30","14h30","Traitement des cas complexes","Simulation d'appels avec objections Cat.5 (conflictuelles). Escalade progressive."),
        ("14h30","15h30","Débrief final & scoring","Restitution des grilles individuelles. Feedback personnalisé. Plan d'action J7."),
    ]
    y = 0.88
    for st,en,ti,de in prg:
        R(s, 0.3, y, 1.38, 0.68, fill=TEAL_D)
        T(s, f"{st}\n{en}", 0.32, y+0.08, 1.34, 0.54, sz=11, bold=True, col=WHITE, align=PP_ALIGN.CENTER)
        R(s, 1.7, y, 3.0, 0.68, fill=TEAL)
        T(s, ti, 1.8, y+0.15, 2.82, 0.42, sz=12, bold=True, col=WHITE)
        R(s, 4.72, y, 8.31, 0.68, fill=WHITE, line=TEAL, lw=Pt(0.8))
        T(s, de, 4.84, y+0.07, 8.1, 0.58, sz=9, col=DARK, wrap=True)
        y += 0.71

    s = p.slides.add_slide(SL(p))
    H_CONTENT(s, "Grille d'Observation Live J6 — Utilisée par Votre Formateur")
    TWO_COL(s,
        "Ce que le formateur observe sur votre appel",
        ["Règle 1 accroche : prénom + nom prononcés clairement",
         "Règle 2 : organisation UNICEF nommée dans les 10 premières secondes",
         "Règle 3 : raison de l'appel annoncée explicitement",
         "Règle 4 : disponibilité du prospect vérifiée",
         "Règle 5 : ton chaleureux et sourire audible",
         "Script 7 étapes respecté dans l'ordre",
         "AAR appliqué sur au moins 2 objections"],
        "Indicateurs de performance mesurés",
        ["CU/H atteint (objectif ≥ 9)",
         "Qualification correcte (DON / INDÉCIS / REFUS)",
         "Mode de validation proposé dans le bon ordre (IBAN → PEL → courrier)",
         "Durée appel dans la norme (3–5 min)",
         "Pas d'interruption ou de tic verbal",
         "Clôture de l'appel avec courtoisie même en cas de refus",
         "Résilience : ton maintenu après 3 refus consécutifs"])
    HL(s,
       "Après chaque appel : feedback en 3 points — 1 force · 1 axe d'amélioration · 1 conseil précis actionnable.",
       0.3, 6.05, 12.73, 0.42)

    MERCI(p, "Atelier Pratique Jour 6 — Formation FIM")
    return save(p, f"FIM_BookJ6_AtelierPratique1_{VER}.pptx")


# ══════════════════════════════════════════════════════════════════════════
#  J7 — ATELIER PRATIQUE 2 : PERFECTIONNEMENT & CERTIFICATION FINALE
# ══════════════════════════════════════════════════════════════════════════

def build_J7():
    p = prs()

    s = p.slides.add_slide(SL(p))
    H_TITLE(s,
        "Atelier Pratique — Jour 7",
        "Perfectionnement · Certification Finale · Prise de Poste",
        tag="BOOK J7  ·  Atelier 2  ·  Certification & Déploiement terrain")
    T(s,
      "Aujourd'hui c'est votre dernier jour de formation.\n"
      "Vous validez votre certification FIM finale et vous construisez votre plan\n"
      "de prise de poste opérationnel dès demain.",
      0.5, 2.75, 11.8, 1.1, sz=14, col=DARK)

    s = p.slides.add_slide(SL(p))
    H_CONTENT(s, "Programme Atelier Perfectionnement J7")
    prg = [
        ("08h30","09h30","Révision ciblée","Traitement des points faibles identifiés en J6 · Focus sur les objections Cat.5 · Réentraînement ciblé"),
        ("09h30","11h00","Simulation finale certifiante","Appel complet en conditions réelles simulées · Grille 24 critères · 1 évaluateur externe"),
        ("11h00","11h30","Résultats & Débriefing","Restitution des scores certifiants · Feedback individuel formateur · Attestation FIM remise"),
        ("11h30","12h30","Quiz 40 questions","Évaluation finale des connaissances théoriques · Seuil 28/40 · Corrigé collectif"),
        ("13h30","14h30","Plan de prise de poste","Construction du plan J+0 à J+30 : objectifs semaine 1 à 4 · Indicateurs de réussite"),
        ("14h30","15h30","Cérémonie de clôture","Remise officielle des certifications · Engagement public · Photo de groupe · Votre premier appel demain"),
    ]
    y = 0.88
    for st,en,ti,de in prg:
        R(s, 0.3, y, 1.38, 0.68, fill=TEAL_D)
        T(s, f"{st}\n{en}", 0.32, y+0.08, 1.34, 0.54, sz=11, bold=True, col=WHITE, align=PP_ALIGN.CENTER)
        R(s, 1.7, y, 3.0, 0.68, fill=TEAL)
        T(s, ti, 1.8, y+0.15, 2.82, 0.42, sz=12, bold=True, col=WHITE)
        R(s, 4.72, y, 8.31, 0.68, fill=WHITE, line=TEAL, lw=Pt(0.8))
        T(s, de, 4.84, y+0.07, 8.1, 0.58, sz=9, col=DARK, wrap=True)
        y += 0.71

    s = p.slides.add_slide(SL(p))
    H_CONTENT(s, "Analyse des Écarts — Recommandations Pédagogiques UNICEF Paris")
    TWO_COL(s,
        "Écarts identifiés vs standards UNICEF Paris",
        ["Durée de formation : FIM = 5J · Standard UNICEF Paris = 7J (J6+J7 ajoutés)",
         "Simulation terrain réelle : manquait avant J6 — intégrée ici",
         "Quiz théorique : absent de V5 · 40 questions intégrées en J5 & J7",
         "Feedback individuel structuré : informel avant · grille 24 critères désormais",
         "Évaluateur externe en certification : absent · intégré en J7"],
        "Recommandations intégrées V6",
        ["2 journées d'ateliers pratiques ajoutées (J6 + J7)",
         "Grille d'observation live utilisée dès J6",
         "Module quiz dédié (40 questions, 6 parties)",
         "Plan de prise de poste J+0 à J+30 formalisé",
         "Slide 'Merci à tous' ritualisant la clôture de chaque module",
         "Logo SHY-Performance sur chaque entête de slide"])
    HL(s,
       "Ces recommandations s'appuient sur les standards de formation des fundraisers téléphoniques\n"
       "UNICEF Paris (protocole FIDELIS 2025) et les bonnes pratiques du secteur en France.",
       0.3, 6.05, 12.73, 0.58)

    s = p.slides.add_slide(SL(p))
    H_CONTENT(s, "Votre Plan de Prise de Poste — J+0 à J+30")
    weeks = [
        ("J+0\nDemain", TEAL_D,
         "Premier appel de production\nObjectif : 9 CU/H dès J1\nQualification FIDELIS conforme"),
        ("SEMAINE 1\nJ+7", TEAL,
         "TX Transfo ≥ cible FIDELIS\nAucune disqualification nomenclature\nFeedback formateur J+7"),
        ("SEMAINE 2\nJ+14", TEAL_M,
         "0 objection Cat.5 non gérée\nPremier PA IBAN à chaud validé\nPoint superviseur FIDELIS"),
        ("SEMAINE 3\nJ+21", TEAL_A,
         "Autonomie complète sur le script\nPDC mensuel en bonne voie\nAutoévaluation grille 24 crit."),
        ("SEMAINE 4\nJ+30", RGBColor(0x33,0x77,0x77),
         "Bilan formateur SHY-Performance\nRévision objectifs M2\nCertification terrain confirmée"),
    ]
    x = 0.3
    for tag, col, body in weeks:
        R(s, x, 0.88, 2.46, 0.55, fill=col)
        T(s, tag, x+0.1, 0.9, 2.26, 0.51, sz=12, bold=True, col=WHITE, align=PP_ALIGN.CENTER)
        R(s, x, 1.43, 2.46, 4.5, fill=WHITE, line=col, lw=Pt(1.5))
        T(s, body, x+0.14, 1.52, 2.18, 4.2, sz=11, col=DARK, wrap=True)
        x += 2.56
    HL(s,
       "Votre formateur SHY-Performance vous contacte à J+15 et J+30. "
       "À J+90 : analyse KPI terrain consolidée transmise à FIDELIS × UNICEF France.",
       0.3, 6.1, 12.73, 0.44)

    MERCI(p, "Atelier Pratique Jour 7 — Certification FIM")
    return save(p, f"FIM_BookJ7_AtelierPratique2_{VER}.pptx")


# ══════════════════════════════════════════════════════════════════════════
#  7 MODULES DE BASE (avec MERCI final sur chacun)
# ══════════════════════════════════════════════════════════════════════════

def build_book0():
    p2 = prs()
    s = p2.slides.add_slide(SL(p2))
    H_TITLE(s,"Bienvenue dans la Formation FIM",
        "Fundraiser Impact Mission — UNICEF France × SHY-Performance × FIDELIS",
        tag="BOOK 0  ·  Fil Conducteur  ·  Document de parcours")
    T(s,"Ce livret est votre guide pour les 5 jours de formation.\n"
        "Il contient votre planning, vos engagements et votre tableau de bord personnel.",
        0.5,2.75,11.8,0.9,sz=14,col=DARK)
    T(s,"Vous allez maîtriser le script officiel UNICEF en 7 étapes, qualifier vos appels\n"
        "selon la nomenclature FIDELIS et atteindre l'objectif de 9 CU/H minimum.",
        0.5,3.75,11.8,0.9,sz=13,bold=True,col=TEAL_D)

    s = p2.slides.add_slide(SL(p2))
    H_CONTENT(s,"Vos 5 Journées + 2 Ateliers — Ce qui vous attend")
    days=[
        ("J1","Fondations\n& Mission","Mission UNICEF · KPI FIDELIS · Nomenclature DON/INDÉCIS/REFUS · Organisation production"),
        ("J2","Monde\nAssociatif","Loi 1901 · Histoire humanitaire · 3 principes · Don régulier · Label Don en Confiance"),
        ("J3","Analyse &\nScript","Script 7 étapes · 5 règles accroche · 4 modes validation · Lecture à voix haute"),
        ("J4","Traitement\nObjections","15 objections · 5 catégories · Méthode AAR · Jeux de rôle · Grille formateur"),
        ("J5","Synthèse &\nCertif.","Simulations · Grille 24 critères · Quiz 40 questions · Remise certifications"),
        ("J6","Atelier\nPratique 1","Appels simulés live · Grille formateur · Feedback immédiat · Cas complexes"),
        ("J7","Atelier\nPratique 2","Perfectionnement · Certification finale · Plan prise de poste J+0→J+30"),
    ]
    x=0.3
    for tag,title,desc in days:
        R(s,x,0.88,1.82,0.34,fill=TEAL_D)
        T(s,tag,x+0.06,0.91,1.7,0.26,sz=9,bold=True,col=WHITE)
        R(s,x,1.22,1.82,0.9,fill=TEAL)
        T(s,title,x+0.08,1.25,1.66,0.84,sz=10,bold=True,col=WHITE,align=PP_ALIGN.CENTER)
        R(s,x,2.12,1.82,3.75,fill=WHITE,line=TEAL,lw=Pt(1))
        T(s,desc,x+0.1,2.18,1.62,3.5,sz=8,col=DARK,wrap=True)
        x+=1.86
    HL(s,"35h formation + 2 ateliers pratiques · 9h30–17h30 · 6h40 production effective · Certification FIDELIS × UNICEF France",
        0.3,6.05,12.73,0.38)

    s = p2.slides.add_slide(SL(p2))
    H_CONTENT(s,"Votre Contrat d'Engagement")
    TWO_COL(s,"Vos engagements",
        ["Participer activement à toutes les mises en situation",
         "Compléter votre livret personnel chaque soir (5 min)",
         "Demander un feedback précis après chaque exercice",
         "Respecter la confidentialité des échanges du groupe",
         "Arriver à l'heure — le groupe démarre ensemble"],
        "Ce que SHY-Performance s'engage à vous offrir",
        ["Un feedback individuel personnalisé chaque jour",
         "Des exercices pratiques progressifs et bienveillants",
         "Un formateur qui adapte le rythme à votre profil",
         "Un suivi à J+15 et J+30 après la formation",
         "Une certification reconnue FIDELIS × UNICEF France"])
    HL(s,"Mantra de la formation : « Le refus d'aujourd'hui peut être le don de demain »",0.3,6.02,12.73,0.44)

    MERCI(p2,"Book 0 — Fil Conducteur · Formation FIM")
    return save(p2, f"FIM_Book0_FilConducteur_{VER}.pptx")


def build_book1():
    p2=prs()
    s=p2.slides.add_slide(SL(p2))
    H_TITLE(s,"Vos Objectifs de Formation",
        "Ce que vous saurez faire à l'issue des 5 jours + 2 ateliers FIM",
        tag="BOOK 1  ·  Synopsis & Objectifs  ·  SHY-Performance × FIDELIS × UNICEF France")
    T(s,"À l'issue de ces 7 jours, vous serez capable de conduire un appel de collecte de dons\n"
        "au nom de l'UNICEF, de la phrase d'accroche jusqu'à la validation du Prélèvement Automatique.",
        0.5,2.75,11.8,0.9,sz=14,col=DARK)

    s=p2.slides.add_slide(SL(p2))
    H_CONTENT(s,"Vos 3 Dimensions de Compétence — SAVOIR / SAVOIR-FAIRE / SAVOIR-ÊTRE")
    cols=[
        ("SAVOIR","Connaissance de la cause",TEAL_D,
         ["Présenter l'UNICEF, sa mission, son histoire (75 ans, 190 pays)",
          "Expliquer la malnutrition aigüe sévère et le rôle des sachets RUTF (92g)",
          "Situer le secteur associatif (Loi 1901, Comité de la Charte, Bloctel, RGPD)",
          "Citer les arguments de légitimité UNICEF",
          "Connaître les 3 principes humanitaires fondamentaux",
          "Comprendre Conquête vs Fidélisation/Réactivation"]),
        ("SAVOIR-FAIRE","Technique d'appel & qualification",TEAL,
         ["Appliquer les 5 règles officielles de l'accroche téléphonique",
          "Conduire le script officiel UNICEF en 7 étapes",
          "Qualifier chaque appel selon la nomenclature FIDELIS (DON/INDÉCIS/REFUS)",
          "Gérer les 4 modes de validation du don",
          "Traiter les 15 objections types par la méthode AAR",
          "Atteindre l'objectif de 9 CU/H minimum en production"]),
        ("SAVOIR-ÊTRE","Posture & conviction",TEAL_M,
         ["Maintenir un ton chaleureux, souriant et naturel tout au long de l'appel",
          "Gérer les objections sensibles avec calme et professionnalisme",
          "Transmettre une conviction sincère pour la mission UNICEF",
          "Intégrer le mantra : «Le refus d'aujourd'hui peut être le don de demain»",
          "Adopter une posture de représentant digne de la cause humanitaire",
          "Respecter systématiquement les règles éthiques du secteur"]),
    ]
    x=0.3
    for label,sub,col,items in cols:
        R(s,x,0.88,4.18,0.52,fill=col)
        T(s,label,x+0.12,0.9,4.0,0.3,sz=14,bold=True,col=WHITE)
        T(s,sub,x+0.12,1.2,4.0,0.22,sz=9,col=RGBColor(0xCC,0xEE,0xEE))
        iy=1.44
        for i,item in enumerate(items):
            bg=WHITE if i%2==0 else TEAL_L
            R(s,x,iy,4.18,0.5,fill=bg,line=col,lw=Pt(0.4))
            T(s,"• "+item,x+0.12,iy+0.06,3.96,0.4,sz=9,col=DARK)
            iy+=0.52
        x+=4.35

    s=p2.slides.add_slide(SL(p2))
    H_CONTENT(s,"Indicateurs de Performance — Nomenclature FIDELIS")
    kpis=[
        ("CU/H","Contacts Utiles / Heure","Nombre de contacts qualifiés par heure de production","≥ 9 CU/H"),
        ("TX Transfo","Taux de Transformation","% de PEL/PA obtenus parmi les Contacts Utiles","Cible FIDELIS"),
        ("PDC","Plan de Charge","Volume mensuel de dons réguliers FIDELIS × UNICEF France","Volume défini"),
        ("PEL","Promesse En Ligne","Contact ayant accepté de valider son don en ligne","Comptabilisé TX"),
        ("PA","Prélèvement Automatique","Don régulier sécurisé par mandat SEPA","Produit unique"),
        ("Qualif.","Conformité des qualifications","Qualifications correctes DON/INDÉCIS/REFUS","≥ 50% qualité"),
    ]
    cxs=[0.3,1.35,4.65,10.38]; cws=[1.02,3.27,5.7,2.65]
    TH(s,["Code","Libellé complet","Définition opérationnelle","Objectif"],cxs,cws,y=0.88)
    y=1.2
    for i,(code,lib,defin,obj) in enumerate(kpis):
        bg=WHITE if i%2==0 else TEAL_L; rh=0.65
        R(s,cxs[0],y,cws[0],rh,fill=TEAL_D)
        T(s,code,cxs[0]+0.06,y+0.1,cws[0]-0.1,rh-0.15,sz=9,bold=True,col=WHITE,align=PP_ALIGN.CENTER)
        R(s,cxs[1],y,cws[1],rh,fill=bg)
        T(s,lib,cxs[1]+0.1,y+0.1,cws[1]-0.15,rh-0.15,sz=10,bold=True,col=TEAL_D)
        R(s,cxs[2],y,cws[2],rh,fill=bg)
        T(s,defin,cxs[2]+0.1,y+0.08,cws[2]-0.15,rh-0.1,sz=9,col=DARK,italic=True)
        R(s,cxs[3],y,cws[3],rh,fill=TEAL)
        T(s,obj,cxs[3]+0.1,y+0.15,cws[3]-0.15,rh-0.22,sz=10,bold=True,col=WHITE,align=PP_ALIGN.CENTER)
        y+=rh+0.02

    MERCI(p2,"Book 1 — Synopsis & Objectifs · Formation FIM")
    return save(p2, f"FIM_Book1_Synopsis_Objectifs_{VER}.pptx")


# ── Books J1–J5 : version compacte avec MERCI final ──────────────────────

def _j_base(title, sub, tag, content_fn):
    p2=prs()
    s=p2.slides.add_slide(SL(p2)); H_TITLE(s,title,sub,tag=tag)
    content_fn(p2)
    MERCI(p2, f"{tag}")
    return p2

def build_bookJ1():
    def fn(p2):
        s=p2.slides.add_slide(SL(p2))
        H_CONTENT(s,"Journée 1 — Ce que vous allez vivre aujourd'hui")
        SEQ4(s,["Témoignage bénéficiaire UNICEF.\nNotez ce que vous ressentez.",
                "En groupe : 'Qu'est-ce qui vous\na le plus touché ?'",
                "Mission UNICEF · KPI FIDELIS\nNomenclature · Organisation journée",
                "Pitch mission 60 secondes\nà votre binôme — sans notes."],
             ["IMMERSION","ÉCHANGE","APPORT","MISE EN PRATIQUE"],y=0.88,h=5.1)
        KPI_BAR(s,["Pitch mission 60s validé","Quiz 5Q UNICEF (seuil 60%)","3 lignes LFI"],y=6.12)

        s=p2.slides.add_slide(SL(p2))
        H_CONTENT(s,"L'UNICEF & La Cause — Vos Arguments de Conviction")
        facts=[("1946","Fondation UNICEF\n(ONU)"),("190","pays présents"),
               ("45%","enfants du monde\nvaccinés par l'UNICEF"),("92g","poids sachet RUTF\nmalnutrition"),
               ("3 princ.","Humanité\nImpartialité\nNeutralité"),("9 CU/H","votre objectif\nquotidien minimum")]
        x=0.3
        for val,label in facts:
            R(s,x,0.88,2.1,1.05,fill=TEAL_D)
            T(s,val,x+0.1,0.9,1.9,0.62,sz=20,bold=True,col=WHITE,align=PP_ALIGN.CENTER)
            T(s,label,x+0.1,1.5,1.9,0.4,sz=9,col=RGBColor(0xCC,0xEE,0xEE),align=PP_ALIGN.CENTER,wrap=True)
            x+=2.16
        HL(s,"La cause centrale : Malnutrition aigüe sévère — 1 enfant meurt de faim toutes les 11 secondes.\n"
            "Les sachets RUTF (92g) permettent à un enfant malnutri de guérir en 6 à 8 semaines.",
            0.3,2.1,12.73,0.72)
        HL(s,"Votre posture vocale : Sourire audible dès le mot 'Bonjour' · Débit posé · Ton chaleureux · "
            "Conviction sincère — vous croyez en la mission UNICEF que vous représentez.",
            0.3,2.95,12.73,0.62)
        HL(s,"Exercice : Lisez votre script à voix haute en souriant physiquement. "
            "Enregistrez-vous 60 sec. Réécoutez. Qu'est-ce que vous changeriez ?",
            0.3,3.72,12.73,0.44)
        kpis_j1=[("CU/H ≥ 9","Contacts Utiles/h"),("TX Transfo","% PEL+PA parmi CU"),
                 ("PDC","Plan de Charge mensuel"),("DON","Contact qui accepte le PA"),
                 ("INDÉCIS","Contact à relancer"),("REFUS","Contact qui refuse définitivement")]
        x=0.3
        for val,label in kpis_j1:
            R(s,x,4.35,2.1,1.05,fill=TEAL)
            T(s,val,x+0.08,4.37,1.94,0.55,sz=13,bold=True,col=WHITE,align=PP_ALIGN.CENTER)
            T(s,label,x+0.08,4.9,1.94,0.46,sz=9,col=RGBColor(0xCC,0xEE,0xEE),align=PP_ALIGN.CENTER,wrap=True)
            x+=2.16
        HL(s,"Organisation journée : 9h30–17h30 · 6h40 production effective · Conquête vs Fidélisation/Réactivation",
            0.3,5.55,12.73,0.38)
    p2=_j_base("Fondations & Mission",
        "Votre identité de fundraiser · La cause · Votre raison d'agir",
        "BOOK J1  ·  Journée 1  ·  UNICEF France × SHY-Performance",fn)
    return save(p2,f"FIM_BookJ1_Fondations_Mission_{VER}.pptx")


def build_bookJ2():
    def fn(p2):
        s=p2.slides.add_slide(SL(p2))
        H_CONTENT(s,"Journée 2 — Ce que vous allez vivre aujourd'hui")
        SEQ4(s,["Analyse de 2 campagnes téléphoniques :\n1 réussie, 1 en difficulté.",
                "'Qu'est-ce qui a fait la différence ?\nPourquoi l'une a converti ?'",
                "Loi 1901 · Histoire humanitaire\nDon régulier · Don en Confiance",
                "'Expliquez le don régulier à un\nprospect sceptique en 90 secondes.'"],
             ["DÉCOUVERTE","ANALYSE","APPORT","MISE EN PRATIQUE"],y=0.88,h=5.1)
        KPI_BAR(s,["Restitution orale étude de cas (5 min)","Fiche secteur associatif complétée","Réflexion LFI : légitimité du PA"],y=6.12)

        s=p2.slides.add_slide(SL(p2))
        H_CONTENT(s,"La France Associative & L'Histoire de l'Humanitaire")
        facts=[("1,5M","associations en France"),("22M","bénévoles en France"),
               ("1er juil\n1901","Loi droit associatif"),("3 types","De fait / Déclarée /\nUtilité publique"),
               ("Don en\nConfiance","Label indépendant\nde contrôle"),("RGPD\nBloctel","Cadre légal données\nprospects")]
        x=0.3
        for val,label in facts:
            R(s,x,0.88,2.1,1.05,fill=TEAL_D)
            T(s,val,x+0.1,0.9,1.9,0.6,sz=15,bold=True,col=WHITE,align=PP_ALIGN.CENTER)
            T(s,label,x+0.1,1.5,1.9,0.4,sz=9,col=RGBColor(0xCC,0xEE,0xEE),align=PP_ALIGN.CENTER,wrap=True)
            x+=2.16
        timeline=[("1859","Bataille de Solférino\n→ Naissance Croix-Rouge"),("1863","Fondation Croix-Rouge\npar Henry Dunant"),
                  ("1946","Fondation UNICEF\n(ONU après WWII)"),("1989","Convention Droits\nde l'Enfant (ONU)"),
                  ("1989","Comité Don\nen Confiance"),("2000s","RGPD & Bloctel\nencadrement légal")]
        x=0.3
        for year,event in timeline:
            R(s,x,2.1,2.1,0.36,fill=TEAL)
            T(s,year,x+0.1,2.13,1.9,0.26,sz=11,bold=True,col=WHITE,align=PP_ALIGN.CENTER)
            R(s,x,2.46,2.1,1.2,fill=TEAL_L,line=TEAL,lw=Pt(0.8))
            T(s,event,x+0.12,2.52,1.86,1.08,sz=10,col=DARK,wrap=True)
            x+=2.16
        HL(s,"3 principes humanitaires que vous incarnez : HUMANITÉ · IMPARTIALITÉ · NEUTRALITÉ",0.3,3.82,12.73,0.38)
        HL(s,"Don régulier PA : valeur à vie × 8 vs don ponctuel — colonne vertébrale des programmes UNICEF France.",0.3,4.32,12.73,0.38)
        HL(s,"Exercice : Expliquez en 90 secondes pourquoi le PA est préférable au don ponctuel. "
            "Votre binôme joue le prospect sceptique.",0.3,4.84,12.73,0.44)
    p2=_j_base("Le Monde Associatif & Humanitaire",
        "Comprendre le secteur pour représenter la cause avec légitimité",
        "BOOK J2  ·  Journée 2  ·  France associative & Histoire humanitaire",fn)
    return save(p2,f"FIM_BookJ2_MondeAssociatif_{VER}.pptx")


def build_bookJ3():
    def fn(p2):
        s=p2.slides.add_slide(SL(p2))
        H_CONTENT(s,"Journée 3 — Ce que vous allez vivre aujourd'hui")
        SEQ4(s,["Lecture intégrale du script\nofficiel UNICEF à voix haute.\nPremière impression.",
                "Analyse phrase par phrase :\n'Quelle est l'intention ici ?\nPourquoi ce mot précis ?'",
                "7 étapes · 5 règles accroche\n4 modes de validation du don\nTon, rythme, intention",
                "Lecture en binômes. Débriefs\ncollectifs. Points forts\n& axes d'amélioration."],
             ["DÉCOUVERTE","ANALYSE","APPORT","ENTRAÎNEMENT"],y=0.88,h=5.1)
        KPI_BAR(s,["Script lu 3 fois sans hésitation","5 règles d'accroche récitées","Score seuil 12/20 pour J4"],y=6.12)

        s=p2.slides.add_slide(SL(p2))
        H_CONTENT(s,"Le Script Officiel UNICEF — Les 7 Étapes")
        steps=[("①\nACCROCHE",TEAL_D,"Ouverture conforme aux 5 règles\nTon chaleureux, prénom+nom,\norganisation nommée, raison annoncée"),
               ("②\nMALNUTRITION",TEAL,"Présentation de la cause\n1 enfant/11sec · Dimension\nhumaine et émotionnelle"),
               ("③\nSACHETS RUTF",TEAL_M,"La solution UNICEF : sachets 92g\n6 à 8 semaines pour guérir\nun enfant sévèrement malnutri"),
               ("④\nAPPEL AU SOUTIEN",TEAL_A,"Proposition du PA mensuel\nMontant suggéré · Impact\nchiffré · Simplicité du geste"),
               ("⑤\nSI DON PONCTUEL",RGBColor(0x33,0x77,0x77),"Réorientation vers le PA\n'Et si on envisageait quelque\nchose de régulier ?'"),
               ("⑥\nINDÉCIS",RGBColor(0x22,0x66,0x66),"Gestion de l'hésitation\nRester dans l'échange\nRelance avec question ouverte"),
               ("⑦\nVALIDATION",RGBColor(0x11,0x55,0x55),"IBAN à chaud / PA en ligne /\nPEL / PA courrier\nConfirmer coordonnées · SMS")]
        x=0.3; sw=(13.33-0.6)/7
        for title,col,body in steps:
            R(s,x,0.88,sw,0.48,fill=col)
            T(s,title,x+0.06,0.9,sw-0.1,0.44,sz=8,bold=True,col=WHITE,align=PP_ALIGN.CENTER)
            R(s,x,1.36,sw,4.85,fill=WHITE,line=col,lw=Pt(1.2))
            T(s,body,x+0.08,1.42,sw-0.14,4.65,sz=8.5,col=DARK,wrap=True)
            x+=sw+0.02
        HL(s,"Durée standard : 3 à 5 min · Chaque étape a une fonction précise — ne pas sauter, ne pas inverser.",0.3,6.34,12.73,0.38)

        s=p2.slides.add_slide(SL(p2))
        H_CONTENT(s,"Les 5 Règles Officielles de la Phrase d'Accroche")
        T(s,"Ces 5 règles sont non négociables — elles définissent un appel conforme FIDELIS :",
           0.35,0.88,12.5,0.32,sz=12,bold=True,col=TEAL_D)
        rules=[("RÈGLE 1","Se présenter avec prénom et nom","'Bonjour, je m'appelle [Prénom NOM]…' — identité claire dès les premières secondes."),
               ("RÈGLE 2","Nommer l'organisation représentée","'…je vous appelle au nom de l'UNICEF France…' — le prospect sait immédiatement qui appelle."),
               ("RÈGLE 3","Annoncer la raison de l'appel","'…pour vous parler d'une initiative importante pour les enfants…' — clarté totale sur l'objet de l'appel."),
               ("RÈGLE 4","Vérifier la disponibilité du prospect","'Est-ce que vous avez quelques minutes ?' — respect de l'interlocuteur."),
               ("RÈGLE 5","Adopter un ton chaleureux et naturel","Sourire audible · Débit posé · Articulation claire — le ton est aussi important que les mots.")]
        y=1.28
        for tag,title,body in rules:
            bg=WHITE if rules.index((tag,title,body))%2==0 else TEAL_L
            R(s,0.3,y,1.1,0.88,fill=TEAL_D)
            T(s,tag,0.32,y+0.18,1.06,0.52,sz=8,bold=True,col=WHITE,align=PP_ALIGN.CENTER)
            R(s,1.42,y,3.4,0.88,fill=TEAL_L,line=TEAL_D,lw=Pt(0.8))
            T(s,title,1.52,y+0.2,3.22,0.52,sz=10,bold=True,col=TEAL_D)
            R(s,4.84,y,8.19,0.88,fill=bg)
            T(s,body,4.96,y+0.1,7.98,0.7,sz=10,col=DARK,italic=True)
            y+=0.92
        HL(s,"Un appel qui viole une de ces 5 règles est disqualifié dans la nomenclature FIDELIS.",
            0.3,5.88,12.73,0.44)
    p2=_j_base("Analyse & Script Officiel UNICEF",
        "Les 7 étapes · Les 5 règles d'accroche · Les 4 modes de validation",
        "BOOK J3  ·  Journée 3  ·  Script UNICEF × FIDELIS",fn)
    return save(p2,f"FIM_BookJ3_Lecture_Script_{VER}.pptx")


def build_bookJ4():
    def fn(p2):
        s=p2.slides.add_slide(SL(p2))
        H_CONTENT(s,"Journée 4 — Ce que vous allez vivre aujourd'hui")
        SEQ4(s,["Écoute de 5 extraits d'appels :\nbonne et mauvaise gestion\nd'une objection.",
                "'Qu'avez-vous entendu ?\nQu'est-ce qui a tout changé\ndans la réponse ?'",
                "15 objections · 5 catégories\nMéthode AAR\nAccuser réception/Argumenter/Relancer",
                "Jeux de rôle en binômes\nsur chaque catégorie.\nGrille d'observation formateur."],
             ["ÉCOUTE","ANALYSE","APPORT","ENTRAÎNEMENT"],y=0.88,h=5.1)
        KPI_BAR(s,["Grille écoute active validée formateur","LFI : 'Quelle objection me challenge le plus ?'","Score seuil 14/20 pour validation J4"],y=6.12)

        s=p2.slides.add_slide(SL(p2))
        H_CONTENT(s,"Votre Méthode AAR — Accuser Réception · Argumenter · Relancer")
        aar=[("A\nACCUSER\nRÉCEPTION",TEAL_D,
              "Montrez que vous avez entendu.\nNe justifiez pas immédiatement.\n\n"
              "'Je comprends tout à fait…'\n'C'est une question que beaucoup\nse posent et c'est légitime…'\n\n"
              "⚠ Pas de 'Oui mais…' ou\n'Non vous avez tort…'"),
             ("A\nARGUMENTER",TEAL,
              "1 argument + 1 preuve + 1 chiffre.\n\n"
              "Règles d'or :\n→ 1 seul argument par réponse\n→ Preuve vérifiable (rapport annuel)\n→ Chiffre simple et mémorisable\n\n"
              "Ex : '10€/mois = 33 centimes/jour\n= 1 enfant traité en 6 semaines'"),
             ("R\nRELANCER",TEAL_M,
              "Revenez à la proposition.\nPas de pression.\n\n"
              "'Est-ce que cela répond à\nvotre question ?'\n'On peut aller plus loin ?'\n'Qu'est-ce qui vous permettrait\nde vous sentir à l'aise ?'\n\n"
              "⚠ Laissez 3 secondes de silence\navant de relancer.")]
        x=0.3
        for title,col,body in aar:
            R(s,x,0.88,4.17,0.68,fill=col)
            T(s,title,x+0.14,0.9,3.93,0.64,sz=15,bold=True,col=WHITE,align=PP_ALIGN.CENTER)
            R(s,x,1.56,4.17,4.68,fill=WHITE,line=col,lw=Pt(1.8))
            T(s,body,x+0.16,1.64,3.85,4.46,sz=10,col=DARK,wrap=True)
            x+=4.3
        HL(s,"Le silence après votre argument est votre allié. Comptez 1–2–3 mentalement avant de relancer.",
            0.3,6.34,12.73,0.44)

        s=p2.slides.add_slide(SL(p2))
        H_CONTENT(s,"Les 5 Catégories d'Objections — 15 Situations à Maîtriser")
        cats=[("CAT. 1\nOBJECTIONS INITIALES",TEAL_D,
               "Phase d'accroche — 8 objections :\n• Pas intéressé\n• Faux numéro\n• Appel pour un don\n"
               "• Je donne déjà\n• Arnaque\n• Je ne donne pas par téléphone\n"
               "• Pas confiance aux associations\n• N'aime pas être contacté"),
              ("CAT. 2\nFINANCIÈRES",TEAL,
               "Après l'appel au don — 3 objections :\n• Je donne déjà ailleurs\n"
               "• Je n'ai pas les moyens\n• Dernier positionnement financier"),
              ("CAT. 3\nCONTRE LE PA",TEAL_M,
               "Sur le prélèvement — 3 objections :\n• Préfère le don ponctuel\n"
               "• N'aime pas l'engagement mensuel\n• Dernier positionnement PA"),
              ("CAT. 4\nMÉFIANCE IBAN",TEAL_A,
               "Sur la sécurité bancaire :\n• Protocole SEPA expliqué\n"
               "• Droits du donateur (arrêt 1 clic)\n• Réassurance données bancaires"),
              ("CAT. 5\nCONFLICTUELLES",RGBColor(0x33,0x77,0x77),
               "Objections sensibles :\n• Origine de l'appel\n• Source du numéro\n"
               "• RGPD / Droits légaux\n• Bloctel / liste rouge\n"
               "• Accent · Identité de l'appelant")]
        x=0.3; cw=(13.33-0.6)/5
        for title,col,body in cats:
            R(s,x,0.88,cw,0.52,fill=col)
            T(s,title,x+0.1,0.9,cw-0.16,0.48,sz=8,bold=True,col=WHITE,align=PP_ALIGN.CENTER)
            R(s,x,1.4,cw,4.85,fill=WHITE,line=col,lw=Pt(1.2))
            T(s,body,x+0.1,1.47,cw-0.16,4.65,sz=9,col=DARK,wrap=True)
            x+=cw+0.02
        HL(s,"La majorité des appels se termine à Cat.1. Maîtriser ces 8 objections = atteindre vos 9 CU/H.",
            0.3,6.34,12.73,0.44)
    p2=_j_base("Traitement des Objections",
        "15 objections · 5 catégories · Méthode AAR · Jeux de rôle téléphonique",
        "BOOK J4  ·  Journée 4  ·  Objections UNICEF France × FIDELIS",fn)
    return save(p2,f"FIM_BookJ4_Objections_{VER}.pptx")


def build_bookJ5():
    def fn(p2):
        s=p2.slides.add_slide(SL(p2))
        H_CONTENT(s,"Programme Journée 5 — Certification FIM")
        prg=[("08h30","09h00","Révision consolidée","Révision des 5 modules FIM · Quiz flash · Météo émotionnelle"),
             ("09h00","11h00","Simulations d'appels complets","Script + objections + closing chronométrés · Jeux de rôle fundraiser vs prospect"),
             ("11h00","11h30","Débriefing collectif","Retour simulations · Célébration progrès · 3 apprentissages collectifs de la semaine"),
             ("11h30","12h30","Grille 24 critères","Évaluation individuelle · Entretien formateur 5 min · Feedback personnalisé"),
             ("13h30","14h30","Quiz 40 questions","Évaluation finale des connaissances · Seuil 28/40 · Corrigé collectif"),
             ("14h30","15h00","Remise certifications","Attestations officielles FIM · Plan d'action terrain · 1er appel de production demain")]
        y=0.88
        for st,en,ti,de in prg:
            R(s,0.3,y,1.38,0.7,fill=TEAL_D)
            T(s,f"{st}\n{en}",0.32,y+0.08,1.34,0.56,sz=11,bold=True,col=WHITE,align=PP_ALIGN.CENTER)
            R(s,1.7,y,3.0,0.7,fill=TEAL)
            T(s,ti,1.8,y+0.15,2.82,0.42,sz=12,bold=True,col=WHITE)
            R(s,4.72,y,8.31,0.7,fill=WHITE,line=TEAL,lw=Pt(0.8))
            T(s,de,4.84,y+0.07,8.1,0.6,sz=9,col=DARK,wrap=True)
            y+=0.73

        s=p2.slides.add_slide(SL(p2))
        H_CONTENT(s,"Grille de Certification — 24 Critères")
        T(s,"Voici exactement ce sur quoi vous serez évalué(e) — aucune surprise :",
           0.35,0.85,12.5,0.3,sz=11,bold=True,col=TEAL_D)
        criteria=[
            ("A1","Respect des 5 règles d'accroche","Accroche"),
            ("A2","Présentation claire de l'organisation","Accroche"),
            ("A3","Ton chaleureux et naturel dès l'ouverture","Accroche"),
            ("A4","Présentation de la cause malnutrition","Script"),
            ("A5","Explication des sachets RUTF avec précision","Script"),
            ("A6","Appel au soutien clair avec montant suggéré","Script"),
            ("A7","Gestion don ponctuel → réorientation PA","Script"),
            ("A8","Gestion indécis — relance constructive","Script"),
            ("A9","Validation coordonnées — mode paiement","Closing"),
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
            ("D1","Qualification conforme FIDELIS (DON/IND/REF)","FIDELIS"),
            ("D2","Nomination correcte du mode de validation","FIDELIS"),
            ("D3","Durée d'appel dans la norme (3–5 min)","FIDELIS"),
            ("D4","Saisie coordonnées complète et exacte","FIDELIS"),
        ]
        cxs=[0.3,0.76,7.0,9.9,10.7,11.5,12.3]
        cws=[0.43,6.22,2.87,0.78,0.78,0.78,0.75]
        TH(s,["#","Ce que vous devez démontrer","Catégorie","1","2","3","4"],cxs,cws,y=1.2)
        y=1.52; rh=0.34
        cat_c={"Accroche":TEAL_D,"Script":TEAL,"Closing":TEAL_M,"Objections":TEAL_A,
               "Posture":RGBColor(0x33,0x77,0x77),"Éthique":RGBColor(0x22,0x66,0x66),"FIDELIS":RGBColor(0x11,0x55,0x55)}
        for i,(code,crit,cat) in enumerate(criteria):
            bg=WHITE if i%2==0 else TEAL_L
            R(s,cxs[0],y,cws[0],rh,fill=TEAL_D)
            T(s,code,cxs[0]+0.04,y+0.06,cws[0]-0.06,rh-0.1,sz=8,bold=True,col=WHITE,align=PP_ALIGN.CENTER)
            R(s,cxs[1],y,cws[1],rh,fill=bg)
            T(s,crit,cxs[1]+0.08,y+0.06,cws[1]-0.12,rh-0.08,sz=9,col=DARK)
            cc=cat_c.get(cat,TEAL)
            R(s,cxs[2],y,cws[2],rh,fill=cc)
            T(s,cat,cxs[2]+0.06,y+0.06,cws[2]-0.1,rh-0.08,sz=8,bold=True,col=WHITE,align=PP_ALIGN.CENTER)
            for x,w in zip(cxs[3:],cws[3:]):
                R(s,x,y,w,rh,fill=WHITE,line=TEAL,lw=Pt(0.4))
            y+=rh+0.01
        R(s,cxs[0],y,sum(cws[:3])+0.04,0.32,fill=TEAL_D)
        T(s,"TOTAL  /24",cxs[0]+0.12,y+0.06,9.2,0.22,sz=10,bold=True,col=WHITE)
        R(s,cxs[3],y,sum(cws[3:])+0.04,0.32,fill=TEAL_L,line=TEAL_D,lw=Pt(2))
        T(s,"__ / 24",cxs[3]+0.1,y+0.06,2.4,0.22,sz=11,bold=True,col=TEAL_D,align=PP_ALIGN.CENTER)
        y+=0.36
        HL(s,"✅ Certifié(e) FIM : ≥ 17/24  ·  ⚠ Formation complémentaire : 12–16/24  ·  🔄 Rattrapage J+15 : < 12/24",
            0.3,y+0.04,12.73,0.36,fill=TEAL_D)

        s=p2.slides.add_slide(SL(p2))
        H_CONTENT(s,"Votre Plan d'Action — 4 Semaines Terrain")
        TWO_COL(s,"Quiz final — 40 questions réparties en 6 parties",
            ["Partie A (Q1–Q8) : Connaissance UNICEF & de la cause",
             "Partie B (Q9–Q15) : Secteur associatif & humanitaire",
             "Partie C (Q16–Q22) : Script officiel & règles d'accroche",
             "Partie D (Q23–Q28) : KPI FIDELIS, nomenclature & production",
             "Partie E (Q29–Q35) : Traitement des objections & méthode AAR",
             "Partie F (Q36–Q40) : Éthique, posture & vocabulaire métier"],
            "Votre plan d'action — 4 semaines terrain",
            ["S1 (J+7) : 9 CU/H atteint et stable",
             "S2 (J+14) : TX Transfo ≥ cible FIDELIS première mesure",
             "S3 (J+21) : 0 disqualification nomenclature FIDELIS",
             "S4 (J+30) : Bilan formateur SHY-Performance + objectifs M2",
             "J+15 : Observation supervisée SHY-Performance",
             "J+90 : Analyse KPI terrain FIDELIS × UNICEF France"])
        HL(s,"Bienvenue dans la famille SHY-Performance × FIDELIS × UNICEF France — Certifié(e) FIM.",
            0.3,6.02,12.73,0.44)
    p2=_j_base("Synthèse Générale & Simulations",
        "Certification FIM · Grille 24 critères · Quiz 40 questions",
        "BOOK J5  ·  Journée 5  ·  Certification FIDELIS × UNICEF France × SHY-Performance",fn)
    return save(p2,f"FIM_BookJ5_Synthese_Simulations_{VER}.pptx")


# ══════════════════════════════════════════════════════════════════════════
#  MAIN
# ══════════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    print(f"\n🎨 FIM {VER} — FINAL PLUS — Logo·Quiz 40Q·Merci·Ateliers J6+J7\n")
    files = []
    files.append(build_book0())
    files.append(build_book1())
    files.append(build_bookJ1())
    files.append(build_bookJ2())
    files.append(build_bookJ3())
    files.append(build_bookJ4())
    files.append(build_bookJ5())
    files.append(build_quiz())
    files.append(build_J6())
    files.append(build_J7())
    print(f"\n✅ {len(files)} modules générés dans {OUT}/")
    for f in files:
        print(f"   {os.path.basename(f)}")

"""
FIM V7-01062026SHY-TE — Corrections complètes
- Police corps 18pt · Titres 26pt · Sous-titres 22pt
- Script accueil corrigé (ALLO + écoute active + identification)
- 3 phases appel sortant (Décollage/Vol/Atterrissage)
- Trame accroche 5 points (Présentation→Dramatisation→Solution→Confiance→Sensibilisation)
- Règles de prélèvement (date du 10, avant/après le 04)
- 4 types de PA + qualification avancée
- Verrouillage / Closing / Repêchage
- 5 jours de formation (non production)
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
GRAY   = RGBColor(0x55, 0x55, 0x55)

SZ_T   = 26   # titres
SZ_S   = 22   # sous-titres
SZ_B   = 18   # corps
SZ_SM  = 14   # secondaire
SZ_XS  = 10   # labels / quiz

LOGO   = "/home/user/fatou/shy_logo_v2.png"
OUT    = "/home/user/fatou"
VER    = "V7-01062026SHY-TE"
CO     = "SHY-Performance"

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

def T(s, text, x, y, w, h, sz=SZ_B, bold=False, col=None,
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
    try:
        s.shapes.add_picture(LOGO, Inches(x), Inches(y), Inches(w), Inches(h))
    except Exception:
        pass

def FOOTER(s):
    R(s, 0, 7.1, 13.33, 0.4, fill=TEAL_D)
    T(s, f"{CO}  ·  Formation FIM  ·  UNICEF France  ·  {VER}",
      0, 7.13, 13.33, 0.3, sz=SZ_XS, col=WHITE, align=PP_ALIGN.CENTER)

def H_TITLE(s, title, sub="", tag=""):
    R(s, 0, 0, 13.33, 7.5, fill=TEAL_L)
    R(s, 0, 0, 13.33, 2.55, fill=TEAL_D)
    R(s, 0, 2.53, 13.33, 0.06, fill=TEAL_A)
    T(s, CO, 0.38, 0.1, 5, 0.52, sz=SZ_S, bold=True, col=WHITE)
    T(s, "Formation Initiale Module — Fundraising Téléphonique UNICEF",
      0.38, 0.62, 9, 0.36, sz=SZ_SM, col=RGBColor(0xBB,0xEE,0xEE), italic=True)
    if tag: T(s, tag, 0.38, 1.0, 9, 0.3, sz=SZ_XS+2, bold=True, col=RGBColor(0x88,0xDD,0xDD))
    T(s, title, 0.38, 1.35, 11.5, 1.05, sz=SZ_T, bold=True, col=WHITE)
    if sub: T(s, sub, 0.38, 2.12, 11.5, 0.42, sz=SZ_SM, col=RGBColor(0xCC,0xEE,0xEE), italic=True)
    LOGO_ADD(s); FOOTER(s)

def H_CONTENT(s, title):
    R(s, 0, 0, 13.33, 7.5, fill=TEAL_L)
    R(s, 0, 0, 13.33, 0.82, fill=TEAL_D)
    R(s, 0, 0.8, 13.33, 0.06, fill=TEAL_A)
    T(s, CO, 0.18, 0.1, 2.2, 0.44, sz=SZ_SM+2, bold=True, col=WHITE)
    R(s, 2.52, 0.1, 0.05, 0.6, fill=TEAL_A)
    T(s, title, 2.68, 0.08, 10.4, 0.65, sz=SZ_T, bold=True, col=WHITE)
    LOGO_ADD(s); FOOTER(s)

def MERCI(p, module_tag=""):
    s = p.slides.add_slide(SL(p))
    R(s, 0, 0, 13.33, 7.5, fill=TEAL_D)
    R(s, 0, 0, 13.33, 0.08, fill=TEAL_A)
    R(s, 0, 7.42, 13.33, 0.08, fill=TEAL_A)
    sh = s.shapes.add_shape(9, Inches(4.5), Inches(1.2), Inches(4.3), Inches(4.3))
    sh.fill.solid(); sh.fill.fore_color.rgb = TEAL_M; sh.line.fill.background()
    T(s, "🙏", 5.5, 2.2, 2.3, 1.5, sz=48, align=PP_ALIGN.CENTER, col=WHITE)
    T(s, "Merci à Tous !", 1.5, 1.4, 10.3, 1.2,
      sz=46, bold=True, col=WHITE, align=PP_ALIGN.CENTER)
    T(s, "Votre engagement fait la différence pour les enfants du monde entier.",
      1.5, 2.85, 10.3, 0.6, sz=SZ_B, col=RGBColor(0xCC,0xEE,0xEE),
      align=PP_ALIGN.CENTER, italic=True)
    if module_tag:
        T(s, module_tag, 1.5, 3.55, 10.3, 0.38, sz=SZ_SM,
          col=RGBColor(0x88,0xDD,0xDD), align=PP_ALIGN.CENTER)
    T(s, "« Le refus d'aujourd'hui peut être le don de demain »",
      2, 4.1, 9.33, 0.5, sz=SZ_SM+2, italic=True,
      col=RGBColor(0xAA,0xDD,0xDD), align=PP_ALIGN.CENTER)
    T(s, f"{CO}  ·  Formation FIM  ·  UNICEF France  ·  {VER}",
      0, 5.0, 13.33, 0.36, sz=SZ_XS+2, col=RGBColor(0x88,0xDD,0xDD), align=PP_ALIGN.CENTER)
    LOGO_ADD(s, x=6.2, y=5.6, w=1.0, h=0.98)
    T(s, "SHY-Performance  ×  FIDELIS  ×  UNICEF France",
      0, 6.6, 13.33, 0.36, sz=SZ_XS+2, bold=True, col=WHITE, align=PP_ALIGN.CENTER)

def TWO_COL(s, lt, li, rt, ri, y=0.9):
    cw = 6.18
    for head, items, ox in [(lt, li, 0.3), (rt, ri, 6.83)]:
        R(s, ox, y, cw, 0.44, fill=TEAL)
        T(s, head, ox+0.12, y+0.06, cw-0.2, 0.34, sz=SZ_SM, bold=True, col=WHITE)
        iy = y + 0.46
        for i, item in enumerate(items):
            bg = WHITE if i%2==0 else TEAL_L
            R(s, ox, iy, cw, 0.54, fill=bg, line=TEAL, lw=Pt(0.4))
            T(s, "• "+item, ox+0.14, iy+0.07, cw-0.24, 0.44, sz=SZ_B, col=DARK)
            iy += 0.56

def HL(s, text, x, y, w, h, fill=None, tc=None, sz=None):
    fill = fill or TEAL; tc = tc or WHITE
    sz = sz or SZ_SM
    R(s, x, y, w, h, fill=fill)
    T(s, text, x+0.15, y+0.08, w-0.3, h-0.16, sz=sz, col=tc, wrap=True)

def KPI_BAR(s, items, y=6.08, label="À retenir"):
    R(s, 0.3, y, 12.73, 0.3, fill=TEAL)
    T(s, label, 0.42, y+0.05, 12.5, 0.22, sz=SZ_XS+2, bold=True, col=WHITE)
    w = 12.73/len(items)
    for i,item in enumerate(items):
        bg = WHITE if i%2==0 else TEAL_L
        R(s, 0.3+i*w, y+0.32, w-0.04, 0.58, fill=bg, line=TEAL, lw=Pt(0.6))
        T(s, "✔ "+item, 0.42+i*w, y+0.36, w-0.18, 0.5, sz=SZ_SM, col=DARK, wrap=True)

def TH(s, cols, cxs, cws, y):
    for h,x,w in zip(cols,cxs,cws):
        R(s, x, y, w, 0.35, fill=TEAL_D)
        T(s, h, x+0.06, y+0.05, w-0.1, 0.28, sz=SZ_XS+2, bold=True, col=WHITE,
          align=PP_ALIGN.CENTER if len(h)<=3 else PP_ALIGN.LEFT)

def SEQ4(s, steps, labels, y=0.9, h=5.0):
    colors = [TEAL_D, TEAL, TEAL_M, TEAL_A]
    sw = (13.33-0.6)/4
    for i,(step,label,col) in enumerate(zip(steps,labels,colors)):
        bx = 0.3+i*(sw+0.02)
        R(s, bx, y, sw, 0.36, fill=col)
        T(s, label, bx+0.1, y+0.05, sw-0.18, 0.28, sz=SZ_XS+2, bold=True, col=WHITE)
        R(s, bx, y+0.36, sw, h-0.36, fill=WHITE, line=col, lw=Pt(1.2))
        T(s, step, bx+0.12, y+0.44, sw-0.22, h-0.6, sz=SZ_SM, col=DARK, wrap=True)

def save(p, name):
    path = os.path.join(OUT, name)
    p.save(path)
    print(f"  ✓ {name}  ({os.path.getsize(path)//1024} KB)")
    return path


# ══════════════════════════════════════════════════════════════════════════
#  QUIZ 40Q (tailles préservées pour lisibilité des 4Q/slide)
# ══════════════════════════════════════════════════════════════════════════

QUIZ = [
    ("A","En quelle année l'UNICEF a-t-il été fondé ?",["a) 1901","b) 1946 ✓","c) 1959","d) 1919"]),
    ("A","Dans combien de pays l'UNICEF est-il présent ?",["a) 120","b) 150","c) 190 ✓","d) 210"]),
    ("A","Quel % des enfants du monde l'UNICEF vaccine-t-il ?",["a) 25%","b) 35%","c) 45% ✓","d) 60%"]),
    ("A","Que signifie RUTF ?",["a) Rapid Urban Treatment Fund","b) Ready-to-Use Therapeutic Food ✓","c) Regional Unified Task Force","d) Routine Universal Treatment"]),
    ("A","Quel est le poids standard d'un sachet RUTF ?",["a) 45g","b) 75g","c) 92g ✓","d) 120g"]),
    ("A","Quelle est la cause centrale de la campagne UNICEF × SHY-Performance ?",["a) Accès à l'éducation","b) Malnutrition aigüe sévère ✓","c) Eau potable","d) Travail des enfants"]),
    ("A","Quelle organisation a précédé l'UNICEF ?",["a) La Croix-Rouge","b) La SDN","c) L'UNRRA ✓","d) Le Fond ONU"]),
    ("A","Lequel NE fait PAS partie des missions UNICEF ?",["a) Vaccination","b) Eau potable","c) Financement partis politiques ✓","d) Protection enfance"]),
    ("B","Date de la loi associative française ?",["a) 14 juil. 1789","b) 1er juil. 1901 ✓","c) 5 mai 1946","d) 10 mars 1972"]),
    ("B","Combien d'associations en France ?",["a) 250 000","b) 800 000","c) 1,5 million ✓","d) 3 millions"]),
    ("B","Quel événement de 1859 est l'acte fondateur de l'humanitaire moderne ?",["a) Conférence Berlin","b) Traité Genève","c) Bataille de Solférino ✓","d) Révolution industrielle"]),
    ("B","Quel type d'association bénéficie du statut le plus reconnu par l'État ?",["a) De fait","b) Déclarée","c) Utilité publique ✓","d) Internationale"]),
    ("B","Rôle du Comité Don en Confiance ?",["a) Gérer finances publiques","b) Certifier transparence et éthique ✓","c) Former agents","d) Contrôler RGPD"]),
    ("B","Combien de bénévoles dans le secteur associatif ?",["a) 5 millions","b) 12 millions","c) 22 millions ✓","d) 35 millions"]),
    ("B","Lequel N'est PAS un principe humanitaire ?",["a) Humanité","b) Impartialité","c) Solidarité nationale ✓","d) Neutralité"]),
    ("C","Combien d'étapes comporte le script UNICEF ?",["a) 4","b) 5","c) 7 ✓","d) 10"]),
    ("C","Combien de règles d'accroche officielles ?",["a) 2","b) 3","c) 5 ✓","d) 7"]),
    ("C","1ère étape après la présentation ?",["a) Coordonnées bancaires","b) Accroche mission UNICEF ✓","c) Montant minimum","d) Vérif identité"]),
    ("C","Combien de modes de validation du don ?",["a) 2","b) 3","c) 4 ✓","d) 6"]),
    ("C","Lequel N'est PAS un mode de validation ?",["a) PA IBAN à chaud","b) PA en ligne direct","c) Virement bancaire unique différé ✓","d) PA courrier"]),
    ("C","À quelle étape présente-t-on les sachets RUTF ?",["a) Étape 7","b) Étapes 2-3 ✓","c) Seulement en objection","d) Dès l'accroche"]),
    ("C","Quel produit est proposé lors de chaque appel ?",["a) Don ponctuel carte","b) Prélèvement automatique régulier ✓","c) Abonnement par chèque","d) Parrainage ponctuel"]),
    ("D","Objectif minimum CU/H ?",["a) 5","b) 7","c) 9 ✓","d) 12"]),
    ("D","Que représente TX Transfo ?",["a) Transferts de dossiers","b) Taux de transformation PEL/PA parmi CU ✓","c) Temps moyen traitement","d) Taux transfert superviseur"]),
    ("D","Qualifications d'un Contact Utile ?",["a) Positif/Neutre/Négatif","b) DON/INDÉCIS/REFUS ✓","c) Chaud/Tiède/Froid","d) PA/PEL/PDC"]),
    ("D","Durée effective de production / journée standard ?",["a) 5h00","b) 6h00","c) 6h40 ✓","d) 8h00"]),
    ("D","Que désigne PDC ?",["a) Protocole Communication","b) Plan de Charge mensuel ✓","c) Programme Collecte Dons","d) Priorité Campagne"]),
    ("D","Différence Conquête vs Fidélisation ?",["a) Conquête = anciens donnateurs","b) Conquête = nouveaux prospects ✓","c) Les deux ciblent les mêmes fichiers","d) Conquête = appels entrants"]),
    ("E","Combien de catégories d'objections ?",["a) 3","b) 5 ✓","c) 7","d) 10"]),
    ("E","Acronyme AAR ?",["a) Analyser/Adapter/Répondre","b) Accepter/Argumenter/Reformuler","c) Accuser réception/Argumenter/Relancer ✓","d) Approcher/Assurer/Recadrer"]),
    ("E","Combien d'objections au total ?",["a) 8","b) 10","c) 15 ✓","d) 20"]),
    ("E","Catégorie 'Vous n'avez pas le droit de m'appeler' ?",["a) Cat.2 Financières","b) Cat.3 Contre PA","c) Cat.4 Méfiance IBAN","d) Cat.5 Conflictuelles ✓"]),
    ("E","Objections initiales (Cat.1) ?",["a) 4","b) 6","c) 8 ✓","d) 10"]),
    ("E","Preuve à invoquer pour objection financière ?",["a) Nombre de pays","b) Impact chiffré du PA (1 enfant 6 semaines) ✓","c) Nombre de bénévoles","d) Date fondation"]),
    ("E","Qu'est-ce que Bloctel ?",["a) Logiciel CRM","b) Liste opposition au démarchage téléphonique ✓","c) Nom du script UNICEF","d) Label qualité associations"]),
    ("F","Règle éthique systématique ?",["a) Insister si hésitation","b) Proposer don ponctuel si PA refusé","c) Respecter tout refus définitif avec courtoisie ✓","d) Enregistrer sans informer"]),
    ("F","PEL signifie ?",["a) Plan Envoi Lien","b) Promesse En Ligne ✓","c) Processus Engagement Légal","d) Protocole Écoute"]),
    ("F","Posture vocale attendue ?",["a) Autoritaire","b) Neutre et distante","c) Chaleureuse, souriante, naturelle ✓","d) Rapide et concise"]),
    ("F","Mantra de la FIM ?",["a) Un refus est un refus","b) Le refus d'aujourd'hui = le don de demain ✓","c) Insistez toujours","d) Le silence vaut mieux"]),
    ("F","Durée standard d'un appel complet ?",["a) 30s–1min","b) 1–2 min","c) 3–5 min ✓","d) 8–10 min"]),
]

def build_quiz():
    p = prs()
    parts = {
        "A":("Partie A","Connaissance de l'UNICEF & de la Cause","Q1–Q8",TEAL_D),
        "B":("Partie B","Secteur Associatif & Humanitaire","Q9–Q15",TEAL),
        "C":("Partie C","Script officiel & Règles d'Accroche","Q16–Q22",TEAL_M),
        "D":("Partie D","KPIs, Nomenclature FIDELIS & Production","Q23–Q28",TEAL_A),
        "E":("Partie E","Traitement des Objections & Méthode AAR","Q29–Q35",RGBColor(0x33,0x77,0x77)),
        "F":("Partie F","Éthique, Posture & Vocabulaire métier","Q36–Q40",RGBColor(0x22,0x66,0x66)),
    }
    s = p.slides.add_slide(SL(p))
    H_TITLE(s,"Quiz d'Évaluation Finale — 40 Questions",
        "Formation Initiale Module FIM  ·  SHY-Performance × FIDELIS × UNICEF France",
        tag="BOOK QUIZ  ·  Évaluation finale des connaissances")
    T(s,"Nom de l'apprenant : ___________________________________",0.5,2.75,9,0.45,sz=SZ_SM+2,bold=True,col=DARK)
    T(s,"Date : _______________________",0.5,3.25,5,0.45,sz=SZ_SM+2,bold=True,col=DARK)
    T(s,"Score : ______ / 40    Seuil de validation : 28/40 (70%)",0.5,3.75,9,0.45,sz=SZ_SM+2,bold=True,col=TEAL_D)
    x=0.3
    for pk in ["A","B","C","D","E","F"]:
        pt,desc,qr,col=parts[pk]
        R(s,x,4.4,2.1,0.55,fill=col)
        T(s,f"{pt}\n{desc}\n{qr}",x+0.08,4.43,1.94,0.49,sz=SZ_XS,bold=True,col=WHITE,align=PP_ALIGN.CENTER)
        x+=2.14
    q_list=[(i+1,q[0],q[1],q[2]) for i,q in enumerate(QUIZ)]
    for slide_i in range(0,len(q_list),4):
        batch=q_list[slide_i:slide_i+4]
        pk=batch[0][1]
        s=p.slides.add_slide(SL(p))
        R(s,0,0,13.33,7.5,fill=TEAL_L)
        R(s,0,0,13.33,0.82,fill=TEAL_D)
        R(s,0,0.8,13.33,0.05,fill=TEAL_A)
        T(s,CO,0.18,0.1,2.2,0.44,sz=SZ_SM,bold=True,col=WHITE)
        R(s,2.52,0.1,0.05,0.6,fill=TEAL_A)
        T(s,f"Quiz FIM — {parts[pk][0]} : {parts[pk][1]} ({parts[pk][2]})",
          2.68,0.08,10.4,0.65,sz=SZ_SM+2,bold=True,col=WHITE)
        LOGO_ADD(s); FOOTER(s)
        y=0.9; cols=[0.3,6.83]; rows_done=0
        for qi,(qnum,qpart,qtext,opts) in enumerate(batch):
            ox=cols[qi%2]
            if qi%2==0 and qi>0: y+=2.85
            if rows_done>=2: break
            pc=parts[qpart][3]
            R(s,ox,y,0.5,2.6,fill=pc)
            T(s,f"Q{qnum}",ox+0.04,y+0.85,0.42,0.6,sz=SZ_XS+2,bold=True,col=WHITE,align=PP_ALIGN.CENTER)
            R(s,ox+0.52,y,5.66,0.68,fill=pc)
            T(s,qtext,ox+0.62,y+0.08,5.46,0.56,sz=SZ_XS+1,bold=True,col=WHITE,wrap=True)
            for oi,opt in enumerate(opts):
                og=WHITE if oi%2==0 else TEAL_L
                R(s,ox+0.52,y+0.7+oi*0.48,5.66,0.46,fill=og,line=pc,lw=Pt(0.4))
                T(s,opt,ox+0.66,y+0.75+oi*0.48,5.42,0.36,sz=SZ_XS,col=DARK)
            if qi%2==1: rows_done+=1
    s=p.slides.add_slide(SL(p))
    H_CONTENT(s,"Corrigé — Récapitulatif des Bonnes Réponses")
    T(s,"Les réponses marquées ✓ sont les réponses correctes.",0.35,0.92,12.6,0.4,sz=SZ_SM,col=DARK,italic=True)
    summary=[("Q1–Q8\nPartie A","b/c/c/b/c/b/c/c",TEAL_D),("Q9–Q15\nPartie B","b/c/c/c/b/c/c",TEAL),
             ("Q16–Q22\nPartie C","c/c/b/c/c/b/b",TEAL_M),("Q23–Q28\nPartie D","c/b/b/c/b/b",TEAL_A),
             ("Q29–Q35\nPartie E","b/c/c/d/c/b/b",RGBColor(0x33,0x77,0x77)),
             ("Q36–Q40\nPartie F","c/b/c/b/c",RGBColor(0x22,0x66,0x66))]
    x=0.3
    for label,answers,col in summary:
        R(s,x,1.5,2.08,0.6,fill=col)
        T(s,label,x+0.08,1.52,1.92,0.56,sz=SZ_XS+2,bold=True,col=WHITE,align=PP_ALIGN.CENTER)
        R(s,x,2.1,2.08,1.6,fill=WHITE,line=col,lw=Pt(1))
        T(s,answers,x+0.1,2.18,1.88,1.42,sz=SZ_SM,col=TEAL_D,wrap=True)
        x+=2.18
    HL(s,"Seuil de validation : 28/40 (70%)  ·  En dessous : révision ciblée + rattrapage J+15  ·  Score ≥ 35/40 : mention Excellent",
       0.3,3.9,12.73,0.48,fill=TEAL_D)
    MERCI(p,"Quiz d'Évaluation Finale — Formation FIM")
    return save(p,f"FIM_BookQUIZ_40Questions_{VER}.pptx")


# ══════════════════════════════════════════════════════════════════════════
#  BOOK 0 — FIL CONDUCTEUR
# ══════════════════════════════════════════════════════════════════════════

def build_book0():
    p2=prs()
    s=p2.slides.add_slide(SL(p2))
    H_TITLE(s,"Bienvenue dans la Formation FIM",
        "Fundraiser Impact Mission — UNICEF France × SHY-Performance × FIDELIS",
        tag="BOOK 0  ·  Fil Conducteur  ·  Document de parcours")
    T(s,"Ce livret est votre guide pour les 5 jours de formation + 2 ateliers pratiques.\n"
        "Il contient votre planning, vos engagements et votre tableau de bord personnel.",
        0.5,2.75,11.8,0.9,sz=SZ_B,col=DARK)
    T(s,"Vous allez maîtriser le script officiel UNICEF en 7 étapes, qualifier vos appels\n"
        "selon la nomenclature FIDELIS et atteindre l'objectif de 9 CU/H minimum.",
        0.5,3.75,11.8,0.9,sz=SZ_B,bold=True,col=TEAL_D)

    s=p2.slides.add_slide(SL(p2))
    H_CONTENT(s,"Vos 5 Journées de Formation + 2 Ateliers")
    days=[("J1","Fondations\n& Mission","Mission UNICEF · KPI FIDELIS · Nomenclature · Organisation"),
          ("J2","Monde\nAssociatif","Loi 1901 · Histoire humanitaire · 3 principes · Don régulier"),
          ("J3","Analyse &\nScript","Script 7 étapes · 5 règles accroche · 4 modes validation"),
          ("J4","Traitement\nObjections","15 objections · 5 catégories · Méthode AAR · Jeux de rôle"),
          ("J5","Synthèse &\nCertif.","Simulations · Grille 24 critères · Quiz 40Q · Certifications"),
          ("J6","Atelier\nPratique 1","Appels simulés live · Grille formateur · Feedback immédiat"),
          ("J7","Atelier\nPratique 2","Certification finale · Plan prise de poste J+0→J+30")]
    x=0.3
    for tag,title,desc in days:
        R(s,x,0.9,1.82,0.38,fill=TEAL_D)
        T(s,tag,x+0.06,0.92,1.7,0.3,sz=SZ_XS+2,bold=True,col=WHITE)
        R(s,x,1.28,1.82,0.95,fill=TEAL)
        T(s,title,x+0.08,1.32,1.66,0.88,sz=SZ_SM,bold=True,col=WHITE,align=PP_ALIGN.CENTER)
        R(s,x,2.23,1.82,3.6,fill=WHITE,line=TEAL,lw=Pt(1))
        T(s,desc,x+0.1,2.3,1.62,3.4,sz=SZ_XS+2,col=DARK,wrap=True)
        x+=1.86
    HL(s,"35h formation + 2 ateliers pratiques · 9h30–17h30 · 6h40 production effective · Certification FIDELIS × UNICEF France",
        0.3,6.08,12.73,0.42)

    s=p2.slides.add_slide(SL(p2))
    H_CONTENT(s,"Votre Contrat d'Engagement")
    TWO_COL(s,"Vos engagements",
        ["Participer activement à toutes les mises en situation",
         "Compléter votre livret personnel chaque soir",
         "Demander un feedback précis après chaque exercice",
         "Respecter la confidentialité des échanges",
         "Arriver à l'heure — le groupe démarre ensemble"],
        "Ce que SHY-Performance vous offre",
        ["Un feedback individuel personnalisé chaque jour",
         "Des exercices pratiques progressifs et bienveillants",
         "Un formateur qui adapte le rythme à votre profil",
         "Un suivi à J+15 et J+30 après la formation",
         "Une certification reconnue FIDELIS × UNICEF France"])
    HL(s,"Mantra : « Le refus d'aujourd'hui peut être le don de demain »",0.3,6.05,12.73,0.46)
    MERCI(p2,"Book 0 — Fil Conducteur · Formation FIM")
    return save(p2,f"FIM_Book0_FilConducteur_{VER}.pptx")


# ══════════════════════════════════════════════════════════════════════════
#  BOOK 1 — SYNOPSIS & OBJECTIFS
# ══════════════════════════════════════════════════════════════════════════

def build_book1():
    p2=prs()
    s=p2.slides.add_slide(SL(p2))
    H_TITLE(s,"Vos Objectifs de Formation",
        "Ce que vous saurez faire à l'issue des 5 jours de formation + 2 ateliers FIM",
        tag="BOOK 1  ·  Synopsis & Objectifs  ·  SHY-Performance × FIDELIS × UNICEF France")
    T(s,"À l'issue de ces 7 jours, vous serez capable de conduire un appel de collecte de dons\n"
        "au nom de l'UNICEF, de la phrase d'accroche jusqu'à la validation du Prélèvement Automatique.",
        0.5,2.75,11.8,0.9,sz=SZ_B,col=DARK)

    s=p2.slides.add_slide(SL(p2))
    H_CONTENT(s,"Vos 3 Dimensions de Compétence")
    cols=[
        ("SAVOIR","Connaissance de la cause",TEAL_D,
         ["Présenter l'UNICEF, sa mission, 75 ans, 190 pays",
          "Expliquer la malnutrition et les sachets RUTF (92g)",
          "Situer le secteur associatif (Loi 1901, Bloctel, RGPD)",
          "Citer les arguments de légitimité UNICEF",
          "Comprendre Conquête vs Fidélisation/Réactivation"]),
        ("SAVOIR-FAIRE","Technique d'appel & qualification",TEAL,
         ["Appliquer les 5 règles officielles de l'accroche",
          "Conduire le script officiel UNICEF en 7 étapes",
          "Qualifier chaque appel (DON / INDÉCIS / REFUS)",
          "Gérer les 4 modes de validation du don",
          "Atteindre l'objectif de 9 CU/H minimum"]),
        ("SAVOIR-ÊTRE","Posture & conviction",TEAL_M,
         ["Maintenir un ton chaleureux et naturel tout au long",
          "Gérer les objections sensibles avec professionnalisme",
          "Transmettre une conviction sincère pour la mission",
          "Intégrer le mantra : «Le refus d'aujourd'hui…»",
          "Respecter les règles éthiques du secteur"]),
    ]
    x=0.3
    for label,sub,col,items in cols:
        R(s,x,0.9,4.18,0.56,fill=col)
        T(s,label,x+0.12,0.92,4.0,0.34,sz=SZ_S,bold=True,col=WHITE)
        T(s,sub,x+0.12,1.26,4.0,0.24,sz=SZ_XS+2,col=RGBColor(0xCC,0xEE,0xEE))
        iy=1.52
        for i,item in enumerate(items):
            bg=WHITE if i%2==0 else TEAL_L
            R(s,x,iy,4.18,0.58,fill=bg,line=col,lw=Pt(0.4))
            T(s,"• "+item,x+0.12,iy+0.08,3.96,0.46,sz=SZ_SM,col=DARK)
            iy+=0.6
        x+=4.35

    s=p2.slides.add_slide(SL(p2))
    H_CONTENT(s,"Indicateurs de Performance — Nomenclature FIDELIS")
    kpis=[("CU/H","Contacts Utiles / Heure","Contacts qualifiés par heure de production","≥ 9 CU/H"),
          ("TX Transfo","Taux de Transformation","% PEL/PA parmi les Contacts Utiles","Cible FIDELIS"),
          ("PDC","Plan de Charge","Volume mensuel de dons réguliers FIDELIS","Volume défini"),
          ("PEL","Promesse En Ligne","Contact ayant accepté de valider son don en ligne","Comptabilisé TX"),
          ("PA","Prélèvement Automatique","Don régulier sécurisé par mandat SEPA","Produit unique"),
          ("Qualif.","Conformité des qualifications","DON / INDÉCIS / REFUS — nomenclature conforme","≥ 50% qualité")]
    cxs=[0.3,1.35,4.65,10.38]; cws=[1.02,3.27,5.7,2.65]
    TH(s,["Code","Libellé","Définition","Objectif"],cxs,cws,y=0.9)
    y=1.28
    for i,(code,lib,defin,obj) in enumerate(kpis):
        bg=WHITE if i%2==0 else TEAL_L; rh=0.7
        R(s,cxs[0],y,cws[0],rh,fill=TEAL_D)
        T(s,code,cxs[0]+0.06,y+0.12,cws[0]-0.1,rh-0.2,sz=SZ_XS+2,bold=True,col=WHITE,align=PP_ALIGN.CENTER)
        R(s,cxs[1],y,cws[1],rh,fill=bg)
        T(s,lib,cxs[1]+0.1,y+0.12,cws[1]-0.15,rh-0.2,sz=SZ_SM,bold=True,col=TEAL_D)
        R(s,cxs[2],y,cws[2],rh,fill=bg)
        T(s,defin,cxs[2]+0.1,y+0.1,cws[2]-0.15,rh-0.14,sz=SZ_SM,col=DARK,italic=True)
        R(s,cxs[3],y,cws[3],rh,fill=TEAL)
        T(s,obj,cxs[3]+0.1,y+0.16,cws[3]-0.15,rh-0.26,sz=SZ_SM,bold=True,col=WHITE,align=PP_ALIGN.CENTER)
        y+=rh+0.02
    MERCI(p2,"Book 1 — Synopsis & Objectifs · Formation FIM")
    return save(p2,f"FIM_Book1_Synopsis_Objectifs_{VER}.pptx")


# ══════════════════════════════════════════════════════════════════════════
#  J1 — FONDATIONS & MISSION
# ══════════════════════════════════════════════════════════════════════════

def build_bookJ1():
    p2=prs()
    s=p2.slides.add_slide(SL(p2))
    H_TITLE(s,"Fondations & Mission",
        "Votre identité de fundraiser · La cause · Votre raison d'agir",
        tag="BOOK J1  ·  Journée 1  ·  UNICEF France × SHY-Performance")

    s=p2.slides.add_slide(SL(p2))
    H_CONTENT(s,"Journée 1 — Ce que vous allez vivre aujourd'hui")
    SEQ4(s,["Témoignage bénéficiaire UNICEF.\nNotez ce que vous ressentez.",
             "En groupe : 'Qu'est-ce qui vous\na le plus touché ?'",
             "Mission UNICEF · KPI FIDELIS\nNomenclature · Organisation journée",
             "Pitch mission 60 secondes\nà votre binôme — sans notes."],
          ["IMMERSION","ÉCHANGE","APPORT","MISE EN PRATIQUE"],y=0.9,h=5.0)
    KPI_BAR(s,["Pitch mission 60s validé","Quiz 5Q UNICEF (seuil 60%)","3 lignes livret"],y=6.1)

    s=p2.slides.add_slide(SL(p2))
    H_CONTENT(s,"L'UNICEF & La Cause — Vos Arguments de Conviction")
    facts=[("1946","Fondation UNICEF (ONU)"),("190","pays présents dans le monde"),
           ("45%","enfants du monde vaccinés"),("92g","poids sachet RUTF malnutrition"),
           ("3 princ.","Humanité · Impartialité · Neutralité"),("9 CU/H","votre objectif quotidien minimum")]
    x=0.3
    for val,label in facts:
        R(s,x,0.9,2.1,1.1,fill=TEAL_D)
        T(s,val,x+0.1,0.92,1.9,0.65,sz=22,bold=True,col=WHITE,align=PP_ALIGN.CENTER)
        T(s,label,x+0.1,1.55,1.9,0.42,sz=SZ_XS+2,col=RGBColor(0xCC,0xEE,0xEE),align=PP_ALIGN.CENTER,wrap=True)
        x+=2.16
    HL(s,"La cause : Malnutrition aigüe sévère — 1 enfant meurt de faim toutes les 11 secondes.\nLes sachets RUTF (92g) permettent à un enfant malnutri de guérir en 6 à 8 semaines.",
        0.3,2.2,12.73,0.72,sz=SZ_SM)
    HL(s,"Posture vocale : Sourire audible dès le mot 'Bonjour' · Débit posé · Ton chaleureux\nVous représentez l'UNICEF — votre conviction doit être sincère et perceptible.",
        0.3,3.05,12.73,0.62,sz=SZ_SM)
    kpis_j1=[("CU/H ≥ 9","Contacts Utiles/h"),("TX Transfo","% PEL+PA parmi CU"),
             ("PDC","Plan de Charge"),("DON","Accepte le PA"),
             ("INDÉCIS","À relancer"),("REFUS","Refuse définitivement")]
    x=0.3
    for val,label in kpis_j1:
        R(s,x,3.88,2.1,1.1,fill=TEAL)
        T(s,val,x+0.08,3.9,1.94,0.58,sz=SZ_SM,bold=True,col=WHITE,align=PP_ALIGN.CENTER)
        T(s,label,x+0.08,4.46,1.94,0.48,sz=SZ_XS+2,col=RGBColor(0xCC,0xEE,0xEE),align=PP_ALIGN.CENTER,wrap=True)
        x+=2.16
    HL(s,"Organisation journée : 9h30–17h30 · 6h40 production effective · Conquête vs Fidélisation/Réactivation",
        0.3,5.12,12.73,0.42,sz=SZ_SM)
    MERCI(p2,"BOOK J1  ·  Journée 1  ·  UNICEF France × SHY-Performance")
    return save(p2,f"FIM_BookJ1_Fondations_Mission_{VER}.pptx")


# ══════════════════════════════════════════════════════════════════════════
#  J2 — MONDE ASSOCIATIF
# ══════════════════════════════════════════════════════════════════════════

def build_bookJ2():
    p2=prs()
    s=p2.slides.add_slide(SL(p2))
    H_TITLE(s,"Le Monde Associatif & Humanitaire",
        "Comprendre le secteur pour représenter la cause avec légitimité",
        tag="BOOK J2  ·  Journée 2  ·  France associative & Histoire humanitaire")

    s=p2.slides.add_slide(SL(p2))
    H_CONTENT(s,"Journée 2 — Ce que vous allez vivre aujourd'hui")
    SEQ4(s,["Analyse de 2 campagnes :\n1 réussie, 1 en difficulté.\nPremières observations.",
             "'Qu'est-ce qui a fait\nla différence ? Pourquoi\nl'une a converti ?'",
             "Loi 1901 · Histoire humanitaire\nDon régulier · Don en Confiance\nRGPD · Bloctel",
             "'Expliquez le don régulier\nà un prospect sceptique\nen 90 secondes.'"],
          ["DÉCOUVERTE","ANALYSE","APPORT","MISE EN PRATIQUE"],y=0.9,h=5.0)
    KPI_BAR(s,["Restitution orale étude de cas (5 min)","Fiche secteur associatif complétée","Réflexion LFI : légitimité du PA"],y=6.1)

    s=p2.slides.add_slide(SL(p2))
    H_CONTENT(s,"La France Associative & L'Histoire de l'Humanitaire")
    facts=[("1,5M","associations en France"),("22M","bénévoles en France"),
           ("1er juil\n1901","Loi droit associatif"),("3 types","De fait / Déclarée /\nUtilité publique"),
           ("Don en\nConfiance","Label indépendant"),("RGPD\nBloctel","Cadre légal prospects")]
    x=0.3
    for val,label in facts:
        R(s,x,0.9,2.1,1.1,fill=TEAL_D)
        T(s,val,x+0.1,0.92,1.9,0.64,sz=16,bold=True,col=WHITE,align=PP_ALIGN.CENTER)
        T(s,label,x+0.1,1.54,1.9,0.42,sz=SZ_XS+2,col=RGBColor(0xCC,0xEE,0xEE),align=PP_ALIGN.CENTER,wrap=True)
        x+=2.16
    timeline=[("1859","Bataille de Solférino → Croix-Rouge"),("1863","Fondation Croix-Rouge"),
              ("1946","Fondation UNICEF (ONU)"),("1989","Convention Droits de l'Enfant"),
              ("1989","Comité Don en Confiance"),("2000s","RGPD & Bloctel")]
    x=0.3
    for year,event in timeline:
        R(s,x,2.2,2.1,0.4,fill=TEAL)
        T(s,year,x+0.1,2.23,1.9,0.3,sz=SZ_SM,bold=True,col=WHITE,align=PP_ALIGN.CENTER)
        R(s,x,2.6,2.1,1.2,fill=TEAL_L,line=TEAL,lw=Pt(0.8))
        T(s,event,x+0.12,2.68,1.86,1.06,sz=SZ_SM,col=DARK,wrap=True)
        x+=2.16
    HL(s,"3 principes humanitaires : HUMANITÉ · IMPARTIALITÉ · NEUTRALITÉ",0.3,3.95,12.73,0.42,sz=SZ_SM)
    HL(s,"Don régulier PA : valeur à vie × 8 vs don ponctuel — colonne vertébrale des programmes UNICEF France.",0.3,4.5,12.73,0.42,sz=SZ_SM)
    MERCI(p2,"BOOK J2  ·  Journée 2  ·  France associative & Histoire humanitaire")
    return save(p2,f"FIM_BookJ2_MondeAssociatif_{VER}.pptx")


# ══════════════════════════════════════════════════════════════════════════
#  J3 — SCRIPT OFFICIEL (VERSION COMPLÈTE CORRIGÉE)
# ══════════════════════════════════════════════════════════════════════════

def build_bookJ3():
    p2=prs()
    s=p2.slides.add_slide(SL(p2))
    H_TITLE(s,"Analyse & Script Officiel UNICEF",
        "Les 7 étapes · Les 5 règles d'accroche · Les 4 modes de validation",
        tag="BOOK J3  ·  Journée 3  ·  Script UNICEF × FIDELIS")

    # Slide 2 — Programme J3
    s=p2.slides.add_slide(SL(p2))
    H_CONTENT(s,"Journée 3 — Ce que vous allez vivre aujourd'hui")
    SEQ4(s,["Lecture intégrale du script\nà voix haute.\nPremière impression à froid.",
             "Analyse phrase par phrase :\n'Quelle est l'intention ici ?\nPourquoi ce mot précis ?'",
             "7 étapes · 5 règles accroche\n4 modes de validation\n3 phases de l'appel",
             "Lecture en binômes. Débriefs\ncollectifs. Points forts\net axes d'amélioration."],
          ["DÉCOUVERTE","ANALYSE","APPORT","ENTRAÎNEMENT"],y=0.9,h=5.0)
    KPI_BAR(s,["Script lu 3 fois sans hésitation","5 règles accroche récitées","Score seuil 12/20 pour J4"],y=6.1)

    # Slide 3 — L'Accueil téléphonique (NOUVEAU)
    s=p2.slides.add_slide(SL(p2))
    H_CONTENT(s,"L'Accueil Téléphonique — La Bonne Séquence")
    steps_acc=[
        (TEAL_D,"ÉTAPE 1 — ALLO",
         "Le fundraiser dit 'ALLO' puis s'arrête.\nÉcoute active : identifier le genre de l'interlocuteur\n(Madame ou Monsieur) avant de parler."),
        (TEAL,"ÉTAPE 2 — IDENTIFICATION",
         "Convention française : prénom D'ABORD, puis nom de famille.\n\n\"Est-ce bien Monsieur François DUPONT ?\"\n\nNE PAS commencer par le nom de famille.\nC'est non naturel en français."),
        (TEAL_M,"ÉTAPE 3 — PRÉSENTATION",
         "Après confirmation du contact :\n\"Bonjour Monsieur Dupont, je m'appelle [Prénom],\nje vous appelle au nom de l'UNICEF France.\"\nPuis suivre le script officiel en 7 étapes."),
    ]
    x=0.3; cw=4.17
    for col,title,body in steps_acc:
        R(s,x,0.9,cw,0.56,fill=col)
        T(s,title,x+0.14,0.93,cw-0.22,0.48,sz=SZ_SM,bold=True,col=WHITE)
        R(s,x,1.46,cw,4.3,fill=WHITE,line=col,lw=Pt(1.5))
        T(s,body,x+0.16,1.56,cw-0.28,4.06,sz=SZ_B,col=DARK,wrap=True)
        x+=4.3
    HL(s,"⚠ Jamais de nom de famille en premier — cela sonne froid et commercial. Prénom + NOM = chaleur + clarté.",
        0.3,5.9,12.73,0.48,fill=RGBColor(0x00,0x55,0x55))

    # Slide 4 — 3 phases de l'appel (NOUVEAU)
    s=p2.slides.add_slide(SL(p2))
    H_CONTENT(s,"Les 3 Phases de l'Appel Sortant")
    phases=[
        (TEAL_D,"✈  DÉCOLLAGE","Phase 1 — Accueil & Accroche",
         "• ALLO + écoute active + identification genre\n• Prénom → NOM de famille (convention française)\n• Présentation : 'Je m'appelle [Prénom], UNICEF France'\n• Vérifier disponibilité : 'Vous avez quelques minutes ?'\n• Lancer la trame d'accroche en 5 points\n• Objectif : dépasser 60 secondes sans raccrocher"),
        (TEAL,"🛫  VOL","Phase 2 — Argumentation & Objections",
         "• Script 7 étapes dans l'ordre\n• Dramatisation → Solution → Confiance\n• Traitement des objections par méthode AAR\n• Maximum 2 objections après l'appel au don\n• Ne pas s'arrêter avant l'appel au don\n• Obtenir un accord de principe clair"),
        (TEAL_M,"🛬  ATTERRISSAGE","Phase 3 — Accord de Principe & Validation",
         "• Accord de principe : oui ferme + montant clair\n• Verrouillage : question fermée sécurisante\n• Validation du mode de PA (IBAN / ligne / courrier)\n• Vérification adresse complète (n° de porte !)\n• Confirmation n° de téléphone\n• Closing : 'C'est formidable, merci !'"),
    ]
    x=0.3; cw=4.17
    for col,title,sub,body in phases:
        R(s,x,0.9,cw,0.42,fill=col)
        T(s,title,x+0.14,0.92,cw-0.22,0.36,sz=SZ_S,bold=True,col=WHITE)
        R(s,x,1.32,cw,0.34,fill=RGBColor(0x00,0x66,0x66))
        T(s,sub,x+0.14,1.35,cw-0.22,0.28,sz=SZ_XS+2,bold=True,col=RGBColor(0xCC,0xEE,0xEE))
        R(s,x,1.66,cw,4.6,fill=WHITE,line=col,lw=Pt(1.5))
        T(s,body,x+0.14,1.74,cw-0.24,4.38,sz=SZ_B,col=DARK,wrap=True)
        x+=4.3
    HL(s,"Chaque phase a un objectif clair. Le vol ne peut pas s'arrêter avant l'appel au don.",
        0.3,6.4,12.73,0.4)

    # Slide 5 — Trame d'accroche 5 points (NOUVEAU)
    s=p2.slides.add_slide(SL(p2))
    H_CONTENT(s,"La Trame d'Accroche — 5 Points Clés")
    T(s,"Si objection précoce : accepter, ne pas argumenter, et démarrer immédiatement cette trame.",
       0.35,0.9,12.5,0.4,sz=SZ_SM,bold=True,col=TEAL_D)
    trame=[
        (TEAL_D,"1 — PRÉSENTATION","QUI on parle",
         "Brève présentation de l'organisme.\n'Je vous appelle au nom de l'UNICEF France,\nl'organisation internationale qui protège\nles enfants dans 190 pays.'"),
        (TEAL,"2 — DRAMATISATION","POURQUOI cet appel",
         "Gravité humaine. Prise de conscience.\n'En ce moment même, un enfant meurt\nde malnutrition toutes les 11 secondes.'\nFaire ressentir sans tomber dans l'excès."),
        (TEAL_M,"3 — SOLUTION","COMMENT on aide",
         "Transformer la solution en confiance.\n'Un sachet RUTF de 92g suffit à traiter\nun enfant en 6 à 8 semaines.\nVotre geste peut tout changer.'"),
        (TEAL_A,"4 — CONFIANCE","FAIRE VOYAGER",
         "Faites imaginer le donateur.\nFaites-le voyager dans ses pensées.\nTransformez ses émotions en réaction\nà chaud : 'Imaginez cet enfant...'"),
        (RGBColor(0x33,0x77,0x77),"5 — SENSIBILISATION","APPEL À L'ACTION",
         "Les émotions provoquent la réaction.\n'Avec 10€/mois, vous permettez de\ntraiter un enfant. Est-ce que vous\nsouhaitez nous rejoindre aujourd'hui ?'"),
    ]
    x=0.3; cw=(13.33-0.6)/5
    for col,title,sub,body in trame:
        R(s,x,1.38,cw,0.42,fill=col)
        T(s,title,x+0.1,1.4,cw-0.16,0.38,sz=SZ_XS+2,bold=True,col=WHITE,align=PP_ALIGN.CENTER)
        R(s,x,1.8,cw,0.28,fill=RGBColor(0x00,0x60,0x60))
        T(s,sub,x+0.08,1.83,cw-0.14,0.22,sz=SZ_XS,bold=True,col=RGBColor(0xBB,0xEE,0xEE),align=PP_ALIGN.CENTER)
        R(s,x,2.08,cw,4.1,fill=WHITE,line=col,lw=Pt(1.2))
        T(s,body,x+0.1,2.16,cw-0.16,3.88,sz=SZ_SM,col=DARK,wrap=True)
        x+=cw+0.02
    HL(s,"Résumé : QUI → QUOI → COMMENT → CONFIANCE → SENSIBILISATION\n"
        "Objection précoce : dépasser 60s avant de conclure au moins 1 CU. Ne pas s'arrêter avant l'appel au don.",
        0.3,6.3,12.73,0.52,fill=TEAL_D)

    # Slide 6 — Script 7 étapes
    s=p2.slides.add_slide(SL(p2))
    H_CONTENT(s,"Le Script Officiel UNICEF — Les 7 Étapes")
    steps=[("①\nACCROCHE",TEAL_D,"Ouverture conforme aux 5 règles\nTon chaleureux, prénom+nom,\norganisation, raison, disponibilité"),
           ("②\nMALNUTRITION",TEAL,"Présentation de la cause\n1 enfant/11sec · Dimension\nhumaine et émotionnelle"),
           ("③\nSACHETS RUTF",TEAL_M,"La solution UNICEF : sachets 92g\n6 à 8 semaines pour guérir\nun enfant malnutri"),
           ("④\nAPPEL AU SOUTIEN",TEAL_A,"Proposition du PA mensuel\nMontant suggéré · Impact\nchiffré · Simplicité"),
           ("⑤\nSI DON PONCTUEL",RGBColor(0x33,0x77,0x77),"Réorientation vers le PA\n'Et si on envisageait\nquelque chose de régulier ?'"),
           ("⑥\nINDÉCIS",RGBColor(0x22,0x66,0x66),"Gestion de l'hésitation\nRester dans l'échange\nRelance question ouverte"),
           ("⑦\nVALIDATION",RGBColor(0x11,0x55,0x55),"IBAN à chaud / PA en ligne\nPEL / PA courrier\nConfirmer coordonnées")]
    x=0.3; sw=(13.33-0.6)/7
    for title,col,body in steps:
        R(s,x,0.9,sw,0.5,fill=col)
        T(s,title,x+0.06,0.92,sw-0.1,0.46,sz=SZ_XS+1,bold=True,col=WHITE,align=PP_ALIGN.CENTER)
        R(s,x,1.4,sw,4.85,fill=WHITE,line=col,lw=Pt(1.2))
        T(s,body,x+0.08,1.48,sw-0.14,4.6,sz=SZ_SM,col=DARK,wrap=True)
        x+=sw+0.02
    HL(s,"Durée standard : 3 à 5 min · Max 2 objections après l'appel au don · Accord de principe avant l'atterrissage.",
        0.3,6.38,12.73,0.4)

    # Slide 7 — 5 règles d'accroche
    s=p2.slides.add_slide(SL(p2))
    H_CONTENT(s,"Les 5 Règles Officielles de la Phrase d'Accroche")
    T(s,"Ces 5 règles sont non négociables — elles définissent un appel conforme FIDELIS :",
       0.35,0.9,12.5,0.38,sz=SZ_SM,bold=True,col=TEAL_D)
    rules=[("RÈGLE 1","Se présenter : prénom EN PREMIER, puis nom","'Bonjour, je m'appelle [Prénom NOM]…' — convention française, jamais le nom de famille d'abord."),
           ("RÈGLE 2","Nommer l'organisation représentée","'…je vous appelle au nom de l'UNICEF France…' — le prospect sait immédiatement qui appelle."),
           ("RÈGLE 3","Annoncer la raison de l'appel","'…pour vous parler d'une initiative importante pour les enfants…' — clarté totale."),
           ("RÈGLE 4","Vérifier la disponibilité du prospect","'Est-ce que vous avez quelques minutes ?' — respect de l'interlocuteur, écoute active."),
           ("RÈGLE 5","Ton chaleureux et naturel — le sourire s'entend","Débit posé · Articulation claire · Conviction sincère — le ton est aussi important que les mots.")]
    y=1.36
    for idx,(tag,title,body) in enumerate(rules):
        bg=WHITE if idx%2==0 else TEAL_L
        R(s,0.3,y,1.1,0.88,fill=TEAL_D)
        T(s,tag,0.32,y+0.18,1.06,0.52,sz=SZ_XS+1,bold=True,col=WHITE,align=PP_ALIGN.CENTER)
        R(s,1.42,y,3.5,0.88,fill=TEAL_L,line=TEAL_D,lw=Pt(0.8))
        T(s,title,1.52,y+0.2,3.32,0.52,sz=SZ_SM,bold=True,col=TEAL_D)
        R(s,4.94,y,8.09,0.88,fill=bg)
        T(s,body,5.06,y+0.12,7.88,0.68,sz=SZ_SM,col=DARK,italic=True)
        y+=0.92
    HL(s,"Un appel qui viole une de ces 5 règles est disqualifié dans la nomenclature FIDELIS.",
        0.3,5.96,12.73,0.44)

    # Slide 8 — Règles de prélèvement (NOUVEAU)
    s=p2.slides.add_slide(SL(p2))
    H_CONTENT(s,"Règles de Prélèvement — Date du 10 & Délais")
    T(s,"Le prélèvement PA est effectué le 10 de chaque mois. La date de saisie détermine le mois de prélèvement.",
       0.35,0.9,12.5,0.48,sz=SZ_B,bold=True,col=TEAL_D)

    R(s,0.3,1.5,6.1,0.44,fill=TEAL_D)
    T(s,"PA SAISI AVANT LE 04 DU MOIS",0.42,1.54,5.9,0.34,sz=SZ_S,bold=True,col=WHITE)
    R(s,0.3,1.94,6.1,1.0,fill=WHITE,line=TEAL_D,lw=Pt(1.5))
    T(s,"→ Prélevé le 10 du MÊME mois\n\nExemple : PA signé le 04 juin\n→ Prélèvement le 10 juin ✅",
       0.44,2.0,5.82,0.88,sz=SZ_B,col=DARK)

    R(s,6.7,1.5,6.1,0.44,fill=TEAL)
    T(s,"PA SAISI APRÈS LE 04 DU MOIS",6.82,1.54,5.9,0.34,sz=SZ_S,bold=True,col=WHITE)
    R(s,6.7,1.94,6.1,1.0,fill=WHITE,line=TEAL,lw=Pt(1.5))
    T(s,"→ Prélevé le 10 du MOIS SUIVANT\n\nExemple : PA signé le 05 juin\n→ Prélèvement le 10 juillet ✅",
       6.84,2.0,5.82,0.88,sz=SZ_B,col=DARK)

    R(s,0.3,3.12,12.5,0.4,fill=TEAL_A)
    T(s,"RÈGLE SIMPLE : 04 = date pivot · Avant le 04 = ce mois · Après le 04 = mois suivant",
       0.44,3.15,12.2,0.32,sz=SZ_SM,bold=True,col=WHITE)

    T(s,"Informez TOUJOURS le donateur du mois de son premier prélèvement.",
       0.35,3.65,12.5,0.38,sz=SZ_B,col=TEAL_D)

    R(s,0.3,4.2,12.5,0.36,fill=TEAL_D)
    T(s,"TX D'AUDIOLYSE : 80%  —  8 donateurs sur 10 arrivent à la validation s'ils restent en ligne après la trame d'accroche.",
       0.44,4.24,12.2,0.28,sz=SZ_SM,bold=True,col=WHITE)

    HL(s,"Rappel au donateur : 'Votre premier prélèvement sera effectué le 10 [mois]. Je vous confirme cela maintenant.'",
        0.3,4.7,12.73,0.48,sz=SZ_SM)
    HL(s,"Si le donateur interroge sur la date de prélèvement, c'est une objection à traiter avec la règle du 04 + exemple concret.",
        0.3,5.32,12.73,0.48,sz=SZ_SM)

    # Slide 9 — Types de PA (NOUVEAU)
    s=p2.slides.add_slide(SL(p2))
    H_CONTENT(s,"Les 4 Types de Prélèvement Automatique")
    types_pa=[
        (TEAL_D,"PA EN LIGNE — AGENT EN DIRECT",
         "L'agent collecte l'IBAN à chaud pendant l'appel.\nLe donateur valide oralement.\nL'agent saisit les coordonnées bancaires en direct.\n\n✅ Mode le plus rapide · Valider l'IBAN mot à mot\n⚠ Voyant VERT = reprendre · Voyant ROUGE = suspendre"),
        (TEAL,"PA EN LIGNE — DONATEUR EN DIRECT",
         "L'agent envoie un lien de paiement sécurisé.\nRecommandé par MAIL (non par SMS).\nLe donateur saisit lui-même ses coordonnées.\n\nEnvoi : mail prioritaire, SMS en secours.\nC'est une promesse — à confirmer J+5."),
        (TEAL_M,"PA EN LIGNE — DIFFÉRÉ PAR LE DONATEUR",
         "Promesse de paiement en ligne.\nLe donateur confirme qu'il a payé\nen lisant à voix haute le message de confirmation\n(même message visible sur l'écran agent).\n\n✅ Qualifier en PAIEMENT DIFFÉRÉ confirmé"),
        (TEAL_A,"PROMESSE + RELANCE CALL 2",
         "Promesse de paiement en ligne non encore confirmée.\nCall 2 rappelle pour relance et confirmation.\n2 cas : Call 1 ou Call 2 finalise le PA.\n\nSi pas de réponse au rappel :\n→ Qualifier en PROMESSE DÉFINITIVE PEL"),
    ]
    x=0.3; cw=(13.33-0.6)/4
    for col,title,body in types_pa:
        R(s,x,0.9,cw,0.48,fill=col)
        T(s,title,x+0.1,0.92,cw-0.16,0.44,sz=SZ_XS+2,bold=True,col=WHITE,align=PP_ALIGN.CENTER)
        R(s,x,1.38,cw,4.95,fill=WHITE,line=col,lw=Pt(1.2))
        T(s,body,x+0.1,1.46,cw-0.16,4.72,sz=SZ_SM,col=DARK,wrap=True)
        x+=cw+0.02
    HL(s,"Objection 'D'où avez-vous eu mon numéro ?' → Réponse : fichiers partenaires (Amazon, M6 boutique, comparateurs, etc.)",
        0.3,6.4,12.73,0.42)

    MERCI(p2,"BOOK J3  ·  Journée 3  ·  Script UNICEF × FIDELIS")
    return save(p2,f"FIM_BookJ3_Lecture_Script_{VER}.pptx")


# ══════════════════════════════════════════════════════════════════════════
#  J4 — OBJECTIONS + VERROUILLAGE COMPLET
# ══════════════════════════════════════════════════════════════════════════

def build_bookJ4():
    p2=prs()
    s=p2.slides.add_slide(SL(p2))
    H_TITLE(s,"Traitement des Objections",
        "15 objections · 5 catégories · Méthode AAR · Jeux de rôle téléphonique",
        tag="BOOK J4  ·  Journée 4  ·  Objections UNICEF France × FIDELIS")

    s=p2.slides.add_slide(SL(p2))
    H_CONTENT(s,"Journée 4 — Ce que vous allez vivre aujourd'hui")
    SEQ4(s,["Écoute de 5 extraits :\nbonne et mauvaise gestion\nd'une objection.",
             "'Qu'avez-vous entendu ?\nQu'est-ce qui a tout changé\ndans la réponse ?'",
             "15 objections · 5 catégories\nMéthode AAR\nRègle 60s / règle 2 objections",
             "Jeux de rôle en binômes.\nGrille d'observation formateur.\nVerrouillage & closing."],
          ["ÉCOUTE","ANALYSE","APPORT","ENTRAÎNEMENT"],y=0.9,h=5.0)
    KPI_BAR(s,["Grille écoute active validée","LFI : 'Quelle objection me challenge le plus ?'","Score seuil 14/20 pour validation J4"],y=6.1)

    # Slide méthode AAR
    s=p2.slides.add_slide(SL(p2))
    H_CONTENT(s,"Votre Méthode AAR — Accuser Réception · Argumenter · Relancer")
    aar=[("A\nACCUSER\nRÉCEPTION",TEAL_D,
          "Montrez que vous avez entendu.\nNe justifiez pas immédiatement.\n\n'Je comprends tout à fait…'\n'C'est une question que\nbeaucoup se posent et\nc'est tout à fait légitime…'\n\n⚠ Jamais 'Oui mais…'"),
         ("A\nARGUMENTER",TEAL,
          "1 argument + 1 preuve + 1 chiffre.\n\nRègles d'or :\n→ 1 seul argument par réponse\n→ Preuve vérifiable\n→ Chiffre simple et mémorisable\n\nEx : '10€/mois = 33 centimes/jour\n= 1 enfant traité en 6 semaines'"),
         ("R\nRELANCER",TEAL_M,
          "Revenez à la proposition.\nPas de pression.\n\n'Est-ce que cela répond à\nvotre question ?'\n'Qu'est-ce qui vous permettrait\nde vous sentir à l'aise ?'\n\n⚠ Laissez 3 secondes de silence\navant de relancer.")]
    x=0.3
    for title,col,body in aar:
        R(s,x,0.9,4.17,0.72,fill=col)
        T(s,title,x+0.14,0.93,3.93,0.66,sz=SZ_S,bold=True,col=WHITE,align=PP_ALIGN.CENTER)
        R(s,x,1.62,4.17,4.64,fill=WHITE,line=col,lw=Pt(1.8))
        T(s,body,x+0.16,1.7,3.85,4.42,sz=SZ_B,col=DARK,wrap=True)
        x+=4.3
    HL(s,"Le silence après votre argument est votre allié. Comptez 1–2–3 mentalement avant de relancer.",
        0.3,6.38,12.73,0.44)

    # Slide règle objection précoce (NOUVEAU)
    s=p2.slides.add_slide(SL(p2))
    H_CONTENT(s,"Règle Fondamentale — Gestion des Objections dans l'Appel")
    R(s,0.3,0.9,12.73,0.52,fill=TEAL_D)
    T(s,"⏱ RÈGLE 60 SECONDES — Objection précoce",0.44,0.94,12.4,0.42,sz=SZ_S,bold=True,col=WHITE)
    R(s,0.3,1.42,12.73,1.05,fill=WHITE,line=TEAL_D,lw=Pt(1.5))
    T(s,"Si le prospect objecte AVANT la trame d'accroche :\n"
       "→ Accepter l'objection sans argumenter\n"
       "→ Démarrer immédiatement la trame (Présentation → Dramatisation → Solution)\n"
       "→ Objectif impératif : DÉPASSER 60 SECONDES pour conclure au moins 1 CU",
       0.44,1.5,12.4,0.9,sz=SZ_B,col=DARK)

    R(s,0.3,2.6,12.73,0.52,fill=TEAL)
    T(s,"🚫 RÈGLE 2 OBJECTIONS — Après l'appel au don",0.44,2.64,12.4,0.42,sz=SZ_S,bold=True,col=WHITE)
    R(s,0.3,3.12,12.73,0.95,fill=WHITE,line=TEAL,lw=Pt(1.5))
    T(s,"Après avoir formulé l'appel au don :\n"
       "→ NE PAS DÉPASSER 2 objections traitées\n"
       "→ Aller JUSQU'À l'appel au don sans s'arrêter\n"
       "→ Obtenir un ACCORD DE PRINCIPE avant l'atterrissage",
       0.44,3.2,12.4,0.82,sz=SZ_B,col=DARK)

    R(s,0.3,4.2,12.73,0.52,fill=TEAL_M)
    T(s,"✅ L'ACCORD DE PRINCIPE — Condition de l'atterrissage",0.44,4.24,12.4,0.42,sz=SZ_S,bold=True,col=WHITE)
    R(s,0.3,4.72,12.73,0.88,fill=WHITE,line=TEAL_M,lw=Pt(1.5))
    T(s,"Un OUI ferme avec montant clair AVANT de passer à la validation.\n"
       "Exemples : 'Oui je suis d'accord pour 15€/mois' ou 'Oui, comment ça marche ?'\n"
       "Sans accord de principe = pas d'atterrissage → relancer ou conclure en INDÉCIS.",
       0.44,4.8,12.4,0.74,sz=SZ_B,col=DARK)
    HL(s,"Ne jamais forcer un atterrissage sans accord de principe. Un PA obtenu sous pression = annulation rapide.",
        0.3,5.75,12.73,0.52,fill=TEAL_D)

    # Slide 5 catégories
    s=p2.slides.add_slide(SL(p2))
    H_CONTENT(s,"Les 5 Catégories d'Objections — 15 Situations à Maîtriser")
    cats=[("CAT. 1\nOBJECTIONS INITIALES",TEAL_D,
           "Phase d'accroche — 8 objections :\n• Pas intéressé · Faux numéro\n• Appel pour un don · Je donne déjà\n• Arnaque · Pas par téléphone\n• Pas confiance aux associations\n• N'aime pas être contacté"),
          ("CAT. 2\nFINANCIÈRES",TEAL,
           "Après l'appel au don — 3 objections :\n• Je donne déjà ailleurs\n• Je n'ai pas les moyens\n• Dernier positionnement financier"),
          ("CAT. 3\nCONTRE LE PA",TEAL_M,
           "Sur le prélèvement — 3 objections :\n• Préfère le don ponctuel\n• N'aime pas l'engagement mensuel\n• Dernier positionnement PA"),
          ("CAT. 4\nMÉFIANCE IBAN",TEAL_A,
           "Sur la sécurité bancaire :\n• Protocole SEPA expliqué\n• Droits du donateur (arrêt 1 clic)\n• Réassurance données bancaires"),
          ("CAT. 5\nCONFLICTUELLES",RGBColor(0x33,0x77,0x77),
           "Objections sensibles :\n• Source du numéro → Amazon etc.\n• RGPD / Droits légaux\n• Bloctel / liste rouge\n• Accent / identité appelant")]
    x=0.3; cw=(13.33-0.6)/5
    for title,col,body in cats:
        R(s,x,0.9,cw,0.54,fill=col)
        T(s,title,x+0.1,0.92,cw-0.16,0.5,sz=SZ_XS+2,bold=True,col=WHITE,align=PP_ALIGN.CENTER)
        R(s,x,1.44,cw,4.85,fill=WHITE,line=col,lw=Pt(1.2))
        T(s,body,x+0.1,1.52,cw-0.16,4.62,sz=SZ_SM,col=DARK,wrap=True)
        x+=cw+0.02
    HL(s,"La majorité des appels se termine à Cat.1. Maîtriser ces 8 objections = atteindre vos 9 CU/H.",
        0.3,6.38,12.73,0.44)

    # Slide verrouillage (NOUVEAU)
    s=p2.slides.add_slide(SL(p2))
    H_CONTENT(s,"Le Verrouillage & le Closing — Sécuriser la Promesse")
    R(s,0.3,0.9,6.1,0.44,fill=TEAL_D)
    T(s,"🔒 VERROUILLAGE — Question fermée",0.44,0.94,5.9,0.34,sz=SZ_S,bold=True,col=WHITE)
    R(s,0.3,1.34,6.1,2.05,fill=WHITE,line=TEAL_D,lw=Pt(1.5))
    T(s,"Posez UNE question fermée qui confirme la promesse :\n\n"
       "'Alors on est bien d'accord pour [montant]€/mois à partir du [date] ?'\n\n"
       "⚠ Impact statistique :\n"
       "✅ Verrouillage réussi = 50% des promesses confirmées\n"
       "❌ Sans verrouillage = 1/10 confirmées, le reste est perdu",
       0.44,1.42,5.82,1.9,sz=SZ_SM,col=DARK)

    R(s,6.7,0.9,6.1,0.44,fill=TEAL)
    T(s,"🌟 CLOSING — Valoriser la décision",6.84,0.94,5.9,0.34,sz=SZ_S,bold=True,col=WHITE)
    R(s,6.7,1.34,6.1,2.05,fill=WHITE,line=TEAL,lw=Pt(1.5))
    T(s,"Après l'accord de principe confirmé :\n\n"
       "Dites : 'C'est formidable ! Merci beaucoup.'\n\n"
       "Valorisez AVEC MÉTHODE :\n"
       "→ Rappeler l'impact : 'Votre geste va permettre de traiter un enfant'\n"
       "→ Sécuriser psychologiquement la promesse\n"
       "→ Enchaîner immédiatement sur la validation",
       6.84,1.42,5.82,1.9,sz=SZ_SM,col=DARK)

    R(s,0.3,3.52,12.73,0.44,fill=TEAL_D)
    T(s,"🔁 REPÊCHAGE — Dernier filet si hésitation finale",0.44,3.56,12.4,0.34,sz=SZ_S,bold=True,col=WHITE)
    R(s,0.3,3.96,12.73,0.88,fill=WHITE,line=TEAL_D,lw=Pt(1.2))
    T(s,"Si le donateur hésite après le closing : poser la question fermée de repêchage.\n"
       "Il faut UN OUI DE DON FERME avec montant clair — pas d'ambiguïté.\n"
       "Si refus définitif : clore avec courtoisie. 'Merci de m'avoir écouté. Bonne journée.'",
       0.44,4.04,12.4,0.74,sz=SZ_B,col=DARK)

    HL(s,"Le verrouillage est psychologique — il sécurise la promesse. Sans lui, 90% des promesses non confirmées disparaissent.",
        0.3,4.98,12.73,0.48,fill=TEAL_M)

    # Slide informations à collecter (NOUVEAU)
    s=p2.slides.add_slide(SL(p2))
    H_CONTENT(s,"Informations à Collecter & Vérifier — Fiche Hermès")
    TWO_COL(s,"Fiche Hermès — Données à revérifier en direct",
        ["Nom, prénom du donateur (confirmation genre : M. ou Mme)",
         "Adresse complète avec NUMÉRO DE PORTE (sinon NPAI)",
         "Code postal + ville — vérifier chaque caractère",
         "Email : lire lettre par lettre · Si erreur → Action Manager",
         "Si pas d'adresse mail : uniceftmk@unicef.fr (à vérifier)"],
        "Téléphone & Lien de paiement",
        ["Fixe → confirmer et prendre le mobile (enrichit la base)",
         "Mobile → pas besoin de confirmer le mobile lui-même",
         "Lien PA : envoyer par MAIL (recommandé) ou SMS",
         "Promesse PA en ligne : confirmer J+5 après validation",
         "Si donateur rappelle : voir historique → CU autre agent → transférer ou reprendre"])
    HL(s,"Courrier postal : 3 à 4 jours max. Argumenter l'adresse : revues trimestrielles, cadeaux adhérents, reçu fiscal.",
        0.3,6.0,12.73,0.52,sz=SZ_SM)

    # Slide notifications & qualification avancée (NOUVEAU)
    s=p2.slides.add_slide(SL(p2))
    H_CONTENT(s,"Qualification Avancée & Gestion des Cas Particuliers")
    R(s,0.3,0.9,6.1,0.44,fill=TEAL_D)
    T(s,"Notification entrante — Call Blinding",0.44,0.94,5.9,0.34,sz=SZ_S,bold=True,col=WHITE)
    R(s,0.3,1.34,6.1,1.65,fill=WHITE,line=TEAL_D,lw=Pt(1.2))
    T(s,"Notification visible si le donateur rappelle.\n"
       "→ Voir l'historique du dossier\n"
       "→ CU d'un autre agent : prendre en charge EN SON NOM ou transférer\n"
       "→ Votre CU : prendre en charge directement",
       0.44,1.42,5.82,1.5,sz=SZ_SM,col=DARK)

    R(s,6.7,0.9,6.1,0.44,fill=TEAL)
    T(s,"Prélèvement assisté — Voyants",6.84,0.94,5.9,0.34,sz=SZ_S,bold=True,col=WHITE)
    R(s,6.7,1.34,6.1,1.65,fill=WHITE,line=TEAL,lw=Pt(1.2))
    T(s,"🟢 Voyant VERT = continuer / reprendre la saisie\n"
       "🔴 Voyant ROUGE = suspendre immédiatement\n\n"
       "TX d'audiolyse : 80% — 8 donateurs sur 10 finalisent s'ils restent en ligne.",
       6.84,1.42,5.82,1.5,sz=SZ_SM,col=DARK)

    R(s,0.3,3.12,12.73,0.44,fill=TEAL_M)
    T(s,"Règles de qualification — Cas spécifiques",0.44,3.16,12.4,0.34,sz=SZ_S,bold=True,col=WHITE)
    qual_cases=[("Paiement différé confirmé","Le donateur lit à voix haute le message de confirmation sur son écran → Qualifier PAIEMENT DIFFÉRÉ"),
                ("Rappel sans réponse","Si Call 2 rappelle et que le donateur ne répond pas → Qualifier PROMESSE DÉFINITIVE PEL"),
                ("PA en ligne par donateur","Envoi du lien par mail (priorité) ou SMS → Promesse à confirmer J+5 · Si erreur mail → Action Manager"),
                ("Adresse manquante","Si le donateur n'a pas d'adresse postale : utiliser uniceftmk@unicef.fr (à vérifier avec superviseur)")]
    y=3.7
    for i,(title,body) in enumerate(qual_cases):
        bg=WHITE if i%2==0 else TEAL_L
        R(s,0.3,y,3.4,0.56,fill=TEAL)
        T(s,title,0.42,y+0.1,3.2,0.38,sz=SZ_SM,bold=True,col=WHITE)
        R(s,3.72,y,9.31,0.56,fill=bg,line=TEAL,lw=Pt(0.4))
        T(s,body,3.86,y+0.1,9.1,0.4,sz=SZ_SM,col=DARK)
        y+=0.6
    HL(s,"En cas de doute sur une qualification, demander au superviseur FIDELIS avant de clore le dossier.",
        0.3,6.1,12.73,0.46)

    MERCI(p2,"BOOK J4  ·  Journée 4  ·  Objections UNICEF France × FIDELIS")
    return save(p2,f"FIM_BookJ4_Objections_{VER}.pptx")


# ══════════════════════════════════════════════════════════════════════════
#  J5 — SYNTHÈSE & CERTIFICATION
# ══════════════════════════════════════════════════════════════════════════

def build_bookJ5():
    p2=prs()
    s=p2.slides.add_slide(SL(p2))
    H_TITLE(s,"Synthèse Générale & Simulations",
        "Certification FIM · Grille 24 critères · Quiz 40 questions",
        tag="BOOK J5  ·  Journée 5  ·  Certification FIDELIS × UNICEF France × SHY-Performance")

    s=p2.slides.add_slide(SL(p2))
    H_CONTENT(s,"Programme Journée 5 — Certification FIM")
    prg=[("08h30","09h00","Révision consolidée","Révision des 5 jours de formation · Quiz flash · Météo émotionnelle"),
         ("09h00","11h00","Simulations d'appels complets","Script + objections + closing · Verrouillage chronométré · Jeux de rôle"),
         ("11h00","11h30","Débriefing collectif","Retour simulations · Célébration progrès · 3 apprentissages collectifs"),
         ("11h30","12h30","Grille 24 critères","Évaluation individuelle · Entretien formateur 5 min · Feedback personnalisé"),
         ("13h30","14h30","Quiz 40 questions","Évaluation finale théorique · Seuil 28/40 · Corrigé collectif"),
         ("14h30","15h00","Remise certifications","Attestations FIM · Plan d'action terrain · 1er appel de production demain")]
    y=0.9
    for st,en,ti,de in prg:
        R(s,0.3,y,1.38,0.72,fill=TEAL_D)
        T(s,f"{st}\n{en}",0.32,y+0.1,1.34,0.56,sz=SZ_XS+2,bold=True,col=WHITE,align=PP_ALIGN.CENTER)
        R(s,1.7,y,3.0,0.72,fill=TEAL)
        T(s,ti,1.8,y+0.16,2.82,0.44,sz=SZ_SM,bold=True,col=WHITE)
        R(s,4.72,y,8.31,0.72,fill=WHITE,line=TEAL,lw=Pt(0.8))
        T(s,de,4.84,y+0.1,8.1,0.58,sz=SZ_SM,col=DARK,wrap=True)
        y+=0.75

    s=p2.slides.add_slide(SL(p2))
    H_CONTENT(s,"Grille de Certification — 24 Critères")
    T(s,"Voici exactement ce sur quoi vous serez évalué(e) — aucune surprise :",
       0.35,0.88,12.5,0.34,sz=SZ_SM,bold=True,col=TEAL_D)
    criteria=[
        ("A1","Respect des 5 règles d'accroche","Accroche"),
        ("A2","Présentation ALLO + écoute active + identification genre","Accroche"),
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
    cxs=[0.3,0.76,7.0,9.9,10.7,11.5,12.3]; cws=[0.43,6.22,2.87,0.78,0.78,0.78,0.75]
    TH(s,["#","Ce que vous devez démontrer","Catégorie","1","2","3","4"],cxs,cws,y=1.25)
    y=1.62; rh=0.34
    cat_c={"Accroche":TEAL_D,"Script":TEAL,"Closing":TEAL_M,"Objections":TEAL_A,
           "Posture":RGBColor(0x33,0x77,0x77),"Éthique":RGBColor(0x22,0x66,0x66),"FIDELIS":RGBColor(0x11,0x55,0x55)}
    for i,(code,crit,cat) in enumerate(criteria):
        bg=WHITE if i%2==0 else TEAL_L
        R(s,cxs[0],y,cws[0],rh,fill=TEAL_D)
        T(s,code,cxs[0]+0.04,y+0.06,cws[0]-0.06,rh-0.1,sz=SZ_XS,bold=True,col=WHITE,align=PP_ALIGN.CENTER)
        R(s,cxs[1],y,cws[1],rh,fill=bg)
        T(s,crit,cxs[1]+0.08,y+0.06,cws[1]-0.12,rh-0.08,sz=SZ_XS+1,col=DARK)
        cc=cat_c.get(cat,TEAL)
        R(s,cxs[2],y,cws[2],rh,fill=cc)
        T(s,cat,cxs[2]+0.06,y+0.06,cws[2]-0.1,rh-0.08,sz=SZ_XS,bold=True,col=WHITE,align=PP_ALIGN.CENTER)
        for x,w in zip(cxs[3:],cws[3:]): R(s,x,y,w,rh,fill=WHITE,line=TEAL,lw=Pt(0.4))
        y+=rh+0.01
    R(s,cxs[0],y,sum(cws[:3])+0.04,0.34,fill=TEAL_D)
    T(s,"TOTAL  /24",cxs[0]+0.12,y+0.08,9.2,0.22,sz=SZ_XS+2,bold=True,col=WHITE)
    R(s,cxs[3],y,sum(cws[3:])+0.04,0.34,fill=TEAL_L,line=TEAL_D,lw=Pt(2))
    T(s,"__ / 24",cxs[3]+0.1,y+0.08,2.4,0.22,sz=SZ_XS+2,bold=True,col=TEAL_D,align=PP_ALIGN.CENTER)
    y+=0.38
    HL(s,"✅ Certifié(e) FIM : ≥ 17/24  ·  ⚠ Complémentaire : 12–16/24  ·  🔄 Rattrapage J+15 : < 12/24",
        0.3,y+0.04,12.73,0.36,fill=TEAL_D)

    s=p2.slides.add_slide(SL(p2))
    H_CONTENT(s,"Votre Plan d'Action — 4 Semaines Terrain")
    TWO_COL(s,"Quiz final — 40 questions en 6 parties",
        ["Partie A (Q1–Q8) : Connaissance UNICEF & de la cause",
         "Partie B (Q9–Q15) : Secteur associatif & humanitaire",
         "Partie C (Q16–Q22) : Script officiel & règles d'accroche",
         "Partie D (Q23–Q28) : KPI FIDELIS, nomenclature & production"],
        "Plan d'action — 4 semaines terrain",
        ["S1 (J+7) : 9 CU/H atteint et stable · 0 disqualification",
         "S2 (J+14) : TX Transfo ≥ cible FIDELIS",
         "S3 (J+21) : Premier PA IBAN à chaud validé",
         "S4 (J+30) : Bilan formateur SHY-Performance"])
    HL(s,"Bienvenue dans la famille SHY-Performance × FIDELIS × UNICEF France — Certifié(e) FIM.",
        0.3,6.05,12.73,0.46)
    MERCI(p2,"BOOK J5  ·  Journée 5  ·  Certification FIDELIS × UNICEF France × SHY-Performance")
    return save(p2,f"FIM_BookJ5_Synthese_Simulations_{VER}.pptx")


# ══════════════════════════════════════════════════════════════════════════
#  J6 — ATELIER PRATIQUE 1
# ══════════════════════════════════════════════════════════════════════════

def build_J6():
    p=prs()
    s=p.slides.add_slide(SL(p))
    H_TITLE(s,"Atelier Pratique — Jour 6",
        "Appels en conditions réelles simulées · Grille live · Débrief individuel",
        tag="BOOK J6  ·  Atelier 1  ·  Mise en situation téléphonique")
    T(s,"Aujourd'hui vous êtes en production simulée.\n"
        "Vous appelez, vous gérez, vous closez — en temps réel, avec retour immédiat.\n"
        "Objectif : atteindre 9 CU/H et votre premier PA validé avec verrouillage.",
        0.5,2.75,11.8,1.1,sz=SZ_B,col=DARK)

    s=p.slides.add_slide(SL(p))
    H_CONTENT(s,"Programme Atelier Pratique J6")
    prg=[("08h30","09h00","Briefing production","Rappel KPI · Répartition fichiers FIDELIS · Règles qualification · Rappel 3 phases"),
         ("09h00","11h00","Production simulée — Round 1","Appels en binômes : 1 appelle, 1 observe avec grille. Rotation toutes les 30 min."),
         ("11h00","11h30","Débrief Round 1","Analyse collective · 3 points forts / 3 axes · Écoute extraits · Verrouillage revu"),
         ("11h30","12h30","Production simulée — Round 2","Appels individuels. Formateur écoute en silence, note grille 24 critères."),
         ("13h30","14h30","Cas complexes & objections Cat.5","Simulation avec objections conflictuelles · Escalade progressive · Closing difficile"),
         ("14h30","15h30","Débrief final & scoring","Grilles individuelles restituées · Feedback personnalisé · Plan d'action J7")]
    y=0.9
    for st,en,ti,de in prg:
        R(s,0.3,y,1.38,0.7,fill=TEAL_D)
        T(s,f"{st}\n{en}",0.32,y+0.1,1.34,0.54,sz=SZ_XS+2,bold=True,col=WHITE,align=PP_ALIGN.CENTER)
        R(s,1.7,y,3.0,0.7,fill=TEAL)
        T(s,ti,1.8,y+0.16,2.82,0.44,sz=SZ_SM,bold=True,col=WHITE)
        R(s,4.72,y,8.31,0.7,fill=WHITE,line=TEAL,lw=Pt(0.8))
        T(s,de,4.84,y+0.1,8.1,0.56,sz=SZ_SM,col=DARK,wrap=True)
        y+=0.73

    s=p.slides.add_slide(SL(p))
    H_CONTENT(s,"Grille d'Observation Live J6")
    TWO_COL(s,"Ce que le formateur observe",
        ["ALLO + écoute active + identification genre",
         "Prénom AVANT nom de famille",
         "3 phases respectées : Décollage / Vol / Atterrissage",
         "Trame accroche 5 points appliquée",
         "Max 2 objections après l'appel au don"],
        "Indicateurs de performance",
        ["Verrouillage : question fermée posée ✓ / ✗",
         "Closing : 'C'est formidable' valorisé ✓ / ✗",
         "CU/H atteint (objectif ≥ 9)",
         "Qualification correcte (DON/INDÉCIS/REFUS)",
         "Durée appel dans la norme (3–5 min)"])
    HL(s,"Après chaque appel : 1 force · 1 axe d'amélioration · 1 conseil actionnable.",0.3,6.08,12.73,0.44)

    MERCI(p,"Atelier Pratique Jour 6 — Formation FIM")
    return save(p,f"FIM_BookJ6_AtelierPratique1_{VER}.pptx")


# ══════════════════════════════════════════════════════════════════════════
#  J7 — ATELIER PRATIQUE 2
# ══════════════════════════════════════════════════════════════════════════

def build_J7():
    p=prs()
    s=p.slides.add_slide(SL(p))
    H_TITLE(s,"Atelier Pratique — Jour 7",
        "Perfectionnement · Certification Finale · Prise de Poste",
        tag="BOOK J7  ·  Atelier 2  ·  Certification & Déploiement terrain")
    T(s,"Aujourd'hui c'est votre dernier jour de formation.\n"
        "Vous validez votre certification FIM finale et vous construisez\n"
        "votre plan de prise de poste opérationnel dès demain.",
        0.5,2.75,11.8,1.1,sz=SZ_B,col=DARK)

    s=p.slides.add_slide(SL(p))
    H_CONTENT(s,"Programme Atelier Perfectionnement J7")
    prg=[("08h30","09h30","Révision ciblée","Points faibles identifiés en J6 · Verrouillage · Objections Cat.5 · Règles de prélèvement"),
         ("09h30","11h00","Simulation finale certifiante","Appel complet · Grille 24 critères · Évaluateur externe"),
         ("11h00","11h30","Résultats & Débriefing","Scores certifiants · Feedback individuel · Attestation FIM remise"),
         ("11h30","12h30","Quiz 40 questions","Évaluation finale théorique · Seuil 28/40 · Corrigé collectif"),
         ("13h30","14h30","Plan de prise de poste","Construction plan J+0 à J+30 : objectifs semaine 1 à 4"),
         ("14h30","15h30","Cérémonie de clôture","Remise des certifications · Engagement public · Votre premier appel demain")]
    y=0.9
    for st,en,ti,de in prg:
        R(s,0.3,y,1.38,0.7,fill=TEAL_D)
        T(s,f"{st}\n{en}",0.32,y+0.1,1.34,0.54,sz=SZ_XS+2,bold=True,col=WHITE,align=PP_ALIGN.CENTER)
        R(s,1.7,y,3.0,0.7,fill=TEAL)
        T(s,ti,1.8,y+0.16,2.82,0.44,sz=SZ_SM,bold=True,col=WHITE)
        R(s,4.72,y,8.31,0.7,fill=WHITE,line=TEAL,lw=Pt(0.8))
        T(s,de,4.84,y+0.1,8.1,0.56,sz=SZ_SM,col=DARK,wrap=True)
        y+=0.73

    s=p.slides.add_slide(SL(p))
    H_CONTENT(s,"Analyse des Écarts — Recommandations Pédagogiques")
    TWO_COL(s,"Écarts identifiés vs standards UNICEF Paris",
        ["Durée formation : FIM = 5J · Standard UNICEF = 7J (J6+J7 ajoutés)",
         "Simulation terrain réelle : absente avant J6 — intégrée ici",
         "Quiz théorique : absent de V5 · 40 questions intégrées",
         "Feedback individuel structuré : grille 24 critères formalisée",
         "Verrouillage & closing : non enseigné avant V7 — intégré J4+J6"],
        "Recommandations intégrées V7",
        ["2 journées d'ateliers pratiques ajoutées (J6 + J7)",
         "Script accueil corrigé : ALLO + écoute active + prénom d'abord",
         "3 phases d'appel (Décollage/Vol/Atterrissage) formalisées",
         "Trame accroche 5 points + règle 60s + règle 2 objections",
         "Règles de prélèvement (date du 04) intégrées en J3"])
    HL(s,"Ces recommandations s'appuient sur les standards UNICEF Paris (protocole FIDELIS 2025) et les bonnes pratiques du secteur.",
        0.3,6.08,12.73,0.52)

    s=p.slides.add_slide(SL(p))
    H_CONTENT(s,"Votre Plan de Prise de Poste — J+0 à J+30")
    weeks=[("J+0\nDemain",TEAL_D,"Premier appel de production\n9 CU/H dès J1\nQualification FIDELIS conforme"),
           ("SEMAINE 1\nJ+7",TEAL,"TX Transfo ≥ cible FIDELIS\nVerrouillage sur chaque appel\nFeedback formateur J+7"),
           ("SEMAINE 2\nJ+14",TEAL_M,"0 objection Cat.5 non gérée\nPremier PA IBAN à chaud\nPoint superviseur FIDELIS"),
           ("SEMAINE 3\nJ+21",TEAL_A,"Autonomie complète script\nPDC mensuel en bonne voie\nAutoévaluation grille 24"),
           ("SEMAINE 4\nJ+30",RGBColor(0x33,0x77,0x77),"Bilan formateur SHY-Performance\nRévision objectifs M2\nCertification terrain confirmée")]
    x=0.3
    for tag,col,body in weeks:
        R(s,x,0.9,2.46,0.58,fill=col)
        T(s,tag,x+0.1,0.92,2.26,0.54,sz=SZ_SM,bold=True,col=WHITE,align=PP_ALIGN.CENTER)
        R(s,x,1.48,2.46,4.4,fill=WHITE,line=col,lw=Pt(1.5))
        T(s,body,x+0.14,1.58,2.18,4.18,sz=SZ_B,col=DARK,wrap=True)
        x+=2.56
    HL(s,"Votre formateur SHY-Performance vous contacte à J+15 et J+30. À J+90 : analyse KPI terrain FIDELIS × UNICEF France.",
        0.3,6.08,12.73,0.46)

    MERCI(p,"Atelier Pratique Jour 7 — Certification FIM")
    return save(p,f"FIM_BookJ7_AtelierPratique2_{VER}.pptx")


# ══════════════════════════════════════════════════════════════════════════
#  MAIN
# ══════════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    print(f"\n🎨 FIM {VER} — CORRECTIONS COMPLÈTES\n")
    files=[]
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
    for f in files: print(f"   {os.path.basename(f)}")

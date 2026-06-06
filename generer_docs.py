from docx import Document
from docx.shared import Pt, RGBColor, Cm, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ─────────────────────────────────────────
# 1. LETTRE DE MOTIVATION
# ─────────────────────────────────────────
doc = Document()

# Marges
for section in doc.sections:
    section.top_margin    = Cm(2)
    section.bottom_margin = Cm(2)
    section.left_margin   = Cm(2.5)
    section.right_margin  = Cm(2.5)

def add_para(doc, text="", bold=False, size=11, align=WD_ALIGN_PARAGRAPH.LEFT, color=None, space_before=0, space_after=6):
    p = doc.add_paragraph()
    p.alignment = align
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after  = Pt(space_after)
    if text:
        run = p.add_run(text)
        run.bold = bold
        run.font.size = Pt(size)
        if color:
            run.font.color.rgb = RGBColor(*color)
    return p

# En-tête candidat
add_para(doc, "Tamou Eljerrari", bold=True, size=14, color=(0,102,153), space_after=2)
add_para(doc, "+212 608 946 665 | madsenfatou@gmail.com | Marrakech, Maroc", size=10, space_after=12)

add_para(doc, "Rabat, le 6 juin 2026", size=11, space_after=6)
add_para(doc, "À l'attention du Service des Ressources Humaines", size=11, space_after=2)
add_para(doc, "ADM Value — Rabat Agdal", bold=True, size=11, space_after=12)

add_para(doc, "Objet : Candidature au poste de Responsable Amélioration Continue (CDI)", bold=True, size=11, space_after=16)

add_para(doc, "Madame, Monsieur,", size=11, space_after=10)

paragraphes = [
    ("Ma culture qualité ne commence pas dans un centre de contacts — elle prend racine dans l'industrie. "
     "Titulaire d'un Bac +2 scientifique, j'ai travaillé de 1997 à 2013 dans des unités de transformation "
     "de produits alimentaires au Maroc, destinées à l'export exclusif vers l'Europe. Cette expérience m'a imposé "
     "dès le départ la rigueur des référentiels ISO, la logique PDCA et le pilotage par processus. "
     "En intégrant le secteur de la relation client en 2014, j'ai transposé naturellement cette culture au CRC."),

    ("Depuis lors, j'ai exercé l'ensemble des fonctions clés du centre de contacts : conseillère commerciale, "
     "contrôleuse qualité, chargée de recrutement, formatrice senior, et aujourd'hui Responsable Qualité & Formation. "
     "Ce parcours de plus de 10 ans m'a permis de maîtriser les indicateurs NPS, IS, IR, DMT et taux de transformation, "
     "de croiser analyses qualitatives et KPIs quantitatifs, et de conduire des démarches d'amélioration continue "
     "aboutissant à la certification ISO 9001 V2015 (AFNOR, 2022). Il m'a également amenée à structurer les processus "
     "de recrutement et de formation, à former les formateurs et managers dans le cadre du cursus compétence métier, "
     "à concevoir des contenus pédagogiques en formation initiale et continue, et à piloter un démarrage opérationnel "
     "from scratch au Sénégal."),

    ("En tant que relais transversal entre le client DO, la direction et l'opérationnel (COPROD, COPIL), "
     "j'ai développé une posture de pont entre les parties prenantes que votre poste appelle précisément. "
     "Mon approche \"Root Cause\" et mon aisance aussi bien en atelier terrain qu'en présentation de livrables "
     "stratégiques complètent ce profil."),

    ("Disponible immédiatement et mobile sur Rabat, je reste à votre disposition pour un entretien."),
]

for texte in paragraphes:
    add_para(doc, texte, size=11, space_before=4, space_after=8)

add_para(doc, "Cordiales salutations,", size=11, space_before=12, space_after=4)
add_para(doc, "Tamou Eljerrari", bold=True, size=11, space_after=0)

doc.save("/home/user/fatou/Lettre_Motivation_Tamou_Eljerrari.docx")
print("Lettre sauvegardée.")

# ─────────────────────────────────────────
# 2. CV MIS À JOUR
# ─────────────────────────────────────────
cv = Document()

for section in cv.sections:
    section.top_margin    = Cm(1.5)
    section.bottom_margin = Cm(1.5)
    section.left_margin   = Cm(2)
    section.right_margin  = Cm(2)

def titre_section(doc, texte):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after  = Pt(4)
    run = p.add_run(texte.upper())
    run.bold = True
    run.font.size = Pt(12)
    run.font.color.rgb = RGBColor(0, 102, 153)
    # ligne sous le titre
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '6')
    bottom.set(qn('w:space'), '1')
    bottom.set(qn('w:color'), '006699')
    pBdr.append(bottom)
    pPr.append(pBdr)

def add_cv_para(doc, text, bold=False, size=10.5, bullet=False, space_after=3):
    if bullet:
        p = doc.add_paragraph(style='List Bullet')
    else:
        p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(space_after)
    run = p.add_run(text)
    run.bold = bold
    run.font.size = Pt(size)
    return p

# Nom & titre
p = cv.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(2)
r = p.add_run("TAMOU ELJERRARI")
r.bold = True; r.font.size = Pt(18); r.font.color.rgb = RGBColor(0,102,153)

p2 = cv.add_paragraph()
p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
p2.paragraph_format.space_after = Pt(4)
r2 = p2.add_run("Responsable Qualité & Formation | Amélioration Continue")
r2.font.size = Pt(11); r2.italic = True

# Contacts
p3 = cv.add_paragraph()
p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
p3.paragraph_format.space_after = Pt(10)
r3 = p3.add_run("+212 608 946 665 / +212 603 993 201  |  madsenfatou@gmail.com  |  Marrakech, Maroc  |  Disponibilité : Immédiate")
r3.font.size = Pt(9.5)

# Profil
titre_section(cv, "Profil")
add_cv_para(cv, (
    "Professionnelle polyvalente avec plus de 25 ans d'expérience combinant industrie agroalimentaire export "
    "(1997-2013) et centres de relation client (2014 à ce jour). Habilitée sur le métier client DO, "
    "certifiée ISO 9001 V2015 (AFNOR), ingénieure de formation et coach qualité terrain. "
    "Expertise reconnue en amélioration continue, pilotage qualité, management transversal et ingénierie pédagogique."
), size=10.5, space_after=6)

# Expérience
titre_section(cv, "Expérience Professionnelle")

experiences = [
    ("2023 – À ce jour", "Responsable Qualité & Formation — Shyperformance", [
        "Mise en place et suivi de la certification ISO 9001 V2015 (délivrée par AFNOR)",
        "Ingénierie de formation et conception de contenus pédagogiques (formation initiale et continue)",
        "Habilitation et formation des formateurs et managers (cursus compétence métier)",
        "Structuration du processus de recrutement et de formation",
        "Relais de communication transversal entre le client DO, la direction et l'opérationnel",
        "Animation des instances de pilotage : COPROD et COPIL",
        "Participation au démarrage opérationnel from scratch au Sénégal",
    ]),
    ("2021 – 2022", "Formatrice Senior — Shyperformance (Énergie & Téléphonie)", [
        "Encadrement des nouvelles recrues et accompagnement des équipes",
        "Animation de formations en relation client et gestion des compétences",
    ]),
    ("2020 – 2021", "Formatrice Senior — Webhelp", []),
    ("2018 – 2020", "Formatrice SFR Sortant & Entraînement Service Client — Webhelp", []),
    ("2016 – 2017", "Chargée de Recrutement — Webhelp", [
        "Sélection et intégration des nouveaux collaborateurs",
        "Élaboration de tests et d'entretiens de validation",
    ]),
    ("2015 – 2016", "Contrôleuse Qualité — Webhelp", [
        "Évaluation des performances des conseillers",
        "Analyse des écarts et proposition d'actions correctives (approche Root Cause / PDCA)",
    ]),
    ("2014 – 2015", "Conseillère Commerciale TLV SFR — Webhelp (9 mois)", []),
    ("1997 – 2013", "Production & Qualité — Industrie agroalimentaire, Maroc (export exclusif Europe)", [
        "Pilotage de la qualité et conformité aux normes européennes d'exportation",
        "Maîtrise des référentiels ISO, traçabilité et audits qualité",
        "Logique PDCA et amélioration continue en environnement industriel",
    ]),
]

for periode, poste, bullets in experiences:
    p = cv.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(1)
    r1 = p.add_run(periode + " — ")
    r1.bold = True; r1.font.size = Pt(10.5); r1.font.color.rgb = RGBColor(0,102,153)
    r2 = p.add_run(poste)
    r2.bold = True; r2.font.size = Pt(10.5)
    for b in bullets:
        add_cv_para(cv, b, bullet=True, size=10, space_after=2)

# Compétences
titre_section(cv, "Compétences Clés")
competences = [
    "Amélioration continue & pilotage qualité (ISO 9001, PDCA, Root Cause Analysis)",
    "Ingénierie de formation & conception de contenus pédagogiques",
    "Habilitation métier client DO — Formation des formateurs et managers",
    "Indicateurs CRC : NPS, IS, IR, DMT, Taux de vente — tableaux de bord croisés Quali/Quanti",
    "Audit interne & auto-contrôle",
    "Management, leadership et cohésion d'équipe",
    "Relais transversal client DO / Direction / Opérationnel (COPROD, COPIL)",
    "Communication d'impact : animation d'ateliers & restitution direction",
    "Soft Skills : gestion du stress, intelligence émotionnelle, gestion des conflits",
]
for c in competences:
    add_cv_para(cv, c, bullet=True, size=10, space_after=2)

# Réalisations
titre_section(cv, "Réalisations")
realisations = [
    "Certification ISO 9001 V2015 délivrée par AFNOR (2022) — financée par l'entreprise",
    "Habilitation à former les formateurs et managers (cursus compétence métier)",
    "Habilitation sur le métier client DO",
    "Démarrage opérationnel from scratch au Sénégal",
    "Mise en place du processus complet de recrutement et de formation",
    "Propositions de valeur et capitalisation des 2R : Résultats & Rétention",
]
for r in realisations:
    add_cv_para(cv, r, bullet=True, size=10, space_after=2)

# Formation
titre_section(cv, "Formation")
formations = [
    "Bac +2 Scientifique",
    "Formation en ingénierie de formation et élaboration de contenus pédagogiques",
    "Expertise en animation Soft Skills : gestion du stress, intelligence émotionnelle, excellence relationnelle, gestion des conflits",
]
for f in formations:
    add_cv_para(cv, f, bullet=True, size=10, space_after=2)

cv.save("/home/user/fatou/CV_Tamou_Eljerrari_MisAJour.docx")
print("CV sauvegardé.")

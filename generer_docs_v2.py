from docx import Document
from docx.shared import Pt, RGBColor, Cm, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.enum.table import WD_TABLE_ALIGNMENT

# ══════════════════════════════════════════════
# COULEURS
# ══════════════════════════════════════════════
BLEU      = RGBColor(0, 82, 136)
BLEU_CLAIR= RGBColor(0, 140, 186)
GRIS      = RGBColor(90, 90, 90)
BLANC     = RGBColor(255, 255, 255)
NOIR      = RGBColor(30, 30, 30)

def set_cell_bg(cell, hex_color):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color)
    tcPr.append(shd)

def set_cell_border(cell, **kwargs):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    for edge in ('top','left','bottom','right','insideH','insideV'):
        tag = OxmlElement(f'w:{edge}')
        tag.set(qn('w:val'), 'none')
        tcBorders.append(tag)
    tcPr.append(tcBorders)

def no_space(para):
    para.paragraph_format.space_before = Pt(0)
    para.paragraph_format.space_after  = Pt(0)

# ══════════════════════════════════════════════
# CV — STRUCTURE 2 COLONNES
# ══════════════════════════════════════════════
cv = Document()
for s in cv.sections:
    s.top_margin    = Cm(0)
    s.bottom_margin = Cm(1)
    s.left_margin   = Cm(0)
    s.right_margin  = Cm(0)

# ── BANDEAU EN-TÊTE ──────────────────────────
header_tbl = cv.add_table(rows=1, cols=1)
header_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
header_tbl.style = 'Table Grid'
hcell = header_tbl.rows[0].cells[0]
set_cell_bg(hcell, '005288')
set_cell_border(hcell)
hcell.width = Inches(8.27)

p = hcell.add_paragraph()
no_space(p)
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(14)
r = p.add_run("TAMOU ELJERRARI")
r.bold = True; r.font.size = Pt(22); r.font.color.rgb = BLANC

p2 = hcell.add_paragraph()
no_space(p2)
p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
p2.paragraph_format.space_after = Pt(6)
r2 = p2.add_run("Responsable Qualité, Formation & Amélioration Continue")
r2.font.size = Pt(12); r2.italic = True; r2.font.color.rgb = RGBColor(180,220,255)

p3 = hcell.add_paragraph()
no_space(p3)
p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
p3.paragraph_format.space_after = Pt(14)
r3 = p3.add_run("📍 Marrakech, Maroc  |  📞 +212 608 946 665  |  ✉ madsenfatou@gmail.com  |  Disponible immédiatement")
r3.font.size = Pt(9); r3.font.color.rgb = RGBColor(200,230,255)

cv.add_paragraph()  # espace

# ── CORPS : 2 COLONNES ───────────────────────
body = cv.add_table(rows=1, cols=2)
body.style = 'Table Grid'
body.alignment = WD_TABLE_ALIGNMENT.CENTER

left  = body.rows[0].cells[0]
right = body.rows[0].cells[1]
set_cell_bg(left,  'EAF4FB')
set_cell_bg(right, 'FFFFFF')
set_cell_border(left)
set_cell_border(right)
left.width  = Cm(6.5)
right.width = Cm(13.5)

def left_titre(cell, text):
    p = cell.add_paragraph()
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after  = Pt(4)
    p.paragraph_format.left_indent  = Cm(0.3)
    r = p.add_run(text.upper())
    r.bold = True; r.font.size = Pt(9.5); r.font.color.rgb = BLEU

def left_item(cell, text, bold=False):
    p = cell.add_paragraph()
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.space_after  = Pt(3)
    p.paragraph_format.left_indent  = Cm(0.3)
    r = p.add_run(text)
    r.font.size = Pt(9); r.bold = bold; r.font.color.rgb = NOIR

def right_titre(cell, text):
    p = cell.add_paragraph()
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.space_after  = Pt(3)
    p.paragraph_format.left_indent  = Cm(0.4)
    r = p.add_run(text.upper())
    r.bold = True; r.font.size = Pt(10.5); r.font.color.rgb = BLEU
    # ligne décorative
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bot = OxmlElement('w:bottom')
    bot.set(qn('w:val'), 'single'); bot.set(qn('w:sz'), '6')
    bot.set(qn('w:space'), '1'); bot.set(qn('w:color'), '005288')
    pBdr.append(bot); pPr.append(pBdr)

def right_exp(cell, periode, poste, bullets):
    p = cell.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after  = Pt(1)
    p.paragraph_format.left_indent  = Cm(0.4)
    r1 = p.add_run(periode)
    r1.bold = True; r1.font.size = Pt(9); r1.font.color.rgb = BLEU_CLAIR
    p2 = cell.add_paragraph()
    p2.paragraph_format.space_before = Pt(0)
    p2.paragraph_format.space_after  = Pt(2)
    p2.paragraph_format.left_indent  = Cm(0.4)
    r2 = p2.add_run(poste)
    r2.bold = True; r2.font.size = Pt(10)
    for b in bullets:
        pb = cell.add_paragraph(style='List Bullet')
        pb.paragraph_format.space_before = Pt(1)
        pb.paragraph_format.space_after  = Pt(1)
        pb.paragraph_format.left_indent  = Cm(0.8)
        rb = pb.add_run(b)
        rb.font.size = Pt(9); rb.font.color.rgb = GRIS

# ─── COLONNE GAUCHE ───
left_titre(left, "Compétences Clés")
comps = [
    "Amélioration continue (ISO 9001, PDCA)",
    "Analyse Root Cause des écarts",
    "Pilotage qualité & audit interne",
    "KPIs : NPS, IS, IR, DMT, Taux de vente",
    "Tableaux de bord croisés Quali/Quanti",
    "Ingénierie de formation",
    "Conception de contenus pédagogiques",
    "Formation des formateurs & managers",
    "Habilitation métier client DO",
    "Recrutement & intégration",
    "Management & leadership d'équipe",
    "Animation COPROD / COPIL",
    "Communication transversale d'impact",
]
for c in comps:
    left_item(left, f"▸  {c}")

left_titre(left, "Réalisations")
reals = [
    "Certification ISO 9001 V2015\n(AFNOR, 2022)",
    "Habilitation formation des\nformateurs et managers",
    "Habilitation métier client DO",
    "Set-up opérationnel au Sénégal",
    "Mise en place du processus\nrecrutement & formation",
    "Capitalisation des 2R :\nRésultats & Rétention",
]
for r in reals:
    left_item(left, f"✔  {r}")

left_titre(left, "Formation")
left_item(left, "Bac +2 Scientifique", bold=True)
left_item(left, "Ingénierie de formation &\nconception pédagogique")
left_item(left, "Soft Skills : gestion du stress,\nintelligence émotionnelle,\nexcellence relationnelle,\ngestion des conflits")

left_titre(left, "Langues")
left_item(left, "Français — Courant")
left_item(left, "Arabe — Langue maternelle")

# ─── COLONNE DROITE ───
right_titre(right, "Profil")
p_profil = right.add_paragraph()
p_profil.paragraph_format.space_before = Pt(4)
p_profil.paragraph_format.space_after  = Pt(4)
p_profil.paragraph_format.left_indent  = Cm(0.4)
r_profil = p_profil.add_run(
    "Professionnelle avec plus de 25 ans d'expérience alliant rigueur industrielle et excellence "
    "en relation client. Après 16 années dans l'industrie agroalimentaire orientée export Europe (1997–2013), "
    "j'ai transposé ma culture qualité ISO et PDCA au secteur des centres de contacts depuis 2014. "
    "Certifiée ISO 9001 V2015 (AFNOR), habilitée sur le métier client DO et sur la formation des formateurs, "
    "je pilote l'amélioration continue de bout en bout : voix du client, data, terrain et livrables stratégiques."
)
r_profil.font.size = Pt(9.5); r_profil.font.color.rgb = GRIS

right_titre(right, "Expérience Professionnelle")

exps = [
    ("2023 – À ce jour", "Responsable Qualité & Formation — Shyperformance", [
        "Pilotage et obtention de la certification ISO 9001 V2015 (AFNOR)",
        "Ingénierie de formation et conception de contenus pédagogiques (initiale & continue)",
        "Formation et habilitation des formateurs et managers (cursus compétence métier)",
        "Habilitation sur le métier client DO",
        "Structuration du processus de recrutement et de formation",
        "Relais transversal entre le client DO, la direction et l'opérationnel",
        "Animation des instances COPROD et COPIL",
        "Démarrage opérationnel from scratch au Sénégal",
    ]),
    ("2021 – 2022", "Formatrice Senior — Shyperformance (Énergie & Téléphonie)", [
        "Encadrement des recrues et accompagnement des équipes terrain",
        "Animation de formations relation client et gestion des compétences",
    ]),
    ("2020 – 2021", "Formatrice Senior — Webhelp", [
        "Formation continue des équipes et transfert des bonnes pratiques",
    ]),
    ("2018 – 2020", "Formatrice SFR Sortant & Entraînement Service Client — Webhelp", [
        "Animation des formations sur les process SFR et la relation client sortante",
    ]),
    ("2016 – 2017", "Chargée de Recrutement — Webhelp", [
        "Sélection, intégration et validation des nouveaux collaborateurs",
    ]),
    ("2015 – 2016", "Contrôleuse Qualité — Webhelp", [
        "Évaluation des conseillers, analyse des écarts et plans d'actions correctifs (Root Cause / PDCA)",
    ]),
    ("2014 – 2015", "Conseillère Commerciale TLV SFR — Webhelp", [
        "Gestion de la relation client et performance commerciale",
    ]),
    ("1997 – 2013", "Responsable Production & Qualité — Industrie agroalimentaire, Maroc (export Europe)", [
        "Pilotage qualité et conformité aux standards européens d'exportation",
        "Maîtrise des référentiels ISO, traçabilité et audits internes",
        "Démarche d'amélioration continue en environnement industriel (PDCA)",
    ]),
]
for periode, poste, bullets in exps:
    right_exp(right, periode, poste, bullets)

cv.save("/home/user/fatou/CV_Tamou_Eljerrari_V2.docx")
print("CV V2 sauvegardé.")

# ══════════════════════════════════════════════
# LETTRE DE MOTIVATION V2
# ══════════════════════════════════════════════
lm = Document()
for s in lm.sections:
    s.top_margin    = Cm(0)
    s.bottom_margin = Cm(1.5)
    s.left_margin   = Cm(0)
    s.right_margin  = Cm(0)

# Bandeau haut
lt = lm.add_table(rows=1, cols=1)
lt.style = 'Table Grid'
lc = lt.rows[0].cells[0]
set_cell_bg(lc, '005288')
set_cell_border(lc)

p = lc.add_paragraph()
no_space(p); p.alignment = WD_ALIGN_PARAGRAPH.LEFT
p.paragraph_format.space_before = Pt(16)
p.paragraph_format.left_indent  = Cm(1.5)
r = p.add_run("TAMOU ELJERRARI")
r.bold = True; r.font.size = Pt(18); r.font.color.rgb = BLANC

p2 = lc.add_paragraph()
no_space(p2); p2.alignment = WD_ALIGN_PARAGRAPH.LEFT
p2.paragraph_format.left_indent  = Cm(1.5)
p2.paragraph_format.space_after  = Pt(6)
r2 = p2.add_run("Responsable Qualité, Formation & Amélioration Continue")
r2.italic = True; r2.font.size = Pt(11); r2.font.color.rgb = RGBColor(180,220,255)

p3 = lc.add_paragraph()
no_space(p3); p3.alignment = WD_ALIGN_PARAGRAPH.LEFT
p3.paragraph_format.left_indent  = Cm(1.5)
p3.paragraph_format.space_after  = Pt(14)
r3 = p3.add_run("+212 608 946 665  |  madsenfatou@gmail.com  |  Marrakech, Maroc  |  Disponible immédiatement")
r3.font.size = Pt(9); r3.font.color.rgb = RGBColor(200,230,255)

# Corps lettre
def lm_para(doc, text, bold=False, size=10.5, space_before=6, space_after=8,
            align=WD_ALIGN_PARAGRAPH.JUSTIFY, color=NOIR, italic=False, indent=Cm(1.5)):
    p = doc.add_paragraph()
    p.alignment = align
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after  = Pt(space_after)
    p.paragraph_format.left_indent  = indent
    p.paragraph_format.right_indent = Cm(1.5)
    if text:
        r = p.add_run(text)
        r.bold = bold; r.italic = italic
        r.font.size = Pt(size)
        r.font.color.rgb = color
    return p

lm_para(lm, "Rabat, le 6 juin 2026", size=10, space_before=16, space_after=4,
        align=WD_ALIGN_PARAGRAPH.RIGHT)
lm_para(lm, "À l'attention du Service des Ressources Humaines", size=10.5, space_before=2, space_after=2)
lm_para(lm, "ADM Value — Rabat Agdal", size=10.5, bold=True, space_before=0, space_after=14)
lm_para(lm, "Objet : Candidature au poste de Responsable Amélioration Continue (CDI)",
        bold=True, size=11, color=BLEU, space_before=0, space_after=16)
lm_para(lm, "Madame, Monsieur,", size=10.5, space_before=0, space_after=12)

contenu = [
    ("Ma démarche qualité ne s'est pas construite dans un bureau — elle est née sur le terrain industriel. "
     "Titulaire d'un Bac +2 scientifique, j'ai exercé de 1997 à 2013 dans des unités de transformation "
     "agroalimentaire au Maroc, dont la production était intégralement destinée à l'export vers l'Europe. "
     "Cette expérience m'a immergée dans la rigueur des référentiels ISO, la logique PDCA et le pilotage "
     "par processus — des réflexes que j'ai ensuite transposés et adaptés au Centre de Relation Client "
     "dès mon entrée dans le secteur en 2014."),

    ("En douze ans de centres de contacts, j'ai occupé chaque maillon de la chaîne de performance : "
     "conseillère commerciale, contrôleuse qualité, chargée de recrutement, formatrice senior, et aujourd'hui "
     "Responsable Qualité & Formation. Ce parcours m'a rendue opérationnelle sur l'ensemble des dimensions "
     "que votre poste requiert : maîtrise des indicateurs NPS, IS, IR, DMT et taux de transformation ; "
     "croisement des analyses qualitatives et quantitatives pour identifier les leviers à fort impact ; "
     "certification ISO 9001 V2015 (AFNOR, 2022) ; ingénierie de formation et conception de contenus "
     "pédagogiques en formation initiale et continue ; habilitation sur le métier client DO et formation "
     "des formateurs et managers dans le cadre du cursus compétence métier ; structuration du processus "
     "de recrutement et de formation ; et pilotage d'un démarrage opérationnel from scratch au Sénégal."),

    ("En tant que relais transversal entre le client DO, la direction et l'opérationnel — à travers "
     "l'animation des instances COPROD et COPIL — j'ai développé l'aisance nécessaire pour porter "
     "une analyse terrain en atelier de conseillers comme pour restituer des livrables stratégiques "
     "à la direction. Mon approche systématique d'analyse en « Root Cause » me permet d'aller au-delà "
     "des symptômes pour agir sur les causes réelles des écarts de performance."),

    ("Créer et déployer un poste stratégique d'amélioration continue au sein du groupe Tessi est "
     "exactement le type de mission pour lequel mon parcours m'a préparée. Je suis disponible "
     "immédiatement et mobile sur Rabat, et serai ravie de vous en convaincre lors d'un entretien."),
]

for i, texte in enumerate(contenu):
    lm_para(lm, texte, space_before=4, space_after=10)

lm_para(lm, "Dans l'attente de votre retour, je vous adresse mes cordiales salutations.", 
        space_before=12, space_after=6)
lm_para(lm, "Tamou Eljerrari", bold=True, size=11, space_before=4, space_after=0)

lm.save("/home/user/fatou/Lettre_Motivation_Tamou_Eljerrari_V2.docx")
print("Lettre V2 sauvegardée.")

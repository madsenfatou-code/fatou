from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import cm, mm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_JUSTIFY, TA_RIGHT
from reportlab.pdfgen import canvas
from reportlab.platypus import BaseDocTemplate, Frame, PageTemplate
from PIL import Image as PILImage

W, H = A4
BLEU       = colors.HexColor('#005288')
BLEU_CLAIR = colors.HexColor('#008CBA')
GRIS       = colors.HexColor('#5A5A5A')
GRIS_CLAIR = colors.HexColor('#EAF4FB')
BLANC      = colors.white
NOIR       = colors.HexColor('#1E1E1E')

# ── Styles ──────────────────────────────────
def s(name, **kw):
    return ParagraphStyle(name, **kw)

st_nom     = s('nom',     fontSize=20, textColor=BLANC,  leading=24, alignment=TA_CENTER, fontName='Helvetica-Bold')
st_titre   = s('titre',   fontSize=11, textColor=colors.HexColor('#B4DCF0'), leading=14, alignment=TA_CENTER, fontName='Helvetica-Oblique')
st_contact = s('contact', fontSize=8.5,textColor=colors.HexColor('#C8E6FF'), leading=11, alignment=TA_CENTER)

st_sec_l   = s('sec_l',  fontSize=8.5, textColor=BLEU, leading=11, fontName='Helvetica-Bold', spaceAfter=3, spaceBefore=10)
st_item_l  = s('item_l', fontSize=8.5, textColor=NOIR, leading=12, leftIndent=4, spaceAfter=2)

st_profil  = s('profil', fontSize=9.5, textColor=GRIS, leading=14, alignment=TA_JUSTIFY, spaceAfter=4)
st_sec_r   = s('sec_r',  fontSize=10.5, textColor=BLEU, leading=13, fontName='Helvetica-Bold', spaceBefore=12, spaceAfter=4)
st_periode = s('periode',fontSize=8.5, textColor=BLEU_CLAIR, leading=11, fontName='Helvetica-Bold', spaceBefore=6)
st_poste   = s('poste',  fontSize=10,  textColor=NOIR, leading=13, fontName='Helvetica-Bold', spaceAfter=2)
st_bullet  = s('bullet', fontSize=8.5, textColor=GRIS, leading=12, leftIndent=10, spaceAfter=1,
               bulletIndent=2, bulletText='•')

def hr():
    return Table([['']], colWidths=[13*cm],
                 style=TableStyle([('LINEABOVE',(0,0),(0,0),0.8,BLEU),
                                   ('TOPPADDING',(0,0),(-1,-1),0),
                                   ('BOTTOMPADDING',(0,0),(-1,-1),2)]))

# ── Contenu colonne gauche ───────────────────
photo = PILImage.open("/home/user/fatou/photo_tamou_crop.png")
photo.save("/home/user/fatou/photo_cv.png")
img = Image("/home/user/fatou/photo_cv.png", width=3.6*cm, height=3.6*cm)

left_items = [
    img,
    Spacer(1, 6),
    Paragraph("COMPÉTENCES CLÉS", st_sec_l),
    *[Paragraph(f"▸ {c}", st_item_l) for c in [
        "Amélioration continue (ISO 9001, PDCA)",
        "Analyse Root Cause des écarts",
        "Pilotage qualité & audit interne",
        "KPIs : NPS, IS, IR, DMT,<br/>Taux de vente",
        "Tableaux de bord Quali/Quanti",
        "Ingénierie de formation",
        "Conception contenus pédagogiques",
        "Formation formateurs & managers",
        "Habilitation métier client DO",
        "Recrutement & intégration",
        "Management & leadership",
        "Animation COPROD / COPIL",
        "Communication transversale",
    ]],
    Spacer(1, 4),
    Paragraph("RÉALISATIONS", st_sec_l),
    *[Paragraph(f"✔ {r}", st_item_l) for r in [
        "Certification ISO 9001 V2015<br/>(AFNOR, 2022)",
        "Habilitation formation des<br/>formateurs et managers",
        "Habilitation métier client DO",
        "Set-up opérationnel au Sénégal",
        "Mise en place processus<br/>recrutement & formation",
        "Capitalisation des 2R :<br/>Résultats & Rétention",
    ]],
    Spacer(1, 4),
    Paragraph("FORMATION", st_sec_l),
    *[Paragraph(f"▸ {f}", st_item_l) for f in [
        "Bac +2 Scientifique",
        "Ingénierie de formation &<br/>conception pédagogique",
        "Soft Skills : gestion du stress,<br/>intelligence émotionnelle,<br/>excellence relationnelle,<br/>gestion des conflits",
    ]],
    Spacer(1, 4),
    Paragraph("LANGUES", st_sec_l),
    Paragraph("▸ Français — Courant", st_item_l),
    Paragraph("▸ Arabe — Langue maternelle", st_item_l),
    Spacer(1, 4),
    Paragraph("VIDÉO DE PRÉSENTATION", st_sec_l),
    Image("/home/user/fatou/qrcode_video.png", width=3.2*cm, height=3.2*cm),
    Paragraph('<link href="https://www.youtube.com/watch?v=cNZ97tWrsNU">'
              '<font color="#005288"><u>▸ Voir ma présentation</u></font></link>', st_item_l),
]

# ── Contenu colonne droite ───────────────────
exps = [
    ("2023 – À ce jour", "Responsable Qualité & Formation — Shyperformance", [
        "Pilotage et obtention de la certification ISO 9001 V2015 (AFNOR)",
        "Ingénierie de formation et conception de contenus pédagogiques",
        "Formation et habilitation des formateurs et managers",
        "Habilitation sur le métier client DO",
        "Structuration du processus de recrutement et de formation",
        "Relais transversal : client DO ↔ direction ↔ opérationnel",
        "Animation des instances COPROD et COPIL",
        "Démarrage opérationnel from scratch au Sénégal",
    ]),
    ("2021 – 2022", "Formatrice Senior — Shyperformance (Énergie & Téléphonie)", [
        "Encadrement des recrues et accompagnement terrain",
        "Animation de formations relation client et compétences",
    ]),
    ("2020 – 2021", "Formatrice Senior — Webhelp", [
        "Formation continue et transfert des bonnes pratiques",
    ]),
    ("2018 – 2020", "Formatrice SFR Sortant & Service Client — Webhelp", [
        "Animation formations process SFR et relation client sortante",
    ]),
    ("2016 – 2017", "Chargée de Recrutement — Webhelp", [
        "Sélection, intégration et validation des collaborateurs",
    ]),
    ("2015 – 2016", "Contrôleuse Qualité — Webhelp", [
        "Évaluation des conseillers, analyse des écarts, plans d'actions (Root Cause / PDCA)",
    ]),
    ("2014 – 2015", "Conseillère Commerciale TLV SFR — Webhelp", [
        "Relation client et performance commerciale",
    ]),
    ("1997 – 2013", "Responsable Production & Qualité — Industrie agroalimentaire, Maroc (export Europe)", [
        "Pilotage qualité et conformité aux standards européens d'exportation",
        "Maîtrise des référentiels ISO, traçabilité et audits internes",
        "Amélioration continue en environnement industriel (PDCA)",
    ]),
]

right_items = [
    Paragraph("PROFIL", st_sec_r),
    hr(),
    Paragraph(
        "Professionnelle avec plus de 25 ans d'expérience alliant rigueur industrielle et excellence "
        "en relation client. Après 16 années dans l'industrie agroalimentaire orientée export Europe (1997–2013), "
        "j'ai transposé ma culture qualité ISO et PDCA au secteur des centres de contacts depuis 2014. "
        "Certifiée ISO 9001 V2015 (AFNOR), habilitée métier client DO et formation des formateurs, "
        "je pilote l'amélioration continue de bout en bout : voix du client, data, terrain et livrables stratégiques.",
        st_profil),
    Paragraph("EXPÉRIENCE PROFESSIONNELLE", st_sec_r),
    hr(),
]
for periode, poste, bullets in exps:
    right_items.append(Paragraph(periode, st_periode))
    right_items.append(Paragraph(poste, st_poste))
    for b in bullets:
        right_items.append(Paragraph(b, st_bullet))
    right_items.append(Spacer(1, 3))

# ── Assemblage PDF ───────────────────────────
LEFT_W  = 5.5*cm
RIGHT_W = 13.5*cm
PAD     = 0.3*cm

def build_col(items, width):
    data = [[item] for item in items]
    t = Table(data, colWidths=[width])
    t.setStyle(TableStyle([
        ('LEFTPADDING',  (0,0),(-1,-1), 4),
        ('RIGHTPADDING', (0,0),(-1,-1), 4),
        ('TOPPADDING',   (0,0),(-1,-1), 1),
        ('BOTTOMPADDING',(0,0),(-1,-1), 1),
        ('VALIGN',       (0,0),(-1,-1),'TOP'),
    ]))
    return t

# Bandeau titre
header_data = [[
    Paragraph("TAMOU ELJERRARI", st_nom),
]]
header = Table(header_data, colWidths=[W])
header.setStyle(TableStyle([
    ('BACKGROUND', (0,0),(-1,-1), BLEU),
    ('TOPPADDING', (0,0),(-1,-1), 14),
    ('BOTTOMPADDING',(0,0),(-1,-1), 4),
]))

sub_data = [[Paragraph("Responsable Qualité, Formation & Amélioration Continue", st_titre)]]
sub = Table(sub_data, colWidths=[W])
sub.setStyle(TableStyle([
    ('BACKGROUND',(0,0),(-1,-1), BLEU),
    ('TOPPADDING',(0,0),(-1,-1), 0),
    ('BOTTOMPADDING',(0,0),(-1,-1), 4),
]))

cont_data = [[Paragraph("📍 Marrakech, Maroc  |  📞 +212 608 946 665  |  ✉ madsenfatou@gmail.com  |  Disponible immédiatement", st_contact)]]
cont = Table(cont_data, colWidths=[W])
cont.setStyle(TableStyle([
    ('BACKGROUND',(0,0),(-1,-1), BLEU),
    ('TOPPADDING',(0,0),(-1,-1), 0),
    ('BOTTOMPADDING',(0,0),(-1,-1), 12),
]))

# Corps 2 colonnes
left_col  = build_col(left_items,  LEFT_W)
right_col = build_col(right_items, RIGHT_W)

body = Table([[left_col, right_col]], colWidths=[LEFT_W, RIGHT_W])
body.setStyle(TableStyle([
    ('BACKGROUND',    (0,0),(0,-1), GRIS_CLAIR),
    ('BACKGROUND',    (1,0),(1,-1), BLANC),
    ('VALIGN',        (0,0),(-1,-1), 'TOP'),
    ('LEFTPADDING',   (0,0),(-1,-1), 6),
    ('RIGHTPADDING',  (0,0),(-1,-1), 6),
    ('TOPPADDING',    (0,0),(-1,-1), 4),
    ('BOTTOMPADDING', (0,0),(-1,-1), 4),
]))

doc = SimpleDocTemplate(
    "/home/user/fatou/CV_Tamou_Eljerrari.pdf",
    pagesize=A4,
    leftMargin=0, rightMargin=0,
    topMargin=0, bottomMargin=1*cm
)
doc.build([header, sub, cont, body])
print("PDF généré.")

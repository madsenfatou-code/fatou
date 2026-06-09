from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_JUSTIFY

W, H = A4
BLEU       = colors.HexColor('#005288')
BLEU_CLAIR = colors.HexColor('#008CBA')
VERT       = colors.HexColor('#1A7A4A')
ORANGE     = colors.HexColor('#D4600A')
GRIS_FOND  = colors.HexColor('#F0F7FB')
GRIS_FONCE = colors.HexColor('#5A5A5A')
BLANC      = colors.white
NOIR       = colors.HexColor('#1E1E1E')

def s(name, **kw):
    return ParagraphStyle(name, **kw)

# Styles
st_titre    = s('titre',   fontSize=15, textColor=BLANC, leading=20, alignment=TA_CENTER, fontName='Helvetica-Bold')
st_sous     = s('sous',    fontSize=9,  textColor=colors.HexColor('#C8E6FF'), leading=12, alignment=TA_CENTER, fontName='Helvetica-Oblique')
st_section  = s('section', fontSize=10, textColor=BLANC, leading=13, fontName='Helvetica-Bold')
st_body     = s('body',    fontSize=9,  textColor=NOIR,  leading=13, alignment=TA_JUSTIFY)
st_pitch    = s('pitch',   fontSize=9.5,textColor=NOIR,  leading=14, alignment=TA_JUSTIFY,
                leftIndent=6, rightIndent=6)
st_kw       = s('kw',      fontSize=9,  textColor=BLEU,  leading=13, fontName='Helvetica-Bold')
st_check    = s('check',   fontSize=9,  textColor=NOIR,  leading=14)
st_label    = s('label',   fontSize=8.5,textColor=BLEU_CLAIR, leading=11, fontName='Helvetica-Bold')
st_star     = s('star',    fontSize=8.5,textColor=GRIS_FONCE, leading=12, leftIndent=8)
st_cloture  = s('cloture', fontSize=9.5,textColor=VERT,  leading=14, alignment=TA_JUSTIFY,
                fontName='Helvetica-Oblique', leftIndent=6, rightIndent=6)

def section_header(texte, couleur=BLEU):
    t = Table([[Paragraph(texte, st_section)]], colWidths=[W - 2*cm])
    t.setStyle(TableStyle([
        ('BACKGROUND',    (0,0),(-1,-1), couleur),
        ('TOPPADDING',    (0,0),(-1,-1), 5),
        ('BOTTOMPADDING', (0,0),(-1,-1), 5),
        ('LEFTPADDING',   (0,0),(-1,-1), 8),
        ('ROUNDEDCORNERS',[4]),
    ]))
    return t

def box(content_rows, bg=GRIS_FOND, border_color=BLEU_CLAIR):
    t = Table(content_rows, colWidths=[W - 2*cm])
    t.setStyle(TableStyle([
        ('BACKGROUND',    (0,0),(-1,-1), bg),
        ('TOPPADDING',    (0,0),(-1,-1), 4),
        ('BOTTOMPADDING', (0,0),(-1,-1), 4),
        ('LEFTPADDING',   (0,0),(-1,-1), 10),
        ('RIGHTPADDING',  (0,0),(-1,-1), 10),
        ('LINEAFTER',     (0,0),(0,-1),  2, border_color),
        ('LINEBEFORE',    (0,0),(0,-1),  3, border_color),
    ]))
    return t

def two_col(left, right, lw=9*cm, rw=None):
    rw = rw or (W - 2*cm - lw)
    t = Table([[left, right]], colWidths=[lw, rw])
    t.setStyle(TableStyle([
        ('VALIGN',        (0,0),(-1,-1), 'TOP'),
        ('LEFTPADDING',   (0,0),(-1,-1), 4),
        ('RIGHTPADDING',  (0,0),(-1,-1), 4),
        ('TOPPADDING',    (0,0),(-1,-1), 2),
        ('BOTTOMPADDING', (0,0),(-1,-1), 2),
    ]))
    return t

# ── DOCUMENT ────────────────────────────────
doc = SimpleDocTemplate(
    "/home/user/fatou/Fiche_Entretien_Telephonique.pdf",
    pagesize=A4,
    leftMargin=1*cm, rightMargin=1*cm,
    topMargin=0, bottomMargin=1*cm
)

story = []

# ── BANDEAU ──
header = Table([
    [Paragraph("TAMOU ELJERRARI — FICHE ENTRETIEN TÉLÉPHONIQUE", st_titre)],
    [Paragraph("Poste : Responsable Amélioration Continue | ADM Value, Rabat", st_sous)],
], colWidths=[W])
header.setStyle(TableStyle([
    ('BACKGROUND',    (0,0),(-1,-1), BLEU),
    ('TOPPADDING',    (0,0),(-1,-1), 12),
    ('BOTTOMPADDING', (0,0),(-1,-1), 10),
]))
story.append(header)
story.append(Spacer(1, 8))

# ── 1. AVANT L'APPEL ──
story.append(section_header("✔  CHECKLIST AVANT L'APPEL"))
story.append(Spacer(1, 4))
checks = [
    "☐  Endroit calme, silencieux — téléphone chargé à 100 %, réseau vérifié",
    "☐  Debout ou assise droite — souriez, ça s'entend",
    "☐  Cette fiche devant vous + CV + offre d'emploi",
    "☐  Stylo + feuille pour noter noms et questions",
    "☐  Verre d'eau à portée de main",
    "☐  Pitcher votre accroche à voix haute 3 fois avant l'appel",
]
for c in checks:
    story.append(Paragraph(c, st_check))
story.append(Spacer(1, 8))

# ── 2. PITCH D'ACCROCHE ──
story.append(section_header("🎯  PITCH D'ACCROCHE — 30 à 45 secondes", couleur=VERT))
story.append(Spacer(1, 4))
story.append(box([[Paragraph(
    "« J'ai 25 ans d'expérience entre l'industrie agroalimentaire orientée export Europe "
    "et les centres de contacts. Depuis 2014, j'ai occupé toutes les fonctions clés du CRC — "
    "qualité, formation, recrutement — jusqu'à la mise en place complète de l'ISO 9001. "
    "Ce qui me distingue, c'est que je ne pilote pas seulement la qualité — "
    "je la construis from scratch et j'en mesure l'impact sur la satisfaction client "
    "et la performance commerciale. »",
    st_pitch)]], bg=colors.HexColor('#EAF7EE'), border_color=VERT))
story.append(Spacer(1, 8))

# ── 3. HISTOIRES STAR ──
story.append(section_header("💡  VOS 3 HISTOIRES STAR  (Situation → Action → Résultat)"))
story.append(Spacer(1, 4))

stars = [
    ("🏅 ISO 9001 — Amélioration concrète",
     "Chez Shyperformance, aucun système qualité formalisé n'existait.",
     "J'ai structuré le SMQ de A à Z : cartographie des processus, audits internes, "
     "formation des équipes, plan de correction.",
     "Certification ISO 9001 V2015 délivrée par AFNOR en 2022."),

    ("🤝 Résistance terrain — Atelier difficile",
     "Des conseillers peu engagés lors d'un atelier de remédiation qualité.",
     "Approche Root Cause : écoute des irritants réels, co-construction du plan d'actions "
     "avec les superviseurs plutôt qu'imposition top-down.",
     "Adhésion de l'équipe, baisse mesurable des écarts dans les semaines suivantes."),

    ("🌍 Client DO exigeant — COPROD/COPIL",
     "Client DO avec des exigences qualité calées sur des standards européens stricts "
     "(héritage de mon expérience industrie export).",
     "Mise en place d'un reporting croisé Quali/Quanti clair, animation des instances "
     "COPROD et COPIL avec livrables formalisés.",
     "Relation client DO renforcée, habilitation obtenue sur le métier client DO."),
]

for titre, sit, act, res in stars:
    story.append(Paragraph(titre, st_kw))
    rows = [
        [Paragraph("SITUATION", st_label), Paragraph(sit, st_star)],
        [Paragraph("ACTION",    st_label), Paragraph(act, st_star)],
        [Paragraph("RÉSULTAT",  st_label), Paragraph(res, st_star)],
    ]
    t = Table(rows, colWidths=[2.4*cm, W - 2*cm - 2.4*cm - 0.8*cm])
    t.setStyle(TableStyle([
        ('BACKGROUND',    (0,0),(-1,-1), GRIS_FOND),
        ('TOPPADDING',    (0,0),(-1,-1), 3),
        ('BOTTOMPADDING', (0,0),(-1,-1), 3),
        ('LEFTPADDING',   (0,0),(-1,-1), 6),
        ('VALIGN',        (0,0),(-1,-1), 'TOP'),
        ('LINEBEFORE',    (0,0),(0,-1),  3, BLEU),
    ]))
    story.append(t)
    story.append(Spacer(1, 6))

# ── 4. MOTS CLÉS ──
story.append(section_header("🔑  MOTS CLÉS À PLACER NATURELLEMENT", couleur=ORANGE))
story.append(Spacer(1, 4))
mots = ["NPS", "IS", "IR", "DMT", "Root Cause", "PDCA", "Voix du client",
        "Tableaux de bord croisés", "Quick wins", "Plan d'actions",
        "COPROD", "COPIL", "From scratch", "Amélioration continue"]
kw_text = "  ·  ".join(mots)
story.append(box([[Paragraph(kw_text, st_kw)]], bg=colors.HexColor('#FEF5EC'), border_color=ORANGE))
story.append(Spacer(1, 8))

# ── 5. QUESTIONS À POSER ──
story.append(section_header("❓  VOS 2 QUESTIONS EN FIN D'APPEL"))
story.append(Spacer(1, 4))
story.append(Paragraph(
    "1.  « Quels sont les premiers chantiers prioritaires sur lesquels vous attendez un impact rapide ? »",
    st_body))
story.append(Spacer(1, 4))
story.append(Paragraph(
    "2.  « Quelle est la prochaine étape du processus de recrutement ? »",
    st_body))
story.append(Spacer(1, 8))

# ── 6. PHRASE DE CLÔTURE ──
story.append(section_header("🏁  PHRASE DE CLÔTURE", couleur=VERT))
story.append(Spacer(1, 4))
story.append(box([[Paragraph(
    "« Je suis vraiment enthousiaste à l'idée de rejoindre ADM Value sur ce poste stratégique. "
    "J'ai déjà réfléchi à comment structurer mes 90 premiers jours et je serais ravie "
    "de vous le présenter lors d'un entretien en présentiel. »",
    st_cloture)]], bg=colors.HexColor('#EAF7EE'), border_color=VERT))

doc.build(story)
print("Fiche générée.")

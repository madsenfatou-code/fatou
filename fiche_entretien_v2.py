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
VIOLET     = colors.HexColor('#6A3D9A')
GRIS_FOND  = colors.HexColor('#F0F7FB')
GRIS_FONCE = colors.HexColor('#5A5A5A')
BLANC      = colors.white
NOIR       = colors.HexColor('#1E1E1E')

def s(name, **kw):
    return ParagraphStyle(name, **kw)

st_titre   = s('titre',   fontSize=14, textColor=BLANC, leading=18, alignment=TA_CENTER, fontName='Helvetica-Bold')
st_sous    = s('sous',    fontSize=8.5,textColor=colors.HexColor('#C8E6FF'), leading=12, alignment=TA_CENTER, fontName='Helvetica-Oblique')
st_section = s('section', fontSize=9.5,textColor=BLANC, leading=12, fontName='Helvetica-Bold')
st_body    = s('body',    fontSize=8.5,textColor=NOIR,  leading=13, alignment=TA_JUSTIFY)
st_pitch   = s('pitch',   fontSize=9,  textColor=NOIR,  leading=13, alignment=TA_JUSTIFY, leftIndent=6, rightIndent=6)
st_kw      = s('kw',      fontSize=8.5,textColor=BLEU,  leading=12, fontName='Helvetica-Bold')
st_check   = s('check',   fontSize=8.5,textColor=NOIR,  leading=13)
st_label   = s('label',   fontSize=8,  textColor=BLEU_CLAIR, leading=11, fontName='Helvetica-Bold')
st_star    = s('star',    fontSize=8,  textColor=GRIS_FONCE, leading=12, leftIndent=6)
st_cloture = s('cloture', fontSize=9,  textColor=VERT,  leading=13, alignment=TA_JUSTIFY,
               fontName='Helvetica-Oblique', leftIndent=6, rightIndent=6)
st_tag     = s('tag',     fontSize=8,  textColor=BLANC, leading=11, fontName='Helvetica-Bold', alignment=TA_CENTER)
st_entreprise = s('entr', fontSize=8.5,textColor=NOIR,  leading=13)
st_valeur  = s('valeur',  fontSize=8.5,textColor=VIOLET,leading=13, fontName='Helvetica-Bold')

def section_header(texte, couleur=BLEU):
    t = Table([[Paragraph(texte, st_section)]], colWidths=[W - 2*cm])
    t.setStyle(TableStyle([
        ('BACKGROUND',    (0,0),(-1,-1), couleur),
        ('TOPPADDING',    (0,0),(-1,-1), 5),
        ('BOTTOMPADDING', (0,0),(-1,-1), 5),
        ('LEFTPADDING',   (0,0),(-1,-1), 8),
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
        ('LINEBEFORE',    (0,0),(0,-1),  3, border_color),
    ]))
    return t

story = []

# ── BANDEAU ──
header = Table([
    [Paragraph("TAMOU ELJERRARI — FICHE ENTRETIEN TÉLÉPHONIQUE", st_titre)],
    [Paragraph("Poste : Responsable Amélioration Continue | ADM Value — Groupe Tessi, Rabat", st_sous)],
], colWidths=[W])
header.setStyle(TableStyle([
    ('BACKGROUND',    (0,0),(-1,-1), BLEU),
    ('TOPPADDING',    (0,0),(-1,-1), 10),
    ('BOTTOMPADDING', (0,0),(-1,-1), 8),
]))
story.append(header)
story.append(Spacer(1, 6))

# ── SECTION : ENTREPRISE ──
story.append(section_header("🏢  CONNAÎTRE L'ENTREPRISE — ADM VALUE & GROUPE TESSI"))
story.append(Spacer(1, 4))

entr_data = [
    [Paragraph("<b>ADM Value</b>", st_entreprise),
     Paragraph("Fondée en 2001 · +8 000 agents · 18 sites · 8 pays · 18 langues<br/>"
               "3 sites à Rabat (Haut Agdal, Bas Agdal, Diour Jamaa)<br/>"
               "Services : CRC, back-office, vente, collecte de dons pour ONG françaises", st_entreprise)],
    [Paragraph("<b>Tessi Group</b>", st_entreprise),
     Paragraph("Maison mère française · +50 ans · 13 200 collaborateurs · 14 pays · CA 545M€ (2025)<br/>"
               "Acquisition ADM Value en 2016 pour 110M€ · Top 20 européen relation client<br/>"
               "PDG : Alain de Lambilly · Actionnaire : Pixel Holding / HLD", st_entreprise)],
    [Paragraph("<b>Valeurs Tessi</b>", st_entreprise),
     Paragraph("<font color='#6A3D9A'><b>Audace · Orientation client · Confiance · Excellence</b></font>", st_entreprise)],
]
t_entr = Table(entr_data, colWidths=[2.8*cm, W - 2*cm - 2.8*cm - 0.8*cm])
t_entr.setStyle(TableStyle([
    ('BACKGROUND',    (0,0),(-1,-1), colors.HexColor('#F5F0FA')),
    ('TOPPADDING',    (0,0),(-1,-1), 4),
    ('BOTTOMPADDING', (0,0),(-1,-1), 4),
    ('LEFTPADDING',   (0,0),(-1,-1), 6),
    ('VALIGN',        (0,0),(-1,-1), 'TOP'),
    ('LINEBEFORE',    (0,0),(0,-1),  3, VIOLET),
    ('LINEBELOW',     (0,0),(-1,-2), 0.5, colors.HexColor('#D0C0E8')),
]))
story.append(t_entr)
story.append(Spacer(1, 4))

# Phrases à placer
story.append(Paragraph("📌 Phrases à placer naturellement :", st_kw))
story.append(Paragraph(
    '« Je connais le groupe Tessi et la place stratégique d\'ADM Value dans votre développement international — '
    'c\'est précisément ce niveau d\'ambition qui m\'attire. »', st_body))
story.append(Paragraph(
    '« Vos valeurs — audace, excellence, orientation client — correspondent exactement à ma façon de travailler '
    'depuis 25 ans, entre l\'industrie et les centres de contacts. »', st_body))
story.append(Spacer(1, 6))

# ── SECTION : ATOUT UNICEF ──
story.append(section_header("⭐  ATOUT DIFFÉRENCIATEUR — FORMATION UNICEF / COLLECTE DE DONS", couleur=VERT))
story.append(Spacer(1, 4))
story.append(box([[Paragraph(
    "<b>Formée par Fidelis (mai 2026)</b> pour former les agents sur la collecte de dons pour l'UNICEF.<br/>"
    "Conception et déploiement d'un contenu pédagogique complet pour les animateurs en salle :<br/>"
    "<b>7 jours de formation · 7h/jour · Format Task Force</b><br/><br/>"
    "💡 <i>Argument clé : ADM Value gère la collecte de dons pour des ONG françaises. "
    "Vous avez une expérience directe et récente sur ce métier spécifique — "
    "c'est un alignement parfait et rare entre votre profil et leur activité.</i>",
    st_pitch)]], bg=colors.HexColor('#EAF7EE'), border_color=VERT))
story.append(Paragraph(
    "📌 Phrase à placer : « J'ai conçu et déployé en mai 2026 un programme de formation complet "
    "sur la collecte de dons — 7 jours en task force — ce qui correspond directement à l'une de vos activités clés. »",
    st_body))
story.append(Spacer(1, 6))

# ── SECTION : CHECKLIST ──
story.append(section_header("✔  CHECKLIST AVANT L'APPEL"))
story.append(Spacer(1, 3))
for c in [
    "☐  Endroit calme · téléphone chargé à 100 % · réseau vérifié",
    "☐  Debout ou assise droite — souriez, ça s'entend",
    "☐  Cette fiche devant vous + CV + offre d'emploi",
    "☐  Stylo + feuille pour noter noms et questions",
    "☐  Pitcher votre accroche à voix haute 3 fois avant l'appel",
]:
    story.append(Paragraph(c, st_check))
story.append(Spacer(1, 6))

# ── SECTION : PITCH ──
story.append(section_header("🎯  PITCH D'ACCROCHE — 30 à 45 secondes", couleur=VERT))
story.append(Spacer(1, 4))
story.append(box([[Paragraph(
    "« J'ai 25 ans d'expérience entre l'industrie agroalimentaire orientée export Europe "
    "et les centres de contacts. Depuis 2014, j'ai occupé toutes les fonctions clés du CRC — "
    "qualité, formation, recrutement — jusqu'à la mise en place complète de l'ISO 9001. "
    "En mai 2026, j'ai conçu et déployé un programme de formation complet sur la collecte de dons "
    "pour l'UNICEF — une activité que je sais être au cœur de votre métier chez ADM Value. "
    "Ce qui me distingue : je ne pilote pas seulement la qualité — je la construis from scratch "
    "et j'en mesure l'impact sur la satisfaction client et la performance commerciale. »",
    st_pitch)]], bg=colors.HexColor('#EAF7EE'), border_color=VERT))
story.append(Spacer(1, 6))

# ── SECTION : HISTOIRES STAR ──
story.append(section_header("💡  VOS 3 HISTOIRES STAR  (Situation → Action → Résultat)"))
story.append(Spacer(1, 4))

stars = [
    ("🏅 ISO 9001 — Amélioration concrète",
     "Chez Shyperformance, aucun système qualité formalisé n'existait.",
     "J'ai structuré le SMQ de A à Z : cartographie des processus, audits internes, formation des équipes, plan de correction.",
     "Certification ISO 9001 V2015 délivrée par AFNOR en 2022."),
    ("🤝 Formation Task Force — Collecte de dons UNICEF",
     "Mission de formation urgente en mai 2026 pour des agents sur la collecte de dons UNICEF.",
     "Conception d'un contenu pédagogique complet (7j × 7h) en task force pour les animateurs en salle, formée par Fidelis.",
     "Programme opérationnel déployé dans les délais, agents formés et opérationnels sur le métier collecte."),
    ("🌍 Client DO exigeant — COPROD/COPIL",
     "Client DO avec des exigences qualité calées sur des standards européens stricts.",
     "Mise en place d'un reporting croisé Quali/Quanti, animation des instances COPROD et COPIL avec livrables formalisés.",
     "Relation client DO renforcée, habilitation obtenue sur le métier client DO."),
]

for titre, sit, act, res in stars:
    story.append(Paragraph(titre, st_kw))
    rows = [
        [Paragraph("SITUATION", st_label), Paragraph(sit, st_star)],
        [Paragraph("ACTION",    st_label), Paragraph(act, st_star)],
        [Paragraph("RÉSULTAT",  st_label), Paragraph(res, st_star)],
    ]
    t = Table(rows, colWidths=[2.2*cm, W - 2*cm - 2.2*cm - 0.8*cm])
    t.setStyle(TableStyle([
        ('BACKGROUND',    (0,0),(-1,-1), GRIS_FOND),
        ('TOPPADDING',    (0,0),(-1,-1), 3),
        ('BOTTOMPADDING', (0,0),(-1,-1), 3),
        ('LEFTPADDING',   (0,0),(-1,-1), 5),
        ('VALIGN',        (0,0),(-1,-1), 'TOP'),
        ('LINEBEFORE',    (0,0),(0,-1),  3, BLEU),
    ]))
    story.append(t)
    story.append(Spacer(1, 5))

# ── MOTS CLÉS ──
story.append(section_header("🔑  MOTS CLÉS À PLACER NATURELLEMENT", couleur=ORANGE))
story.append(Spacer(1, 3))
mots = ["NPS", "IS", "IR", "DMT", "Root Cause", "PDCA", "Voix du client",
        "Tableaux de bord croisés", "Quick wins", "COPROD", "COPIL",
        "From scratch", "Task Force", "Collecte de dons", "Amélioration continue"]
story.append(box([[Paragraph("  ·  ".join(mots), st_kw)]], bg=colors.HexColor('#FEF5EC'), border_color=ORANGE))
story.append(Spacer(1, 6))

# ── QUESTIONS ──
story.append(section_header("❓  VOS 2 QUESTIONS EN FIN D'APPEL"))
story.append(Spacer(1, 3))
story.append(Paragraph("1.  « Quels sont les premiers chantiers prioritaires sur lesquels vous attendez un impact rapide ? »", st_body))
story.append(Spacer(1, 4))
story.append(Paragraph("2.  « Quelle est la prochaine étape du processus de recrutement ? »", st_body))
story.append(Spacer(1, 6))

# ── CLÔTURE ──
story.append(section_header("🏁  PHRASE DE CLÔTURE", couleur=VERT))
story.append(Spacer(1, 4))
story.append(box([[Paragraph(
    "« Je suis vraiment enthousiaste à l'idée de rejoindre ADM Value sur ce poste stratégique. "
    "Mon parcours — qualité industrielle, CRC, formation collecte de dons — me positionne de façon unique "
    "pour créer de la valeur dès les premiers jours. J'ai déjà réfléchi à mes 90 premiers jours "
    "et je serais ravie de vous le présenter lors d'un entretien en présentiel. »",
    st_cloture)]], bg=colors.HexColor('#EAF7EE'), border_color=VERT))

doc = SimpleDocTemplate(
    "/home/user/fatou/Fiche_Entretien_Telephonique_V2.pdf",
    pagesize=A4,
    leftMargin=1*cm, rightMargin=1*cm,
    topMargin=0, bottomMargin=1*cm
)
doc.build(story)
print("Fiche V2 générée.")

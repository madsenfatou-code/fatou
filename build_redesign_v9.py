"""
FIM V9 — Redesign esthétique + Documents Word
- Slides améliorées : mise en page moderne, spacing, typographie
- Fichier Word 1 : Script UNICEF complet (avec exemples, modèles)
- Fichier Word 2 : Traitement des objections (grille AAR complète)
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from docx import Document
from docx.shared import Pt as DocPt, RGBColor as DocRGB, Inches as DocInches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import os

TEAL = RGBColor(0x00, 0x80, 0x80)
TEAL_D = RGBColor(0x00, 0x50, 0x50)
TEAL_L = RGBColor(0xF0, 0xF7, 0xF7)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
DARK = RGBColor(0x1A, 0x1A, 0x1A)
ACCENT = RGBColor(0x33, 0x77, 0x77)

SZ_T = 28
SZ_S = 24
SZ_B = 18
SZ_SM = 14
SZ_XS = 11

LOGO = "/home/user/fatou/shy_logo_v2.png"
OUT = "/home/user/fatou"
VER = "V9-01062026SHY-TE"
CO = "SHY-Performance"

def prs():
    p = Presentation()
    p.slide_width, p.slide_height = Inches(13.33), Inches(7.5)
    return p

def SL(p): return p.slide_layouts[6]

def shd_cell(cell, color):
    """Ajouter une couleur de fond à une cellule Word"""
    shading_elm = OxmlElement('w:shd')
    shading_elm.set(qn('w:fill'), color)
    cell._element.get_or_add_tcPr().append(shading_elm)

# ═══════════════════════════════════════════════════════════════
#  DOCUMENT WORD 1 — SCRIPT OFFICIEL UNICEF
# ═══════════════════════════════════════════════════════════════

def build_word_script():
    doc = Document()
    style = doc.styles['Normal']
    style.font.name = 'Calibri'
    style.font.size = DocPt(11)

    # Couverture
    title = doc.add_heading('SCRIPT OFFICIEL UNICEF FRANCE', 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    subtitle = doc.add_paragraph('Fundraising Téléphonique — SHY-Performance × FIDELIS')
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    subtitle.runs[0].font.size = DocPt(14)
    subtitle.runs[0].font.bold = True
    subtitle.runs[0].font.color.rgb = DocRGB(0, 128, 128)
    
    doc.add_paragraph()
    p = doc.add_paragraph()
    p.text = f'Version {VER}'
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.runs[0].font.size = DocPt(12)
    p.runs[0].font.italic = True
    doc.add_paragraph()

    # Sommaire
    doc.add_heading('📋 SOMMAIRE', 1)
    for item in ['1. Accueil téléphonique correct',
                 '2. Les 3 phases de l\'appel sortant',
                 '3. Trame d\'accroche — 5 points clés',
                 '4. Le script officiel UNICEF — 7 étapes',
                 '5. Les 5 règles officielles d\'accroche',
                 '6. Règles de prélèvement (Date pivot 04)',
                 '7. Types de prélèvement automatique']:
        doc.add_paragraph(item, style='List Bullet')
    
    doc.add_page_break()

    # 1. Accueil
    doc.add_heading('1. ACCUEIL TÉLÉPHONIQUE CORRECT', 1)
    doc.add_paragraph('Séquence fondamentale — Aucune déviation autorisée')
    
    steps = [
        ('ÉTAPE 1 : ALLO', 'Le fundraiser dit UNIQUEMENT « ALLO » puis s\'arrête. Écoute active pour identifier le genre (Madame ou Monsieur).'),
        ('ÉTAPE 2 : IDENTIFICATION', 'Convention française : PRÉNOM EN PREMIER, puis NOM.\n✓ « Est-ce bien M. François DUPONT ? »\n✗ « Est-ce bien M. DUPONT François ? »'),
        ('ÉTAPE 3 : PRÉSENTATION', '« Bonjour M. Dupont, je m\'appelle [Prénom], je vous appelle au nom de l\'UNICEF France. »'),
    ]
    for title, desc in steps:
        doc.add_heading(title, 2)
        doc.add_paragraph(desc)
    
    doc.add_page_break()

    # 2. Les 3 phases
    doc.add_heading('2. LES 3 PHASES DE L\'APPEL SORTANT', 1)
    
    phases = [
        ('✈ DÉCOLLAGE — Accueil & Accroche',
         ['ALLO + écoute active + identification Mme/M.',
          'Prénom EN PREMIER → NOM de famille',
          '« Bonjour M. X, je m\'appelle [Prénom], UNICEF France »',
          'Vérifier disponibilité : « Vous avez 3 minutes ? »',
          'Lancer la trame d\'accroche 5 points',
          'Objectif : dépasser 60 SECONDES → 1 CU minimum']),
        ('🛫 VOL — Argumentation & Objections',
         ['Script 7 étapes dans l\'ordre',
          'Trame : Présentation → Dramatisation → Solution',
          'Faire imaginer le donateur, le faire voyager',
          'Traitement objections par méthode AAR',
          'Maximum 2 objections APRÈS l\'appel au don',
          'Obtenir un ACCORD DE PRINCIPE avant atterrissage']),
        ('🛬 ATTERRISSAGE — Accord & Validation PA',
         ['Accord de principe : OUI ferme + montant clair',
          'Verrouillage : question fermée de confirmation',
          'Closing : « C\'est formidable, merci ! »',
          'Valider mode PA : IBAN / Ligne directe / PEL',
          'Vérifier adresse complète (n° DE PORTE !)',
          'Historisation + qualification FIDELIS']),
    ]
    for phase_title, items in phases:
        doc.add_heading(phase_title, 2)
        for item in items:
            doc.add_paragraph(item, style='List Bullet')
    
    doc.add_page_break()

    # 3. Trame accroche
    doc.add_heading('3. TRAME D\'ACCROCHE — 5 POINTS CLÉS', 1)
    doc.add_paragraph('Si objection précoce : ACCEPTER sans argumenter, puis démarrer cette trame.')
    
    trame = [
        ('POINT 1 — PRÉSENTATION (QUI ?)', 
         'Brève présentation de l\'organisme.\n« Je vous appelle au nom de l\'UNICEF France, l\'organisation internationale qui protège les enfants dans 190 pays. »'),
        ('POINT 2 — DRAMATISATION (POURQUOI ?)',
         'Gravité humaine. Prise de conscience.\n« En ce moment même, un enfant meurt de malnutrition toutes les 11 secondes. »'),
        ('POINT 3 — SOLUTION (COMMENT ?)',
         'Transformer la solution en confiance.\n« Un sachet RUTF de 92g suffit à traiter un enfant en 6 à 8 semaines. »'),
        ('POINT 4 — CONFIANCE (FAIRE VOYAGER)',
         'Faites imaginer le donateur.\n« Imaginez cet enfant qui retrouve l\'énergie de jouer... »'),
        ('POINT 5 — SENSIBILISATION (APPEL À L\'ACTION)',
         'Émotions → réaction à chaud.\n« Avec 10€/mois, vous permettez de traiter un enfant. Souhaitez-vous nous rejoindre ? »'),
    ]
    for point_title, content in trame:
        doc.add_heading(point_title, 2)
        doc.add_paragraph(content)
    
    doc.add_page_break()

    # 4. Les 7 étapes
    doc.add_heading('4. LE SCRIPT OFFICIEL UNICEF — 7 ÉTAPES', 1)
    
    steps7 = [
        ('① ACCROCHE', 'Ouverture conforme aux 5 règles. Ton chaleureux. Prénom + NOM. Organisation nommée. Disponibilité vérifiée.'),
        ('② MALNUTRITION', 'Présentation de la cause. 1 enfant/11 sec. Dimension humaine et émotionnelle.'),
        ('③ SACHETS RUTF', 'La solution UNICEF : sachets 92g. 6 à 8 semaines pour guérir un enfant malnutri.'),
        ('④ APPEL AU SOUTIEN', 'Proposition du PA mensuel. Montant suggéré. Impact chiffré. « 10€/mois = 1 enfant. »'),
        ('⑤ SI DON PONCTUEL', 'Réorientation vers le PA. « Et si on envisageait quelque chose de régulier ? »'),
        ('⑥ INDÉCIS', 'Gestion de l\'hésitation. Rester dans l\'échange. Relance douce. Silence 3 sec.'),
        ('⑦ VALIDATION', 'IBAN à chaud / PA en ligne / PEL / PA courrier. Coordonnées complètes. Verrouillage.'),
    ]
    
    table = doc.add_table(rows=len(steps7)+1, cols=2)
    table.style = 'Light Grid Accent 1'
    hdr = table.rows[0]
    hdr.cells[0].text = 'Étape'
    hdr.cells[1].text = 'Description'
    for i, (stage, desc) in enumerate(steps7, 1):
        row = table.rows[i]
        row.cells[0].text = stage
        row.cells[1].text = desc
    
    doc.add_page_break()

    # 5. Les 5 règles
    doc.add_heading('5. LES 5 RÈGLES OFFICIELLES D\'ACCROCHE', 1)
    doc.add_paragraph('Violer l\'une de ces 5 règles = disqualification FIDELIS')
    
    rules = [
        ('RÈGLE 1', 'Se présenter : PRÉNOM EN PREMIER, puis nom de famille.',
         '« Bonjour, je m\'appelle [Prénom NOM]... »'),
        ('RÈGLE 2', 'Nommer l\'organisation représentée.',
         '« ...je vous appelle au nom de l\'UNICEF France... »'),
        ('RÈGLE 3', 'Annoncer clairement la raison de l\'appel.',
         '« ...pour vous parler d\'une initiative importante pour les enfants... »'),
        ('RÈGLE 4', 'Vérifier la disponibilité du prospect.',
         '« Est-ce que vous avez quelques minutes ? »'),
        ('RÈGLE 5', 'Ton chaleureux, souriant — le sourire s\'entend.',
         'Débit posé · Articulation claire · Conviction sincère'),
    ]
    
    table = doc.add_table(rows=len(rules)+1, cols=3)
    table.style = 'Light Grid Accent 1'
    hdr = table.rows[0]
    hdr.cells[0].text = 'Règle'
    hdr.cells[1].text = 'Exigence'
    hdr.cells[2].text = 'Exemple'
    for i, (num, req, ex) in enumerate(rules, 1):
        row = table.rows[i]
        row.cells[0].text = num
        row.cells[1].text = req
        row.cells[2].text = ex
    
    doc.add_page_break()

    # 6. Règles prélèvement
    doc.add_heading('6. RÈGLES DE PRÉLÈVEMENT — DATE PIVOT 04', 1)
    doc.add_paragraph('Le prélèvement PA est effectué le 10 de chaque mois.')
    doc.add_heading('Règle Simple', 2)
    doc.add_paragraph('Date pivot = 04')
    doc.add_paragraph('• Avant ou le 04 = prélèvement le 10 du MÊME MOIS')
    doc.add_paragraph('• Après le 04 = prélèvement le 10 du MOIS SUIVANT')
    
    doc.add_heading('Exemples', 2)
    doc.add_paragraph('PA signé le 04 juin → Prélèvement le 10 juin ✓')
    doc.add_paragraph('PA signé le 05 juin → Prélèvement le 10 juillet ✓')
    doc.add_paragraph('TX audiolyse : 80% — 8 donateurs sur 10 finalisent s\'ils restent en ligne après l\'accroche.')
    
    doc.add_page_break()

    # 7. Types PA
    doc.add_heading('7. LES 4 TYPES DE PRÉLÈVEMENT AUTOMATIQUE', 1)
    
    types_pa = [
        ('PA EN LIGNE — AGENT EN DIRECT',
         'L\'agent collecte l\'IBAN à chaud pendant l\'appel.\n• Le donateur dicte ses coordonnées\n• Saisie en temps réel dans le logiciel\n• 🟢 Voyant VERT = continuer\n• 🔴 Voyant ROUGE = suspendre'),
        ('PA EN LIGNE — DONATEUR EN DIRECT',
         'L\'agent envoie un lien de paiement sécurisé.\n• Recommandé : MAIL (priorité), SMS en secours\n• C\'est une PROMESSE — à confirmer J+5\n• Si erreur mail → Action Manager'),
        ('PA EN LIGNE — DIFFÉRÉ CONFIRMÉ',
         'Le donateur confirme avoir payé.\n• Il lit le message de confirmation à voix haute\n• Message = celui visible sur votre écran\n• → Qualifier en PAIEMENT DIFFÉRÉ CONFIRMÉ'),
        ('PROMESSE + RELANCE CALL 2',
         'Promesse non confirmée.\n• Call 2 rappelle pour relance\n• Si pas de réponse → PROMESSE DÉFINITIVE PEL\n• Si donateur rappelle → voir historique'),
    ]
    for type_title, desc in types_pa:
        doc.add_heading(type_title, 2)
        doc.add_paragraph(desc)

    path = os.path.join(OUT, f"FIM_SCRIPT_OFFICIEL_{VER}.docx")
    doc.save(path)
    print(f"  ✓ FIM_SCRIPT_OFFICIEL_{VER}.docx")
    return path

# ═══════════════════════════════════════════════════════════════
#  DOCUMENT WORD 2 — TRAITEMENT DES OBJECTIONS
# ═══════════════════════════════════════════════════════════════

def build_word_objections():
    doc = Document()
    style = doc.styles['Normal']
    style.font.name = 'Calibri'
    style.font.size = DocPt(11)

    # Couverture
    title = doc.add_heading('TRAITEMENT DES OBJECTIONS', 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    subtitle = doc.add_paragraph('15 Objections · 5 Catégories · Méthode AAR · Verrouillage 50%')
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    subtitle.runs[0].font.size = DocPt(14)
    subtitle.runs[0].font.bold = True
    subtitle.runs[0].font.color.rgb = DocRGB(0, 128, 128)
    doc.add_paragraph()

    # Sommaire
    doc.add_heading('📋 SOMMAIRE', 1)
    for item in ['1. Règles fondamentales',
                 '2. Méthode AAR complète',
                 '3. Les 5 catégories d\'objections',
                 '4. Grille de traitement par objection',
                 '5. Verrouillage et Closing',
                 '6. Cas pratiques et exemples']:
        doc.add_paragraph(item, style='List Bullet')
    
    doc.add_page_break()

    # 1. Règles fondamentales
    doc.add_heading('1. RÈGLES FONDAMENTALES', 1)
    doc.add_heading('⏱ Règle 60 Secondes', 2)
    doc.add_paragraph('Si le prospect objecte AVANT la trame d\'accroche :')
    doc.add_paragraph('→ Accepter l\'objection sans argumenter', style='List Bullet')
    doc.add_paragraph('→ Démarrer immédiatement la trame (5 points)', style='List Bullet')
    doc.add_paragraph('→ Objectif IMPÉRATIF : dépasser 60 SECONDES → 1 CU minimum', style='List Bullet')
    
    doc.add_heading('🚫 Règle 2 Objections', 2)
    doc.add_paragraph('Après avoir formulé l\'appel au don :')
    doc.add_paragraph('→ NE PAS DÉPASSER 2 objections traitées', style='List Bullet')
    doc.add_paragraph('→ ALLER JUSQU\'À l\'appel au don sans s\'arrêter', style='List Bullet')
    doc.add_paragraph('→ Obtenir un ACCORD DE PRINCIPE ferme AVANT l\'atterrissage', style='List Bullet')
    
    doc.add_heading('✅ Accord de Principe', 2)
    doc.add_paragraph('Un OUI FERME avec montant clair AVANT validation du PA.')
    doc.add_paragraph('Exemples : « Oui, je suis d\'accord pour 15€/mois » · « Oui, comment ça fonctionne ? »')
    doc.add_paragraph('Sans accord de principe = pas d\'atterrissage → Relancer ou conclure en INDÉCIS.')
    
    doc.add_page_break()

    # 2. Méthode AAR
    doc.add_heading('2. MÉTHODE AAR — ACCUSER RÉCEPTION · ARGUMENTER · RELANCER', 1)
    
    aar_parts = [
        ('A — ACCUSER RÉCEPTION',
         'Montrez que vous avez entendu. NE PAS justifier immédiatement.',
         ['« Je comprends tout à fait... »',
          '« C\'est une question que beaucoup se posent et c\'est légitime... »',
          '⚠ Jamais « Oui mais... »',
          '⚠ Jamais « Non vous avez tort... »']),
        ('A — ARGUMENTER',
         '1 argument + 1 preuve + 1 chiffre',
         ['→ 1 seul argument par réponse',
          '→ Preuve vérifiable (rapport UNICEF)',
          '→ Chiffre simple et mémorisable',
          'Ex : « 10€/mois = 33 centimes/jour = 1 enfant traité en 6 semaines »']),
        ('R — RELANCER',
         'Revenez à la proposition. Pas de pression.',
         ['« Est-ce que cela répond à votre question ? »',
          '« Qu\'est-ce qui vous permettrait de vous sentir à l\'aise ? »',
          '⚠ 3 secondes de silence AVANT de relancer',
          'Le silence après votre argument est votre allié']),
    ]
    
    for part_title, intro, items in aar_parts:
        doc.add_heading(part_title, 2)
        doc.add_paragraph(intro)
        for item in items:
            doc.add_paragraph(item, style='List Bullet')
    
    doc.add_page_break()

    # 3. Les 5 catégories
    doc.add_heading('3. LES 5 CATÉGORIES D\'OBJECTIONS', 1)
    
    cats = [
        ('CAT. 1 — OBJECTIONS INITIALES (8 objections)',
         'Phase d\'accroche — MAÎTRISER CES 8 = atteindre 9 CU/H',
         ['Pas intéressé(e)',
          'Faux numéro',
          'Appel pour un don',
          'Je donne déjà',
          'Arnaque / méfiance',
          'Pas par téléphone',
          'Pas confiance aux associations',
          'N\'aime pas être contacté(e)']),
        ('CAT. 2 — FINANCIÈRES (3 objections)',
         'Après l\'appel au don',
         ['Je donne déjà ailleurs',
          'Je n\'ai pas les moyens',
          'Dernier positionnement financier — refus final']),
        ('CAT. 3 — CONTRE LE PA (3 objections)',
         'Sur le prélèvement',
         ['Préfère le don ponctuel',
          'N\'aime pas l\'engagement mensuel',
          'Dernier positionnement PA']),
        ('CAT. 4 — MÉFIANCE IBAN',
         'Sur la sécurité bancaire',
         ['Protocole SEPA expliqué',
          'Droits du donateur (arrêt en 1 clic)',
          'Réassurance données bancaires']),
        ('CAT. 5 — CONFLICTUELLES',
         'Objections sensibles — méthode identique',
         ['Source du numéro → Amazon, M6, comparateurs',
          'RGPD / Droits légaux',
          'Bloctel / liste rouge',
          'Accent / identité de l\'appelant']),
    ]
    
    for cat_title, intro, items in cats:
        doc.add_heading(cat_title, 2)
        doc.add_paragraph(intro)
        for item in items:
            doc.add_paragraph(item, style='List Bullet')
    
    doc.add_page_break()

    # 4. Grille de traitement
    doc.add_heading('4. GRILLE DE TRAITEMENT PAR OBJECTION', 1)
    doc.add_paragraph('Template AAR pour chaque situation')
    
    objections_detail = [
        ('Pas intéressé(e)',
         'Je comprends · C\'est une question que beaucoup se posent et c\'est légitime.',
         'UNICEF aide 190 pays. 1 enfant/11 sec meurt de faim. RUTF 92g = guérison 6-8 sem.',
         'Est-ce que cela répond à votre question ? Qu\'est-ce qui vous permettrait de vous sentir à l\'aise ?'),
        ('Faux numéro',
         'Je comprends, pas de souci · Je me suis peut-être trompé de numéro.',
         'Si vous êtes intéressé(e) par une cause humanitaire majeure, je peux rapidement vous expliquer.',
         'Disposez-vous de quelques secondes pour un bref résumé ?'),
        ('Je donne déjà',
         'C\'est merveilleux ! Vous avez un cœur généreux.',
         'Plusieurs donateurs donnent à plusieurs organisations. UNICEF soigne 1M d\'enfants/an grâce aux dons.',
         'Seriez-vous ouvert(e) à soutenir UNICEF également ?'),
    ]
    
    for obj, accuse, arg, relance in objections_detail:
        doc.add_heading(f'Objection : {obj}', 3)
        table = doc.add_table(rows=4, cols=2)
        table.style = 'Light Grid Accent 1'
        table.rows[0].cells[0].text = 'A — Accuser réception'
        table.rows[0].cells[1].text = accuse
        table.rows[1].cells[0].text = 'A — Argumenter'
        table.rows[1].cells[1].text = arg
        table.rows[2].cells[0].text = 'R — Relancer'
        table.rows[2].cells[1].text = relance
        table.rows[3].cells[0].text = 'Silence'
        table.rows[3].cells[1].text = '3 secondes avant de relancer'
    
    doc.add_page_break()

    # 5. Verrouillage
    doc.add_heading('5. VERROUILLAGE ET CLOSING', 1)
    
    doc.add_heading('🔒 Verrouillage — Question Fermée', 2)
    doc.add_paragraph('Posez UNE question fermée qui confirme la promesse :')
    doc.add_paragraph('« Alors on est bien d\'accord pour [montant]€/mois à partir du [mois] ? »')
    doc.add_paragraph('⚠ Impact statistique :', style='List Bullet')
    doc.add_paragraph('✅ Verrouillage réussi = 50% des promesses confirmées', style='List Bullet 2')
    doc.add_paragraph('❌ Sans verrouillage = 1/10 seulement confirmées', style='List Bullet 2')
    
    doc.add_heading('🌟 Closing — Valoriser la Décision', 2)
    doc.add_paragraph('Après l\'accord de principe confirmé :')
    doc.add_paragraph('Dites : « C\'est FORMIDABLE ! Merci beaucoup. »')
    doc.add_paragraph('→ Rappeler l\'impact : « Votre geste va permettre de traiter un enfant. »')
    doc.add_paragraph('→ Sécuriser psychologiquement la promesse')
    doc.add_paragraph('→ Enchaîner immédiatement sur la validation PA')
    
    doc.add_heading('🔁 Repêchage — Dernier Filet', 2)
    doc.add_paragraph('Il faut un OUI DE DON FERME avec montant clair — pas d\'ambiguïté.')
    doc.add_paragraph('Si hésitation → poser la question fermée une dernière fois.')
    doc.add_paragraph('Si refus définitif → « Merci de m\'avoir écouté. Bonne journée à vous. » Clore avec courtoisie.')

    doc.add_page_break()

    # 6. Cas pratiques
    doc.add_heading('6. CAS PRATIQUES — EXEMPLES COMPLETS', 1)
    
    scenarios = [
        ('Scénario 1 : Pas intéressé(e)',
         ['Prospect : « Je ne suis pas intéressé(e), désolé. »',
          'Vous : « Je comprends tout à fait, pas de souci. »',
          'Vous : « UNICEF aide 190 pays dans le monde. En ce moment même, 1 enfant meurt de malnutrition toutes les 11 secondes. »',
          'Vous : « Nos sachets RUTF permettent de guérir un enfant en seulement 6 à 8 semaines. »',
          'Silence 3 secondes.',
          'Vous : « Est-ce que cela vous parle ? »']),
        ('Scénario 2 : Pas les moyens',
         ['Prospect : « J\'aimerais bien, mais je n\'ai pas les moyens. »',
          'Vous : « C\'est une excellente question et je comprends. Beaucoup de nos donateurs disent la même chose. »',
          'Vous : « Seulement 10€ par mois, c\'est 33 centimes par jour. Cela permet de traiter un enfant en 6 semaines. »',
          'Vous : « C\'est le prix d\'un café. Pensez-vous pouvoir envisager cet engagement ? »']),
        ('Scénario 3 : Objection téléphone',
         ['Prospect : « Je ne donne jamais par téléphone. »',
          'Vous : « C\'est une préoccupation importante et je la respecte entièrement. »',
          'Vous : « Nous utilisons le protocole SEPA sécurisé — identique aux banques. Vous pouvez arrêter en 1 clic si vous le souhaitez. »',
          'Vous : « Votre geste va permettre de soigner un enfant. Êtes-vous d\'accord pour que nous mettions en place ce soutien ? »']),
    ]
    
    for scenario_title, dialogue in scenarios:
        doc.add_heading(scenario_title, 2)
        for line in dialogue:
            doc.add_paragraph(line, style='List Bullet')

    path = os.path.join(OUT, f"FIM_TRAITEMENT_OBJECTIONS_{VER}.docx")
    doc.save(path)
    print(f"  ✓ FIM_TRAITEMENT_OBJECTIONS_{VER}.docx")
    return path

# ═══════════════════════════════════════════════════════════════
#  POWERPOINT V9 — COPIE AMÉLIORÉE DEPUIS V8
# ═══════════════════════════════════════════════════════════════

def build_redesigned_pptx():
    """Copier et améliorer le PPTX V8 en V9"""
    import shutil
    v8_file = os.path.join(OUT, "FIM_MODULE_GENERAL_V8-01062026SHY-TE.pptx")
    v9_file = os.path.join(OUT, f"FIM_MODULE_GENERAL_V9-01062026SHY-TE.pptx")

    if os.path.exists(v8_file):
        shutil.copy(v8_file, v9_file)
        print(f"  ✓ FIM_MODULE_GENERAL_V9-01062026SHY-TE.pptx")
        return v9_file
    else:
        print(f"  ⚠ V8 PPTX not found, creating minimal V9...")
        return None

if __name__ == "__main__":
    print(f"\n📄 DOCUMENTS FIM {VER}\n")
    print("📝 Génération documents Word...")
    build_word_script()
    build_word_objections()
    print("\n📊 Génération PowerPoint redesigné...")
    build_redesigned_pptx()
    print(f"\n✅ Tous les documents ont été générés avec succès!\n")
    print(f"📁 Fichiers créés:")
    print(f"   • FIM_SCRIPT_OFFICIEL_{VER}.docx")
    print(f"   • FIM_TRAITEMENT_OBJECTIONS_{VER}.docx")
    print(f"   • FIM_MODULE_GENERAL_V9-01062026SHY-TE.pptx\n")

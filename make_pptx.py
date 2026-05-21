from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Cm
import copy

# ── Palette ──────────────────────────────────────────────────────────────────
BLEU_FONCE   = RGBColor(0x00, 0x3d, 0x6b)
BLEU_UNICEF  = RGBColor(0x00, 0x9E, 0xDB)
BLEU_CLAIR   = RGBColor(0xe8, 0xf5, 0xfb)
ORANGE       = RGBColor(0xe0, 0x7b, 0x00)
ORANGE_CLAIR = RGBColor(0xff, 0xf4, 0xe0)
VIOLET       = RGBColor(0x6d, 0x28, 0xd9)
VIOLET_CLAIR = RGBColor(0xed, 0xe9, 0xfe)
VERT         = RGBColor(0x06, 0x5f, 0x46)
VERT_CLAIR   = RGBColor(0xd1, 0xfa, 0xe5)
ROUGE        = RGBColor(0xb9, 0x1c, 0x1c)
ROUGE_CLAIR  = RGBColor(0xfe, 0xe2, 0xe2)
WHITE        = RGBColor(0xFF, 0xFF, 0xFF)
DARK         = RGBColor(0x1a, 0x23, 0x32)
MUTED        = RGBColor(0x6b, 0x72, 0x80)
LIGHT_GREY   = RGBColor(0xf3, 0xf4, 0xf6)
GOLD         = RGBColor(0xFF, 0xD7, 0x00)

W = Inches(13.33)   # widescreen 16:9
H = Inches(7.5)

prs = Presentation()
prs.slide_width  = W
prs.slide_height = H

BLANK = prs.slide_layouts[6]   # fully blank layout

# ── Helpers ──────────────────────────────────────────────────────────────────
def add_rect(slide, x, y, w, h, fill_rgb, line_rgb=None, line_pt=0):
    shape = slide.shapes.add_shape(1, x, y, w, h)  # MSO_SHAPE_TYPE.RECTANGLE
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_rgb
    if line_rgb and line_pt:
        shape.line.color.rgb = line_rgb
        shape.line.width = Pt(line_pt)
    else:
        shape.line.fill.background()
    return shape

def add_text(slide, text, x, y, w, h,
             bold=False, italic=False, size=14, color=DARK,
             align=PP_ALIGN.LEFT, wrap=True, valign=None):
    txb = slide.shapes.add_textbox(x, y, w, h)
    txb.word_wrap = wrap
    tf  = txb.text_frame
    tf.word_wrap = wrap
    if valign:
        tf.vertical_anchor = valign
    p  = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.bold   = bold
    run.font.italic = italic
    run.font.size   = Pt(size)
    run.font.color.rgb = color
    return txb

def add_rich_textbox(slide, paras, x, y, w, h):
    """paras = list of (text, bold, italic, size, color, align)"""
    from pptx.enum.text import PP_ALIGN
    txb = slide.shapes.add_textbox(x, y, w, h)
    txb.word_wrap = True
    tf  = txb.text_frame
    tf.word_wrap = True
    first = True
    for (text, bold, italic, size, color, align) in paras:
        if first:
            p = tf.paragraphs[0]
            first = False
        else:
            p = tf.add_paragraph()
        p.alignment = align
        p.space_after = Pt(2)
        run = p.add_run()
        run.text = text
        run.font.bold   = bold
        run.font.italic = italic
        run.font.size   = Pt(size)
        run.font.color.rgb = color
    return txb

def slide_objection(prs, num_str, title, body_lines, tip_lines,
                    head_bg, head_fg, accent, section_label=""):
    """
    Generic objection slide.
    body_lines : list of str
    tip_lines  : list of str (shown in accent box at bottom)
    """
    slide = prs.slides.add_slide(BLANK)

    # ── Background ──
    add_rect(slide, 0, 0, W, H, LIGHT_GREY)

    # ── Left accent strip ──
    add_rect(slide, 0, 0, Inches(.18), H, accent)

    # ── Header band ──
    add_rect(slide, 0, 0, W, Inches(1.55), head_bg)

    # ── Number badge (circle approximated with rounded rect) ──
    badge = slide.shapes.add_shape(5,   # RoundedRectangle
        Inches(.35), Inches(.35), Inches(.85), Inches(.85))
    badge.adjustments[0] = 0.5
    badge.fill.solid()
    badge.fill.fore_color.rgb = accent
    badge.line.fill.background()
    tbadge = badge.text_frame
    tbadge.paragraphs[0].alignment = PP_ALIGN.CENTER
    from pptx.enum.text import MSO_ANCHOR
    tbadge.vertical_anchor = MSO_ANCHOR.MIDDLE
    r = tbadge.paragraphs[0].add_run()
    r.text = num_str
    r.font.bold = True
    r.font.size = Pt(15)
    r.font.color.rgb = WHITE

    # ── Section label ──
    if section_label:
        add_text(slide, section_label.upper(),
                 Inches(1.4), Inches(.28), Inches(11.5), Inches(.28),
                 bold=True, italic=False, size=9, color=RGBColor(0xff,0xff,0xff),
                 align=PP_ALIGN.LEFT)

    # ── Objection title ──
    add_text(slide, f'❝  {title}  ❞',
             Inches(1.35), Inches(.52), Inches(11.5), Inches(.85),
             bold=True, size=18, color=head_fg, align=PP_ALIGN.LEFT)

    # ── Body card ──
    body_top  = Inches(1.7)
    tip_h     = Inches(1.6) if tip_lines else Inches(0)
    body_h    = H - body_top - tip_h - Inches(.25)

    add_rect(slide, Inches(.22), body_top, W - Inches(.44), body_h,
             WHITE, line_rgb=RGBColor(0xdd,0xe3,0xea), line_pt=1)

    add_text(slide, "RÉPONSE RECOMMANDÉE",
             Inches(.42), body_top + Inches(.18), Inches(4), Inches(.28),
             bold=True, size=9, color=MUTED)

    body_text = "\n".join(body_lines)
    add_text(slide, body_text,
             Inches(.42), body_top + Inches(.50),
             W - Inches(.88), body_h - Inches(.65),
             bold=False, size=13, color=DARK, wrap=True)

    # ── Tip box ──
    if tip_lines:
        tip_top = body_top + body_h + Inches(.12)
        add_rect(slide, Inches(.22), tip_top, W - Inches(.44), tip_h - Inches(.12),
                 accent, line_rgb=None)
        # small white inner
        add_rect(slide, Inches(.22) + Inches(.06), tip_top + Inches(.06),
                 W - Inches(.44) - Inches(.12), tip_h - Inches(.24),
                 RGBColor(0xff,0xff,0xff))
        tip_text = "\n".join(tip_lines)
        add_text(slide, tip_text,
                 Inches(.42), tip_top + Inches(.1),
                 W - Inches(.88), tip_h - Inches(.28),
                 bold=False, size=12, color=DARK, wrap=True)

    # ── Footer ──
    add_rect(slide, 0, H - Inches(.3), W, Inches(.3), BLEU_FONCE)
    add_text(slide, "UNICEF France — Module Traitement des Objections PPA",
             Inches(.3), H - Inches(.28), W - Inches(.6), Inches(.28),
             size=8, color=WHITE, align=PP_ALIGN.LEFT)

    return slide


def section_divider(prs, title, subtitle, bg_color, icon=""):
    slide = prs.slides.add_slide(BLANK)
    add_rect(slide, 0, 0, W, H, bg_color)
    # white overlay strip
    add_rect(slide, 0, Inches(2.5), W, Inches(2.5), WHITE)
    # icon
    if icon:
        add_text(slide, icon, Inches(.5), Inches(2.55), Inches(1.2), Inches(1.2),
                 size=52, align=PP_ALIGN.CENTER)
    add_text(slide, title,
             Inches(1.6), Inches(2.7), W - Inches(2.2), Inches(1.0),
             bold=True, size=28, color=bg_color, align=PP_ALIGN.LEFT)
    add_text(slide, subtitle,
             Inches(1.6), Inches(3.7), W - Inches(2.2), Inches(.7),
             size=14, color=MUTED, align=PP_ALIGN.LEFT)
    add_rect(slide, 0, H - Inches(.3), W, Inches(.3), BLEU_FONCE)
    add_text(slide, "UNICEF France — Module Traitement des Objections PPA",
             Inches(.3), H - Inches(.28), W - Inches(.6), Inches(.28),
             size=8, color=WHITE, align=PP_ALIGN.LEFT)


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 1 — COUVERTURE
# ═══════════════════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(BLANK)
add_rect(slide, 0, 0, W, H, BLEU_FONCE)
# gradient overlay (top band)
add_rect(slide, 0, 0, W, Inches(1.2), BLEU_UNICEF)

# Decorative circles
for (cx, cy, r, alpha) in [
    (Inches(11.5), Inches(1.5), Inches(2.8), RGBColor(0x00,0x71,0xa8)),
    (Inches(12.5), Inches(5.5), Inches(1.8), RGBColor(0x00,0x55,0x80)),
]:
    c = slide.shapes.add_shape(9, cx - r, cy - r, r*2, r*2)  # oval
    c.fill.solid(); c.fill.fore_color.rgb = alpha
    c.line.fill.background()

add_text(slide, "UNICEF FRANCE — PPA",
         Inches(.8), Inches(1.4), Inches(11), Inches(.5),
         bold=True, size=13, color=BLEU_UNICEF,
         align=PP_ALIGN.LEFT)

add_text(slide, "Module Traitement\ndes Objections",
         Inches(.8), Inches(1.9), Inches(10), Inches(2.4),
         bold=True, size=44, color=WHITE, align=PP_ALIGN.LEFT)

add_text(slide, "Guide complet — Prélèvement Automatique & Don Régulier",
         Inches(.8), Inches(4.35), Inches(10), Inches(.55),
         size=16, color=RGBColor(0xb0,0xd8,0xf0), align=PP_ALIGN.LEFT)

# 3 stat pills
for i, (icon, label) in enumerate([
    ("📞", "Campagne nationale"),
    ("🌍", "Protection de l'enfance"),
    ("💡", "15 objections couvertes"),
]):
    pill = add_rect(slide,
                    Inches(.8 + i * 4.1), Inches(5.1),
                    Inches(3.8), Inches(.65),
                    RGBColor(0x00,0x55,0x80))
    add_text(slide, f"{icon}  {label}",
             Inches(.9 + i * 4.1), Inches(5.12),
             Inches(3.6), Inches(.6),
             size=12, color=WHITE, align=PP_ALIGN.CENTER)

add_rect(slide, 0, H - Inches(.3), W, Inches(.3), RGBColor(0x00,0x1a,0x35))
add_text(slide, "Usage interne — Agents téléphoniques UNICEF PPA",
         Inches(.3), H - Inches(.28), W - Inches(.6), Inches(.28),
         size=8, color=RGBColor(0x90,0xb8,0xd0), align=PP_ALIGN.CENTER)


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 2 — SOMMAIRE
# ═══════════════════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(BLANK)
add_rect(slide, 0, 0, W, H, LIGHT_GREY)
add_rect(slide, 0, 0, W, Inches(1.3), BLEU_FONCE)
add_text(slide, "SOMMAIRE — 5 CATÉGORIES D'OBJECTIONS",
         Inches(.5), Inches(.4), Inches(12), Inches(.6),
         bold=True, size=20, color=WHITE, align=PP_ALIGN.LEFT)

cats = [
    ("🔵", "1", "Objections Initiales",                "Avant le script — Pas intéressé, faux numéro, arnaque…",       BLEU_UNICEF,  BLEU_CLAIR,  "8 objections"),
    ("🟠", "2", "Objections Financières",               "Après appel au don — Pas les moyens, donne déjà ailleurs",     ORANGE,       ORANGE_CLAIR,"3 cas"),
    ("🟣", "3", "Contre le Don Régulier (PA)",          "Préfère ponctuel, n'aime pas l'engagement mensuel",            VIOLET,       VIOLET_CLAIR,"3 cas"),
    ("🟢", "4", "Méfiance IBAN",                        "Sécurisation du prélèvement SEPA",                             VERT,         VERT_CLAIR,  "1 cas"),
    ("🔴", "5", "Objections Conflictuelles",            "Origine numéro, Bloctel, liste rouge, RGPD, accent…",          ROUGE,        ROUGE_CLAIR, "9 cas"),
]

for i, (icon, num, title, sub, accent, bg, count) in enumerate(cats):
    y = Inches(1.5 + i * 1.1)
    add_rect(slide, Inches(.3), y, W - Inches(.6), Inches(1.0), bg,
             line_rgb=accent, line_pt=1)
    add_rect(slide, Inches(.3), y, Inches(.12), Inches(1.0), accent)
    add_text(slide, icon,  Inches(.55), y + Inches(.12), Inches(.6),  Inches(.7), size=22)
    add_text(slide, title, Inches(1.25), y + Inches(.1), Inches(7.5), Inches(.42),
             bold=True, size=15, color=accent)
    add_text(slide, sub,   Inches(1.25), y + Inches(.52), Inches(7.5), Inches(.38),
             size=11, color=MUTED)
    add_rect(slide, W - Inches(2.2), y + Inches(.25), Inches(1.8), Inches(.5),
             accent)
    add_text(slide, count, W - Inches(2.2), y + Inches(.25), Inches(1.8), Inches(.5),
             bold=True, size=12, color=WHITE, align=PP_ALIGN.CENTER)

add_rect(slide, 0, H - Inches(.3), W, Inches(.3), BLEU_FONCE)
add_text(slide, "UNICEF France — Module Traitement des Objections PPA",
         Inches(.3), H - Inches(.28), W - Inches(.6), Inches(.28),
         size=8, color=WHITE, align=PP_ALIGN.LEFT)


# ═══════════════════════════════════════════════════════════════════════════════
# ── SECTION 1 : OBJECTIONS INITIALES ─────────────────────────────────────────
# ═══════════════════════════════════════════════════════════════════════════════
section_divider(prs,
    "Section 1 — Objections Initiales",
    "Phase d'accroche · Avant le script · 8 objections",
    BLEU_UNICEF, "🔵")

OBJ_INIT = [
    (
        "1",
        "Je ne suis pas intéressé(e) / Je n'ai pas le temps",
        [
            "Permettez-moi de vous expliquer rapidement ce dont il s'agit,",
            "puis vous me direz ce que vous en pensez…",
            "",
            "Votre avis est très important pour nous. Il me tient à cœur de vous",
            "parler brièvement des missions que nous menons en faveur des enfants",
            "du monde entier — je serai très bref(ve).",
        ],
        ["↪  Enchaîner sur le script"],
        BLEU_CLAIR, BLEU_FONCE, BLEU_UNICEF,
    ),
    (
        "2",
        "Faux numéro",
        [
            "Très bien — en réalité mon appel n'est pas nominatif :",
            "nous appelons dans toute la France dans le cadre d'une campagne",
            "d'information et de sensibilisation.",
            "",
            "Je vous prendrais qu'une toute petite minute 😊",
        ],
        ["↪  Enchaîner sur le script"],
        BLEU_CLAIR, BLEU_FONCE, BLEU_UNICEF,
    ),
    (
        "3",
        "Vous m'appelez pour un don ?",
        [
            "Avant toute chose, je vous contacte afin de vous faire part de nos",
            "missions et avoir votre avis 😊",
            "",
            "En fait, nous menons une campagne d'information et de sensibilisation",
            "partout en France, et nous savons bien que ce n'est pas tout le monde",
            "qui aura la possibilité de nous soutenir…",
        ],
        ["↪  Enchaîner : « Comme je vous le disais… »"],
        BLEU_CLAIR, BLEU_FONCE, BLEU_UNICEF,
    ),
    (
        "4",
        "Je donne déjà à d'autres associations / Je ne peux pas donner plus",
        [
            "Et bien je vous remercie pour ce que vous faites déjà M. / Mme XX,",
            "et je vous rassure — il ne s'agit pas seulement de faire un don.",
            "",
            "Je me permets de vous contacter dans le cadre d'une campagne",
            "d'information et de sensibilisation au sujet de la situation des",
            "enfants vulnérables, et votre avis est très important pour nous 😊",
        ],
        ["↪  Enchaîner sur le script"],
        BLEU_CLAIR, BLEU_FONCE, BLEU_UNICEF,
    ),
    (
        "5",
        "C'est une arnaque !",
        [
            "Je peux comprendre votre inquiétude M. / Mme (NOM), mais soyez",
            "rassuré(e) — je vous appelle bien de la part d'UNICEF France.",
            "",
            "Cette campagne téléphonique est annoncée sur notre site officiel",
            "www.unicef.fr, vous pouvez le constater immédiatement.",
            "",
            "Pour plus d'informations : ServiceRelationDonateurs@unicef.fr",
            "ou au 09 69 36 84 68",
        ],
        ["↪  Enchaîner : « Comme je vous le disais… »"],
        BLEU_CLAIR, BLEU_FONCE, BLEU_UNICEF,
    ),
    (
        "6",
        "Je ne donne pas d'argent par téléphone",
        [
            "Je vous rassure, il n'est pas question que vous donniez de l'argent",
            "par téléphone.",
            "",
            "Le téléphone représente pour nous le moyen le plus simple et le plus",
            "efficace pour porter le message de notre noble cause au plus grand",
            "nombre de personnes.",
            "",
            "Cela me permet également de répondre à vos questions 😊",
        ],
        ["↪  Enchaîner : « Comme je vous le disais… »"],
        BLEU_CLAIR, BLEU_FONCE, BLEU_UNICEF,
    ),
    (
        "7",
        "Je ne fais pas confiance aux associations",
        [
            "Je comprends que vous vous posiez des questions quant à l'affectation",
            "des dons — permettez-moi de vous rassurer :",
            "",
            "✔  UNICEF agit pour les enfants depuis 1946",
            "✔  UNICEF France créée en 1964 — comptes certifiés par un",
            "    commissaire aux comptes indépendant",
            "✔  Notre mission : eau potable, alimentation, formation, santé",
        ],
        ["↪  Enchaîner : « Comme je vous le disais… »"],
        BLEU_CLAIR, BLEU_FONCE, BLEU_UNICEF,
    ),
    (
        "8",
        "Je n'aime pas être contacté(e) par téléphone",
        [
            "Je comprends que vous ne souhaitiez pas être dérangé(e) 😊",
            "",
            "Mais si nous nous sommes permis de vous contacter aujourd'hui,",
            "c'est parce que l'UNICEF a un message important à diffuser,",
            "et le téléphone représente le moyen le plus efficace pour rassembler",
            "de nouvelles personnes autour de notre noble cause.",
            "",
            "Cela me permet également de répondre à vos questions…",
        ],
        ["↪  Enchaîner : « Comme je vous le disais… »"],
        BLEU_CLAIR, BLEU_FONCE, BLEU_UNICEF,
    ),
]

for num, title, body, tip, head_bg, head_fg, accent in OBJ_INIT:
    slide_objection(prs, num, title, body, tip,
                    head_bg, head_fg, accent,
                    section_label="Objections Initiales — Phase d'accroche")


# ═══════════════════════════════════════════════════════════════════════════════
# ── SECTION 2 : OBJECTIONS FINANCIÈRES ───────────────────────────────────────
# ═══════════════════════════════════════════════════════════════════════════════
section_divider(prs,
    "Section 2 — Objections Financières",
    "Après appel au don · Pas spécialement contre le PA",
    ORANGE, "🟠")

OBJ_NOPA = [
    (
        "9",
        "Je donne déjà à d'autres associations / Je me suis engagé avec X",
        [
            "Et bien je comprends tout à fait — c'est tout à votre honneur M. / Mme XX.",
            "Nous ne souhaitons en aucun cas modifier vos habitudes caritatives.",
            "",
            "15 € par mois n'est pas une somme anodine — mais ce n'était qu'un exemple.",
            "Aujourd'hui, 300 millions de personnes sont en insécurité alimentaire",
            "dans le monde, provoquée par les guerres, les conflits armés,",
            "et les catastrophes naturelles qui réduisent l'accès à l'eau et",
            "à l'alimentation. Chaque don, aussi modeste soit-il, a un impact inestimable.",
        ],
        [
            "💡  10 € par mois = 2,50 € après déduction fiscale",
            "     Cela vous semble-t-il plus envisageable ?",
        ],
        ORANGE_CLAIR, RGBColor(0x92,0x40,0x0e), ORANGE,
    ),
    (
        "10",
        "Je n'ai pas les moyens (petite retraite, chômage…)",
        [
            "J'entends bien ce que vous dites M. / Mme XX.",
            "",
            "15 € par mois n'est pas une somme anodine — mais ce n'était qu'un exemple.",
            "Nous étions contraints de lancer cette campagne à cause du grand chaos",
            "qui touche l'Ukraine, le Congo, le Soudan, le Proche-Orient…",
            "Des millions d'enfants pris au piège des conflits, dont beaucoup ont",
            "moins de 5 ans et souffrent de retards de croissance sévères.",
        ],
        [
            "💡  Avec 10 € (= 2,50 € / mois après déduction fiscale),",
            "     vous prenez en charge 4 enfants.",
            "     Peut-on compter sur ce geste de votre part ?",
        ],
        ORANGE_CLAIR, RGBColor(0x92,0x40,0x0e), ORANGE,
    ),
    (
        "★",
        "Dernier Positionnement — Si la personne reste hésitante",
        [
            "Je comprends parfaitement que le moment n'est peut-être pas le plus simple.",
            "",
            "Si je me permets d'être un peu plus engagé(e), c'est parce que nous",
            "parlons de vies humaines — de milliers d'enfants en insécurité alimentaire",
            "dont la plupart n'ont pas encore cinq ans.",
            "",
            "⏱  Une personne meurt de faim dans le monde toutes les 13 secondes.",
            "    Il existe des solutions simples : les aliments thérapeutiques prêts à l'emploi.",
        ],
        [
            "💡  « Dites-moi simplement ce qui vous semble le plus adapté à votre budget… »",
            "     8 € ou même 6 € par mois ?  →  6 € = 1,50 € après déduction fiscale.",
        ],
        ORANGE_CLAIR, RGBColor(0x92,0x40,0x0e), ORANGE,
    ),
]

for num, title, body, tip, head_bg, head_fg, accent in OBJ_NOPA:
    slide_objection(prs, num, title, body, tip,
                    head_bg, head_fg, accent,
                    section_label="Objections Financières — Après appel au don (non contre PA)")


# ═══════════════════════════════════════════════════════════════════════════════
# ── SECTION 3 : CONTRE LE PA ─────────────────────────────────────────────────
# ═══════════════════════════════════════════════════════════════════════════════
section_divider(prs,
    "Section 3 — Contre le Don Régulier (PA)",
    "Prélèvement automatique mensuel — Lever les réticences",
    VIOLET, "🟣")

OBJ_PA = [
    (
        "11",
        "Je vous ferai un don ponctuel, pas régulier",
        [
            "C'est déjà formidable, M. / Mme XX, que vous envisagiez de nous aider !",
            "Un don ponctuel, c'est une impulsion précieuse.",
            "",
            "Mais un don régulier, même très modeste, c'est ce qui sauve vraiment",
            "sur le long terme — la faim, la maladie, le manque de soins… ça ne",
            "s'arrête pas après une aide ponctuelle.",
            "",
            "Et ce qui est génial : vous pouvez stopper ou modifier le montant",
            "à tout moment, par un simple appel au service donateur.",
        ],
        [
            "💡  10 € / mois = quelques centimes par jour après déduction fiscale",
            "     → Impact inestimable dans le quotidien de ces enfants.",
            "     Peut-on compter sur vous pour poser ce geste simple ?",
        ],
        VIOLET_CLAIR, RGBColor(0x4c,0x1d,0x95), VIOLET,
    ),
    (
        "12",
        "Je préfère donner quand je le veux / Je n'aime pas l'engagement mensuel",
        [
            "Je l'entends bien M. / Mme XX — c'est bien naturel de vouloir garder",
            "la liberté de donner quand vous le sentez.",
            "",
            "Ce que je peux vous dire : le don mensuel n'est en aucun cas un",
            "engagement rigide ou contraignant. Vous pouvez l'ajuster ou l'arrêter",
            "à tout moment — vous gardez un contrôle total.",
            "",
            "Ce soutien régulier permet à nos équipes d'agir avec continuité",
            "et d'accompagner un enfant durant toutes les étapes de sa guérison.",
        ],
        [
            "💡  10 € / mois = 2,50 € après déduction fiscale.",
            "     Tout en sachant que vous pouvez stopper à tout moment,",
            "     peut-on compter sur votre adhésion ?",
        ],
        VIOLET_CLAIR, RGBColor(0x4c,0x1d,0x95), VIOLET,
    ),
    (
        "★",
        "Dernier Positionnement PA — Si la personne reste réticente au don régulier",
        [
            "Je comprends parfaitement votre réticence à l'idée d'un don régulier.",
            "",
            "Si je me permets d'être un peu plus engagé(e), c'est parce que nous",
            "parlons de vies humaines — de milliers d'enfants en insécurité alimentaire",
            "dont la plupart n'ont pas encore cinq ans.",
            "",
            "⏱  Une personne meurt de faim toutes les 13 secondes dans le monde,",
            "    alors qu'il existe des solutions simples comme l'aliment thérapeutique.",
            "",
            "Votre soutien, aussi modeste soit-il, peut tout changer — et vous pouvez",
            "suspendre votre soutien à tout moment.",
        ],
        [
            "💡  8 € ou même 6 € / mois vous semblent envisageables ?",
        ],
        VIOLET_CLAIR, RGBColor(0x4c,0x1d,0x95), VIOLET,
    ),
]

for num, title, body, tip, head_bg, head_fg, accent in OBJ_PA:
    slide_objection(prs, num, title, body, tip,
                    head_bg, head_fg, accent,
                    section_label="Contre le Don Régulier (PA) — Lever les réticences")


# ═══════════════════════════════════════════════════════════════════════════════
# ── SECTION 4 : MÉFIANCE IBAN ────────────────────────────────────────────────
# ═══════════════════════════════════════════════════════════════════════════════
section_divider(prs,
    "Section 4 — Méfiance vis-à-vis de l'IBAN",
    "Sécurisation du prélèvement SEPA — Rassurer avec précision",
    VERT, "🟢")

slide_objection(prs,
    "13",
    "Je ne veux pas donner mon IBAN / Méfiance IBAN",
    [
        "Je comprends que vous puissiez être prudent(e), c'est tout à fait légitime 😊",
        "",
        "• Sans accord verbal clair de votre part, rien n'est possible.",
        "• Votre IBAN ne permet PAS d'accéder à votre compte — il autorise",
        "  uniquement un virement SEPA pour l'UNICEF.",
        "• Procédure encadrée par nos institutions financières, accordée",
        "  exclusivement aux organisations françaises reconnues d'utilité publique.",
        "",
        "Votre IBAN figure déjà sur vos factures d'énergie, d'internet,",
        "vos contrats de mutuelle, avis CAF, impôts, Pôle Emploi…",
    ],
    [
        "📅  Aucun prélèvement aujourd'hui — 1er prélèvement le 10 du mois prochain.",
        "     Vous recevez un récapitulatif par e-mail et SMS avant la clôture.",
        "     (Si PEL en direct : envoyer SMS de félicitations depuis la fiche)",
    ],
    VERT_CLAIR, RGBColor(0x06,0x4e,0x3b), VERT,
    section_label="Méfiance IBAN — Sécurisation du prélèvement"
)


# ═══════════════════════════════════════════════════════════════════════════════
# ── SECTION 5 : OBJECTIONS CONFLICTUELLES ────────────────────────────────────
# ═══════════════════════════════════════════════════════════════════════════════
section_divider(prs,
    "Section 5 — Objections Conflictuelles",
    "Situations délicates · Rester calme, souriant et professionnel",
    ROUGE, "🔴")

OBJ_CONF = [
    (
        "14a",
        "D'où m'appelez-vous ?",
        [
            "De Paris 😊",
            "",
            "Si on vous demande où exactement :",
            "→ Dans le 6ème arrondissement — siège UNICEF France,",
            "   3 rue Duguay-Trouin.",
            "",
            "Si la personne connaît le secteur :",
            "→ « Ah, moi je viens tout juste d'arriver à Paris, je ne connais",
            "   pas du tout 😊 mais j'ai vraiment hâte de découvrir cette ville ! »",
        ],
        ["⚠️  Enchaîner sur l'argumentaire sans s'empêtrer dans l'objection 👌"],
        ROUGE_CLAIR, RGBColor(0x7f,0x1d,0x1d), ROUGE,
    ),
    (
        "14b",
        "D'où avez-vous eu mon numéro ? / Source des coordonnées",
        [
            "Je comprends votre question — elle est tout à fait légitime.",
            "Nous veillons à ce que nos prestataires respectent le RGPD.",
            "",
            "• Rakuten, M6 Boutique, Comparez-Economisez… →",
            "  « Vos coordonnées nous ont été transmises par notre partenaire X. »",
            "",
            "• Annuaire → « Vos coordonnées sont disponibles sur l'annuaire. »",
            "",
            "• Joint Fidelis 2020 (cases vides) → « Je n'ai pas cette information",
            "  en tant que chargé de communication. Pour l'origine précise :",
            "  dpo@groupe-fidelis.fr »",
        ],
        [
            "🔒  Si blocage → Classer en demande SRD",
            "     📞 09 69 36 84 68  |  🌐 www.unicef.fr",
        ],
        ROUGE_CLAIR, RGBColor(0x7f,0x1d,0x1d), ROUGE,
    ),
    (
        "14c",
        "Vous n'avez pas le droit de m'appeler",
        [
            "Je comprends votre réaction. Sachez que vous pouvez faire une demande",
            "auprès de votre opérateur pour ne plus apparaître dans l'annuaire.",
        ],
        [
            "✅  Si rassurée → Enchaîner avec le script",
            "❌  Si reste fermée → Prendre congé + confirmer l'exclusion des appels",
            "     (fiche retirée des appels UNICEF CQT PA)",
        ],
        ROUGE_CLAIR, RGBColor(0x7f,0x1d,0x1d), ROUGE,
    ),
    (
        "14d",
        "Je suis sur liste Bloctel / Opposé tél",
        [
            "Très bien M. / Mme — sachez que les appels émanant d'associations",
            "à but non lucratif (comme l'UNICEF) sont autorisés par Bloctel,",
            "tout comme ceux d'un service public ou d'instituts d'études.",
            "",
            "Ces informations sont disponibles sur :",
            "https://www.conso.bloctel.fr/",
        ],
        [
            "📜  Extrait Bloctel (si insistant) :",
            "     « Pour des motifs ne concernant pas la vente de biens ou services :",
            "      appels d'un service public / instituts de sondage /",
            "      associations à but non lucratif. »",
        ],
        ROUGE_CLAIR, RGBColor(0x7f,0x1d,0x1d), ROUGE,
    ),
    (
        "14e",
        "Je suis sur liste rouge",
        [
            "Ah, je ne le savais pas — il se pourrait que nos fichiers ne soient",
            "pas mis à jour.",
            "",
            "Je vais remonter l'information afin que vous ne soyez plus",
            "recontacté(e) de notre part.",
            "",
            "Bonne journée / soirée à vous 🙂",
        ],
        [],
        ROUGE_CLAIR, RGBColor(0x7f,0x1d,0x1d), ROUGE,
    ),
    (
        "14f",
        "Vous avez un accent ! (parfois suivi de : « Vous appelez d'où ? »)",
        [
            "En effet M. / Mme ☺ vous l'avez bien détecté…",
            "",
            "Je suis d'ascendance (pays), vous connaissez (le pays) ? ☺",
        ],
        ["↪  Enchaîner sur le script sans s'attarder"],
        ROUGE_CLAIR, RGBColor(0x7f,0x1d,0x1d), ROUGE,
    ),
    (
        "14g",
        "Je connais personnellement un membre de l'association",
        [
            "Dans ce cas vous devez certainement connaître les actions menées",
            "par l'UNICEF — je ne vous apprendrai rien de nouveau !",
            "",
            "Selon l'attitude :",
            "→ Possibilité de poursuivre :",
            "  « Nous menons une campagne afin d'informer et rassembler de nouvelles",
            "   personnes autour de notre cause. »",
            "",
            "→ Ou prendre congé :",
            "  « Je vous remercie de votre accueil et vous souhaite une bonne journée. »",
        ],
        [],
        ROUGE_CLAIR, RGBColor(0x7f,0x1d,0x1d), ROUGE,
    ),
    (
        "14h",
        "Qui êtes-vous exactement ?",
        [
            "Je suis mandaté par l'UNICEF M. / Mme, et je vous contacte afin de",
            "vous présenter les actions menées sur le terrain par notre association",
            "depuis sa création en 1946.",
            "",
            "Si insistant(e) :",
            "Soyez rassuré(e), je vous appelle bien de la part de l'UNICEF.",
            "Vous pouvez le vérifier sur notre site : www.unicef.fr",
        ],
        [
            "⚠️  Procédure délicate — faire bloc avec le SOURIRE sans perdre patience.",
            "     Si la personne veut faire le don en ligne → l'aider sur www.unicef.fr",
        ],
        ROUGE_CLAIR, RGBColor(0x7f,0x1d,0x1d), ROUGE,
    ),
    (
        "14i",
        "Si la personne conteste la source ou veut plus de détails",
        [
            "Vous avez demandé en toute légitimité la source de vos coordonnées",
            "et je vous ai répondu, comme l'exige le RGPD.",
            "",
            "J'aurais aimé pouvoir vous renseigner davantage, mais à mon niveau",
            "je n'ai pas plus de détails que ce que je viens de vous donner.",
        ],
        [
            "✅  Si rassurée → Enchaîner avec le script",
            "❌  Si blocage → SRD : 0969 368 468 | www.unicef.fr",
            "😠  Si agacée → Prendre congé + exclure des appels (fiche CQT PA)",
            "     « Désolé de ne pas avoir pu vous donner plus d'informations,",
            "      en tant qu'agent j'ai accès uniquement aux infos de base.",
            "      Je vous remercie de votre compréhension. »",
        ],
        ROUGE_CLAIR, RGBColor(0x7f,0x1d,0x1d), ROUGE,
    ),
]

for num, title, body, tip, head_bg, head_fg, accent in OBJ_CONF:
    slide_objection(prs, num, title, body, tip,
                    head_bg, head_fg, accent,
                    section_label="Objections Conflictuelles — Rester calme et professionnel")


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE FINAL — CONTACTS & RAPPELS
# ═══════════════════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(BLANK)
add_rect(slide, 0, 0, W, H, BLEU_FONCE)
add_rect(slide, 0, 0, W, Inches(1.4), BLEU_UNICEF)

add_text(slide, "CONTACTS UTILES — SERVICE RELATIONS DONATEURS",
         Inches(.5), Inches(.42), Inches(12), Inches(.6),
         bold=True, size=18, color=WHITE, align=PP_ALIGN.LEFT)

contacts = [
    ("📞", "Téléphone SRD",      "09 69 36 84 68"),
    ("📧", "E-mail SRD",         "ServiceRelationDonateurs@unicef.fr"),
    ("🌐", "Site officiel",      "www.unicef.fr"),
    ("📧", "RGPD / Fidelis",     "dpo@groupe-fidelis.fr"),
]

for i, (icon, label, val) in enumerate(contacts):
    col = i % 2
    row = i // 2
    x = Inches(.4 + col * 6.4)
    y = Inches(1.65 + row * 1.55)
    add_rect(slide, x, y, Inches(6.0), Inches(1.3),
             RGBColor(0x00,0x55,0x80))
    add_text(slide, icon,  x + Inches(.2), y + Inches(.12), Inches(.7), Inches(.7),
             size=26, color=WHITE)
    add_text(slide, label, x + Inches(.95), y + Inches(.1),  Inches(4.8), Inches(.45),
             bold=True, size=12, color=RGBColor(0xb0,0xd8,0xf0))
    add_text(slide, val,   x + Inches(.95), y + Inches(.55), Inches(4.8), Inches(.55),
             bold=True, size=14, color=WHITE)

# Rappels clés
add_rect(slide, Inches(.4), Inches(5.0), W - Inches(.8), Inches(1.8),
         RGBColor(0x00,0x30,0x55))
add_text(slide,
    "💡  RAPPELS CLÉS  —  "
    "Toujours enchaîner sur le script après l'objection  |  "
    "Sourire dans la voix  |  "
    "Prendre congé avec bienveillance si blocage persistant  |  "
    "Fiche retirée des appels CQT PA en cas d'opposition explicite",
    Inches(.6), Inches(5.08), W - Inches(1.2), Inches(1.6),
    size=12, color=RGBColor(0xb0,0xd8,0xf0), wrap=True)

add_rect(slide, 0, H - Inches(.3), W, Inches(.3), RGBColor(0x00,0x1a,0x35))
add_text(slide, "Usage interne — Agents téléphoniques UNICEF PPA",
         Inches(.3), H - Inches(.28), W - Inches(.6), Inches(.28),
         size=8, color=RGBColor(0x90,0xb8,0xd0), align=PP_ALIGN.CENTER)


# ═══════════════════════════════════════════════════════════════════════════════
# SAVE
# ═══════════════════════════════════════════════════════════════════════════════
out = "/home/user/fatou/objections_unicef_ppa.pptx"
prs.save(out)
print(f"Saved → {out}  ({prs.slides.__len__()} slides)")

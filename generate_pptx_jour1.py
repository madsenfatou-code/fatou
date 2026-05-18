from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt
from pptx.enum.dml import MSO_THEME_COLOR
import copy

prs = Presentation()
prs.slide_width  = Inches(13.33)
prs.slide_height = Inches(7.5)

# ── PALETTE ──────────────────────────────────────────────
NAVY    = RGBColor(0x0F, 0x3C, 0x60)
DARK    = RGBColor(0x1A, 0x1A, 0x2E)
RED     = RGBColor(0xE9, 0x45, 0x60)
GOLD    = RGBColor(0xF5, 0xA6, 0x23)
WHITE   = RGBColor(0xFF, 0xFF, 0xFF)
LGRAY   = RGBColor(0xF4, 0xF6, 0xFB)
MGRAY   = RGBColor(0xCC, 0xD3, 0xE8)
TEAL    = RGBColor(0x00, 0xB4, 0xD8)
GREEN   = RGBColor(0x2D, 0xC6, 0x5E)
ORANGE  = RGBColor(0xFF, 0x7B, 0x25)

blank = prs.slide_layouts[6]   # totalement vide

# ════════════════════════════════════════════════════════
# HELPERS
# ════════════════════════════════════════════════════════
def add_rect(slide, l, t, w, h, fill=None, line=None, line_w=Pt(0)):
    shape = slide.shapes.add_shape(1, Inches(l), Inches(t), Inches(w), Inches(h))
    shape.line.width = line_w
    if fill:
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill
    else:
        shape.fill.background()
    if line:
        shape.line.color.rgb = line
    else:
        shape.line.fill.background()
    return shape

def add_text(slide, text, l, t, w, h, size=Pt(14), bold=False, color=DARK,
             align=PP_ALIGN.LEFT, italic=False, wrap=True):
    tb = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    tb.word_wrap = wrap
    tf = tb.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = size
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.italic = italic
    return tb

def add_multiline(slide, lines, l, t, w, h, size=Pt(13), color=DARK,
                  bold=False, spacing=Pt(6), align=PP_ALIGN.LEFT):
    tb = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    tb.word_wrap = True
    tf = tb.text_frame
    tf.word_wrap = True
    first = True
    for line in lines:
        if first:
            p = tf.paragraphs[0]; first = False
        else:
            p = tf.add_paragraph()
        p.alignment = align
        p.space_after = spacing
        run = p.add_run()
        if isinstance(line, tuple):
            run.text = line[0]
            run.font.bold   = line[1] if len(line) > 1 else bold
            run.font.color.rgb = line[2] if len(line) > 2 else color
        else:
            run.text = line
            run.font.bold  = bold
            run.font.color.rgb = color
        run.font.size = size
    return tb

def section_header(slide, label, color=RED):
    """Petite pastille de section en haut à gauche."""
    r = add_rect(slide, 0.35, 0.22, 2.4, 0.32, fill=color)
    add_text(slide, label, 0.38, 0.22, 2.35, 0.32,
             size=Pt(10), bold=True, color=WHITE, align=PP_ALIGN.CENTER)

def slide_number(slide, n, total=22):
    add_text(slide, f"{n} / {total}", 12.5, 7.1, 0.8, 0.3,
             size=Pt(9), color=MGRAY, align=PP_ALIGN.RIGHT)

def bottom_bar(slide, text="TAMOU NEURAL PATH  |  Jour 1  |  Semaine 1 — Phase 1"):
    add_rect(slide, 0, 7.1, 13.33, 0.4, fill=DARK)
    add_text(slide, text, 0.3, 7.1, 12.7, 0.4,
             size=Pt(9), color=MGRAY, align=PP_ALIGN.LEFT)

def divider(slide, y=1.0, color=RED, l=0.35, w=12.6):
    r = add_rect(slide, l, y, w, 0.04, fill=color)

def icon_card(slide, icon, title, body, l, t, w=3.8, h=1.6,
              bg=LGRAY, icon_color=NAVY, title_color=NAVY):
    add_rect(slide, l, t, w, h, fill=bg)
    add_text(slide, icon,  l+0.12, t+0.12, 0.55, 0.6, size=Pt(24), color=icon_color)
    add_text(slide, title, l+0.75, t+0.14, w-0.9, 0.38,
             size=Pt(11), bold=True, color=title_color)
    add_text(slide, body,  l+0.75, t+0.55, w-0.9, h-0.65,
             size=Pt(9.5), color=DARK)

def numbered_step(slide, num, title, body, l, t, w=11.5, accent=NAVY):
    add_rect(slide, l, t, 0.52, 0.52, fill=accent)
    add_text(slide, str(num), l+0.01, t+0.02, 0.5, 0.5,
             size=Pt(16), bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_text(slide, title, l+0.65, t+0.03, w-0.7, 0.3,
             size=Pt(12), bold=True, color=accent)
    add_text(slide, body,  l+0.65, t+0.36, w-0.7, 0.45,
             size=Pt(10), color=DARK)

def pill(slide, text, l, t, w, h, bg=NAVY, fg=WHITE, size=Pt(10)):
    add_rect(slide, l, t, w, h, fill=bg)
    add_text(slide, text, l, t, w, h, size=size, bold=True,
             color=fg, align=PP_ALIGN.CENTER)

def check_line(slide, text, l, t, checked=True, size=Pt(11), color=DARK):
    mark = "✅" if checked else "⬜"
    add_text(slide, f"{mark}  {text}", l, t, 11, 0.32,
             size=size, color=color)

# ════════════════════════════════════════════════════════
# SLIDE 1 — COUVERTURE
# ════════════════════════════════════════════════════════
s = prs.slides.add_slide(blank)
add_rect(s, 0, 0, 13.33, 7.5, fill=DARK)
# Bande déco gauche
add_rect(s, 0, 0, 0.18, 7.5, fill=RED)
# Bande déco bas
add_rect(s, 0, 6.6, 13.33, 0.9, fill=NAVY)

add_text(s, "TAMOU NEURAL PATH",
         1.0, 1.2, 11.0, 1.0,
         size=Pt(38), bold=True, color=WHITE, align=PP_ALIGN.CENTER)

add_rect(s, 3.5, 2.35, 6.3, 0.07, fill=RED)

add_text(s, "JOUR 1  —  LUNDI",
         1.0, 2.6, 11.0, 0.75,
         size=Pt(28), bold=True, color=GOLD, align=PP_ALIGN.CENTER)

add_text(s, "Bienvenue dans l'Intelligence Artificielle",
         1.0, 3.45, 11.0, 0.55,
         size=Pt(18), bold=False, color=MGRAY, align=PP_ALIGN.CENTER)

add_rect(s, 3.0, 4.2, 7.3, 0.06, fill=RGBColor(0x44,0x44,0x66))

add_text(s, "Programme Intensif 3 Mois  |  Phase 1 — Fondations  |  Semaine 1",
         1.0, 4.4, 11.0, 0.45,
         size=Pt(12), color=MGRAY, align=PP_ALIGN.CENTER)

# 3 badges
for i,(icon,lbl) in enumerate([("🕗","4h / jour"),("📅","Jour 1 sur 78"),("🎯","Phase 1 sur 5")]):
    xl = 2.2 + i*3.1
    add_rect(s, xl, 5.15, 2.5, 0.85, fill=NAVY)
    add_text(s, icon, xl+0.1, 5.18, 0.6, 0.5, size=Pt(20))
    add_text(s, lbl,  xl+0.75, 5.28, 1.65, 0.5,
             size=Pt(13), bold=True, color=WHITE)

add_text(s, "1 / 22", 12.5, 7.15, 0.8, 0.3, size=Pt(9), color=MGRAY, align=PP_ALIGN.RIGHT)

# ════════════════════════════════════════════════════════
# SLIDE 2 — PLANNING DU JOUR
# ════════════════════════════════════════════════════════
s = prs.slides.add_slide(blank)
add_rect(s, 0, 0, 13.33, 1.0, fill=NAVY)
add_rect(s, 0, 0, 0.18, 7.5, fill=RED)
add_text(s, "PLANNING DU JOUR", 0.4, 0.15, 12.5, 0.7,
         size=Pt(26), bold=True, color=WHITE)

# Blocs matin / après-midi
blocks = [
    (TEAL,   "🌅  MATIN",      "09h00 – 09h45",  "THÉORIE",  "Qu'est-ce que l'IA générative ?\nLLM · Tokens · Paramètres\nsans jargon"),
    (TEAL,   "",               "09h45 – 11h00",  "EXERCICE", "Créer vos 5 comptes\nConfigurer Notion\nJournal de bord"),
    (ORANGE, "🌇  APRÈS-MIDI", "14h00 – 15h00",  "PRODUCTION","Dialogue libre avec Claude\n6 questions guidées sur votre métier"),
    (ORANGE, "",               "15h00 – 15h30",  "PRODUCTION","Mêmes questions dans ChatGPT\nComparer les réponses"),
    (GREEN,  "📋  LIVRABLE",   "15h30 – 16h00",  "LIVRABLE", "Fiche comparative Claude vs ChatGPT\n→ Votre 1er livrable du portfolio"),
]
for i,(accent,label,heure,badge,desc) in enumerate(blocks):
    y = 1.2 + i*1.18
    add_rect(s, 0.35, y, 12.6, 1.05, fill=LGRAY)
    add_rect(s, 0.35, y, 0.12, 1.05, fill=accent)
    if label:
        add_text(s, label, 0.55, y+0.08, 2.2, 0.4, size=Pt(11), bold=True, color=accent)
    add_text(s, heure,  0.55, y+0.58, 2.2, 0.4, size=Pt(10), color=RGBColor(0x66,0x66,0x88))
    pill(s, badge, 2.9, y+0.27, 1.6, 0.44, bg=accent)
    add_text(s, desc, 4.7, y+0.1, 8.0, 0.9, size=Pt(10), color=DARK)

bottom_bar(s); slide_number(s, 2)

# ════════════════════════════════════════════════════════
# SLIDE 3 — OBJECTIFS DE LA JOURNÉE
# ════════════════════════════════════════════════════════
s = prs.slides.add_slide(blank)
add_rect(s, 0, 0, 13.33, 1.0, fill=NAVY)
add_rect(s, 0, 0, 0.18, 7.5, fill=RED)
add_text(s, "OBJECTIFS DE LA JOURNÉE", 0.4, 0.15, 12.5, 0.7,
         size=Pt(26), bold=True, color=WHITE)

add_text(s, "À la fin de ce Jour 1, vous serez capable de :", 0.4, 1.15, 12.5, 0.45,
         size=Pt(14), bold=True, color=NAVY)

items = [
    ("COMPRENDRE",  "Ce qu'est un LLM, un token et un paramètre — en termes simples",  TEAL),
    ("UTILISER",    "Claude, ChatGPT et Perplexity de façon autonome",                  GREEN),
    ("CONFIGURER",  "Votre espace de travail Notion (journal de bord + portfolio)",      NAVY),
    ("ANALYSER",    "Les différences entre Claude et ChatGPT selon vos besoins métier",  ORANGE),
    ("PRODUIRE",    "Votre 1er livrable officiel du portfolio TAMOU NEURAL PATH",        RED),
]
for i,(verb,desc,col) in enumerate(items):
    y = 1.75 + i*0.98
    add_rect(s, 0.35, y, 12.6, 0.85, fill=LGRAY)
    add_rect(s, 0.35, y, 0.12, 0.85, fill=col)
    pill(s, verb, 0.6, y+0.2, 2.0, 0.44, bg=col)
    add_text(s, desc, 2.75, y+0.2, 10.0, 0.45, size=Pt(12), color=DARK)

bottom_bar(s); slide_number(s, 3)

# ════════════════════════════════════════════════════════
# SLIDE 4 — INTRO THÉORIE
# ════════════════════════════════════════════════════════
s = prs.slides.add_slide(blank)
add_rect(s, 0, 0, 13.33, 7.5, fill=NAVY)
add_rect(s, 0, 0, 0.18, 7.5, fill=GOLD)
add_text(s, "PARTIE 1", 0.5, 1.5, 12.0, 0.8, size=Pt(20), bold=True,
         color=GOLD, align=PP_ALIGN.CENTER)
add_text(s, "THÉORIE", 0.5, 2.35, 12.0, 1.2, size=Pt(52), bold=True,
         color=WHITE, align=PP_ALIGN.CENTER)
add_rect(s, 3.5, 3.65, 6.3, 0.07, fill=GOLD)
add_text(s, "45 minutes  ·  Lisez, observez, posez-vous des questions",
         0.5, 3.85, 12.0, 0.55, size=Pt(16), color=MGRAY, align=PP_ALIGN.CENTER)

add_text(s, "\"L'IA ne remplace pas votre expertise.\nElle lui donne des ailes.\"",
         1.5, 4.65, 10.0, 1.2, size=Pt(18), italic=True,
         color=GOLD, align=PP_ALIGN.CENTER)
slide_number(s, 4)

# ════════════════════════════════════════════════════════
# SLIDE 5 — QU'EST-CE QUE L'IA GÉNÉRATIVE ?
# ════════════════════════════════════════════════════════
s = prs.slides.add_slide(blank)
add_rect(s, 0, 0, 13.33, 1.0, fill=NAVY)
add_rect(s, 0, 0, 0.18, 7.5, fill=GOLD)
section_header(s, "THÉORIE  1/6", GOLD)
add_text(s, "Qu'est-ce que l'IA générative ?", 0.4, 0.15, 12.5, 0.7,
         size=Pt(24), bold=True, color=WHITE)

# 2 colonnes
add_rect(s, 0.35, 1.1, 5.9, 5.7, fill=LGRAY)
add_rect(s, 0.35, 1.1, 0.1, 5.7, fill=GOLD)
add_text(s, "🤖  L'IA ANCIENNE", 0.6, 1.2, 5.5, 0.45, size=Pt(13), bold=True, color=NAVY)
add_text(s, "(dite discriminante)", 0.6, 1.62, 5.5, 0.35, size=Pt(10), italic=True, color=RGBColor(0x77,0x77,0x99))
items_old = ["Filtre anti-spam (reconnaît)","Reconnaissance vocale (transcrit)","Recommandation Netflix (prédit)","→  Elle classe et prédit","→  Elle ne crée pas"]
for i,t in enumerate(items_old):
    bold = t.startswith("→")
    col  = RED if bold else DARK
    add_text(s, ("• " if not bold else "") + t,
             0.65, 2.05+i*0.58, 5.4, 0.5, size=Pt(11),
             bold=bold, color=col)

add_rect(s, 6.9, 1.1, 6.05, 5.7, fill=RGBColor(0xE8,0xF4,0xFF))
add_rect(s, 6.9, 1.1, 0.1, 5.7, fill=TEAL)
add_text(s, "✨  L'IA GÉNÉRATIVE", 7.1, 1.2, 5.7, 0.45, size=Pt(13), bold=True, color=NAVY)
add_text(s, "(celle que vous allez utiliser)", 7.1, 1.62, 5.7, 0.35, size=Pt(10), italic=True, color=RGBColor(0x77,0x77,0x99))
items_new = ["Claude / ChatGPT (génèrent du texte)","Canva AI (génère des images)","Synthesia (génère des vidéos)","→  Elle CRÉE du contenu nouveau","→  Jamais deux fois pareil"]
for i,t in enumerate(items_new):
    bold = t.startswith("→")
    col  = TEAL if bold else DARK
    add_text(s, ("• " if not bold else "") + t,
             7.15, 2.05+i*0.58, 5.6, 0.5, size=Pt(11),
             bold=bold, color=col)

add_rect(s, 0.35, 6.88, 12.6, 0.04, fill=GOLD)
add_text(s, "💡  Pour vous : à chaque fois que vous posez une question, l'IA construit une réponse entièrement nouvelle — ce n'est pas du copier-coller.",
         0.4, 6.95, 12.5, 0.45, size=Pt(10), italic=True, color=NAVY)
bottom_bar(s); slide_number(s, 5)

# ════════════════════════════════════════════════════════
# SLIDE 6 — LE LLM : L'ANALOGIE DE L'AGENT ULTRA-LU
# ════════════════════════════════════════════════════════
s = prs.slides.add_slide(blank)
add_rect(s, 0, 0, 13.33, 1.0, fill=NAVY)
add_rect(s, 0, 0, 0.18, 7.5, fill=GOLD)
section_header(s, "THÉORIE  2/6", GOLD)
add_text(s, "Le LLM — Grand Modèle de Langage", 0.4, 0.15, 12.5, 0.7,
         size=Pt(24), bold=True, color=WHITE)

add_rect(s, 0.35, 1.1, 12.6, 1.35, fill=RGBColor(0xFF,0xF5,0xE0))
add_rect(s, 0.35, 1.1, 0.12, 1.35, fill=GOLD)
add_text(s, "🎯  L'ANALOGIE :", 0.6, 1.18, 3.0, 0.38, size=Pt(12), bold=True, color=GOLD)
add_text(s,
    "Imaginez un agent de centre d'appels qui aurait lu 10 millions de conversations téléphoniques, "
    "500 000 manuels de formation, l'encyclopédie entière et tous les livres de management — en 30 langues. "
    "Quand vous lui posez une question, il ne cherche pas dans ses notes : il reconstruit une réponse cohérente, "
    "mot après mot, grâce à tout ce qu'il a intégré.",
    0.6, 1.55, 12.2, 0.8, size=Pt(11), color=DARK)

# 3 LLM cards
llms = [
    (NAVY,  "CLAUDE",   "Anthropic", "Rigueur · Nuance\nIdéal : formation, qualité, ISO"),
    (GREEN, "ChatGPT",  "OpenAI",    "Polyvalence · Écosystème\nIdéal : tous usages"),
    (TEAL,  "Perplexity","Perplexity AI","Recherche internet\nIdéal : veille, benchmarks"),
]
for i,(col,name,maker,desc) in enumerate(llms):
    xl = 0.35 + i*4.32
    add_rect(s, xl, 2.7, 4.1, 2.1, fill=col)
    add_text(s, name,  xl+0.15, 2.82, 3.8, 0.55, size=Pt(20), bold=True, color=WHITE)
    add_text(s, maker, xl+0.15, 3.38, 3.8, 0.35, size=Pt(10), color=RGBColor(0xCC,0xDD,0xFF))
    add_rect(s, xl, 4.05, 4.1, 0.06, fill=WHITE)
    add_text(s, desc,  xl+0.15, 4.18, 3.8, 0.55, size=Pt(10), color=WHITE)

add_text(s, "⭐  VOTRE CHOIX PRINCIPAL", 0.4, 4.98, 4.5, 0.38, size=Pt(11), bold=True, color=NAVY)
add_rect(s, 0.35, 5.38, 4.1, 0.06, fill=NAVY)
add_text(s, "Claude : rigueur ISO, nuance pédagogique, respect des instructions complexes", 0.4, 5.5, 4.5, 0.5, size=Pt(10), color=DARK)

add_text(s, "🔑  LLM = Large Language Model", 5.2, 5.0, 7.8, 0.38, size=Pt(12), bold=True, color=NAVY)
add_text(s,
    "Un programme qui a analysé des milliards de phrases\n"
    "pour apprendre comment les mots et les idées s'enchaînent.\n"
    "Il produit du texte mot à mot, selon les probabilités apprises.",
    5.2, 5.45, 7.8, 0.9, size=Pt(10), color=DARK)

bottom_bar(s); slide_number(s, 6)

# ════════════════════════════════════════════════════════
# SLIDE 7 — LES TOKENS
# ════════════════════════════════════════════════════════
s = prs.slides.add_slide(blank)
add_rect(s, 0, 0, 13.33, 1.0, fill=NAVY)
add_rect(s, 0, 0, 0.18, 7.5, fill=GOLD)
section_header(s, "THÉORIE  3/6", GOLD)
add_text(s, "Les Tokens — Comment l'IA lit le texte", 0.4, 0.15, 12.5, 0.7,
         size=Pt(24), bold=True, color=WHITE)

add_text(s, "Un token = un fragment de mot ou un mot entier", 0.4, 1.15, 12.5, 0.45,
         size=Pt(14), bold=True, color=NAVY)

# Visualisation tokens
phrase_tokens = [
    ("Je ", TEAL), ("suis ", GREEN), ("for", ORANGE), ("ma", RED),
    ("trice ", GOLD), ("en ", TEAL), ("cen", GREEN), ("tres ", ORANGE),
    ("d'ap", RED), ("pels", NAVY)
]
add_text(s, "Exemple :", 0.4, 1.72, 2.0, 0.38, size=Pt(11), bold=True, color=DARK)
xl = 2.5
for tok, col in phrase_tokens:
    w = max(0.45, len(tok)*0.14)
    add_rect(s, xl, 1.7, w, 0.42, fill=col)
    add_text(s, tok.strip(), xl+0.03, 1.73, w-0.05, 0.38,
             size=Pt(11), bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    xl += w + 0.06
add_text(s, "→  10 tokens pour cette phrase",
         xl+0.1, 1.75, 3.0, 0.38, size=Pt(10), italic=True, color=DARK)

# 3 règles pratiques
rules = [
    (TEAL,   "💡  Donnez du contexte",
     "Plus vous donnez d'informations sur votre situation,\nmeilleure sera la réponse de l'IA."),
    (GREEN,  "📏  Limite de mémoire",
     "Claude peut traiter jusqu'à ~200 000 tokens\n≈ un livre entier dans une seule conversation."),
    (ORANGE, "⚠️  Règle pratique",
     "Pour une session normale : restez sous 50 pages\nde texte dans la même conversation."),
]
for i,(col,title,body) in enumerate(rules):
    xl = 0.35 + i*4.32
    add_rect(s, xl, 2.45, 4.1, 2.05, fill=LGRAY)
    add_rect(s, xl, 2.45, 0.12, 2.05, fill=col)
    add_text(s, title, xl+0.25, 2.55, 3.7, 0.42, size=Pt(11), bold=True, color=col)
    add_text(s, body,  xl+0.25, 3.02, 3.7, 0.9,  size=Pt(10), color=DARK)

# Tableau comparatif mémoire
add_text(s, "Capacité mémoire des outils :", 0.4, 4.68, 5.0, 0.38, size=Pt(11), bold=True, color=NAVY)
mem = [("Claude","~200 000 tokens","Livre entier"),
       ("ChatGPT GPT-4o","~128 000 tokens","~400 pages"),
       ("Perplexity","Variable (recherche live)","Pas de limite conversation")]
for i,(tool,tok,eq) in enumerate(mem):
    y = 5.12 + i*0.58
    bg = LGRAY if i%2==0 else WHITE
    add_rect(s, 0.35, y, 12.6, 0.52, fill=bg)
    add_text(s, tool, 0.5, y+0.1, 3.5, 0.35, size=Pt(10), bold=True, color=NAVY)
    add_text(s, tok,  4.2, y+0.1, 4.0, 0.35, size=Pt(10), color=DARK)
    add_text(s, eq,   8.5, y+0.1, 4.2, 0.35, size=Pt(10), italic=True, color=RGBColor(0x66,0x66,0x88))

bottom_bar(s); slide_number(s, 7)

# ════════════════════════════════════════════════════════
# SLIDE 8 — LES PARAMÈTRES
# ════════════════════════════════════════════════════════
s = prs.slides.add_slide(blank)
add_rect(s, 0, 0, 13.33, 1.0, fill=NAVY)
add_rect(s, 0, 0, 0.18, 7.5, fill=GOLD)
section_header(s, "THÉORIE  4/6", GOLD)
add_text(s, "Les Paramètres — La Personnalité de l'IA", 0.4, 0.15, 12.5, 0.7,
         size=Pt(24), bold=True, color=WHITE)

add_text(s, "Vous ne voyez pas ces curseurs — mais vous pouvez les influencer par vos mots.",
         0.4, 1.15, 12.5, 0.42, size=Pt(13), color=DARK, italic=True)

# Curseur visuel
add_text(s, "🌡  TEMPÉRATURE  (créativité / rigueur)", 0.4, 1.72, 7.0, 0.4, size=Pt(12), bold=True, color=NAVY)
add_rect(s, 0.4, 2.18, 8.5, 0.18, fill=MGRAY)
# Gradient visuel simulé
colors_grad = [NAVY, TEAL, GREEN, GOLD, ORANGE, RED]
for i,c in enumerate(colors_grad):
    add_rect(s, 0.4+i*1.42, 2.18, 1.42, 0.18, fill=c)
for lbl,xl in [("RIGOUREUX",0.4),("ÉQUILIBRÉ",3.5),("CRÉATIF",7.0)]:
    add_text(s, lbl, xl, 2.42, 2.0, 0.32, size=Pt(9), bold=True, color=DARK)

# Tableau d'influence
rows = [
    ("Vous écrivez…",             "L'IA répond avec…",           "Usage idéal pour vous"),
    ('"Sois précis et factuel"',  "Rigueur, peu de créativité",  "Grilles QA, rapports ISO"),
    ('"Sois créatif et innovant"',"Variété, surprises",          "Jeux de rôle, icebreakers"),
    ('"Propose 3 variantes"',     "Diversité maximale",          "Emails prospection, intros"),
]
for i,row in enumerate(rows):
    y = 2.88 + i*0.72
    bg = NAVY if i==0 else (LGRAY if i%2==0 else WHITE)
    add_rect(s, 0.35, y, 12.6, 0.66, fill=bg)
    for j,(cell, w, xl) in enumerate(zip(row,[4.1,4.2,4.1],[0.5,4.75,9.1])):
        fc = WHITE if i==0 else DARK
        bold_ = (i==0)
        add_text(s, cell, xl, y+0.14, w, 0.42,
                 size=Pt(10 if i>0 else 11), bold=bold_, color=fc)

add_text(s, "💡  Retenir :", 0.4, 5.8, 2.0, 0.38, size=Pt(12), bold=True, color=NAVY)
add_text(s, "Vous contrôlez la personnalité de l'IA par le vocabulaire que vous utilisez dans vos instructions.",
         2.5, 5.8, 10.5, 0.38, size=Pt(11), color=DARK)
add_text(s, "→  ISO / audit → mots clés : \"précis\", \"conforme\", \"selon la norme\"",
         0.5, 6.25, 12.5, 0.35, size=Pt(10), color=NAVY, bold=True)
add_text(s, "→  Formation / créativité → mots clés : \"innovant\", \"ludique\", \"variantes\", \"surprenant\"",
         0.5, 6.62, 12.5, 0.35, size=Pt(10), color=NAVY, bold=True)

bottom_bar(s); slide_number(s, 8)

# ════════════════════════════════════════════════════════
# SLIDE 9 — CE QUE L'IA PEUT / NE PEUT PAS
# ════════════════════════════════════════════════════════
s = prs.slides.add_slide(blank)
add_rect(s, 0, 0, 13.33, 1.0, fill=NAVY)
add_rect(s, 0, 0, 0.18, 7.5, fill=GOLD)
section_header(s, "THÉORIE  5/6", GOLD)
add_text(s, "Ce que l'IA peut — et ne peut pas — faire", 0.4, 0.15, 12.5, 0.7,
         size=Pt(24), bold=True, color=WHITE)

add_text(s, "👉  C'est ici que votre expertise de 15 ans reste IRREMPLAÇABLE.",
         0.4, 1.1, 12.5, 0.42, size=Pt(13), bold=True, color=RED)

can = [
    "Rédiger vite et structurer logiquement",
    "Proposer des variantes et des idées",
    "Synthétiser de l'information volumineuse",
    "Générer des cas pratiques et des exemples",
    "Respecter une structure et un format imposés",
    "Adapter le ton selon vos instructions",
]
cannot = [
    "Valider si le résultat est juste pour VOTRE client",
    "Connaître la culture de votre entreprise",
    "Ressentir les dynamiques d'équipe",
    "Juger si un agent est réellement prêt",
    "Adapter à la personnalité d'un apprenant",
    "Prendre la responsabilité d'une décision",
]

add_rect(s, 0.35, 1.65, 5.9, 0.48, fill=GREEN)
add_text(s, "✅  L'IA PEUT", 0.5, 1.68, 5.6, 0.42, size=Pt(13), bold=True, color=WHITE)
add_rect(s, 7.1, 1.65, 5.9, 0.48, fill=RED)
add_text(s, "❌  L'IA NE PEUT PAS", 7.25, 1.68, 5.6, 0.42, size=Pt(13), bold=True, color=WHITE)

for i,(c,nc) in enumerate(zip(can, cannot)):
    y = 2.22 + i*0.69
    bg = LGRAY if i%2==0 else WHITE
    add_rect(s, 0.35, y, 5.9, 0.62, fill=bg)
    add_text(s, "✅  "+c, 0.5, y+0.1, 5.6, 0.5, size=Pt(10), color=DARK)
    add_rect(s, 7.1, y, 5.9, 0.62, fill=bg)
    add_text(s, "❌  "+nc, 7.25, y+0.1, 5.6, 0.5, size=Pt(10), color=DARK)

add_rect(s, 6.25, 1.65, 0.6, 4.15+0.07, fill=GOLD)
add_text(s, "VS", 6.28, 3.5, 0.55, 0.55, size=Pt(14), bold=True, color=WHITE, align=PP_ALIGN.CENTER)

add_rect(s, 0.35, 6.45, 12.6, 0.52, fill=NAVY)
add_text(s, "🏆  LA FORMULE GAGNANTE :", 0.5, 6.5, 4.0, 0.42, size=Pt(11), bold=True, color=GOLD)
add_text(s, "IA rapide  +  Expertise humaine juste  =  Résultat exceptionnel en peu de temps",
         4.6, 6.5, 8.5, 0.42, size=Pt(12), bold=True, color=WHITE)

bottom_bar(s); slide_number(s, 9)

# ════════════════════════════════════════════════════════
# SLIDE 10 — 5 POINTS CLÉS
# ════════════════════════════════════════════════════════
s = prs.slides.add_slide(blank)
add_rect(s, 0, 0, 13.33, 1.0, fill=NAVY)
add_rect(s, 0, 0, 0.18, 7.5, fill=GOLD)
section_header(s, "THÉORIE  6/6", GOLD)
add_text(s, "5 Points Clés à Mémoriser", 0.4, 0.15, 12.5, 0.7,
         size=Pt(24), bold=True, color=WHITE)

points = [
    (RED,    "1", "L'IA générative CRÉE",   "Elle ne copie pas — elle produit du contenu nouveau à chaque fois"),
    (ORANGE, "2", "Le LLM a tout lu",        "Il reconstruit une réponse mot par mot grâce à des milliards de textes appris"),
    (GOLD,   "3", "Les tokens = le carburant","Donnez du contexte → l'IA produit mieux. Plus vous êtes précise, plus c'est juste"),
    (GREEN,  "4", "Les paramètres obéissent", "Vos mots influencent le ton, la rigueur, la créativité de l'IA"),
    (TEAL,   "5", "Votre expertise valide",  "L'IA produit vite — vous produisez juste. Ensemble : vite ET juste"),
]
for i,(col,num,title,body) in enumerate(points):
    y = 1.18 + i*1.1
    add_rect(s, 0.35, y, 12.6, 0.98, fill=LGRAY)
    add_rect(s, 0.35, y, 0.12, 0.98, fill=col)
    add_rect(s, 0.55, y+0.22, 0.56, 0.56, fill=col)
    add_text(s, num, 0.56, y+0.22, 0.56, 0.56, size=Pt(20), bold=True,
             color=WHITE, align=PP_ALIGN.CENTER)
    add_text(s, title, 1.28, y+0.12, 4.5, 0.4, size=Pt(13), bold=True, color=col)
    add_text(s, body,  1.28, y+0.55, 11.2, 0.4, size=Pt(10), color=DARK)

bottom_bar(s); slide_number(s, 10)

# ════════════════════════════════════════════════════════
# SLIDE 11 — INTERLUDE EXERCICE
# ════════════════════════════════════════════════════════
s = prs.slides.add_slide(blank)
add_rect(s, 0, 0, 13.33, 7.5, fill=TEAL)
add_rect(s, 0, 0, 0.18, 7.5, fill=WHITE)
add_text(s, "PARTIE 2", 0.5, 1.5, 12.0, 0.8, size=Pt(20), bold=True,
         color=WHITE, align=PP_ALIGN.CENTER)
add_text(s, "EXERCICE", 0.5, 2.35, 12.0, 1.2, size=Pt(52), bold=True,
         color=WHITE, align=PP_ALIGN.CENTER)
add_rect(s, 3.5, 3.65, 6.3, 0.07, fill=WHITE)
add_text(s, "75 minutes  ·  Créer vos 5 comptes  ·  Configurer Notion",
         0.5, 3.85, 12.0, 0.55, size=Pt(16), color=RGBColor(0xCC,0xF0,0xF8), align=PP_ALIGN.CENTER)
add_text(s, "🎯  Objectif : être opérationnelle sur vos 5 outils essentiels",
         1.5, 4.7, 10.0, 0.55, size=Pt(16), bold=True, color=WHITE, align=PP_ALIGN.CENTER)
slide_number(s, 11)

# ════════════════════════════════════════════════════════
# SLIDE 12 — LES 5 COMPTES
# ════════════════════════════════════════════════════════
s = prs.slides.add_slide(blank)
add_rect(s, 0, 0, 13.33, 1.0, fill=TEAL)
add_rect(s, 0, 0, 0.18, 7.5, fill=WHITE)
section_header(s, "EXERCICE  1/2", TEAL)
add_text(s, "Créer vos 5 Comptes Essentiels", 0.4, 0.15, 12.5, 0.7,
         size=Pt(24), bold=True, color=WHITE)

comptes = [
    ("1","🤖","CLAUDE","claude.ai","Votre outil principal — formation, qualité, ISO", NAVY),
    ("2","💬","CHATGPT","chat.openai.com","Polyvalent, large écosystème d'outils", GREEN),
    ("3","🔍","PERPLEXITY","perplexity.ai","Recherche internet avec sources citées", ORANGE),
    ("4","📋","NOTION","notion.so","Votre journal de bord et portfolio IA", RED),
    ("5","🎨","CANVA","canva.com","Supports de formation visuels IA", TEAL),
]
for i,(num,icon,name,url,usage,col) in enumerate(comptes):
    y = 1.15 + i*1.1
    add_rect(s, 0.35, y, 12.6, 0.98, fill=LGRAY)
    add_rect(s, 0.35, y, 0.12, 0.98, fill=col)
    add_rect(s, 0.55, y+0.22, 0.56, 0.56, fill=col)
    add_text(s, num, 0.56, y+0.22, 0.56, 0.56, size=Pt(18), bold=True,
             color=WHITE, align=PP_ALIGN.CENTER)
    add_text(s, icon+" "+name, 1.25, y+0.12, 2.8, 0.4, size=Pt(13), bold=True, color=col)
    pill(s, url, 4.2, y+0.24, 3.2, 0.42, bg=col, size=Pt(10))
    add_text(s, usage, 7.6, y+0.28, 5.2, 0.42, size=Pt(10), color=DARK)

add_rect(s, 0.35, 6.72, 12.6, 0.04, fill=TEAL)
add_text(s, "💡  Utilisez la même adresse email partout — plus simple à gérer. Commencez par Claude.",
         0.4, 6.8, 12.5, 0.38, size=Pt(10), italic=True, color=NAVY)

bottom_bar(s); slide_number(s, 12)

# ════════════════════════════════════════════════════════
# SLIDE 13 — PREMIER TEST CLAUDE
# ════════════════════════════════════════════════════════
s = prs.slides.add_slide(blank)
add_rect(s, 0, 0, 13.33, 1.0, fill=TEAL)
add_rect(s, 0, 0, 0.18, 7.5, fill=WHITE)
section_header(s, "EXERCICE  1/2", TEAL)
add_text(s, "Premier Test sur Claude — À faire maintenant", 0.4, 0.15, 12.5, 0.7,
         size=Pt(22), bold=True, color=WHITE)

add_text(s, "Dès que votre compte Claude est créé, copiez-collez ce prompt :", 0.4, 1.1, 12.5, 0.4,
         size=Pt(12), bold=True, color=NAVY)

add_rect(s, 0.35, 1.6, 12.6, 2.6, fill=DARK)
add_rect(s, 0.35, 1.6, 0.15, 2.6, fill=RED)
prompt_lines = [
    "Bonjour Claude. Je suis formatrice en centres d'appels",
    "avec 15 ans d'expérience. Je commence aujourd'hui",
    "à apprendre l'IA.",
    "",
    "Présente-toi en 5 lignes et explique-moi",
    "comment tu peux m'aider dans mon métier.",
]
for i,line in enumerate(prompt_lines):
    add_text(s, line, 0.65, 1.72+i*0.37, 12.1, 0.38,
             size=Pt(12), color=RGBColor(0xA8,0xF0,0xE0))

add_text(s, "Ce que vous observez :", 0.4, 4.38, 3.5, 0.38, size=Pt(12), bold=True, color=TEAL)
obs = [
    "→  Le ton : formel ? chaleureux ? professionnel ?",
    "→  Les exemples : sont-ils vraiment liés aux centres d'appels ?",
    "→  La précision : est-ce que ça correspond à votre réalité terrain ?",
    "→  Notez TOUT dans votre journal Notion — même ce qui semble évident",
]
for i,o in enumerate(obs):
    add_text(s, o, 0.5, 4.82+i*0.48, 12.5, 0.42, size=Pt(11), color=DARK)

add_rect(s, 0.35, 6.78, 12.6, 0.04, fill=TEAL)
add_text(s, "⚠️  Votre œil d'expert(e) va détecter des imprécisions — c'est NORMAL et c'est votre valeur ajoutée.",
         0.4, 6.85, 12.5, 0.38, size=Pt(10), bold=True, color=RED)
bottom_bar(s); slide_number(s, 13)

# ════════════════════════════════════════════════════════
# SLIDE 14 — CONFIGURER NOTION
# ════════════════════════════════════════════════════════
s = prs.slides.add_slide(blank)
add_rect(s, 0, 0, 13.33, 1.0, fill=TEAL)
add_rect(s, 0, 0, 0.18, 7.5, fill=WHITE)
section_header(s, "EXERCICE  2/2", TEAL)
add_text(s, "Configurer votre Espace Notion — TAMOU NEURAL PATH", 0.4, 0.15, 12.5, 0.7,
         size=Pt(22), bold=True, color=WHITE)

steps = [
    (NAVY,  "Créer l'espace",         "Nouvelle page → Titre : \"TAMOU NEURAL PATH — Mon Parcours IA\"\nAjoutez l'icône 🧠"),
    (TEAL,  "Journal de Bord",        "Base de données Tableau → colonnes : Date / Outil / Produit /\nMeilleur Prompt / Ce qui n'a pas marché / Découverte / Niveau /10"),
    (GREEN, "Portfolio Livrables",    "Nouvelle page → 5 phases avec checkboxes\nCochez ✅ au fur et à mesure de vos productions"),
    (RED,   "Bibliothèque de Prompts","Base de données → colonnes : Nom / Thème / Prompt complet /\nRésultat / Note /5 — votre capital de productivité"),
]
for i,(col,title,body) in enumerate(steps):
    y = 1.2 + i*1.4
    add_rect(s, 0.35, y, 12.6, 1.28, fill=LGRAY)
    add_rect(s, 0.35, y, 0.12, 1.28, fill=col)
    add_rect(s, 0.55, y+0.36, 0.5, 0.5, fill=col)
    add_text(s, str(i+1), 0.56, y+0.36, 0.5, 0.5, size=Pt(16), bold=True,
             color=WHITE, align=PP_ALIGN.CENTER)
    add_text(s, title, 1.22, y+0.12, 4.2, 0.42, size=Pt(13), bold=True, color=col)
    add_text(s, body,  1.22, y+0.58, 11.2, 0.62, size=Pt(10), color=DARK)

bottom_bar(s); slide_number(s, 14)

# ════════════════════════════════════════════════════════
# SLIDE 15 — TEMPLATE JOURNAL DE BORD
# ════════════════════════════════════════════════════════
s = prs.slides.add_slide(blank)
add_rect(s, 0, 0, 13.33, 1.0, fill=TEAL)
add_rect(s, 0, 0, 0.18, 7.5, fill=WHITE)
section_header(s, "EXERCICE  2/2", TEAL)
add_text(s, "Template Journal de Bord — À remplir chaque soir", 0.4, 0.15, 12.5, 0.7,
         size=Pt(22), bold=True, color=WHITE)

fields = [
    ("📅  DATE",               "Cliquez sur la date dans Notion",                    NAVY),
    ("🛠  OUTIL UTILISÉ",      "Claude / ChatGPT / Perplexity / Notion / Canva",     TEAL),
    ("📄  CE QUE J'AI PRODUIT","Décrivez le livrable — soyez précise",               GREEN),
    ("⭐  MEILLEUR PROMPT",    "Copiez-collez le prompt qui a donné le meilleur résultat", GOLD),
    ("❌  CE QUI N'A PAS MARCHÉ","Notez honnêtement — c'est votre meilleur outil de progression", RED),
    ("💡  DÉCOUVERTE DU JOUR", "La chose la plus surprenante que vous avez apprise", ORANGE),
    ("📊  NIVEAU CONFIANCE /10","De 1 (perdu) à 10 (à l'aise) — sans jugement",      TEAL),
]
for i,(label,hint,col) in enumerate(fields):
    y = 1.12 + i*0.82
    add_rect(s, 0.35, y, 12.6, 0.72, fill=LGRAY if i%2==0 else WHITE)
    add_rect(s, 0.35, y, 0.1, 0.72, fill=col)
    add_text(s, label, 0.6, y+0.08, 3.8, 0.38, size=Pt(11), bold=True, color=col)
    add_text(s, hint,  4.6, y+0.12, 8.2, 0.42, size=Pt(10), italic=True, color=DARK)

add_rect(s, 0.35, 6.9, 12.6, 0.04, fill=TEAL)
add_text(s, "💡  Ce journal est votre preuves de progression — et une partie de votre portfolio freelance.",
         0.4, 6.95, 12.5, 0.38, size=Pt(10), italic=True, color=NAVY)
bottom_bar(s); slide_number(s, 15)

# ════════════════════════════════════════════════════════
# SLIDE 16 — INTERLUDE APRÈS-MIDI
# ════════════════════════════════════════════════════════
s = prs.slides.add_slide(blank)
add_rect(s, 0, 0, 13.33, 7.5, fill=ORANGE)
add_rect(s, 0, 0, 0.18, 7.5, fill=WHITE)
add_text(s, "PARTIE 3", 0.5, 1.5, 12.0, 0.8, size=Pt(20), bold=True,
         color=WHITE, align=PP_ALIGN.CENTER)
add_text(s, "PRODUCTION", 0.5, 2.35, 12.0, 1.2, size=Pt(48), bold=True,
         color=WHITE, align=PP_ALIGN.CENTER)
add_rect(s, 3.5, 3.65, 6.3, 0.07, fill=WHITE)
add_text(s, "2 heures  ·  Dialogue guidé  ·  Fiche comparative",
         0.5, 3.85, 12.0, 0.55, size=Pt(16), color=RGBColor(0xFF,0xE5,0xCC), align=PP_ALIGN.CENTER)
add_text(s, "🎯  Objectif : Produire votre 1er livrable du portfolio TAMOU NEURAL PATH",
         1.5, 4.7, 10.0, 0.55, size=Pt(15), bold=True, color=WHITE, align=PP_ALIGN.CENTER)
slide_number(s, 16)

# ════════════════════════════════════════════════════════
# SLIDE 17 — 6 PROMPTS GUIDÉS CLAUDE
# ════════════════════════════════════════════════════════
s = prs.slides.add_slide(blank)
add_rect(s, 0, 0, 13.33, 1.0, fill=ORANGE)
add_rect(s, 0, 0, 0.18, 7.5, fill=WHITE)
section_header(s, "PRODUCTION  1h", ORANGE)
add_text(s, "Dialogue guidé avec Claude — 6 Questions Séquencées", 0.4, 0.15, 12.5, 0.7,
         size=Pt(22), bold=True, color=WHITE)

add_text(s, "Posez ces 6 questions UNE PAR UNE. Lisez chaque réponse complète avant de continuer.",
         0.4, 1.08, 12.5, 0.38, size=Pt(11), bold=True, color=NAVY)

prompts = [
    ("1","Connaissance métier","Demandez à Claude de vous expliquer en quoi il peut transformer votre quotidien\n→ Observez la précision, les exemples terrain"),
    ("2","Objectif pédagogique","Demandez un objectif SMART pour \"Gérer un client mécontent\"\n→ Vérifiez la méthode SMART avec votre expertise"),
    ("3","ISO 9001","Demandez les éléments du chapitre 8 ISO 9001 appliqués à la formation\n→ Utilisez votre certification pour détecter les erreurs"),
    ("4","Message de motivation","Demandez un message 150 mots pour une équipe après une semaine difficile\n→ Jugez le ton : humain ou artificiel ?"),
    ("5","Plan formation 3 jours","Demandez un plan détaillé pour 8 superviseurs de centres d'appels sortants\n→ Vérifiez la cohérence pédagogique"),
    ("6","Itération","Répondez : \"Le jour 1 est trop chargé, allège et ajoute un brise-glace\"\n→ Observe-t-il votre retour sans tout refaire ?"),
]
for i,(num,theme,desc) in enumerate(prompts):
    col  = [NAVY,TEAL,GREEN,ORANGE,RED,GOLD][i]
    xl   = 0.35 + (i%3)*4.32
    y    = 1.58 + (i//3)*2.42
    add_rect(s, xl, y, 4.1, 2.18, fill=LGRAY)
    add_rect(s, xl, y, 0.12, 2.18, fill=col)
    add_rect(s, xl+0.18, y+0.12, 0.46, 0.46, fill=col)
    add_text(s, num,   xl+0.19, y+0.12, 0.46, 0.46, size=Pt(14), bold=True,
             color=WHITE, align=PP_ALIGN.CENTER)
    add_text(s, theme, xl+0.75, y+0.14, 3.2, 0.38, size=Pt(10), bold=True, color=col)
    add_text(s, desc,  xl+0.2, y+0.65, 3.75, 1.42, size=Pt(9), color=DARK)

bottom_bar(s); slide_number(s, 17)

# ════════════════════════════════════════════════════════
# SLIDE 18 — MÊME CHOSE AVEC CHATGPT
# ════════════════════════════════════════════════════════
s = prs.slides.add_slide(blank)
add_rect(s, 0, 0, 13.33, 1.0, fill=ORANGE)
add_rect(s, 0, 0, 0.18, 7.5, fill=WHITE)
section_header(s, "PRODUCTION  30 min", ORANGE)
add_text(s, "Même Exercice dans ChatGPT — Comparer à périmètre égal", 0.4, 0.15, 12.5, 0.7,
         size=Pt(22), bold=True, color=WHITE)

add_text(s, "Copiez-collez exactement les mêmes 6 questions dans ChatGPT. N'adaptez rien.",
         0.4, 1.1, 12.5, 0.42, size=Pt(12), bold=True, color=NAVY)

add_rect(s, 0.35, 1.62, 5.9, 4.85, fill=RGBColor(0xE8,0xF8,0xFF))
add_rect(s, 0.35, 1.62, 0.12, 4.85, fill=NAVY)
add_text(s, "🤖  CLAUDE — Points à noter", 0.6, 1.75, 5.5, 0.42, size=Pt(12), bold=True, color=NAVY)
claude_pts = [
    "Niveau de précision sur le métier",
    "Respect des instructions complexes",
    "Intégration de vos corrections (Q6)",
    "Ton et qualité rédactionnelle",
    "Exemples centres d'appels",
    "Rigueur sur l'ISO 9001",
]
for i,p in enumerate(claude_pts):
    add_text(s, f"□  {p}", 0.6, 2.28+i*0.58, 5.5, 0.5, size=Pt(10), color=DARK)

add_rect(s, 7.1, 1.62, 5.9, 4.85, fill=RGBColor(0xF0,0xFF,0xF0))
add_rect(s, 7.1, 1.62, 0.12, 4.85, fill=GREEN)
add_text(s, "💬  CHATGPT — Points à noter", 7.35, 1.75, 5.5, 0.42, size=Pt(12), bold=True, color=GREEN)
for i,p in enumerate(claude_pts):
    add_text(s, f"□  {p}", 7.35, 2.28+i*0.58, 5.5, 0.5, size=Pt(10), color=DARK)

add_rect(s, 6.25, 1.62, 0.6, 4.85, fill=ORANGE)
add_text(s, "VS", 6.28, 3.9, 0.55, 0.55, size=Pt(14), bold=True, color=WHITE, align=PP_ALIGN.CENTER)

add_rect(s, 0.35, 6.58, 12.6, 0.52, fill=LGRAY)
add_text(s, "💡  Pas de bonne ou mauvaise réponse — les deux outils sont différents, pas concurrents. Vous les utiliserez pour des tâches différentes.",
         0.5, 6.64, 12.4, 0.42, size=Pt(10), italic=True, color=DARK)

bottom_bar(s); slide_number(s, 18)

# ════════════════════════════════════════════════════════
# SLIDE 19 — FICHE COMPARATIVE (LIVRABLE 1)
# ════════════════════════════════════════════════════════
s = prs.slides.add_slide(blank)
add_rect(s, 0, 0, 13.33, 1.0, fill=GREEN)
add_rect(s, 0, 0, 0.18, 7.5, fill=WHITE)
section_header(s, "LIVRABLE  ✅", GREEN)
add_text(s, "Fiche Comparative Claude vs ChatGPT — Livrable 1", 0.4, 0.15, 12.5, 0.7,
         size=Pt(22), bold=True, color=WHITE)

add_rect(s, 0.35, 1.1, 12.6, 0.42, fill=GREEN)
headers = ["CRITÈRE", "CLAUDE /5", "CHATGPT /5", "MON CHOIX"]
widths  = [5.2, 2.3, 2.3, 2.7]
positions = [0.5, 5.85, 8.3, 10.75]
for h,w,xl in zip(headers,widths,positions):
    add_text(s, h, xl, 1.15, w, 0.34, size=Pt(10), bold=True, color=WHITE)

criteria = [
    "Connaissance centres d'appels",
    "Qualité des contenus formation",
    "Rigueur sur les normes ISO",
    "Ton et qualité rédactionnelle",
    "Intégration de mes corrections",
    "Pertinence des exemples métier",
]
for i,crit in enumerate(criteria):
    y = 1.58 + i*0.68
    bg = LGRAY if i%2==0 else WHITE
    add_rect(s, 0.35, y, 12.6, 0.62, fill=bg)
    add_text(s, crit, 0.5, y+0.12, 5.1, 0.42, size=Pt(10), color=DARK)
    for xl in [5.85, 8.3, 10.75]:
        add_rect(s, xl, y+0.1, 2.0, 0.44, fill=WHITE)
        add_text(s, "__ / 5" if xl < 10 else "Claude  /  ChatGPT",
                 xl+0.1, y+0.14, 1.8, 0.34, size=Pt(9),
                 color=RGBColor(0x99,0x99,0xBB), italic=True)

add_rect(s, 0.35, 5.72, 12.6, 0.5, fill=NAVY)
add_text(s, "TOTAL", 0.5, 5.8, 5.1, 0.38, size=Pt(11), bold=True, color=WHITE)
add_text(s, "__ / 30", 5.85, 5.8, 2.0, 0.38, size=Pt(11), bold=True, color=GOLD)
add_text(s, "__ / 30", 8.3, 5.8, 2.0, 0.38, size=Pt(11), bold=True, color=GOLD)

add_rect(s, 0.35, 6.32, 12.6, 0.04, fill=GREEN)
add_text(s, "📌  Sauvegardez ce tableau dans Notion → Portfolio → Livrable 1  ✅  Premier livrable validé !",
         0.4, 6.42, 12.5, 0.42, size=Pt(11), bold=True, color=NAVY)

bottom_bar(s); slide_number(s, 19)

# ════════════════════════════════════════════════════════
# SLIDE 20 — BILAN DU JOUR
# ════════════════════════════════════════════════════════
s = prs.slides.add_slide(blank)
add_rect(s, 0, 0, 13.33, 1.0, fill=NAVY)
add_rect(s, 0, 0, 0.18, 7.5, fill=GREEN)
add_text(s, "BILAN DU JOUR 1", 0.4, 0.15, 12.5, 0.7, size=Pt(26), bold=True, color=WHITE)

add_text(s, "Ce que vous avez accompli aujourd'hui :", 0.4, 1.12, 12.5, 0.42, size=Pt(14), bold=True, color=NAVY)

accomplishments = [
    "Compris ce qu'est un LLM, un token et un paramètre — sans jargon technique",
    "Créé vos 5 comptes essentiels (Claude, ChatGPT, Perplexity, Notion, Canva)",
    "Configuré votre espace de travail Notion — journal, portfolio, bibliothèque de prompts",
    "Testé Claude et ChatGPT sur 6 cas concrets de votre métier",
    "Produit votre LIVRABLE 1 — Fiche comparative Claude vs ChatGPT  ✅",
    "Identifié vos premières erreurs de l'IA grâce à votre expertise terrain",
]
for i,a in enumerate(accomplishments):
    y = 1.65 + i*0.75
    add_rect(s, 0.35, y, 12.6, 0.64, fill=LGRAY if i%2==0 else WHITE)
    add_rect(s, 0.35, y, 0.12, 0.64, fill=GREEN)
    add_text(s, "✅  "+a, 0.6, y+0.12, 12.2, 0.42, size=Pt(11), color=DARK)

add_rect(s, 0.35, 6.3, 12.6, 0.06, fill=GREEN)
add_text(s, "📊  Remplissez votre journal de bord ce soir avant de fermer votre ordinateur",
         0.4, 6.42, 12.5, 0.38, size=Pt(11), bold=True, color=NAVY)
add_text(s, "Niveau de confiance : __ / 10  |  Meilleur prompt de la journée : sauvegardé dans Notion ?  □",
         0.4, 6.82, 12.5, 0.38, size=Pt(10), color=DARK)

bottom_bar(s); slide_number(s, 20)

# ════════════════════════════════════════════════════════
# SLIDE 21 — DEMAIN JOUR 2
# ════════════════════════════════════════════════════════
s = prs.slides.add_slide(blank)
add_rect(s, 0, 0, 13.33, 1.0, fill=NAVY)
add_rect(s, 0, 0, 0.18, 7.5, fill=RED)
add_text(s, "DEMAIN — JOUR 2 : La Formule Magique des Prompts", 0.4, 0.15, 12.5, 0.7,
         size=Pt(22), bold=True, color=WHITE)

add_text(s, "Vous allez apprendre la méthode RCTF et produire 10 prompts professionnels prêts à l'emploi.",
         0.4, 1.1, 12.5, 0.42, size=Pt(12), color=DARK)

add_text(s, "La Méthode RCTF — Aperçu", 0.4, 1.68, 5.0, 0.42, size=Pt(13), bold=True, color=NAVY)
rctf = [
    (RED,    "R", "RÔLE",    "Tu es expert(e) qualité en centres d'appels…"),
    (ORANGE, "C", "CONTEXTE","Dans le cadre d'un audit ISO 9001 v2015…"),
    (GREEN,  "T", "TÂCHE",   "Crée une grille d'évaluation appels entrants…"),
    (TEAL,   "F", "FORMAT",  "Format : tableau avec 5 colonnes, 7 catégories…"),
]
for i,(col,letter,label,ex) in enumerate(rctf):
    y = 2.2 + i*1.1
    add_rect(s, 0.35, y, 5.9, 0.98, fill=LGRAY)
    add_rect(s, 0.35, y, 0.12, 0.98, fill=col)
    add_rect(s, 0.55, y+0.22, 0.56, 0.56, fill=col)
    add_text(s, letter, 0.56, y+0.22, 0.56, 0.56, size=Pt(20), bold=True,
             color=WHITE, align=PP_ALIGN.CENTER)
    add_text(s, label, 1.26, y+0.1, 1.8, 0.38, size=Pt(12), bold=True, color=col)
    add_text(s, ex,    3.2, y+0.14, 2.9, 0.72, size=Pt(9.5), color=DARK, italic=True)

add_rect(s, 6.6, 1.68, 6.38, 4.7, fill=DARK)
add_rect(s, 6.6, 1.68, 0.15, 4.7, fill=RED)
add_text(s, "EXEMPLE — Livrable Jour 2", 6.88, 1.78, 5.9, 0.42, size=Pt(11), bold=True, color=RED)
ex_lines = [
    "Tu es expert qualité en centres",
    "d'appels, certifié ISO 9001 v2015.",
    "",
    "Dans le contexte du contrôle qualité",
    "d'un service client de 50 agents",
    "traitant 500 appels/jour.",
    "",
    "Crée une grille d'évaluation complète",
    "compatible ISO 9001.",
    "",
    "Format : tableau 5 colonnes,",
    "7 catégories, total 100 points.",
]
for i,line in enumerate(ex_lines):
    add_text(s, line, 6.88, 2.26+i*0.3, 5.9, 0.28, size=Pt(9.5),
             color=RGBColor(0xA8,0xF0,0xE0))

add_rect(s, 0.35, 6.42, 12.6, 0.5, fill=NAVY)
add_text(s, "🎯  À la fin du Jour 2 : 10 prompts professionnels dans votre bibliothèque — utilisables immédiatement avec vos clients",
         0.5, 6.5, 12.4, 0.38, size=Pt(11), bold=True, color=WHITE)
bottom_bar(s); slide_number(s, 21)

# ════════════════════════════════════════════════════════
# SLIDE 22 — CLÔTURE
# ════════════════════════════════════════════════════════
s = prs.slides.add_slide(blank)
add_rect(s, 0, 0, 13.33, 7.5, fill=DARK)
add_rect(s, 0, 0, 0.18, 7.5, fill=GOLD)
add_rect(s, 0, 7.05, 13.33, 0.45, fill=NAVY)

add_text(s, "TAMOU NEURAL PATH",
         0.5, 0.9, 12.0, 0.8, size=Pt(36), bold=True, color=WHITE, align=PP_ALIGN.CENTER)
add_rect(s, 3.5, 1.82, 6.3, 0.07, fill=GOLD)
add_text(s, "Jour 1 terminé  ·  1 livrable validé  ·  77 jours devant vous",
         0.5, 2.0, 12.0, 0.55, size=Pt(16), color=MGRAY, align=PP_ALIGN.CENTER)

add_text(s,
    "\"Vous venez de faire quelque chose que beaucoup de professionnels\n"
    "expérimentés n'ont pas encore fait :\n"
    "vous avez ouvert la porte de l'IA avec votre expertise comme boussole.\"",
    1.2, 2.85, 10.9, 1.6, size=Pt(17), italic=True,
    color=GOLD, align=PP_ALIGN.CENTER)

add_rect(s, 3.5, 4.58, 6.3, 0.07, fill=RGBColor(0x44,0x44,0x66))

for i,(icon,txt) in enumerate([("🤖","L'IA produit vite"),("👩‍💼","Vous produisez juste"),("🏆","Ensemble : exceptionnel")]):
    xl = 1.5 + i*3.6
    add_rect(s, xl, 4.78, 3.2, 1.15, fill=NAVY)
    add_text(s, icon, xl+0.1, 4.88, 0.6, 0.6, size=Pt(24))
    add_text(s, txt,  xl+0.75, 5.05, 2.3, 0.55, size=Pt(13), bold=True, color=WHITE)

add_text(s, "TAMOU NEURAL PATH  |  Jour 1 sur 78  |  Phase 1 sur 5",
         0.3, 7.1, 12.7, 0.3, size=Pt(9), color=MGRAY, align=PP_ALIGN.CENTER)
slide_number(s, 22)

# ════════════════════════════════════════════════════════
# SAVE
# ════════════════════════════════════════════════════════
output = "/home/user/fatou/TAMOU_NEURAL_PATH_Jour1_Lundi.pptx"
prs.save(output)
print(f"PPTX généré : {output}")

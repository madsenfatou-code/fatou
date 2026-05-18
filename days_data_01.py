# Days 2–30 data
from pptx_helpers import *

DAYS = []

# ════════════════════════════════════════════════
# PHASE 1 — FONDATIONS (Semaines 1-2)
# ════════════════════════════════════════════════

DAYS.append(dict(
    day_num=2, week=1, phase=1, phase_name="Fondations IA", day_name="Mardi",
    title="La Formule Magique des Prompts", subtitle="Méthode RCTF — Rôle · Contexte · Tâche · Format",
    accent_color=TEAL,
    theory_title="La Méthode RCTF",
    theory_points=[
        ("R — RÔLE", "Définissez qui est l'IA : 'Tu es expert qualité en centres d'appels...'"),
        ("C — CONTEXTE", "Décrivez la situation : 'Dans le cadre d'un audit ISO 9001...'"),
        ("T — TÂCHE", "Soyez précis sur ce que vous voulez : 'Crée une grille d'évaluation...'"),
        ("F — FORMAT", "Imposez la structure : 'Format : tableau 5 colonnes, total 100 pts'"),
        "Plus vous êtes précise, meilleure est la réponse — l'IA n'est pas devin",
        "Un bon prompt = une bonne commande passée à un stagiaire brillant",
    ],
    theory_insight="La méthode RCTF réduit de 80% les réponses inutilisables. C'est votre outil #1.",
    exercise_title="Écrire 10 Prompts Métier avec RCTF",
    exercise_steps=[
        ("📝","Prompt 1 : Grille QA","Appliquez RCTF pour générer une grille d'évaluation appels entrants — testez dans Claude",NAVY),
        ("📝","Prompts 2-4 : Formation","Plan de formation / Objectif SMART / Module micro-learning — un prompt par livrable",TEAL),
        ("📝","Prompts 5-7 : Qualité","Rapport d'audit / Non-conformité 8D / Plan d'action ISO — appliquez RCTF",GREEN),
        ("📝","Prompts 8-10 : Management","Email de motivation / Compte-rendu réunion / Fiche recrutement",ORANGE),
    ],
    production_title="Affiner et Comparer 3 Itérations",
    production_items=[
        ("Choisir votre meilleur prompt de l'exercice","Demandez la même chose 3 fois avec des RCTF de plus en plus précis"),
        ("Comparer les 3 versions","Notez ce qui change : précision, pertinence, longueur, ton"),
        ("Identifier le pattern gagnant","Quel niveau de détail donne systématiquement le meilleur résultat ?"),
        "Sauvegarder vos 10 prompts dans votre Bibliothèque Notion avec note /5",
    ],
    livrable="Bibliothèque de 10 Prompts RCTF",
    livrable_desc="10 prompts métier organisés par thème dans Notion — votre premier capital de productivité IA.",
    tomorrow_title="Limites et Esprit Critique",
    tomorrow_desc="Vous allez apprendre à détecter les erreurs de l'IA grâce à votre expertise — et à créer votre Charte IA."
))

DAYS.append(dict(
    day_num=3, week=1, phase=1, phase_name="Fondations IA", day_name="Mercredi",
    title="Limites et Esprit Critique", subtitle="Hallucinations · RGPD · Votre Œil d'Expert",
    accent_color=RED,
    theory_title="Ce que l'IA Invente — et Comment le Détecter",
    theory_points=[
        ("Hallucination IA","L'IA peut inventer des faits, des chiffres, des normes — avec le même ton assuré que pour les vraies informations"),
        ("Biais algorithmiques","L'IA reproduit les biais présents dans ses données d'apprentissage — vigilance sur RH et recrutement"),
        ("Données périmées","Les LLM ont une date de coupure — ils ignorent les actualités récentes"),
        ("RGPD : règle absolue","Ne jamais entrer de données personnelles réelles (noms, emails, données clients) dans un LLM public"),
        ("Confidentialité client","Tout ce que vous tapez peut être utilisé pour réentraîner le modèle — anonymisez toujours"),
        "Votre expertise = votre bouclier contre les erreurs de l'IA",
    ],
    theory_insight="Règle d'or : l'IA produit, vous validez. Votre expertise est le filtre qualité irremplaçable.",
    exercise_title="Chasse aux Erreurs — Tester l'IA sur votre Domaine",
    exercise_steps=[
        ("🎯","Test ISO 9001","Posez 5 questions ISO que vous connaissez parfaitement. Notez chaque erreur ou imprécision détectée",RED),
        ("🎯","Test Centres d'appels","Demandez des définitions KPIs, des benchmarks sectoriels. Vérifiez avec vos 15 ans d'expérience",ORANGE),
        ("📋","Documenter les erreurs","Créez une page Notion 'Erreurs IA détectées' — c'est votre preuve d'expertise critique",NAVY),
        ("⚖️","Tester RGPD","Essayez d'entrer un nom fictif + données fictives. Observez comment Claude réagit",TEAL),
    ],
    production_title="Rédiger votre Charte IA Personnelle",
    production_items=[
        ("Charte d'utilisation responsable","Ce que vous faites / ne faites jamais avec l'IA dans votre activité professionnelle"),
        ("Politique de vérification","Pour chaque type de livrable : quand vérifier systématiquement, quand faire confiance"),
        ("Clause RGPD pour vos clients","Comment vous garantissez la protection des données lorsque vous utilisez l'IA en mission"),
        "Ce document est vendable : les clients apprécient un prestataire qui maîtrise la sécurité IA",
    ],
    livrable="Charte d'Utilisation Responsable de l'IA",
    livrable_desc="Document 1-2 pages utilisable avec vos clients pour rassurer sur l'usage éthique et sécurisé de l'IA.",
    tomorrow_title="Perplexity et la Recherche Documentaire",
    tomorrow_desc="Vous allez produire votre premier rapport de veille sectorielle 'IA × Centres d'appels' avec sources citées."
))

DAYS.append(dict(
    day_num=4, week=1, phase=1, phase_name="Fondations IA", day_name="Jeudi",
    title="Perplexity et la Recherche Documentaire", subtitle="Veille sectorielle · Sources citées · Benchmark",
    accent_color=ORANGE,
    theory_title="Perplexity vs Claude — Deux Usages Différents",
    theory_points=[
        ("Claude / ChatGPT","Génèrent du contenu à partir de leurs connaissances — pas de connexion internet en temps réel"),
        ("Perplexity AI","Se connecte à internet, cite ses sources, idéal pour la recherche et la veille"),
        ("Quand utiliser Perplexity","Benchmarks sectoriels / actualités IA / statistiques récentes / tendances marché"),
        ("Quand utiliser Claude","Rédaction / conception formation / analyse / prompts complexes en plusieurs étapes"),
        ("Vérification croisée","Perplexity pour trouver les faits → Claude pour les mettre en forme et analyser"),
        "Combinés, ces deux outils couvrent 90% de vos besoins de contenu professionnel",
    ],
    theory_insight="Perplexity est votre outil de veille automatisé — 1h de recherche manuelle = 5 minutes avec Perplexity.",
    exercise_title="Recherches Sectorielles Guidées",
    exercise_steps=[
        ("🔍","Recherche 1","'IA dans les centres d'appels France 2024 2025' — notez les 5 informations les plus utiles",ORANGE),
        ("🔍","Recherche 2","'Formation professionnelle intelligence artificielle tendances AFNOR' — quelles opportunités ?",TEAL),
        ("🔍","Recherche 3","'Outils IA qualité centre d appels analyse sentiment' — liste des solutions disponibles",GREEN),
        ("🔍","Recherche 4","'Freelance formation IA tarifs France 2024' — benchmark tarifaire de votre futur marché",NAVY),
    ],
    production_title="Rapport de Veille Sectorielle",
    production_items=[
        ("Synthèse Perplexity → Claude","Copiez vos recherches Perplexity. Demandez à Claude d'en faire une synthèse structurée"),
        ("Rapport 5-7 pages","Titre / Contexte / Tendances clés / Opportunités pour formateurs freelance / Conclusion"),
        ("Mise en forme","Exportez ou copiez dans Notion avec mise en page propre"),
        "Ce rapport peut être transformé en article LinkedIn — première étape de votre personal branding",
    ],
    livrable="Rapport de Veille : IA × Centres d'Appels France 2025",
    livrable_desc="Document 5-7 pages avec sources — positionnement d'expert et base de prospection commerciale.",
    tomorrow_title="Google NotebookLM — Votre Expert ISO Personnel",
    tomorrow_desc="Chargez la norme ISO 9001 et interrogez-la comme un expert — créez votre fiche synthèse vulgarisée."
))

DAYS.append(dict(
    day_num=5, week=1, phase=1, phase_name="Fondations IA", day_name="Vendredi",
    title="Google NotebookLM — Expert ISO Personnel", subtitle="Analyser la norme ISO 9001 avec l'IA · Fiche synthèse",
    accent_color=GREEN,
    theory_title="NotebookLM — L'IA qui Lit vos Documents",
    theory_points=[
        ("Principe NotebookLM","Vous chargez vos propres documents — l'IA répond uniquement à partir de ces sources"),
        ("Avantage clé","Zéro hallucination sur le contenu chargé — elle cite exactement ce qui est dans votre document"),
        ("Usages formation","Analyser une norme / synthétiser un manuel / interroger un rapport d'audit / préparer une formation"),
        ("Usages qualité","Extraire les exigences clés / comparer deux versions / créer des FAQ pour managers"),
        ("Format des sources","PDF, Word, Google Docs, sites web, YouTube — grande flexibilité"),
        "NotebookLM transforme n'importe quel document dense en base de connaissances interrogeable",
    ],
    theory_insight="Avec NotebookLM + votre expertise ISO, vous pouvez créer des formations et audits 5x plus vite.",
    exercise_title="Charger ISO 9001 et Interroger",
    exercise_steps=[
        ("📄","Préparer les sources","Trouvez un résumé ou extrait ISO 9001 v2015 en PDF (AFNOR, formation existante, votre cours)",GREEN),
        ("🔗","Charger dans NotebookLM","Nouveau notebook → Ajouter source → chargez votre PDF ISO 9001",TEAL),
        ("❓","15 questions métier","Posez 15 questions précises : chapitres clés, exigences formation, indicateurs, audits internes",ORANGE),
        ("📝","Documenter les réponses","Notez les réponses les plus pertinentes — comparez avec votre connaissance expert",NAVY),
    ],
    production_title="Fiche Synthèse ISO 9001 pour Managers",
    production_items=[
        ("Cible : manager de centre d'appels","Pas un qualitologue — quelqu'un qui doit comprendre sans être expert"),
        ("Structure : 1 page par chapitre clé","Chapitres 4 (Contexte) / 6 (Planification) / 8 (Opérations) / 9 (Évaluation) / 10 (Amélioration)"),
        ("Langage terrain","Traduire chaque exigence en actions concrètes pour un centre d'appels"),
        "Livrable rare et très demandé — peu de formateurs font cette passerelle ISO × terrain",
    ],
    livrable="Guide ISO 9001 'Version Terrain' pour Managers Centres d'Appels",
    livrable_desc="5 fiches d'une page chacune — document unique qui démontre votre double expertise ISO + centres d'appels.",
    tomorrow_title="Bilan Semaine 1 — Portfolio et Évaluation",
    tomorrow_desc="Organisez vos 5 livrables, complétez votre tableau de progression et préparez la Semaine 2."
))

DAYS.append(dict(
    day_num=6, week=1, phase=1, phase_name="Fondations IA", day_name="Samedi",
    title="Bilan Semaine 1 — Portfolio et Auto-Évaluation", subtitle="Consolider · Organiser · Mesurer sa progression",
    accent_color=NAVY,
    theory_title="Bilan et Méthode d'Auto-Évaluation",
    theory_points=[
        "Revisitez chaque livrable de la semaine avec un œil critique et frais",
        "Identifiez votre meilleur livrable et votre prompt le plus efficace",
        ("Quiz Semaine 1","10 questions : LLM / tokens / RCTF / hallucinations / RGPD / NotebookLM"),
        "Mesurez votre progression : niveau de confiance J1 vs J6 — la courbe est votre motivation",
        "Organisez Notion : portfolio structuré, bibliothèque de prompts annotée",
        "Planifiez la Semaine 2 : identifiez les 2 points à renforcer",
    ],
    theory_insight="Un bilan hebdomadaire = 30% de mémorisation supplémentaire. C'est votre rituel de consolidation.",
    exercise_title="Quiz Semaine 1 — 10 Questions",
    exercise_steps=[
        ("❓","Questions 1-3","Qu'est-ce qu'un LLM ? / Expliquez token en 2 phrases / Citez 3 limites de l'IA",NAVY),
        ("❓","Questions 4-6","Écrivez un prompt RCTF complet / Qu'est-ce qu'une hallucination ? / 3 règles RGPD",RED),
        ("❓","Questions 7-9","Différence Claude vs Perplexity / À quoi sert NotebookLM ? / Quelle est votre PVU IA ?",TEAL),
        ("❓","Question 10","Quel est votre livrable préféré de la semaine et pourquoi ? (réponse libre)",GREEN),
    ],
    production_title="Organiser le Portfolio et Améliorer un Livrable",
    production_items=[
        ("Portfolio Notion","5 livrables numérotés, décrits, datés — comme un dossier client professionnel"),
        ("Améliorer le plus faible","Reprenez le livrable que vous trouvez le moins bon. Demandez à Claude comment l'améliorer"),
        ("Tableau de progression","Remplissez la ligne Semaine 2 de votre tableau de bord TAMOU NEURAL PATH"),
        "Post LinkedIn optionnel : partagez une découverte de la semaine — commencez à construire votre audience",
    ],
    livrable="Portfolio Semaine 1 — 5 Livrables Documentés + Score Quiz",
    livrable_desc="Portfolio organisé dans Notion + score quiz /100. Score ≥ 60 = Niveau 1 validé. Sinon : révision ciblée.",
    tomorrow_title="Canva AI — Supports Visuels Professionnels",
    tomorrow_desc="Semaine 2 commence ! Vous allez créer vos premiers supports de formation visuels avec Canva AI en 1h."
))

# ── SEMAINE 2 ──────────────────────────────────────────

DAYS.append(dict(
    day_num=7, week=2, phase=1, phase_name="Fondations IA", day_name="Lundi",
    title="Canva AI — Supports Visuels Professionnels", subtitle="Design IA · Infographies · Supports de formation",
    accent_color=TEAL,
    theory_title="Canva AI — Design Sans Être Graphiste",
    theory_points=[
        ("Magic Design","Décrivez ce que vous voulez en texte → Canva génère la mise en page automatiquement"),
        ("Magic Write","Générez du texte directement dans vos slides — titres, bullets, descriptions"),
        ("Redimensionnement IA","Un design créé une fois → adapté automatiquement à tous les formats"),
        ("Bibliothèque IA","Images générées par IA directement intégrées dans vos supports"),
        ("Usage éthique","Images IA : vérifiez les droits. Canva garantit l'usage commercial de ses générés"),
        "Un support professionnel augmente la valeur perçue de vos formations de 40% selon les apprenants",
    ],
    theory_insight="Canva AI vous donne l'apparence d'une agence de design pour le prix d'un abonnement basique.",
    exercise_title="Explorer les Fonctionnalités Canva AI",
    exercise_steps=[
        ("🎨","Magic Design","Tapez : 'Support formation centres d'appels soft skills professionnel' → générez 5 variantes",TEAL),
        ("✍️","Magic Write","Dans une slide vide : demandez à Magic Write de rédiger 5 bullets 'Techniques d'écoute active'",GREEN),
        ("🖼️","Génération d'images","Générez 3 images : agent téléphonique professionnel / équipe formation / tableau de bord qualité",ORANGE),
        ("📐","Formats","Créez votre design en format Présentation → redimensionnez en Post LinkedIn → Infographie",NAVY),
    ],
    production_title="Support de Formation Complet — 10 Slides",
    production_items=[
        ("Thème : 'Soft Skills au Téléphone — 7 Techniques'","Slides : titre / objectifs / 7 techniques (1 slide chacune) / synthèse / évaluation"),
        ("Design unifié","Choisissez 1 palette de couleurs, 1 police — cohérence professionnelle"),
        ("Visuels IA","Au moins 3 images générées par IA intégrées dans les slides"),
        "Exportez en PDF + PPTX — deux formats pour deux types de clients",
    ],
    livrable="Support de Formation Visuel — Soft Skills Téléphoniques",
    livrable_desc="10 slides design professionnel, prêtes à utiliser en présentiel ou distanciel — premier support vendable.",
    tomorrow_title="Gamma.app — La Présentation en 60 Secondes",
    tomorrow_desc="Vous allez générer une présentation complète de 25 slides en moins d'une heure avec Gamma.app."
))

DAYS.append(dict(
    day_num=8, week=2, phase=1, phase_name="Fondations IA", day_name="Mardi",
    title="Gamma.app — La Présentation en 60 Secondes", subtitle="Pitch décks · Présentations clients · Propositions",
    accent_color=RED,
    theory_title="Gamma.app — IA de Présentation",
    theory_points=[
        ("Principe","Décrivez votre présentation en texte → Gamma génère toute la structure, le design et le contenu"),
        ("Avantage vs PowerPoint","10x plus rapide / Design automatique / Export PPTX ou PDF / Partage en ligne"),
        ("Limite à connaître","Le contenu généré est un premier jet — vous devez valider et affiner avec votre expertise"),
        ("Usages prioritaires pour vous","Pitch client / Présentation offre freelance / Support atelier / Rapport qualité visuel"),
        ("Plan optimal","Donnez à Gamma : public cible + objectif + 5-7 points clés + ton souhaité"),
        "Une présentation Gamma de qualité professionnelle se produit en 45 minutes contre 3 heures en manuel",
    ],
    theory_insight="Gamma est votre outil de prototypage rapide — générez, validez avec votre expertise, affinez.",
    exercise_title="Générer 3 Présentations Différentes",
    exercise_steps=[
        ("🚀","Présentation 1","'Soft Skills et IA : les compétences de demain pour les agents centres d'appels' — 15 slides",RED),
        ("🚀","Présentation 2","'ISO 9001 v2015 expliqué simplement aux managers' — 12 slides vulgarisées",NAVY),
        ("🚀","Présentation 3","'Pourquoi intégrer l'IA dans votre programme de formation' — 10 slides convaincantes",TEAL),
        ("🔍","Analyse critique","Pour chaque présentation : qu'est-ce qui est juste ? qu'est-ce que vous corrigeriez ?",GREEN),
    ],
    production_title="Pitch Deck Freelance IA — Version 1",
    production_items=[
        ("Contenu du pitch","Qui je suis / Mon expertise / Mes offres IA / Mes livrables exemples / Tarifs / Contact"),
        ("Prompt Gamma précis","Incluez : '15 ans expérience centres d'appels / formatrice / ISO 9001 / freelance IA / ton expert et accessible'"),
        ("Affiner avec votre expertise","Corrigez les inexactitudes, ajoutez vos exemples concrets, personnalisez le ton"),
        "Ce pitch deck sera votre support de prospection — investissez du temps pour le personnaliser",
    ],
    livrable="Pitch Deck Freelance IA — 20-25 Slides Professionnelles",
    livrable_desc="Votre présentation de positionnement commercial — à envoyer aux prospects et à publier sur LinkedIn.",
    tomorrow_title="Prompts en Chaîne — Chain of Thought",
    tomorrow_desc="Apprenez à décomposer des missions complexes en séquences de prompts enchaînés — créez un module complet."
))

DAYS.append(dict(
    day_num=9, week=2, phase=1, phase_name="Fondations IA", day_name="Mercredi",
    title="Prompts en Chaîne — Chain of Thought", subtitle="Décomposer les missions complexes · Module formation complet",
    accent_color=ORANGE,
    theory_title="Chain of Thought — Penser en Séquences",
    theory_points=[
        ("Problème","Un seul prompt long donne souvent une réponse superficielle et peu structurée"),
        ("Solution : Chain of Thought","Décomposez la mission en 4-6 étapes logiques, chacune avec son propre prompt"),
        ("Avantage 1","Chaque étape peut être vérifiée et affinée avant de passer à la suivante"),
        ("Avantage 2","L'IA garde le contexte d'une réponse à l'autre dans la même conversation"),
        ("Structure type","Objectifs → Plan → Contenu → Exercices → Évaluation → Guide formateur"),
        "Un module de formation complet se construit en 5 prompts enchaînés en moins d'1h30",
    ],
    theory_insight="Chain of Thought = penser comme un ingénieur pédagogique, faire faire à l'IA chaque pièce du puzzle.",
    exercise_title="Chaîne de 5 Prompts — Module Gestion Client Difficile",
    exercise_steps=[
        ("1️⃣","Prompt 1 — Objectifs","'Définis 4 objectifs SMART pour un module Gestion client difficile, agents 6 mois expérience'",ORANGE),
        ("2️⃣","Prompts 2-3","Plan de cours 5 parties → Développer le contenu de la partie 'Désescalade émotionnelle'",TEAL),
        ("3️⃣","Prompt 4","'Crée 3 exercices pratiques de mise en situation pour cette partie, avec grilles d'observation'",GREEN),
        ("4️⃣","Prompt 5","'Rédige l'évaluation finale du module avec corrigé — QCM 10 questions + cas pratique'",NAVY),
    ],
    production_title="Module Complet — Gestion Client Difficile (4h)",
    production_items=[
        ("Assembler les 5 livrables","Objectifs + plan + contenu + exercices + évaluation = module complet clé en main"),
        ("Guide formateur","Demandez à Claude de générer le guide formateur avec timings et conseils d'animation"),
        ("Vérification experte","Relisez avec votre regard de formatrice — corrigez ce qui sonne faux"),
        "Ce module est immédiatement vendable — estimez sa valeur (temps économisé × tarif horaire)",
    ],
    livrable="Module de Formation Complet — Gestion Client Difficile (4h)",
    livrable_desc="Objectifs + plan + contenu + 3 exercices + évaluation + guide formateur — module clé en main.",
    tomorrow_title="Personas et Simulation d'Apprenants",
    tomorrow_desc="Créez 4 personas d'apprenants typiques de vos formations et des parcours différenciés pour chacun."
))

DAYS.append(dict(
    day_num=10, week=2, phase=1, phase_name="Fondations IA", day_name="Jeudi",
    title="Personas et Simulation d'Apprenants", subtitle="Différenciation pédagogique · Parcours adaptatifs",
    accent_color=GOLD,
    theory_title="Les Personas — Connaître ses Apprenants",
    theory_points=[
        ("Qu'est-ce qu'un persona","Portrait-robot d'un type d'apprenant : profil, motivations, freins, mode d'apprentissage préféré"),
        ("Pourquoi c'est crucial","Une formation identique pour tous = 40% d'efficacité. Différenciée = 80%"),
        ("IA et personas","Claude peut simuler un apprenant donné pour tester vos exercices et détecter les difficultés"),
        ("4 personas types CA","Agent débutant anxieux / Agent expérimenté résistant / Manager sceptique / Manager motivé"),
        ("Application pratique","Pour chaque livrable : demandez à Claude 'Comment cet apprenant réagirait à ce module ?'"),
        "Les personas transforment une formation générique en expérience personnalisée — c'est votre valeur ingénierie",
    ],
    theory_insight="L'IA peut jouer le rôle de n'importe quel apprenant — testez vos formations avant de les délivrer.",
    exercise_title="Créer et Tester 4 Personas",
    exercise_steps=[
        ("👤","Persona 1 — L'Anxieux","Agent 3 mois / peur de l'erreur / besoin de rassurance. Demandez à Claude de jouer ce rôle",GOLD),
        ("👤","Persona 2 — Le Résistant","Agent 8 ans / 'on fait déjà ça' / valorise son expérience. Comment l'engager ?",RED),
        ("👤","Persona 3 — Le Manager Sceptique","'L'IA ça ne marchera jamais chez nous' — argumentaire adapté",ORANGE),
        ("👤","Persona 4 — Le Curieux","Apprenant enthousiaste mais dispersé — comment canaliser l'énergie ?",TEAL),
    ],
    production_title="4 Parcours Différenciés — Module Gestion Client Difficile",
    production_items=[
        ("Adapter le module Jour 9","Pour chaque persona : mêmes objectifs, approches pédagogiques différentes"),
        ("Personnaliser le ton","L'Anxieux a besoin de bienveillance / Le Résistant a besoin de légitimité / etc."),
        ("Créer les variantes d'exercices","Même mise en situation, niveaux de difficulté différents"),
        "Ces variantes font partie de votre ingénierie pédagogique avancée — différenciateur commercial fort",
    ],
    livrable="4 Fiches Personas + Parcours Différenciés",
    livrable_desc="4 profils d'apprenants détaillés + adaptation du module J9 pour chaque profil — ingénierie pédagogique IA.",
    tomorrow_title="Claude Projects — Votre Assistant Personnel",
    tomorrow_desc="Configurez 3 assistants IA spécialisés qui connaissent votre méthode, vos offres et vos chartes métier."
))

DAYS.append(dict(
    day_num=11, week=2, phase=1, phase_name="Fondations IA", day_name="Vendredi",
    title="Claude Projects — Vos Assistants Personnels", subtitle="Mémoire permanente · Base de connaissances · 3 assistants",
    accent_color=NAVY,
    theory_title="Claude Projects — La Mémoire de l'IA",
    theory_points=[
        ("Problème sans Projects","Chaque conversation repart de zéro — vous réexpliquez votre contexte à chaque fois"),
        ("Solution : Claude Projects","Vous configurez une fois : profil, méthode, chartes, offres → disponibles en permanence"),
        ("Instructions système","Texte que Claude lit avant chaque conversation dans ce project — sa 'fiche de poste'"),
        ("Documents de base","Chargez vos propres documents : chartes d'appel, grilles QA, procédures ISO"),
        ("3 assistants à créer","Formation Expert / Quality Coach / Commercial Coach — chacun avec sa spécialité"),
        "Claude Projects = votre équipe IA permanente qui connaît votre métier par cœur",
    ],
    theory_insight="Avec 3 Projects bien configurés, chaque prompt devient 3x plus efficace car le contexte est déjà là.",
    exercise_title="Configurer vos 3 Claude Projects",
    exercise_steps=[
        ("🤖","Project 1 — Formation Expert","Instructions : méthodes pédagogiques, niveaux apprenants, chartes formation, ton professionnel bienveillant",NAVY),
        ("🎯","Project 2 — Quality Coach","Instructions : grilles QA, normes ISO 9001, KPIs CA, procédures audit, rigueur et précision",RED),
        ("💼","Project 3 — Commercial Coach","Instructions : vos offres freelance, PVU, tarifs, objections fréquentes, ton confiant et accessible",GREEN),
        ("🧪","Tester les 3","Posez la même question dans les 3 projects — observez comment chaque assistant répond différemment",ORANGE),
    ],
    production_title="Charger vos Documents dans les Projects",
    production_items=[
        ("Project Formation","Chargez : module Jour 9 / fiches personas Jour 10 / grille QA Jour 2"),
        ("Project Quality","Chargez : charte utilisation IA Jour 3 / fiche ISO 9001 Jour 5 / grille QA"),
        ("Project Commercial","Chargez : pitch deck Jour 8 / rapport veille Jour 4 / bibliothèque prompts"),
        "Testez chaque project avec 5 cas concrets de votre métier — affinez les instructions",
    ],
    livrable="3 Claude Projects Opérationnels",
    livrable_desc="Formation Expert + Quality Coach + Commercial Coach — vos 3 assistants IA permanents configurés et testés.",
    tomorrow_title="Évaluation Phase 1 — Niveau 1 Validé ?",
    tomorrow_desc="Quiz complet Phase 1 + organisation de votre portfolio de 10 livrables. Score ≥ 60/100 = Niveau 1 certifié."
))

DAYS.append(dict(
    day_num=12, week=2, phase=1, phase_name="Fondations IA", day_name="Samedi",
    title="Évaluation Phase 1 — Niveau 1 Certifié", subtitle="Quiz complet · Portfolio · Passage Phase 2",
    accent_color=GREEN,
    theory_title="Grille d'Évaluation — Niveau 1 Découverte",
    theory_points=[
        ("Critère 1 /20","Écrire un prompt RCTF efficace sans aide — testé sur 2 cas inédits"),
        ("Critère 2 /20","Utiliser Claude, ChatGPT, Perplexity, Notion, Canva de façon autonome"),
        ("Critère 3 /20","Expliquer les risques RGPD à un client — simulé avec Claude"),
        ("Critère 4 /20","Portfolio de 10 livrables professionnels documentés et organisés"),
        ("Critère 5 /20","Identifier les erreurs dans une réponse IA sur votre domaine"),
        "Score ≥ 60/100 = Niveau 1 validé → Phase 2 autorisée. Score < 60 = révision ciblée 1 jour.",
    ],
    theory_insight="L'évaluation n'est pas un examen — c'est une mesure honnête pour cibler les révisions utiles.",
    exercise_title="Quiz Phase 1 — 20 Questions Chrono",
    exercise_steps=[
        ("📝","Questions 1-5 : Théorie","LLM / tokens / paramètres / hallucinations / RGPD — 5 min",GREEN),
        ("📝","Questions 6-10 : Outils","Claude / ChatGPT / Perplexity / Notion / Canva — 5 min",TEAL),
        ("📝","Questions 11-15 : Pratique","Écrire 2 prompts RCTF / corriger une erreur IA — 10 min",ORANGE),
        ("📝","Questions 16-20 : Portfolio","Présenter 3 livrables, expliquer leur valeur ajoutée — 10 min",NAVY),
    ],
    production_title="Portfolio Phase 1 Finalisé",
    production_items=[
        ("Organiser les 10 livrables","Chaque livrable : titre / date / outils utilisés / valeur commerciale estimée"),
        ("Calculer votre score","Remplissez la grille d'évaluation honnêtement — corrigez avec Claude"),
        ("Plan de révision si besoin","Score < 60 : identifiez les 2 points faibles, planifiez 1 journée de révision"),
        "Célébrez ! En 2 semaines vous avez produit 10 livrables professionnels avec l'IA — c'est remarquable.",
    ],
    livrable="Portfolio Phase 1 — 10 Livrables + Score /100",
    livrable_desc="Phase 1 complétée. Score ≥ 60/100 = prête pour la Phase 2 : IA × Centres d'Appels.",
    tomorrow_title="Phase 2 — IA × Centres d'Appels",
    tomorrow_desc="Analyse de sentiment, NLP, outils QA IA — vous devenez la référence IA qualité centres d'appels."
))

# ════════════════════════════════════════════════
# PHASE 2 — IA × CENTRES D'APPELS (Semaines 3-5)
# ════════════════════════════════════════════════

DAYS.append(dict(
    day_num=13, week=3, phase=2, phase_name="IA × Centres d'Appels", day_name="Lundi",
    title="NLP et Analyse de Sentiment", subtitle="Comment l'IA comprend la voix · Applications QA",
    accent_color=GREEN,
    theory_title="NLP — Le Langage Naturel Compris par l'IA",
    theory_points=[
        ("NLP = Natural Language Processing","Branche de l'IA qui analyse et comprend le langage humain (texte et voix)"),
        ("Analyse de sentiment","L'IA détecte si un texte est positif / négatif / neutre / mixte — et mesure l'intensité"),
        ("Applications CA","Scoring automatique des appels / détection insatisfaction / alertes en temps réel"),
        ("Transcription automatique","La voix → texte → analyse IA — en quelques secondes"),
        ("Ce que vous pouvez faire aujourd'hui","Simuler des transcriptions et les faire analyser par Claude"),
        "Secteur CA : 68% des entreprises prévoient d'utiliser l'analyse sentiment IA d'ici 2026",
    ],
    theory_insight="Vous connaissez déjà l'analyse qualité manuelle — l'IA l'automatise. Votre expertise définit les critères.",
    exercise_title="Simuler et Analyser 5 Transcriptions",
    exercise_steps=[
        ("📞","Créer 5 transcriptions","Demandez à Claude de générer 5 transcriptions appels entrants : 1 excellent / 2 moyens / 2 mauvais",GREEN),
        ("🔍","Analyser chaque appel","Pour chaque transcription : demandez sentiment client / sentiment agent / points de friction",TEAL),
        ("📊","Scorer avec votre grille","Appliquez votre grille QA Jour 2 à ces transcriptions — comparez le score IA vs votre note",ORANGE),
        ("⚖️","Calibrage","Là où IA et expert divergent : analysez pourquoi — l'IA peut manquer le contexte",NAVY),
    ],
    production_title="Protocole d'Analyse Qualité IA",
    production_items=[
        ("Procédure standard","Étapes : réception transcription → analyse Claude → scoring → validation experte → feedback agent"),
        ("Prompts d'analyse","Prompt type pour analyser une transcription en 5 dimensions qualité"),
        ("Calibrage humain-IA","Dans quels cas l'expert prime toujours sur l'IA — liste des situations"),
        "Ce protocole est vendable comme service de conseil à des responsables qualité CA",
    ],
    livrable="Protocole d'Analyse Qualité IA — Appels Entrants",
    livrable_desc="Procédure standardisée + prompts d'analyse + grille de calibrage — livrable conseil à forte valeur.",
    tomorrow_title="Outils IA de Qualité Appels — Panorama",
    tomorrow_desc="Modjo, Gong, Refract, Scorebuddy — vous devenez experte en recommandation d'outils QA IA."
))

DAYS.append(dict(
    day_num=14, week=3, phase=2, phase_name="IA × Centres d'Appels", day_name="Mardi",
    title="Outils IA de Qualité Appels — Panorama", subtitle="Modjo · Gong · Refract · Guide de recommandation",
    accent_color=TEAL,
    theory_title="Marché des Outils QA IA pour Centres d'Appels",
    theory_points=[
        ("Modjo","Outil FR / analyse appels / coaching IA / intégration CRM — cible : équipes commerciales et CA"),
        ("Gong","Leader mondial / Revenue Intelligence / idéal grands comptes internationaux"),
        ("Refract / Jiminny","Coaching IA vidéo + audio / feedback automatisé / learning paths"),
        ("Scorebuddy","Spécialiste QA CA / grilles configurables / reporting avancé / conformité"),
        ("EvaluAgent","QA automation / scoring IA / intégration omnicanale"),
        "Critères de recommandation : taille équipe / budget / intégrations existantes / langue / conformité RGPD",
    ],
    theory_insight="Votre valeur = vous ne vendez pas un outil, vous recommandez le bon outil pour le bon contexte.",
    exercise_title="Benchmark Comparatif — 5 Outils",
    exercise_steps=[
        ("🔍","Recherche Perplexity","Pour chaque outil : fonctionnalités / prix / avis clients / cas d'usage France",TEAL),
        ("📊","Grille comparative","Critères : prix / fonctionnalités QA / IA niveau / RGPD / intégrations / langue FR",GREEN),
        ("🎯","Scénarios clients","Quel outil pour : 30 agents PME / 200 agents ETI / télétravail / conformité bancaire ?",ORANGE),
        ("💡","Rédiger les recommandations","Pour chaque scénario : outil recommandé + justification + alternatives",NAVY),
    ],
    production_title="Guide de Recommandation Outils QA IA",
    production_items=[
        ("Structure du guide","Introduction / Tableau comparatif / 4 scénarios clients / Méthodologie de choix / Conclusion"),
        ("Ton consultant","Neutre, factuel, basé sur des critères objectifs — pas de préférence a priori"),
        ("Mise en page","Canva : tableau coloré + icônes + score visuel par critère"),
        "Ce guide positionne comme experte en conseil IA — très rare sur le marché FR centres d'appels",
    ],
    livrable="Guide Comparatif Outils QA IA — Recommandation Client",
    livrable_desc="Document conseil 8-10 pages — différenciateur fort dans votre offre d'audit qualité.",
    tomorrow_title="Scripts Intelligents et Arbres de Décision",
    tomorrow_desc="Créez un script d'appel entrant réclamation avec 3 niveaux d'escalade et gestion émotionnelle IA."
))

DAYS.append(dict(
    day_num=15, week=3, phase=2, phase_name="IA × Centres d'Appels", day_name="Mercredi",
    title="Scripts Intelligents et Arbres de Décision", subtitle="Scripts adaptatifs · Escalade · Désescalade émotionnelle",
    accent_color=ORANGE,
    theory_title="Scripts IA vs Scripts Classiques",
    theory_points=[
        ("Script classique","Linéaire / rigide / ne s'adapte pas à l'émotionnel du client / frustrant pour l'agent"),
        ("Script IA adaptatif","Branché sur le sentiment client / propose des variantes selon le contexte / s'enrichit"),
        ("Arbre de décision","Si client calme → procédure standard / Si client agité → désescalade / Si urgent → escalade"),
        ("3 niveaux d'escalade","Niveau 1 : agent / Niveau 2 : superviseur / Niveau 3 : responsable — avec critères clairs"),
        ("Formules de désescalade","Empathie verbale / validation du ressenti / reformulation positive / engagement de résolution"),
        "Un bon script IA réduit le temps de traitement de 20% et augmente la satisfaction de 35%",
    ],
    theory_insight="Votre expertise terrain vous permet de rédiger des scripts que l'IA seule ne peut pas produire — combinaison gagnante.",
    exercise_title="Construire le Script Réclamation",
    exercise_steps=[
        ("🗺️","Cartographier les flux","Listez les 5 types de réclamations les plus fréquentes en CA entrant — avec Claude",ORANGE),
        ("📝","Script niveau 1","Rédiger l'ouverture / identification / reformulation / proposition solution / clôture",RED),
        ("📝","Scripts escalade 2 & 3","Formules de transition / informations à transmettre / posture agent",TEAL),
        ("🎭","Tester avec jeu de rôle","Claude joue un client réclamant — testez votre script en temps réel",GREEN),
    ],
    production_title="Script Complet Appels Entrants Réclamation",
    production_items=[
        ("Structure complète","Accueil → Identification → Écoute → Reformulation → Solution → Escalade si besoin → Clôture"),
        ("Variantes émotionnelles","3 registres : client calme / irrité / en colère — formules adaptées à chaque",),
        ("Guide d'utilisation","Comment lire et adapter le script sans le réciter — garder le naturel",),
        "Format : document Word structuré + version quick reference une page pour l'agent",
    ],
    livrable="Script Appels Entrants Réclamation — 3 Niveaux",
    livrable_desc="Script complet avec variantes émotionnelles et protocole escalade — modèle adaptable par secteur.",
    tomorrow_title="KPIs et Tableaux de Bord IA",
    tomorrow_desc="Créez un dashboard KPI centres d'appels augmenté par l'IA avec zones d'alerte et recommandations auto."
))

DAYS.append(dict(
    day_num=16, week=3, phase=2, phase_name="IA × Centres d'Appels", day_name="Jeudi",
    title="KPIs et Tableaux de Bord IA", subtitle="Performance agents · Alertes prédictives · Dashboard",
    accent_color=NAVY,
    theory_title="KPIs Centres d'Appels × IA",
    theory_points=[
        ("KPIs fondamentaux","DMT / Taux de décroché / FCR (First Call Resolution) / Taux d'abandon / NPS / Taux conformité"),
        ("KPIs qualité","Score qualité moyen / Taux de non-conformité / Conformité charte / Escalades évitées"),
        ("KPIs formation","Taux de progression post-formation / Ancrage mémoriel J+30 / Application terrain"),
        ("IA prédictive","L'IA détecte les tendances avant qu'elles deviennent des problèmes — alertes précoces"),
        ("Dashboard intelligent","Seuils d'alerte automatiques / visualisation tendances / recommandations contextuelles"),
        "Un dashboard IA bien configuré = 30% de temps de reporting économisé pour les superviseurs",
    ],
    theory_insight="Vous connaissez les KPIs — l'IA les surveille en continu et vous alerte. Vous interprétez et décidez.",
    exercise_title="Construire les Indicateurs et Seuils",
    exercise_steps=[
        ("📊","Sélectionner 10 KPIs","Avec Claude : choisissez les 10 KPIs les plus pertinents pour votre type de CA",NAVY),
        ("⚠️","Définir les seuils","Pour chaque KPI : seuil vert / orange / rouge + cause probable par niveau",RED),
        ("💡","Recommandations auto","Pour chaque passage en rouge : quelle action recommander automatiquement ?",TEAL),
        ("🎨","Maquetter le dashboard","Canva : tableau de bord visuel avec codes couleur, jauges, tendances",GREEN),
    ],
    production_title="Dashboard KPI IA-Augmenté",
    production_items=[
        ("Template Canva","Dashboard 1 page : 10 KPIs / codes couleur / zones alerte / espace commentaires superviseur"),
        ("Fiche d'interprétation","Guide d'utilisation : comment lire le dashboard / quand escalader / quoi faire"),
        ("Prompt de mise à jour","Prompt Claude pour analyser nouvelles données et mettre à jour les recommandations"),
        "Ce dashboard est un livrable clé de votre offre d'accompagnement — montrez-le à chaque prospect",
    ],
    livrable="Dashboard KPI Centres d'Appels IA-Augmenté",
    livrable_desc="Template visuel Canva + fiche interprétation + prompt de mise à jour — outil clé de votre offre conseil.",
    tomorrow_title="Mailing Back-Office et QA Multicanal",
    tomorrow_desc="Créez votre kit complet d'évaluation qualité multicanal : appels + emails + chat."
))

DAYS.append(dict(
    day_num=17, week=3, phase=2, phase_name="IA × Centres d'Appels", day_name="Vendredi",
    title="Mailing Back-Office et QA Multicanal", subtitle="Évaluation emails · Chat · Kit QA complet",
    accent_color=RED,
    theory_title="Qualité Multicanale — Au-delà du Téléphone",
    theory_points=[
        ("Mailing back-office","Emails clients : délai de réponse / personnalisation / conformité / qualité rédactionnelle"),
        ("Chat en ligne","Réactivité / fluidité / cohérence avec le verbal / escalade vers téléphone"),
        ("Spécificités email vs appel","Email : traçabilité / relecture possible / moins d'urgence / RGPD données écrites"),
        ("IA et back-office","Catégorisation automatique / suggestion de réponse / contrôle qualité avant envoi"),
        ("Cohérence omnicanale","Le client attend la même qualité quel que soit le canal — votre rôle de garant"),
        "Les entreprises avec QA multicanale ont un NPS 23 points supérieur à celles téléphone seul",
    ],
    theory_insight="Rares sont les formateurs qui maîtrisent la QA multicanale — c'est un vrai différenciateur commercial.",
    exercise_title="Créer les Grilles Mailing et Chat",
    exercise_steps=[
        ("📧","Grille email back-office","Critères : délai / objet / personnalisation / clarté / solution proposée / signature / conformité",RED),
        ("💬","Grille chat","Critères : réactivité / ton / cohérence / proactivité / escalade appropriée",ORANGE),
        ("🔗","Cohérence inter-canaux","Prompt pour vérifier qu'un client ayant appelé puis écrit reçoit des réponses cohérentes",TEAL),
        ("📋","Assembler le kit","Regrouper : grille appel J2 + grille email + grille chat + procédure multicanale",GREEN),
    ],
    production_title="Kit QA Multicanal Complet",
    production_items=[
        ("3 grilles harmonisées","Appels entrants / emails back-office / chat — mêmes dimensions qualité, formats adaptés"),
        ("Guide d'utilisation","Comment calibrer les évaluateurs / fréquence d'évaluation par canal / seuils d'alerte"),
        ("Formation évaluateurs","Micro-module 1h pour former les superviseurs à utiliser les 3 grilles"),
        "Ce kit est l'une de vos offres commerciales les plus demandées — packagez-le avec votre tarif",
    ],
    livrable="Kit QA Multicanal — Appels + Emails + Chat",
    livrable_desc="3 grilles harmonisées + guide utilisation + micro-module formation évaluateurs — offre packagée.",
    tomorrow_title="Bilan Semaine 3 — Offre Audit Qualité",
    tomorrow_desc="Assemblez vos livrables en une première offre commerciale packagée et chiffrez-la."
))

DAYS.append(dict(
    day_num=18, week=3, phase=2, phase_name="IA × Centres d'Appels", day_name="Samedi",
    title="Bilan Semaine 3 — Offre Audit Qualité", subtitle="Packager · Chiffrer · Présenter",
    accent_color=GREEN,
    theory_title="Construire une Offre Commerciale Packagée",
    theory_points=[
        ("Principe du package","Regrouper plusieurs livrables en une offre cohérente avec un prix global"),
        ("Vos livrables de la semaine","Protocole analyse IA / Guide outils QA / Script réclamation / Dashboard KPI / Kit QA multicanal"),
        ("Valeur perçue vs valeur réelle","Le client achète un résultat (centre d'appels amélioré), pas des documents"),
        ("Tarification par valeur","Prix basé sur l'économie générée pour le client, pas sur votre temps passé"),
        ("Structure d'une offre","Problème client → Votre solution → Livrables → Preuves → Prix → Appel à l'action"),
        "Une offre bien packagée se vend 3x plus facilement qu'une liste de prestations à l'heure",
    ],
    theory_insight="Vous avez produit cette semaine les composants de votre Offre 1 — il ne reste qu'à les assembler.",
    exercise_title="Assembler l'Offre Audit Qualité",
    exercise_steps=[
        ("💼","Définir les livrables inclus","Sélectionnez les 4-5 livrables les plus impactants de la semaine",GREEN),
        ("💰","Estimer la valeur","Temps économisé pour le client / coût d'une erreur qualité / valeur formation interne",TEAL),
        ("📝","Rédiger la description","Claude rédige la description de l'offre — vous validez et personnalisez",ORANGE),
        ("🎨","Mise en page","Canva : fiche offre 1 page — titre accrocheur / livrables / bénéfices / prix",NAVY),
    ],
    production_title="Fiche Offre 1 — Audit Qualité IA",
    production_items=[
        ("Contenu","Problème / Solution / Livrables inclus / Durée mission / Prix / Conditions"),
        ("Présentation Gamma","15 slides : contexte / approche / livrables / preuves / témoignages / CTA"),
        ("Email de présentation","Template email pour envoyer cette offre à un prospect responsable qualité CA"),
        "Votre première offre commerciale packagée est prête — montrez-la à 3 personnes pour feedback",
    ],
    livrable="Offre 1 Packagée — Audit Qualité IA Centres d'Appels",
    livrable_desc="Fiche offre 1 page + présentation 15 slides + email prospect — prête à envoyer.",
    tomorrow_title="Adaptive Learning — Parcours Personnalisés",
    tomorrow_desc="Semaine 4 : concevoir des parcours de formation adaptatifs pour 3 profils d'agents différents."
))

DAYS.append(dict(
    day_num=19, week=4, phase=2, phase_name="IA × Centres d'Appels", day_name="Lundi",
    title="Adaptive Learning — Parcours Personnalisés", subtitle="Formation différenciée · 3 niveaux · IA pédagogique",
    accent_color=TEAL,
    theory_title="L'Adaptive Learning — La Formation qui s'Adapte",
    theory_points=[
        ("Définition","Parcours qui s'adapte au niveau, au rythme et aux résultats de chaque apprenant"),
        ("IA et adaptation","L'IA analyse les réponses → ajuste la difficulté → personnalise les exercices"),
        ("3 profils types CA","Débutant (0-6 mois) / Confirmé (6 mois-2 ans) / Expert (2 ans+)"),
        ("Mêmes objectifs, voies différentes","Tous doivent maîtriser les mêmes compétences — les chemins varient"),
        ("Sans LMS dédié","Vous pouvez créer des parcours adaptatifs simples avec Claude + questionnaires"),
        "Un parcours adaptatif réduit le temps de formation de 40% tout en augmentant la rétention de 60%",
    ],
    theory_insight="L'adaptive learning est l'avenir de la formation — vous pouvez le proposer aujourd'hui avec des outils accessibles.",
    exercise_title="Architecturer 3 Parcours selon le Niveau",
    exercise_steps=[
        ("🎓","Parcours Débutant","Fondamentaux / beaucoup de pratique guidée / feedback fréquent / progression linéaire",TEAL),
        ("🎓","Parcours Confirmé","Consolidation ciblée / cas complexes / autonomie croissante / évaluation par situations",GREEN),
        ("🎓","Parcours Expert","Perfectionnement / pairs learning / cas rares / contribution au collectif",ORANGE),
        ("🔗","Points de jonction","Moments où les parcours se rejoignent — cohésion d'équipe et partage",NAVY),
    ],
    production_title="Architecture Formation Adaptative Agents CA",
    production_items=[
        ("Parcours complet 3 niveaux","Objectifs / contenus / activités / évaluations / durées pour chaque profil"),
        ("Outils simples","Typeform pour diagnostic initial / Claude pour personnaliser / Notion pour suivre"),
        ("Guide de l'animateur","Comment identifier le bon parcours pour chaque agent / comment progresser"),
        "Cette architecture est la base de votre offre d'ingénierie de formation — valorisez-la à 2000-5000€",
    ],
    livrable="Architecture Formation Adaptative — 3 Niveaux Agents CA",
    livrable_desc="3 parcours complets avec objectifs, activités et évaluations différenciés — ingénierie pédagogique premium.",
    tomorrow_title="Jeux de Rôle IA et Simulation",
    tomorrow_desc="Apprenez à configurer Claude comme client simulé — créez 10 scénarios de jeux de rôle prêts à l'emploi."
))

DAYS.append(dict(
    day_num=20, week=4, phase=2, phase_name="IA × Centres d'Appels", day_name="Mardi",
    title="Jeux de Rôle IA et Simulation", subtitle="Client simulé · 10 scénarios · Livret formation",
    accent_color=RED,
    theory_title="La Simulation IA — Former sans Risque",
    theory_points=[
        ("Principe","Claude joue le rôle d'un client — vous ou vos apprenants jouent l'agent"),
        ("Avantage pédagogique","Pratique illimitée / sans enjeu réel / feedback immédiat / personnalisation maximale"),
        ("Configuration du persona client","Plus le persona est détaillé, plus la simulation est réaliste et formatrice"),
        ("5 typologies de clients","Client en colère / impatient / confus / manipulateur / de bonne foi mais exigeant"),
        ("Débriefing IA","Après chaque simulation : Claude analyse la performance et donne du feedback structuré"),
        "Les simulations IA augmentent la confiance des agents de 47% avant leur première vraie prise d'appel",
    ],
    theory_insight="Un livret de 10 scénarios bien construits vous permet d'animer des formations pratiques sans préparer à chaque fois.",
    exercise_title="Créer et Tester 5 Scénarios",
    exercise_steps=[
        ("🎭","Scénario 1 — Client en colère","Facture incorrecte / 3e appel / menace de partir — testez la désescalade",RED),
        ("🎭","Scénario 2 — Client confus","Ne comprend pas sa facture / répète la même question / besoin de pédagogie",ORANGE),
        ("🎭","Scénario 3 — Client manipulateur","Invente des faits / cherche à obtenir une remise injustifiée",TEAL),
        ("🎭","Scénarios 4-5","Client impatient (file d'attente) / Client fidèle mécontent — construisez le prompt vous-même",GREEN),
    ],
    production_title="Livret de Jeux de Rôle IA — 10 Scénarios",
    production_items=[
        ("5 scénarios restants","Complétez les scénarios 6-10 : résiliation / urgence / recouvrement / vente additionnelle / complexe"),
        ("Pour chaque scénario","Contexte client / prompt de configuration / objectifs pédagogiques / grille d'observation / débriefing"),
        ("Guide formateur","Comment animer une session jeux de rôle IA / durée par scénario / variantes possibles"),
        "Ce livret est un produit formation clé en main — vendable seul ou inclus dans votre offre formation",
    ],
    livrable="Livret Jeux de Rôle IA — 10 Scénarios Centres d'Appels",
    livrable_desc="10 scénarios configurés avec prompts, grilles d'observation et guide formateur — prêt à animer.",
    tomorrow_title="Feedback Automatisé et Coaching IA",
    tomorrow_desc="Configurez un outil Coach Qualité IA qui analyse une transcription et génère un plan de développement personnalisé."
))

DAYS.append(dict(
    day_num=21, week=4, phase=2, phase_name="IA × Centres d'Appels", day_name="Mercredi",
    title="Feedback Automatisé et Coaching IA", subtitle="Analyse transcriptions · Feedback structuré · Plan développement",
    accent_color=ORANGE,
    theory_title="Le Coaching IA — Feedback Personnalisé à l'Échelle",
    theory_points=[
        ("Problème du coaching classique","Superviseur : 1 session coaching / semaine / agent = coaching insuffisant pour progresser"),
        ("Solution IA","Chaque transcription analysée → feedback structuré immédiat → plan de développement personnalisé"),
        ("5 dimensions de feedback","Technique (procédure) / Communication / Empathie / Efficacité / Conformité"),
        ("Ton du feedback IA","Spécifique / factuel / bienveillant / orienté action — à paramétrer dans votre prompt"),
        ("Rôle du superviseur","Valide le feedback IA / ajoute la dimension relationnelle / suit le plan de développement"),
        "Le coaching IA ne remplace pas le superviseur — il lui donne 5x plus de matière à travailler",
    ],
    theory_insight="Un outil de coaching IA bien configuré est l'un des livrables les plus demandés par les responsables CA.",
    exercise_title="Configurer le Prompt Coach Qualité IA",
    exercise_steps=[
        ("⚙️","Écrire le prompt coach","Dimensions analysées / format du feedback / ton / exemples positifs et axes d'amélioration",ORANGE),
        ("🧪","Tester sur 3 transcriptions","Appliquez votre prompt coach sur les 3 transcriptions créées Jour 13 — comparez",TEAL),
        ("📝","Calibrer le feedback","Ajustez le prompt : trop sévère ? trop vague ? pas assez actionnable ?",GREEN),
        ("🗂️","Plan de développement","Créez le prompt qui génère un plan de développement 30 jours à partir du feedback",NAVY),
    ],
    production_title="Outil Coach Qualité IA Complet",
    production_items=[
        ("Prompt d'analyse","Analyse une transcription → feedback 5 dimensions → points forts / axes amélioration"),
        ("Prompt plan de développement","À partir du feedback → plan d'action 30 jours / 3 objectifs / activités / indicateurs"),
        ("Guide superviseur","Comment utiliser l'outil / fréquence / comment présenter le feedback à l'agent"),
        "Démontrez cet outil à un prospect : 10 minutes de démo = plus convaincant que 10 pages de brochure",
    ],
    livrable="Outil Coach Qualité IA — Feedback + Plan Développement",
    livrable_desc="Prompts configurés + guide superviseur — outil différenciateur dans votre offre accompagnement.",
    tomorrow_title="Micro-Learning Mobile",
    tomorrow_desc="Créez une série de 10 modules micro-learning 'Techniques de vente' pour mobile — format 2-5 minutes."
))

DAYS.append(dict(
    day_num=22, week=4, phase=2, phase_name="IA × Centres d'Appels", day_name="Jeudi",
    title="Micro-Learning Mobile", subtitle="Formats courts · Ancrage mémoriel · Série techniques vente",
    accent_color=GOLD,
    theory_title="Micro-Learning — Apprendre en 3 Minutes",
    theory_points=[
        ("Définition","Contenus courts (2-5 min) / un seul objectif / format mobile-first / répétition espacée"),
        ("Neuro-pédagogie","Courbes d'oubli : sans rappel 70% oublié en 24h / avec micro-learning espacé : 90% retenu"),
        ("Formats efficaces","Vidéo courte / mini-quiz / carte mémo / simulation rapide / audio"),
        ("Série cohérente","10 modules = 1 programme complet consultable en 30 min / module"),
        ("IA et micro-learning","Claude génère le script / Canva produit les visuels / Loom enregistre la voix"),
        "Le micro-learning est le format le plus adopté en CA (96% de taux de complétion vs 15% e-learning classique)",
    ],
    theory_insight="10 micro-modules bien construits = programme formation complet que les agents consultent sur leur téléphone.",
    exercise_title="Structurer 10 Modules Micro-Learning",
    exercise_steps=[
        ("📱","Définir les 10 thèmes","Techniques de vente : accroche / qualification / écoute / reformulation / objections / closing / etc.",GOLD),
        ("📝","Template par module","Contexte (30s) / Technique expliquée (90s) / Exemple audio (60s) / Mémo visuel / Mini-quiz",ORANGE),
        ("✍️","Rédiger 3 scripts","Modules 1-3 complets avec script mot à mot — à lire devant Loom ou enregistreur",TEAL),
        ("🎨","Visuels Canva","Pour chaque module : 1 carte mémo A4 téléchargeable — design cohérent série",GREEN),
    ],
    production_title="Série Micro-Learning — 10 Techniques de Vente",
    production_items=[
        ("10 scripts complets","Un script par technique — rédigés par Claude, validés et personnalisés par vous"),
        ("10 cartes mémo","Visuels Canva prêts à télécharger / imprimer / afficher au poste de travail"),
        ("Guide déploiement","Comment intégrer la série dans le quotidien des agents — rappels, fréquence, suivi"),
        "Cette série est un produit digital vendable (PDF + scripts) — estimez entre 200 et 500€ le package",
    ],
    livrable="Série Micro-Learning — 10 Techniques de Vente Téléphonique",
    livrable_desc="10 scripts + 10 cartes mémo + guide déploiement — produit digital commercialisable.",
    tomorrow_title="Otter.ai et Transcription Automatique",
    tomorrow_desc="Transcrivez une session de travail et transformez-la automatiquement en supports : slides, FAQ, synthèse."
))

DAYS.append(dict(
    day_num=23, week=4, phase=2, phase_name="IA × Centres d'Appels", day_name="Vendredi",
    title="Otter.ai — Transcription et Valorisation", subtitle="Voix → Texte → Livrables en 30 minutes",
    accent_color=TEAL,
    theory_title="Otter.ai — Votre Secrétaire IA",
    theory_points=[
        ("Transcription automatique","Enregistrez n'importe quelle session → texte formaté en quelques minutes"),
        ("Cas d'usage formation","Transcription de vos propres formations → base de contenu réutilisable"),
        ("Cas d'usage qualité","Transcription de réunions de débriefing → compte-rendu automatique"),
        ("Workflow valorisation","Transcription → Claude → Résumé / FAQ / Slides / Article / Post LinkedIn"),
        ("RGPD attention","Informez toujours les participants avant d'enregistrer — obligation légale"),
        ("Alternative : Whisper","Outil open-source de transcription — option hors cloud pour données sensibles"),
        "Une heure de formation enregistrée = 5 livrables différents produits en 45 minutes",
    ],
    theory_insight="Chaque formation que vous animez est une mine de contenu — ne laissez plus passer une seule session sans la valoriser.",
    exercise_title="Enregistrer et Transcrire",
    exercise_steps=[
        ("🎙️","Enregistrement","Enregistrez 15 minutes sur votre expertise (monologue sur les techniques de vente CA)",TEAL),
        ("📝","Transcription Otter","Uploadez l'audio dans Otter.ai / transcrivez / vérifiez la qualité",GREEN),
        ("🔄","Workflow Claude","Copiez la transcription dans Claude — demandez : résumé / 5 points clés / FAQ / titres slides",ORANGE),
        ("📊","Vérifier les livrables","Chaque livrable issu de la transcription est-il publiable immédiatement ?",NAVY),
    ],
    production_title="Workflow — Réunion/Formation → 5 Livrables",
    production_items=[
        ("Livrable 1","Résumé exécutif 1 page — pour envoyer après une réunion client"),
        ("Livrable 2","FAQ 10 questions-réponses — à publier sur votre site ou Notion"),
        ("Livrable 3","Slides Gamma 10 slides — présentation de la session"),
        "Documentez ce workflow — vendez-le comme service de valorisation de contenus existants",
    ],
    livrable="Workflow 'Réunion → 5 Livrables en 45 min'",
    livrable_desc="Process documenté + exemple concret — service vendable aux entreprises qui veulent valoriser leurs réunions.",
    tomorrow_title="Bilan Semaine 4 — Catalogue Formations",
    tomorrow_desc="Assemblez un catalogue de vos formations commercialisables et calculez votre CA potentiel mensuel."
))

DAYS.append(dict(
    day_num=24, week=4, phase=2, phase_name="IA × Centres d'Appels", day_name="Samedi",
    title="Bilan Semaine 4 — Catalogue Formations", subtitle="Packager · Chiffrer · Préparer le catalogue",
    accent_color=GREEN,
    theory_title="Construire un Catalogue de Formations",
    theory_points=[
        ("Inventaire des modules","Listez tous vos modules créés — durée / public / objectifs / format"),
        ("Regrouper en programmes","Assembler des modules en parcours cohérents de 1 jour / 2 jours / 3 mois"),
        ("Tarification formation","Journée présentiel / Module e-learning / Accompagnement retainer — 3 modèles"),
        ("Qualiopi et IA","Si vous êtes ou souhaitez être certifiée Qualiopi — comment intégrer l'IA dans le dossier"),
        ("Catalogue visuel","Canva : une page par formation — accroche / objectifs / public / durée / prix / contact"),
        "Un catalogue structuré multiplie par 3 la conversion des prospects en clients",
    ],
    theory_insight="Vous avez maintenant assez de modules pour construire un catalogue complet — il ne manque que la mise en forme.",
    exercise_title="Inventaire et Assemblage des Modules",
    exercise_steps=[
        ("📋","Inventaire complet","Listez tous les modules créés Jours 1-24 avec durée et public cible",GREEN),
        ("🗂️","3 programmes","Assemblez : Programme Agent CA (3 jours) / Manager IA (2 jours) / Formateur IA (2 jours)",TEAL),
        ("💰","Chiffrer chaque formation","TJM × jours + préparation + supports — prix marché pour référence",ORANGE),
        ("🎨","Mise en page Canva","1 fiche par formation — design cohérent avec votre charte graphique",NAVY),
    ],
    production_title="Catalogue Formations TAMOU NEURAL PATH",
    production_items=[
        ("Page de couverture","Votre nom / positionnement / accroche / coordonnées"),
        ("3 fiches formations","Agent CA / Manager IA / Formateur IA — design Canva professionnel"),
        ("Page tarifs et conditions","Modalités / durées / formats / réductions volume / CGV résumées"),
        "Envoyez ce catalogue à 5 contacts ce weekend — feedback = or pour votre développement commercial",
    ],
    livrable="Catalogue Formations Freelance IA",
    livrable_desc="3 formations packagées + tarifs — document commercial prêt à diffuser.",
    tomorrow_title="Communication Managériale Augmentée",
    tomorrow_desc="Semaine 5 : créez une bibliothèque de 20 templates de communication managériale pour centres d'appels."
))

DAYS.append(dict(
    day_num=25, week=5, phase=2, phase_name="IA × Centres d'Appels", day_name="Lundi",
    title="Communication Managériale Augmentée", subtitle="Templates · Registres · Bibliothèque de 20 modèles",
    accent_color=NAVY,
    theory_title="IA et Communication Managériale",
    theory_points=[
        ("Enjeu","La communication managériale impacte directement la motivation, la conformité et la rétention"),
        ("Communication descendante","Du manager vers l'équipe : instructions / résultats / recadrage / félicitations"),
        ("Communication ascendante","De l'agent vers le manager : remontée terrain / alertes / idées d'amélioration"),
        ("IA et adaptation","L'IA adapte le ton, le registre et la longueur selon le contexte — vous définissez le message"),
        ("5 registres","Formel / bienveillant / direct / motivationnel / recadrage — mêmes infos, ton différent"),
        "Un message bien formulé = 80% d'adhésion. Mal formulé = résistance même pour un bon projet",
    ],
    theory_insight="Vos 15 ans de communication terrain vous permettent de valider les nuances que l'IA peut rater.",
    exercise_title="Rédiger le même message en 5 Registres",
    exercise_steps=[
        ("📝","Message source","'Les résultats de la semaine sont en dessous de l'objectif. Réunion vendredi.'",NAVY),
        ("✍️","5 versions","Claude génère : email formel / message vocal / réunion équipe / note écrite / SMS urgent",RED),
        ("🔍","Analyse critique","Quelle version sonne le plus juste ? Qu'est-ce que vous modifieriez ?",TEAL),
        ("📚","Bibliothèque","Créez un template pour chacun des 20 cas les plus fréquents de votre expérience",GREEN),
    ],
    production_title="Bibliothèque de 20 Templates Communication",
    production_items=[
        ("10 templates descendants","Félicitations / résultats / objectifs / recadrage / annonce changement / formation obligatoire / etc."),
        ("10 templates ascendants","Remontée problème / idée amélioration / alerte qualité / demande ressource / retour formation"),
        ("Guide d'utilisation","Comment adapter chaque template à son propre style — pas de copier-coller brut"),
        "Ce kit est un produit digital vendable aux managers CA — entre 49 et 149€",
    ],
    livrable="Bibliothèque 20 Templates Communication Managériale CA",
    livrable_desc="20 modèles rédigés pour 20 situations clés — produit digital commercialisable.",
    tomorrow_title="Soft Skills et IA",
    tomorrow_desc="Créez un atelier complet 7h 'Soft Skills × IA' avec activités interactives et jeux de rôle simulés."
))

DAYS.append(dict(
    day_num=26, week=5, phase=2, phase_name="IA × Centres d'Appels", day_name="Mardi",
    title="Soft Skills et IA — Atelier Complet", subtitle="Empathie · Assertivité · Écoute · Programme 7h",
    accent_color=TEAL,
    theory_title="enseigner les Soft Skills avec l'IA",
    theory_points=[
        ("Soft skills prioritaires CA","Empathie / écoute active / assertivité / gestion du stress / leadership / communication"),
        ("Limite de l'IA","L'IA ne peut pas évaluer l'authenticité d'une empathie — votre regard humain est irremplaçable"),
        ("IA comme outil de pratique","Simulations répétées / feedback immédiat / cas variés infinis"),
        ("Ancrage comportemental","La simulation IA crée de la mémoire musculaire verbale — ancrage durable"),
        ("Format atelier idéal","Alternance : concept (15 min) / démonstration IA (10 min) / pratique (25 min) / débriefing (10 min)"),
        "Les soft skills développées avec IA + coaching humain progressent 2x plus vite qu'en formation classique",
    ],
    theory_insight="Votre rôle n'est pas de transmettre des soft skills — c'est de créer les conditions pour que les apprenants les découvrent.",
    exercise_title="Concevoir l'Architecture de l'Atelier",
    exercise_steps=[
        ("🏗️","Structure 7h","Matin : concepts et démonstrations (3h) / Après-midi : pratique intensive et débriefing (4h)",TEAL),
        ("🎭","Activités IA","3 simulations Claude / 2 jeux de rôle inter-apprenants / 1 analyse transcription collective",GREEN),
        ("📊","Évaluation","Grille d'observation comportementale / auto-évaluation avant-après / plan de développement",ORANGE),
        ("📦","Kit participant","Livret apprenant / fiches mémo / plan d'engagement personnel post-atelier",NAVY),
    ],
    production_title="Programme Atelier Soft Skills × IA — 7h",
    production_items=[
        ("Programme détaillé","Minute par minute : activité / durée / matériel / animation / objectif"),
        ("Supports formateur","Guide animateur / fiches activités / grilles d'observation / corrigés débriefing"),
        ("Kit participant","Livret 20 pages / fiches mémo / plan d'engagement / ressources complémentaires"),
        "Atelier vendable en intra-entreprise : 800-1500€ la journée — l'un de vos produits phares",
    ],
    livrable="Programme Atelier Soft Skills × IA — 7h Clé en Main",
    livrable_desc="Guide formateur + supports + kit participant — formation intra-entreprise prête à délivrer.",
    tomorrow_title="Recrutement IA-assisté",
    tomorrow_desc="Créez un kit recrutement complet pour centres d'appels : fiche poste + grille entretien + onboarding 30 jours."
))

DAYS.append(dict(
    day_num=27, week=5, phase=2, phase_name="IA × Centres d'Appels", day_name="Mercredi",
    title="Recrutement IA-assisté", subtitle="Fiche poste · Grille entretien · Onboarding 30 jours",
    accent_color=ORANGE,
    theory_title="IA et Recrutement — Opportunités et Vigilance",
    theory_points=[
        ("Gain de temps IA","Fiche de poste / annonce / grille entretien / scoring → 4h de travail en 45 min"),
        ("Biais algorithmiques","L'IA peut reproduire des biais de genre, d'âge, de diplôme — validation humaine obligatoire"),
        ("RGPD recrutement","Données candidats : minimisation / durée conservation / droit à l'information / pas de scoring opaque"),
        ("Scoring structuré","Grille pondérée = décisions objectives / traçables / défendables légalement"),
        ("Onboarding IA","Parcours d'intégration personnalisé / ressources adaptées / suivi automatique des étapes"),
        "Les entreprises avec process recrutement structuré retiennent 50% mieux leurs nouvelles recrues",
    ],
    theory_insight="Vous connaissez les critères de réussite d'un agent CA — utilisez l'IA pour les formaliser et les systématiser.",
    exercise_title="Créer les 3 Documents Recrutement",
    exercise_steps=[
        ("📄","Fiche de poste","Conseiller clientèle CA entrant : missions / compétences / profil / conditions — version IA affinée par vous",ORANGE),
        ("📋","Grille entretien","10 questions structurées avec critères d'évaluation / 5 niveaux de scoring / pondération",RED),
        ("🗺️","Plan onboarding","30 jours : semaine 1 (découverte) / semaine 2-3 (montée en puissance) / semaine 4 (autonomie)",TEAL),
        ("✅","Checklist RGPD","Vérification conformité : information candidat / conservation données / critères objectifs",GREEN),
    ],
    production_title="Kit Recrutement Complet CA",
    production_items=[
        ("4 documents","Fiche de poste + annonce multi-canal + grille entretien scoring + plan onboarding 30 jours"),
        ("Checklist RGPD","Document de conformité inclus — rassure les clients DRH"),
        ("Mise en page","Canva : design professionnel cohérent avec votre charte"),
        "Ce kit est une offre standalone vendable 500-1500€ aux RH de centres d'appels",
    ],
    livrable="Kit Recrutement CA Complet — 4 Documents",
    livrable_desc="Fiche poste + annonce + grille entretien + onboarding + checklist RGPD — offre RH packagée.",
    tomorrow_title="Fédérer et Motiver avec l'IA",
    tomorrow_desc="Créez un outil de diagnostic motivation d'équipe avec scoring automatique et plan d'action managérial."
))

DAYS.append(dict(
    day_num=28, week=5, phase=2, phase_name="IA × Centres d'Appels", day_name="Jeudi",
    title="Fédérer et Motiver avec l'IA", subtitle="Diagnostic motivation · Scoring auto · Plan d'action manager",
    accent_color=RED,
    theory_title="Motivation et IA — Modèles et Outils",
    theory_points=[
        ("Modèles de référence","Maslow (besoins) / Herzberg (facteurs) / Self-Determination Theory (autonomie, compétence, appartenance)"),
        ("Signaux faibles à détecter","Absentéisme / retards / qualité en baisse / interactions réduite / erreurs inhabituelles"),
        ("IA et diagnostic","15 questions ciblées → scoring automatique → identification des leviers → plan d'action personnalisé"),
        ("Communication individuelle","Adapter le discours managérial au profil motivationnel de chaque agent"),
        ("Limite IA","L'IA détecte des patterns — le manager comprend les causes. Toujours croiser avec le terrain"),
        "Un diagnostic motivation IA prend 10 minutes / agent — vs 1h d'entretien classique pour moins de données",
    ],
    theory_insight="Votre expertise des dynamiques d'équipe CA est la clé — l'IA structure et mesure, vous interprétez et agissez.",
    exercise_title="Construire le Baromètre Motivation",
    exercise_steps=[
        ("❓","15 questions","Rédigez avec Claude 15 questions sur : charge de travail / relations / sens / autonomie / reconnaissance",RED),
        ("📊","Système de scoring","Pondération par dimension / score global / profil motivationnel / zones d'alerte",ORANGE),
        ("💡","Recommandations auto","Pour chaque profil : 3 leviers managériaux concrets recommandés automatiquement",TEAL),
        ("🧪","Tester l'outil","Remplissez vous-même le baromètre — les résultats sont-ils pertinents ?",GREEN),
    ],
    production_title="Outil Baromètre Motivation Équipe",
    production_items=[
        ("Questionnaire Typeform","15 questions avec scoring automatique — lien partageable à l'équipe"),
        ("Grille d'analyse Claude","Prompt qui analyse les résultats collectifs et génère un plan d'action équipe"),
        ("Rapport manager","Template de présentation des résultats à l'équipe — transparence et co-construction"),
        "Proposez cet outil gratuit en lead magnet — génère des conversations commerciales avec les managers CA",
    ],
    livrable="Baromètre Motivation Équipe CA — Outil Complet",
    livrable_desc="Questionnaire + scoring + analyse IA + rapport manager — outil diagnostic RH à forte valeur ajoutée.",
    tomorrow_title="Atelier Amélioration Continue",
    tomorrow_desc="Concevez un atelier 3h d'amélioration continue avec toutes les activités, le guide animateur et les supports."
))

DAYS.append(dict(
    day_num=29, week=5, phase=2, phase_name="IA × Centres d'Appels", day_name="Vendredi",
    title="Atelier Amélioration Continue IA", subtitle="Facilitation · PDCA · Intelligence collective · 3h",
    accent_color=GOLD,
    theory_title="Faciliter un Atelier avec l'IA",
    theory_points=[
        ("Rôle du facilitateur","Créer les conditions pour que le groupe trouve ses propres solutions — pas imposer les vôtres"),
        ("IA comme préparateur","L'IA génère les activités, les questions, les supports — vous animez avec votre expertise humaine"),
        ("Structure idéale 3h","Problème partagé (30 min) / Analyse causes (45 min) / Solutions (45 min) / Plan d'action (30 min) / Clôture (30 min)"),
        ("Techniques d'animation IA","World Café / Ishikawa collectif / Dot voting / PDCA en équipe"),
        ("Compte-rendu IA","Photo du tableau blanc → Claude → compte-rendu structuré en 5 minutes"),
        "Un atelier bien facilité génère 10x plus d'adhésion qu'une décision top-down",
    ],
    theory_insight="Votre valeur en facilitation : vous lisez la salle. L'IA prépare le contenu. Vous créez la magie collective.",
    exercise_title="Concevoir l'Atelier Minute par Minute",
    exercise_steps=[
        ("🗺️","Séquençage","Découpez les 3h en blocs avec Claude — transitions / durées / objectifs par bloc",GOLD),
        ("🎯","Activités","3 activités collaboratives : problème partagé / brainstorming structuré / vote priorités",ORANGE),
        ("📋","Supports","Fiche facilitateur / slides visuels Canva / post-its virtuels Miro si distanciel",TEAL),
        ("📝","Compte-rendu type","Créez le template de compte-rendu qui sera rempli automatiquement après l'atelier",GREEN),
    ],
    production_title="Kit Atelier Amélioration Continue — 3h",
    production_items=[
        ("Guide facilitateur","Minute par minute : activité / durée / consigne / matériel / observation clé"),
        ("Supports visuels","Slides Canva pour chaque étape / fiches activité / exemple rempli"),
        ("Template compte-rendu","Structure à remplir pendant / après l'atelier — action / responsable / délai / indicateur"),
        "Proposez cet atelier en intra à 800-1500€ — livrable à forte valeur perçue car format participatif",
    ],
    livrable="Kit Atelier Amélioration Continue — 3h Clé en Main",
    livrable_desc="Guide facilitateur + supports visuels + template compte-rendu — atelier intra-entreprise prêt à animer.",
    tomorrow_title="Évaluation Phase 2 — Niveau 2 Validé ?",
    tomorrow_desc="Quiz Phase 2 + portfolio 25 livrables + vérification de votre niveau 2 d'application."
))

DAYS.append(dict(
    day_num=30, week=5, phase=2, phase_name="IA × Centres d'Appels", day_name="Samedi",
    title="Évaluation Phase 2 — Niveau 2 Validé", subtitle="Quiz · Portfolio · Passage Phase 3",
    accent_color=GREEN,
    theory_title="Grille Évaluation — Niveau 2 Application",
    theory_points=[
        ("Critère 1 /20","Générer un module formation complet en moins de 3h — testé en conditions réelles"),
        ("Critère 2 /20","Créer une grille QA appels entrants conforme ISO 9001 — sans aide"),
        ("Critère 3 /20","Configurer un jeu de rôle IA fonctionnel pour la formation"),
        ("Critère 4 /20","25 livrables cumulés dans le portfolio — documentés et organisés"),
        ("Critère 5 /20","Maîtriser 7 outils IA de façon professionnelle et autonome"),
        "Score ≥ 60/100 = Niveau 2 validé. Vous passez en Phase 3 : IA × Qualité ISO.",
    ],
    theory_insight="Phase 2 terminée — vous êtes maintenant la référence IA × Formation × Centres d'Appels dans votre région.",
    exercise_title="Quiz Phase 2 — 20 Questions",
    exercise_steps=[
        ("📝","Questions 1-5","NLP / analyse sentiment / outils QA IA / scripts adaptatifs / KPIs prédictifs",GREEN),
        ("📝","Questions 6-10","Adaptive learning / jeux de rôle IA / coaching IA / micro-learning / transcription",TEAL),
        ("📝","Questions 11-15","Communication managériale / soft skills IA / recrutement RGPD / motivation / facilitation",ORANGE),
        ("📝","Questions 16-20","Présenter 3 offres packagées / décrire 2 livrables à valeur / expliquer votre PVU",NAVY),
    ],
    production_title="Portfolio Phase 2 — Consolidation",
    production_items=[
        ("25 livrables documentés","Chaque livrable : titre / date / valeur commerciale estimée / outils utilisés"),
        ("2 offres packagées finalisées","Offre 1 Audit Qualité + Offre 2 Formation CA — prêtes à envoyer"),
        ("Score et plan de suite","Calculez votre score / identifiez les 2 points à renforcer en Phase 3"),
        "Vous avez maintenant assez de livrables pour décrocher votre première mission freelance — candidatez !",
    ],
    livrable="Portfolio Phase 2 — 25 Livrables + 2 Offres Packagées",
    livrable_desc="Phase 2 complétée. Score ≥ 60/100 = Niveau 2 certifié. Prête pour Phase 3 : IA × Qualité ISO.",
    tomorrow_title="Phase 3 — IA × Qualité ISO & Amélioration Continue",
    tomorrow_desc="Cartographie processus, audits internes, 8D, PDCA IA — vous devenez la référence ISO × IA en France."
))

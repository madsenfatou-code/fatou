# Days 31–54
from pptx_helpers import *

DAYS = []

DAYS.append(dict(
    day_num=31,week=6,phase=3,phase_name="Qualité & ISO",day_name="Lundi",
    title="Cartographie des Processus IA",subtitle="SIPOC augmenté · Logigrammes · Documentation SMQ",
    accent_color=ORANGE,
    theory_title="Cartographie Processus avec l'IA",
    theory_points=[
        ("SIPOC","Suppliers / Inputs / Process / Outputs / Customers — cartographie rapide assistée IA"),
        ("Logigramme textuel","Claude génère la description étape par étape — vous validez et transformez en visuel"),
        ("Gains IA","Une cartographie qui prenait 2 jours se produit en 3 heures avec l'IA"),
        ("ISO 9001 ch.4 & 8","Approche processus obligatoire — l'IA structure et formate selon les exigences"),
        ("Processus prioritaires CA","Traitement réclamation / Escalade / Monitoring qualité / Intégration nouvel agent"),
        "Une cartographie claire = base de tout SMQ et de toute démarche d'amélioration",
    ],
    theory_insight="Votre connaissance terrain des processus CA est irremplaçable — l'IA les met en forme 10x plus vite.",
    exercise_title="Cartographier 3 Processus Clés",
    exercise_steps=[
        ("🗺️","Processus 1","Traitement réclamation client : SIPOC complet avec Claude",ORANGE),
        ("🗺️","Processus 2","Escalade appel : arbre de décision + logigramme",RED),
        ("🗺️","Processus 3","Monitoring qualité appels : de l'enregistrement au feedback agent",TEAL),
        ("📄","Documentation","Formatez chaque cartographie selon les exigences ISO 9001 v2015",GREEN),
    ],
    production_title="3 Cartographies Processus Formatées ISO",
    production_items=[
        ("SIPOC × 3","Un SIPOC par processus — format tableau clair et utilisable"),
        ("Logigrammes textuels","Description étape par étape avec points de décision et jalons"),
        ("Fiches processus","Fiche 1 page par processus : propriétaire / indicateurs / risques / documents associés"),
        "Ces 3 cartographies constituent la base documentaire d'un SMQ — valeur conseil élevée",
    ],
    livrable="3 Cartographies Processus CA — Format ISO 9001",
    livrable_desc="SIPOC + logigrammes + fiches processus — base documentaire SMQ prête à intégrer.",
    tomorrow_title="Audit Interne IA-assisté",
    tomorrow_desc="Créez un plan d'audit interne complet avec check-list de terrain et template de rapport automatisé."
))

DAYS.append(dict(
    day_num=32,week=6,phase=3,phase_name="Qualité & ISO",day_name="Mardi",
    title="Audit Interne IA-assisté",subtitle="Plan d'audit · Check-list terrain · Rapport automatisé",
    accent_color=NAVY,
    theory_title="L'Audit Interne Augmenté par l'IA",
    theory_points=[
        ("Gain de temps IA","Plan d'audit + check-list + rapport = 1 journée → 2 heures avec l'IA"),
        ("Qualité préservée","L'IA structure — votre expertise terrain garantit la pertinence des constats"),
        ("Plan d'audit ISO","Périmètre / objectifs / critères / méthode / équipe / calendrier / livrables"),
        ("Check-list intelligente","Questions adaptées au contexte CA — pas générique, spécifique à votre secteur"),
        ("Rapport automatisé","À partir de vos notes de terrain → Claude génère le rapport structuré en 20 minutes"),
        "Un audit interne bien préparé avec l'IA détecte 40% de non-conformités supplémentaires",
    ],
    theory_insight="L'IA prépare et formate — vous observez, constatez et jugez. Le jugement audit reste humain.",
    exercise_title="Préparer un Audit Fictif Complet",
    exercise_steps=[
        ("📋","Plan d'audit","Département formation d'un CA fictif de 100 agents — périmètre / objectifs / calendrier",NAVY),
        ("✅","Check-list terrain","40 questions spécifiques CA : formation initiale / continue / évaluation / documentation",RED),
        ("📝","Simulation terrain","Jouez l'auditeur : Claude joue le responsable formation — 3 constats fictifs",TEAL),
        ("📄","Rapport auto","À partir de vos 3 constats : demandez à Claude de générer le rapport d'audit complet",GREEN),
    ],
    production_title="Kit Audit Interne Complet",
    production_items=[
        ("Plan d'audit","Template réutilisable pour tout département formation CA"),
        ("Check-list 40 questions","Organisée par chapitre ISO — coche et notes terrain"),
        ("Template rapport","Structure standard : contexte / méthode / constats / non-conformités / recommandations"),
        "Ce kit vaut 1 à 3 jours de prestation — packagez-le dans votre offre conseil qualité",
    ],
    livrable="Kit Audit Interne Formation CA — Plan + Check-list + Rapport",
    livrable_desc="3 documents prêts à l'emploi — intégrés dans votre offre d'audit qualité.",
    tomorrow_title="Non-Conformités et Méthode 8D",
    tomorrow_desc="Maîtrisez la méthode 8D augmentée par l'IA — traitez 3 non-conformités fictives en moins de 2 heures."
))

DAYS.append(dict(
    day_num=33,week=6,phase=3,phase_name="Qualité & ISO",day_name="Mercredi",
    title="Non-Conformités et Méthode 8D IA",subtitle="Analyse causes racines · 5 Pourquoi · Plans d'action",
    accent_color=RED,
    theory_title="La Méthode 8D Augmentée",
    theory_points=[
        ("8D = 8 Disciplines","D1 Équipe / D2 Problème / D3 Containment / D4 Causes / D5 Actions / D6 Vérification / D7 Prévention / D8 Clôture"),
        ("IA et D4","Les 5 Pourquoi assistés par IA — l'IA propose des causes probables, vous validez avec le terrain"),
        ("IA et D5","Génération automatique d'actions correctives avec responsables et délais"),
        ("Bibliothèque de causes","L'IA connaît les causes les plus fréquentes en CA — gain de temps pour le diagnostic"),
        ("Traçabilité ISO","Chaque 8D documenté dans le format requis pour l'audit — automatiquement"),
        "Un 8D complet et bien documenté est la preuve la plus forte de votre maîtrise SMQ",
    ],
    theory_insight="La méthode 8D n'a pas de secret pour vous — l'IA vous libère de la mise en forme pour vous concentrer sur le fond.",
    exercise_title="Traiter 3 Non-Conformités Fictives",
    exercise_steps=[
        ("🔴","NC1","Taux de réclamations 3 points au-dessus de l'objectif — 8D complet avec IA",RED),
        ("🔴","NC2","Module formation initiale non délivré dans les délais ISO — analyse causes + actions",ORANGE),
        ("🔴","NC3","Grilles d'évaluation non remplies par 30% des superviseurs — plan de correction",TEAL),
        ("📚","Template 8D","Créez le template réutilisable pour tout type de NC en CA",GREEN),
    ],
    production_title="Template 8D IA-assisté + 3 Exemples",
    production_items=[
        ("Template 8D","Format Word/Notion avec les 8 sections — champs à remplir + exemples par section"),
        ("Bibliothèque causes CA","Top 20 causes racines les plus fréquentes en CA — avec facteurs contributifs"),
        ("3 8D complétés","Les 3 NC fictives traitées de A à Z — références portfolio"),
        "Ce template est vendable en format numérique ou intégré dans votre mission d'audit",
    ],
    livrable="Template 8D IA-assisté + Bibliothèque Causes CA",
    livrable_desc="Outil de traitement NC complet — 3 exemples concrets inclus pour démonstration client.",
    tomorrow_title="Indicateurs Qualité Prédictifs",
    tomorrow_desc="Construisez un système d'alertes qualité précoces qui anticipe les dérives avant qu'elles deviennent des problèmes."
))

DAYS.append(dict(
    day_num=34,week=6,phase=3,phase_name="Qualité & ISO",day_name="Jeudi",
    title="Indicateurs Qualité Prédictifs",subtitle="Signaux faibles · Alertes précoces · IA prédictive",
    accent_color=GOLD,
    theory_title="L'IA Prédictive pour la Qualité",
    theory_points=[
        ("Différence curatif/prédictif","Curatif : traiter le problème après / Prédictif : détecter le signal avant"),
        ("Signaux faibles CA","Légère hausse DMT / micro-augmentation escalades / baisse imperceptible satisfaction"),
        ("Corrélations IA","L'IA détecte des liens non évidents : absentéisme ↔ qualité / météo ↔ agressivité clients"),
        ("Seuils d'alerte intelligents","Pas un seuil fixe — un seuil adaptatif selon la saisonnalité et le contexte"),
        ("Tableau de bord prédictif","Tendances / projections / alertes précoces / recommandations contextuelles"),
        "Passer du curatif au prédictif = passer de manager de crise à architecte de la performance",
    ],
    theory_insight="Votre expérience vous permet de définir les bons signaux faibles — l'IA les surveille en permanence.",
    exercise_title="Identifier et Modéliser les Signaux Faibles",
    exercise_steps=[
        ("📡","Brainstorm signaux","Listez avec Claude les 15 signaux faibles les plus pertinents pour un CA entrant",GOLD),
        ("🔗","Corrélations","Pour chaque signal : quelle dérive prédit-il ? dans quel délai ? quelle action préventive ?",ORANGE),
        ("⚠️","Seuils adaptatifs","Définissez les seuils d'alerte avec contexte (lundi matin / veille vacances / etc.)",RED),
        ("📊","Maquette dashboard","Canva : tableau de bord prédictif avec jauges tendances et zones d'alerte couleur",TEAL),
    ],
    production_title="Système d'Alertes Qualité Précoces",
    production_items=[
        ("Catalogue 15 signaux faibles","Chaque signal : définition / méthode de mesure / seuil alerte / action recommandée"),
        ("Dashboard prédictif","Maquette visuelle Canva — opérationnel à adapter dans n'importe quel outil BI"),
        ("Prompt d'analyse hebdomadaire","Chaque lundi : collez vos données → Claude génère l'analyse prédictive de la semaine"),
        "Ce système transforme votre offre de conseil : vous ne traitez plus les problèmes, vous les évitez",
    ],
    livrable="Système d'Alertes Qualité Précoces — 15 Signaux + Dashboard",
    livrable_desc="Catalogue signaux + dashboard prédictif + prompt analyse — offre conseil à très forte valeur ajoutée.",
    tomorrow_title="PDCA × IA — Plan d'Amélioration Continue",
    tomorrow_desc="Créez un Plan d'Amélioration Continue complet pour un CA fictif — 90 jours, 5 chantiers, 25 actions."
))

DAYS.append(dict(
    day_num=35,week=6,phase=3,phase_name="Qualité & ISO",day_name="Vendredi",
    title="PDCA × IA — Plan d'Amélioration Continue",subtitle="90 jours · 5 chantiers · 25 actions · Indicateurs",
    accent_color=TEAL,
    theory_title="Le PDCA Accéléré par l'IA",
    theory_points=[
        ("PDCA classique","Plan (planifier) / Do (faire) / Check (vérifier) / Act (ajuster) — cycle vertueux"),
        ("IA et PLAN","Génération automatique d'actions à partir du diagnostic — priorisation intelligente"),
        ("IA et CHECK","Analyse des données de suivi — détection des écarts et causes en temps réel"),
        ("IA et ACT","Recommandations d'ajustement contextualisées — apprentissage continu"),
        ("Structure PAC","Chantiers prioritaires / actions / responsables / délais / indicateurs / jalons"),
        "Un PAC IA-structuré est produit 5x plus vite et intègre automatiquement les meilleures pratiques sectorielles",
    ],
    theory_insight="Votre valeur dans le PAC : vous connaissez les vrais obstacles humains et organisationnels que l'IA ignore.",
    exercise_title="Construire le PAC 90 Jours",
    exercise_steps=[
        ("🎯","Diagnostic de départ","CA fictif : 3 NC identifiées + 2 risques détectés → base du PAC",TEAL),
        ("🗂️","5 chantiers","Qualité appels / Formation continue / Processus / Management / Outils — avec Claude",GREEN),
        ("📋","25 actions","5 actions par chantier : quoi / qui / quand / comment mesurer",ORANGE),
        ("📊","Tableau de pilotage","Format Notion ou Excel : RAG status / avancement / alertes",NAVY),
    ],
    production_title="PAC Complet — CA Fictif 90 Jours",
    production_items=[
        ("Document PAC","Introduction / diagnostic / 5 chantiers / 25 actions / planning Gantt / indicateurs"),
        ("Tableau de pilotage","Outil de suivi hebdomadaire avec RAG et calcul d'avancement automatique"),
        ("Présentation comité","10 slides Gamma pour présenter le PAC à la direction"),
        "Ce PAC est le livrable central de votre Offre 1 Audit Qualité — votre preuve de valeur",
    ],
    livrable="Plan d'Amélioration Continue 90 jours — CA Fictif Complet",
    livrable_desc="PAC structuré + tableau de pilotage + présentation direction — livrable phare de votre offre conseil.",
    tomorrow_title="Offre SMQ & Amélioration Continue",
    tomorrow_desc="Assemblez vos livrables de la semaine en votre Offre 2 packagée — SMQ et amélioration continue IA."
))

DAYS.append(dict(
    day_num=36,week=6,phase=3,phase_name="Qualité & ISO",day_name="Samedi",
    title="Bilan Semaine 6 — Offre SMQ & Amélioration Continue",subtitle="Packager l'Offre 2 · Étude de cas · Présentation",
    accent_color=GREEN,
    theory_title="Packager une Offre Conseil Qualité",
    theory_points=[
        ("Vos livrables de la semaine","3 cartographies / kit audit / template 8D / système alertes / PAC 90 jours"),
        ("Offre 2 = mission complète","Audit → diagnostic → PAC → accompagnement suivi — 3 à 6 mois"),
        ("Différenciation IA","Vous livrez en 3 semaines ce qu'un cabinet fait en 3 mois — et avec plus de précision"),
        ("Étude de cas fictive","Racontez la mission du début à la fin — le prospect se reconnaît dans l'histoire"),
        ("Tarification","Mission complète 3 mois : 3000 à 8000€ selon taille du CA — justifiez par les économies"),
        "Une offre avec étude de cas se vend 40% plus facilement qu'une offre sans preuve",
    ],
    theory_insight="L'étude de cas fictive est aussi convaincante qu'une vraie référence — elle montre votre méthode en action.",
    exercise_title="Construire l'Étude de Cas et l'Offre",
    exercise_steps=[
        ("📖","Étude de cas","CA fictif de 80 agents : problème / intervention / méthode / résultats chiffrés",GREEN),
        ("💼","Fiche offre 2","Titre accrocheur / problème client / votre solution / livrables / durée / prix",TEAL),
        ("🎯","Présentation Gamma","15 slides : contexte marché / votre approche / livrables / étude de cas / CTA",ORANGE),
        ("📧","Email de prospection","Template email pour envoyer l'offre à un responsable qualité ou DG de CA",NAVY),
    ],
    production_title="Offre 2 Packagée — SMQ & Amélioration Continue IA",
    production_items=[
        ("Étude de cas 3 pages","Narrative complète : avant / pendant / après — résultats quantifiés"),
        ("Fiche offre 1 page","Canva design — accroche / livrables / prix / contact"),
        ("Présentation 15 slides","Gamma — prête à projeter chez un client"),
        "Vous avez maintenant 2 offres packagées — assez pour prospecter activement",
    ],
    livrable="Offre 2 — SMQ & Amélioration Continue + Étude de Cas",
    livrable_desc="Fiche offre + étude de cas + présentation + email — Offre 2 prête à diffuser.",
    tomorrow_title="Introduction à Make — Automatisation No-Code",
    tomorrow_desc="Semaine 7 : créez votre premier workflow automatisé — formulaire → analyse IA → email feedback personnalisé."
))

DAYS.append(dict(
    day_num=37,week=7,phase=3,phase_name="Qualité & ISO",day_name="Lundi",
    title="Make — Automatisation No-Code",subtitle="Workflows · Triggers · Premier scénario opérationnel",
    accent_color=ORANGE,
    theory_title="Make — La Logique des Workflows Automatisés",
    theory_points=[
        ("Principe","Connecter des applications entre elles sans coder — si X se passe → faire Y automatiquement"),
        ("Vocabulaire clé","Scénario (workflow) / Module (application) / Trigger (déclencheur) / Action (ce qui se passe)"),
        ("Make vs Zapier","Make : plus puissant et flexible / Zapier : plus simple / Make recommandé pour usage avancé"),
        ("Cas d'usage formation","Formulaire rempli → analyse Claude → email feedback → enregistrement Airtable"),
        ("Cas d'usage qualité","Données KPI → analyse seuils → email alerte manager → création tâche suivi"),
        "1 workflow bien conçu = 2-5 heures de travail manuel économisées chaque semaine",
    ],
    theory_insight="L'automatisation no-code est votre levier de scalabilité — servir plus de clients sans travailler plus.",
    exercise_title="Créer votre Premier Scénario Make",
    exercise_steps=[
        ("🔧","Créer le compte Make","make.com → inscription → interface découverte — 30 minutes",ORANGE),
        ("🔗","Scénario 1 simple","Google Forms → Gmail : quand formulaire soumis → envoyer email de confirmation",RED),
        ("🔗","Scénario 2 avec IA","Typeform → Claude API → Gmail : évaluation → analyse → feedback personnalisé",TEAL),
        ("🧪","Tester et déboguer","Exécutez chaque scénario avec des données test — corrigez les erreurs",GREEN),
    ],
    production_title="Workflow Évaluation Formation Automatisée",
    production_items=[
        ("Flux complet","Apprenant remplit Typeform → Make analyse → Claude génère feedback → email envoyé automatiquement"),
        ("Documentation","Schéma du workflow + guide de configuration pour le reproduire chez un client"),
        ("Test complet","3 soumissions test avec profils différents — vérifiez la cohérence des feedbacks"),
        "Ce workflow seul justifie votre tarif journalier — montrez-le en démo à vos prospects",
    ],
    livrable="Workflow Évaluation Formation Automatisée — Opérationnel",
    livrable_desc="Scénario Make fonctionnel + documentation + guide de déploiement chez un client.",
    tomorrow_title="Typeform et Évaluations Intelligentes",
    tomorrow_desc="Créez 3 outils d'évaluation automatisés : à chaud, à froid J+30, satisfaction formateur."
))

DAYS.append(dict(
    day_num=38,week=7,phase=3,phase_name="Qualité & ISO",day_name="Mardi",
    title="Typeform — Évaluations Intelligentes",subtitle="Logique conditionnelle · Scoring auto · 3 outils",
    accent_color=TEAL,
    theory_title="Typeform — Formulaires Intelligents",
    theory_points=[
        ("Logique conditionnelle","Les questions s'adaptent aux réponses précédentes — expérience personnalisée"),
        ("Scoring automatique","Chaque réponse vaut des points → score calculé automatiquement → action déclenchée"),
        ("3 évaluations essentielles","À chaud (fin formation) / à froid J+30 (ancrage) / satisfaction formateur"),
        ("Intégration","Typeform + Make + Claude = évaluation intelligente + feedback IA automatique"),
        ("Analytics","Tableau de bord Typeform : taux de réponse / scores moyens / évolutions / verbatims"),
        "Une évaluation à froid J+30 est la preuve la plus solide de l'efficacité de votre formation",
    ],
    theory_insight="3 évaluations automatisées = mesure continue de l'impact de vos formations sans travail manuel.",
    exercise_title="Créer les 3 Formulaires Typeform",
    exercise_steps=[
        ("📝","Éval à chaud","Kirkpatrick niveau 1 : satisfaction / pertinence / qualité animateur — 8 questions + NPS",TEAL),
        ("📝","Éval à froid J+30","Kirkpatrick niveau 2-3 : mémorisation / application terrain / impact mesurable — 10 questions",GREEN),
        ("📝","Éval satisfaction formateur","Auto-évaluation : objectifs atteints / ajustements / points forts / axes amélioration",ORANGE),
        ("🔗","Connecter à Make","Chaque formulaire → Make → email résultats automatique + enregistrement Notion",NAVY),
    ],
    production_title="Système d'Évaluation Formation Automatisé",
    production_items=[
        ("3 formulaires Typeform","Actifs, testés, liens générés — prêts à envoyer"),
        ("3 workflows Make","Chaque formulaire déclenche automatiquement l'analyse et l'envoi des résultats"),
        ("Dashboard résultats","Vue consolidée dans Notion : toutes les évaluations / scores / tendances"),
        "Ce système prouve à vos clients que vos formations ont un impact mesurable — argument commercial fort",
    ],
    livrable="Système Évaluation Formation Automatisé — 3 Niveaux",
    livrable_desc="3 formulaires Typeform + 3 workflows Make + dashboard Notion — mesure d'impact automatisée.",
    tomorrow_title="Airtable — Votre CRM Formation Freelance",
    tomorrow_desc="Construisez votre cockpit de gestion d'activité freelance : clients, missions, modules, facturation."
))

DAYS.append(dict(
    day_num=39,week=7,phase=3,phase_name="Qualité & ISO",day_name="Mercredi",
    title="Airtable — CRM Formation Freelance",subtitle="Base de données · Vues · Automatisations · Cockpit",
    accent_color=NAVY,
    theory_title="Airtable — La Base de Données Intelligente",
    theory_points=[
        ("Au-delà d'Excel","Airtable = tableur + base de données + automatisations + intégrations"),
        ("Relations entre tables","Clients ↔ Missions ↔ Modules ↔ Évaluations — tout est lié"),
        ("Vues multiples","Vue tableau / calendrier / Kanban / galerie / timeline — même donnée, angles différents"),
        ("Automatisations natives","Rappels automatiques / emails / mises à jour de statut sans Make"),
        ("CRM freelance","Prospects / clients / missions / livrables / facturation / suivi — tout en un"),
        "Un CRM bien configuré vous permet de gérer 10 clients simultanément sans perdre de fil",
    ],
    theory_insight="Votre temps est votre ressource la plus précieuse — un CRM IA vous libère 5h/semaine de gestion administrative.",
    exercise_title="Construire le CRM Freelance",
    exercise_steps=[
        ("🗄️","Table Clients","Nom / secteur / contact / statut / CA généré / NPS — vue Kanban par statut",NAVY),
        ("🗄️","Table Missions","Client lié / type / dates / livrables / statut / montant / facturé",RED),
        ("🗄️","Table Modules","Titre / type / durée / public / prix unitaire / utilisations",TEAL),
        ("🔗","Relations et vues","Lier les 3 tables / créer vue calendrier missions / vue Kanban prospects",GREEN),
    ],
    production_title="CRM Formation Freelance Opérationnel",
    production_items=[
        ("3 tables configurées","Clients + Missions + Modules — toutes reliées et peuplées avec données fictives"),
        ("5 vues utiles","Kanban prospects / calendrier missions / liste factures / galerie modules / timeline"),
        ("Automatisations","Rappel 3 jours avant mission / email confirmation livraison / alerte facture impayée"),
        "Remplissez le CRM avec vos vraies données — votre cockpit freelance est opérationnel",
    ],
    livrable="CRM Formation Freelance — Airtable Opérationnel",
    livrable_desc="Base de données complète + 5 vues + automatisations — gérez 10 clients sans effort.",
    tomorrow_title="Suivi Post-Formation Automatisé",
    tomorrow_desc="Configurez une séquence de 5 messages post-formation automatisés sur 30 jours pour ancrer les acquis."
))

DAYS.append(dict(
    day_num=40,week=7,phase=3,phase_name="Qualité & ISO",day_name="Jeudi",
    title="Suivi Post-Formation Automatisé",subtitle="Séquence 30 jours · Ancrage mémoriel · Relances intelligentes",
    accent_color=RED,
    theory_title="Le Suivi Post-Formation — La Clé de l'Ancrage",
    theory_points=[
        ("Courbe d'oubli Ebbinghaus","Sans rappel : 50% oublié en 1h / 70% en 24h / 90% en 1 semaine"),
        ("Répétition espacée","Rappels à J+1 / J+7 / J+14 / J+21 / J+30 — ancrage mémoriel prouvé"),
        ("Contenu des rappels","Mini-quiz / cas pratique / conseil du jour / ressource complémentaire / défi terrain"),
        ("IA personnalisation","Rappel adapté au score de l'apprenant — ceux qui ont moins bien réussi reçoivent plus"),
        ("Make automation","Séquence configurable une fois → tourne automatiquement pour tous les apprenants"),
        "Les formations avec suivi post J+30 ont un taux d'application terrain 3x supérieur",
    ],
    theory_insight="La formation ne finit pas le dernier jour — elle commence. Le suivi est votre valeur ajoutée invisible mais décisive.",
    exercise_title="Créer la Séquence de 5 Messages",
    exercise_steps=[
        ("✍️","Message J+1","Récapitulatif des 3 points clés + 1 défi à réaliser aujourd'hui — ton motivant",RED),
        ("✍️","Message J+7","Mini-quiz 3 questions sur les concepts clés + correction commentée",ORANGE),
        ("✍️","Messages J+14/21/30","Cas pratique terrain / ressource complémentaire / bilan personnel",TEAL),
        ("🔗","Automatisation Make","Séquence programmée — déclenchée automatiquement à la fin de chaque formation",GREEN),
    ],
    production_title="Programme Suivi Post-Formation Automatisé",
    production_items=[
        ("5 messages rédigés","J+1 / J+7 / J+14 / J+21 / J+30 — chacun avec objectif pédagogique clair"),
        ("Workflow Make","Déclencheur fin formation → séquence automatique de 5 emails sur 30 jours"),
        ("Rapport de complétion","Dashboard : taux d'ouverture / quiz scores / engagement par message"),
        "Ce suivi différencie radicalement votre offre des formations classiques — argument ROI client",
    ],
    livrable="Programme Suivi Post-Formation — 5 Messages sur 30 Jours",
    livrable_desc="Séquence rédigée + workflow Make + dashboard suivi — valeur ajoutée mesurable incluse dans vos formations.",
    tomorrow_title="Intégration No-Code Globale",
    tomorrow_desc="Connectez tous vos outils en un cockpit freelance cohérent — de la prospection à la facturation."
))

DAYS.append(dict(
    day_num=41,week=7,phase=3,phase_name="Qualité & ISO",day_name="Vendredi",
    title="Intégration No-Code — Cockpit Freelance IA",subtitle="Connecter tous les outils · Flux bout en bout",
    accent_color=GOLD,
    theory_title="L'Architecture No-Code Freelance",
    theory_points=[
        ("Vision d'ensemble","Typeform → Make → Airtable → Claude → Gmail → Notion — un flux continu"),
        ("Prospection","LinkedIn contact → Airtable prospect → email automatique → relance si pas de réponse"),
        ("Mission","Signature → onboarding client → livrables → évaluations → facture → suivi"),
        ("Contenu","Formation → transcription Otter → Claude synthèse → post LinkedIn → newsletter"),
        ("Comptabilité","Mission terminée → Airtable → facture générée → relance impayé automatique"),
        "Un système bien intégré = 8h/semaine économisées = 1 journée de plus pour votre activité principale",
    ],
    theory_insight="Vous n'avez pas à tout automatiser d'un coup — commencez par les 3 tâches qui vous coûtent le plus de temps.",
    exercise_title="Mapper et Connecter les Flux Prioritaires",
    exercise_steps=[
        ("🗺️","Mapper vos 3 flux prioritaires","Prospection / livraison formation / facturation — schéma de chaque flux",GOLD),
        ("🔗","Connecter le flux formation","Typeform éval → Make → Airtable mission → email client → Notion livrable",ORANGE),
        ("🔗","Connecter le flux prospection","Contact LinkedIn → Airtable prospect → séquence email Make → relance",TEAL),
        ("🧪","Tester bout en bout","Simulez une vraie mission de A à Z — détectez les points de friction",GREEN),
    ],
    production_title="Cockpit Freelance IA Opérationnel",
    production_items=[
        ("Schéma d'architecture","Visio ou draw.io : tous les outils / flux / connexions — vue d'ensemble"),
        ("2 flux automatisés","Formation complète + Prospection — testés et fonctionnels"),
        ("Guide d'utilisation","Procédure pour chaque situation : nouvelle mission / relance / clôture"),
        "Ce cockpit est un service vendable aux autres formateurs freelance — modèle 'clé en main'",
    ],
    livrable="Cockpit Freelance IA — Architecture No-Code Complète",
    livrable_desc="Schéma + 2 flux automatisés + guide utilisation — gestion d'activité freelance optimisée.",
    tomorrow_title="Stack Technologique et Contenu LinkedIn",
    tomorrow_desc="Documentez votre stack IA et transformez-le en contenu LinkedIn à fort engagement."
))

DAYS.append(dict(
    day_num=42,week=7,phase=3,phase_name="Qualité & ISO",day_name="Samedi",
    title="Stack Tech et Premier Contenu LinkedIn IA",subtitle="Documenter · Partager · Construire l'audience",
    accent_color=TEAL,
    theory_title="LinkedIn pour Expert Freelance IA",
    theory_points=[
        ("Pourquoi LinkedIn maintenant","Vous avez 6 semaines d'expérience IA concrète — c'est déjà expert pour 95% du marché"),
        ("Formats qui fonctionnent","Listes numérotées / avant-après / 'j'ai testé X voici ce que j'ai appris' / cas concret"),
        ("Algorithme LinkedIn","Commentaires > Partages > Likes — posez des questions à la fin de chaque post"),
        ("Régularité > Perfection","2 posts/semaine réguliers valent mieux que 1 chef d'œuvre par mois"),
        ("Votre niche rare","IA × Formation × Centres d'Appels × ISO 9001 — pratiquement personne ne fait ça"),
        "Un profil LinkedIn actif génère 3x plus de prospects entrants qu'un profil statique",
    ],
    theory_insight="Chaque livrable que vous avez créé ces 6 semaines est un post LinkedIn en puissance.",
    exercise_title="Créer votre Premier Contenu LinkedIn IA",
    exercise_steps=[
        ("✍️","Post 1 — Stack tech","'Ma stack IA de formatrice freelance en 2025 : 8 outils / 10x ma productivité' — liste + photo",TEAL),
        ("✍️","Post 2 — Avant/après","Un livrable que vous faisiez en 3h / maintenant en 20 min avec l'IA — concret",GREEN),
        ("✍️","Post 3 — Tip ISO × IA","'Comment j'utilise NotebookLM pour former les managers à l'ISO 9001 en 1h'",ORANGE),
        ("📅","Calendrier 30 jours","Planifiez 8 posts pour les 4 prochaines semaines — thèmes + dates",NAVY),
    ],
    production_title="3 Posts LinkedIn + Calendrier Éditorial",
    production_items=[
        ("3 posts rédigés","Prêts à publier — relu et validé par Claude pour ton et impact"),
        ("Calendrier 30 jours","8 posts planifiés sur 4 semaines — variété des thèmes"),
        ("Optimiser le profil","Titre LinkedIn mis à jour : 'Expert(e) Formation & Qualité CA × IA | Freelance'"),
        "Publiez votre premier post aujourd'hui — chaque heure de délai est une heure sans visibilité",
    ],
    livrable="3 Posts LinkedIn + Calendrier Éditorial 30 Jours",
    livrable_desc="Contenu prêt à publier + stratégie éditoriale — début de votre personal branding IA.",
    tomorrow_title="Lean × IA — Identifier les Gaspillages",
    tomorrow_desc="Semaine 8 : analysez un processus de formation avec l'IA pour identifier les 7 gaspillages Lean."
))

DAYS.append(dict(
    day_num=43,week=8,phase=3,phase_name="Qualité & ISO",day_name="Lundi",
    title="Lean × IA — Identifier les Gaspillages",subtitle="7 Muda · Diagnostic Lean · Plan d'élimination",
    accent_color=RED,
    theory_title="Lean Management Augmenté par l'IA",
    theory_points=[
        ("7 gaspillages Muda","Surproduction / Attentes / Transport / Stocks / Mouvements / Défauts / Sur-traitement"),
        ("Lean en formation","Gaspillages typiques : contenu inutile / attentes administratives / reprises / sur-documentation"),
        ("Lean en CA","Escalades évitables / double traitement / temps de mise en attente / formations redondantes"),
        ("IA et diagnostic Lean","Analyser un processus décrit → identifier les gaspillages probables → prioriser"),
        ("Priorité ROI","Éliminer d'abord les gaspillages avec le plus fort impact sur satisfaction client et coût"),
        "Une démarche Lean bien conduite réduit le temps de traitement de 25% sans investissement",
    ],
    theory_insight="Votre œil terrain détecte les gaspillages réels — l'IA les catégorise et les priorise automatiquement.",
    exercise_title="Diagnostic Lean sur 2 Processus",
    exercise_steps=[
        ("🔍","Processus 1 — Formation initiale","Décrivez le processus actuel → Claude identifie les 7 gaspillages potentiels",RED),
        ("🔍","Processus 2 — Traitement email","Back-office email → analyse Lean → gaspillages + quick wins",ORANGE),
        ("📊","Priorisation","Matrice impact/effort : quel gaspillage éliminer en premier pour le meilleur ROI",TEAL),
        ("💡","Plan d'élimination","5 actions concrètes avec délai et indicateur de résultat",GREEN),
    ],
    production_title="Diagnostic Lean Formation CA",
    production_items=[
        ("Rapport diagnostic","Les 7 gaspillages identifiés / niveau d'impact / fréquence / coût estimé"),
        ("Matrice priorisation","Visuel Canva : impact vs effort — 4 quadrants / actions par quadrant"),
        ("Plan d'actions Lean","5 quick wins / 3 chantiers moyen terme / 2 projets long terme"),
        "Le diagnostic Lean est une porte d'entrée commerciale puissante — offrez-le en 'audit flash' gratuit",
    ],
    livrable="Diagnostic Lean Formation CA — Rapport + Plan d'Actions",
    livrable_desc="Analyse gaspillages + matrice priorisation + plan d'actions — offre audit flash commercialisable.",
    tomorrow_title="Benchmark IA — Rapport Sectoriel",
    tomorrow_desc="Produisez un rapport de benchmark 'Qualité Formation Centres d'Appels France 2025' avec Perplexity."
))

DAYS.append(dict(
    day_num=44,week=8,phase=3,phase_name="Qualité & ISO",day_name="Mardi",
    title="Benchmark IA — Rapport Sectoriel 2025",subtitle="Meilleures pratiques · Écarts · Recommandations",
    accent_color=ORANGE,
    theory_title="Le Benchmark Assisté par l'IA",
    theory_points=[
        ("Méthode benchmark","Identifier les leaders / collecter données / comparer à votre cible / analyser écarts / recommander"),
        ("Sources fiables","Perplexity / rapports sectoriels / études AFNOR / benchmarks EY-KPMG / articles spécialisés"),
        ("Analyse d'écart","Situation actuelle vs best practice → gap → priorité de combler le gap"),
        ("Rapport structuré","Contexte / méthodologie / résultats / analyse / recommandations / conclusion"),
        ("Valeur commerciale","Un benchmark bien documenté justifie à lui seul une prestation de 1000 à 3000€"),
        "Rares sont les consultants qui produisent des benchmarks sectoriels documentés — c'est votre différenciateur",
    ],
    theory_insight="Ce rapport positionne comme analyste de marché — au-delà du formateur, vous devenez experte-conseil.",
    exercise_title="Rechercher et Structurer le Benchmark",
    exercise_steps=[
        ("🔍","Recherches Perplexity","10 requêtes ciblées : pratiques QA / outils IA / indicateurs performance / tendances RH CA FR",ORANGE),
        ("📊","Données comparatives","Collectez : KPIs moyens secteur / taux adoption IA / budgets formation / certifications",RED),
        ("🔍","Analyse écarts","Pour chaque dimension : où en est le marché / où devraient être vos clients / comment y arriver",TEAL),
        ("📝","Rédiger avec Claude","Structure + données → Claude rédige les sections → vous validez et enrichissez",GREEN),
    ],
    production_title="Rapport Benchmark QF CA France 2025",
    production_items=[
        ("Rapport 10-15 pages","Contexte / méthodologie / 5 dimensions benchmarkées / recommandations / conclusion"),
        ("Infographie résumé","Canva : les 5 chiffres clés du benchmark — partageable sur LinkedIn"),
        ("Version client","Même contenu adapté en présentation 15 slides Gamma pour réunion direction"),
        "Publiez l'infographie sur LinkedIn — c'est votre meilleur contenu de génération de leads",
    ],
    livrable="Rapport Benchmark Qualité Formation CA France 2025",
    livrable_desc="Rapport 10-15 pages + infographie + présentation — contenu d'autorité à forte valeur commerciale.",
    tomorrow_title="Storytelling Qualité et Data",
    tomorrow_desc="Transformez un rapport technique en présentation narrative engageante — data storytelling avec l'IA."
))

DAYS.append(dict(
    day_num=45,week=8,phase=3,phase_name="Qualité & ISO",day_name="Mercredi",
    title="Storytelling Qualité — Données → Histoire",subtitle="Data storytelling · Présentation narrative · Impact",
    accent_color=NAVY,
    theory_title="Le Data Storytelling avec l'IA",
    theory_points=[
        ("Problème du rapport classique","Un rapport technique n'est lu que par 20% des destinataires — les décideurs ne lisent pas"),
        ("Solution : storytelling","Même données + narration + visuels = 5x plus d'impact et de mémorisation"),
        ("Structure narrative","Situation (contexte) → Complication (problème) → Résolution (votre solution) → Résultats"),
        ("IA et narration","Claude transforme n'importe quel rapport en histoire cohérente et engageante"),
        ("Visuels impactants","Un seul chiffre bien mis en forme vaut mieux qu'un tableau de 50 lignes"),
        "Un dirigeant retient 65% d'une histoire bien racontée vs 5% d'un rapport standard",
    ],
    theory_insight="Vous savez analyser — maintenant apprenez à convaincre. Le storytelling est la clé de la vente de vos analyses.",
    exercise_title="Transformer le Rapport Benchmark en Présentation",
    exercise_steps=[
        ("📖","Identifier les 3 messages clés","Du rapport Jour 44 : quels sont les 3 insights les plus impactants ?",NAVY),
        ("✍️","Écrire la narrative","Claude : transforme ce rapport en histoire — problème secteur / solutions / appel à l'action",RED),
        ("🎨","Gamma — 25 slides","Narrative + visuels + données clés — présentation engageante et mémorable",TEAL),
        ("🎙️","Préparer le discours","Pour chaque slide : que dire en 30 secondes ? Entrainement avec Claude comme public",GREEN),
    ],
    production_title="Présentation Storytelling Qualité CA 2025",
    production_items=[
        ("Présentation 25 slides","Narrative / données / visuels / call to action — format conférence ou webinaire"),
        ("Script de présentation","Word pour mot ou bullet points — 20 minutes de présentation chrono"),
        ("Version condensée","5 slides pour une réunion de 10 minutes chez un client"),
        "Cette présentation peut être utilisée en webinaire LinkedIn — générateur de leads puissant",
    ],
    livrable="Présentation Storytelling 'Qualité CA 2025' — Conférence/Webinaire",
    livrable_desc="25 slides narratives + script + version condensée — contenu de conférence prêt à délivrer.",
    tomorrow_title="Système de Veille IA Automatisé",
    tomorrow_desc="Configurez votre veille IA permanente sur 5 thèmes — newsletter automatique chaque lundi matin."
))

DAYS.append(dict(
    day_num=46,week=8,phase=3,phase_name="Qualité & ISO",day_name="Jeudi",
    title="Système de Veille IA Automatisé",subtitle="5 thèmes · Curation · Newsletter hebdomadaire",
    accent_color=TEAL,
    theory_title="La Veille Continue avec l'IA",
    theory_points=[
        ("Pourquoi veiller","L'IA évolue toutes les semaines — rester à jour = rester pertinente = maintenir la valeur"),
        ("5 thèmes de veille","IA générative / Centres d'appels / Formation professionnelle / Qualité-ISO / Freelance-conseil"),
        ("Outils de veille","Perplexity quotidien / Google Alertes / LinkedIn flux / newsletters spécialisées"),
        ("Synthèse automatique","Chaque lundi : coller les news de la semaine → Claude synthèse + recommandations"),
        ("Partager la veille","Votre newsletter veille = service à valeur ajoutée pour vos clients et réseau"),
        "5 minutes de veille quotidienne > 2 heures de recherche hebdomadaire. La régularité prime.",
    ],
    theory_insight="Votre veille devient un produit : une newsletter mensuelle positionnée expert se monétise 20 à 99€/abonné.",
    exercise_title="Configurer le Système de Veille",
    exercise_steps=[
        ("⚙️","Google Alertes","5 alertes configurées : 'IA centres appels' / 'formation IA France' / 'ISO 9001 IA' / etc.",TEAL),
        ("📱","LinkedIn flux","Suivre 20 comptes experts : formateurs IA / consultants CA / AFNOR / éditeurs outils",GREEN),
        ("📋","Template synthèse Notion","Page hebdomadaire : sources lues / insights clés / opportunités / posts LinkedIn potentiels",ORANGE),
        ("🔄","Workflow Make","Chaque lundi 8h : email auto avec liens à lire + espace de synthèse pré-rempli",NAVY),
    ],
    production_title="Tableau de Bord Veille IA Opérationnel",
    production_items=[
        ("5 alertes Google actives","Configurées et testées — premières alertes reçues"),
        ("Template synthèse Notion","Page hebdomadaire prête à remplir — champs structurés"),
        ("Workflow Make","Rappel hebdomadaire automatique + compilation des sources"),
        "Dans 4 semaines vous aurez assez de matière pour lancer une newsletter mensuelle",
    ],
    livrable="Système de Veille IA — 5 Thèmes Automatisés",
    livrable_desc="Alertes + flux LinkedIn + template Notion + workflow Make — veille permanente sans effort.",
    tomorrow_title="Révision et Consolidation Phase 3",
    tomorrow_desc="Révisez les points faibles, améliorez vos 2 meilleurs livrables et préparez l'évaluation Phase 3."
))

DAYS.append(dict(
    day_num=47,week=8,phase=3,phase_name="Qualité & ISO",day_name="Vendredi",
    title="Révision et Consolidation Phase 3",subtitle="Identifier les lacunes · Améliorer · Préparer l'éval",
    accent_color=RED,
    theory_title="La Révision Active — Plus Efficace que la Relecture",
    theory_points=[
        ("Révision active vs passive","Relire = passif / Reproduire sans regarder = actif → ancrage 3x supérieur"),
        ("Technique du rappel","Fermez vos notes. Qu'avez-vous retenu de la Phase 3 ? Listez sans regarder. Puis vérifiez."),
        ("Identifier les lacunes","Comparez votre liste avec les livrables Phase 3 — les oublis sont vos points faibles"),
        ("Améliorer, pas rajouter","Mieux vaut 3 livrables excellents que 10 moyens — choisissez vos meilleurs et perfectionnez"),
        ("Préparer l'évaluation","Relisez la grille de niveau 3 — soyez honnête sur ce que vous maîtrisez vraiment"),
        "30 minutes de révision active valent plus que 3 heures de lecture passive",
    ],
    theory_insight="Les lacunes identifiées aujourd'hui sont des opportunités de croissance — pas des échecs.",
    exercise_title="Révision Active Phase 3",
    exercise_steps=[
        ("📝","Rappel sans notes","Listez tout ce que vous avez appris en Phase 3 sans regarder — 15 minutes",RED),
        ("🔍","Identifier les lacunes","Comparez avec le programme — qu'avez-vous oublié ou peu maîtrisé ?",ORANGE),
        ("💪","Révision ciblée","Pour chaque lacune : relisez le livrable / refaites l'exercice clé / posez une question à Claude",TEAL),
        ("🏆","Améliorer 2 livrables","Choisissez vos 2 meilleurs livrables Phase 3 et demandez à Claude comment les rendre excellents",GREEN),
    ],
    production_title="2 Livrables Phase 3 Améliorés",
    production_items=[
        ("Livrable 1 amélioré","Le meilleur de la semaine — version 2.0 avec les retours Claude"),
        ("Livrable 2 amélioré","Second choix — version renforcée"),
        ("Plan de révision personnalisé","Claude génère un plan de révision sur vos 3 points faibles — pour demain"),
        "Qualité > quantité. Deux livrables parfaits ouvrent plus de portes que dix livrables passables.",
    ],
    livrable="2 Livrables Phase 3 Version 2.0 — Qualité Renforcée",
    livrable_desc="Versions améliorées de vos meilleurs livrables + plan de révision personnalisé.",
    tomorrow_title="Évaluation Phase 3 — Niveau 3 Intégration",
    tomorrow_desc="Quiz complet Phase 3 + portfolio 37 livrables + vérification de votre niveau 3."
))

DAYS.append(dict(
    day_num=48,week=8,phase=3,phase_name="Qualité & ISO",day_name="Samedi",
    title="Évaluation Phase 3 — Niveau 3 Validé",subtitle="Quiz · Portfolio · Passage Phase 4",
    accent_color=GREEN,
    theory_title="Grille Évaluation — Niveau 3 Intégration",
    theory_points=[
        ("Critère 1 /20","Automatiser au moins 2 processus de son activité — démonstration live"),
        ("Critère 2 /20","3 offres freelance packagées avec prix et livrables décrits"),
        ("Critère 3 /20","Former quelqu'un d'autre à l'utilisation de l'IA — simulé avec Claude"),
        ("Critère 4 /20","Portfolio 37 livrables documentés et valorisés"),
        ("Critère 5 /20","Contenu LinkedIn publié ou prêt à publier — 3 posts minimum"),
        "Score ≥ 60/100 = Niveau 3 validé. Entrée en Phase 4 : Freelance IA.",
    ],
    theory_insight="Phase 3 terminée — vous êtes maintenant l'une des rares expertes ISO × Formation × IA en France.",
    exercise_title="Quiz Phase 3 — 20 Questions",
    exercise_steps=[
        ("📝","Questions 1-5","SMQ IA / cartographie processus / audit interne IA / 8D IA / indicateurs prédictifs",GREEN),
        ("📝","Questions 6-10","PDCA IA / Make workflows / Typeform / Airtable CRM / suivi post-formation",TEAL),
        ("📝","Questions 11-15","Lean × IA / benchmark / storytelling / veille IA / LinkedIn stratégie",ORANGE),
        ("📝","Questions 16-20","Décrire votre architecture no-code / vos 3 offres / votre PVU / votre valeur différenciante",NAVY),
    ],
    production_title="Portfolio Phase 3 — Consolidation",
    production_items=[
        ("37 livrables documentés","Chaque livrable : valeur commerciale estimée + outil principal + temps de production"),
        ("3 offres finalisées","Offre 1 Audit Qualité / Offre 2 SMQ / Offre 3 Formation — prêtes à envoyer"),
        ("Score et projection","Votre score Phase 3 / écart à 60 / plan si besoin de révision"),
        "37 livrables professionnels en 8 semaines — c'est une performance remarquable. Continuez.",
    ],
    livrable="Portfolio Phase 3 — 37 Livrables + 3 Offres Packagées",
    livrable_desc="Phase 3 complétée. Score ≥ 60 = Niveau 3 certifié. Entrée en Phase 4 : Freelance IA.",
    tomorrow_title="Phase 4 — Développer votre Activité Freelance IA",
    tomorrow_desc="PVU / tarifs / GPT Custom / personal branding — vous allez transformer votre expertise en revenus."
))

# ══════════════════════════════════════
# PHASE 4 — FREELANCE IA (Semaines 9-11)
# ══════════════════════════════════════

DAYS.append(dict(
    day_num=49,week=9,phase=4,phase_name="Freelance IA",day_name="Lundi",
    title="Votre Proposition de Valeur Unique",subtitle="PVU · Positionnement · Niche · Message percutant",
    accent_color=RED,
    theory_title="Construire une PVU Irrésistible",
    theory_points=[
        ("PVU = Proposition de Valeur Unique","La phrase qui explique en 10 secondes pourquoi vous et pas un autre"),
        ("Structure PVU","Je aide [qui] à [résoudre quoi] grâce à [comment] pour [résultat mesurable]"),
        ("Votre niche rare","Formation + Qualité + Centres d'Appels + ISO 9001 + IA = quasi monopole en France"),
        ("Test PVU","La tester sur 3 personnes de votre cible — leur réaction en 5 secondes dit tout"),
        ("Déclinaisons","LinkedIn titre / email signature / présentation orale / page site / devis"),
        "Une PVU claire attire les bons clients et repousse les mauvais — économisez votre énergie",
    ],
    theory_insight="Votre niche est si précise que la concurrence est quasi inexistante — c'est votre force principale.",
    exercise_title="Définir et Tester votre PVU",
    exercise_steps=[
        ("🎯","Brainstorm PVU","Demandez à Claude 5 versions de PVU basées sur votre profil — choisissez la plus percutante",RED),
        ("✍️","Affiner","Prenez la meilleure — testez-la avec Claude jouant 3 types de clients différents",ORANGE),
        ("🔄","Décliner","5 versions adaptées à chaque support : LinkedIn / email / oral / site / devis",TEAL),
        ("✅","Finaliser","La PVU validée devient votre signature sur tous vos supports — mise en cohérence",GREEN),
    ],
    production_title="PVU Finalisée et Déclinée sur 5 Supports",
    production_items=[
        ("PVU principale","1 phrase / maximum 20 mots / claire / différenciante / mémorable"),
        ("5 déclinaisons","LinkedIn / email / oral 30s / site web / devis — chacune adaptée au contexte"),
        ("Message de positionnement","1 paragraphe de 5 lignes développant la PVU — pour page About site"),
        "Votre PVU est maintenant le fil directeur de tout votre marketing — appliquez-la partout",
    ],
    livrable="PVU Finalisée + 5 Déclinaisons Support",
    livrable_desc="Proposition de valeur unique validée et adaptée à chaque canal de communication.",
    tomorrow_title="Grille Tarifaire et Packages",
    tomorrow_desc="Définissez votre TJM, construisez 4 offres packagées avec prix et créez votre politique tarifaire."
))

DAYS.append(dict(
    day_num=50,week=9,phase=4,phase_name="Freelance IA",day_name="Mardi",
    title="Grille Tarifaire et Packages Freelance",subtitle="TJM · 4 offres packagées · Politique tarifaire",
    accent_color=GOLD,
    theory_title="Tarifier l'Expertise Augmentée par l'IA",
    theory_points=[
        ("TJM consultant formation","Marché France : 400-800€/jour junior / 800-1500€/jour expert / 1500€+ spécialiste rare"),
        ("Votre positionnement","Expert 15 ans + IA + ISO + CA = haut de gamme justifié — ne bradez pas"),
        ("Tarification par valeur","Prix basé sur économie générée pour le client — pas sur votre temps"),
        ("4 modèles de revenus","TJM pour missions longues / forfait livrable / retainer mensuel / produit digital"),
        ("Psychologie des prix","3 offres : Essentiel / Pro / Premium — le client compare et choisit souvent le milieu"),
        "Augmenter son TJM de 20% ne perd que 10% des clients — le revenu net augmente",
    ],
    theory_insight="L'IA multiplie votre productivité — votre tarif doit refléter la valeur livrée, pas les heures passées.",
    exercise_title="Construire la Grille Tarifaire",
    exercise_steps=[
        ("💰","Calculer votre TJM cible","Charges + salaire souhaité + frais + marge = TJM minimum → TJM marché",GOLD),
        ("📦","4 offres packagées","Audit flash (1 jour) / Mission audit (5 jours) / Formation (2 jours) / Accompagnement (3 mois)",ORANGE),
        ("🏷️","Tarifer chaque offre","Prix par valeur + comparaison marché + marge confort — justifier chaque prix",TEAL),
        ("📊","Tableau comparatif","3 niveaux Essentiel/Pro/Premium avec ce qui est inclus à chaque niveau",NAVY),
    ],
    production_title="Politique Tarifaire Complète",
    production_items=[
        ("TJM documenté","Calcul détaillé + benchmark marché + positionnement"),
        ("4 fiches offres avec prix","Chaque offre : contenu / durée / livrables / prix / conditions"),
        ("Tableau 3 niveaux","Essentiel vs Pro vs Premium — pour présenter aux clients hésitants"),
        "Votre grille tarifaire est confidentielle mais doit être disponible immédiatement — préparez-la",
    ],
    livrable="Grille Tarifaire Complète — 4 Offres + TJM",
    livrable_desc="Politique tarifaire documentée + 4 fiches offres avec prix — base commerciale solide.",
    tomorrow_title="GPT Custom — Vos Outils IA Propriétaires",
    tomorrow_desc="Créez 3 GPT Custom spécialisés : QualityCoach IA / Formation Expert / ISO Assistant."
))

DAYS.append(dict(
    day_num=51,week=9,phase=4,phase_name="Freelance IA",day_name="Mercredi",
    title="GPT Custom — Vos Outils IA Propriétaires",subtitle="QualityCoach · Formation Expert · ISO Assistant",
    accent_color=TEAL,
    theory_title="Les GPT Custom — Votre Marque IA",
    theory_points=[
        ("Qu'est-ce qu'un GPT Custom","Assistant Claude ou ChatGPT configuré avec vos instructions permanentes et vos documents"),
        ("Différenciateur commercial","Vous ne vendez plus des heures — vous vendez l'accès à des outils IA experts"),
        ("3 GPT à créer","QualityCoach IA (qualité CA) / Formation Expert (ingénierie pédagogique) / ISO Helper (norme vulgarisée)"),
        ("Instructions système","La 'fiche de poste' de votre GPT — plus elle est précise, plus l'outil est expert"),
        ("Documents de base","Chargez vos chartes, grilles, procédures — le GPT les connaît par cœur"),
        "Un GPT Custom bien configuré est un produit vendable — accès mensuel 20-50€ par utilisateur",
    ],
    theory_insight="Vos GPT Custom sont votre propriété intellectuelle IA — ils ne peuvent être copiés sans votre expertise.",
    exercise_title="Configurer les 3 GPT Custom",
    exercise_steps=[
        ("🤖","GPT 1 — QualityCoach IA","Instructions : expert CA / ISO 9001 / analyse appels / feedback structuré / plans d'action",TEAL),
        ("🎓","GPT 2 — Formation Expert","Instructions : ingénierie pédagogique / modules CA / objectifs SMART / évaluations Kirkpatrick",GREEN),
        ("📋","GPT 3 — ISO Helper","Instructions : vulgariser ISO 9001 / répondre aux questions non-experts / exemples CA concrets",ORANGE),
        ("🧪","Tester les 3","10 questions par GPT — sont-ils vraiment experts dans leur domaine ?",NAVY),
    ],
    production_title="3 GPT Custom Documentés",
    production_items=[
        ("3 GPT opérationnels","Testés et affinés — répondent de façon experte dans leur domaine"),
        ("Fiches de description","Pour chaque GPT : ce qu'il fait / ce qu'il ne fait pas / comment l'utiliser"),
        ("Cas d'usage clients","5 exemples concrets d'utilisation par un responsable CA / formateur / manager"),
        "Présentez vos GPT en démo à vos prospects — 15 minutes de démonstration = conviction forte",
    ],
    livrable="3 GPT Custom Opérationnels + Documentation",
    livrable_desc="QualityCoach IA + Formation Expert + ISO Helper — vos outils propriétaires différenciants.",
    tomorrow_title="Bibliothèque de Prompts Premium",
    tomorrow_desc="Organisez vos 50 meilleurs prompts en 5 packs thématiques vendables — votre premier produit digital."
))

DAYS.append(dict(
    day_num=52,week=9,phase=4,phase_name="Freelance IA",day_name="Jeudi",
    title="Bibliothèque de Prompts Premium",subtitle="50 prompts · 5 packs · Produit digital vendable",
    accent_color=ORANGE,
    theory_title="Les Prompt Packs — Produit Digital Scalable",
    theory_points=[
        ("Produit digital","Vendu une fois / livré infiniment / zéro coût marginal — idéal pour scaler"),
        ("Valeur perçue","Un prompt bien rédigé économise 2h de travail — facile à valoriser à 5-15€ le prompt"),
        ("5 packs thématiques","QA Appels / Formation / ISO Qualité / Management / Recrutement — 10 prompts chacun"),
        ("Format de vente","PDF + fichier texte + guide d'utilisation / vendu sur Gumroad / Payhip / votre site"),
        ("Pricing","Pack simple : 29€ / Collection 5 packs : 99€ / Mise à jour annuelle : +20€"),
        "Un produit digital généré pendant 3 mois peut générer des revenus passifs pendant 3 ans",
    ],
    theory_insight="Vos prompts représentent 50+ heures de travail condensé — leur valeur est bien supérieure à leur prix de vente.",
    exercise_title="Trier et Organiser les 50 Meilleurs Prompts",
    exercise_steps=[
        ("📋","Inventaire","Parcourez vos 50+ prompts créés depuis Jour 2 — notez chaque prompt /5",ORANGE),
        ("🗂️","Organiser en 5 packs","10 prompts par thème — choisissez les mieux notés pour chaque pack",RED),
        ("✍️","Rédiger les guides","Pour chaque pack : introduction / comment utiliser / cas d'usage / résultat attendu",TEAL),
        ("🎨","Mise en page PDF","Canva : PDF professionnel pour chaque pack — design cohérent avec votre marque",GREEN),
    ],
    production_title="5 Prompt Packs PDF Prêts à Vendre",
    production_items=[
        ("5 PDFs finalisés","Chaque pack : couverture / introduction / 10 prompts commentés / guide utilisation"),
        ("Page de vente","Description de chaque pack / bénéfices / exemples de résultats / prix / achat"),
        ("Bundle 5 packs","Offre groupée à prix réduit — augmente le panier moyen"),
        "Publiez un pack sur Gumroad cette semaine — vos premiers revenus passifs",
    ],
    livrable="5 Prompt Packs PDF — Produits Digitaux Prêts à Vendre",
    livrable_desc="50 prompts organisés en 5 packs thématiques avec guides — revenus passifs immédiats.",
    tomorrow_title="Personal Branding LinkedIn — Calendrier 30 Jours",
    tomorrow_desc="Optimisez votre profil LinkedIn et créez 20 posts pré-rédigés pour un mois de visibilité IA."
))

DAYS.append(dict(
    day_num=53,week=9,phase=4,phase_name="Freelance IA",day_name="Vendredi",
    title="Personal Branding LinkedIn — 20 Posts Prêts",subtitle="Profil optimisé · Calendrier 30 jours · Audience expert",
    accent_color=NAVY,
    theory_title="LinkedIn pour Expert Freelance — Stratégie Avancée",
    theory_points=[
        ("Profil = vitrine","Titre / résumé / expériences reformulés avec IA — 3x plus de vues en 30 jours"),
        ("20 formats efficaces","Astuce / liste / coulisses / avant-après / cas client / poll / carrousel / vidéo"),
        ("Algorithme 2025","Commentaires > Partages > Likes / premières 60 min critiques / réseau proche prime"),
        ("Votre angle unique","IA + Formation + CA + ISO = 0 concurrent direct sur ce positionnement"),
        ("Régularité","3 posts/semaine pendant 3 mois = audience de 500+ followers qualifiés"),
        "Un follower qualifié sur LinkedIn vaut 10x un abonné généraliste — construisez petit mais solide",
    ],
    theory_insight="Votre expertise de 15 ans + 2 mois d'IA pratique vous donne plus de légitimité que 95% des 'experts IA' LinkedIn.",
    exercise_title="Optimiser Profil + Créer 20 Posts",
    exercise_steps=[
        ("👤","Optimiser le profil","Titre / résumé / expériences / compétences — tout mis à jour avec votre PVU",NAVY),
        ("✍️","10 posts semaines 10-11","5 posts expertise CA / 3 posts IA pratique / 2 posts ISO — rédigés avec Claude",RED),
        ("✍️","10 posts semaines 12-13","5 posts témoignages outils / 3 posts tips / 2 posts coulisses formation — planifiés",TEAL),
        ("📅","Calendrier éditorial","Date / thème / format / objectif / hashtags — 20 lignes dans Notion",GREEN),
    ],
    production_title="Profil LinkedIn + 20 Posts Prêts à Publier",
    production_items=[
        ("Profil optimisé","Titre accrocheur + résumé PVU + photo pro + bannière Canva"),
        ("20 posts rédigés","Variété de formats et thèmes — prêts à copier-coller avec image si besoin"),
        ("Calendrier Notion","Planning visuel des 20 prochains posts — ne jamais manquer une publication"),
        "Publiez le premier post aujourd'hui — chaque semaine sans publication est une semaine perdue",
    ],
    livrable="Profil LinkedIn Optimisé + 20 Posts Prêts",
    livrable_desc="Vitrine professionnelle + 30 jours de contenu prêt — visibilité immédiate sur votre marché cible.",
    tomorrow_title="Lead Magnet — Guide PDF Gratuit",
    tomorrow_desc="Créez un guide PDF gratuit '5 façons dont l'IA améliore votre centre d'appels' pour attirer vos clients idéaux."
))

DAYS.append(dict(
    day_num=54,week=9,phase=4,phase_name="Freelance IA",day_name="Samedi",
    title="Lead Magnet — Guide PDF Gratuit",subtitle="Attirer les clients idéaux · Démontrer l'expertise",
    accent_color=GREEN,
    theory_title="Le Lead Magnet — Attirer avant de Vendre",
    theory_points=[
        ("Définition","Contenu gratuit à forte valeur qui attire des prospects qualifiés en échange de leur email"),
        ("Pourquoi ça marche","Le client test votre expertise gratuitement → fait confiance → achète"),
        ("Format idéal","Guide PDF 8-15 pages / checklist / mini-formation / webinaire — lisible en 15 min"),
        ("Votre thème idéal","'5 façons dont l'IA améliore la qualité et la formation de votre centre d'appels'"),
        ("Distribution","LinkedIn post avec lien / site web / email de prospection / événements sectoriels"),
        "Un bon lead magnet génère 3x plus de conversations commerciales qu'un email de prospection à froid",
    ],
    theory_insight="Le lead magnet transforme votre expertise en aimant — les bons clients viennent à vous.",
    exercise_title="Créer le Guide PDF",
    exercise_steps=[
        ("📝","Structurer le guide","5 chapitres : 1 problème CA / 1 solution IA / résultats chiffrés / mise en œuvre / prochaine étape",GREEN),
        ("✍️","Rédiger avec Claude","Chapitre par chapitre avec vos exemples concrets — Claude structure, vous validez",TEAL),
        ("🎨","Mise en page Canva","Design professionnel : couverture / chapitres / call to action final",ORANGE),
        ("📢","Page de téléchargement","Notion ou Carrd : formulaire email + téléchargement automatique du PDF",NAVY),
    ],
    production_title="Guide PDF + Page de Téléchargement",
    production_items=[
        ("Guide PDF 10 pages","'5 façons dont l'IA améliore votre centre d'appels' — design Canva professionnel"),
        ("Page Notion/Carrd","Formulaire email + téléchargement automatique + message de bienvenue"),
        ("Post LinkedIn d'annonce","Annoncer le guide : problème / teaser des 5 façons / lien de téléchargement"),
        "Publiez le post LinkedIn d'annonce aujourd'hui — premiers téléchargements ce weekend",
    ],
    livrable="Lead Magnet PDF + Page de Téléchargement Opérationnelle",
    livrable_desc="Guide PDF + page web + post LinkedIn — premier outil d'acquisition de prospects automatisé.",
    tomorrow_title="Formation 'IA pour Formateurs' — Conception Jour 1",
    tomorrow_desc="Semaine 10 : concevez votre formation 2 jours 'IA pour Formateurs' — le programme, les supports, les exercices."
))

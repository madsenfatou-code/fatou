# Days 55–78
from pptx_helpers import *

DAYS = []

DAYS.append(dict(
    day_num=55,week=10,phase=4,phase_name="Freelance IA",day_name="Lundi",
    title="Formation 'IA pour Formateurs' — Jour 1",subtitle="Conception programme · Objectifs · Supports Matinée",
    accent_color=TEAL,
    theory_title="Ingénierie d'une Formation sur l'IA",
    theory_points=[
        ("Public cible","Formateurs en activité / sans expérience IA / résistance possible / adultes exigeants"),
        ("Positionnement","Pas une formation théorique — 80% pratique / livrables réels produits pendant la formation"),
        ("Objectifs J1","Comprendre l'IA générative / maîtriser le prompt engineering / produire 3 livrables"),
        ("Structure J1","Matin : théorie + premiers outils (3h) / Après-midi : ateliers pratiques guidés (4h)"),
        ("Méthodes actives","Démonstration live / ateliers en binôme / challenge production / débriefing collectif"),
        "Un formateur qui a produit quelque chose pendant la formation repart avec une conviction concrète",
    ],
    theory_insight="Vous êtes la mieux placée pour former des formateurs à l'IA — vous êtes formatrice ET praticienne IA.",
    exercise_title="Concevoir le Programme J1 et les Supports Matin",
    exercise_steps=[
        ("🗺️","Programme J1 détaillé","Minute par minute : 9h00 à 17h30 — activités / durées / matériel / animation",TEAL),
        ("🎨","Slides matin J1","Canva/Gamma : 20 slides — accueil / IA générative / prompts RCTF / outils essentiels",GREEN),
        ("📋","Livret participant","Pages matin J1 : espace notes / prompts à compléter / exercices guidés",ORANGE),
        ("🧪","Tester le programme","Claude joue un formateur résistant — testez votre pitch de la matinée",NAVY),
    ],
    production_title="Supports Formation J1 Matin",
    production_items=[
        ("Programme J1 complet","Planning heure par heure / activités / matériel / transitions"),
        ("20 slides matin","Design cohérent / exemples terrain / exercices intégrés"),
        ("Livret participant matin","10 pages / espace structuré pour chaque activité"),
        "Cette formation se vend 800-1200€ par participant en inter-entreprises",
    ],
    livrable="Formation IA pour Formateurs — Programme J1 + Supports Matin",
    livrable_desc="Programme détaillé + slides + livret participant — moitié de votre formation commercialisable.",
    tomorrow_title="Formation 'IA pour Formateurs' — Jour 1 Après-midi & Jour 2",
    tomorrow_desc="Finalisez les ateliers pratiques J1 et concevez le programme complet du Jour 2."
))

DAYS.append(dict(
    day_num=56,week=10,phase=4,phase_name="Freelance IA",day_name="Mardi",
    title="Formation 'IA pour Formateurs' — Finalisation",subtitle="Ateliers J1 · Programme J2 · Kit complet",
    accent_color=GREEN,
    theory_title="Les Ateliers Pratiques — Cœur de la Formation",
    theory_points=[
        ("Principe","1 atelier = 1 compétence + 1 livrable produit — le formateur repart avec quelque chose de concret"),
        ("3 ateliers J1 après-midi","Atelier prompts / Atelier Canva IA / Atelier Gamma — 45 min chacun"),
        ("Structure J2","Matin : outils avancés + automatisation (3h) / Après-midi : projet intégrateur (4h)"),
        ("Projet intégrateur J2","Chaque participant conçoit UN module de formation complet avec l'IA — en 3h"),
        ("Évaluation et certification","Quiz + livrable + auto-évaluation — attestation de formation remise"),
        "Un atelier production > 3 heures de théorie — les adultes apprennent en faisant",
    ],
    theory_insight="Le projet intégrateur J2 est votre argument commercial le plus fort — les participants repartent avec un livrable réel.",
    exercise_title="Créer les Ateliers et le Programme J2",
    exercise_steps=[
        ("🛠️","3 fiches atelier J1","Chaque atelier : objectif / consigne / matériel / durée / débriefing / livrable attendu",GREEN),
        ("🗺️","Programme J2","9h-17h30 minute par minute — outils avancés + projet intégrateur + remise attestation",TEAL),
        ("📋","Livret J2 + éval","Livret participant J2 + quiz 20 questions + grille auto-évaluation + attestation type",ORANGE),
        ("📦","Assembler le kit","Rassemblez tout : programme / slides / livrets / ateliers / évaluation",NAVY),
    ],
    production_title="Formation IA pour Formateurs — Kit Complet 2 Jours",
    production_items=[
        ("Guide formateur","Programme J1+J2 / fiches ateliers / guide animation / réponses aux objections"),
        ("Supports participants","Livret J1+J2 / slides imprimables / ressources complémentaires"),
        ("Évaluation complète","Quiz + grille observation + attestation type — conformité Qualiopi"),
        "Formation prête à délivrer — première session peut être proposée dès la semaine prochaine",
    ],
    livrable="Formation 'IA pour Formateurs' — Kit Complet 2 Jours",
    livrable_desc="Guide formateur + supports participants + évaluation — formation clé en main à 800-1200€/participant.",
    tomorrow_title="Atelier 'Manager IA-Ready' — Conception",
    tomorrow_desc="Concevez votre atelier 1 journée pour managers de centres d'appels — programme, supports, activités."
))

DAYS.append(dict(
    day_num=57,week=10,phase=4,phase_name="Freelance IA",day_name="Mercredi",
    title="Atelier 'Manager IA-Ready' — Conception",subtitle="Programme 7h · Managers CA · Outils IA managériaux",
    accent_color=ORANGE,
    theory_title="Former les Managers à l'IA — Enjeux Spécifiques",
    theory_points=[
        ("Public manager","Moins de résistance à la nouveauté / plus d'enjeux de légitimité / attentes ROI fortes"),
        ("Ce qu'ils veulent savoir","Comment l'IA améliore mes KPIs / comment motiver l'équipe avec l'IA / risques à gérer"),
        ("3 axes de la journée","Piloter la qualité avec l'IA / Coacher avec l'IA / Communiquer avec l'IA"),
        ("Format adapté","Plus court en théorie / plus long en mise en situation / peer learning entre managers"),
        ("Livrables de la journée","Chaque manager repart avec 1 outil IA opérationnel pour son équipe"),
        "Un manager convaincu est votre meilleur prescripteur — soignez son expérience",
    ],
    theory_insight="Les managers ont peur d'être dépassés par leurs équipes sur l'IA — votre rôle est de les rassurer et outiller.",
    exercise_title="Concevoir le Programme et les Activités",
    exercise_steps=[
        ("🗺️","Programme 7h","9h-17h : accueil / 3 modules de 1h30 / ateliers / déjeuner / projet après-midi / clôture",ORANGE),
        ("🛠️","3 ateliers managers","Tableau de bord IA / Feedback IA / Email de communication — chacun produit un outil",RED),
        ("🎨","Slides journée","Gamma : 30 slides — ton professionnel / exemples ROI / outils concrets",TEAL),
        ("📋","Kit manager","Livret + 3 outils IA à ramener + plan d'engagement 30 jours post-atelier",GREEN),
    ],
    production_title="Atelier Manager IA-Ready — Supports Complets",
    production_items=[
        ("Programme 7h détaillé","Minute par minute avec animations et transitions"),
        ("3 fiches ateliers","Objectif / matériel / consigne / débriefing pour chaque atelier"),
        ("Kit participant manager","Livret + outils IA + plan 30 jours + ressources"),
        "Tarif : 800-1500€ la journée en intra / 400-600€ par manager en inter — marché important",
    ],
    livrable="Atelier 'Manager IA-Ready' — Kit Complet 7h",
    livrable_desc="Programme + supports + 3 ateliers + kit participant — atelier intra-entreprise prêt à délivrer.",
    tomorrow_title="Atelier 'Manager IA-Ready' — Finalisation",
    tomorrow_desc="Finalisez l'évaluation, le guide formateur et la fiche commerciale de l'atelier."
))

DAYS.append(dict(
    day_num=58,week=10,phase=4,phase_name="Freelance IA",day_name="Jeudi",
    title="Atelier Manager IA-Ready — Finalisation",subtitle="Évaluation · Guide formateur · Fiche commerciale",
    accent_color=RED,
    theory_title="Finaliser un Produit de Formation",
    theory_points=[
        ("Check-list de finalisation","Programme / supports / ateliers / évaluation / guide formateur / fiche commerciale"),
        ("Guide formateur","Instructions d'animation / réponses aux objections courantes / timing de précision"),
        ("Évaluation adaptée","Pour les managers : évaluation action (plan 30 jours) plutôt que quiz théorique"),
        ("Fiche commerciale","1 page : public / objectifs / programme résumé / durée / format / prix / contact"),
        ("Témoignage fictif","Créez un témoignage type de manager satisfait — utile en attendant les vrais"),
        "Un produit finalisé avec guide formateur peut être délivré par un autre formateur — scalabilité",
    ],
    theory_insight="La fiche commerciale est la première chose que voit un DRH ou un acheteur formation — soignez-la.",
    exercise_title="Compléter le Kit et Créer la Fiche Commerciale",
    exercise_steps=[
        ("📋","Évaluation action","Plan d'engagement 30 jours : 3 actions IA à mettre en place / indicateurs / date bilan",RED),
        ("📖","Guide formateur","Conseils animation / gestion du groupe / questions anticipées / adaptations possibles",ORANGE),
        ("🎨","Fiche commerciale Canva","1 page A4 : accroche / objectifs / programme / formateur / tarif / contact",TEAL),
        ("🧪","Test complet","Simulez la journée entière avec Claude — identifiez les points à améliorer",GREEN),
    ],
    production_title="Atelier Manager — Kit Finalisé",
    production_items=[
        ("Kit complet","Programme + slides + ateliers + livret + évaluation + guide formateur"),
        ("Fiche commerciale","Design Canva professionnel — prête à envoyer aux DRH CA"),
        ("Email de prospection","Template email pour proposer l'atelier à un responsable RH ou directeur CA"),
        "Vous avez maintenant 3 formations commercialisables — votre catalogue est complet",
    ],
    livrable="Atelier Manager IA-Ready — Kit Finalisé + Fiche Commerciale",
    livrable_desc="Kit complet + fiche commerciale + email prospect — prêt à commercialiser immédiatement.",
    tomorrow_title="Site Web Freelance — Textes Complets",
    tomorrow_desc="Rédigez tous les textes de votre site web avec l'IA : accueil, offres, à propos, contact, blog."
))

DAYS.append(dict(
    day_num=59,week=10,phase=4,phase_name="Freelance IA",day_name="Vendredi",
    title="Site Web Freelance — Textes Complets",subtitle="Accueil · Offres · À propos · SEO · Contact",
    accent_color=GOLD,
    theory_title="Le Site Web Freelance — Votre Vitrine Permanente",
    theory_points=[
        ("Pourquoi un site","Crédibilité immédiate / disponible 24h/7j / base pour le SEO / portfolio visible"),
        ("Pages essentielles","Accueil / Offres (4 pages) / À propos / Portfolio / Blog / Contact"),
        ("Textes SEO","Mots-clés : 'formatrice freelance IA' / 'audit qualité centres d'appels' / 'formation ISO IA'"),
        ("Options no-code","Notion Sites (simple) / Carrd (rapide) / Webflow (puissant) / WordPress (complet)"),
        ("Ce que l'IA fait","Claude rédige tous les textes — vous validez, personnalisez, corrigez"),
        "Un site bien référencé génère des leads entrants sans effort — investissement une fois, revenus longtemps",
    ],
    theory_insight="L'IA rédige en 2 heures ce qui vous prendrait 2 semaines — concentrez-vous sur la vérité et la précision.",
    exercise_title="Rédiger les Textes du Site",
    exercise_steps=[
        ("🏠","Page Accueil","Accroche PVU / problème client / votre solution / 3 preuves / CTA — 300 mots",GOLD),
        ("💼","Pages Offres × 4","Chaque offre : titre / problème / solution / livrables / témoignage / prix / CTA",ORANGE),
        ("👤","Page À Propos","Votre histoire / expertise / certifications / approche / valeurs — authentique",TEAL),
        ("✍️","Page Blog × 3","3 articles de 500 mots sur vos thèmes d'expertise — SEO ciblé",GREEN),
    ],
    production_title="Textes Complets Site Web — 10 Pages",
    production_items=[
        ("10 pages rédigées","Accueil / 4 offres / à propos / portfolio / blog ×3 / contact"),
        ("Méta-descriptions SEO","Pour chaque page : titre SEO + description 155 caractères + mots-clés"),
        ("Appels à l'action","CTA cohérents sur toutes les pages — formulaire / email / LinkedIn / téléphone"),
        "Publiez sur Notion Sites ou Carrd ce weekend — version 1 vaut mieux que perfection en attente",
    ],
    livrable="Textes Complets Site Web — 10 Pages Prêtes à Publier",
    livrable_desc="Tous les textes rédigés avec SEO intégré — publiables immédiatement sur n'importe quelle plateforme.",
    tomorrow_title="Dossier Commercial Complet",
    tomorrow_desc="Assemblez votre dossier commercial : présentation + 4 offres + catalogue + tarifs + références fictives."
))

DAYS.append(dict(
    day_num=60,week=10,phase=4,phase_name="Freelance IA",day_name="Samedi",
    title="Dossier Commercial Complet",subtitle="Assembler · Cohérence · Prêt à prospecter",
    accent_color=GREEN,
    theory_title="Le Dossier Commercial — Votre Arsenal de Vente",
    theory_points=[
        ("Contenu du dossier","Pitch deck / 4 fiches offres / catalogue formations / grille tarifaire / références / CGV"),
        ("Cohérence visuelle","Même charte graphique / mêmes couleurs / mêmes polices — identité professionnelle forte"),
        ("Formats adaptés","PDF pour email / PPTX pour réunion / Notion pour lien web — 3 formats utiles"),
        ("Références fictives","3 études de cas fictives bien écrites valent des vraies en attendant — mentionnez 'exemple'"),
        ("Prochaine étape","Prospecter activement dès lundi — 10 messages par jour / objectif : 1 rendez-vous par semaine"),
        "Un dossier commercial cohérent et professionnel multiplie votre taux de conversion par 3",
    ],
    theory_insight="La qualité de votre dossier commercial reflète la qualité de votre travail — c'est votre première impression.",
    exercise_title="Assembler et Vérifier la Cohérence",
    exercise_steps=[
        ("📦","Rassembler tous les documents","Pitch deck / 4 fiches offres / catalogue / grille tarifaire / guide PDF",GREEN),
        ("🎨","Harmoniser la charte","Même logo / couleurs / polices / style photo sur tous les documents",TEAL),
        ("📝","Vérifier les contenus","Textes cohérents avec votre PVU / prix cohérents / offres complémentaires",ORANGE),
        ("📤","Créer le dossier numérique","Dossier Google Drive ou Notion : bien organisé / liens fonctionnels / partage facile",NAVY),
    ],
    production_title="Dossier Commercial Complet et Cohérent",
    production_items=[
        ("Dossier structuré","Tous les documents / 3 formats (PDF/PPTX/lien) / bien nommés / à jour"),
        ("Email de présentation","Template email pour partager le dossier à un prospect"),
        ("Check-list pré-envoi","Vérifications avant chaque envoi prospect : personnalisation / mise à jour prix / lien actif"),
        "Lundi matin : envoyez votre premier message de prospection avec le dossier — démarrez la machine",
    ],
    livrable="Dossier Commercial Complet — Prêt à Prospecter",
    livrable_desc="Tous les supports commerciaux harmonisés + guide d'envoi — arsenal de vente opérationnel.",
    tomorrow_title="Prospection IA — 30 Messages Personnalisés",
    tomorrow_desc="Semaine 11 : identifiez 30 cibles et rédigez des messages de prospection personnalisés avec l'IA."
))

DAYS.append(dict(
    day_num=61,week=11,phase=4,phase_name="Freelance IA",day_name="Lundi",
    title="Prospection IA — 30 Messages Personnalisés",subtitle="ICP · Ciblage LinkedIn · Séquence emails",
    accent_color=NAVY,
    theory_title="La Prospection Augmentée par l'IA",
    theory_points=[
        ("ICP = Ideal Client Profile","Qui est votre client idéal ? Taille CA / secteur / poste décideur / problème principal"),
        ("Sources de cibles","LinkedIn Sales Navigator / annuaire AFRC / événements sectoriels / réseau existant"),
        ("Personnalisation IA","Claude analyse le profil LinkedIn d'un prospect → génère un message personnalisé en 30s"),
        ("Séquence en 3 temps","Contact 1 : valeur / Contact 2 (J+7) : relance soft / Contact 3 (J+14) : dernière chance"),
        ("Taux de réponse","Message personnalisé : 15-25% / message générique : 2-5% — la personnalisation vaut l'effort"),
        "30 contacts ciblés bien prospectés valent mieux que 300 messages génériques",
    ],
    theory_insight="Votre réseau de 15 ans + ciblage IA = pipeline commercial solide en moins d'une semaine.",
    exercise_title="Identifier 30 Cibles et Rédiger les Messages",
    exercise_steps=[
        ("🎯","Définir l'ICP","Avec Claude : profil idéal en 5 critères / 3 types de décideurs / 2 secteurs prioritaires",NAVY),
        ("🔍","Identifier 30 cibles","LinkedIn : directeurs CA / responsables formation / DRH / responsables qualité — 30 profils",RED),
        ("✍️","30 messages personnalisés","Pour chaque cible : contexte perso + accroche + valeur + CTA — 3-5 lignes max",TEAL),
        ("📋","Séquence complète","Contact 1 + relance J+7 + relance J+14 pour chaque cible — CRM Airtable",GREEN),
    ],
    production_title="Séquence de Prospection — 30 Cibles",
    production_items=[
        ("30 profils cibles","Dans Airtable : nom / poste / entreprise / canal de contact / message J1 rédigé"),
        ("30 messages J1","Personnalisés et prêts à envoyer — 1 par cible"),
        ("Séquences J+7 et J+14","Templates de relance adaptables — 2 modèles par type de cible"),
        "Envoyez 5 messages aujourd'hui — le meilleur moment pour prospecter c'est maintenant",
    ],
    livrable="Séquence Prospection — 30 Cibles × 3 Messages",
    livrable_desc="30 profils ciblés + 30 messages personnalisés + relances — pipeline commercial activé.",
    tomorrow_title="Proposition Commerciale et Devis Type",
    tomorrow_desc="Créez votre template de proposition commerciale de 8-10 pages — prête à personnaliser en 30 minutes."
))

DAYS.append(dict(
    day_num=62,week=11,phase=4,phase_name="Freelance IA",day_name="Mardi",
    title="Proposition Commerciale et Devis Type",subtitle="Structure · Argumentation · Template réutilisable",
    accent_color=RED,
    theory_title="La Proposition Commerciale Convaincante",
    theory_points=[
        ("Structure en 7 parties","Contexte client / Problème identifié / Votre solution / Livrables / Planning / Prix / CTA"),
        ("Règle fondamentale","80% sur le client et son problème / 20% sur vous — pas l'inverse"),
        ("Preuve sociale","Étude de cas / témoignage / chiffre concret — une preuve vaut mille promesses"),
        ("Clarté du prix","Prix visible et justifié — ne jamais cacher le tarif en fin de document"),
        ("Format","8-10 pages maximum / design professionnel / PDF envoyé + lien web pour suivi"),
        "Une proposition bien structurée se lit en 5 minutes et se décide en 5 secondes",
    ],
    theory_insight="La proposition commerciale est la traduction de votre diagnostic — elle doit montrer que vous avez compris le problème.",
    exercise_title="Créer le Template Proposition",
    exercise_steps=[
        ("📝","Template 8 pages","Claude structure les 7 sections avec placeholders — vous personnalisez pour chaque client",RED),
        ("🎨","Mise en page Canva","Design professionnel : couverture / sections claires / photos / charte graphique",ORANGE),
        ("💰","Section prix","3 options avec ce qui est inclus / exclu / tarif clairement affiché",TEAL),
        ("✅","Simulation","Remplissez le template pour un client fictif — est-ce convaincant pour un DRH CA ?",GREEN),
    ],
    production_title="Template Proposition Commerciale Complet",
    production_items=[
        ("Template 8-10 pages","PDF éditable / sections pré-structurées / placeholders clairement indiqués"),
        ("3 versions prix","Essentiel / Pro / Premium — avec différenciation claire des livrables"),
        ("Check-list pré-envoi","10 points à vérifier avant chaque envoi : personnalisation / cohérence prix / fautes"),
        "Avec ce template : une proposition personnalisée se produit en 30 minutes — gain majeur",
    ],
    livrable="Template Proposition Commerciale 8-10 Pages",
    livrable_desc="Template réutilisable avec 3 niveaux de prix — proposition personnalisée en 30 minutes.",
    tomorrow_title="Simulation d'Entretien Client",
    tomorrow_desc="Préparez et simulez un entretien de vente avec Claude jouant un DRH sceptique — maîtrisez les objections."
))

DAYS.append(dict(
    day_num=63,week=11,phase=4,phase_name="Freelance IA",day_name="Mercredi",
    title="Simulation d'Entretien Client",subtitle="Vente consultative · Objections · SPIN selling",
    accent_color=ORANGE,
    theory_title="La Vente Consultative pour Expert",
    theory_points=[
        ("SPIN Selling","Situation / Problème / Implication / Need-payoff — 4 types de questions pour révéler le besoin"),
        ("Écoute avant proposition","80% du temps : comprendre / 20% : proposer — les experts qui écoutent vendent plus"),
        ("10 objections fréquentes","'Trop cher' / 'On a déjà un formateur' / 'L'IA c'est pas pour nous' / 'On verra l'année prochaine'"),
        ("Réponse aux objections","Valider / Questionner / Reformuler / Proposer — jamais argumenter frontalement"),
        ("Closing","Demander clairement la décision / proposer une prochaine étape concrète / relancer si besoin"),
        "Un expert qui maîtrise la vente consultative peut facturer 30% de plus qu'un expert qui ne sait pas vendre",
    ],
    theory_insight="Vous avez l'expertise — apprenez à la vendre. La vente n'est pas de la manipulation, c'est de la communication.",
    exercise_title="Simuler 3 Entretiens Clients",
    exercise_steps=[
        ("🎭","Entretien 1 — DRH bienveillant","Claude joue un DRH intéressé — pratiquez SPIN / la proposition / le closing",ORANGE),
        ("🎭","Entretien 2 — Directeur sceptique","Claude joue un directeur CA résistant — gérez les 5 objections les plus dures",RED),
        ("🎭","Entretien 3 — Budget serré","Claude joue un responsable avec budget limité — créez une offre adaptée",TEAL),
        ("📝","Après chaque entretien","Demandez à Claude le feedback : qu'est-ce qui a bien marché / à améliorer",GREEN),
    ],
    production_title="Guide de Vente et Gestion des Objections",
    production_items=[
        ("10 objections + réponses","Format : objection / réponse recommandée / question de rebond / exemple",),
        ("Guide SPIN","Vos 15 questions SPIN adaptées aux contextes CA / formation / qualité",),
        ("Script de closing","3 versions : soft / direct / alternatif — selon le profil du client",),
        "Pratiquez 1 entretien simulé par semaine — la vente s'améliore comme toute compétence",
    ],
    livrable="Guide Vente Consultative + 10 Objections Répondues",
    livrable_desc="Argumentaire complet + script SPIN + réponses objections + techniques closing.",
    tomorrow_title="CGV, Contrat et Cadre Juridique IA",
    tomorrow_desc="Rédigez vos CGV, votre contrat de prestation et vos clauses IA — avec Claude comme assistant juridique."
))

DAYS.append(dict(
    day_num=64,week=11,phase=4,phase_name="Freelance IA",day_name="Jeudi",
    title="CGV, Contrat et Cadre Juridique IA",subtitle="RGPD · CGV · Contrat prestation · Clauses IA",
    accent_color=NAVY,
    theory_title="Le Cadre Juridique du Freelance IA en France",
    theory_points=[
        ("Statut juridique","Auto-entrepreneur / SASU / EURL — comparatif selon votre situation et ambition"),
        ("Obligations RGPD","En mission : traitement données clients / sous-traitant / registre / DPA"),
        ("Clauses IA indispensables","Usage des outils IA / confidentialité / propriété intellectuelle / limitation responsabilité"),
        ("CGV essentielles","Conditions de paiement / révisions incluses / droits d'auteur / résiliation"),
        ("Attention","Claude génère des bases juridiques — faites valider par un avocat pour les documents finaux"),
        "Un contrat bien rédigé prévient 90% des litiges — investissement minimal pour protection maximale",
    ],
    theory_insight="Vos documents juridiques montrent votre professionnalisme — les clients sérieux les apprécient et les attendent.",
    exercise_title="Rédiger les Documents Juridiques de Base",
    exercise_steps=[
        ("📄","CGV 2 pages","Claude génère une base : conditions paiement / délais / révisions / propriété IP / résiliation",NAVY),
        ("📄","Contrat prestation","Template : parties / objet / livrables / délais / prix / paiement / clauses IA",RED),
        ("📄","Clause IA spécifique","Rédiger une clause sur l'usage des outils IA dans vos prestations — transparence",TEAL),
        ("📄","Mentions légales","Pour votre site web : éditeur / hébergeur / RGPD / cookies — base légale",GREEN),
    ],
    production_title="Kit Juridique Freelance — 4 Documents",
    production_items=[
        ("CGV","Conditions générales de vente — 2 pages, à annexer à chaque devis"),
        ("Contrat de prestation","Template personnalisable en 30 minutes pour chaque nouvelle mission"),
        ("Clause IA","Paragraphe à insérer dans tout contrat de prestation — transparence et responsabilité"),
        "Faites relire par un avocat spécialisé freelance avant utilisation réelle — 1-2h de consultation suffisent",
    ],
    livrable="Kit Juridique Freelance IA — CGV + Contrat + Clause IA",
    livrable_desc="4 documents juridiques de base — à faire valider par un avocat avant usage client.",
    tomorrow_title="Kit Administratif Freelance Complet",
    tomorrow_desc="Finalisez votre kit administratif : devis type, facture, relance impayé, onboarding client."
))

DAYS.append(dict(
    day_num=65,week=11,phase=4,phase_name="Freelance IA",day_name="Vendredi",
    title="Kit Administratif Freelance Complet",subtitle="Devis · Facture · Onboarding client · Clôture mission",
    accent_color=RED,
    theory_title="L'Administration Freelance — Automatiser le Routin",
    theory_points=[
        ("Documents essentiels","Devis numéroté / facture conforme / relance impayé / onboarding client / clôture mission"),
        ("Devis professionnel","Numéro / date validité / livrables détaillés / prix HT+TVA / conditions de paiement"),
        ("Facturation","Numérotation séquentielle / mentions légales obligatoires / délai légal paiement 30 jours"),
        ("Onboarding client","Email bienvenue + accès outils partagés + planning + canal de communication"),
        ("Clôture mission","Bilan livraison + évaluation client + demande témoignage + proposition suite"),
        "10 minutes d'administration bien faite évite 10 heures de gestion de conflits",
    ],
    theory_insight="Un client bien onboardé et bien suivi renouvelle et recommande — votre meilleur commercial est votre dernier client satisfait.",
    exercise_title="Créer les 5 Documents Administratifs",
    exercise_steps=[
        ("📋","Devis type","Template Canva : numéro auto / lignes livrables / sous-total / TVA / conditions",RED),
        ("💰","Facture type","Même design que devis / mentions légales obligatoires / IBAN / conditions",ORANGE),
        ("📧","Onboarding email","Email bienvenue : liens / planning / interlocuteur / premières étapes / urgences",TEAL),
        ("🏁","Email clôture","Bilan mission / livrables remis / demande témoignage / proposition de suite",GREEN),
    ],
    production_title="Kit Administratif Complet",
    production_items=[
        ("5 documents finalisés","Devis / Facture / Relance impayé / Onboarding / Clôture mission"),
        ("Intégration Airtable","Chaque document déclenché automatiquement par le statut dans votre CRM"),
        ("Email relance impayé","3 niveaux : rappel cordial / relance ferme / mise en demeure — avec délais"),
        "Ce kit vous fait gagner 2h/semaine d'administration — 100h/an récupérées",
    ],
    livrable="Kit Administratif Freelance — 5 Documents Opérationnels",
    livrable_desc="Devis + facture + relance + onboarding + clôture — gestion administrative automatisée.",
    tomorrow_title="Évaluation Phase 4 — Niveau 4 Maîtrise ?",
    tomorrow_desc="Quiz Phase 4 + portfolio 47 livrables + vérification niveau 4. Dernière phase avant la certification !"
))

DAYS.append(dict(
    day_num=66,week=11,phase=4,phase_name="Freelance IA",day_name="Samedi",
    title="Évaluation Phase 4 — Niveau 4 Maîtrise",subtitle="Quiz · Portfolio 47 livrables · Passage Phase 5",
    accent_color=GREEN,
    theory_title="Grille Évaluation — Niveau 4 Maîtrise",
    theory_points=[
        ("Critère 1 /20","Positionnement clair et différenciant — PVU tenue en 30 secondes sans hésitation"),
        ("Critère 2 /20","3 offres packagées avec prix justifiés + 3 GPT Custom opérationnels"),
        ("Critère 3 /20","Former quelqu'un d'autre à l'IA — démonstration avec Claude comme apprenant"),
        ("Critère 4 /20","47 livrables dans le portfolio + dossier commercial complet"),
        ("Critère 5 /20","Prospection active : 10+ messages envoyés + 1 rendez-vous obtenu ou en cours"),
        "Score ≥ 60/100 = Niveau 4 validé. Passage en Phase 5 : Maîtrise et Certification.",
    ],
    theory_insight="Phase 4 terminée — vous êtes maintenant une consultante freelance IA opérationnelle. Il ne manque plus que la certification.",
    exercise_title="Quiz Phase 4 — 20 Questions",
    exercise_steps=[
        ("📝","Questions 1-5","PVU / tarification / GPT Custom / prompt packs / LinkedIn stratégie",GREEN),
        ("📝","Questions 6-10","Formations créées / outils déployés / prospection / proposition commerciale / vente consultative",TEAL),
        ("📝","Questions 11-15","CGV / contrat prestation / kit admin / CRM / automatisations",ORANGE),
        ("📝","Questions 16-20","Pitchez votre activité en 60s / présentez une offre / gérez une objection / décrivez votre valeur",NAVY),
    ],
    production_title="Portfolio Phase 4 + Bilan Commercial",
    production_items=[
        ("47 livrables documentés","Avec valeur commerciale estimée pour chaque — votre bilan de valeur créée"),
        ("Bilan prospection","Messages envoyés / réponses / rendez-vous / pipeline estimé"),
        ("Score Phase 4","Honnête et documenté — plan de renforcement si < 60"),
        "Vous avez créé plus de valeur en 11 semaines que la plupart des consultants en 1 an — continuez.",
    ],
    livrable="Portfolio Phase 4 — 47 Livrables + Bilan Commercial",
    livrable_desc="Phase 4 complétée. Score ≥ 60 = Niveau 4 certifié. Entrée en Phase 5 : Maîtrise finale.",
    tomorrow_title="Phase 5 — Projet Capstone : Mission Complète",
    tomorrow_desc="Dernière phase — vous conduisez une mission freelance complète fictive de A à Z. C'est votre chef-d'œuvre."
))

# ═══════════════════════════════════════
# PHASE 5 — MAÎTRISE (Semaines 12-13)
# ═══════════════════════════════════════

DAYS.append(dict(
    day_num=67,week=12,phase=5,phase_name="Maîtrise",day_name="Lundi",
    title="Projet Capstone — Diagnostic Initial",subtitle="Mission fictive complète · CA 80 agents · Audit initial",
    accent_color=GOLD,
    theory_title="Le Projet Capstone — Votre Chef-d'Œuvre",
    theory_points=[
        ("Principe","Conduire une mission freelance complète fictive — de l'audit au bilan de mission"),
        ("Client fictif","CenterCall France — 80 agents / CA entrant / secteur télécom / taux réclamation 9%"),
        ("Votre mission","Audit qualité formation + diagnostic + PAC + formations + suivi — 6 semaines simulées"),
        ("Jour 67 : Diagnostic","Analyser le brief client / conduire l'audit initial / rédiger le rapport de diagnostic"),
        ("Votre démarche","Utilisez tous vos outils IA — c'est la démonstration de votre maîtrise"),
        "Ce projet est votre référence portfolio la plus solide — un dossier de mission réel de 50 pages",
    ],
    theory_insight="Traitez ce projet fictif comme un vrai client — chaque document doit être d'une qualité que vous seriez fière de remettre.",
    exercise_title="Analyser le Brief et Conduire l'Audit",
    exercise_steps=[
        ("📋","Analyser le brief","Lire le brief CenterCall / identifier les enjeux / préparer les questions de cadrage",GOLD),
        ("🔍","Audit initial","Checklist audit Jour 32 + analyse données fictives KPIs / formation / qualité",ORANGE),
        ("📊","Collecter les informations","Simuler des entretiens avec Claude (DRH / responsable qualité / superviseur / agent)",TEAL),
        ("📄","Rapport de diagnostic","15 pages : contexte / méthodologie / constats / synthèse / priorités",GREEN),
    ],
    production_title="Rapport de Diagnostic — CenterCall France",
    production_items=[
        ("Brief client","Document de cadrage signé — périmètre / objectifs / livrables attendus / délais"),
        ("Rapport diagnostic 15 pages","Constats terrain + analyse + cartographie des problèmes + priorités"),
        ("Plan de la mission","6 semaines / 5 chantiers / jalons / livrables / budget prévisionnel"),
        "Ce rapport seul justifie une prestation de 1500-3000€ en réel — prenez-le au sérieux",
    ],
    livrable="Rapport Diagnostic — CenterCall France (15 pages)",
    livrable_desc="Audit initial complet avec constats, analyse et plan de mission — premier livrable du capstone.",
    tomorrow_title="Projet Capstone — Plan d'Action",
    tomorrow_desc="À partir du diagnostic : construire le PAC 90 jours avec 5 chantiers et 25 actions priorisées."
))

DAYS.append(dict(
    day_num=68,week=12,phase=5,phase_name="Maîtrise",day_name="Mardi",
    title="Projet Capstone — Plan d'Action 90 Jours",subtitle="5 chantiers · 25 actions · Priorisation · Budget",
    accent_color=TEAL,
    theory_title="Du Diagnostic au Plan — La Logique de Mission",
    theory_points=[
        ("Lien diagnostic-plan","Chaque constat du diagnostic génère une ou plusieurs actions dans le PAC"),
        ("5 chantiers CenterCall","Qualité appels / Formation continue / Management superviseurs / Outils IA / Processus"),
        ("25 actions priorisées","Matrice impact/effort : quick wins (semaine 1) / court terme (mois 1) / moyen terme"),
        ("Budget réaliste","Coût interne (temps agents/superviseurs) + coût externe (formation / outils) / ROI estimé"),
        ("Jalons et indicateurs","Pour chaque chantier : jalons mensuels / indicateurs de succès / seuils d'alerte"),
        "Un PAC sans priorisation est une liste de vœux — la priorisation est votre valeur de consultante",
    ],
    theory_insight="Le client ne peut pas tout faire en même temps — votre rôle est de lui dire quoi faire en premier et pourquoi.",
    exercise_title="Construire le PAC CenterCall",
    exercise_steps=[
        ("🗂️","5 chantiers","Pour chaque chantier : objectif / enjeu / actions / responsable / délai / indicateur",TEAL),
        ("📊","Matrice priorisation","Quick wins (impact fort / effort faible) en priorité — visuel Canva",GREEN),
        ("💰","Budget prévisionnel","Coûts directs / coûts indirects / ROI attendu à 3 mois et 12 mois",ORANGE),
        ("🗓️","Planning Gantt","Gantt 12 semaines dans Notion ou Excel — toutes les actions / dépendances",NAVY),
    ],
    production_title="PAC 90 Jours — CenterCall France",
    production_items=[
        ("PAC structuré","5 chantiers / 25 actions / responsables / délais / indicateurs — format consultant"),
        ("Budget prévisionnel","Tableau coûts / économies / ROI — argument de vente fort"),
        ("Planning Gantt","Visuel 12 semaines — séquençage logique des actions"),
        "Ce PAC + budget + Gantt = livrables types d'une mission de conseil à 5000-8000€",
    ],
    livrable="PAC 90 Jours CenterCall — 5 Chantiers / 25 Actions",
    livrable_desc="Plan d'action complet + budget + Gantt — cœur de la mission capstone.",
    tomorrow_title="Projet Capstone — Programme Formation",
    tomorrow_desc="Concevez le programme de formation associé au PAC — 3 modules clés pour CenterCall."
))

DAYS.append(dict(
    day_num=69,week=12,phase=5,phase_name="Maîtrise",day_name="Mercredi",
    title="Projet Capstone — Programme de Formation",subtitle="3 modules clés · Supports complets · Plan de déploiement",
    accent_color=GREEN,
    theory_title="La Formation comme Levier du PAC",
    theory_points=[
        ("Lien PAC-formation","Chaque chantier qui nécessite un changement de comportement = module de formation"),
        ("3 modules CenterCall","Qualité appels + gestion client difficile + outils IA pour agents — liés aux chantiers"),
        ("Format adapté","Agents CA : préférence présentiel / groupes de 10-12 / durée max 3h par session"),
        ("Plan de déploiement","Qui / quand / comment / avec quoi — planning de déploiement sur 3 mois"),
        ("Mesure d'impact","Kirkpatrick 4 niveaux / indicateurs PAC / suivi post-formation — preuve du ROI"),
        "La formation sans suivi est un coût — avec suivi c'est un investissement avec ROI mesurable",
    ],
    theory_insight="Ces 3 modules sont les livrables formation du capstone — ils doivent être de la même qualité que vos modules commerciaux.",
    exercise_title="Concevoir les 3 Modules Formation",
    exercise_steps=[
        ("📚","Module 1 — Qualité appels","Objectifs / plan / contenu clé / 2 exercices / évaluation — 3h format",GREEN),
        ("📚","Module 2 — Client difficile","Réutilisez et adaptez le module Jour 9 au contexte CenterCall",TEAL),
        ("📚","Module 3 — Outils IA agents","Introduction IA pour agents non-experts — pratique / rassurante / utile",ORANGE),
        ("🗓️","Plan de déploiement","Qui suit quoi / quand / en quelle priorité / avec quel formateur",NAVY),
    ],
    production_title="3 Modules Formation CenterCall",
    production_items=[
        ("3 plans de cours","Objectifs / plan / contenu / exercices / évaluation — format synthétique",),
        ("Supports essentiels","Slides clés + livret participant pour chaque module",),
        ("Plan de déploiement","Tableau : module / groupe / date / formateur / salle / pré-requis",),
        "Ces 3 modules démontrent votre capacité à aligner formation et stratégie — compétence rare",
    ],
    livrable="3 Modules Formation CenterCall + Plan de Déploiement",
    livrable_desc="Plans de cours + supports + plan de déploiement — volet formation du capstone.",
    tomorrow_title="Projet Capstone — Outils et Automatisation",
    tomorrow_desc="Déployez les outils IA opérationnels pour CenterCall : dashboard qualité, coaching IA, évaluations auto."
))

DAYS.append(dict(
    day_num=70,week=12,phase=5,phase_name="Maîtrise",day_name="Jeudi",
    title="Projet Capstone — Outils et Automatisation",subtitle="Dashboard qualité · Coaching IA · Évaluations auto",
    accent_color=ORANGE,
    theory_title="Déployer les Outils IA en Mission",
    theory_points=[
        ("Logique de déploiement","Les outils IA ne se larguent pas — ils s'introduisent graduellement avec formation et suivi"),
        ("3 outils CenterCall","Dashboard KPI IA / Outil coaching superviseur / Système évaluation automatisé"),
        ("Change management","Résistance probable des superviseurs — présentation bénéfices / pas de contrainte initiale"),
        ("Formation aux outils","30 minutes par outil suffisent si l'interface est simple et le bénéfice immédiat"),
        ("Suivi adoption","Métriques d'utilisation / feedback utilisateurs / ajustements rapides"),
        "Un outil non utilisé n'a aucune valeur — l'accompagnement au changement est votre différenciateur",
    ],
    theory_insight="Vous ne livrez pas des outils — vous accompagnez l'adoption. C'est cette posture qui fidélise les clients.",
    exercise_title="Configurer les 3 Outils CenterCall",
    exercise_steps=[
        ("📊","Dashboard KPI CenterCall","Adaptez le dashboard Jour 16 aux données et KPIs spécifiques de CenterCall",ORANGE),
        ("🤖","Outil coaching superviseur","Configurez le prompt coach Jour 21 avec la grille QA et les valeurs CenterCall",RED),
        ("📝","Évaluations automatisées","3 formulaires Typeform adaptés au contexte CenterCall — formations + qualité",TEAL),
        ("📋","Guide déploiement","Procédure de déploiement de chaque outil : qui / comment / dans quel ordre",GREEN),
    ],
    production_title="3 Outils IA Configurés pour CenterCall",
    production_items=[
        ("Dashboard KPI","Adapté aux données CenterCall — opérationnel immédiatement"),
        ("Outil coaching","Configuré avec grille QA CenterCall — prêt à utiliser par les superviseurs"),
        ("3 formulaires évaluation","Typés CenterCall — liens actifs et testés"),
        "Ce volet démontre votre capacité à personnaliser des outils génériques — compétence commercialement précieuse",
    ],
    livrable="3 Outils IA Déployés + Guide Déploiement CenterCall",
    livrable_desc="Dashboard + coaching IA + évaluations auto + guide — volet outils du capstone.",
    tomorrow_title="Projet Capstone — Présentation Direction",
    tomorrow_desc="Préparez et simulez la présentation de restitution de mission devant le comité de direction CenterCall."
))

DAYS.append(dict(
    day_num=71,week=12,phase=5,phase_name="Maîtrise",day_name="Vendredi",
    title="Projet Capstone — Présentation Direction",subtitle="Restitution de mission · 30 slides · Soutenance simulée",
    accent_color=NAVY,
    theory_title="Présenter une Mission au Comité de Direction",
    theory_points=[
        ("Enjeu","Le CODIR n'a pas suivi la mission — il veut des résultats, pas du détail"),
        ("Structure executive","Situation initiale / Ce que nous avons fait / Résultats obtenus / Prochaines étapes / Budget"),
        ("En 20 minutes","3 minutes contexte / 10 minutes résultats / 5 minutes recommandations / 2 minutes questions"),
        ("Langage CODIR","Chiffres / ROI / risques / opportunités — pas de jargon pédagogique"),
        ("Recommandations","3 priorités stratégiques pour les 6 prochains mois — simples et actionnables"),
        "Un CODIR convaincu = renouvellement de contrat + recommandation à d'autres directions",
    ],
    theory_insight="Votre capacité à parler le langage des dirigeants est aussi importante que votre expertise technique.",
    exercise_title="Préparer et Simuler la Présentation",
    exercise_steps=[
        ("🎨","30 slides Gamma","Structure executive : contexte / résultats / ROI / recommandations / prochaines étapes",NAVY),
        ("📝","Notes de présentation","Pour chaque slide : ce que vous dites en 30 secondes — cohérence + fluidité",RED),
        ("🎭","Simulation CODIR","Claude joue 3 directeurs avec des questions difficiles — pratiquez les réponses",TEAL),
        ("🔄","Affiner après simulation","Quelles questions vous ont pris par surprise ? Préparez les réponses",GREEN),
    ],
    production_title="Présentation Restitution + Simulation Validée",
    production_items=[
        ("30 slides executive","Gamma — design professionnel / données clés / visuels impactants / CTA clair"),
        ("Notes de présentation","Script complet — 20 minutes chrono maîtrisé"),
        ("Q&A préparé","15 questions difficiles + réponses calibrées"),
        "Cette présentation est l'aboutissement de toute la mission — soignez-la comme une vraie restitution",
    ],
    livrable="Présentation Restitution CODIR — 30 Slides + Notes",
    livrable_desc="Présentation executive CenterCall + script + Q&A — soutenance de mission professionnelle.",
    tomorrow_title="Assemblage Final du Dossier de Mission",
    tomorrow_desc="Assemblez tous les livrables du capstone en un dossier de mission complet de 50+ pages."
))

DAYS.append(dict(
    day_num=72,week=12,phase=5,phase_name="Maîtrise",day_name="Samedi",
    title="Assemblage Final — Dossier de Mission Complet",subtitle="50+ pages · Portfolio référence · Mise en page pro",
    accent_color=GREEN,
    theory_title="Le Dossier de Mission — Votre Référence Ultime",
    theory_points=[
        ("Contenu","Brief client / rapport diagnostic / PAC / programme formation / outils / présentation CODIR"),
        ("Table des matières","Structure claire — un lecteur doit trouver n'importe quelle section en 30 secondes"),
        ("Mise en page","Cohérence visuelle / pagination / numérotation des livrables / index"),
        ("Version portfolio","Anonymisez le client fictif → 'Client CA Télécom' — prêt à montrer à des prospects"),
        ("Version complète","Avec nom client fictif + tous les détails — version d'apprentissage"),
        "Ce dossier de 50 pages représente l'équivalent de 3 à 6 mois de mission réelle — c'est votre proof of concept",
    ],
    theory_insight="Ce dossier est votre meilleure carte de visite — il prouve par A+B que vous pouvez délivrer une mission complète.",
    exercise_title="Assembler et Mettre en Page",
    exercise_steps=[
        ("📦","Rassembler tous les livrables","Jours 67-71 : brief / diagnostic / PAC / formations / outils / présentation",GREEN),
        ("📝","Rédiger la synthèse executive","2 pages : la mission en résumé / résultats clés / valeur créée",TEAL),
        ("📐","Mise en page Canva/Word","Table des matières / pagination / couverture professionnelle / cohérence",ORANGE),
        ("✅","Relecture finale","Claude relit et identifie les incohérences / manques / formulations à améliorer",NAVY),
    ],
    production_title="Dossier de Mission Complet — 50+ Pages",
    production_items=[
        ("Dossier complet","Couverture / synthèse / 6 sections / annexes — 50+ pages professionnelles"),
        ("Version portfolio","Anonymisée / prête à montrer / résumé 1 page en couverture"),
        ("Pitch du dossier","Comment présenter ce dossier en 3 minutes à un prospect"),
        "Partagez ce dossier avec 3 personnes de confiance ce weekend — leur feedback est précieux",
    ],
    livrable="Dossier de Mission Complet — CenterCall France (50+ pages)",
    livrable_desc="Référence portfolio ultime — équivalent d'une mission réelle de conseil de 6 semaines.",
    tomorrow_title="Révision et Auto-Évaluation Finale",
    tomorrow_desc="Semaine 13 — révisez tout le programme et préparez l'évaluation finale de certification."
))

DAYS.append(dict(
    day_num=73,week=13,phase=5,phase_name="Maîtrise",day_name="Lundi",
    title="Révision Finale — Tout le Programme",subtitle="Rappel actif · Lacunes ciblées · Préparation certification",
    accent_color=RED,
    theory_title="La Révision Finale — Consolider et Confirmer",
    theory_points=[
        ("5 phases en rappel actif","Listez mentalement les 3 points clés de chaque phase — sans regarder vos notes"),
        ("Tester sa maîtrise","Expliquez chaque concept à Claude comme si c'était un apprenant novice"),
        ("Combler les dernières lacunes","Toutes les phases ont des gaps — identifiez et combler les 3 derniers"),
        ("Quiz de préparation","50 questions aléatoires sur tout le programme — simulation d'examen"),
        ("Confiance","À ce stade vous avez 78 jours de pratique IA intensive — vous êtes experte"),
        "La révision finale n'est pas pour combler des lacunes — c'est pour confirmer sa maîtrise",
    ],
    theory_insight="Vous avez parcouru 73 jours — ce que vous maîtrisez aujourd'hui dépasse la grande majorité du marché.",
    exercise_title="Quiz de Préparation Finale — 50 Questions",
    exercise_steps=[
        ("📝","Phase 1 — 10 questions","Fondamentaux IA / prompts / outils / éthique — 15 min",RED),
        ("📝","Phase 2 — 10 questions","NLP / coaching / formation / KPIs / soft skills — 15 min",ORANGE),
        ("📝","Phase 3 — 10 questions","ISO / audit / automatisation / lean / benchmark — 15 min",TEAL),
        ("📝","Phases 4+5 — 20 questions","Freelance / offres / prospection / mission capstone — 20 min",GREEN),
    ],
    production_title="Plan de Révision Personnalisé + Score 50 Questions",
    production_items=[
        ("Score des 50 questions","Identifiez précisément les 5 questions les plus difficiles"),
        ("Plan révision 2 jours","Ciblé sur vos 5 points faibles — 2h par point maximum"),
        ("Mind map du programme","Claude génère une carte mentale de tout le programme — vue d'ensemble"),
        "Un score ≥ 40/50 aujourd'hui = certification assurée vendredi. Continuez ainsi.",
    ],
    livrable="Quiz 50 Questions + Plan de Révision Personnalisé",
    livrable_desc="Auto-évaluation complète sur tout le programme + ciblage des dernières révisions.",
    tomorrow_title="Veille IA et Communauté",
    tomorrow_desc="Activez votre présence dans les communautés IA francophones et publiez votre première contribution."
))

DAYS.append(dict(
    day_num=74,week=13,phase=5,phase_name="Maîtrise",day_name="Mardi",
    title="Veille IA et Intégration Communauté",subtitle="Communautés francophones · Contribution · Réseau",
    accent_color=GOLD,
    theory_title="Faire Partie de la Communauté IA",
    theory_points=[
        ("Pourquoi la communauté","Veille accélérée / opportunités mission / partenariats / recommandations / visibilité"),
        ("Communautés FR actives","LinkedIn #IAFormation / Slack IA-France / Discord prompt engineers / AFNOR groupes"),
        ("Contribuer avant de demander","Partagez votre expertise avant de chercher des missions — donner pour recevoir"),
        ("Votre valeur dans la communauté","Vous êtes l'une des rares avec formation + CA + ISO + IA — votre niche est recherchée"),
        ("Objectif 3 mois","5 contributions / 100 nouveaux contacts qualifiés / 1 collaboration"),
        "Votre réseau est votre actif le plus précieux — l'IA peut le développer mais pas le remplacer",
    ],
    theory_insight="Dans une communauté d'experts, celui qui donne le plus reçoit le plus. Commencez par partager.",
    exercise_title="S'Inscrire et Contribuer",
    exercise_steps=[
        ("🤝","Rejoindre 3 communautés","LinkedIn groupe Formation IA / Slack IA-France / Discord ou groupe WhatsApp",GOLD),
        ("✍️","Première contribution","Partagez un insight concret de votre parcours — court / utile / authentique",ORANGE),
        ("🔗","Connecter 10 personnes","Demandes de connexion personnalisées à 10 experts de votre domaine",TEAL),
        ("📅","Calendrier contributions","1 contribution par semaine sur LinkedIn / 2 par mois en communauté spécialisée",GREEN),
    ],
    production_title="Présence Communauté Activée",
    production_items=[
        ("3 communautés rejointes","Profils actifs / première contribution publiée dans chacune"),
        ("10 connexions envoyées","Messages personnalisés — pas de copier-coller"),
        ("Post LinkedIn contribution","Votre premier post de partage d'expérience TAMOU NEURAL PATH"),
        "Votre réseau actif dans 6 mois commencera à vous envoyer des opportunités passives",
    ],
    livrable="Présence Communauté IA — 3 Plateformes Activées",
    livrable_desc="Inscriptions + premières contributions + connexions — réseau IA francophone en construction.",
    tomorrow_title="Plan de Développement Professionnel 12 Mois",
    tomorrow_desc="Construisez votre roadmap professionnelle IA pour les 12 prochains mois — formations, certifications, objectifs CA."
))

DAYS.append(dict(
    day_num=75,week=13,phase=5,phase_name="Maîtrise",day_name="Mercredi",
    title="Roadmap Professionnelle IA — 12 Mois",subtitle="Certifications · Objectifs CA · Nouvelles compétences",
    accent_color=TEAL,
    theory_title="Planifier la Suite avec l'IA",
    theory_points=[
        ("Certifications recommandées","Google AI Essentials (gratuit) / Microsoft AI-900 / HubSpot AI / Qualiopi + IA"),
        ("Compétences à approfondir","Agents IA / analyse de données / vidéo IA / prompt engineering avancé"),
        ("Objectifs commerciaux","CA cible à 6 mois / 12 mois / nb missions / nb clients récurrents"),
        ("Stratégie de contenu","Newsletter mensuelle / webinaire trimestriel / livre blanc annuel"),
        ("Partenariats","1 cabinet conseil / 1 éditeur LMS / 1 association AFNOR — pour co-développement"),
        "Le marché IA-Formation va croître de 35% par an — vous êtes positionnée pour en capturer une part",
    ],
    theory_insight="La meilleure stratégie pour les 12 prochains mois : aller chercher les 3 premières missions et les réussir parfaitement.",
    exercise_title="Construire la Roadmap",
    exercise_steps=[
        ("🗓️","Roadmap trimestrielle","Q1 : premières missions / Q2 : scale / Q3 : certifications / Q4 : nouvelles offres",TEAL),
        ("🎯","3 objectifs par trimestre","Mesurables / atteignables / avec indicateurs de suivi",GREEN),
        ("📚","Plan de formation continue","2 formations ou certifications par an minimum — budget et calendrier",ORANGE),
        ("💰","Objectifs financiers","CA mensuel cible / nb clients / TJM cible / charges prévisionnelles",NAVY),
    ],
    production_title="Roadmap Professionnelle IA — 12 Mois",
    production_items=[
        ("Document roadmap","Trimestriel / objectifs / actions / formations / jalons / indicateurs"),
        ("Tableau de bord objectifs","Notion : suivi mensuel de vos indicateurs clés freelance"),
        ("Plan formation 12 mois","Certifications / formations / lectures / communautés — avec dates"),
        "Votre roadmap est votre contrat avec vous-même — révisez-la tous les 3 mois",
    ],
    livrable="Roadmap Professionnelle IA — 12 Mois",
    livrable_desc="Plan de développement trimestriel + objectifs + formations — votre GPS professionnel IA.",
    tomorrow_title="Soutenance Finale — Jury IA",
    tomorrow_desc="Présentez votre portfolio complet et votre mission capstone devant Claude jouant un jury de 3 experts."
))

DAYS.append(dict(
    day_num=76,week=13,phase=5,phase_name="Maîtrise",day_name="Jeudi",
    title="Soutenance Finale — Jury IA",subtitle="Présenter le portfolio · Défendre les choix · Jury simulé",
    accent_color=RED,
    theory_title="La Soutenance — Démonstration de Maîtrise",
    theory_points=[
        ("Format","20 minutes de présentation / 20 minutes de questions / 5 minutes de délibération"),
        ("Ce que le jury évalue","Maîtrise théorique / qualité des livrables / autonomie / valeur commerciale / vision"),
        ("Préparer les questions difficiles","Pourquoi ce prix ? / Si le client refuse ? / Comment vous différenciez-vous ? / Prochaine offre ?"),
        ("Posture","Expert(e) qui partage son expérience — pas un(e) étudiant(e) qui passe un examen"),
        ("Conclusion","Votre vision à 2 ans — montrez que vous pensez déjà au-delà de la certification"),
        "Une soutenance réussie n'est pas sans erreurs — c'est une avec des réponses honnêtes et confiantes",
    ],
    theory_insight="Après 76 jours, vous savez plus sur l'IA appliquée à la formation CA que 99% des gens qui se présentent sur ce marché.",
    exercise_title="Préparer et Simuler la Soutenance",
    exercise_steps=[
        ("📊","Préparer la présentation 20 min","Portfolio en 5 sections / livrable phare / mission capstone / vision",RED),
        ("🎭","Simuler avec Claude jury","Claude joue 3 experts aux personnalités différentes — questions variées et difficiles",ORANGE),
        ("💪","Renforcer les réponses faibles","Identifiez les 3 réponses les plus hésitantes — préparez des réponses plus solides",TEAL),
        ("🔄","Deuxième simulation","Seconde soutenance plus fluide — validez la progression",GREEN),
    ],
    production_title="Soutenance Validée — Compte-Rendu",
    production_items=[
        ("Présentation 20 min","Slides + notes — version finale après 2 simulations"),
        ("Compte-rendu jury","Feedback Claude : points forts / axes d'amélioration / recommandation"),
        ("Réponses Q&A finalisées","20 questions avec réponses polies — votre aide-mémoire"),
        "Dormez bien ce soir — la certification est demain et vous êtes prête.",
    ],
    livrable="Soutenance Simulée — Compte-Rendu + Présentation Finale",
    livrable_desc="2 simulations effectuées + compte-rendu jury IA + présentation finalisée — prête pour la certification.",
    tomorrow_title="CERTIFICATION — Évaluation Finale",
    tomorrow_desc="Demain : évaluation finale. Votre parcours de 78 jours se conclut par la certification TAMOU NEURAL PATH."
))

DAYS.append(dict(
    day_num=77,week=13,phase=5,phase_name="Maîtrise",day_name="Vendredi",
    title="ÉVALUATION FINALE — Certification TAMOU NEURAL PATH",subtitle="Quiz 50 questions · Portfolio final · Score global",
    accent_color=GOLD,
    theory_title="Grille de Certification — Niveau Maîtrise",
    theory_points=[
        ("Critère 1 /20","Mission capstone complète + dossier 50 pages — qualité professionnelle"),
        ("Critère 2 /20","Portfolio 52+ livrables documentés avec valeur commerciale"),
        ("Critère 3 /20","Autonomie complète sur 12+ outils IA — démonstration live"),
        ("Critère 4 /20","3 offres packagées opérationnelles + prospection active engagée"),
        ("Critère 5 /20","Vision et plan de développement 12 mois cohérent et réaliste"),
        "Score ≥ 80/100 = CERTIFIÉE TAMOU NEURAL PATH — Expert(e) IA Formation Centres d'Appels",
    ],
    theory_insight="Ce n'est pas la fin d'un apprentissage — c'est le début d'une nouvelle expertise qui va encore croître.",
    exercise_title="Évaluation Finale — 50 Questions Complètes",
    exercise_steps=[
        ("📝","Quiz théorique","20 questions sur les 5 phases — sans aide / 30 minutes",GOLD),
        ("🛠️","Quiz pratique","10 situations : que feriez-vous ? / quel outil ? / quel prompt ? / quel livrable ?",ORANGE),
        ("💼","Quiz commercial","10 questions : pitcher une offre / répondre à une objection / justifier un tarif",TEAL),
        ("📊","Évaluation portfolio","10 questions sur vos livrables : valeur / originalité / qualité",GREEN),
    ],
    production_title="Score de Certification + Récapitulatif Global",
    production_items=[
        ("Score global /100","Calcul honnête des 5 critères — votre niveau certifié"),
        ("Attestation de parcours","Document TAMOU NEURAL PATH : compétences acquises / livrables produits / durée"),
        ("Bilan de valeur créée","Estimez la valeur commerciale totale de vos 52+ livrables"),
        "Score ≥ 80 : CERTIFIÉE ✓ / Score 60-79 : Compétente ✓ / Score < 60 : Révision ciblée 1 semaine",
    ],
    livrable="CERTIFICATION TAMOU NEURAL PATH — Score Final",
    livrable_desc="Évaluation finale + score + attestation de parcours — votre certification IA Formation CA.",
    tomorrow_title="LANCEMENT — Votre Activité Freelance IA",
    tomorrow_desc="Demain : célébration et lancement officiel. Publiez votre post LinkedIn d'annonce et envoyez vos 10 premiers messages."
))

DAYS.append(dict(
    day_num=78,week=13,phase=5,phase_name="Maîtrise",day_name="Samedi",
    title="LANCEMENT — Activité Freelance IA Officiellement Ouverte",subtitle="Célébrer · Annoncer · Premier pas commercial",
    accent_color=GREEN,
    theory_title="Le Lancement — De la Formation à l'Action",
    theory_points=[
        ("Célébrez d'abord","78 jours d'effort intensif méritent une vraie célébration — marquez ce moment"),
        ("Post LinkedIn d'annonce","Votre histoire en 3 paragraphes : avant / parcours / maintenant — authentique et humain"),
        ("10 messages de prospection","Les 10 premières cibles de votre liste Jour 61 — envoyez aujourd'hui"),
        ("Votre première mission","Peut-être déjà en cours — sinon elle arrive dans les prochaines semaines"),
        ("L'IA continue d'évoluer","Votre veille hebdomadaire est maintenant essentielle — le marché bouge vite"),
        "Le diplôme le plus précieux en IA, c'est la pratique — vous l'avez. Maintenant, livrez.",
    ],
    theory_insight="Vous n'êtes plus une novice en IA. Vous êtes une experte qui a 15 ans de terrain ET une maîtrise IA concrète. C'est exceptionnel.",
    exercise_title="Lancement Officiel",
    exercise_steps=[
        ("🎉","Célébration","Prenez le temps de mesurer le chemin parcouru — 78 jours / 52 livrables / 312 heures",GREEN),
        ("📢","Post LinkedIn d'annonce","'Après 3 mois intensifs, je suis officiellement...' — votre histoire vraie",TEAL),
        ("📤","10 messages de prospection","Les 10 premières cibles — messages personnalisés Jour 61",ORANGE),
        ("🗓️","Agenda semaine 1 post-formation","Définissez vos 3 priorités commerciales pour les 7 prochains jours",NAVY),
    ],
    production_title="Lancement Officiel TAMOU NEURAL PATH",
    production_items=[
        ("Post LinkedIn d'annonce","Authentique / inspirant / avec appel à l'action discret"),
        ("10 messages envoyés","Dans votre CRM Airtable : statuts mis à jour"),
        ("Agenda semaine prochaine","3 priorités commerciales claires + 1 objectif de veille"),
        "Félicitations ! Vous avez accompli quelque chose de remarquable. L'aventure continue.",
    ],
    livrable="LANCEMENT OFFICIEL — Post LinkedIn + 10 Prospections",
    livrable_desc="Annonce publique de votre activité + premières prospections — Jour 1 de votre vie professionnelle IA.",
    tomorrow_title="La Suite — Votre Avenir IA",
    tomorrow_desc="TAMOU NEURAL PATH est terminé. Votre expertise IA × Formation × CA × ISO est certifiée. Le marché vous attend."
))

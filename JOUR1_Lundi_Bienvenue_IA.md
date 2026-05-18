# TAMOU NEURAL PATH
## LUNDI — Jour 1 : Bienvenue dans l'IA
**Semaine 1 | Phase 1 — Fondations**

---

```
PLANNING DU JOUR
┌──────────────────────────────────────────────────────────────────┐
│  09h00 – 09h45  │  THÉORIE       Qu'est-ce que l'IA générative ? │
│  09h45 – 11h00  │  EXERCICE      Créer vos 5 comptes + Notion    │
│  ─────────────  │  ─────────────────────────────────────────────  │
│  14h00 – 15h00  │  PRODUCTION    Dialogue avec Claude (1h)       │
│  15h00 – 15h30  │  PRODUCTION    Dialogue avec ChatGPT (30 min)  │
│  15h30 – 16h00  │  LIVRABLE      Fiche comparative (30 min)      │
└──────────────────────────────────────────────────────────────────┘
```

---

# PARTIE 1 — THÉORIE (45 minutes)
## Qu'est-ce que l'IA générative ? LLM, tokens, paramètres — sans jargon

---

### 1.1 — L'intelligence artificielle : une idée simple derrière un grand mot

L'intelligence artificielle n'est pas une magie. C'est un programme informatique qui a appris à reconnaître des schémas dans d'énormes quantités de textes, exactement comme un enfant apprend à parler en écoutant des milliers de phrases autour de lui.

**Une analogie pour comprendre :**

Imaginez un agent de centre d'appels qui aurait lu et mémorisé 10 millions de conversations téléphoniques, 500 000 manuels de formation, toute l'encyclopédie, tous les livres de management jamais écrits — en 30 langues. Quand vous lui posez une question, il ne "cherche" pas la réponse dans ses notes. Il a intégré tellement de connaissances qu'il peut construire une réponse cohérente, phrase après phrase.

C'est exactement ce que fait un LLM.

---

### 1.2 — Le LLM : votre nouveau collègue ultra-lu

**LLM** signifie **Large Language Model** — en français : *Grand Modèle de Langage*.

Ce que vous devez retenir :

> Un LLM est un programme qui a analysé des milliards de phrases pour apprendre comment les mots s'enchaînent, comment les idées se construisent, comment les humains communiquent.

**Les LLM que vous utiliserez :**

| Nom commercial | LLM derrière | Créé par |
|---|---|---|
| **Claude** | Claude 3.5 / 4 | Anthropic |
| **ChatGPT** | GPT-4o | OpenAI |
| **Gemini** | Gemini | Google |
| **Perplexity** | Plusieurs LLM | Perplexity AI |

**Différence clé pour votre usage :**
- Claude est reconnu pour sa rigueur, sa nuance et sa capacité à suivre des instructions complexes — idéal pour la formation et la qualité
- ChatGPT est très polyvalent, avec un écosystème d'outils très large
- Perplexity recherche sur internet en temps réel — idéal pour la veille

---

### 1.3 — L'IA générative : elle crée, elle ne copie pas

Il existe deux grandes familles d'IA :

**L'IA ancienne** (dite discriminante) : elle classe, reconnaît, prédit.
- Exemple : le filtre anti-spam de vos emails reconnaît un spam
- Exemple : la reconnaissance vocale de votre téléphone transcrit votre voix

**L'IA générative** : elle crée du contenu nouveau.
- Elle génère du texte, des images, des vidéos, du son
- Elle ne copie pas un document existant — elle produit quelque chose de nouveau à chaque fois
- C'est celle que vous allez utiliser

**Appliqué à votre métier :**

| Ce que vous faisiez | Ce que l'IA générative fait à votre place ou avec vous |
|---|---|
| Rédiger un plan de formation (2h) | Le générer en 5 minutes, vous l'affinez en 20 min |
| Créer une grille d'évaluation QA | La produire en 3 minutes selon vos critères |
| Rédiger un rapport d'audit | Le structurer automatiquement à partir de vos notes |
| Préparer un jeu de rôle formation | Simuler le client difficile en temps réel |
| Écrire un email de prospection | Proposer 3 versions adaptées à chaque cible |

---

### 1.4 — Les tokens : comment l'IA lit et compte

Quand vous écrivez à Claude, il ne lit pas les mots comme vous. Il découpe tout en **tokens** — des fragments de mots ou des mots entiers.

**Exemple concret :**

La phrase *"Bonjour, je suis formatrice en centres d'appels"* est découpée ainsi :
```
"Bon" | "jour" | "," | " je" | " suis" | " forma" | "trice" | " en" | " centres" | " d'" | "appels"
```
→ environ 11 tokens pour cette phrase

**Pourquoi c'est important pour vous :**
- Chaque LLM a une limite de tokens par conversation (sa "mémoire de travail")
- Claude peut traiter environ 200 000 tokens — soit un livre entier d'un coup
- Plus vous donnez de contexte, meilleure est la réponse — mais ne dépassez pas la limite
- **Règle pratique :** pour une session de travail normale, ne dépassez pas 50 pages de texte dans une même conversation

---

### 1.5 — Les paramètres : la personnalité de l'IA

Quand les ingénieurs créent un LLM, ils ajustent des milliards de **paramètres** — pensez à des millions de petits curseurs qui définissent comment l'IA répond.

**Paramètre clé que vous pouvez contrôler : la température**

Même si vous ne voyez pas ce curseur directement, vous pouvez l'influencer par vos instructions :

| Vous demandez... | L'IA répond avec... |
|---|---|
| "Sois précis et factuel" | Réponses rigoureuses, peu de créativité |
| "Sois créatif et innovant" | Réponses plus variées, plus surprenantes |
| "Propose 5 variantes différentes" | Maximum de diversité dans les réponses |

**Appliqué à votre métier :**
- Pour une grille d'audit ISO → demandez de la rigueur
- Pour un jeu de rôle formation → demandez de la créativité
- Pour un email de prospection → demandez 3 variantes de ton différentes

---

### 1.6 — Ce que l'IA ne sait PAS faire (et c'est votre valeur ajoutée)

L'IA générative est puissante mais elle a des limites claires. C'est là où votre expertise de 15 ans reste irremplaçable.

```
L'IA PEUT :                          L'IA NE PEUT PAS :
✅ Rédiger vite                       ❌ Valider si c'est juste pour VOTRE client
✅ Structurer logiquement             ❌ Connaître la culture de l'entreprise
✅ Proposer des variantes             ❌ Ressentir les dynamiques d'équipe
✅ Synthétiser de l'information       ❌ Juger si un agent est "prêt" en vrai
✅ Générer des cas pratiques          ❌ Adapter à la personnalité d'un apprenant
✅ Respecter une structure imposée    ❌ Prendre la responsabilité d'une décision
```

**La formule gagnante :**
> IA rapide + Expertise humaine juste = Résultat exceptionnel en peu de temps

---

### 1.7 — Résumé théorique en 5 points à retenir

```
1. L'IA générative CRÉE du contenu nouveau — elle ne copie pas
2. Le LLM a "lu" des milliards de textes et construit des réponses mot à mot
3. Les tokens = la monnaie d'échange entre vous et l'IA (donnez du contexte)
4. Les paramètres = la personnalité de l'IA (vous pouvez l'influencer)
5. Votre expertise valide — l'IA produit — ensemble vous êtes imbattables
```

---

# PARTIE 2 — EXERCICE (75 minutes)
## Créer vos 5 comptes et configurer votre journal Notion

---

### EXERCICE 1 — Créer vos 5 comptes (30 minutes)

Ouvrez votre navigateur (Chrome ou Edge recommandé). Créez les comptes dans cet ordre :

---

#### Compte 1 — CLAUDE (15 min maximum)
**Adresse :** claude.ai

**Étapes :**
1. Cliquez sur "Sign up"
2. Entrez votre adresse email
3. Choisissez un mot de passe solide (notez-le quelque part)
4. Vérifiez votre email et cliquez sur le lien de confirmation
5. Choisissez le plan **gratuit** pour commencer

**Premier test immédiat :** tapez exactement cette phrase dans Claude :
```
Bonjour Claude. Je suis formatrice en centres d'appels avec 15 ans d'expérience.
Je commence aujourd'hui à apprendre l'IA. Présente-toi en 5 lignes
et explique-moi comment tu peux m'aider dans mon métier.
```
Lisez la réponse. Ne cherchez pas à comprendre tout — observez simplement le ton, la précision, la pertinence.

---

#### Compte 2 — CHATGPT (5 min)
**Adresse :** chat.openai.com

**Étapes :**
1. Cliquez "Sign up"
2. Utilisez la même adresse email que Claude (plus simple à gérer)
3. Vérifiez votre email
4. Plan gratuit pour commencer

---

#### Compte 3 — PERPLEXITY (3 min)
**Adresse :** perplexity.ai

1. Cliquez "Sign up" ou "Continue with Google"
2. Compte gratuit suffisant pour débuter

---

#### Compte 4 — NOTION (3 min)
**Adresse :** notion.so

1. Cliquez "Get Notion free"
2. Inscrivez-vous avec votre email
3. Choisissez "Personal" quand on vous demande l'usage
4. Passez le tutoriel rapide

---

#### Compte 5 — CANVA (3 min)
**Adresse :** canva.com

1. Cliquez "S'inscrire"
2. Choisissez "Continuer avec Google" ou votre email
3. Sélectionnez "Éducation / Formation" comme usage
4. Plan gratuit pour l'instant

---

### EXERCICE 2 — Configurer votre Journal de Bord Notion (45 minutes)

Le journal de bord est votre outil de progression le plus important. Il vous permet de mesurer vos avancées, de garder vos meilleurs prompts, et de préparer votre portfolio.

---

#### Étape 1 — Créer votre espace de travail (10 min)

Dans Notion, cliquez sur **"+ Nouvelle page"** dans la barre latérale gauche.

Tapez comme titre : **TAMOU NEURAL PATH — Mon Parcours IA**

Appuyez sur Entrée. Choisissez l'icône 🧠 (cliquez sur "Ajouter une icône").

---

#### Étape 2 — Créer la page Journal de Bord (10 min)

Dans votre nouvelle page, cliquez à l'intérieur et tapez `/` pour ouvrir le menu.

Choisissez **"Base de données — Vue tableau"**.

Nommez la base de données : **Journal de Bord Quotidien**

Créez ces colonnes (cliquez sur le **+** pour ajouter chaque colonne) :

| Nom de la colonne | Type |
|---|---|
| DATE | Date |
| OUTIL UTILISÉ | Sélection (options : Claude / ChatGPT / Perplexity / Notion / Canva / Autre) |
| CE QUE J'AI PRODUIT | Texte |
| MEILLEUR PROMPT | Texte long |
| CE QUI N'A PAS MARCHÉ | Texte |
| DÉCOUVERTE DU JOUR | Texte |
| NIVEAU CONFIANCE /10 | Nombre |

---

#### Étape 3 — Remplir votre première entrée (10 min)

Cliquez sur **"+ Nouvelle ligne"** et remplissez votre première entrée du jour :

```
DATE : Aujourd'hui (cliquez sur la date)

OUTIL UTILISÉ : Claude

CE QUE J'AI PRODUIT : Première prise en main — test de présentation

MEILLEUR PROMPT : [copiez le prompt que vous avez utilisé avec Claude]

CE QUI N'A PAS MARCHÉ : [notez honnêtement ce que vous n'avez pas compris]

DÉCOUVERTE DU JOUR : L'IA répond différemment selon le niveau de détail qu'on lui donne

NIVEAU CONFIANCE /10 : [votre note honnête — 1 si vous êtes perdu(e), 5 si c'est clair]
```

---

#### Étape 4 — Créer votre page Portfolio (5 min)

Dans votre espace TAMOU NEURAL PATH, créez une deuxième page :

Titre : **Portfolio — Mes Livrables IA**

Structure à copier :
```
## PHASE 1 — Fondations (Semaines 1-2)
[ ] Livrable 1 — Fiche comparative Claude vs ChatGPT
[ ] Livrable 2 — Grille QA appels entrants
[ ] Livrable 3 — Charte utilisation responsable IA
[ ] Livrable 4 — Fiche synthèse ISO 9001 vulgarisée
[ ] Livrable 5 — ...

## PHASE 2 — Centres d'Appels (Semaines 3-5)
[ ] ...
```

---

#### Étape 5 — Créer votre Bibliothèque de Prompts (10 min)

Troisième page dans votre espace :

Titre : **Ma Bibliothèque de Prompts**

Créez une base de données avec ces colonnes :

| Colonne | Type |
|---|---|
| NOM DU PROMPT | Titre |
| THÈME | Sélection (Formation / Qualité / ISO / Management / Recrutement / Commercial) |
| LE PROMPT COMPLET | Texte long |
| RÉSULTAT OBTENU | Texte |
| NOTE /5 | Nombre |
| DATE | Date |

Vous alimenterez cette bibliothèque chaque jour. À la fin des 13 semaines, vous aurez 50+ prompts premium organisés — c'est un produit digital vendable.

---

# PARTIE 3 — PRODUCTION APRÈS-MIDI (2 heures)
## Dialogue libre et Fiche comparative Claude vs ChatGPT

---

### PRODUCTION 1 — Dialogue avec Claude (1 heure)

L'objectif n'est pas de produire un livrable parfait. C'est d'**observer comment Claude pense, répond et s'adapte** à votre domaine.

Voici une séquence de 6 questions à poser une par une. Laissez Claude répondre complètement avant de poser la suivante.

---

**Question 1 — Tester sa connaissance de votre domaine**
```
Je suis formatrice et ingénieure en formation, spécialisée dans les centres
d'appels entrants et sortants, avec 15 ans d'expérience. J'ai une certification
ISO 9001 v2015 délivrée par l'AFNOR.

Explique-moi en quoi ton utilisation peut transformer mon quotidien professionnel.
Sois très concret et cite des exemples directement liés aux centres d'appels
et à la formation professionnelle.
```

*Observez : Est-il précis ? Connaît-il les réalités du terrain ? Ses exemples sonnent-ils juste ?*

---

**Question 2 — Tester sa capacité à créer du contenu formation**
```
Génère un objectif pédagogique SMART pour une formation
"Gérer un client mécontent au téléphone" destinée à des agents
de centre d'appels avec 6 mois d'expérience.
```

*Observez : Connaît-il la méthode SMART ? L'objectif est-il réellement applicable ?*

---

**Question 3 — Tester sa rigueur ISO**
```
Dans le cadre de la norme ISO 9001 v2015, quels sont les éléments clés
du chapitre 8 (Réalisation des activités opérationnelles) qui s'appliquent
directement à un service de formation interne d'un centre d'appels ?
```

*Observez : Est-il rigoureux ? Fait-il des erreurs ? Utilisez vos connaissances pour vérifier.*

---

**Question 4 — Tester sa capacité à s'adapter à votre ton**
```
Rédige un message de motivation à envoyer à une équipe d'agents qui vient
de vivre une semaine difficile avec beaucoup de réclamations clients.
Le manager veut reconnaître l'effort, remotiver sans minimiser les difficultés,
et rappeler les valeurs de l'équipe. Ton chaleureux et sincère, 150 mots maximum.
```

*Observez : Le ton est-il juste ? Le message sonne-t-il humain ou artificiel ?*

---

**Question 5 — Tester la gestion d'une demande complexe**
```
Je dois concevoir un plan de formation de 3 jours pour des superviseurs
de centre d'appels sortants. Thèmes à couvrir : techniques de vente
consultative, management de la performance, gestion des conflits,
utilisation des KPIs. Public : 8 superviseurs, 3 à 5 ans d'expérience,
résistants aux formations théoriques.

Propose-moi un plan détaillé avec les horaires, les méthodes pédagogiques
adaptées à ce public et les modalités d'évaluation.
```

*Observez : Respecte-t-il vos contraintes ? La structure est-elle logique et réaliste ?*

---

**Question 6 — Tester l'itération (amélioration d'une réponse)**

Après avoir reçu la réponse à la question 5, répondez :
```
C'est bien mais le premier jour est trop chargé. Allège-le
en déplaçant la gestion des conflits au jour 2 après-midi.
Ajoute une activité brise-glace le matin du jour 1.
```

*Observez : Comprend-il vos retours ? Intègre-t-il les modifications sans tout refaire ?*

---

### PRODUCTION 2 — Dialogue avec ChatGPT (30 minutes)

Posez exactement les **mêmes 6 questions** dans ChatGPT. Copiez-collez vos questions une par une.

L'objectif est de comparer les réponses à périmètre identique.

---

### LIVRABLE — Fiche Comparative Claude vs ChatGPT (30 minutes)

Créez une nouvelle page dans Notion : **Livrable 1 — Fiche Comparative Claude vs ChatGPT**

Remplissez ce tableau en vous basant sur vos 6 échanges :

---

**TABLEAU DE COMPARAISON**

| Critère | Claude | ChatGPT | Mon choix |
|---|---|---|---|
| Connaissance des centres d'appels | /5 | /5 | |
| Qualité des contenus formation | /5 | /5 | |
| Rigueur sur les normes ISO | /5 | /5 | |
| Ton et qualité rédactionnelle | /5 | /5 | |
| Capacité à intégrer mes corrections | /5 | /5 | |
| Pertinence des exemples métier | /5 | /5 | |
| **TOTAL** | **/30** | **/30** | |

---

**MES OBSERVATIONS PERSONNELLES**

Répondez à ces 5 questions dans Notion (5 à 10 lignes au total) :

```
1. Quelle a été ma première surprise (positive ou négative) ?

2. Quelle réponse m'a le plus impressionnée et pourquoi ?

3. Quelle erreur ou imprécision ai-je détectée grâce à mon expertise ?
   (c'est votre valeur ajoutée — notez-la précieusement)

4. Pour quel type de tâche vais-je utiliser Claude en priorité ?

5. Pour quel type de tâche vais-je utiliser ChatGPT en priorité ?
```

---

**DÉCISION D'USAGE**

Complétez ces phrases :

```
Je vais utiliser CLAUDE principalement pour :
→ ____________________________________________

Je vais utiliser CHATGPT principalement pour :
→ ____________________________________________

Je vais utiliser PERPLEXITY principalement pour :
→ ____________________________________________
```

---

# JOURNAL DE BORD — FIN DE JOURNÉE

Remplissez votre entrée complète dans Notion avant de fermer votre ordinateur.

```
DATE : [Aujourd'hui]

OUTIL UTILISÉ : Claude + ChatGPT + Notion + Canva (création comptes)

CE QUE J'AI PRODUIT :
→ Fiche comparative Claude vs ChatGPT (Livrable 1 ✅)
→ Journal de bord configuré
→ Bibliothèque de prompts créée
→ Portfolio initialisé

MEILLEUR PROMPT :
→ [Copiez le prompt qui a donné la meilleure réponse de la journée]

CE QUI N'A PAS MARCHÉ :
→ [Soyez honnête — c'est votre meilleur outil de progression]

DÉCOUVERTE DU JOUR :
→ [La chose la plus surprenante que vous avez apprise]

NIVEAU CONFIANCE /10 :
→ [Votre note — sans jugement]
```

---

# BILAN DU JOUR 1

## Ce que vous avez accompli aujourd'hui

```
✅ Compris ce qu'est un LLM, un token, un paramètre — sans jargon
✅ Créé vos 5 comptes essentiels (Claude, ChatGPT, Perplexity, Notion, Canva)
✅ Configuré votre journal de bord Notion
✅ Produit votre LIVRABLE 1 — Fiche comparative Claude vs ChatGPT
✅ Identifié vos premières erreurs de l'IA grâce à votre expertise
✅ Posé les bases de votre bibliothèque de prompts
```

## Ce qui vous attend demain — Jour 2

**MARDI — La Formule Magique des Prompts**

Vous allez apprendre la méthode RCTF et l'appliquer sur 10 cas concrets de votre métier. À la fin de la journée, vous aurez votre première bibliothèque de 10 prompts professionnels prêts à l'emploi.

---

## Message pour vous ce soir

Vous venez de faire quelque chose que beaucoup de professionnels expérimentés n'ont pas encore fait : vous avez ouvert la porte de l'IA avec votre expertise comme boussole.

Vous avez probablement remarqué que l'IA fait des réponses impressionnantes — et que votre œil d'experte a immédiatement repéré ce qui cloche ou manque de profondeur.

**C'est exactement ça, votre valeur.** L'IA produit vite. Vous produisez juste.

Ensemble, vous produirez vite ET juste.

---

*TAMOU NEURAL PATH — Jour 1 sur 78 | Phase 1 sur 5*

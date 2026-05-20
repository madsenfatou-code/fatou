"""
Quiz Interactif — Module J2 Le Monde Associatif
Fidelis × SHY Performance
"""

import json
import sqlite3
import os
from datetime import datetime
from functools import wraps
from flask import (Flask, request, jsonify, session,
                   redirect, url_for, render_template_string)

app = Flask(__name__)
app.secret_key = "fidelis_shy_2026_quiz_secret"

DB_PATH = os.path.join(os.path.dirname(__file__), "quiz_results.db")
ADMIN_USER = "Tamou Eljerrari"
ADMIN_PASS = "1969"

QUESTIONS = [
    {
        "id": 1,
        "text": "Qu'est-ce que le Fundraising ?",
        "choices": [
            {"key": "a", "text": "Vente par correspondance"},
            {"key": "b", "text": "Collecte de fonds"},
            {"key": "c", "text": "Relation clients"},
        ],
        "correct": "b",
        "explanation": "Le Fundraising désigne la collecte de fonds — c'est l'activité centrale du fundraiser : solliciter des dons réguliers au nom d'une association."
    },
    {
        "id": 2,
        "text": "En quelle année fut créé le Comité de la Charte du Don en Confiance ?",
        "choices": [
            {"key": "a", "text": "1989"},
            {"key": "b", "text": "1996"},
            {"key": "c", "text": "1995"},
        ],
        "correct": "a",
        "explanation": "Le Comité de la Charte fut créé en 1989, puis renforcé après l'Affaire ARC en 1996. Il garantit la transparence et l'éthique des associations collectant des dons."
    },
    {
        "id": 3,
        "text": "Quelle autorité reconnaît les associations d'utilité publique ?",
        "choices": [
            {"key": "a", "text": "Le Comité de la Charte"},
            {"key": "b", "text": "Le Conseil d'État"},
            {"key": "c", "text": "Le ministère des Affaires Sociales"},
        ],
        "correct": "b",
        "explanation": "La Reconnaissance d'Utilité Publique (RUP) est accordée par décret en Conseil d'État — c'est le statut le plus élevé pour une association. UNICEF France en bénéficie."
    },
    {
        "id": 4,
        "text": "Combien y a-t-il d'associations en France ?",
        "choices": [
            {"key": "a", "text": "130 000"},
            {"key": "b", "text": "260 000"},
            {"key": "c", "text": "1 500 000"},
        ],
        "correct": "c",
        "explanation": "Il y a 1,5 million d'associations en France, dont 23 millions de bénévoles. La France est le 2e pays le plus associatif d'Europe, après la Suède."
    },
    {
        "id": 5,
        "text": "Chasser l'intrus — Le don ouvre droit à une déduction fiscale de :",
        "choices": [
            {"key": "a", "text": "66 %"},
            {"key": "b", "text": "75 %"},
            {"key": "c", "text": "100 %"},
        ],
        "correct": "c",
        "explanation": "L'intrus est 100 %. Les taux réels sont : 66 % pour les dons aux associations d'intérêt général, et 75 % (dans la limite de 513 €) pour les associations d'aide aux personnes en difficulté."
    },
    {
        "id": 6,
        "text": "Combien d'associations sont Reconnues d'Utilité Publique (RUP) en France ?",
        "choices": [
            {"key": "a", "text": "20 000"},
            {"key": "b", "text": "2 000"},
            {"key": "c", "text": "200 000"},
        ],
        "correct": "b",
        "explanation": "Environ 2 000 associations ont le statut RUP en France — soit moins de 0,2 % des associations. Parmi elles : UNICEF France, Croix-Rouge française, Restos du Cœur, MSF."
    },
    {
        "id": 7,
        "text": "Quelle est la limite maximale ouvrant droit à la déduction fiscale de 75 % ?",
        "choices": [
            {"key": "a", "text": "100 €"},
            {"key": "b", "text": "2 000 €"},
            {"key": "c", "text": "513 €"},
        ],
        "correct": "c",
        "explanation": "Le taux de 75 % s'applique jusqu'à 513 € de dons par an pour les associations d'aide aux personnes en difficulté. Au-delà, la réduction est de 66 %. Un vrai argument pour le donateur !"
    },
    {
        "id": 8,
        "text": "Quelle catastrophe naturelle réveille particulièrement la générosité des donateurs ?",
        "choices": [
            {"key": "a", "text": "Tremblement de terre au Mexique"},
            {"key": "b", "text": "Le Tsunami"},
            {"key": "c", "text": "Tremblement de terre en Haïti"},
        ],
        "correct": "c",
        "explanation": "Le séisme en Haïti (2010) a déclenché une mobilisation mondiale record. Les catastrophes majeures « proches » culturellement ou médiatisées amplifient fortement la générosité des donateurs."
    },
    {
        "id": 9,
        "text": "Comment les associations collectaient-elles des fonds avant 1990 ?",
        "choices": [
            {"key": "a", "text": "Les galas de bienfaisance"},
            {"key": "b", "text": "Les shows télévisés"},
            {"key": "c", "text": "Le mailing postal"},
        ],
        "correct": "c",
        "explanation": "Avant 1990, le mailing postal était le principal canal de collecte. L'émergence du phoning professionnel date des années 1990-1995. Aujourd'hui le phoning représente 55 % du marché."
    },
    {
        "id": 10,
        "text": "Qui est Jacques Crozemarie ?",
        "choices": [
            {"key": "a", "text": "Fondateur de l'UNICEF"},
            {"key": "b", "text": "Directeur de l'ARC impliqué dans le scandale 1996"},
            {"key": "c", "text": "Premier président du Comité de la Charte"},
        ],
        "correct": "b",
        "explanation": "Jacques Crozemarie était directeur de l'Association de Recherche sur le Cancer (ARC). En 1996, un scandale de détournements massifs de fonds donors éclate, provoquant une crise de confiance majeure dans le secteur associatif."
    },
]


def init_db():
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("""
        CREATE TABLE IF NOT EXISTS sessions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            participant_name TEXT NOT NULL,
            score INTEGER NOT NULL,
            total INTEGER NOT NULL,
            answers TEXT NOT NULL,
            completed_at TEXT NOT NULL
        )
    """)
    conn.commit()
    conn.close()


def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def admin_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if not session.get("admin"):
            return redirect(url_for("admin_login"))
        return f(*args, **kwargs)
    return decorated


# ─────────────────────────── QUIZ HTML ───────────────────────────────────────

QUIZ_HTML = """<!DOCTYPE html>
<html lang="fr">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Quiz — Le Monde Associatif | Fidelis × SHY Performance</title>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;600;700;800&display=swap" rel="stylesheet">
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Inter',sans-serif;background:linear-gradient(135deg,#003087 0%,#0057B8 60%,#00AEEF 100%);min-height:100vh;display:flex;align-items:center;justify-content:center;padding:20px}
.card{background:#fff;border-radius:20px;box-shadow:0 20px 60px rgba(0,0,0,.3);width:100%;max-width:700px;overflow:hidden}
.header{background:linear-gradient(135deg,#003087,#0057B8);padding:28px 32px;position:relative}
.header::after{content:'';position:absolute;bottom:0;left:0;right:0;height:4px;background:linear-gradient(90deg,#F7941D,#00AEEF)}
.logo{font-size:11px;font-weight:700;color:#00AEEF;letter-spacing:2px;text-transform:uppercase;margin-bottom:8px}
.header h1{font-size:22px;font-weight:800;color:#fff;line-height:1.3}
.header p{color:rgba(255,255,255,.7);font-size:13px;margin-top:6px}
.progress-bar-wrap{background:rgba(255,255,255,.15);border-radius:10px;height:8px;margin-top:18px;overflow:hidden}
.progress-bar{background:linear-gradient(90deg,#F7941D,#00AEEF);height:100%;border-radius:10px;transition:width .5s ease}
.body{padding:32px}
/* Landing */
.landing-icon{font-size:64px;text-align:center;margin-bottom:16px}
.landing h2{font-size:24px;font-weight:800;color:#003087;text-align:center;margin-bottom:8px}
.landing p{color:#555;text-align:center;margin-bottom:28px;font-size:14px;line-height:1.6}
.input-group{margin-bottom:20px}
.input-group label{display:block;font-size:13px;font-weight:600;color:#003087;margin-bottom:8px}
.input-group input{width:100%;padding:14px 18px;border:2px solid #dce3f0;border-radius:12px;font-size:15px;font-family:inherit;outline:none;transition:border-color .2s}
.input-group input:focus{border-color:#0057B8}
/* Question */
.q-number{font-size:11px;font-weight:700;color:#00AEEF;letter-spacing:2px;text-transform:uppercase;margin-bottom:12px}
.q-text{font-size:18px;font-weight:700;color:#003087;line-height:1.5;margin-bottom:24px}
.choices{display:flex;flex-direction:column;gap:12px}
.choice{display:flex;align-items:center;gap:14px;padding:14px 18px;border:2px solid #dce3f0;border-radius:14px;cursor:pointer;transition:all .2s;background:#fff}
.choice:hover:not(.disabled){border-color:#0057B8;background:#e8f0fb;transform:translateX(4px)}
.choice.disabled{cursor:default}
.choice .badge{width:34px;height:34px;border-radius:50%;border:2px solid #dce3f0;display:flex;align-items:center;justify-content:center;font-weight:700;font-size:13px;color:#555;flex-shrink:0;transition:all .2s}
.choice .label{font-size:14px;font-weight:500;color:#333;flex:1}
.choice.correct{border-color:#27AE60;background:#eafaf1}
.choice.correct .badge{background:#27AE60;border-color:#27AE60;color:#fff}
.choice.wrong{border-color:#E74C3C;background:#fdf0ef}
.choice.wrong .badge{background:#E74C3C;border-color:#E74C3C;color:#fff}
.choice.reveal-correct{border-color:#27AE60;background:#eafaf1}
.choice.reveal-correct .badge{background:#27AE60;border-color:#27AE60;color:#fff}
/* Feedback */
.feedback{margin-top:20px;padding:16px 20px;border-radius:14px;display:none;animation:fadeIn .4s ease}
.feedback.correct-fb{background:#eafaf1;border:2px solid #27AE60}
.feedback.wrong-fb{background:#fdf0ef;border:2px solid #E74C3C}
.feedback .fb-title{font-size:16px;font-weight:700;margin-bottom:6px}
.feedback.correct-fb .fb-title{color:#27AE60}
.feedback.wrong-fb .fb-title{color:#E74C3C}
.feedback .fb-score{font-size:13px;font-weight:600;margin-bottom:8px}
.feedback .fb-expl{font-size:13px;color:#555;line-height:1.6}
/* Score */
.score-wrap{text-align:center}
.score-circle{width:140px;height:140px;border-radius:50%;margin:0 auto 24px;display:flex;flex-direction:column;align-items:center;justify-content:center;font-size:42px;font-weight:800;color:#fff;background:linear-gradient(135deg,#003087,#0057B8);box-shadow:0 8px 24px rgba(0,48,135,.3)}
.score-circle span{font-size:14px;font-weight:600;opacity:.8}
.score-message{font-size:20px;font-weight:800;color:#003087;margin-bottom:12px}
.score-sub{font-size:14px;color:#555;line-height:1.6;margin-bottom:28px}
/* Review */
.review-item{padding:14px;border-radius:12px;margin-bottom:12px;border-left:4px solid}
.review-item.ok{background:#eafaf1;border-color:#27AE60}
.review-item.ko{background:#fdf0ef;border-color:#E74C3C}
.review-item .ri-q{font-size:13px;font-weight:600;color:#003087;margin-bottom:4px}
.review-item .ri-a{font-size:12px;color:#555}
/* Buttons */
.btn{width:100%;padding:15px;border:none;border-radius:14px;font-size:15px;font-weight:700;cursor:pointer;transition:all .2s;font-family:inherit;margin-top:18px}
.btn-primary{background:linear-gradient(135deg,#003087,#0057B8);color:#fff;box-shadow:0 4px 14px rgba(0,48,135,.3)}
.btn-primary:hover{transform:translateY(-2px);box-shadow:0 6px 20px rgba(0,48,135,.4)}
.btn-orange{background:linear-gradient(135deg,#F7611D,#F7941D);color:#fff;box-shadow:0 4px 14px rgba(247,148,29,.3)}
.btn-orange:hover{transform:translateY(-2px)}
.btn-sm{width:auto;padding:10px 22px;font-size:13px;margin-top:0}
.pts-badge{display:inline-block;padding:3px 10px;border-radius:20px;font-size:11px;font-weight:700;margin-left:8px}
.pts-1{background:#27AE60;color:#fff}
.pts-0{background:#E74C3C;color:#fff}
@keyframes fadeIn{from{opacity:0;transform:translateY(8px)}to{opacity:1;transform:translateY(0)}}
@keyframes popIn{0%{transform:scale(.8);opacity:0}80%{transform:scale(1.05)}100%{transform:scale(1);opacity:1}}
.pop{animation:popIn .5s ease}
.hidden{display:none}
</style>
</head>
<body>
<div class="card">
  <div class="header">
    <div class="logo">Fidelis × SHY Performance — Formation J2</div>
    <h1 id="hdr-title">Quiz — Le Monde Associatif & Humanitaire</h1>
    <p id="hdr-sub">10 questions · 1 point par bonne réponse · Score affiché à la fin</p>
    <div class="progress-bar-wrap"><div class="progress-bar" id="progress" style="width:0%"></div></div>
  </div>
  <div class="body">

    <!-- LANDING -->
    <div id="view-landing">
      <div class="landing-icon">🎯</div>
      <div class="landing">
        <h2>Bienvenue au Quiz J2 !</h2>
        <p>Testez vos connaissances sur le Monde Associatif & Humanitaire.<br>
           <strong>10 questions</strong> — 1 point par bonne réponse.<br>
           Bravo si vous faites ≥ 7/10 !</p>
      </div>
      <div class="input-group">
        <label>Votre prénom et nom *</label>
        <input type="text" id="participant-name" placeholder="ex : Jean Dupont" autocomplete="name">
      </div>
      <button class="btn btn-primary" onclick="startQuiz()">Commencer le Quiz →</button>
    </div>

    <!-- QUESTION -->
    <div id="view-question" class="hidden">
      <div class="q-number" id="q-number"></div>
      <div class="q-text" id="q-text"></div>
      <div class="choices" id="choices"></div>
      <div class="feedback" id="feedback">
        <div class="fb-title" id="fb-title"></div>
        <div class="fb-score" id="fb-score"></div>
        <div class="fb-expl" id="fb-expl"></div>
      </div>
      <button class="btn btn-primary hidden" id="btn-next" onclick="nextQuestion()">Question suivante →</button>
      <button class="btn btn-orange hidden" id="btn-finish" onclick="finishQuiz()">Voir mon score 🏆</button>
    </div>

    <!-- SCORE -->
    <div id="view-score" class="hidden score-wrap">
      <div class="score-circle pop" id="score-circle">
        <span id="score-val"></span>
        <span>/ 10</span>
      </div>
      <div class="score-message" id="score-msg"></div>
      <div class="score-sub" id="score-sub"></div>
      <details style="text-align:left;margin-bottom:16px">
        <summary style="cursor:pointer;font-size:13px;font-weight:600;color:#0057B8;margin-bottom:12px">
          📋 Voir le détail de mes réponses
        </summary>
        <div id="review-list"></div>
      </details>
      <button class="btn btn-primary" onclick="restartQuiz()">Recommencer le quiz 🔄</button>
    </div>

  </div>
</div>

<script>
const QUESTIONS = {{ questions | tojson }};
let state = {
  name: "",
  current: 0,
  answers: [],
  score: 0
};

function show(id) {
  ["view-landing","view-question","view-score"].forEach(v => {
    document.getElementById(v).classList.add("hidden");
  });
  document.getElementById(id).classList.remove("hidden");
}

function startQuiz() {
  const n = document.getElementById("participant-name").value.trim();
  if (!n) { alert("Merci d'entrer votre prénom et nom."); return; }
  state.name = n;
  state.current = 0;
  state.answers = [];
  state.score = 0;
  showQuestion();
  show("view-question");
}

function showQuestion() {
  const q = QUESTIONS[state.current];
  const total = QUESTIONS.length;
  document.getElementById("q-number").textContent =
    "Question " + (state.current + 1) + " / " + total;
  document.getElementById("q-text").textContent = q.text;
  document.getElementById("progress").style.width = (state.current / total * 100) + "%";

  const fb = document.getElementById("feedback");
  fb.style.display = "none";
  fb.className = "feedback";
  document.getElementById("btn-next").classList.add("hidden");
  document.getElementById("btn-finish").classList.add("hidden");

  const choicesDiv = document.getElementById("choices");
  choicesDiv.innerHTML = "";
  q.choices.forEach(ch => {
    const div = document.createElement("div");
    div.className = "choice";
    div.id = "choice-" + ch.key;
    div.innerHTML = `<div class="badge">${ch.key.toUpperCase()}</div><div class="label">${ch.text}</div>`;
    div.onclick = () => answer(ch.key);
    choicesDiv.appendChild(div);
  });
}

function answer(key) {
  const q = QUESTIONS[state.current];
  const isCorrect = key === q.correct;
  const pts = isCorrect ? 1 : 0;
  if (isCorrect) state.score++;

  state.answers.push({ qid: q.id, chosen: key, correct: q.correct, pts });

  // Disable all choices
  q.choices.forEach(ch => {
    const el = document.getElementById("choice-" + ch.key);
    el.classList.add("disabled");
    el.onclick = null;
    if (ch.key === q.correct) el.classList.add(isCorrect && ch.key === key ? "correct" : "reveal-correct");
    if (ch.key === key && !isCorrect) el.classList.add("wrong");
  });

  // Feedback
  const fb = document.getElementById("feedback");
  fb.className = "feedback " + (isCorrect ? "correct-fb" : "wrong-fb");
  document.getElementById("fb-title").textContent = isCorrect ? "✅  Bonne réponse !" : "❌  Pas tout à fait…";
  document.getElementById("fb-score").innerHTML =
    isCorrect
      ? '<span class="pts-badge pts-1">+1 point</span>'
      : '<span class="pts-badge pts-0">0 point</span>';
  document.getElementById("fb-expl").textContent = q.explanation;
  fb.style.display = "block";

  const isLast = state.current === QUESTIONS.length - 1;
  if (isLast) document.getElementById("btn-finish").classList.remove("hidden");
  else document.getElementById("btn-next").classList.remove("hidden");
}

function nextQuestion() {
  state.current++;
  showQuestion();
}

function finishQuiz() {
  document.getElementById("progress").style.width = "100%";

  const s = state.score;
  document.getElementById("score-val").textContent = s;
  document.getElementById("score-circle").className = "score-circle pop";
  void document.getElementById("score-circle").offsetWidth;
  document.getElementById("score-circle").className = "score-circle pop";

  let msg, sub;
  if (s >= 9) {
    msg = "🏆 Excellent ! Bravo !";
    sub = "Score parfait ou quasi-parfait ! Vous maîtrisez parfaitement le module J2. Vous êtes prêt(e) à convaincre avec confiance et légitimité.";
  } else if (s >= 7) {
    msg = "🎉 Bravo ! Très bon score !";
    sub = "Vous avez bien assimilé les points clés du module. Revoyez les questions ratées pour consolider vos connaissances.";
  } else if (s >= 5) {
    msg = "💪 Tu peux faire mieux !";
    sub = "Vous avez les bases ! Mais certains points importants méritent d'être revus. Consultez le détail ci-dessous et relisez vos supports de formation.";
  } else {
    msg = "📚 À retravailler !";
    sub = "Le score est insuffisant. Prenez le temps de relire le module J2 — Le Monde Associatif & Humanitaire — avant de recommencer.";
  }
  document.getElementById("score-msg").textContent = msg;
  document.getElementById("score-sub").textContent = sub;

  // Review list
  const rl = document.getElementById("review-list");
  rl.innerHTML = "";
  state.answers.forEach((a, i) => {
    const q = QUESTIONS[i];
    const ok = a.pts === 1;
    const chosenText = q.choices.find(c => c.key === a.chosen)?.text || "";
    const correctText = q.choices.find(c => c.key === a.correct)?.text || "";
    rl.innerHTML += `<div class="review-item ${ok ? 'ok' : 'ko'}">
      <div class="ri-q">Q${i+1}: ${q.text}</div>
      <div class="ri-a">
        ${ok
          ? "✅ Votre réponse : " + chosenText
          : "❌ Votre réponse : " + chosenText + " &nbsp;|&nbsp; ✅ Bonne réponse : " + correctText
        }
      </div>
    </div>`;
  });

  // Submit to server
  fetch("/submit", {
    method: "POST",
    headers: {"Content-Type": "application/json"},
    body: JSON.stringify({
      name: state.name,
      score: state.score,
      total: QUESTIONS.length,
      answers: state.answers
    })
  });

  show("view-score");
}

function restartQuiz() {
  state.current = 0;
  state.answers = [];
  state.score = 0;
  showQuestion();
  show("view-question");
}
</script>
</body>
</html>"""


ADMIN_LOGIN_HTML = """<!DOCTYPE html>
<html lang="fr">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Admin — Quiz Fidelis</title>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700;800&display=swap" rel="stylesheet">
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Inter',sans-serif;background:linear-gradient(135deg,#003087,#0057B8);min-height:100vh;display:flex;align-items:center;justify-content:center;padding:20px}
.card{background:#fff;border-radius:20px;box-shadow:0 20px 60px rgba(0,0,0,.3);width:100%;max-width:400px;overflow:hidden}
.header{background:linear-gradient(135deg,#003087,#0057B8);padding:24px 28px}
.header::after{content:'';display:block;height:4px;background:linear-gradient(90deg,#F7941D,#00AEEF);margin-top:16px;border-radius:2px}
.header h1{font-size:20px;font-weight:800;color:#fff}
.header p{color:rgba(255,255,255,.7);font-size:12px;margin-top:4px}
.body{padding:28px}
.icon{font-size:48px;text-align:center;margin-bottom:16px}
label{display:block;font-size:13px;font-weight:600;color:#003087;margin-bottom:6px}
input{width:100%;padding:13px 16px;border:2px solid #dce3f0;border-radius:12px;font-size:14px;font-family:inherit;outline:none;margin-bottom:16px}
input:focus{border-color:#0057B8}
button{width:100%;padding:14px;background:linear-gradient(135deg,#003087,#0057B8);color:#fff;border:none;border-radius:12px;font-size:15px;font-weight:700;cursor:pointer;font-family:inherit}
.error{color:#E74C3C;font-size:13px;margin-bottom:12px;padding:10px;background:#fdf0ef;border-radius:8px;border-left:3px solid #E74C3C}
</style>
</head>
<body>
<div class="card">
  <div class="header"><h1>🔐 Espace Admin</h1><p>Quiz — Fidelis × SHY Performance</p></div>
  <div class="body">
    <div class="icon">👤</div>
    {% if error %}<div class="error">{{ error }}</div>{% endif %}
    <form method="POST">
      <label>Identifiant admin</label>
      <input type="text" name="username" placeholder="Nom prénom" value="{{ username or '' }}" autocomplete="off">
      <label>Mot de passe</label>
      <input type="password" name="password" placeholder="••••">
      <button type="submit">Accéder au tableau de bord →</button>
    </form>
  </div>
</div>
</body>
</html>"""


ADMIN_DASHBOARD_HTML = """<!DOCTYPE html>
<html lang="fr">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Dashboard Admin — Quiz Fidelis</title>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;600;700;800&display=swap" rel="stylesheet">
<script src="https://cdn.jsdelivr.net/npm/chart.js@4/dist/chart.umd.min.js"></script>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Inter',sans-serif;background:#f0f4f9;min-height:100vh}
.topbar{background:linear-gradient(135deg,#003087,#0057B8);padding:16px 32px;display:flex;align-items:center;justify-content:space-between}
.topbar h1{font-size:18px;font-weight:800;color:#fff}
.topbar p{color:rgba(255,255,255,.7);font-size:12px}
.topbar .right{display:flex;gap:12px;align-items:center}
.pill{background:rgba(255,255,255,.15);color:#fff;padding:6px 14px;border-radius:20px;font-size:12px;font-weight:600}
.logout{background:#F7941D;color:#fff;padding:6px 14px;border-radius:20px;font-size:12px;font-weight:600;text-decoration:none}
.topbar::after{content:'';display:block;position:absolute;top:0;left:0;right:0;height:4px;background:linear-gradient(90deg,#F7941D,#00AEEF)}
.container{max-width:1200px;margin:0 auto;padding:28px 20px}
.stats-row{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:16px;margin-bottom:28px}
.stat-card{background:#fff;border-radius:16px;padding:20px 24px;box-shadow:0 2px 8px rgba(0,0,0,.07)}
.stat-card .val{font-size:36px;font-weight:800;color:#003087;line-height:1}
.stat-card .lbl{font-size:12px;color:#555;margin-top:6px;font-weight:500}
.stat-card.orange .val{color:#F7941D}
.stat-card.green .val{color:#27AE60}
.stat-card.cyan .val{color:#00AEEF}
.grid2{display:grid;grid-template-columns:1fr 1fr;gap:20px;margin-bottom:28px}
@media(max-width:700px){.grid2{grid-template-columns:1fr}}
.chart-card{background:#fff;border-radius:16px;padding:24px;box-shadow:0 2px 8px rgba(0,0,0,.07)}
.chart-card h3{font-size:14px;font-weight:700;color:#003087;margin-bottom:16px}
.table-card{background:#fff;border-radius:16px;padding:24px;box-shadow:0 2px 8px rgba(0,0,0,.07);margin-bottom:28px;overflow-x:auto}
.table-card h3{font-size:14px;font-weight:700;color:#003087;margin-bottom:16px}
table{width:100%;border-collapse:collapse}
th{text-align:left;padding:10px 14px;font-size:11px;font-weight:700;color:#555;text-transform:uppercase;letter-spacing:.5px;border-bottom:2px solid #dce3f0}
td{padding:12px 14px;font-size:13px;color:#333;border-bottom:1px solid #f0f4f9}
tr:last-child td{border-bottom:none}
tr:hover td{background:#f8faff}
.badge-score{display:inline-block;padding:3px 10px;border-radius:20px;font-size:12px;font-weight:700}
.score-high{background:#eafaf1;color:#27AE60}
.score-mid{background:#fff8e6;color:#F7941D}
.score-low{background:#fdf0ef;color:#E74C3C}
.refresh-btn{background:#003087;color:#fff;border:none;padding:8px 18px;border-radius:10px;font-size:13px;font-weight:600;cursor:pointer;font-family:inherit;margin-bottom:20px}
.empty{text-align:center;padding:40px;color:#999;font-size:14px}
</style>
</head>
<body>
<div class="topbar" style="position:relative">
  <div>
    <h1>📊 Tableau de Bord — Quiz J2</h1>
    <p>Administration — Fidelis × SHY Performance</p>
  </div>
  <div class="right">
    <span class="pill">👤 {{ admin_name }}</span>
    <a href="/admin/logout" class="logout">Déconnexion</a>
  </div>
</div>

<div class="container">
  <button class="refresh-btn" onclick="location.reload()">🔄 Actualiser</button>

  <div class="stats-row">
    <div class="stat-card">
      <div class="val">{{ stats.total_participants }}</div>
      <div class="lbl">Participants</div>
    </div>
    <div class="stat-card green">
      <div class="val">{{ stats.avg_score }}</div>
      <div class="lbl">Score moyen / 10</div>
    </div>
    <div class="stat-card orange">
      <div class="val">{{ stats.max_score }}</div>
      <div class="lbl">Meilleur score</div>
    </div>
    <div class="stat-card cyan">
      <div class="val">{{ stats.pct_success }}%</div>
      <div class="lbl">Réussite (≥ 7/10)</div>
    </div>
  </div>

  <div class="grid2">
    <div class="chart-card">
      <h3>📊 Scores par participant</h3>
      <canvas id="barChart"></canvas>
    </div>
    <div class="chart-card">
      <h3>🥧 Répartition des scores</h3>
      <canvas id="pieChart"></canvas>
    </div>
  </div>

  <div class="chart-card" style="margin-bottom:28px">
    <h3>❓ Taux de réussite par question</h3>
    <canvas id="questionChart" style="max-height:260px"></canvas>
  </div>

  <div class="table-card">
    <h3>📋 Détail des participants</h3>
    {% if rows %}
    <table>
      <tr>
        <th>#</th><th>Participant</th><th>Score</th><th>Résultat</th><th>Date/Heure</th>
      </tr>
      {% for row in rows %}
      <tr>
        <td style="color:#999">{{ loop.index }}</td>
        <td><strong>{{ row.participant_name }}</strong></td>
        <td>
          <span class="badge-score {% if row.score >= 7 %}score-high{% elif row.score >= 5 %}score-mid{% else %}score-low{% endif %}">
            {{ row.score }} / {{ row.total }}
          </span>
        </td>
        <td>{% if row.score >= 7 %}✅ Réussi{% else %}⚠️ À améliorer{% endif %}</td>
        <td style="color:#777;font-size:12px">{{ row.completed_at }}</td>
      </tr>
      {% endfor %}
    </table>
    {% else %}
    <div class="empty">Aucun participant n'a encore complété le quiz.</div>
    {% endif %}
  </div>
</div>

<script>
const chartData = {{ chart_data | tojson }};

if (chartData.names.length > 0) {
  new Chart(document.getElementById("barChart"), {
    type: "bar",
    data: {
      labels: chartData.names,
      datasets: [{
        label: "Score /10",
        data: chartData.scores,
        backgroundColor: chartData.scores.map(s =>
          s >= 7 ? "rgba(39,174,96,.7)" : s >= 5 ? "rgba(247,148,29,.7)" : "rgba(231,76,60,.7)"
        ),
        borderRadius: 8,
        borderSkipped: false
      }]
    },
    options: {
      responsive: true,
      scales: {y:{min:0,max:10,ticks:{stepSize:1}},x:{ticks:{font:{size:11}}}},
      plugins: {legend: {display: false}}
    }
  });

  new Chart(document.getElementById("pieChart"), {
    type: "doughnut",
    data: {
      labels: ["≥ 7/10 (Réussi)", "5-6/10 (Moyen)", "< 5/10 (À revoir)"],
      datasets: [{
        data: [chartData.high, chartData.mid, chartData.low],
        backgroundColor: ["#27AE60","#F7941D","#E74C3C"],
        borderWidth: 0
      }]
    },
    options: {responsive: true, plugins: {legend: {position: "bottom"}}}
  });
}

if (chartData.q_labels.length > 0) {
  new Chart(document.getElementById("questionChart"), {
    type: "bar",
    data: {
      labels: chartData.q_labels,
      datasets: [{
        label: "% bonnes réponses",
        data: chartData.q_pcts,
        backgroundColor: chartData.q_pcts.map(p =>
          p >= 70 ? "rgba(39,174,96,.7)" : p >= 50 ? "rgba(247,148,29,.7)" : "rgba(231,76,60,.7)"
        ),
        borderRadius: 6,
        borderSkipped: false
      }]
    },
    options: {
      responsive: true,
      scales: {y:{min:0,max:100,ticks:{callback: v => v + "%"}},x:{ticks:{font:{size:11}}}},
      plugins: {legend: {display: false}}
    }
  });
}
</script>
</body>
</html>"""


# ─────────────────────────── ROUTES ──────────────────────────────────────────

@app.route("/")
def index():
    name = request.args.get("nom", "")
    return render_template_string(QUIZ_HTML, questions=QUESTIONS, prefill_name=name)


@app.route("/submit", methods=["POST"])
def submit():
    data = request.get_json()
    conn = get_db()
    conn.execute(
        "INSERT INTO sessions (participant_name, score, total, answers, completed_at) "
        "VALUES (?, ?, ?, ?, ?)",
        (
            data["name"],
            data["score"],
            data["total"],
            json.dumps(data["answers"]),
            datetime.now().strftime("%d/%m/%Y %H:%M"),
        )
    )
    conn.commit()
    conn.close()
    return jsonify({"ok": True})


@app.route("/admin", methods=["GET", "POST"])
def admin_login():
    if session.get("admin"):
        return redirect(url_for("admin_dashboard"))
    error = None
    username = ""
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "").strip()
        if username == ADMIN_USER and password == ADMIN_PASS:
            session["admin"] = True
            session["admin_name"] = username
            return redirect(url_for("admin_dashboard"))
        error = "Identifiant ou mot de passe incorrect."
    return render_template_string(ADMIN_LOGIN_HTML, error=error, username=username)


@app.route("/admin/dashboard")
@admin_required
def admin_dashboard():
    conn = get_db()
    rows = conn.execute(
        "SELECT * FROM sessions ORDER BY completed_at DESC"
    ).fetchall()
    conn.close()
    rows = [dict(r) for r in rows]

    total = len(rows)
    avg = round(sum(r["score"] for r in rows) / total, 1) if total else 0
    maxi = max((r["score"] for r in rows), default=0)
    success = sum(1 for r in rows if r["score"] >= 7)
    pct_success = round(success / total * 100) if total else 0

    stats = {
        "total_participants": total,
        "avg_score": avg,
        "max_score": maxi,
        "pct_success": pct_success,
    }

    names = [r["participant_name"] for r in rows]
    scores = [r["score"] for r in rows]
    high = sum(1 for s in scores if s >= 7)
    mid = sum(1 for s in scores if 5 <= s < 7)
    low = sum(1 for s in scores if s < 5)

    # Per-question stats
    q_correct = [0] * len(QUESTIONS)
    for r in rows:
        try:
            answers = json.loads(r["answers"])
            for a in answers:
                idx = a["qid"] - 1
                if 0 <= idx < len(QUESTIONS):
                    q_correct[idx] += a.get("pts", 0)
        except Exception:
            pass
    q_labels = [f"Q{i+1}" for i in range(len(QUESTIONS))]
    q_pcts = [round(q_correct[i] / total * 100) if total else 0 for i in range(len(QUESTIONS))]

    chart_data = {
        "names": names,
        "scores": scores,
        "high": high,
        "mid": mid,
        "low": low,
        "q_labels": q_labels,
        "q_pcts": q_pcts,
    }

    return render_template_string(
        ADMIN_DASHBOARD_HTML,
        stats=stats,
        rows=rows,
        chart_data=chart_data,
        admin_name=session.get("admin_name", "Admin"),
    )


@app.route("/admin/logout")
def admin_logout():
    session.clear()
    return redirect(url_for("admin_login"))


@app.route("/ping")
def ping():
    return "OK"


if __name__ == "__main__":
    init_db()
    print("=" * 60)
    print("  Quiz Fidelis × SHY Performance — Module J2")
    print("=" * 60)
    print("  Quiz participants : http://0.0.0.0:5000/")
    print("  Lien avec nom     : http://0.0.0.0:5000/?nom=Prenom+Nom")
    print("  Admin dashboard   : http://0.0.0.0:5000/admin")
    print("  Admin login       : Tamou Eljerrari / 1969")
    print("=" * 60)
    app.run(host="0.0.0.0", port=5000, debug=False)

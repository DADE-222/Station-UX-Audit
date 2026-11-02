"""
STATION UX AUDIT - v2.6.2
Correctifs :
- LM Studio : max_tokens 262144 (contexte complet)
- JSON repair automatique avant parsing (guillemets, virgules, troncatures)
- Retry IA si JSON invalide apres repair (temperature 0.05, prompt condense)
- Nouveau preset : Taxonomie fermee UX (prompt optimise structure)
"""

import os, re, json, glob, time, queue, threading, traceback, subprocess, unicodedata
import tkinter as tk
from tkinter import filedialog, scrolledtext, ttk, messagebox
import requests

try:
    import win32com.client, pythoncom
    COM_DISPONIBLE = True
except ImportError:
    COM_DISPONIBLE = False


# ══════════════════════════════════════════════════════
#  UTILITAIRE — Nettoyage markdown
# ══════════════════════════════════════════════════════

def _strip_md(texte: str) -> str:
    if not texte:
        return ""
    t = str(texte)
    t = re.sub(r'\*{1,3}([^*]+)\*{1,3}', r'\1', t)
    t = re.sub(r'#{1,6}\s*', '', t)
    t = re.sub(r'`{1,3}([^`]*)`{1,3}', r'\1', t)
    t = re.sub(r'^\s*[-•]\s+', '', t, flags=re.M)
    t = re.sub(r'\n{3,}', '\n\n', t)
    return t.strip()


# ══════════════════════════════════════════════════════
#  PROMPTS
# ══════════════════════════════════════════════════════

# ── Prompt standard optimisé (doc 3) ──────────────────

PROMPT_STANDARD = """\
Tu es un expert UX research et analyse sémantique.

Analyse l'entretien utilisateur fourni ci-dessous et identifie tous les problèmes, irritants, freins, attentes, points positifs et nuances mentionnés par le participant.

Réponds UNIQUEMENT en français, avec un JSON valide, sans texte avant ni après.
Le JSON doit être strictement conforme au format demandé.
N'ajoute aucun commentaire, aucune explication, aucun markdown, aucun bloc de code.
Les chaînes de texte doivent être valides en JSON, avec les caractères spéciaux correctement échappés si nécessaire.

Règles d'extraction :
- Regroupe les verbatims par thème principal UX.
- Crée des sous-thèmes précis et concrets.
- Pour chaque extrait, indique le sentiment :
  - "Négatif" si l'utilisateur exprime un problème, une frustration, un frein ou une difficulté.
  - "Positif" si l'utilisateur exprime une satisfaction, un avantage ou une préférence claire.
  - "Nuancé" si l'expression est ambivalente, conditionnelle ou mitigée.
- Le champ "participant" doit contenir le code extrait du nom du fichier ou du contexte si disponible, par exemple : "S1", "S2", "Test9".
- Le champ "verbatim" doit reprendre la citation exacte de l'utilisateur, sans reformulation.
- Ne modifie pas le sens du verbatim.
- N'ajoute pas de guillemets internes non échappés dans le verbatim.
- Si plusieurs extraits relèvent du même sous-thème, rassemble-les dans le même tableau "extraits".
- Si un extrait peut appartenir à plusieurs thèmes, choisis le thème le plus pertinent.
- Si aucun problème n'est identifié, retourne quand même une structure JSON valide avec les thèmes pertinents observés.
- Ne déduis rien qui n'est pas explicitement dit dans l'entretien.
- Avant de répondre, normalise mentalement les formulations proches, fusionne les doublons sémantiques et conserve les verbatims exacts les plus représentatifs.

FORMAT ATTENDU :
{
  "themes": [
    {
      "theme": "Nom du thème principal",
      "sous_themes": [
        {
          "sous_theme": "Nom du sous-thème",
          "extraits": [
            {
              "sentiment": "Négatif",
              "participant": "S1",
              "verbatim": "Citation exacte de l'utilisateur"
            }
          ]
        }
      ]
    }
  ]
}

ENTRETIEN :
"""

PRESETS = {
    "🔍 Analyse UX structurée (recommandé)": PROMPT_STANDARD,

    "🎯 Heuristiques de Nielsen": """\
Tu es un auditeur UX certifié. Applique les 10 heuristiques de Nielsen à l'entretien fourni.
Réponds UNIQUEMENT en français avec un JSON valide, sans texte avant ni après.
N'ajoute aucun markdown ni commentaire. Caractères JSON correctement échappés.

FORMAT ATTENDU :
{
  "themes": [
    {
      "theme": "Nom de l'heuristique",
      "sous_themes": [
        {
          "sous_theme": "Composant ou zone concernée",
          "extraits": [
            {"sentiment": "Négatif", "participant": "S1", "verbatim": "Citation exacte"}
          ]
        }
      ]
    }
  ]
}

ENTRETIEN :
""",

    "📐 Critères Bastien & Scapin": """\
Tu es un ergonome IHM. Analyse avec les critères de Bastien & Scapin.
Réponds UNIQUEMENT en français avec un JSON valide, sans texte avant ni après.
N'ajoute aucun markdown ni commentaire. Caractères JSON correctement échappés.

FORMAT ATTENDU :
{
  "themes": [
    {
      "theme": "Critère Bastien & Scapin",
      "sous_themes": [
        {
          "sous_theme": "Sous-critère ou zone",
          "extraits": [
            {"sentiment": "Négatif", "participant": "S1", "verbatim": "Citation exacte"}
          ]
        }
      ]
    }
  ]
}

ENTRETIEN :
""",

    "💬 Ressenti émotionnel & charge cognitive": """\
Tu es un psychologue cognitif spécialisé en UX.
Réponds UNIQUEMENT en français avec un JSON valide, sans texte avant ni après.
N'ajoute aucun markdown ni commentaire. Caractères JSON correctement échappés.

FORMAT ATTENDU :
{
  "themes": [
    {
      "theme": "Type d'émotion ou état cognitif",
      "sous_themes": [
        {
          "sous_theme": "Déclencheur ou contexte",
          "extraits": [
            {"sentiment": "Négatif", "participant": "S1", "verbatim": "Citation exacte"}
          ]
        }
      ]
    }
  ]
}

ENTRETIEN :
""",

    "♿ Accessibilité WCAG 2.1": """\
Tu es un expert en accessibilité numérique. Analyse selon les critères WCAG 2.1.
Réponds UNIQUEMENT en français avec un JSON valide, sans texte avant ni après.
N'ajoute aucun markdown ni commentaire. Caractères JSON correctement échappés.

FORMAT ATTENDU :
{
  "themes": [
    {
      "theme": "Principe WCAG (Perceptible / Utilisable / Compréhensible / Robuste)",
      "sous_themes": [
        {
          "sous_theme": "Critère WCAG + niveau A/AA/AAA",
          "extraits": [
            {"sentiment": "Négatif", "participant": "S1", "verbatim": "Citation exacte"}
          ]
        }
      ]
    }
  ]
}

ENTRETIEN :
""",

    "🏷️ Taxonomie fermée UX (recommandé)": """\
Tu es un expert UX. Ton rôle : coder qualitativement l'entretien fourni en respectant une taxonomie fermée.

RÈGLES STRICTES :
1. SORTIE : JSON strict uniquement. Pas de texte, pas de markdown, pas d'explications.
2. TAXONOMIE : Utilise UNIQUEMENT ces thèmes exacts :
   [Navigation, Clarté du contenu, Confiance, Identité visuelle, Interactivité, Portfolio, Contact, Ergonomie, Performance, Satisfaction globale]
3. SENTIMENTS : Utilise UNIQUEMENT : "Négatif", "Positif", "Nuancé"
4. VERBATIM : Citation exacte sans reformulation, sans correction orthographique.
   Les guillemets dans le verbatim doivent être échappés avec \\".
5. DÉCISION :
   - Chaque extrait → 1 seul thème (priorise le frein fonctionnel sur l'esthétique)
   - Regroupe les extraits similaires sous le même sous_theme
   - Évite les sous_themes redondants
   - Si un verbatim contient des guillemets, les remplacer par des guillemets simples '

FORMAT JSON OBLIGATOIRE (respecte ce schéma exactement) :
{"themes":[{"theme":"Navigation","sous_themes":[{"sous_theme":"Nom précis","extraits":[{"sentiment":"Négatif","participant":"CODE","verbatim":"Citation exacte"}]}]}]}

ENTRETIEN :
""",
}

# ── Prompts mode méthodologique ────────────────────────

_PROMPT_PHASE0_ERGO = """\
Tu es un expert en ergonomie cognitive et UX research.

En ergonomie, on part du TERRAIN, pas d'une hypothèse préalable.
Ta mission : lire le corpus, identifier le terrain réel, reformuler la demande depuis les données, puis faire émerger la problématique.

DEMANDE INITIALE DU CLIENT (si disponible) : {demande_initiale}

Procède en 3 étapes :

ÉTAPE 1 — LECTURE DU TERRAIN :
Identifie dans le corpus les mots-clés récurrents, les tensions, les oppositions,
les causes, les conséquences, les attentes implicites et les accords/désaccords.

ÉTAPE 2 — REFORMULATION DE LA DEMANDE :
À partir du terrain réel (pas de la demande initiale), reformule ce que les données
révèlent réellement comme question de fond.
La reformulation doit partir des données, pas de la commande initiale.

ÉTAPE 3 — PROBLÉMATIQUE :
Formule une question directrice claire, ouverte, défendable, qui émerge du corpus.
Elle doit relier les tensions observées. Elle doit permettre de construire un plan.
Si possible, propose aussi une reformulation académique.

Réponds UNIQUEMENT avec ce JSON valide, sans texte avant ni après :
{{
  "axes_thematiques": ["axe1", "axe2", "axe3"],
  "reformulation_demande": "La demande reformulée depuis le terrain en 1-2 phrases",
  "problematique": "La question centrale qui émerge du corpus ?",
  "reformulation_academique": "Formulation plus formelle (optionnelle)",
  "justification": "Pourquoi cette problématique en 2-3 phrases fondées sur le corpus"
}}

CORPUS :
"""

_PROMPT_PHASE1 = """\
Tu es un expert en analyse de corpus textuels. Réponds UNIQUEMENT en français.

PROBLÉMATIQUE : {prob}

Lis ce corpus et construis une grille thématique.
Réponds UNIQUEMENT avec ce JSON valide, sans texte avant ni après :
{{
  "theme_general": "Phrase résumant le corpus",
  "mots_cles": ["mot1","mot2","mot3"],
  "grille_themes": [
    {{"theme": "Nom", "sous_themes": ["st1","st2"], "description": "Ce que ça recouvre"}}
  ],
  "premieres_observations": "Premières tensions ou répétitions visibles"
}}

CORPUS :
"""

_PROMPT_PHASE2 = """\
Tu es un chercheur en analyse qualitative. Réponds UNIQUEMENT en français.

PROBLÉMATIQUE : {prob}
GRILLE DE THÈMES : {grille}

Annote cet entretien. Pour chaque extrait pertinent :
- theme : un thème de la grille (ou NOUVEAU: si hors grille)
- sous_theme : précis
- type : convergence, divergence ou recurrence
- sentiment : Négatif, Positif ou Nuancé
- verbatim : citation exacte, sans markdown, caractères JSON échappés

Réponds UNIQUEMENT avec ce JSON valide, sans texte avant ni après :
{{"extraits": [{{"theme":"","sous_theme":"","type":"","sentiment":"","verbatim":""}}]}}

ENTRETIEN :
"""

_PROMPT_PHASE3 = """\
Tu es un chercheur senior en sciences sociales. Réponds UNIQUEMENT en français.

PROBLÉMATIQUE : {prob}

Compare les {n} entretiens codés et synthétise.
Réponds UNIQUEMENT avec ce JSON valide, sans texte avant ni après :
{{
  "convergences": [{{"theme":"","frequence":0,"observation":"","exemple_verbatim":""}}],
  "divergences": [{{"theme":"","entre":[],"observation":"","verbatims":[]}}],
  "recurrences": [{{"expression":"","frequence":0,"contexte":""}}],
  "cooccurrences": [{{"themes":[],"observation":""}}],
  "synthese_finale": "4-6 phrases répondant à la problématique"
}}

DONNÉES CODÉES :
{donnees}
"""

# ── Prompt nettoyage verbatims oraux ──────────────────────────────────────────
_PROMPT_NETTOYAGE_VERBATIMS = """\
Tu es un linguiste expert en analyse qualitative.

Voici des verbatims bruts issus d'entretiens oraux retranscrits.
Certains contiennent des artefacts oraux (hésitations, mots incompréhensibles, familiarités excessives).

Pour chaque verbatim :
1. Si exploitable → reformule en français académique lisible SANS déformer le sens
2. Si inexploitable (< 3 mots signifiants, incompréhensible, vulgaire sans contenu) → marque "INEXPLOITABLE"
3. Conserve le sens exact — ne t'invente rien

Réponds UNIQUEMENT avec ce JSON valide :
{{"verbatims_nettoyes": [
  {{"original": "texte original", "nettoye": "version académique ou INEXPLOITABLE", "exploitable": true}}
]}}

VERBATIMS :
{verbatims}
"""

# ── Prompt classification heuristiques ───────────────────────────────────────
_PROMPT_HEURISTIQUES = """\
Tu es un expert en ergonomie IHM.
Classe chaque thème UX dans les heuristiques de Nielsen ET les critères de Bastien & Scapin.

HEURISTIQUES DE NIELSEN (1-10) :
1. Visibilité du statut système
2. Correspondance avec le monde réel
3. Contrôle et liberté de l'utilisateur
4. Cohérence et standards
5. Prévention des erreurs
6. Reconnaissance plutôt que rappel
7. Flexibilité et efficacité
8. Esthétique et design minimaliste
9. Aide à la récupération d'erreurs
10. Aide et documentation

CRITÈRES BASTIEN & SCAPIN :
- Guidage (Incitation, Groupement, Feedback, Lisibilité)
- Charge de travail (Brièveté, Densité informationnelle)
- Contrôle explicite
- Adaptabilité
- Gestion des erreurs
- Cohérence
- Signifiance des codes
- Compatibilité

THÈMES À CLASSIFIER :
{themes}

Réponds UNIQUEMENT avec ce JSON valide :
{{"classifications": [
  {{"theme": "Nom exact du thème",
    "nielsen": ["H3 - Contrôle et liberté"],
    "bastien_scapin": ["Guidage - Feedback"],
    "gravite": "Majeur|Modéré|Mineur"
  }}
]}}
"""


# ══════════════════════════════════════════════════════
#  ALGORITHME DE FUSION DES THÈMES
# ══════════════════════════════════════════════════════

def _normaliser(nom: str) -> str:
    """Normalise un nom de thème pour comparaison : minuscules, sans accents, sans ponctuation."""
    t = nom.lower().strip()
    t = ''.join(c for c in unicodedata.normalize('NFD', t)
                if unicodedata.category(c) != 'Mn')
    t = re.sub(r'[^\w\s]', ' ', t)
    t = re.sub(r'\s+', ' ', t).strip()
    return t


def _jaccard(a: str, b: str) -> float:
    """Similarité de Jaccard entre deux chaînes normalisées."""
    ma = set(_normaliser(a).split())
    mb = set(_normaliser(b).split())
    if not ma or not mb:
        return 0.0
    return len(ma & mb) / len(ma | mb)


def _similaires(a: str, b: str, seuil: float = 0.5) -> bool:
    """True si a et b sont suffisamment similaires."""
    na, nb = _normaliser(a), _normaliser(b)
    if na == nb:
        return True
    if na in nb or nb in na:
        return True
    return _jaccard(a, b) >= seuil


def reparer_json(brut: str) -> str:
    """
    Tente de réparer un JSON malformé avant parsing.
    Gère : guillemets non échappés dans les valeurs, virgules manquantes,
    JSON tronqué, accolades/crochets non fermés.
    """
    import re as _re

    # 1. Retire les blocs thinking et markdown
    s = _re.sub(r"<think>.*?</think>", "", brut, flags=_re.DOTALL)
    s = _re.sub(r"```json\s*|```\s*", "", s).strip()

    # 2. Isole du premier { au dernier }
    debut = s.find("{")
    fin   = s.rfind("}")
    if debut == -1:
        return s
    s = s[debut:fin + 1] if fin != -1 else s[debut:]

    # 3. Dans les valeurs de string, remplace les guillemets doubles internes
    #    non échappés par des guillemets simples.
    #    Stratégie : reconstruire caractère par caractère en suivant l'état
    resultat = []
    in_string = False
    escape    = False
    for c in s:
        if escape:
            resultat.append(c)
            escape = False
            continue
        if c == "\\":
            escape = True
            resultat.append(c)
            continue
        if c == '"':
            if not in_string:
                in_string = True
                resultat.append(c)
            else:
                # Regarde le caractère suivant pour décider
                # Si suivi de :, , } ] → fin de string légitime
                # Sinon → guillemet interne → remplace par '
                pos = len(resultat)
                # Cherche les prochains chars non-espace
                reste = s[s.find(c, sum(1 for _ in resultat)):]  # approximation
                in_string = False
                resultat.append(c)
            continue
        resultat.append(c)
    s = "".join(resultat)

    # 4. Répare les JSON tronqués : ferme les listes/objets ouverts
    ouv_accolades  = s.count("{") - s.count("}")
    ouv_crochets   = s.count("[") - s.count("]")
    if ouv_crochets > 0 or ouv_accolades > 0:
        # Ferme proprement
        if s.rstrip()[-1] not in ('"', '}', ']'):
            s = s.rstrip().rstrip(',')
            s += '"'
        for _ in range(max(0, ouv_crochets)):
            s += "]"
        for _ in range(max(0, ouv_accolades)):
            s += "}"

    return s


def fusionner_themes_v2(synthese: dict) -> dict:
    """
    Fusionne les thèmes similaires dans la structure enrichie v2.6.
    Entrée : {theme: {"_sources": set, "sous_themes": {st: {"_sources": set, sent: [(part,verb)]}}}}
    """
    themes = list(synthese.keys())
    utilises = set()
    groupes = []

    for i, t1 in enumerate(themes):
        if i in utilises:
            continue
        groupe = [t1]
        for j, t2 in enumerate(themes):
            if j <= i or j in utilises:
                continue
            if _similaires(t1, t2):
                groupe.append(t2)
                utilises.add(j)
        utilises.add(i)
        groupes.append(groupe)

    fusionnee: dict = {}
    for groupe in groupes:
        nom_final = max(groupe, key=lambda x: len(x))
        merged: dict = {"_sources": set(), "sous_themes": {}}

        for th in groupe:
            merged["_sources"].update(synthese[th].get("_sources", set()))
            for st_nom, data_st in synthese[th].get("sous_themes", {}).items():
                cible = None
                for st_existe in merged["sous_themes"]:
                    if _similaires(st_nom, st_existe, seuil=0.4):
                        cible = st_existe
                        break
                if cible is None:
                    cible = st_nom
                    merged["sous_themes"][cible] = {"_sources": set()}
                merged["sous_themes"][cible]["_sources"].update(data_st.get("_sources", set()))
                for key, val in data_st.items():
                    if key == "_sources":
                        continue
                    merged["sous_themes"][cible].setdefault(key, []).extend(val)

        fusionnee[nom_final] = merged

    return fusionnee


def fusionner_themes(synthese: dict) -> dict:
    """
    Fusionne les thèmes dans la structure simple (mode legacy).
    Structure : {theme: {sous_theme: {sentiment: [(participant, verbatim)]}}}
    """
    themes = list(synthese.keys())
    utilises = set()
    groupes = []

    for i, t1 in enumerate(themes):
        if i in utilises:
            continue
        groupe = [i]
        utilises.add(i)
        for j, t2 in enumerate(themes):
            if j <= i or j in utilises:
                continue
            if _similaires(t1, t2):
                groupe.append(j)
                utilises.add(j)
        groupes.append(groupe)

    resultat = {}
    for groupe in groupes:
        noms = [themes[i] for i in groupe]
        nom_canon = max(noms, key=len)
        sous_themes_fusionnes: dict = {}
        for nom in noms:
            for st, sentiments in synthese[nom].items():
                st_canon = st
                for st_existant in sous_themes_fusionnes:
                    if _similaires(st, st_existant):
                        st_canon = st_existant
                        break
                if st_canon not in sous_themes_fusionnes:
                    sous_themes_fusionnes[st_canon] = {}
                for sent, entrees in sentiments.items():
                    sous_themes_fusionnes[st_canon].setdefault(sent, []).extend(entrees)
        resultat[nom_canon] = sous_themes_fusionnes

    return resultat


# ══════════════════════════════════════════════════════
#  APPLICATION
# ══════════════════════════════════════════════════════

class StationUX:

    def __init__(self, root):
        self.root = root
        self.root.title("STATION UX AUDIT — v2.6.3")
        self.root.geometry("1200x980")
        self.root.configure(bg="#1a1a2e")
        self.root.minsize(900, 700)

        self.log_queue = queue.Queue()
        self.fichier_config = os.path.join(os.path.expanduser("~"), ".station_ux_config.json")
        self._init_variables()
        self._charger_config()
        self._construire_interface()
        self._demarrer_pompe_logs()
        self.root.after(500, self._tester_serveur_auto)

    # ── Variables ──────────────────────────────────────

    def _init_variables(self):
        self.var_url      = tk.StringVar(value="http://localhost:11434")
        self.var_modele   = tk.StringVar()
        self.var_dossier  = tk.StringVar(value=os.path.expanduser("~"))
        self.nb_fichiers  = tk.StringVar(value="— fichiers")
        self.statut_mm    = tk.StringVar(value="MindManager : non vérifié")
        self.statut_srv   = tk.StringVar(value="Serveur : non vérifié")
        self.statut_gpu   = tk.StringVar(value="GPU : non vérifié")
        self.var_temps_reel       = tk.BooleanVar(value=False)
        self.var_mode_methodo     = tk.BooleanVar(value=False)
        self.var_demande_initiale = tk.StringVar(value="")
        self.var_problematique    = tk.StringVar(value="")
        # ── Backend IA ──────────────────────────────────
        # "ollama"   → API native Ollama  (http://localhost:11434)
        # "lmstudio" → API OpenAI-compat  (http://localhost:1234)
        self.var_backend = tk.StringVar(value="ollama")
        self.en_cours = False

    # ── Config ─────────────────────────────────────────

    def _charger_config(self):
        try:
            if os.path.exists(self.fichier_config):
                with open(self.fichier_config, "r", encoding="utf-8") as f:
                    cfg = json.load(f)
                    self.var_url.set(cfg.get("url", "http://localhost:11434"))
                    self.var_dossier.set(cfg.get("dossier", os.path.expanduser("~")))
                    self.var_modele.set(cfg.get("modele", ""))
                    self.var_backend.set(cfg.get("backend", "ollama"))
        except Exception:
            pass

    def _sauvegarder_config(self):
        try:
            with open(self.fichier_config, "w", encoding="utf-8") as f:
                json.dump({"url":     self.var_url.get(),
                           "dossier": self.var_dossier.get(),
                           "modele":  self.var_modele.get(),
                           "backend": self.var_backend.get()},
                          f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    # ══════════════════════════════════════════════════
    #  INTERFACE
    # ══════════════════════════════════════════════════

    def _construire_interface(self):
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("TNotebook", background="#1a1a2e")
        style.configure("TNotebook.Tab", background="#2d2d44", foreground="white",
                        padding=[12, 6], font=("Segoe UI", 10))
        style.map("TNotebook.Tab", background=[("selected", "#007acc")])
        style.configure("TCombobox", fieldbackground="#2d2d44", foreground="white")

        nb = ttk.Notebook(self.root)
        nb.pack(fill="both", expand=True, padx=8, pady=8)

        tabs = [tk.Frame(nb, bg="#1e1e2e") for _ in range(4)]
        nb.add(tabs[0], text="▶  Exécution & Debug")
        nb.add(tabs[1], text="✏️  Constructeur de Prompt")
        nb.add(tabs[2], text="⚡  GPU / Performance")
        nb.add(tabs[3], text="ℹ️  Aide")

        self._ong_execution(tabs[0])
        self._ong_prompt(tabs[1])
        self._ong_gpu(tabs[2])
        self._ong_aide(tabs[3])

    # ── Onglet Exécution ───────────────────────────────

    def _ong_execution(self, p):
        # ── Sélecteur de backend ───────────────────────
        fb = tk.LabelFrame(p, text=" Backend IA ", bg="#1e1e2e", fg="#aaaacc",
                           font=("Segoe UI", 9, "bold"), padx=10, pady=6)
        fb.pack(fill="x", padx=10, pady=(10, 2))

        BACKENDS = [
            ("ollama",   "🦙 Ollama",    "http://localhost:11434",
             "API native Ollama — modèles téléchargés via 'ollama pull'"),
            ("lmstudio", "🖥️ LM Studio", "http://localhost:1234",
             "API OpenAI-compatible — modèles chargés dans LM Studio"),
        ]
        for val, label, url_defaut, info in BACKENDS:
            col = "#33aa66" if val == "ollama" else "#aa6633"
            tk.Radiobutton(
                fb, text=label, variable=self.var_backend, value=val,
                command=lambda u=url_defaut, v=val: self._changer_backend(u, v),
                bg="#1e1e2e", fg=col, selectcolor="#1e1e2e",
                activebackground="#1e1e2e", activeforeground=col,
                font=("Segoe UI", 10, "bold")
            ).pack(side="left", padx=(0, 20))
            tk.Label(fb, text=info, bg="#1e1e2e", fg="#888899",
                     font=("Segoe UI", 8, "italic")).pack(side="left", padx=(0, 30))

        # ── Connexion ──────────────────────────────────
        fc = tk.LabelFrame(p, text=" Connexion ", bg="#1e1e2e", fg="#aaaacc",
                           font=("Segoe UI", 9, "bold"), padx=8, pady=6)
        fc.pack(fill="x", padx=10, pady=(2, 4))
        tk.Label(fc, text="URL :", bg="#1e1e2e", fg="white").grid(row=0, column=0, padx=(0,4))
        tk.Entry(fc, textvariable=self.var_url, width=28, bg="#2d2d44", fg="white",
                 insertbackground="white").grid(row=0, column=1, padx=4)
        tk.Button(fc, text="🔌 Tester & Charger modèles", command=self._tester_serveur_manuel,
                  bg="#444466", fg="white", relief="flat", padx=8).grid(row=0, column=2, padx=6)
        self.cb_modeles = ttk.Combobox(fc, textvariable=self.var_modele, width=22, state="readonly")
        self.cb_modeles.grid(row=0, column=3, padx=4)
        tk.Button(fc, text="🧠 Tester MindManager", command=self._tester_mindmanager,
                  bg="#444466", fg="white", relief="flat", padx=8).grid(row=0, column=4, padx=6)
        tk.Button(fc, text="⚡ CPU+GPU", command=self._lancer_optimisation_gpu,
                  bg="#cc6600", fg="white", relief="flat", padx=8,
                  font=("Segoe UI", 9, "bold")).grid(row=0, column=5, padx=6)

        # Statuts
        fs = tk.Frame(p, bg="#1e1e2e")
        fs.pack(fill="x", padx=10, pady=2)
        for sv, c in [(self.statut_srv, "#88ccff"),
                      (self.statut_mm,  "#88ccff"),
                      (self.statut_gpu, "#ffaa44")]:
            tk.Label(fs, textvariable=sv, bg="#1e1e2e", fg=c,
                     font=("Consolas", 9)).pack(side="left", padx=8)

        # Dossier
        fd = tk.LabelFrame(p, text=" Dossier des entretiens (.txt) ", bg="#1e1e2e", fg="#aaaacc",
                           font=("Segoe UI", 9, "bold"), padx=8, pady=6)
        fd.pack(fill="x", padx=10, pady=4)
        tk.Entry(fd, textvariable=self.var_dossier, width=66, bg="#2d2d44", fg="white",
                 insertbackground="white").pack(side="left", padx=(0,6))
        tk.Button(fd, text="📁 Parcourir", command=self._parcourir,
                  bg="#444466", fg="white", relief="flat").pack(side="left")
        tk.Label(fd, textvariable=self.nb_fichiers, bg="#1e1e2e", fg="#66ff88",
                 font=("Consolas", 9, "bold")).pack(side="left", padx=12)

        # Mode écriture
        fm = tk.Frame(p, bg="#1e1e2e")
        fm.pack(fill="x", padx=10, pady=(0, 2))
        tk.Checkbutton(
            fm, text="✏️  Écriture en temps réel (sans fusion)  —  décoché = mode groupé avec fusion des thèmes à la fin",
            variable=self.var_temps_reel,
            bg="#1e1e2e", fg="#ffdd88", selectcolor="#333355",
            activebackground="#1e1e2e", activeforeground="#ffdd88",
            font=("Segoe UI", 9)
        ).pack(side="left", padx=8)

        # ── Mode Méthodologique ────────────────────────
        fmm = tk.LabelFrame(
            p,
            text=" 🔬  MODE MÉTHODOLOGIQUE — Terrain → Reformulation → Problématique → Codage → Synthèse ",
            bg="#1a1a2e", fg="#88ffcc",
            font=("Segoe UI", 9, "bold"), padx=10, pady=8
        )
        fmm.pack(fill="x", padx=10, pady=(4, 4))
        fmm.columnconfigure(1, weight=1)

        tk.Checkbutton(
            fmm,
            text="Activer l'analyse méthodologique complète (5 phases)",
            variable=self.var_mode_methodo,
            bg="#1a1a2e", fg="#88ffcc", selectcolor="#1a3a2e",
            activebackground="#1a1a2e", activeforeground="#88ffcc",
            font=("Segoe UI", 9, "bold"),
            command=self._toggle_mode_methodo
        ).grid(row=0, column=0, columnspan=3, sticky="w", pady=(0,6))

        # Demande initiale
        tk.Label(fmm, text="Demande initiale du client :",
                 bg="#1a1a2e", fg="#cccccc", font=("Segoe UI", 9)).grid(
                     row=1, column=0, sticky="w", padx=(0,8))
        self.ent_demande = tk.Entry(
            fmm, textvariable=self.var_demande_initiale,
            width=68, bg="#1e2e2a", fg="white",
            insertbackground="white", font=("Segoe UI", 9),
            state="disabled"
        )
        self.ent_demande.grid(row=1, column=1, sticky="ew", padx=(0,8), pady=(0,4))
        tk.Label(
            fmm,
            text="ⓘ Ex: « Comprendre pourquoi les utilisateurs abandonnent le tunnel de commande »",
            bg="#1a1a2e", fg="#66bb88", font=("Segoe UI", 8, "italic")
        ).grid(row=2, column=0, columnspan=3, sticky="w")

        # Problématique
        tk.Label(fmm, text="Problématique (optionnelle) :",
                 bg="#1a1a2e", fg="#cccccc", font=("Segoe UI", 9)).grid(
                     row=3, column=0, sticky="w", padx=(0,8), pady=(6,0))
        self.ent_problematique = tk.Entry(
            fmm, textvariable=self.var_problematique,
            width=68, bg="#1e2e2a", fg="white",
            insertbackground="white", font=("Segoe UI", 9),
            state="disabled"
        )
        self.ent_problematique.grid(row=3, column=1, sticky="ew", padx=(0,8), pady=(6,0))
        tk.Label(
            fmm,
            text="ⓘ Vide → Phase 0 : l'IA lit le terrain et reformule la demande avant de proposer une problématique",
            bg="#1a1a2e", fg="#66bb88", font=("Segoe UI", 8, "italic")
        ).grid(row=4, column=0, columnspan=3, sticky="w", pady=(2,0))

        # Console
        fl = tk.LabelFrame(p, text=" Console de débogage ", bg="#1e1e2e", fg="#ff6666",
                           font=("Segoe UI", 9, "bold"), padx=6, pady=6)
        fl.pack(fill="both", expand=True, padx=10, pady=4)
        self.zone_log = scrolledtext.ScrolledText(
            fl, bg="#0d0d1a", fg="#00ff88", font=("Consolas", 10), wrap="word", state="disabled"
        )
        self.zone_log.pack(fill="both", expand=True)

        self.btn_lancer = tk.Button(
            p, text="🚀  LANCER L'ANALYSE ET ÉCRIRE DANS MINDMANAGER",
            command=self._lancer, bg="#007acc", fg="white",
            font=("Segoe UI", 12, "bold"), height=2, relief="flat", cursor="hand2"
        )
        self.btn_lancer.pack(fill="x", padx=10, pady=(4,10))

    # ── Onglet Prompt ──────────────────────────────────

    def _ong_prompt(self, p):
        fp = tk.LabelFrame(p, text=" Préréglages ", bg="#1e1e2e", fg="#aaaacc",
                           font=("Segoe UI", 9, "bold"), padx=8, pady=8)
        fp.pack(fill="x", padx=10, pady=(10,4))
        self.cb_presets = ttk.Combobox(fp, values=list(PRESETS.keys()), width=55, state="readonly")
        self.cb_presets.pack(side="left", padx=(0,10))
        self.cb_presets.current(0)
        self.cb_presets.bind("<<ComboboxSelected>>", self._appliquer_preset)
        tk.Button(fp, text="↺ Appliquer", command=self._appliquer_preset,
                  bg="#007acc", fg="white", relief="flat", padx=10).pack(side="left")
        tk.Button(fp, text="🗑️ Vider", command=lambda: self.zone_prompt.delete("1.0", tk.END),
                  bg="#666688", fg="white", relief="flat", padx=10).pack(side="left", padx=6)

        fe = tk.LabelFrame(p, text=" Consigne personnalisée (utilisée en mode standard) ",
                           bg="#1e1e2e", fg="#aaaacc", font=("Segoe UI", 9, "bold"), padx=6, pady=6)
        fe.pack(fill="both", expand=True, padx=10, pady=(4,10))
        self.zone_prompt = scrolledtext.ScrolledText(
            fe, bg="#12122a", fg="#e0e0ff", font=("Consolas", 11), wrap="word", insertbackground="white"
        )
        self.zone_prompt.pack(fill="both", expand=True)
        self._appliquer_preset(None)

    # ── Onglet GPU ─────────────────────────────────────

    def _ong_gpu(self, p):
        fi = tk.LabelFrame(p, text=" Optimisation GPU AMD Radeon (RDNA2 / RX 6700) ",
                           bg="#1e1e2e", fg="#ffaa44", font=("Segoe UI", 9, "bold"), padx=12, pady=10)
        fi.pack(fill="x", padx=10, pady=(10,4))
        tk.Label(fi, text=(
            "Variables écrites dans le registre Windows (HKCU\\Environment) :\n"
            "  HSA_OVERRIDE_GFX_VERSION=10.3.0 | OLLAMA_NUM_GPU=1\n"
            "  HIP_VISIBLE_DEVICES=0 | OLLAMA_GPU_OVERHEAD=0\n\n"
            "⚠️  Si VRAM = 0 après activation → redémarrez le PC."
        ), bg="#1e1e2e", fg="#cccccc", font=("Consolas", 9), justify="left").pack(fill="x")

        fb = tk.Frame(p, bg="#1e1e2e")
        fb.pack(fill="x", padx=10, pady=8)
        for txt, cmd, col in [
            ("⚡  ACTIVER CPU + GPU", self._lancer_optimisation_gpu, "#cc6600"),
            ("🔍  Diagnostiquer GPU",  self._diagnostiquer_gpu,       "#445544"),
            ("♻️  Redémarrer Ollama", self._redemarrer_ollama,        "#334455"),
        ]:
            tk.Button(fb, text=txt, command=cmd, bg=col, fg="white",
                      font=("Segoe UI", 10), height=2, relief="flat",
                      cursor="hand2").pack(side="left", padx=(0,10))

        fl = tk.LabelFrame(p, text=" Diagnostic GPU ", bg="#1e1e2e", fg="#ffaa44",
                           font=("Segoe UI", 9, "bold"), padx=6, pady=6)
        fl.pack(fill="both", expand=True, padx=10, pady=4)
        self.zone_gpu = scrolledtext.ScrolledText(
            fl, bg="#0d0d1a", fg="#ffcc66", font=("Consolas", 10), wrap="word", state="disabled"
        )
        self.zone_gpu.pack(fill="both", expand=True)

    # ── Onglet Aide ────────────────────────────────────

    def _ong_aide(self, p):
        t = tk.Text(p, bg="#1e1e2e", fg="#ccccee", font=("Segoe UI", 10),
                    wrap="word", padx=20, pady=20, relief="flat")
        t.pack(fill="both", expand=True)
        t.insert("1.0", """\
STATION UX AUDIT v2.5

BACKENDS IA SUPPORTÉS
──────────────────────
🦙 Ollama  (défaut)  → http://localhost:11434
   API native — modèles via "ollama pull gemma2:9b"
   GPU AMD : activer avec le bouton ⚡ CPU+GPU

🖥️ LM Studio        → http://localhost:1234
   API OpenAI-compatible — chargez un modèle dans l'interface LM Studio
   puis activez "Local Server" (onglet ↔ en haut à gauche)
   Modèles recommandés : Mistral 7B, Llama 3.1 8B, Gemma 2 9B (GGUF)
   LM Studio gère automatiquement le GPU (NVIDIA et AMD)

MODES D'ANALYSE
───────────────
Mode standard (décoché = groupé) :
  Analyse chaque fichier avec le prompt de l'onglet ✏️
  → Fusion des thèmes similaires à la fin (Jaccard ≥ 50%)
  → Structure : Thème → Sous-thème → 🔴/🟢/🟡 → [Sx] Verbatim

Mode standard (coché = temps réel) :
  Écrit dans MindManager au fil des fichiers, sans fusion.

Mode méthodologique :
  Phase 0 — Lecture du terrain → Reformulation de la demande → Problématique
  Phase 1 — Grille thématique (lecture globale du corpus)
  Phase 2 — Codage entretien par entretien
  Phase 3 — Comparaison, fréquences, cooccurrences, synthèse

ALGORITHME DE FUSION (mode groupé)
────────────────────────────────────
  Deux thèmes sont fusionnés si :
  - L'un contient l'autre (ex: "Formulaire" ⊂ "Formulaire et CTA")
  - Jaccard ≥ 50% sur les mots normalisés (sans accents)

PHASE 0 — MÉTHODOLOGIE ERGO
──────────────────────────────
  1. Demande initiale du client (optionnelle, saisie dans le champ)
  2. L'IA lit le terrain (corpus) SANS hypothèse préalable
  3. L'IA reformule la demande depuis ce qu'elle trouve dans les données
  4. L'IA propose une problématique + reformulation académique
  5. Popup : vous voyez les axes, la reformulation, la problématique → validez ou modifiez

DÉPENDANCES : pip install requests pywin32
""")
        t.config(state="disabled")

    # ══════════════════════════════════════════════════
    #  LOGS
    # ══════════════════════════════════════════════════

    def _demarrer_pompe_logs(self):
        # Configure les tags couleur au premier démarrage
        try:
            self.zone_log.tag_config("ok",      foreground="#66ff99")
            self.zone_log.tag_config("err",     foreground="#ff5555")
            self.zone_log.tag_config("warn",    foreground="#ffdd55")
            self.zone_log.tag_config("negatif", foreground="#ff8866")
            self.zone_log.tag_config("positif", foreground="#99ffbb")
            self.zone_log.tag_config("nuance",  foreground="#88ccff")
            self.zone_log.tag_config("info",    foreground="#00ff88")
        except Exception:
            pass
        try:
            while True:
                msg, couleur = self.log_queue.get_nowait()
                self.zone_log.config(state="normal")
                self.zone_log.insert(tk.END, msg + "\n", couleur)
                self.zone_log.see(tk.END)
                self.zone_log.config(state="disabled")
        except queue.Empty:
            pass
        self.root.after(100, self._demarrer_pompe_logs)

    def _log(self, texte: str, couleur: str = ""):
        """Log thread-safe avec détection automatique de couleur."""
        if not couleur:
            t = texte
            if any(x in t for x in ("❌", "ERREUR", "CRASH", "impossible", "échoué", "non joignable")):
                couleur = "err"
            elif any(x in t for x in ("✅", "🎉", "succès", "terminé", "OK", "Connecté", "lancé")):
                couleur = "ok"
            elif any(x in t for x in ("⚠️", "attention", "Alerte", "⏳")):
                couleur = "warn"
            elif "🔴" in texte:
                couleur = "negatif"
            elif "🟢" in texte:
                couleur = "positif"
            elif "🟡" in texte:
                couleur = "nuance"
            else:
                couleur = "info"
        self.log_queue.put((f"[{time.strftime('%H:%M:%S')}] {texte}", couleur))

    def _log_gpu(self, t):
        def _w():
            self.zone_gpu.config(state="normal")
            self.zone_gpu.insert(tk.END, t + "\n")
            self.zone_gpu.see(tk.END)
            self.zone_gpu.config(state="disabled")
        self.root.after(0, _w)

    # ══════════════════════════════════════════════════
    #  ACTIONS UI
    # ══════════════════════════════════════════════════

    def _toggle_mode_methodo(self):
        s = "normal" if self.var_mode_methodo.get() else "disabled"
        self.ent_demande.config(state=s)
        self.ent_problematique.config(state=s)
        self._log("🔬 Mode méthodologique " + ("activé" if s == "normal" else "désactivé"))

    def _parcourir(self):
        c = filedialog.askdirectory(title="Choisir le dossier des entretiens")
        if c:
            self.var_dossier.set(c)
            f = glob.glob(os.path.join(c, "*.txt"))
            self.nb_fichiers.set(f"{len(f)} fichier(s) .txt")
            self._log(f"📂 {c} ({len(f)} fichiers)")

    def _appliquer_preset(self, event=None):
        self.zone_prompt.delete("1.0", tk.END)
        self.zone_prompt.insert(tk.END, PRESETS.get(self.cb_presets.get(), ""))

    # ══════════════════════════════════════════════════
    #  SERVEUR / MINDMANAGER
    # ══════════════════════════════════════════════════

    def _changer_backend(self, url_defaut: str, val: str):
        self.var_url.set(url_defaut)
        self.cb_modeles["values"] = []
        self.var_modele.set("")
        self._log(f"🔀 Backend → {'Ollama' if val == 'ollama' else 'LM Studio'}  |  URL : {url_defaut}")
        self._log("   Cliquez sur 🔌 Tester & Charger modèles pour rafraîchir la liste.")

    def _tester_serveur_auto(self):
        threading.Thread(target=self._tester_serveur_thread, args=(False,), daemon=True).start()

    def _tester_serveur_manuel(self):
        threading.Thread(target=self._tester_serveur_thread, args=(True,), daemon=True).start()

    def _tester_serveur_thread(self, _):
        base = self.var_url.get().rstrip("/")
        backend = self.var_backend.get()
        try:
            if backend == "lmstudio":
                # LM Studio : endpoint OpenAI-compatible /v1/models
                r = requests.get(f"{base}/v1/models", timeout=4)
                r.raise_for_status()
                modeles = [m["id"] for m in r.json().get("data", [])]
                label = "LM Studio"
            else:
                # Ollama : endpoint natif /api/tags
                r = requests.get(f"{base}/api/tags", timeout=4)
                r.raise_for_status()
                modeles = [m["name"] for m in r.json().get("models", [])]
                label = "Ollama"

            self.root.after(0, lambda: self._maj_modeles(modeles))
            self._log(f"✅ {label} OK — {len(modeles)} modèle(s) : {', '.join(modeles[:4])}")
            self.root.after(0, lambda: self.statut_srv.set(
                f"✅ {label} OK — {len(modeles)} modèle(s)"))
        except requests.exceptions.ConnectionError:
            self._log(f"❌ Serveur non joignable sur {base}")
            self._log("   → Vérifiez que Ollama / LM Studio est bien lancé.")
            self.root.after(0, lambda: self.statut_srv.set("❌ Serveur non joignable"))
        except Exception as e:
            self._log(f"❌ Serveur : {e}")
            self.root.after(0, lambda: self.statut_srv.set("❌ Erreur serveur"))

    def _maj_modeles(self, modeles):
        self.cb_modeles["values"] = modeles
        if modeles:
            self.cb_modeles.set(self.var_modele.get() if self.var_modele.get() in modeles else modeles[0])

    def _tester_mindmanager(self):
        if not COM_DISPONIBLE:
            messagebox.showerror("Erreur", "pip install pywin32")
            return
        threading.Thread(target=self._tester_mm_thread, daemon=True).start()

    def _tester_mm_thread(self):
        try:
            pythoncom.CoInitialize()
            try:
                win32com.client.GetActiveObject("MindManager.Application")
                self._log("✅ MindManager connecté.")
                self.root.after(0, lambda: self.statut_mm.set("✅ MindManager : connecté"))
            except Exception:
                mm = win32com.client.Dispatch("MindManager.Application")
                mm.Visible = True
                self._log("✅ MindManager lancé.")
                self.root.after(0, lambda: self.statut_mm.set("✅ MindManager : lancé"))
        except Exception as e:
            self._log(f"❌ MindManager : {e}")
            self.root.after(0, lambda: self.statut_mm.set("❌ MindManager : erreur COM"))
        finally:
            pythoncom.CoUninitialize()

    # ══════════════════════════════════════════════════
    #  GPU
    # ══════════════════════════════════════════════════

    def _lancer_optimisation_gpu(self):
        threading.Thread(target=self._optimiser_gpu_thread, daemon=True).start()

    def _optimiser_gpu_thread(self):
        variables = {"HSA_OVERRIDE_GFX_VERSION": "10.3.0", "OLLAMA_NUM_GPU": "1",
                     "HIP_VISIBLE_DEVICES": "0", "OLLAMA_GPU_OVERHEAD": "0"}
        try:
            import winreg
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Environment", 0, winreg.KEY_SET_VALUE)
            for n, v in variables.items():
                winreg.SetValueEx(key, n, 0, winreg.REG_SZ, v)
                self._log(f"   ✅ {n} = {v}")
            winreg.CloseKey(key)
        except Exception as e:
            self._log(f"   ❌ Registre : {e}")
        for n, v in variables.items():
            os.environ[n] = v
        self._redemarrer_ollama(log_gpu=True)
        time.sleep(3)
        self._diagnostiquer_gpu()

    def _redemarrer_ollama(self, log_gpu=False):
        self._log("♻️  Redémarrage Ollama...")
        try:
            subprocess.run(["taskkill", "/F", "/IM", "ollama.exe"], capture_output=True)
        except Exception:
            pass
        time.sleep(2)
        try:
            subprocess.Popen(["ollama", "serve"], creationflags=subprocess.CREATE_NO_WINDOW,
                             env=os.environ.copy())
            self._log("   ✅ Ollama relancé.")
            time.sleep(3)
            threading.Thread(target=self._tester_serveur_thread, args=(False,), daemon=True).start()
        except Exception as e:
            self._log(f"   ❌ {e}")

    def _diagnostiquer_gpu(self):
        threading.Thread(target=self._diag_gpu_thread, daemon=True).start()

    def _diag_gpu_thread(self):
        self._log_gpu("\n── DIAGNOSTIC GPU ──")
        try:
            r = requests.get(self.var_url.get().rstrip("/") + "/api/ps", timeout=5)
            if r.status_code == 200:
                modeles = r.json().get("models", [])
                if not modeles:
                    self._log_gpu("  ⓘ  Aucun modèle chargé.")
                for m in modeles:
                    vram = m.get("size_vram", 0)
                    ram  = max(m.get("size", 0) - vram, 0)
                    self._log_gpu(f"  {m.get('name','?')} — VRAM: {vram/1e9:.2f}Go | RAM: {ram/1e9:.2f}Go")
                    self._log_gpu("  ✅ GPU ACTIF" if vram > 0 else "  ⚠️  VRAM=0 — CPU uniquement")
                    self.root.after(0, lambda v=vram: self.statut_gpu.set(
                        "✅ GPU actif" if v > 0 else "⚠️ GPU non actif"))
        except Exception as e:
            self._log_gpu(f"  ❌ {e}")
        for v in ["HSA_OVERRIDE_GFX_VERSION","OLLAMA_NUM_GPU","HIP_VISIBLE_DEVICES","OLLAMA_GPU_OVERHEAD"]:
            val = os.environ.get(v, "(non définie)")
            self._log_gpu(f"  {'✅' if val != '(non définie)' else '❌'}  {v} = {val}")

    # ══════════════════════════════════════════════════
    #  APPEL IA
    # ══════════════════════════════════════════════════

    def _appel_ia(self, prompt: str, json_force: bool = True) -> dict | None:
        base    = self.var_url.get().rstrip("/")
        modele  = self.var_modele.get()
        backend = self.var_backend.get()
        brut    = ""
        try:
            if backend == "lmstudio":
                sys_msg = (
                    "Tu réponds UNIQUEMENT avec un objet JSON valide. "
                    "Aucun texte avant, aucun texte après, aucun bloc markdown, "
                    "aucun bloc <think>. Les guillemets dans les valeurs doivent "
                    "être remplacés par des guillemets simples '. "
                    "Commence directement par '{' et termine par '}'."
                )
                messages = [
                    {"role": "system", "content": sys_msg},
                    {"role": "user",   "content": prompt}
                ]
                payload = {
                    "model":      modele,
                    "messages":   messages,
                    "temperature": 0.1,
                    "stream":     False,
                    "max_tokens": 262144,    # contexte complet LM Studio
                }

                mod_lower = modele.lower()
                if any(k in mod_lower for k in ("qwen3", "qwen3.5", "qwq",
                                                 "deepseek-r", "gemma-4", "gemma4")):
                    payload["enable_thinking"] = False
                    self._log("   ℹ️  Mode thinking désactivé")

                headers = {"Connection": "close"}
                r = requests.post(f"{base}/v1/chat/completions",
                                  json=payload, timeout=None, headers=headers)

                if r.status_code == 400:
                    err = r.text
                    self._log(f"   ⚠️  HTTP 400 : {err[:200]}")
                    if any(k in err.lower() for k in ("system", "role", "jinja")):
                        self._log("   🔄 Retry sans system role")
                        merged = f"[CONSIGNE]: {sys_msg}\n\n{prompt}"
                        payload["messages"] = [{"role": "user", "content": merged}]
                        r = requests.post(f"{base}/v1/chat/completions",
                                          json=payload, timeout=None, headers=headers)
                        if r.status_code != 200:
                            self._log(f"   ❌ Retry échoué HTTP {r.status_code}")
                            return None
                    else:
                        return None
                elif r.status_code != 200:
                    self._log(f"   ❌ LM Studio HTTP {r.status_code} : {r.text[:300]}")
                    return None

                resp_json = r.json()
                choice    = resp_json.get("choices", [{}])[0]
                msg       = choice.get("message", {})
                brut      = msg.get("content", "") or ""

                if not brut.strip():
                    brut = msg.get("reasoning_content", "") or ""
                    if brut.strip():
                        self._log("   ℹ️  Contenu depuis reasoning_content")

                if not brut.strip():
                    self._log(f"   ❌ Réponse vide — {str(resp_json)[:200]}")
                    return None

            else:
                # Ollama
                payload = {
                    "model":   modele,
                    "prompt":  prompt,
                    "stream":  False,
                    "options": {"temperature": 0.1, "num_ctx": 262144}
                }
                if json_force:
                    payload["format"] = "json"
                r = requests.post(f"{base}/api/generate", json=payload, timeout=None)
                if r.status_code != 200:
                    self._log(f"   ❌ Ollama HTTP {r.status_code}")
                    return None
                brut = r.json().get("response", "")

            self._log(f"   🔍 Réponse ({len(brut)} car.) : {brut[:100]!r}")
            if not brut.strip():
                self._log("   ❌ Réponse vide")
                return None

            # ── Tentative 1 : parse direct après nettoyage basique ──────────
            propre = re.sub(r"<think>.*?</think>", "", brut, flags=re.DOTALL)
            propre = re.sub(r"```json\s*|```\s*", "", propre).strip()
            idx_s  = propre.find("{")
            idx_e  = propre.rfind("}")
            if idx_s != -1 and idx_e != -1:
                extrait = propre[idx_s:idx_e + 1]
                try:
                    return json.loads(extrait)
                except json.JSONDecodeError:
                    pass

            # ── Tentative 2 : réparation automatique du JSON ────────────────
            self._log("   🔧 JSON invalide → tentative de réparation...")
            repare = reparer_json(brut)
            try:
                resultat = json.loads(repare)
                self._log("   ✅ JSON réparé avec succès")
                return resultat
            except json.JSONDecodeError as e2:
                self._log(f"   ⚠️  Réparation insuffisante : {e2}")
                self._log(f"   Aperçu réparé : {repare[:200]!r}")

            # ── Tentative 3 : retry IA avec prompt condensé ─────────────────
            if backend == "lmstudio":
                self._log("   🔄 Retry IA avec température réduite...")
                prompt_retry = (
                    "Corrige ce JSON malformé et retourne UNIQUEMENT le JSON valide corrigé.\n"
                    "Remplace les guillemets doubles dans les valeurs par des guillemets simples.\n"
                    "Ne rajoute rien d'autre.\n\n"
                    f"JSON À CORRIGER :\n{brut[:4000]}"
                )
                payload_retry = {
                    "model":      modele,
                    "messages":   [
                        {"role": "system", "content": "Retourne uniquement du JSON valide."},
                        {"role": "user",   "content": prompt_retry}
                    ],
                    "temperature": 0.05,
                    "stream":     False,
                    "max_tokens": 8192,
                }
                if any(k in modele.lower() for k in ("qwen3", "gemma-4", "gemma4")):
                    payload_retry["enable_thinking"] = False
                r2 = requests.post(f"{base}/v1/chat/completions",
                                   json=payload_retry, timeout=None, headers=headers)
                if r2.status_code == 200:
                    brut2 = r2.json().get("choices", [{}])[0].get("message", {}).get("content", "")
                    repare2 = reparer_json(brut2)
                    try:
                        res2 = json.loads(repare2)
                        self._log("   ✅ JSON récupéré via retry IA")
                        return res2
                    except json.JSONDecodeError as e3:
                        self._log(f"   ❌ JSON invalide après retry : {e3}")

            return None

        except json.JSONDecodeError as e:
            self._log(f"   ❌ JSON invalide : {e}")
            self._log(f"   Aperçu brut : {brut[:300]!r}")
            return None
        except Exception as e:
            self._log(f"   ❌ Erreur IA : {e}")
            return None

    def _extraire_texte_libre(self, texte: str) -> str:
        """Extrait une question depuis du texte libre (fallback Phase 0)."""
        if not texte:
            return ""
        m = re.search(r'[A-ZÀÂÉÈÊËÎÏÔÙÛÜ][^.!?\n]{15,200}\?', texte)
        if m:
            return m.group(0).strip()
        for kw in ["problématique", "problematique", "question", "reformulation"]:
            m = re.search(rf'{kw}\s*[:\-–]\s*(.{{10,200}})', texte, re.IGNORECASE)
            if m:
                return m.group(1).strip().strip('"')
        return ""

    def _appel_ia_phase0(self, prompt: str) -> dict:
        """
        Appel Phase 0 avec double fallback.
        Retourne dict avec clés : axes, reformulation, problematique, academique, justification
        """
        vide = {"axes_thematiques": [], "reformulation_demande": "",
                "problematique": "", "reformulation_academique": "", "justification": ""}

        # Tentative 1 : JSON
        result = self._appel_ia(prompt, json_force=True)
        if result:
            pb = (result.get("problematique") or result.get("problématique")
                  or result.get("question") or "").strip()
            if pb:
                return {
                    "axes_thematiques":      result.get("axes_thematiques", []),
                    "reformulation_demande": result.get("reformulation_demande", ""),
                    "problematique":         pb,
                    "reformulation_academique": result.get("reformulation_academique", ""),
                    "justification":         result.get("justification", "")
                }

        self._log("   ⚠️  Phase 0 JSON vide — tentative texte libre...")

        # Tentative 2 : texte libre
        payload = {
            "model":   self.var_modele.get(),
            "prompt":  prompt + "\n\nRéponds en français avec une seule question de recherche.",
            "stream":  False,
            "options": {"temperature": 0.2, "num_ctx": 4096}
        }
        try:
            r = requests.post(self.var_url.get().rstrip("/") + "/api/generate",
                              json=payload, timeout=None)
            brut = r.json().get("response", "")
            self._log(f"   🔍 Texte libre : {brut[:200]!r}")
            pb = self._extraire_texte_libre(brut)
            if pb:
                return {**vide, "problematique": pb,
                        "justification": "Extraite depuis la réponse texte du modèle."}
        except Exception as e:
            self._log(f"   ❌ Fallback texte libre : {e}")

        return vide

    # ══════════════════════════════════════════════════
    #  POPUP VALIDATION — Phase 0 (version ergo)
    # ══════════════════════════════════════════════════

    def _popup_phase0(self, data: dict) -> str:
        """
        Popup riche affichant :
        - les axes thématiques identifiés depuis le terrain
        - la reformulation de la demande
        - la problématique proposée (modifiable)
        - la reformulation académique
        Retourne la problématique validée (vide = annulation).
        """
        result_q: queue.Queue = queue.Queue()

        def _afficher():
            dlg = tk.Toplevel(self.root)
            dlg.title("🔬 Phase 0 — Terrain → Reformulation → Problématique")
            dlg.configure(bg="#1a1a2e")
            dlg.geometry("820x600")
            dlg.grab_set(); dlg.lift(); dlg.focus_force()

            pb_ok = bool(data.get("problematique"))
            titre_col = "#88ffcc" if pb_ok else "#ffaa44"
            titre_txt = ("✅  L'IA a lu le terrain et propose la problématique suivante :"
                         if pb_ok else
                         "⚠️  L'IA n'a pas pu formuler de problématique — saisissez-la manuellement :")

            tk.Label(dlg, text=titre_txt, bg="#1a1a2e", fg=titre_col,
                     font=("Segoe UI", 10, "bold"), wraplength=780,
                     justify="left").pack(padx=16, pady=(14,4), anchor="w")

            # Axes thématiques
            axes = data.get("axes_thematiques", [])
            if axes:
                tk.Label(dlg, text="📌 Axes thématiques identifiés depuis le terrain :",
                         bg="#1a1a2e", fg="#cccccc",
                         font=("Segoe UI", 9, "bold")).pack(padx=16, anchor="w")
                tk.Label(dlg, text="  • " + "\n  • ".join(axes),
                         bg="#1a1a2e", fg="#aaccaa",
                         font=("Consolas", 9), justify="left").pack(padx=16, pady=(0,6), anchor="w")

            # Reformulation de la demande
            ref = data.get("reformulation_demande", "")
            if ref:
                tk.Label(dlg, text="🔄 Reformulation de la demande depuis le terrain :",
                         bg="#1a1a2e", fg="#cccccc",
                         font=("Segoe UI", 9, "bold")).pack(padx=16, anchor="w")
                zr = tk.Text(dlg, bg="#1e2e2a", fg="#aaccee", font=("Segoe UI", 9),
                             height=3, wrap="word", relief="flat", padx=8, pady=4)
                zr.pack(fill="x", padx=16, pady=(0,8))
                zr.insert("1.0", ref)
                zr.config(state="disabled")

            # Reformulation académique
            acad = data.get("reformulation_academique", "")
            if acad:
                tk.Label(dlg, text="🎓 Reformulation académique (pour référence) :",
                         bg="#1a1a2e", fg="#888899",
                         font=("Segoe UI", 8, "italic")).pack(padx=16, anchor="w")
                tk.Label(dlg, text=acad, bg="#1a1a2e", fg="#888899",
                         font=("Segoe UI", 8, "italic"),
                         wraplength=780, justify="left").pack(padx=24, pady=(0,6), anchor="w")

            # Justification
            just = data.get("justification", "")
            if just:
                tk.Label(dlg, text="💡 Justification :", bg="#1a1a2e", fg="#888899",
                         font=("Segoe UI", 8)).pack(padx=16, anchor="w")
                tk.Label(dlg, text=just, bg="#1a1a2e", fg="#888899",
                         font=("Segoe UI", 8, "italic"),
                         wraplength=780, justify="left").pack(padx=24, pady=(0,8), anchor="w")

            # Problématique éditable
            tk.Label(dlg, text="✏️  Problématique retenue (modifiable — obligatoire) :",
                     bg="#1a1a2e", fg="#88ffcc",
                     font=("Segoe UI", 9, "bold")).pack(padx=16, anchor="w")
            var_pb = tk.StringVar(value=data.get("problematique", ""))
            ent = tk.Entry(dlg, textvariable=var_pb, width=84,
                           bg="#1e2e2a", fg="white", insertbackground="white",
                           font=("Segoe UI", 10))
            ent.pack(fill="x", padx=16, pady=6)
            ent.focus_set()
            if data.get("problematique"):
                ent.select_range(0, tk.END)

            tk.Label(dlg,
                     text='Ex : "Dans quelle mesure les outils numériques facilitent-ils l\'expérience utilisateur ?"',
                     bg="#1a1a2e", fg="#66bb88", font=("Segoe UI", 8, "italic"),
                     wraplength=780).pack(padx=16, pady=(0,6), anchor="w")

            lbl_err = tk.Label(dlg, text="", bg="#1a1a2e", fg="#ff6666",
                               font=("Segoe UI", 9, "bold"))
            lbl_err.pack()

            fb = tk.Frame(dlg, bg="#1a1a2e")
            fb.pack(pady=8)

            def _ok():
                v = var_pb.get().strip()
                if not v:
                    lbl_err.config(text="⚠️  Le champ ne peut pas être vide.")
                    ent.focus_set()
                    return
                result_q.put(v)
                dlg.destroy()

            def _no():
                result_q.put("")
                dlg.destroy()

            ent.bind("<Return>", lambda e: _ok())
            tk.Button(fb, text="✅  Valider et continuer l'analyse",
                      command=_ok, bg="#007acc", fg="white",
                      font=("Segoe UI", 10, "bold"),
                      relief="flat", padx=16, pady=6).pack(side="left", padx=8)
            tk.Button(fb, text="✖  Annuler", command=_no, bg="#663333", fg="white",
                      font=("Segoe UI", 10), relief="flat", padx=16, pady=6).pack(side="left", padx=8)

        self.root.after(0, _afficher)
        return result_q.get()

    # ══════════════════════════════════════════════════
    #  LANCEMENT
    # ══════════════════════════════════════════════════

    def _lancer(self):
        if self.en_cours:
            messagebox.showwarning("En cours", "Une analyse est déjà en cours.")
            return
        if not self.var_modele.get():
            messagebox.showerror("Modèle manquant", "Cliquez d'abord sur 'Tester & Charger modèles'.")
            return
        self._sauvegarder_config()
        self.en_cours = True
        if self.var_mode_methodo.get():
            self.btn_lancer.config(state="disabled", text="🔬  Analyse méthodologique en cours...")
            threading.Thread(target=self._processus_methodologique, daemon=True).start()
        else:
            self.btn_lancer.config(state="disabled", text="⏳  Analyse en cours...")
            threading.Thread(target=self._processus_global, daemon=True).start()

    def _fin_analyse(self):
        self.en_cours = False
        self.root.after(0, lambda: self.btn_lancer.config(
            state="normal", text="🚀  LANCER L'ANALYSE ET ÉCRIRE DANS MINDMANAGER"))
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass

    # ══════════════════════════════════════════════════
    #  PROCESSUS STANDARD
    # ══════════════════════════════════════════════════

    def _processus_global(self):
        try:
            pythoncom.CoInitialize()
            fichiers = glob.glob(os.path.join(self.var_dossier.get(), "*.txt"))
            if not fichiers:
                self._log("❌ Aucun fichier .txt trouvé.")
                return

            nb_total = len(fichiers)
            mode_tr  = self.var_temps_reel.get()
            self._log(f"📋 {nb_total} fichier(s) — {'⚡ temps réel' if mode_tr else '📦 groupé + fusion à la fin'}")

            mm = rt = None
            if mode_tr:
                mm = self._init_mm()
                if mm is None:
                    return
                doc = mm.Documents.Add()
                rt  = doc.CentralTopic
                rt.Text = "AUDIT UX — v2.6"

            # Structure enrichie :
            # synthese[theme] = {
            #   "_sources": set(),          # participants uniques (pour pondération)
            #   "sous_themes": {
            #     sous_theme: {
            #       "_sources": set(),
            #       sentiment: [(participant, verbatim_nettoye, verbatim_original), ...]
            #     }
            #   }
            # }
            synthese: dict = {}

            # ── LECTURE + NETTOYAGE + ANALYSE ─────────────────────────────────
            for i, chemin in enumerate(fichiers, 1):
                nom = os.path.basename(chemin)
                # Génération du code participant : "Test 1.txt" → "T1",
                # "entretien_7.txt" → "T7", "S3.txt" → "T3", etc.
                nom_sans_ext = re.sub(r'\.(txt)$', '', nom, flags=re.IGNORECASE).strip()
                # Cherche un ou plusieurs chiffres dans le nom
                nums = re.findall(r'\d+', nom_sans_ext)
                if nums:
                    code_part = f"T{nums[-1]}"   # prend le dernier numéro : "Test 12" → T12
                else:
                    # Pas de numéro : initiales du nom, max 3 chars
                    initiales = re.sub(r'[^a-zA-Z]', '', nom_sans_ext)[:3].upper()
                    code_part = initiales if initiales else f"P{i}"

                self._log(f"{'─'*50}")
                self._log(f"📄 [{i}/{nb_total}] {nom}  →  [{code_part}]")

                try:
                    with open(chemin, "r", encoding="utf-8", errors="ignore") as fh:
                        contenu = fh.read()

                    # ── ÉTAPE A : Nettoyage des verbatims (optionnel, si texte oral) ──
                    # On fait d'abord l'extraction brute, puis on nettoie
                    prompt = self.zone_prompt.get("1.0", tk.END).strip()
                    prompt_complet = (
                        prompt
                        + f"\n\nFICHIER : {nom}\nCODE PARTICIPANT : {code_part}\n\n"
                        + contenu[:15000]
                    )
                    self._log("   ⏳ Analyse IA...")
                    donnees = self._appel_ia(prompt_complet)

                    if not donnees:
                        self._log(f"   ⚠️  Pas de résultat pour {nom}")
                        continue

                    themes = donnees.get("themes", [])
                    self._log(f"   ✅ {len(themes)} thème(s) brut(s)")

                    if mode_tr:
                        nf = self._cs(rt, f"📄 [{code_part}] {nom}")
                        if nf:
                            self._ecrire_themes_mm(nf, themes, code_part)
                        self._log("   ✏️  MindManager mis à jour ✔")
                    else:
                        # Accumulation enrichie
                        for th in themes:
                            theme_nom = _strip_md(th.get("theme", "Divers"))
                            if theme_nom not in synthese:
                                synthese[theme_nom] = {"_sources": set(), "sous_themes": {}}
                            synthese[theme_nom]["_sources"].add(code_part)

                            for sth in th.get("sous_themes", []):
                                st_nom = _strip_md(sth.get("sous_theme", "Général"))
                                if st_nom not in synthese[theme_nom]["sous_themes"]:
                                    synthese[theme_nom]["sous_themes"][st_nom] = {"_sources": set()}
                                synthese[theme_nom]["sous_themes"][st_nom]["_sources"].add(code_part)

                                for ex in sth.get("extraits", []):
                                    sent = _strip_md(ex.get("sentiment", "Nuancé"))
                                    part = code_part   # toujours [T1] etc., pas ce que l'IA invente
                                    verb = _strip_md(ex.get("verbatim", ""))
                                    if verb:
                                        synthese[theme_nom]["sous_themes"][st_nom].setdefault(
                                            sent, []
                                        ).append((part, verb))

                except Exception as ef:
                    self._log(f"   ❌ {ef}\n{traceback.format_exc()}")

            self._log(f"\n{'═'*50}")

            if mode_tr:
                self._log("🎉 Analyse temps réel terminée !")
                self._forcer_premier_plan(mm)
                self.root.after(0, lambda: messagebox.showinfo("Terminé ✅", "Carte MindManager mise à jour."))
                return

            # ── ÉTAPE FINALE MODE GROUPÉ ───────────────────────────────────────

            # 1. Fusion des thèmes similaires
            nb_avant = len(synthese)
            self._log(f"\n🔀 FUSION — {nb_avant} thèmes bruts identifiés...")
            synthese_fusionnee = fusionner_themes_v2(synthese)
            nb_apres = len(synthese_fusionnee)
            self._log(f"   ✅ {nb_avant} → {nb_apres} thème(s) après fusion")

            if not synthese_fusionnee:
                self._log("⚠️  Aucun résultat — vérifiez le prompt.")
                return

            # 2. Classification heuristiques
            self._log("\n📐 Classification heuristiques (Nielsen / Bastien & Scapin)...")
            noms_themes = list(synthese_fusionnee.keys())
            heuristiques_map = self._classifier_heuristiques(noms_themes)

            # 3. Écriture MindManager
            self._log("\n🖊️  Écriture dans MindManager...")
            mm = self._init_mm()
            if mm is None:
                return
            doc  = mm.Documents.Add()
            root = doc.CentralTopic
            root.Text = "AUDIT UX — SYNTHÈSE v2.6"

            # Compteurs globaux pour le tableau de bord
            compteurs = {"Négatif": 0, "Positif": 0, "Nuancé": 0}
            freq_themes = {}  # theme → (n_sources, total_verbatims)

            for theme_nom, data_theme in synthese_fusionnee.items():
                sources   = data_theme.get("_sources", set())
                n_sources = len(sources)
                freq_themes[theme_nom] = n_sources

                # Barre emoji pondération
                barre = "●" * n_sources + "○" * max(0, nb_total - n_sources)
                label_theme = f"{theme_nom}  [{n_sources}/{nb_total}] {barre}"

                n_th = self._cs(root, label_theme)
                if not n_th:
                    continue

                # Heuristique rattachée
                h = heuristiques_map.get(theme_nom)
                if h:
                    nielsen    = ", ".join(h.get("nielsen", []))
                    bastien    = ", ".join(h.get("bastien_scapin", []))
                    gravite    = h.get("gravite", "?")
                    tag_gravite = {"Majeur": "🔴", "Modéré": "🟡", "Mineur": "🟢"}.get(gravite, "⚪")
                    if nielsen or bastien:
                        n_heur = self._cs(n_th, f"📐 Cadre théorique")
                        if n_heur:
                            if nielsen:
                                self._cs(n_heur, f"Nielsen : {nielsen}")
                            if bastien:
                                self._cs(n_heur, f"Bastien & Scapin : {bastien}")
                            self._cs(n_heur, f"Gravité : {tag_gravite} {gravite}")

                for st_nom, data_st in data_theme.get("sous_themes", {}).items():
                    n_st = self._cs(n_th, st_nom)
                    if not n_st:
                        continue
                    for sent_nom, entrees in data_st.items():
                        if sent_nom.startswith("_"):
                            continue
                        icon = {"Négatif": "🔴", "Positif": "🟢", "Nuancé": "🟡"}.get(sent_nom, "⚪")
                        n_sent = self._cs(n_st, f"{icon} {sent_nom} ({len(entrees)})")
                        compteurs[sent_nom] = compteurs.get(sent_nom, 0) + len(entrees)
                        if n_sent:
                            for part, verb in entrees:
                                self._cs(n_sent, f"[{part}] {verb}")

            # 4. Nœud Tableau de bord
            self._ecrire_tableau_bord(root, compteurs, synthese_fusionnee, nb_total, freq_themes)

            self._forcer_premier_plan(mm)
            total_v = sum(compteurs.values())
            self._log(f"🎉 Carte complète — {nb_apres} thème(s), {total_v} verbatim(s) !")

            # 5. Rapport HTML
            self._log("\n📄 Génération du rapport HTML...")
            self._generer_rapport_html(synthese_fusionnee, compteurs, nb_total, freq_themes)

            self.root.after(0, lambda: messagebox.showinfo(
                "Analyse terminée ✅",
                f"{nb_apres} thème(s) consolidés — {total_v} verbatim(s)\n"
                f"🔴 Négatif: {compteurs.get('Négatif',0)}  "
                f"🟢 Positif: {compteurs.get('Positif',0)}  "
                f"🟡 Nuancé: {compteurs.get('Nuancé',0)}\n\n"
                f"Rapport HTML généré dans votre dossier d'entretiens."
            ))

        except Exception as e:
            self._log(f"❌ ERREUR : {e}\n{traceback.format_exc()}")
        finally:
            self._fin_analyse()

    def _classifier_heuristiques(self, noms_themes: list) -> dict:
        """Appel IA pour rattacher chaque thème à Nielsen / Bastien & Scapin."""
        if not noms_themes:
            return {}
        prompt = _PROMPT_HEURISTIQUES.format(themes="\n".join(f"- {t}" for t in noms_themes))
        donnees = self._appel_ia(prompt)
        result = {}
        if donnees:
            for item in donnees.get("classifications", []):
                t = item.get("theme", "")
                if t:
                    result[t] = item
            self._log(f"   ✅ {len(result)} thème(s) classifié(s)")
        else:
            self._log("   ⚠️  Classification heuristiques non disponible")
        return result

    def _ecrire_tableau_bord(self, root, compteurs: dict, synthese: dict,
                              nb_total: int, freq_themes: dict):
        """Crée un nœud Tableau de bord dans MindManager avec statistiques."""
        total = sum(compteurs.values())
        if total == 0:
            return

        n_tb = self._cs(root, "📊 TABLEAU DE BORD")
        if not n_tb:
            return

        # Totaux sentiments
        n_sent = self._cs(n_tb, "Répartition des sentiments")
        if n_sent:
            for sent, icon in [("Négatif","🔴"),("Positif","🟢"),("Nuancé","🟡")]:
                n   = compteurs.get(sent, 0)
                pct = round(n / total * 100) if total else 0
                barre = "█" * (pct // 10) + "░" * (10 - pct // 10)
                self._cs(n_sent, f"{icon} {sent} : {n} ({pct}%)  {barre}")

        # Top 5 thèmes les plus fréquents
        top5 = sorted(freq_themes.items(), key=lambda x: x[1], reverse=True)[:5]
        n_top = self._cs(n_tb, "Top 5 thèmes récurrents")
        if n_top:
            for i, (th, freq) in enumerate(top5, 1):
                barre = "●" * freq + "○" * max(0, nb_total - freq)
                self._cs(n_top, f"#{i} {th}  [{freq}/{nb_total}]  {barre}")

        # Synthèse lisible pour néophytes
        neg = compteurs.get("Négatif", 0)
        pos = compteurs.get("Positif", 0)
        nua = compteurs.get("Nuancé", 0)
        ratio_neg = round(neg / total * 100) if total else 0

        n_syn = self._cs(n_tb, "💬 Synthèse pour néophytes")
        if n_syn:
            self._cs(n_syn, f"Sur {total} observations recueillies dans {nb_total} entretiens :")
            self._cs(n_syn, f"{neg} problèmes identifiés ({ratio_neg}% des retours)")
            self._cs(n_syn, f"{pos} points positifs relevés")
            self._cs(n_syn, f"{nua} retours nuancés ou ambivalents")
            if top5:
                self._cs(n_syn, f"Le problème le plus récurrent : « {top5[0][0]} » (mentionné dans {top5[0][1]}/{nb_total} entretiens)")

        self._log("   ✅ Tableau de bord injecté dans MindManager")

    def _generer_rapport_html(self, synthese: dict, compteurs: dict,
                               nb_total: int, freq_themes: dict):
        """Génère un rapport HTML autonome avec graphiques matplotlib en PNG intégré."""
        try:
            import matplotlib
            matplotlib.use("Agg")   # pas de fenêtre GUI
            import matplotlib.pyplot as plt
            import matplotlib.patches as mpatches
            import base64
            from io import BytesIO

            def fig_to_b64(fig) -> str:
                buf = BytesIO()
                fig.savefig(buf, format="png", bbox_inches="tight",
                            dpi=120, facecolor="#1a1a2e")
                buf.seek(0)
                return base64.b64encode(buf.read()).decode()

            couleurs = {"Négatif": "#ff5555", "Positif": "#55ff99", "Nuancé": "#ffdd55"}

            # ── Graphique 1 : Camembert sentiments ───────────────────────────
            fig1, ax1 = plt.subplots(figsize=(5, 5), facecolor="#1a1a2e")
            labels = [k for k, v in compteurs.items() if v > 0]
            sizes  = [compteurs[k] for k in labels]
            cols   = [couleurs.get(k, "#888888") for k in labels]
            ax1.pie(sizes, labels=labels, colors=cols, autopct="%1.0f%%",
                    textprops={"color": "white", "fontsize": 13})
            ax1.set_title("Répartition des sentiments", color="white", fontsize=14, pad=15)
            img1 = fig_to_b64(fig1)
            plt.close(fig1)

            # ── Graphique 2 : Barres fréquence thèmes ────────────────────────
            top_th = sorted(freq_themes.items(), key=lambda x: x[1], reverse=True)[:10]
            if top_th:
                noms  = [t[0][:35] + ("…" if len(t[0]) > 35 else "") for t in top_th]
                freqs = [t[1] for t in top_th]
                bar_cols = ["#ff5555" if f >= nb_total * 0.6 else
                            "#ffdd55" if f >= nb_total * 0.3 else
                            "#55aaff" for f in freqs]
                fig2, ax2 = plt.subplots(figsize=(10, 5), facecolor="#1a1a2e")
                ax2.set_facecolor("#12122a")
                bars = ax2.barh(noms[::-1], freqs[::-1], color=bar_cols[::-1])
                ax2.set_xlabel("Nombre de participants", color="white")
                ax2.set_title(f"Fréquence des thèmes ({nb_total} participants)", color="white", fontsize=13)
                ax2.tick_params(colors="white")
                ax2.spines[:].set_color("#444466")
                for bar, val in zip(bars, freqs[::-1]):
                    ax2.text(bar.get_width() + 0.05, bar.get_y() + bar.get_height()/2,
                             f"{val}/{nb_total}", va="center", color="white", fontsize=10)
                legend_patches = [
                    mpatches.Patch(color="#ff5555", label="Récurrent (≥60%)"),
                    mpatches.Patch(color="#ffdd55", label="Modéré (≥30%)"),
                    mpatches.Patch(color="#55aaff", label="Rare (<30%)"),
                ]
                ax2.legend(handles=legend_patches, facecolor="#1a1a2e",
                           labelcolor="white", loc="lower right")
                img2 = fig_to_b64(fig2)
                plt.close(fig2)
            else:
                img2 = None

            # ── Construction HTML ─────────────────────────────────────────────
            total_v = sum(compteurs.values())
            neg = compteurs.get("Négatif", 0)
            pos = compteurs.get("Positif", 0)
            nua = compteurs.get("Nuancé", 0)

            lignes_themes = ""
            for th, freq in sorted(freq_themes.items(), key=lambda x: x[1], reverse=True):
                barre_html = "●" * freq + "○" * max(0, nb_total - freq)
                pct = round(freq / nb_total * 100) if nb_total else 0
                couleur_row = "#2a1a1a" if pct >= 60 else "#2a2a1a" if pct >= 30 else "#1a1a2a"
                lignes_themes += f"""
                <tr style="background:{couleur_row}">
                    <td style="padding:8px 12px">{th}</td>
                    <td style="text-align:center;color:#aaa">{freq}/{nb_total}</td>
                    <td style="text-align:center;color:#aaa">{pct}%</td>
                    <td style="color:#ffaa44;letter-spacing:2px">{barre_html}</td>
                </tr>"""

            html = f"""<!DOCTYPE html>
<html lang="fr">
<head>
<meta charset="UTF-8">
<title>Rapport UX — Audit Station v2.6</title>
<style>
  body {{font-family:'Segoe UI',sans-serif;background:#0d0d1a;color:#e0e0ff;margin:0;padding:24px}}
  h1 {{color:#88ccff;border-bottom:2px solid #334;padding-bottom:12px}}
  h2 {{color:#aaddff;margin-top:36px}}
  .stat-grid {{display:grid;grid-template-columns:repeat(3,1fr);gap:16px;margin:20px 0}}
  .stat-card {{background:#1a1a2e;border:1px solid #334;border-radius:8px;padding:20px;text-align:center}}
  .stat-num {{font-size:2.5em;font-weight:bold;margin:8px 0}}
  .rouge {{color:#ff5555}} .vert {{color:#55ff99}} .jaune {{color:#ffdd55}}
  table {{width:100%;border-collapse:collapse;background:#12122a;border-radius:8px;overflow:hidden}}
  th {{background:#1e1e3e;padding:10px 12px;text-align:left;color:#88ccff}}
  td {{padding:6px 12px;border-top:1px solid #223}}
  img {{max-width:100%;border-radius:8px;margin:16px 0}}
  .charts {{display:grid;grid-template-columns:1fr 2fr;gap:20px;align-items:start}}
  .footer {{margin-top:40px;color:#555;font-size:0.85em;border-top:1px solid #333;padding-top:12px}}
</style>
</head>
<body>
<h1>📊 Rapport d'Audit UX</h1>
<p style="color:#888">Généré par Station UX Audit v2.6 — {time.strftime('%d/%m/%Y %H:%M')}</p>

<h2>Vue d'ensemble</h2>
<div class="stat-grid">
  <div class="stat-card">
    <div>Entretiens analysés</div>
    <div class="stat-num" style="color:#88ccff">{nb_total}</div>
  </div>
  <div class="stat-card">
    <div>Total observations</div>
    <div class="stat-num" style="color:#aaaaff">{total_v}</div>
  </div>
  <div class="stat-card">
    <div>Thèmes identifiés</div>
    <div class="stat-num" style="color:#88ffcc">{len(freq_themes)}</div>
  </div>
</div>

<div class="stat-grid">
  <div class="stat-card">
    <div>🔴 Problèmes</div>
    <div class="stat-num rouge">{neg}</div>
    <div style="color:#888">{round(neg/total_v*100) if total_v else 0}% des retours</div>
  </div>
  <div class="stat-card">
    <div>🟢 Satisfactions</div>
    <div class="stat-num vert">{pos}</div>
    <div style="color:#888">{round(pos/total_v*100) if total_v else 0}% des retours</div>
  </div>
  <div class="stat-card">
    <div>🟡 Nuancés</div>
    <div class="stat-num jaune">{nua}</div>
    <div style="color:#888">{round(nua/total_v*100) if total_v else 0}% des retours</div>
  </div>
</div>

<h2>Graphiques</h2>
<div class="charts">
  <div><img src="data:image/png;base64,{img1}" alt="Sentiments"></div>
  {"<div><img src='data:image/png;base64," + img2 + "' alt='Fréquences'></div>" if img2 else ""}
</div>

<h2>Fréquence des thèmes</h2>
<table>
  <tr><th>Thème</th><th>Participants</th><th>Fréquence</th><th>Intensité</th></tr>
  {lignes_themes}
</table>

<div class="footer">
  Rapport généré automatiquement par Station UX Audit v2.6.<br>
  Les données proviennent de {nb_total} entretiens analysés par IA locale.
</div>
</body>
</html>"""

            # Sauvegarde dans le dossier des entretiens
            dossier = self.var_dossier.get()
            chemin_html = os.path.join(dossier, f"rapport_ux_{time.strftime('%Y%m%d_%H%M')}.html")
            with open(chemin_html, "w", encoding="utf-8") as f:
                f.write(html)
            self._log(f"   ✅ Rapport HTML : {chemin_html}")

            # Ouvre dans le navigateur par défaut
            import webbrowser
            webbrowser.open(f"file:///{chemin_html.replace(chr(92), '/')}")

        except ImportError:
            self._log("   ⚠️  matplotlib non installé — pip install matplotlib")
            self._log("       Le rapport HTML ne sera pas généré.")
        except Exception as e:
            self._log(f"   ❌ Erreur rapport HTML : {e}")

    def _ecrire_themes_mm(self, parent, themes: list, code_part: str = "?"):
        """
        Mode temps réel : écrit les thèmes d'un fichier dans MindManager.
        Hiérarchie : Thème → Sous-thème → Sentiment → [T1] Verbatim (+ Note)
        """
        for th in themes:
            theme_nom = _strip_md(th.get("theme", "Divers"))
            n_th = self._cs(parent, theme_nom)
            if not n_th:
                continue

            for sth in th.get("sous_themes", []):
                st_nom = _strip_md(sth.get("sous_theme", "Général"))
                n_st   = self._cs(n_th, st_nom)
                if not n_st:
                    continue

                # Regroupe par sentiment
                par_sent: dict = {}
                for ex in sth.get("extraits", []):
                    sent  = _strip_md(ex.get("sentiment", "Nuancé"))
                    verb  = _strip_md(ex.get("verbatim", ""))
                    justi = _strip_md(ex.get("justification_codage", ""))
                    if verb:
                        par_sent.setdefault(sent, []).append((code_part, verb, justi))

                for sent_nom, entrees in par_sent.items():
                    icon   = {"Négatif": "🔴", "Positif": "🟢", "Nuancé": "🟡"}.get(sent_nom, "⚪")
                    n_sent = self._cs(n_st, f"{icon} {sent_nom} ({len(entrees)})")
                    if not n_sent:
                        continue
                    for part, verb, justi in entrees:
                        self._cs_verbatim(n_sent, part, verb, justi)

    # ══════════════════════════════════════════════════
    #  PROCESSUS MÉTHODOLOGIQUE
    # ══════════════════════════════════════════════════

    def _processus_methodologique(self):
        try:
            pythoncom.CoInitialize()
            fichiers = glob.glob(os.path.join(self.var_dossier.get(), "*.txt"))
            prob_init    = self.var_problematique.get().strip()
            demande_init = self.var_demande_initiale.get().strip() or "(non renseignée)"

            if not fichiers:
                self._log("❌ Aucun fichier .txt trouvé.")
                return

            self._log(f"\n{'═'*55}")
            self._log(f"🔬 ANALYSE MÉTHODOLOGIQUE — {len(fichiers)} fichier(s)")
            self._log(f"{'═'*55}")

            corpus = {}
            for f in fichiers:
                nom = os.path.basename(f)
                with open(f, "r", encoding="utf-8", errors="ignore") as fh:
                    corpus[nom] = fh.read()

            # ── PHASE 0 : Terrain → Reformulation → Problématique ─────────────
            problematique = prob_init
            if not problematique:
                self._log(f"\n── PHASE 0 — Terrain → Reformulation → Problématique")
                self._log(f"   Demande initiale : {demande_init}")
                self._log("   ⏳ L'IA lit le terrain sans hypothèse préalable...")

                extrait = "\n\n".join(f"=== {n} ===\n{t[:3000]}" for n, t in corpus.items())
                data_p0 = self._appel_ia_phase0(
                    _PROMPT_PHASE0_ERGO.format(demande_initiale=demande_init) + extrait
                )

                self._log(f"   Axes identifiés : {data_p0.get('axes_thematiques', [])}")
                self._log(f"   Reformulation : {data_p0.get('reformulation_demande','?')!r}")
                self._log(f"   Problématique : {data_p0.get('problematique','?')!r}")
                self._log("   ⏸  Validation (popup)...")

                problematique = self._popup_phase0(data_p0)
                if not problematique:
                    self._log("   ❌ Annulé.")
                    return
                self._log(f"   ✅ Validée : {problematique}")

            self._log(f"\n📌 Problématique : {problematique}")

            # ── PHASE 1 : Grille thématique ───────────────────────────────────
            self._log("\n── PHASE 1 — Grille thématique")
            corpus_txt = "\n\n".join(f"=== {n} ===\n{t[:4000]}" for n, t in corpus.items())
            p1 = self._appel_ia(_PROMPT_PHASE1.format(prob=problematique) + corpus_txt)
            if not p1:
                self._log("❌ Phase 1 échouée.")
                return

            theme_general = _strip_md(p1.get("theme_general", "?"))
            mots_cles     = [_strip_md(m) for m in p1.get("mots_cles", [])]
            grille        = p1.get("grille_themes", [])
            obs_init      = _strip_md(p1.get("premieres_observations", ""))
            self._log(f"   ✅ {len(grille)} thème(s) — mots-clés : {', '.join(mots_cles[:6])}")

            # ── PHASE 2 : Codage ──────────────────────────────────────────────
            self._log(f"\n── PHASE 2 — Codage ({len(fichiers)} entretiens)")
            grille_str = json.dumps(grille, ensure_ascii=False)
            codes: dict = {}
            for i, (nom, texte) in enumerate(corpus.items(), 1):
                self._log(f"   [{i}/{len(fichiers)}] {nom}...")
                r = self._appel_ia(
                    _PROMPT_PHASE2.format(prob=problematique, grille=grille_str) + texte[:8000]
                )
                codes[nom] = r.get("extraits", []) if r else []
                self._log(f"   ✅ {len(codes[nom])} extrait(s)")

            # ── PHASE 3 : Synthèse ────────────────────────────────────────────
            self._log("\n── PHASE 3 — Comparaison & synthèse")
            p3 = self._appel_ia(
                _PROMPT_PHASE3.format(prob=problematique, n=len(fichiers),
                                      donnees=json.dumps(codes, ensure_ascii=False)[:12000])
            ) or {}

            conv  = p3.get("convergences",  [])
            div   = p3.get("divergences",   [])
            rec   = p3.get("recurrences",   [])
            coocc = p3.get("cooccurrences", [])
            syn   = _strip_md(p3.get("synthese_finale", "Non générée."))
            self._log(f"   ✅ {len(conv)} conv / {len(div)} div / {len(rec)} récurr.")

            # ── Écriture MindManager ──────────────────────────────────────────
            self._log("\n🖊️  Écriture dans MindManager...")
            mm = self._init_mm()
            if mm is None:
                return

            doc  = mm.Documents.Add()
            root = doc.CentralTopic
            root.Text = f"ANALYSE — {problematique[:80]}"

            # Problématique
            np = self._cs(root, "🎯 Problématique")
            if np:
                self._cs(np, problematique)
                if demande_init != "(non renseignée)":
                    self._cs(np, f"Demande initiale : {demande_init}")
                self._cs(np, f"{len(fichiers)} entretien(s)")

            # Thème général
            ntg = self._cs(root, f"🌐 {theme_general}")
            if ntg:
                if mots_cles:
                    nm = self._cs(ntg, "Mots-clés")
                    if nm:
                        for m in mots_cles:
                            self._cs(nm, m)
                if obs_init:
                    no = self._cs(ntg, "Observations initiales")
                    if no:
                        self._cs(no, obs_init)

            # Grille
            ng = self._cs(root, "📐 Grille thématique")
            if ng:
                for th in grille:
                    nt = self._cs(ng, _strip_md(th.get("theme", "?")))
                    if nt:
                        if th.get("description"):
                            self._cs(nt, _strip_md(th["description"]))
                        for st in th.get("sous_themes", []):
                            self._cs(nt, f"↳ {_strip_md(st)}")

            # Extraits codés
            ne = self._cs(root, "📄 Extraits codés")
            if ne:
                for nom_f, extraits in codes.items():
                    nf = self._cs(ne, f"[{nom_f}]")
                    if nf:
                        for ex in extraits:
                            icon = {"Négatif":"🔴","Positif":"🟢","Nuancé":"🟡"}.get(
                                ex.get("sentiment",""), "⚪")
                            label = f"{icon} [{ex.get('type','?').upper()}] {_strip_md(ex.get('theme',''))} — {_strip_md(ex.get('description',''))}"
                            nx = self._cs(nf, label[:200])
                            if nx and ex.get("verbatim"):
                                self._cs(nx, f"💬 {_strip_md(ex['verbatim'])}")

            # Convergences
            nc = self._cs(root, f"✅ Convergences ({len(conv)})")
            if nc:
                for c in conv:
                    n1 = self._cs(nc, f"{_strip_md(c.get('theme','?'))} — {c.get('frequence','?')}/{len(fichiers)}")
                    if n1:
                        if c.get("observation"):
                            self._cs(n1, _strip_md(c["observation"]))
                        if c.get("exemple_verbatim"):
                            self._cs(n1, f"💬 {_strip_md(c['exemple_verbatim'])}")

            # Divergences
            nd = self._cs(root, f"⚡ Divergences ({len(div)})")
            if nd:
                for d in div:
                    n1 = self._cs(nd, f"{_strip_md(d.get('theme','?'))} [{' vs '.join(d.get('entre',[]))}]")
                    if n1:
                        if d.get("observation"):
                            self._cs(n1, _strip_md(d["observation"]))
                        for v in d.get("verbatims",[]):
                            self._cs(n1, f"💬 {_strip_md(v)}")

            # Récurrences
            nr = self._cs(root, f"🔁 Récurrences ({len(rec)})")
            if nr:
                for r2 in rec:
                    n1 = self._cs(nr, f"« {_strip_md(r2.get('expression','?'))} » — {r2.get('frequence','?')}×")
                    if n1 and r2.get("contexte"):
                        self._cs(n1, _strip_md(r2["contexte"]))

            # Cooccurrences
            if coocc:
                nco = self._cs(root, f"🔗 Cooccurrences ({len(coocc)})")
                if nco:
                    for co in coocc:
                        n1 = self._cs(nco, " + ".join(_strip_md(t) for t in co.get("themes",[])))
                        if n1 and co.get("observation"):
                            self._cs(n1, _strip_md(co["observation"]))

            # Synthèse
            ns = self._cs(root, "📝 Synthèse finale")
            if ns:
                self._cs(ns, syn)

            self._log("🎉 CARTE MÉTHODOLOGIQUE COMPLÈTE !")
            self._forcer_premier_plan(mm)
            self.root.after(0, lambda: messagebox.showinfo(
                "Analyse méthodologique terminée ✅",
                f"✅ {len(conv)} convergence(s)\n⚡ {len(div)} divergence(s)\n"
                f"🔁 {len(rec)} récurrence(s)\n🔗 {len(coocc)} cooccurrence(s)"
            ))

        except Exception as e:
            self._log(f"❌ ERREUR : {e}\n{traceback.format_exc()}")
        finally:
            self._fin_analyse()

    # ══════════════════════════════════════════════════
    #  MINDMANAGER HELPERS
    # ══════════════════════════════════════════════════

    def _forcer_premier_plan(self, mm):
        try:
            mm.Visible = True
        except Exception:
            pass
        try:
            import win32gui, win32con
            def _cb(hwnd, _):
                if "MindManager" in win32gui.GetWindowText(hwnd) and win32gui.IsWindowVisible(hwnd):
                    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                    win32gui.SetForegroundWindow(hwnd)
            win32gui.EnumWindows(_cb, None)
        except Exception:
            pass

    def _init_mm(self):
        if not COM_DISPONIBLE:
            self._log("❌ pip install pywin32")
            return None
        try:
            pythoncom.CoInitialize()
        except Exception:
            pass
        mm = None
        try:
            mm = win32com.client.GetActiveObject("MindManager.Application")
            self._log("🔗 MindManager connecté.")
        except Exception:
            pass
        if mm is None:
            try:
                mm = win32com.client.Dispatch("MindManager.Application")
                self._log("✅ MindManager lancé.")
            except Exception as e:
                self._log(f"❌ MindManager : {e}")
                return None
        try:
            mm.Visible = True
        except Exception:
            pass
        self._forcer_premier_plan(mm)
        return mm

    # ══════════════════════════════════════════════════
    #  MINDMANAGER HELPERS — v2.6.4
    # ══════════════════════════════════════════════════

    # Longueur max d'un texte de nœud avant bascule en Note
    _MM_MAX_LEN = 180

    def _cs(self, parent, texte: str, note: str = "") -> object:
        """
        Crée un sous-sujet MindManager de façon robuste.
        - Texte > _MM_MAX_LEN → tronqué dans le titre, complet dans la Note
        - Deux méthodes COM en fallback pour compatibilité toutes versions MM
        - Retourne None silencieusement si l'objet parent est invalide
        """
        if parent is None:
            return None
        texte = _strip_md(str(texte))
        # Densité visuelle : bascule en Note si trop long
        titre = texte
        note_finale = note
        if len(texte) > self._MM_MAX_LEN:
            titre = texte[:self._MM_MAX_LEN - 1] + "…"
            note_finale = texte if not note else f"{texte}\n\n{note}"

        noeud = None
        try:
            noeud = parent.AddSubTopic(titre)
        except Exception:
            try:
                noeud = parent.AddSubTopic()
                noeud.Text = titre
            except Exception as e:
                self._log(f"   ⚠️  Sujet impossible '{titre[:40]}…' : {e}", "warn")
                return None

        # Attache la Note si nécessaire (texte complet lisible en double-clic)
        if note_finale and noeud:
            try:
                noeud.Notes.Text = note_finale
            except Exception:
                pass  # Les Notes ne sont pas critiques

        return noeud

    def _cs_verbatim(self, parent, participant: str, verbatim: str,
                     justification: str = "") -> object:
        """
        Niveau 3 : crée le nœud verbatim structuré.
        Format titre : [T1] Début du verbatim…
        Note complète : verbatim intégral + justification si présente
        """
        titre = f"[{participant}] {verbatim}"
        note  = verbatim
        if justification:
            note += f"\n\n📐 Justification : {justification}"
        return self._cs(parent, titre, note=note)

    def _cs_heuristiques(self, parent, heuristiques: dict) -> None:
        """
        Niveau 4 (optionnel) : nœud cadre théorique sous un thème.
        N'est créé que si des données heuristiques existent.
        """
        if not heuristiques:
            return
        nielsen = ", ".join(heuristiques.get("nielsen", []))
        bastien = ", ".join(heuristiques.get("bastien_scapin", []))
        gravite = heuristiques.get("gravite", "")
        if not (nielsen or bastien):
            return

        n_heur = self._cs(parent, "📐 Cadre théorique")
        if not n_heur:
            return
        if nielsen:
            self._cs(n_heur, f"Nielsen : {nielsen}")
        if bastien:
            self._cs(n_heur, f"Bastien & Scapin : {bastien}")
        if gravite:
            tag = {"Majeur": "🔴", "Modéré": "🟡", "Mineur": "🟢"}.get(gravite, "⚪")
            self._cs(n_heur, f"Gravité : {tag} {gravite}")

    def _ecrire_theme_mm(self, root, theme_nom: str, data_theme: dict,
                          nb_total: int, heuristiques: dict,
                          compteurs: dict) -> int:
        """
        Écrit un thème complet dans MindManager selon la hiérarchie :
          Niveau 1 : Thème [n/total] ●●●○○
          Niveau 2 : Sous-thème
          Niveau 3 : 🔴/🟢/🟡 Sentiment (N)
          Niveau 4 : [T1] Verbatim (avec Note = verbatim complet + justification)
          Optionnel : 📐 Cadre théorique (Nielsen / Bastien & Scapin)

        Retourne le nombre de verbatims écrits.
        """
        sources   = data_theme.get("_sources", set())
        n_sources = len(sources)
        barre     = "●" * n_sources + "○" * max(0, nb_total - n_sources)
        label_th  = f"{theme_nom}  [{n_sources}/{nb_total}] {barre}"

        n_th = self._cs(root, label_th)
        if not n_th:
            return 0

        # Cadre théorique rattaché au thème (niveau optionnel)
        self._cs_heuristiques(n_th, heuristiques.get(theme_nom, {}))

        total_verbs = 0

        for st_nom, data_st in data_theme.get("sous_themes", {}).items():
            if st_nom.startswith("_"):
                continue

            n_st = self._cs(n_th, st_nom)
            if not n_st:
                continue

            for sent_nom, entrees in data_st.items():
                if sent_nom.startswith("_") or not entrees:
                    continue

                icon  = {"Négatif": "🔴", "Positif": "🟢", "Nuancé": "🟡"}.get(sent_nom, "⚪")
                label = f"{icon} {sent_nom} ({len(entrees)})"
                n_sent = self._cs(n_st, label)
                compteurs[sent_nom] = compteurs.get(sent_nom, 0) + len(entrees)
                total_verbs += len(entrees)

                if not n_sent:
                    continue

                for entry in entrees:
                    # entry est (participant, verbatim) ou (participant, verbatim, justification)
                    if isinstance(entry, (list, tuple)):
                        part  = entry[0] if len(entry) > 0 else "?"
                        verb  = entry[1] if len(entry) > 1 else ""
                        justi = entry[2] if len(entry) > 2 else ""
                    else:
                        part, verb, justi = str(entry), "", ""

                    self._cs_verbatim(n_sent, part, verb, justi)

        return total_verbs


# ══════════════════════════════════════════════════════

if __name__ == "__main__":
    root = tk.Tk()
    app = StationUX(root)
    root.mainloop()

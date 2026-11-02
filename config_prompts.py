"""
config_prompts.py
Configuration centralisée des prompts et constantes
Station UX Audit v2.6.4
"""

import re
import unicodedata

# ══════════════════════════════════════════════════════
#  CONSTANTES GLOBALES
# ══════════════════════════════════════════════════════

APP_VERSION = "2.6.4"
APP_TITLE = "STATION UX AUDIT — v2.6.4"

# ══════════════════════════════════════════════════════
#  UTILITAIRES — Nettoyage markdown
# ══════════════════════════════════════════════════════

def _strip_md(texte: str) -> str:
    """Supprime la mise en forme markdown du texte."""
    if not texte:
        return ""
    t = str(texte)
    t = re.sub(r'\*{1,3}([^*]+)\*{1,3}', r'\1', t)
    t = re.sub(r'#{1,6}\s*', '', t)
    t = re.sub(r'`{1,3}([^`]*)`{1,3}', r'\1', t)
    t = re.sub(r'^\s*[-•]\s+', '', t, flags=re.M)
    t = re.sub(r'\n{3,}', '\n\n', t)
    return t.strip()


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


# ══════════════════════════════════════════════════════
#  PROMPTS STANDARDS
# ══════════════════════════════════════════════════════

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

PROMPT_NIELSEN = """\
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
"""

PROMPT_BASTIEN_SCAPIN = """\
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
"""

PROMPT_EMOTIONNEL = """\
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
"""

PROMPT_WCAG = """\
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
"""

PROMPT_DOCTORAL = """\
TU ES UN CHERCHEUR EXPERT EN USER RESEARCH ET ANALYSE SÉMANTIQUE.

OBJECTIF : Analyser un entretien utilisateur pour en extraire des insights ergonomiques structurés pour une soutenance de thèse.

RÈGLES D'EXTRACTION (Strictes) :
1. ANALYSE : Chaque verbatim doit être rattaché à une heuristique (Nielsen) ou un critère (Bastien & Scapin).
2. RÉSUMÉ : Normalise les sous-thèmes. Ne crée pas de sous-thèmes uniques pour un seul verbatim si un sous-thème proche existe déjà.
3. PRÉCISION : Si un participant exprime plusieurs points, traite-les séparément.
4. TRAÇABILITÉ : Le champ "participant" est obligatoire. Le champ "verbatim" est une copie brute exacte.

CONTRAINTES DE FORMAT (JSON STRICT) :
- Sortie uniquement en JSON. Aucun texte avant, aucun après, pas de markdown.
- Si le JSON est trop volumineux, PRIORISE les verbatims les plus critiques (freins majeurs).
- Échappe correctement tous les guillemets et caractères spéciaux.

TAXONOMIE (Unique, respecter strictement) :
[Navigation, Clarté du contenu, Confiance, Identité visuelle, Interactivité, Portfolio, Contact, Ergonomie, Performance, Satisfaction globale]

RÈGLES ANTI-SATURATION :
- Maximum 3 extraits par sous-thème (choisis les plus percutants).
- Pas de doublons sémantiques dans les sous-thèmes.
- Priorité : Frein fonctionnel > Compréhension > Confiance > Esthétique.
- Sentiment : "Négatif" si frein/frustration, "Positif" si satisfaction claire, "Nuancé" si ambigu.

FORMAT JSON ATTENDU :
{"themes":[{"theme":"Thème","sous_themes":[{"sous_theme":"Libellé court et concret","extraits":[{"sentiment":"Négatif","participant":"CODE","verbatim":"Texte exact","justification":"Nielsen H5 - Prévention des erreurs"}]}]}]}

ENTRETIEN :
"""

PRESETS = {
    "🔍 Analyse UX structurée (recommandé)": PROMPT_STANDARD,
    "🎯 Heuristiques de Nielsen": PROMPT_NIELSEN,
    "📐 Critères Bastien & Scapin": PROMPT_BASTIEN_SCAPIN,
    "💬 Ressenti émotionnel & charge cognitive": PROMPT_EMOTIONNEL,
    "♿ Accessibilité WCAG 2.1": PROMPT_WCAG,
    "🎓 Méga-Prompt Doctoral (recommandé thèse)": PROMPT_DOCTORAL,
}

# ══════════════════════════════════════════════════════
#  PROMPTS MODE MÉTHODOLOGIQUE
# ══════════════════════════════════════════════════════

PROMPT_PHASE0_ERGO = """\
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

PROMPT_PHASE1 = """\
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

PROMPT_PHASE2 = """\
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

PROMPT_PHASE3 = """\
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

PROMPT_NETTOYAGE_VERBATIMS = """\
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

PROMPT_HEURISTIQUES = """\
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
#  UTILITAIRES JSON
# ══════════════════════════════════════════════════════

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
                in_string = False
                resultat.append(c)
            continue
        resultat.append(c)
    s = "".join(resultat)

    # 4. Répare les JSON tronqués
    ouv_accolades  = s.count("{") - s.count("}")
    ouv_crochets   = s.count("[") - s.count("]")
    if ouv_crochets > 0 or ouv_accolades > 0:
        if s.rstrip()[-1] not in ('"', '}', ']'):
            s = s.rstrip().rstrip(',')
            s += '"'
        for _ in range(max(0, ouv_crochets)):
            s += "]"
        for _ in range(max(0, ouv_accolades)):
            s += "}"

    return s


# ══════════════════════════════════════════════════════
#  UTILITAIRES DE FUSION DE THÈMES
# ══════════════════════════════════════════════════════

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

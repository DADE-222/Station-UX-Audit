# Station-UX-Audit
Station UX Audit automatise l'analyse de corpus textuels pour Mind Manager grâce à l'IA en local.

Station UX Audit (v2.6.4)
Station UX Audit automatise l'analyse qualitative de corpus textuels (entretiens utilisateurs, retranscriptions, tests d'usabilité) grâce à l'intelligence artificielle en local.
L'application traite vos données textuelles en toute confidentialité, applique des cadres méthodologiques UX reconnus, puis génère automatiquement une carte mentale interactive dans Mind Manager ainsi qu'un rapport statistique complet au format HTML.

📌 Fonctionnalités principales
*Analyse automatisée de corpus textuels : Traitement par lots de fichiers d'entretiens au format .txt.
*Génération automatique de Mind Map : Injection directe des thèmes, sous-thèmes, sentiments (🔴 Négatif, 🟢 Positif, 🟡 Nuancé) et citations exactes (verbatims) dans Mind Manager.
*Visualisation HTML interactive : Export d'un rapport visuel incluant la répartition des sentiments et le top des thèmes récurrents (via des graphiques statistiques).

Double mode d'analyse :
*Mode Standard / Groupé : Analyse et fusion sémantique avancée (algorithme de Jaccard) pour consolider les thèmes similaires.
*Mode Méthodologique (5 phases) : Démarche ergo-scientifique complète (Lecture du terrain → Reformulation de la demande → Émergence de la problématique → Codage → Synthèse des convergences/divergences).
*Confidentialité totale : Exécution 100% locale sans envoi de données vers des serveurs externes.

🛠️ Prérequis et Dépendances
1. Système d'exploitation & Logiciels requis
*Windows (indispensable pour l'interfaçage COM avec Mind Manager).
*Python 3.10+ installé sur votre machine.
*Mind Manager (installé et activé sur Windows).

Serveur d'IA local :
*Ollama (Port 11434) — Recommandé pour les cartes graphiques NVIDIA / Intel.
*LM Studio (Port 1234) — Recommandé pour les cartes graphiques AMD Radeon ou pour charger des modèles GGUF sur mesure.

2. Modèles d'IA recommandés
*Mistral Nemo (12B) : Offre les meilleurs résultats lors des benchmarks d'extraction JSON et d'analyse sémantique.
*Autres alternatives compatibles : Llama 3.1 (8B), Gemma 2 (9B), Qwen 2.5.

3. Bibliothèques Python (requirements.txt)
*Le code repose sur trois bibliothèques externes principales :
*requests (>= 2.28.0) : Communication HTTP avec les API locales d'Ollama et LM Studio.
*pywin32 (>= 300) : Pilotage et automatisation COM de l'application Mind Manager.
*matplotlib (>= 3.5.0) : Génération et encodage des graphiques statistiques du rapport HTML.

🚀 Installation et Démarrage
*Cloner ou télécharger le dépôt sur votre ordinateur :
*git clone https://github.com/votre-compte/Station-UX-Audit.git
cd Station-UX-Audit

Installer les dépendances Python :
*pip install -r requirements.txt
(Ou manuellement : pip install requests pywin32 matplotlib)

Lancer le serveur d'IA local :
*Démarrez Ollama (ollama serve) ou activez le Local Server dans LM Studio.
*Assurez-vous d'avoir téléchargé votre modèle (ex: ollama pull mistral-nemo).

Exécuter l'application :
*python Station_UX_Audit_v263.pyw

⚙️ Configuration & Choix des Modèles
*Utilisateurs AMD Radeon : Privilégiez LM Studio pour une prise en charge fluide du GPU. Si vous utilisez Ollama sur AMD, utilisez l'onglet ⚡ GPU / Performance dans l'application pour injecter automatiquement les variables *d'environnement nécessaires (HSA_OVERRIDE_GFX_VERSION=10.3.0, HIP_VISIBLE_DEVICES=0).
*Utilisateurs NVIDIA / Intel : Ollama fonctionne directement out-of-the-box.
*Température : L'application règle par défaut la température à 0.1 pour garantir un respect strict des structures JSON et limiter les hallucinations.

📐 Protocoles & Grilles UX Intégrés
*L'application intègre plusieurs grilles et prompts méthodologiques préconfigurés :

🔍 Analyse UX structurée (Protocole recommandé par défaut).
🎯 Heuristiques de Nielsen (Les 10 critères d'usabilité).
📐 Critères de Bastien & Scapin (Guidage, charge de travail, contrôle explicite, etc.).
💬 Ressenti émotionnel & Charge cognitive (Analyse de la frustration et de la charge mentale).
♿ Accessibilité WCAG 2.1 (Critères de perception, d'utilisabilité et de compréhension).
🎓 Méga-Prompt Doctoral (Analyse approfondie pour recherche & thèse).
✍️ Prompt personnalisé : Rédigez votre propre consigne directement dans le constructeur de prompt.

Exportation d'un rapport visuel au format HTML.
<img width="1201" height="963" alt="image" src="https://github.com/user-attachments/assets/0ad42eff-a7aa-46dd-86e8-ab1e5bfe6629" />
<img width="1197" height="961" alt="image" src="https://github.com/user-attachments/assets/a8ab75d3-5889-4558-83b9-6854d39bfe73" />
<img width="1199" height="961" alt="image" src="https://github.com/user-attachments/assets/307df907-3660-4b0e-8c76-a341b3fc5f7f" />
<img width="1198" height="960" alt="image" src="https://github.com/user-attachments/assets/b90cf886-40c0-4e9d-8b02-103614e4c257" />

Exemple de résultat anonymisé.
<img width="1669" height="996" alt="image" src="https://github.com/user-attachments/assets/3998b55e-2a70-4eb4-963d-58e1f7e82837" />
<img width="1661" height="709" alt="image" src="https://github.com/user-attachments/assets/f912fbe3-7262-4f0f-8dd0-07bb45b0be97" />



❓ Foire Aux Questions (FAQ) & Dépannage

❌ Le serveur IA affiche "Serveur non joignable"
Solution : Vérifiez que l'URL et le backend sélectionnés correspondent à votre outil.
Ollama : http://localhost:11434
LM Studio : http://localhost:1234
Assurez-vous que le serveur local est bien démarré avant de cliquer sur 🔌 Tester & Charger modèles.

❌ Erreur "MindManager : erreur COM" ou "pywin32 non installé"
Solution : L'interfaçage automatique fonctionne uniquement sur Windows. Assurez-vous d'avoir exécuté pip install pywin32. Lancez Mind Manager une première fois manuellement avant d'exécuter l'analyse.

❌ L'IA retourne une erreur JSON ou une réponse vide
Solution : Station UX Audit intègre un système automatique de nettoyage et de réparation JSON en cas de troncature. Si l'erreur persiste, baissez la température du modèle ou passez sur un modèle plus performant (ex: Mistral Nemo ou Llama 3.1 8B).

❌ La VRAM reste à 0 Go sur GPU AMD avec Ollama
Solution : Rendez-vous dans l'onglet ⚡ GPU / Performance, cliquez sur ⚡ ACTIVER CPU + GPU, puis redémarrez Ollama et votre ordinateur pour que les variables système soient prises en compte.

❌ Le rapport HTML ne génère aucun graphique
Solution : Le module matplotlib est requis pour dessiner les graphiques. Installez-le avec pip install matplotlib.

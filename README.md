# Station-UX-Audit
Station UX Audit automatise l'analyse de corpus textuels pour Mind Manager grâce à l'IA en local.

Description complete du projet

Station UX Audit automatise ton analyse de corpus textuels pour Mind Manager grace a l'intelligence artificielle en local.
Le logiciel traite de grands volumes d'entretiens textuels pour ta recherche.
Il produit une carte mentale structuree et un rapport visuel en format HTML pour tes resultats.

Dependances techniques du code

Le code source repose sur plusieurs bibliotheques indispensables a installer:
* requests (version 2.28 ou superieure) pour communiquer avec tes serveurs locaux Ollama ou LM Studio via leurs API HTTP.
* pywin32 (version 300 ou superieure, systeme Windows requis) pour piloter ton application MindManager par automatisation COM.
* matplotlib (version 3.5 ou superieure) pour generer les graphiques statistiques integres a ton rapport HTML final.

Prérequis
Python installé sur ton ordinateur.
Un outil d'IA local tel que LLM Studio ou Ollama.

Installation
Télécharge le projet sur ta machine.
Installe Python si ce n'est pas déjà fait.
Lance l'application en local.

Configuration des modèles
Privilégie LLM Studio si tu possèdes une carte graphique AMD.
Utilise Ollama si tu es sous Nvidia ou Intel.
Le modèle Mistral Nemo offre d'excellents résultats lors des benchmarks. Réalise tes propres tests selon ton matériel et tes paramètres de température.

Protocoles d'analyse disponibles
L'application intègre plusieurs méthodes prêtes à l'emploi :
Heuristiques de Nielsen.
Critères de Bastien et Scapin.
Ressenti émotionnel et charge cognitive.
Accessibilité WCAG 2.1.
Taxonomie fermée UX.
Analyse UX structurelle.
Protocoles personnalisés rédigés par tes soins.
Module de test CPU et performance.

Fonctionnalités
Traitement de grands volumes de textes d'entretiens.
Génération automatique d'une carte mentale pour Mind Manager.

Exportation d'un rapport visuel au format HTML.
<img width="1201" height="963" alt="image" src="https://github.com/user-attachments/assets/0ad42eff-a7aa-46dd-86e8-ab1e5bfe6629" />
<img width="1197" height="961" alt="image" src="https://github.com/user-attachments/assets/a8ab75d3-5889-4558-83b9-6854d39bfe73" />
<img width="1199" height="961" alt="image" src="https://github.com/user-attachments/assets/307df907-3660-4b0e-8c76-a341b3fc5f7f" />
<img width="1198" height="960" alt="image" src="https://github.com/user-attachments/assets/b90cf886-40c0-4e9d-8b02-103614e4c257" />

Foire aux questions et resolution des pannes

Question: Ton serveur local affiche une erreur de connexion.
Reponse: Verifie ton backend selectionne dans l'interface. Ollama utilise le port 11434 et LM Studio utilise le port 1234. Lance ton application Ollama ou active ton serveur local dans LM Studio avant de lancer ton traitement.

Question: MindManager refuse de demarrer ou retourne une erreur COM.
Reponse: Cette fonctionnalite fonctionne uniquement sous Windows. Installe la bibliotheque pywin32 avec la commande pip install pywin32 dans ton terminal. Ouvre ton logiciel MindManager manuellement une premiere fois pour valider son activation.

Question: L'intelligence artificielle retourne un format JSON invalide.
Reponse: Le code integre une reparation automatique du texte et une relance automatique en cas d'echec. Choisis un modele performant d'au moins sept milliards de parametres comme Mistral Nemo. Garde une temperature basse autour de 0.1 pour stabiliser ta structure de sortie.

Question: La memoire VRAM reste a zero avec ta carte graphique AMD sous Ollama.
Reponse: Utilise le bouton approprie dans ton onglet dedie pour injecter les variables d'environnement necessaires dans le registre Windows. Redemarre ton ordinateur si le probleme persiste afin de forcer ton systeme a reconnaitre ton materiel.

Question: Ton rapport HTML ne se genere pas a la fin du traitement.
Reponse: Verifie l'installation du module matplotlib sur ta machine. Tape la commande pip install matplotlib dans ton environnement pour autoriser ton script a dessiner tes graphiques statistiques.



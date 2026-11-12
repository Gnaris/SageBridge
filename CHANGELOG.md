# Historique des versions

## 1.4.1 — 12/11/2026

- **Remplacer ou compléter une facture provisoire** : `POST /facture` accepte `"mode": "remplacer"` ou `"mode": "ajouter"` avec `"numero": "FA…"`. « remplacer » supprime toutes les lignes de la facture puis met celles de la commande ; « ajouter » met les lignes de la commande à la suite. Le numéro et le client de la facture ne changent pas ; la TVA suit le mode TVA de la fiche du client (0 % hors « Local »).
  - Refusé si la facture est validée, introuvable, ou si le numéro manque. `"simulation": true` indique ce qui serait fait sans rien modifier.
  - Après l'enregistrement, la facture est relue pour vérifier son nombre de lignes ; en cas d'écart, l'erreur est signalée et l'opération n'est jamais relancée automatiquement (pas de lignes en double).
  - En mode « ajouter », une ligne **identique** à une ligne déjà présente (même article, même prix HT, même TVA, même première ligne de description — désignation et référence) **augmente la quantité** de cette ligne au lieu d'en créer une nouvelle ; les lignes « Couleur xN » sont additionnées (« Doré x1 » + « Doré x3 » → « Doré x4 »). Une ligne avec remise n'est jamais fusionnée. `"fusionner": false` désactive la fusion. Contrôle après enregistrement : nombre de lignes **et** total HT. Réponse : `lignesFusionnees`.
  - Pas de contrôle anti-doublon dans ces deux modes ; la commande est quand même inscrite dans l'historique.
  - Réponse enrichie : `mode`, `lignesAvant`, `lignesApres`.
  - Utilisé par l'extension « Regroupement facture » 3.40.0 (menu « Créer dans Sage »).

## 1.4.0 — 03/11/2026

- **Enregistrer les factures en PDF** (nouveau bouton « Enregistrer les PDF… ») : on choisit le **mois** et l'**année**, SageBridge liste les factures et avoirs **validés** de ce mois dans Sage et indique pour chacun s'il est déjà enregistré dans le dossier des factures ou s'il manque. Ce dossier est **choisi par l'utilisateur** au premier usage (bouton « Choisir… », retenu ensuite ; aucun dossier imposé). Le bouton « Enregistrer les PDF manquants » crée ceux qui manquent, au modèle **« Format A4 Complet »** (modifiable), exactement comme à la main (Aperçu > Exporter en PDF > Exporter tout en PDF).
  - Rangement : `Recap\<année>\<n°>. <Mois>\Facture\<numéro>.pdf` (avoirs dans `Avoir`). Les dossiers manquants sont créés (« 2028 » puis « 11. Novembre » puis « Facture ») ; un dossier existant est réutilisé tel quel (« Novembre » comme « 11. Novembre », « Facture » comme « Factures »).
  - Fichier nommé par le **numéro de la pièce** (`FA10005913.pdf`, `AV10000115.pdf`). Une pièce est considérée comme déjà enregistrée si un PDF commençant par son numéro existe **n'importe où** sous le dossier des factures (fichiers renommés « FA10005662 Reçu le 06-10-2026.pdf » ou déplacés compris). **Aucun fichier n'est jamais écrasé.**
  - Les factures provisoires sont ignorées (leur numéro n'est pas définitif).
  - Pendant l'enregistrement (≈ 3 s par pièce), les aperçus de Sage s'ouvrent et se ferment tout seuls : ne pas utiliser Sage. Bouton « Arrêter » ; arrêt automatique après 3 erreurs de suite.
  - Technique : Sage n'offre pas d'export PDF par COM (destination « Ficher PDF » de la tâche d'impression → « Paramètres incorrects »). SageBridge ouvre l'aperçu par la tâche COM d'impression (sélection de la pièce, modèle, destination « Ecran »), puis pilote l'aperçu par messages Windows (sans toucher à la souris ni au clavier). Testé le 03/11/2026 : PDF identiques à l'octet près à l'export manuel (FA10005913), avoir AV10000115 correct.
- Réglages `recap` (dossier des factures, vide tant qu’il n’a pas été choisi) et `modele_pdf` dans `settings.ini`, modifiables depuis la fenêtre.

## 1.3.10 — 03/11/2026

- **« Marquer en soldé » : on choisit le mode de paiement.** En 1.3.9, chaque règlement était enregistré d'office en mode `VIR`. Désormais, avant de solder, une fenêtre propose la **liste des modes de paiement du dossier Sage** (table « Modes de paiement » : CB, VIR, PFS…) et l'utilisateur choisit celui des règlements. Le mode inscrit sur la plupart des factures concernées est présélectionné ; rien n'est enregistré sans choix. Le mode retenu figure dans la colonne « Règlement Sage » et dans le journal.
- Nouvelle route `GET /modes` (lecture seule) : `[{ code, intitule }]`. `GET /factures` renvoie aussi `modeReglement` (mode inscrit sur la facture) ; `POST /solder` accepte `mode`.

## 1.3.9 — 03/11/2026

- **Rapprochement des rapports de paiement** (nouveau bouton « Rapport de paiement… ») : on choisit l'export Excel des paiements Paris Fashion Shop (`Payment_Export_….xlsx`), SageBridge le lit ligne par ligne et retrouve la facture Sage de chaque paiement, avec une barre de progression. Il produit à côté du fichier une copie « … - rapproché.xlsx » avec 6 colonnes en plus : Rapprochement (Trouvée en vert, À vérifier en jaune, Introuvable en rouge, Déjà soldée en gris), N° facture Sage, Client Sage, HT facture Sage, Écart HT, Remarque. **Rien n'est modifié dans Sage** (lecture seule).
  - Règle : facture **validée**, **non soldée**, au **même nom** (Société de la facture ou nom du client, sans accents ni forme juridique) et au **même total HT** que le « Montant validé HT » du rapport (à 1 centime près). La commission PFS n'entre pas en compte. Si le n° de commande figure dans l'historique de SageBridge, la facture de cette commande est prise en priorité.
  - Les factures provisoires sont ignorées et les factures soldées ne sont pas retraitées (signalées « Déjà soldée »).
  - Si rien ne correspond, la remarque liste les autres factures non soldées à ce nom (avec l'écart), les avoirs non soldés, une facture égale au montant **commandé** (avoir à faire ?), une combinaison facture − avoir égale au montant validé, ou une facture du même montant sous un autre nom.
  - Testé sur l'export du 03/11/2026 (53 lignes) : 52 trouvées, 1 à vérifier (facture au montant commandé, avoir de 3,80 € HT à faire).
- **Marquer les factures trouvées en soldé** : à la fin du rapprochement, un bouton enregistre dans Sage un **règlement client par facture « Trouvée »** (moteur COM `WgcEngRegCli`) : mode de paiement `VIR` (corrigé en 1.3.10), date du jour, montant = reste dû TTC, référence = n° de commande. Confirmation demandée avant ; les factures sont relues juste avant (une facture soldée, devenue provisoire ou dont le HT a changé entre-temps est ignorée) ; arrêt après 3 erreurs de suite ; vérification finale que le statut Sage est bien « Soldé ». Le fichier « … - rapproché.xlsx » est réécrit avec une colonne « Règlement Sage » (n° du règlement ou motif). Les lignes À vérifier / Introuvable / Déjà soldée ne sont jamais touchées. Testé le 03/11/2026 sur une facture réelle : règlement créé, statut passé à « Soldé ».
- Nouvelle route `POST /solder` `{ facture, client, montant, reference, simulation }` : **simulation par défaut** (seul `"simulation": false` enregistre).
- Nouvelle route `GET /factures` (lecture seule) : pièces clients non soldées (`?tout=1` pour toutes) avec n°, type, client, société, totaux, solde dû, validée, soldée, statut.
- Compilation : références `System.Xml`, `System.Xml.Linq`, `System.IO.Compression` ajoutées dans `build.cmd`.

## 1.3.8 — 02/11/2026

- **Rafales de factures** : plus de risque de facture « fantôme » créée dans Sage sans trace côté SageBridge quand plusieurs commandes sont envoyées coup sur coup.
  - **Avant** : le délai de 2 min était compté depuis l'arrivée de la demande, pas depuis le début effectif du traitement. Avec 2-3 factures qui se chevauchaient en file (dossier en Données Mobiles où chaque création prend ~60 s), la 3ᵉ pouvait voir son HTTP expirer avant même que le worker ne la traite. L'action continuait en arrière-plan, la facture était bien créée dans Sage, **mais** sans ligne « Facture X créée » dans le journal ni enregistrement dans `factures.csv` → l'anti-doublon de SageBridge ne détectait pas la commande à la retentative de l'extension, et le client pouvait se retrouver avec une facture dupliquée.
  - **Après** : les 2 min sont comptées à partir du début du traitement effectif (quand le worker prend l'action en main). L'historique et le journal sont maintenant écrits côté worker, avant même que l'HTTP ne reprenne — même si l'HTTP a abandonné sur un timeout, le journal reflète ce que Sage a réellement fait, et toute retentative de l'extension est bloquée en 409 Doublon. Garde-fou de 15 min sur l'attente en file pour éviter un blocage éternel. Incident observé le 30/10/2026 chez ISSYMA sur une rafale de 7 factures (dossier Données Mobiles de 11 000 fiches) : la facture FG1733 avait été créée dans Sage sans trace côté SageBridge.

## 1.3.7 — 02/11/2026

- **Clients DOM/TOM derrière un label de groupe** : fiche créée sans erreur quand la boutique envoie « France DOM/TOM » (ou tout label d'outre-mer non reconnu) comme pays. SageBridge résout maintenant le territoire à partir du code postal : 971 → Guadeloupe, 972 → Martinique, 973 → Guyane, 974 → La Réunion, 975 → Saint-Pierre-et-Miquelon, 976 → Mayotte, 97133 → Saint-Barthélemy, 97150 → Saint-Martin, 986 → Wallis-et-Futuna, 987 → Polynésie française, 988 → Nouvelle-Calédonie. Avant, le code pays restait vide et Sage refusait la fiche (« La rubrique "Code Pays" doit être remplie »). Un pays reconnu n'est jamais écrasé — une facture « France » + CP 97400 reste FR. Incident observé le 02/11/2026, client « INTERIEUR D AUJOURD HUI ».

## 1.3.6 — 30/10/2026

- **Clients étrangers avec accents dans le nom du pays** : fiche créée sans erreur quand le pays est aux États-Unis, en Grèce, en Côte d'Ivoire, au Sénégal, au Brésil, au Monténégro, etc. SageBridge écrivait le nom du pays avec ses accents (« États-Unis »), que la table pays interne de Sage 50 stocke sans accents (« Etats-Unis ») — Sage ne reconnaissait pas le nom, effaçait silencieusement le code pays associé, et refusait l'enregistrement avec « La rubrique "Code Pays" doit être remplie ». Correction : nom écrit sans accents, et écrit **avant** le code pays (Sage recouple nom↔code quand on écrit le nom — en écrivant d'abord le nom puis le code, le code reste la dernière valeur posée). Incident observé le 30/10/2026 sur la commande KJ25KJNK, client « Accent Piece LLC » (US).

## 1.3.5 — 30/10/2026

- **Cache clients local** : SageBridge n'a plus besoin de parcourir toute la table des clients de Sage à chaque facture. Au premier appel sur un dossier il balaie la table une fois et enregistre un index dans `%APPDATA%\SageBridge\clients-<Dossier>.cache` ; les factures suivantes utilisent cet index. Le cache est reconstruit automatiquement si une fiche est ajoutée ou supprimée côté Sage (variation du nombre de fiches). Impact mesuré sur un dossier en Données Mobiles de 11 000 fiches : résolution client ~40 s → < 1 s, soit une facture d'environ 42 s à ~3 s. Les clients créés par SageBridge sont ajoutés au cache immédiatement. Si une fiche est modifiée **dans Sage** sans ajout/suppression, supprimer le fichier `.cache` force la reconstruction.

## 1.3.4 — 29/10/2026

- **Dossiers en mode Données Mobiles (ou ouverts par un serveur Sage)** : on peut désormais coller dans Réglages > Dossier Sage 50 le « Chemin d'accès » affiché par Sage (Dossier > Options > Informations), qui commence par `!Srv!` (ex. `!Srv!<Smart>/S:/Données Ciel/Dossiers/…/X.SgWgc`). SageBridge le transmet tel quel à Sage au lieu de le refuser comme fichier introuvable.
- La recherche automatique des dossiers ignore les anciennes copies `…_OLD.SgWgc` laissées par Sage lors d'une conversion (Sage refuse de les ouvrir : « Il est impossible d'ouvrir une base locale »).
- Message d'erreur plus juste quand Sage refuse d'ouvrir le dossier : copie `_OLD`, dossier servi par Sage (Données Mobiles) ou cas général. Avant, tout refus accusait l'autorisation d'accès COM, même quand elle était active.

## 1.3.3 — 29/10/2026

- `/articles` retente automatiquement une fois si Sage s'est reconnecté entre-temps (« Le serveur RPC n'est pas disponible »). Même garde-fou que la création de factures depuis la 1.1.0.

## 1.3.2 — 29/10/2026

- L'admin de la boutique (`beliandjolie.com`, `issyma.fr`) peut désormais interroger SageBridge pour afficher la liste des articles Sage dans le mapping « Code article Sage » d'une catégorie. Pas d'exposition supplémentaire : chaque admin tourne dans le navigateur de son propriétaire, SageBridge reste accessible uniquement en local sur `127.0.0.1`.

## 1.3.1 — 29/10/2026

- **Correction** : la route `/articles` échouait sur toutes les installations (« Table Art introuvable »). L'id de la table est maintenant la bonne valeur vérifiée live sur Sage 50 v6.71 (97002, nom interne « Article »).

## 1.3.0 — 29/10/2026

- Nouvelle route **`GET /articles`** : liste tous les articles du dossier Sage (code + nom court). L'extension s'en sert pour peupler le choix d'article à la volée quand un produit n'a pas de code renseigné côté boutique — plus besoin de connaître les codes par cœur.
- L'identifiant de la table articles est **résolu automatiquement** (scanne 97000→97012 et retient celle dont le nom interne est « Art ») : marche sur tous les dossiers Sage sans config manuelle.

## 1.2.3 — 20/10/2026

- Le champ **Référence** de la facture reste vide : la référence envoyée avec la commande (ex. « Faire 1234ABCD ») n'est plus écrite. Option pour la réactiver dans Réglages (« Écrire la référence de la commande… »).

## 1.2.2 — 20/10/2026

- **Correction importante** : en quittant (ou pendant une mise à jour), SageBridge relâchait le dossier Sage ouvert par COM, ce qui pouvait figer Sage 50. Il s'arrête désormais sans toucher à la connexion COM, immédiatement (l'installateur de mise à jour n'attend plus).
- Installateur : la touche Entrée valide toujours l'étape en cours (elle pouvait déclencher « Annuler »).

## 1.2.1 — 20/10/2026

- Fenêtre : distingue « à jour » de « aucune version publiée » ; `/ping` indique la dernière version publiée, l'état des mises à jour automatiques et si le programme est la copie installée.
- Publication : récupération fiable du compte GitHub enregistré, envoi du code par l'API, nouvel essai automatique des envois interrompus, reprise d'une publication inachevée.
- Première version livrée par la mise à jour automatique (test de bout en bout de la 1.2.0).

## 1.2.0 — 20/10/2026

Première version distribuée avec un installateur.

- **Installateur Windows** `SageBridge-Setup.exe` : assistant (détection de Sage 50, dossier, raccourcis Bureau et menu Démarrer, démarrage avec Windows), installation sans droits administrateur, inscription dans *Paramètres > Applications* / *Programmes et fonctionnalités* avec désinstalleur.
- **Mises à jour automatiques** depuis GitHub : vérification au démarrage puis toutes les heures, téléchargement contrôlé (empreinte SHA-256), installation et redémarrage automatiques quand aucune facture n'est en cours. Bouton « Mises à jour… » et réglage pour désactiver l'installation automatique.
- **Nouveau logo**.
- **Sécurité** : seules les extensions de navigateur et les outils locaux peuvent appeler SageBridge ; une page web quelconque ne peut plus créer de facture ni lire la liste des clients (contrôle de l'origine, CORS restreint, taille de requête limitée).
- **Partageable** : plus aucun réglage propre à un dossier (dossier Sage choisi au premier lancement parmi ceux trouvés sur le PC, client par défaut et article de test réglables).
- **Pays** : table complète des 27 États membres de l'UE (noms français, anglais, langue locale, codes ISO 2 et 3, « EL » pour la Grèce), DOM-TOM et principaux pays hors UE ; message quand un pays n'est pas reconnu.
- **Robustesse** : rubriques de la table clients résolues par leur nom (avec repli sur les identifiants connus), contrôle du nom de la table, instance unique, journal limité à 5 Mo, gestion des erreurs inattendues.
- **Sage figé** : seules les instances de Sage *sans fenêtre* (COM orphelin) de plus d'une minute sont fermées, jamais un Sage ouvert par l'utilisateur.
- Mode **simulation** (`"simulation": true`) : vérifie client, TVA et lignes sans rien créer.
- Historique des factures affiché dans la fenêtre, indicateur « Sage 50 ouvert / fermé » rafraîchi toutes les 5 secondes, bouton Documentation.

## 1.1.0 — octobre 2026

- Recherche du client existant (nom, adresse, code postal, ville, pays — sans accents ni casse) et création automatique de la fiche sinon.
- Mode TVA du client selon le pays : Local (France, Monaco), CEE (UE), Hors CEE (reste du monde) ; lignes à 0 % hors France.
- Contrôle du SIRET (clé de Luhn) et du N° TVA français avant enregistrement.

## 1.0.0 — octobre 2026

- Création de factures provisoires dans Sage 50 Gestion commerciale par l'interface COM, à partir d'un JSON reçu sur `http://127.0.0.1:8765/facture`.
- Anti-doublon par commande, journal, fenêtre de suivi et icône dans la zone de notification.

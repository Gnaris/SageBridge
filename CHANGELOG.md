# Historique des versions

## 1.3.9 — 03/11/2026

- **Rapprochement des rapports de paiement** (nouveau bouton « Rapport de paiement… ») : on choisit l'export Excel des paiements Paris Fashion Shop (`Payment_Export_….xlsx`), SageBridge le lit ligne par ligne et retrouve la facture Sage de chaque paiement, avec une barre de progression. Il produit à côté du fichier une copie « … - rapproché.xlsx » avec 6 colonnes en plus : Rapprochement (Trouvée en vert, À vérifier en jaune, Introuvable en rouge, Déjà soldée en gris), N° facture Sage, Client Sage, HT facture Sage, Écart HT, Remarque. **Rien n'est modifié dans Sage** (lecture seule).
  - Règle : facture **validée**, **non soldée**, au **même nom** (Société de la facture ou nom du client, sans accents ni forme juridique) et au **même total HT** que le « Montant validé HT » du rapport (à 1 centime près). La commission PFS n'entre pas en compte. Si le n° de commande figure dans l'historique de SageBridge, la facture de cette commande est prise en priorité.
  - Les factures provisoires sont ignorées et les factures soldées ne sont pas retraitées (signalées « Déjà soldée »).
  - Si rien ne correspond, la remarque liste les autres factures non soldées à ce nom (avec l'écart), les avoirs non soldés, une facture égale au montant **commandé** (avoir à faire ?), une combinaison facture − avoir égale au montant validé, ou une facture du même montant sous un autre nom.
  - Testé sur l'export du 03/11/2026 (53 lignes) : 52 trouvées, 1 à vérifier (facture au montant commandé, avoir de 3,80 € HT à faire).
- **Marquer les factures trouvées en soldé** : à la fin du rapprochement, un bouton enregistre dans Sage un **règlement client par facture « Trouvée »** (moteur COM `WgcEngRegCli`) : mode de paiement `VIR`, date du jour, montant = reste dû TTC, référence = n° de commande. Confirmation demandée avant ; les factures sont relues juste avant (une facture soldée, devenue provisoire ou dont le HT a changé entre-temps est ignorée) ; arrêt après 3 erreurs de suite ; vérification finale que le statut Sage est bien « Soldé ». Le fichier « … - rapproché.xlsx » est réécrit avec une colonne « Règlement Sage » (n° du règlement ou motif). Les lignes À vérifier / Introuvable / Déjà soldée ne sont jamais touchées. Testé le 03/11/2026 sur une facture réelle : règlement créé, statut passé à « Soldé ».
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

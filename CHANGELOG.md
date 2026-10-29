# Historique des versions

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

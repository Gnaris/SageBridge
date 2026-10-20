<p align="center"><img src="logo/SageBridge-256.png" width="128" alt="Logo SageBridge"></p>

<h1 align="center">SageBridge</h1>

<p align="center">Crée vos factures <b>directement dans Sage 50 Gestion commerciale</b>, sans import CSV, grâce à l'interface COM de Sage.</p>

<p align="center"><a href="../../releases/latest"><b>⬇ Télécharger l'installateur (SageBridge-Setup.exe)</b></a></p>

---

## À quoi ça sert

SageBridge est un petit programme Windows qui reste ouvert à côté de Sage 50. Il reçoit des commandes (par exemple depuis l'extension Chrome *Regroupement facture*, ou tout outil capable d'envoyer du JSON) et, pour chacune :

1. **retrouve le client** dans le dossier Sage (nom, adresse, code postal, ville, pays — comparés sans accents ni majuscules) ou **crée sa fiche** ;
2. choisit le **régime de TVA** selon le pays : *Local* (France, Monaco), *CEE* (Union européenne), *Hors CEE* (reste du monde, DOM-TOM compris) — lignes à 0 % hors France ;
3. crée la **facture provisoire** (non validée) avec ses lignes, prix et descriptions ;
4. refuse les **doublons** (même commande envoyée deux fois).

La facture reste modifiable dans Sage jusqu'à sa validation.

## Installation

1. Téléchargez **`SageBridge-Setup.exe`** depuis la [dernière version](../../releases/latest) et lancez-le.
   Le fichier n'est pas signé : si Windows affiche *« Windows a protégé votre ordinateur »*, cliquez sur *Informations complémentaires* puis *Exécuter quand même*.
2. Suivez l'assistant (aucun droit administrateur nécessaire).
3. Dans Sage 50, autorisez l'accès COM : menu **Utilitaires > Mode avancé > Divers > Autorisation accès COM** (autorisation liée au compte Sage, à renouveler environ tous les 6 mois).
4. Ouvrez Sage 50 sur votre dossier, puis dans SageBridge : **Réglages…** (dossier Sage, client par défaut) et **Facture de test…**.

**Prérequis** : Windows 10 ou 11, Sage 50 Gestion commerciale (ex-Ciel) installé sur le même PC et **ouvert** pendant l'utilisation.

Désinstallation : *Paramètres > Applications > SageBridge > Désinstaller* (ou *Panneau de configuration > Programmes et fonctionnalités*). Les réglages et l'historique (`%APPDATA%\SageBridge`) peuvent être conservés.

## Mises à jour

SageBridge vérifie cette page de versions **au démarrage puis toutes les heures**. Quand une nouvelle version est publiée, il télécharge l'installateur, vérifie son empreinte SHA-256, attend qu'aucune facture ne soit en cours, s'installe et redémarre tout seul. L'installation automatique peut être désactivée dans *Réglages* (le bouton *Mises à jour…* reste disponible).

## Utilisation par programme (API locale)

SageBridge écoute uniquement sur `http://127.0.0.1:8765` (ce PC). Les appels sont acceptés depuis une extension de navigateur ou un outil local (PowerShell, script) ; les pages web ordinaires sont refusées.

| Méthode | Adresse | Rôle |
|---|---|---|
| `GET` | `/ping` | état : version, Sage ouvert, dossier |
| `POST` | `/facture` | crée une facture (ou la simule avec `"simulation": true`) |
| `GET` | `/clients` | liste des clients du dossier |

Exemple :

```powershell
$facture = @{
  site = "boutique"; orderNo = "CMD-1042"; reference = "Boutique CMD-1042"
  customer = @{ company = "Bijoux Martin"; address1 = "12 rue des Lilas"; postalCode = "69003"; city = "Lyon"; country = "France"; email = "contact@exemple.fr" }
  lines = @(@{ code = "ART0001"; qty = 3; puht = 4.5; tva = 20; description = "Collier doré" })
} | ConvertTo-Json -Depth 4
Invoke-RestMethod -Method Post -Uri http://127.0.0.1:8765/facture -Body $facture -ContentType "application/json; charset=utf-8"
```

Réponse : `{ ok, numero, totalHT, totalTTC, clientCode, clientCreated, clientModeTVA, messages, … }` — HTTP 200 (créée), 409 (commande déjà envoyée, renvoyer avec `"force": true`), 500 (refus de Sage, détail dans `error`).

Le détail complet (fonctionnement du COM de Sage 50, règles, erreurs rencontrées et corrigées) est dans la **[documentation PDF](docs/SageBridge%20-%20Documentation.pdf)**.

## Compiler soi-même

Aucun outil à installer : double-cliquez sur **`build.cmd`**. Il compile `SageBridge.cs` puis `Setup.cs` avec le compilateur C# du .NET Framework 4 présent dans Windows, et produit `bin\SageBridge-Setup.exe` et son empreinte `.sha256`.

Publier une nouvelle version : augmentez `AssemblyVersion` dans `SageBridge.cs`, ajoutez une section au `CHANGELOG.md`, lancez `build.cmd` puis **`publier-version.cmd`** (utilise le compte GitHub enregistré dans Git sur le PC, crée la release et y joint l'installateur).

| Fichier | Contenu |
|---|---|
| `SageBridge.cs` | le programme (serveur local, accès COM, fenêtre, mises à jour) |
| `Setup.cs` | l'installateur / désinstalleur |
| `app.manifest` | exécution sans droits administrateur, affichage haute résolution |
| `logo/` | icône et logo |
| `docs/` | documentation PDF |

---

SageBridge n'est ni édité ni approuvé par Sage. « Sage 50 » et « Ciel » sont des marques de leurs propriétaires respectifs.

# CLAUDE.md — SageBridge

Ce fichier guide Claude (et tout développeur) pour modifier SageBridge sans tout redécouvrir. Lire aussi `docs/SageBridge - Documentation.pdf` (référence COM complète, chapitre 8 = toutes les erreurs rencontrées et leur correction) et `CHANGELOG.md`.

## Le projet en bref

SageBridge est un programme Windows (C# WinForms, .NET Framework 4) qui crée des **factures provisoires** dans **Sage 50 Gestion commerciale** (ex-Ciel, testé en version 6.71) par son **interface COM**, sans import CSV. Il reste ouvert à côté de Sage, écoute en local sur `http://127.0.0.1:8765/`, reçoit une commande en JSON, **retrouve ou crée le client**, applique le **régime de TVA** selon le pays, crée la facture et renvoie son numéro.

- Client principal : l'extension Chrome **« Regroupement facture »** de Boris (dossier `Bureau\extension Google\Regroupement facture\`, bouton vert « Créer dans Sage », voir son propre `CLAUDE.md`, sections SageBridge). L'extension **ne fait pas partie** de ce dépôt (elle contient des codes articles et une clé API propres à Boris).
- Dépôt **public** : `github.com/Gnaris/SageBridge` (séparé du dépôt privé `beli-jolie`). Ne jamais y mettre de clé, de mot de passe, de liste de clients ou de données propres à un dossier Sage.
- Distribué par l'installateur `SageBridge-Setup.exe`, mis à jour automatiquement depuis les releases GitHub.

## Fichiers

| Fichier | Rôle |
|---|---|
| `SageBridge.cs` | tout le programme (un seul fichier, ~2 200 lignes) |
| `Setup.cs` | installateur / mise à jour silencieuse / désinstalleur (un seul exe) |
| `app.manifest` | `asInvoker` (jamais d'élévation), compatibilité Windows 10/11, `dpiAware` |
| `build.cmd` | compile les deux exe dans `bin\`, calcule l'empreinte SHA-256, recopie l'installateur dans `..\Installateur\` et le PDF dans `..\Documentation\` |
| `publier-version.cmd` + `outils\publier-version.ps1` | publie la version compilée sur GitHub (dépôt, code, release, fichiers joints) |
| `logo\` | `SageBridge.ico` (16→256 px, versions simplifiées ≤ 32 px), PNG 256/512/1024 |
| `docs\SageBridge - Documentation.pdf` | documentation utilisateur + référence COM (embarquée dans l'installateur) |
| `README.md` | page d'accueil GitHub |
| `CHANGELOG.md` | historique ; la section `## X.Y.Z` devient la note de la release |
| `bin\` | sorties de compilation (ignoré par Git) |

Chez Boris, le dossier de travail est `C:\Users\Admin\Desktop\SageBridge\` : `Code source\` (= ce dépôt), `Installateur\` (exe à partager), `Documentation\` (PDF), `LISEZ-MOI.txt`. L'ancien `Téléchargements\SageBridge` (version 1.1) n'est plus utilisé.

## Compiler

- **Aucun outil à installer** : `build.cmd` utilise `%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\csc.exe` (fourni avec Windows). Résultat : `bin\SageBridge.exe`, `bin\SageBridge-Setup.exe`, `bin\SageBridge-Setup.exe.sha256`, journal `build.log` (contient `OK X.Y.Z` ou `ECHEC` + erreurs).
- Ce compilateur n'accepte que **C# 5** : pas d'interpolation `$"…"`, pas de `?.`, pas de `nameof`, pas de `=>` sur les membres, pas de `out var`, pas de tuples. Utiliser `delegate { … }`, `string.Format` / concaténation.
- Références : `System`, `System.Core`, `System.Windows.Forms`, `System.Drawing`, `Microsoft.CSharp` (pour `dynamic`), `System.Web.Extensions` (JSON, `JavaScriptSerializer`).
- `build.cmd` lit la version dans `bin\SageBridge.exe` et génère `bin\SetupVersion.cs` : **la seule source de vérité de la version est `AssemblyVersion` / `AssemblyFileVersion` dans `SageBridge.cs`**. L'installateur lit aussi la version de l'exe embarqué (`Assembly.ReflectionOnlyLoad`).
- Ressources embarquées dans l'installateur : `SageBridge.exe`, `logo.png` (= `logo\SageBridge-256.png`), `Documentation.pdf` (si `docs\…pdf` existe).
- Vérification rapide hors Windows : `mcs -langversion:5 …` (mono) attrape la plupart des erreurs de syntaxe, **mais pas toutes** : csc a refusé un `new[] { … }` dont les éléments étaient `dynamic` (→ `dynamic[]`, pas `string[]`) que mcs acceptait. Toujours finir par `build.cmd` et lire `build.log`.
- Un exe en cours d'exécution est verrouillé : fermer SageBridge avant de recompiler (l'installateur, lui, le ferme tout seul).

## Architecture de `SageBridge.cs`

| Classe | Rôle |
|---|---|
| `Program` | point d'entrée ; **instance unique** (mutex `SageBridge-instance-unique` ; une 2ᵉ instance signale l'événement `SageBridge-afficher-fenetre` et quitte) ; gestion globale des exceptions ; `Version` lue dans l'assembly ; argument `/apres-maj=ancienne` après une mise à jour |
| `Settings` | `%APPDATA%\SageBridge\settings.ini` (`dossier`, `port`, `client`, `article_test`, `maj_auto`, `reference_facture`, `depot`) ; journal `journal.log` (archivé à 5 Mo en `journal.ancien.log`) ; `FindDossiers()` cherche les `*.SgWgc` (hors points de restauration) |
| `InvoiceLine`, `CustomerInfo`, `InvoiceRequest`, `InvoiceResult` | modèle d'une facture et du résultat |
| `Norm` | normalisation sans accents/casse, table des pays (ISO 2, ISO 3, noms FR/EN/locaux), `VatModeForCountry`, `RtfToText` |
| `SageWorker` | **tout l'accès COM**, sur **un seul thread STA** alimenté par une file (`BlockingCollection<Action>`) ; `Run<T>` refuse si Sage n'est pas ouvert et borne chaque opération à 2 min ; recherche/création client ; création facture ; `ListClients`, `Diag` |
| `UpdateInfo`, `Updater` | mises à jour par les releases GitHub |
| `History` | `factures.csv` : historique + clé anti-doublon `site|orderNo` |
| `Server` | `HttpListener` sur `127.0.0.1` et `localhost` ; contrôle d'origine ; routes `/ping`, `/facture`, `/clients`, `/diag` |
| `MainForm` | fenêtre (état, indicateur Sage ouvert/fermé rafraîchi toutes les 5 s, historique, journal, boutons), icône de notification, minuterie des mises à jour |
| `Ui2`, `DossierPicker`, `SettingsForm`, `TestInvoiceForm` | petites fenêtres ; `Ui2.Px()` met les tailles à l'échelle DPI |

Les fenêtres sont construites en code (pas de designer). Le programme est déclaré DPI-aware : toute taille en pixels doit passer par `S()` (MainForm) ou `Ui2.Px()`.

## L'interface COM de Sage 50 — ce qu'il faut savoir

Aucune documentation officielle : tout a été établi avec `WGC.tlb` (dossier d'installation, bibliothèque « WGestCom ») et `CamCOM.tlb` (`Common Files\Ciel`, « Ciel Standard Library »), puis vérifié par des appels réels.

- **Connexion** : `Type.GetTypeFromProgID("CielGestionCommerciale.Application")` → `Activator.CreateInstance` (CLSID inscrit `{10017AE8-0000-84DE-AEB4-0000E80C11B9}`, serveur `WGC.exe` **32 bits**). La coclasse déclarée dans la tlb (`…84D8…`) n'est pas inscrite. `GetActiveObject` renvoie un objet inutilisable (E_NOTIMPL) → toujours `CreateInstance`. Quand Sage est ouvert, l'instance créée travaille dans le Sage de l'utilisateur.
- `app.OpenFile(cheminDossier.SgWgc, "", "", 0)` → document ; `doc.Name` = nom du dossier.
- **Moteurs** `doc.CreateEngine(id)` : 97001 pièce client (`WgcEngPieceCli`), 97002 ligne de pièce client, 97004 fiche client (`WgcEngClient`). Types de pièce : 1 devis, 2 commande, 3 BL, **4 facture**, 5 avoir, 6 acompte.
- **Facture** : `Create(4)`, `Client = code`, `reference`, `Remarque` ; par ligne `RazLine`, `Article`, `Quantite`, `PUHT`, `TauxTVA`, `Description`, `piece.AddLine(ligne)` ; lire `Numero`, `TotalHT`, `TotalTTC` **avant** `Save(verif)`.
- **Messages** : `app.CreateVerif()` ; après `Save` : `verif.Total`, `verif.Desc(k)` (k = 1…Total), `verif.Reset()`.
- **Tables** : `doc.OpenTable(id)` avec l'id **sur 24 bits bas** de l'énumération (`kWgcTableIdCli = 0x8C017AEB` → `0x017AEB` = 97003) ; `TableName` (« Cli »), `NbN`, `First()`/`Next()`, `GetVal(idRubrique)` ; `Fields` numérotée **à partir de 1** (`Item(0)` = null) avec `Id`, `Name`, `Title`. Rubriques clients utiles : Code 1, Nom 2, FacSociete 3, FacAdresse1 5, FacAdresse2 6, FacCP 8, FacVille 9, FacCtyCode 10 (ISO **3**), FacCtyName 11, Liv… 12–20, FacTel 25, ModeTVA 32, NII 33 (N° TVA), Email 54, SIRET 56, LivTel 71, NomCtc 88, IsProspect 129, TypeCli 143. SageBridge les résout **par nom** au chargement (`LoadClientFields`) avec ces valeurs en secours.
- **Fiche client** : `Create()`, `Exist(code)`, `Code`, `Name`, `SetFieldIdValStr(id, valeur, 0, 0)`, `Save(verif, ref code)`.

### Règles impératives (chacune vient d'un vrai incident)

1. **Sage 50 doit être ouvert avec sa fenêtre.** Sans fenêtre, l'activation COM lance un `WGC.exe` invisible qui ne répond jamais et empêche Sage de s'ouvrir. `SageWindowOpen()` vérifie un processus `WGC` avec `MainWindowHandle`.
2. **Ne jamais relâcher `app` ni `doc` tant que Sage est ouvert, et ne rien appeler en COM à la fermeture.** Le 20/10/2026, la fermeture de SageBridge (pendant une mise à jour) faisait `FinalReleaseComObject(doc/app)` → **Sage 50 s'est figé** et a dû être tué. Depuis 1.2.2, `OnFormClosing` arrête le processus avec `Process.GetCurrentProcess().Kill()` (même effet qu'un `taskkill /F`, sans conséquence pour Sage). Seuls les objets de travail (moteurs, tables, vérificateurs) sont libérés après usage. `Reset()` (qui relâche tout) ne sert qu'après une erreur COM pour se reconnecter, et quand on change de dossier.
3. **Ne jamais tuer un Sage avec fenêtre**, même figé : seuls les `WGC.exe` sans fenêtre de plus de 60 s, quand aucun Sage visible n'est ouvert, sont fermés (`KillOrphans`). Si Sage est figé, c'est à l'utilisateur de le fermer.
4. **Un seul thread STA** pour tous les appels COM (sinon erreurs aléatoires).
5. Après un redémarrage de Sage, la connexion lève « Le serveur RPC n'est pas disponible » (0x800706BA) : `Create` fait `Reset()` puis réessaie une fois.
6. `AddLine` (prix fourni) et **pas** `AddLineAutoCalc` (remplace le prix par le tarif, souvent 0).
7. Description en **texte brut** (`\n` respectés) : le COM stocke le RTF tel quel (`RtfToText` convertit ce qui arrive en RTF).
8. Création de client : `TypeCli` **obligatoire** (« Professionnel ») ; `NomCtc` refusé par `SetFieldIdValStr` (contact non écrit) ; Sage recopie la facturation dans la livraison → écrire chaque rubrique de livraison, une espace `" "` pour vider ; un SIRET/N° TVA invalide bloque toute la fiche → contrôle avant envoi, puis second `Save` sans SIRET/TVA si Sage refuse encore.
9. Le CLSID de `WGC.exe` est dans la **vue 32 bits** du registre : lire avec `RegistryKey.OpenBaseKey(RegistryHive.ClassesRoot, RegistryView.Registry32)` (puis 64), sinon Sage est « introuvable » depuis un programme 64 bits (`SageWorker.SageExePath`, `Setup.SageExe`).
10. **Autorisation d'accès COM** à activer dans Sage : Utilitaires > Mode avancé > Divers > Autorisation accès COM (liée au compte Sage, ~6 mois ; chez Boris faite le 17/10/2026, à renouveler vers avril 2027). Sans elle, `OpenFile` échoue.

## Règles métier

- **Recherche du client** (`ResolveClient`) : texte normalisé (NFD sans accents, minuscules, non-alphanumérique → espace), CP = lettres+chiffres, pays → ISO 2 (vide = FR). Correspondance exacte sur nom + adresse 1 + CP + ville + pays. Bloc facturation de la commande comparé à la facturation **et** à la livraison de chaque fiche (nom contre `Nom` ou société du bloc), puis bloc livraison contre livraison. Prospects ignorés. Sans adresse exploitable → client par défaut des Réglages (sinon erreur).
- **Création** : code = suite du format le plus fréquent du dossier (« 281 » → « 282 », « CLI0041 » → « CLI0042 »), vérifié par `Exist`.
- **TVA** (`Norm.VatModeForCountry`) : FR, MC → « Local » ; 26 autres États UE → « CEE » ; reste du monde (Suisse, Royaume-Uni, DOM, COM…) → « Hors CEE ». Le pays seul décide. Pour un client existant, le `ModeTVA` de sa fiche fait foi. Si mode ≠ Local, toutes les lignes à `TauxTVA = 0`. Pays inconnu → Hors CEE + message.
- **SIRET** : 14 chiffres + Luhn. **N° TVA FR** : clé = (12 + 3 × (SIREN mod 97)) mod 97, SIREN cohérent avec le SIRET ; étrangers non contrôlés.
- **Référence de la facture** : laissée **vide** depuis 1.2.3 (demande de Boris : le n° de commande n'y est pas utile). Le champ `reference` du JSON n'est écrit que si `reference_facture=1` (case dans Réglages). L'extension n'envoie plus de référence depuis 3.25.2.
- **Anti-doublon** : clé `site|orderNo` dans `factures.csv` → HTTP 409 ; `force: true` pour recréer. Les simulations et les erreurs ne sont pas enregistrées.
- Activité de Boris : B2B uniquement (revendeurs), d'où le type de client « Professionnel » par défaut.

## API HTTP locale

| Route | Effet |
|---|---|
| `GET /ping` | `{ ok, app, version, sageOuvert, sageFenetre, dossier, clientParDefaut, derniereVersionPubliee, majAuto, copieInstallee }` |
| `POST /facture` | crée (ou simule) une facture |
| `GET /clients` | `[{ code, nom, ville, pays, modeTVA, tva, prospect }]` |
| `GET /diag?id=97003&fids=1,2,3` | outil de développement : nom/titre de la table, rubriques, valeurs de la 1ʳᵉ fiche |

Requête : `{ site, orderNo, reference, remarque, client, force, simulation, customer: {company, contact, address1, address2, postalCode, city, country, phone, email, siret, vat}, shipping: {…}, lines: [{ code, qty, puht, tva, description }] }`.
Réponse : `{ ok, numero, totalHT, totalTTC, nbLines, error, duplicate, simulation, clientCode, clientName, clientCreated, clientMatched, clientModeTVA, messages }` ; HTTP 200 / 409 doublon / 500 refus / 400 JSON invalide / 403 origine refusée / 413 > 2 Mo.

**Sécurité** (`Server.OriginAllowed`) : avec `Origin`, seuls `chrome-extension://`, `moz-extension://`, `safari-web-extension://` et `http://127.0.0.1:port` / `http://localhost:port` ; sans `Origin` (PowerShell, scripts), refus si `Sec-Fetch-Site` vaut `cross-site` ou `same-site`. CORS renvoyé uniquement à l'origine autorisée. Ne jamais remettre `Access-Control-Allow-Origin: *` : n'importe quel site pourrait alors créer des factures et lire les clients.

L'extension appelle depuis son **service worker** (un content script serait bloqué par *Private Network Access*), délai 150 s (SageBridge borne Sage à 120 s).

## Mises à jour automatiques

- `Updater.Latest` : `GET https://api.github.com/repos/{depot}/releases/latest` (sans compte, 60 requêtes/h/IP, TLS 1.2 forcé), **8 s après le démarrage puis toutes les heures**. Version = `tag_name` (`vX.Y.Z`). 404 = aucune version publiée.
- Si plus récente : téléchargement de `SageBridge-Setup.exe` et `SageBridge-Setup.exe.sha256` (URLs github.com / *.githubusercontent.com en HTTPS uniquement) dans `%TEMP%\SageBridge-maj\`, contrôle **SHA-256** et `ProductName == "SageBridge"`, attente qu'aucune facture ne soit en cours (`busy`) ni aucune fenêtre ouverte (10 min max), puis `SageBridge-Setup.exe /update /pid=<pid>` et fermeture.
- Automatique seulement si `maj_auto=1` **et** si le programme lancé est la copie installée (`InstallLocation` de la clé Uninstall) ; sinon notification + bouton « Mises à jour… » (vérification et installation à la demande). Pendant une mise à jour, `/facture` répond « réessayez dans une minute ».
- Testé de bout en bout le 20/10/2026 (1.2.0 → 1.2.1). C'est ce test qui a révélé la règle COM n° 2.

## Installateur (`Setup.cs`)

- `SageBridge-Setup.exe` : assistant (Bienvenue avec détection de Sage et de la version installée → Options → Progression → Terminé avec la liste des vérifications avant la première facture). Entrée valide toujours l'étape (le bouton principal garde le focus).
- `/silent` : sans questions, options précédentes. `/update /pid=N` : petite fenêtre de progression, attend la fin du processus N (15 s puis arrêt forcé), réinstalle, relance avec `/apres-maj=ancienne`.
- Installation **par utilisateur, sans droits admin** : `%LOCALAPPDATA%\Programs\SageBridge\` (`SageBridge.exe`, `SageBridge - Documentation.pdf`, `SageBridge-Uninstall.exe` = copie de l'installateur). Raccourcis Bureau / menu Démarrer via `WScript.Shell`.
- Registre (HKCU) : `Software\Microsoft\Windows\CurrentVersion\Uninstall\SageBridge` (DisplayName, DisplayVersion, Publisher, DisplayIcon, InstallLocation, UninstallString, QuietUninstallString, EstimatedSize, NoModify, NoRepair, `SB_Bureau`, `SB_Demarrer`) → visible dans Paramètres > Applications / Programmes et fonctionnalités ; `Software\Microsoft\Windows\CurrentVersion\Run\SageBridge` si « Démarrer avec Windows ».
- Désinstallation (`SageBridge-Uninstall.exe /uninstall`, lancé par Windows) : ferme SageBridge, supprime raccourcis, clé Run, clé Uninstall, fichiers ; option « supprimer mes données » (`%APPDATA%\SageBridge`) ; l'exe se supprime via un `cmd /c ping … & del` différé. Ne touche jamais aux données Sage.
- Journal : `%APPDATA%\SageBridge\installation.log`. Exe non signé → SmartScreen au premier lancement d'un fichier téléchargé par le navigateur (pas pour les mises à jour automatiques).

## Publier une nouvelle version

1. Modifier le code ; augmenter `AssemblyVersion` **et** `AssemblyFileVersion` dans `SageBridge.cs` (ex. `1.2.3.0`).
2. Ajouter en tête de `CHANGELOG.md` une section `## 1.2.3 — JJ/MM/AAAA` (devient la note de version GitHub).
3. Si la doc change, régénérer le PDF dans `docs\` (même nom).
4. Lancer `build.cmd`, vérifier `build.log` (`OK 1.2.3`).
5. **Installer d'abord la nouvelle version sur le PC de Boris** avec `Installateur\SageBridge-Setup.exe` si le changement touche à la fermeture ou à la mise à jour (le code de fermeture qui s'exécute lors d'une mise à jour est celui de l'**ancienne** version).
6. Lancer `publier-version.cmd`, vérifier `publier-version.log` (`PUBLICATION OK`). Le script :
   - lit le jeton du compte GitHub enregistré dans Git sur le PC (`git credential fill`, Git Credential Manager, compte **Gnaris**) — l'entrée standard de git est écrite **sans BOM** (`[Console]::InputEncoding` UTF-8 sans BOM), sinon « missing protocol field » ;
   - crée le dépôt public s'il n'existe pas ;
   - envoie le code **par l'API Contents** fichier par fichier, seulement les fichiers modifiés (empreinte « blob » Git) — un `git push` échouait en HTTP 408 sur la connexion de Boris ; exclus : `bin\`, `obj\`, `.git\`, `.vs\`, `*.log` ;
   - crée la release en **brouillon**, y joint l'installateur, l'empreinte et le PDF (4 essais chacun, `Expect: 100-continue` désactivé), puis la publie avec l'étiquette `vX.Y.Z` ; un brouillon laissé par une publication interrompue est repris ; une version publiée sans étiquette est réparée et les étiquettes `untagged-…` supprimées ;
   - refuse de republier une version existante.
7. Les SageBridge installés se mettent à jour au démarrage suivant ou dans l'heure.

## Tester sans risque

- `POST /facture` avec `"simulation": true` : résout le client et la TVA, calcule le total, **ne crée rien** (ni client ni facture). À privilégier.
- Bouton « Facture de test… » : crée une vraie facture provisoire d'une ligne (à supprimer ensuite dans Sage). Ne pas en créer sans prévenir Boris : son dossier Princesse est son dossier de production.
- Ne jamais supprimer de pièce ou de client dans Sage sans son accord. Attention : les numéros de facture provisoire peuvent être réutilisés ; vérifier une pièce avant de la croire « de test ».
- `GET /ping` pour l'état ; `journal.log` pour le détail de chaque opération et des messages de Sage.

## Environnement de travail de Claude chez Boris

- Session cloud reliée au PC de Boris (Windows 10, Sage 50 « Sage 50 (6.71) », dossier `C:\ProgramData\Ciel\Données Ciel\Dossiers\Princesse\Ciel Gestion Commerciale\Princesse.SgWgc`, client par défaut 003).
- Dossiers accessibles en fichiers : `Bureau\SageBridge`, `%APPDATA%\SageBridge`, l'extension `Regroupement facture`, `Téléchargements`, l'installation de Sage. Écrire les fichiers dans l'espace de travail puis les déposer sur le PC ; vérifier ensuite par lecture.
- **Compiler / lancer** : l'Explorateur n'est utilisable qu'en clic simple → sélectionner `build.cmd` (ou l'installateur, ou `publier-version.cmd`) puis ruban Accueil > **Ouvrir**. Lire ensuite `build.log` / `publier-version.log`. L'application Claude repasse souvent au premier plan : regrouper les clics (barre des tâches → miniature de la fenêtre → action) dans un même lot.
- **Appeler l'API** : le panneau navigateur de l'application Claude ouvert sur `http://127.0.0.1:8765/ping` permet des `fetch('/facture', …)` en même origine (acceptée par le contrôle d'origine).
- Le conteneur cloud ne joint pas `api.github.com` (proxy) ; pour vérifier une release, passer par ce même panneau navigateur (l'API GitHub autorise CORS).

## Limites connues / pistes

- Un seul dossier Sage à la fois ; Sage doit être ouvert sur ce PC.
- Contact du client non renseigné (il faudrait les méthodes de contacts du moteur client, non explorées).
- Clients créés en « Professionnel » ; régimes particuliers (franchise, suspension de taxe) à régler sur la fiche dans Sage.
- Les articles doivent exister dans Sage (ligne refusée sinon, rien n'est créé).
- Installateur non signé (un certificat de signature de code supprimerait l'avertissement SmartScreen).
- Identifiants de tables autres que Clients (97003) calculés mais non vérifiés : contrôler avec `/diag` avant usage.

## Historique

1.0 (création de factures par COM) → 1.1 (client automatique, TVA par pays, contrôle SIRET/TVA) → 1.2.0 (installateur, mises à jour GitHub, logo, sécurité de l'API, partageable, détection de Sage en registre 32 bits) → 1.2.1 (publication fiabilisée, `/ping` enrichi) → 1.2.2 (fermeture sans appel COM — corrige le gel de Sage —, focus de l'assistant) → 1.2.3 (Référence de la facture laissée vide, option dans Réglages). Détail dans `CHANGELOG.md` et chapitre 8 du PDF.

// ════════════════════════════════════════════════════════════════════════════
//  SageBridge 1.2 — création de factures Sage 50 Gestion commerciale par COM
// ════════════════════════════════════════════════════════════════════════════
//  Reçoit des factures en JSON sur http://127.0.0.1:8765/facture (extension
//  Chrome « Regroupement facture » ou tout autre outil), retrouve ou crée le
//  client, puis crée la facture PROVISOIRE (non validée) dans le dossier Sage
//  par l'interface COM « CielGestionCommerciale.Application ».
//
//  Documentation complète : « SageBridge - Documentation COM Sage 50.pdf ».
//  Compilation : build.cmd (csc.exe du .NET Framework 4 fourni avec Windows,
//  syntaxe C# 5 uniquement : pas d'interpolation $"", pas de ?. ni nameof).
// ════════════════════════════════════════════════════════════════════════════

using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Web.Script.Serialization;
using System.Windows.Forms;
using Microsoft.Win32;

[assembly: AssemblyTitle("SageBridge")]
[assembly: AssemblyDescription("Création de factures Sage 50 Gestion commerciale par l'interface COM")]
[assembly: AssemblyCompany("Beli & Jolie")]
[assembly: AssemblyProduct("SageBridge")]
[assembly: AssemblyCopyright("© 2026 Beli & Jolie")]
[assembly: AssemblyVersion("1.4.0.0")]
[assembly: AssemblyFileVersion("1.4.0.0")]

namespace SageBridge
{
    // ───────────────────────── Programme ─────────────────────────

    static class Program
    {
        public static readonly string Version = ShortVersion(Assembly.GetExecutingAssembly().GetName().Version);
        public static string PreviousVersion = "";   // renseigné après une mise à jour (argument /apres-maj=…)
        public const string AppName = "SageBridge";
        public const string ShowEventName = "SageBridge-afficher-fenetre";
        public const string DocFileName = "SageBridge - Documentation.pdf";

        public static string ShortVersion(Version v)
        {
            return v.Major + "." + v.Minor + "." + Math.Max(0, v.Build);
        }

        [STAThread]
        static void Main(string[] args)
        {
            foreach (string a in args)
                if (a.StartsWith("/apres-maj=", StringComparison.OrdinalIgnoreCase)) PreviousVersion = a.Substring(11);
            bool created;
            using (var mutex = new Mutex(true, "SageBridge-instance-unique", out created))
            {
                if (!created)
                {
                    // Déjà lancé : on demande à l'instance existante de s'afficher.
                    try { using (var ev = EventWaitHandle.OpenExisting(ShowEventName)) ev.Set(); }
                    catch { MessageBox.Show("SageBridge est déjà lancé (icône en bas à droite, près de l'horloge).", AppName, MessageBoxButtons.OK, MessageBoxIcon.Information); }
                    return;
                }
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);
                Application.ThreadException += delegate(object s, ThreadExceptionEventArgs e) { Fatal(e.Exception); };
                AppDomain.CurrentDomain.UnhandledException += delegate(object s, UnhandledExceptionEventArgs e) { Fatal(e.ExceptionObject as Exception); };
                Application.Run(new MainForm());
            }
        }

        static void Fatal(Exception e)
        {
            Settings.AppendLog("ERREUR INATTENDUE : " + (e == null ? "?" : e.ToString()));
            MessageBox.Show("Erreur inattendue dans SageBridge :\n\n" + (e == null ? "?" : e.Message) +
                "\n\nLe détail est enregistré dans le journal (bouton « Journal »).", AppName, MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        public static Icon AppIcon
        {
            get { try { return Icon.ExtractAssociatedIcon(Application.ExecutablePath); } catch { return SystemIcons.Application; } }
        }

        public static string DocPath { get { return Path.Combine(Path.GetDirectoryName(Application.ExecutablePath), DocFileName); } }
    }

    // ───────────────────────── Réglages et fichiers ─────────────────────────

    class Settings
    {
        public string DossierPath = "";
        public int Port = 8765;
        public string DefaultClient = "";   // client utilisé quand la commande n'a pas d'adresse exploitable
        public string TestArticle = "";     // dernier code article utilisé pour la facture de test
        public bool AutoUpdate = true;      // installe seul les nouvelles versions publiées sur GitHub
        public bool WriteReference = false; // écrire le champ « reference » reçu dans la Référence de la facture (désactivé par défaut)
        public string Repo = Updater.DefaultRepo;
        public string RecapDir = "";        // dossier des PDF de factures, choisi par l'utilisateur (aucun par défaut)
        public string PdfModel = "Format A4 Complet"; // modèle Sage utilisé pour l'aperçu exporté en PDF

        public string RecapRoot { get { return RecapDir; } }

        static readonly object logSync = new object();

        public static string Dir
        {
            get
            {
                string d = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), Program.AppName);
                Directory.CreateDirectory(d);
                return d;
            }
        }
        public static string SettingsFile { get { return Path.Combine(Dir, "settings.ini"); } }
        public static string HistoryFile { get { return Path.Combine(Dir, "factures.csv"); } }
        public static string LogFile { get { return Path.Combine(Dir, "journal.log"); } }

        public static void AppendLog(string message)
        {
            lock (logSync)
            {
                try
                {
                    var fi = new FileInfo(LogFile);
                    if (fi.Exists && fi.Length > 5 * 1024 * 1024)
                    {
                        string old = Path.Combine(Dir, "journal.ancien.log");
                        if (File.Exists(old)) File.Delete(old);
                        File.Move(LogFile, old);
                    }
                    File.AppendAllText(LogFile, DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss") + "  " + message + "\r\n", Encoding.UTF8);
                }
                catch { }
            }
        }

        public static Settings Load()
        {
            var s = new Settings();
            if (!File.Exists(SettingsFile)) return s;
            foreach (string raw in File.ReadAllLines(SettingsFile, Encoding.UTF8))
            {
                string line = raw.Trim();
                int eq = line.IndexOf('=');
                if (eq <= 0) continue;
                string k = line.Substring(0, eq).Trim(), v = line.Substring(eq + 1).Trim();
                if (k == "dossier") s.DossierPath = v;
                else if (k == "port") { int p; if (int.TryParse(v, out p) && p > 0 && p < 65536) s.Port = p; }
                else if (k == "client") s.DefaultClient = v;
                else if (k == "article_test") s.TestArticle = v;
                else if (k == "maj_auto") s.AutoUpdate = v != "0";
                else if (k == "reference_facture") s.WriteReference = v == "1";
                else if (k == "recap") s.RecapDir = v;
                else if (k == "modele_pdf" && v.Length > 0) s.PdfModel = v;
                else if (k == "depot" && Regex.IsMatch(v, @"^[\w.-]+/[\w.-]+$")) s.Repo = v;
            }
            return s;
        }

        public void Save()
        {
            File.WriteAllText(SettingsFile,
                "dossier=" + DossierPath + "\r\n" +
                "port=" + Port + "\r\n" +
                "client=" + DefaultClient + "\r\n" +
                "article_test=" + TestArticle + "\r\n" +
                "maj_auto=" + (AutoUpdate ? "1" : "0") + "\r\n" +
                "reference_facture=" + (WriteReference ? "1" : "0") + "\r\n" +
                "recap=" + RecapDir + "\r\n" +
                "modele_pdf=" + PdfModel + "\r\n" +
                "depot=" + Repo + "\r\n", Encoding.UTF8);
        }

        public string DossierName
        {
            get { return DossierLabel(DossierPath); }
        }

        // Dossier ouvert par le serveur de Sage (mode Données Mobiles ou réseau) : chemin tel qu'affiché dans
        // Dossier > Options > Informations > Chemin d'accès, ex. « !Srv!<Smart>/S:/Données Ciel/Dossiers/X/Ciel Gestion Commerciale/X.SgWgc ».
        public static bool IsServerPath(string p)
        {
            return !string.IsNullOrEmpty(p) && p.TrimStart().StartsWith("!Srv!", StringComparison.OrdinalIgnoreCase);
        }

        // Un chemin serveur ne peut pas être vérifié sur le disque : c'est Sage qui tranche à l'ouverture.
        public static bool DossierExists(string p)
        {
            if (string.IsNullOrEmpty(p)) return false;
            if (IsServerPath(p)) return true;
            try { return File.Exists(p); } catch { return false; }
        }

        // Nom du dossier sans extension, sans passer par Path (qui refuse les « < » « > » des chemins serveur).
        public static string DossierLabel(string p)
        {
            if (string.IsNullOrEmpty(p)) return "";
            string s = p.Trim();
            int cut = Math.Max(s.LastIndexOf('/'), s.LastIndexOf('\\'));
            if (cut >= 0) s = s.Substring(cut + 1);
            int dot = s.LastIndexOf('.');
            if (dot > 0) s = s.Substring(0, dot);
            return s;
        }

        // Ancienne copie laissée par Sage lors d'une conversion (ex. passage en Données Mobiles) : « X_OLD.SgWgc ».
        // Sage refuse de l'ouvrir (« Il est impossible d'ouvrir une base locale »).
        public static bool IsOldCopy(string p)
        {
            return DossierLabel(p).EndsWith("_OLD", StringComparison.OrdinalIgnoreCase);
        }

        // Dossiers Sage 50 Gestion commerciale (*.SgWgc), hors points de restauration, modèles et anciennes copies « _OLD ».
        public static List<string> FindDossiers()
        {
            var roots = new List<string>();
            string pd = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData);
            roots.Add(Path.Combine(pd, @"Ciel\Données Ciel\Dossiers"));
            roots.Add(Path.Combine(pd, "Sage"));
            roots.Add(@"C:\Données Ciel");
            roots.Add(Environment.GetFolderPath(Environment.SpecialFolder.CommonDocuments));
            roots.Add(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));
            var found = new List<string>();
            foreach (string root in roots)
            {
                if (!Directory.Exists(root)) continue;
                try
                {
                    foreach (string f in Directory.GetFiles(root, "*.SgWgc", SearchOption.AllDirectories))
                    {
                        string low = f.ToLowerInvariant();
                        if (low.Contains("restorepoints") || low.Contains("\\statics\\") || low.Contains("compléments") || low.Contains("complements")) continue;
                        if (IsOldCopy(f)) continue;
                        if (!found.Contains(f)) found.Add(f);
                    }
                }
                catch { }
            }
            return found;
        }
    }

    // ───────────────────────── Modèle d'une facture ─────────────────────────

    class InvoiceLine
    {
        public string Code = "";
        public double Qty;
        public double Puht;
        public bool HasPuht;
        public double Tva = 20;
        public string Description = "";
    }

    class CustomerInfo
    {
        public string Company = "", Contact = "", Address1 = "", Address2 = "", PostalCode = "", City = "", Country = "";
        public string Phone = "", Email = "", Siret = "", Vat = "";

        public bool IsEmpty
        {
            get { return (Company + Contact).Trim().Length == 0 || (Address1 + City + PostalCode).Trim().Length == 0; }
        }
        public string DisplayName { get { return !string.IsNullOrEmpty(Company) ? Company.Trim() : Contact.Trim(); } }
    }

    class InvoiceRequest
    {
        public string Site = "", OrderNo = "", Client = "", Reference = "", Remarque = "";
        public bool Force, Simulation;
        public CustomerInfo Customer;   // adresse de facturation
        public CustomerInfo Shipping;   // adresse de livraison (facultative)
        public List<InvoiceLine> Lines = new List<InvoiceLine>();
    }

    class InvoiceResult
    {
        public bool Ok, Duplicate, Simulated, ClientCreated, ClientMatched;
        public string Numero = "", Error = "", ClientCode = "", ClientName = "", ClientModeTVA = "";
        public int NbLines;
        public double TotalHT, TotalTTC;
        public List<string> Messages = new List<string>();
    }

    class SolderResult
    {
        public bool Ok, Simulated;
        public string Numero = "", Error = "", Mode = "";
        public double Ecart, TotalFac;
        public int Lignes;
        public List<string> Messages = new List<string>();
    }

    // Pièce client lue dans la table PieceHCli (lecture seule).
    class PieceRow
    {
        public string Num = "", Type = "", TypePiece = "", Date = "", Client = "", Societe = "", CliName = "", Statut = "", ModeRglt = "";
        public double TotalHT, TotalTTC, SoldeDu, MtRegle;
        public bool Valide, Soldee;

        // « facture », « avoir », « acompte » ou « autre », d'après le libellé du type de pièce.
        public string Kind
        {
            get
            {
                string t = Norm.Text(TypePiece);
                if (t.Contains("avoir")) return "avoir";
                if (t.Contains("acompte")) return "acompte";
                if (t.Contains("facture")) return "facture";
                return "autre";
            }
        }

        public string NomAffiche { get { return Societe.Length > 0 ? Societe : CliName; } }
    }

    // Pièce à enregistrer en PDF (facture ou avoir validé), avec son UD dans la table PieceHCli.
    class PdfPiece
    {
        public string Num = "", TypePiece = "", Kind = "", Nom = "";
        public DateTime Date;
        public double TotalTTC;
        public bool Valide;
        public int UD;
        // Résultat de l'export
        public string Existing = "", Target = "", Status = "", Error = "";
    }

    // ───────────────────────── Normalisation (sans accents, sans casse) ─────────────────────────

    static class Norm
    {
        // Minuscules, sans accents ; tout ce qui n'est ni lettre ni chiffre devient un espace unique.
        public static string Text(string s)
        {
            if (string.IsNullOrEmpty(s)) return "";
            string d = s.Normalize(NormalizationForm.FormD);
            var sb = new StringBuilder(d.Length);
            foreach (char c in d)
            {
                if (CharUnicodeInfo.GetUnicodeCategory(c) == UnicodeCategory.NonSpacingMark) continue;
                sb.Append(char.IsLetterOrDigit(c) ? char.ToLowerInvariant(c) : ' ');
            }
            return Regex.Replace(sb.ToString(), " +", " ").Trim();
        }

        // Lettres et chiffres uniquement, en majuscules (codes postaux, SIRET, N° TVA).
        public static string Digits(string s)
        {
            if (string.IsNullOrEmpty(s)) return "";
            var sb = new StringBuilder();
            foreach (char c in s) if (char.IsLetterOrDigit(c)) sb.Append(char.ToUpperInvariant(c));
            return sb.ToString();
        }

        // ISO 2, ISO 3 (stocké par Sage dans FacCtyCode), nom affiché, puis variantes acceptées.
        static readonly string[][] CountryTable = {
            new[] { "FR", "FRA", "France", "Métropole", "France métropolitaine" },
            new[] { "MC", "MCO", "Monaco" },
            // Union européenne (hors France)
            new[] { "DE", "DEU", "Allemagne", "Germany", "Deutschland" },
            new[] { "AT", "AUT", "Autriche", "Austria", "Österreich" },
            new[] { "BE", "BEL", "Belgique", "Belgium", "België", "Belgien" },
            new[] { "BG", "BGR", "Bulgarie", "Bulgaria" },
            new[] { "CY", "CYP", "Chypre", "Cyprus" },
            new[] { "HR", "HRV", "Croatie", "Croatia", "Hrvatska" },
            new[] { "DK", "DNK", "Danemark", "Denmark", "Danmark" },
            new[] { "ES", "ESP", "Espagne", "Spain", "España" },
            new[] { "EE", "EST", "Estonie", "Estonia", "Eesti" },
            new[] { "FI", "FIN", "Finlande", "Finland", "Suomi" },
            new[] { "GR", "GRC", "Grèce", "Greece", "Hellas", "EL" },
            new[] { "HU", "HUN", "Hongrie", "Hungary", "Magyarország" },
            new[] { "IE", "IRL", "Irlande", "Ireland", "Éire" },
            new[] { "IT", "ITA", "Italie", "Italy", "Italia" },
            new[] { "LV", "LVA", "Lettonie", "Latvia", "Latvija" },
            new[] { "LT", "LTU", "Lituanie", "Lithuania", "Lietuva" },
            new[] { "LU", "LUX", "Luxembourg", "Luxemburg" },
            new[] { "MT", "MLT", "Malte", "Malta" },
            new[] { "NL", "NLD", "Pays-Bas", "Netherlands", "Nederland", "Hollande", "Holland" },
            new[] { "PL", "POL", "Pologne", "Poland", "Polska" },
            new[] { "PT", "PRT", "Portugal" },
            new[] { "CZ", "CZE", "République tchèque", "Tchéquie", "Czech Republic", "Czechia" },
            new[] { "RO", "ROU", "Roumanie", "Romania" },
            new[] { "SK", "SVK", "Slovaquie", "Slovakia", "Slovensko" },
            new[] { "SI", "SVN", "Slovénie", "Slovenia", "Slovenija" },
            new[] { "SE", "SWE", "Suède", "Sweden", "Sverige" },
            // Hors Union européenne
            new[] { "CH", "CHE", "Suisse", "Switzerland", "Schweiz", "Svizzera" },
            new[] { "GB", "GBR", "Royaume-Uni", "United Kingdom", "UK", "Grande-Bretagne", "Great Britain", "Angleterre", "England", "Écosse", "Scotland", "Pays de Galles", "Wales", "Irlande du Nord", "Northern Ireland" },
            new[] { "NO", "NOR", "Norvège", "Norway", "Norge" },
            new[] { "IS", "ISL", "Islande", "Iceland" },
            new[] { "LI", "LIE", "Liechtenstein" },
            new[] { "AD", "AND", "Andorre", "Andorra" },
            new[] { "SM", "SMR", "Saint-Marin", "San Marino" },
            new[] { "US", "USA", "États-Unis", "United States", "United States of America", "Etats-Unis d'Amérique" },
            new[] { "CA", "CAN", "Canada" },
            new[] { "MA", "MAR", "Maroc", "Morocco" },
            new[] { "DZ", "DZA", "Algérie", "Algeria" },
            new[] { "TN", "TUN", "Tunisie", "Tunisia" },
            new[] { "SN", "SEN", "Sénégal", "Senegal" },
            new[] { "CI", "CIV", "Côte d'Ivoire", "Ivory Coast" },
            new[] { "TR", "TUR", "Turquie", "Turkey", "Türkiye" },
            new[] { "IL", "ISR", "Israël", "Israel" },
            new[] { "AE", "ARE", "Émirats arabes unis", "United Arab Emirates", "UAE", "Dubaï", "Dubai" },
            new[] { "UA", "UKR", "Ukraine" },
            new[] { "RS", "SRB", "Serbie", "Serbia" },
            new[] { "AL", "ALB", "Albanie", "Albania" },
            new[] { "BA", "BIH", "Bosnie-Herzégovine", "Bosnia and Herzegovina" },
            new[] { "ME", "MNE", "Monténégro", "Montenegro" },
            new[] { "MK", "MKD", "Macédoine du Nord", "North Macedonia" },
            new[] { "CN", "CHN", "Chine", "China" },
            new[] { "JP", "JPN", "Japon", "Japan" },
            new[] { "KR", "KOR", "Corée du Sud", "South Korea", "Korea" },
            new[] { "AU", "AUS", "Australie", "Australia" },
            new[] { "BR", "BRA", "Brésil", "Brazil", "Brasil" },
            new[] { "MX", "MEX", "Mexique", "Mexico" },
            // Outre-mer : hors territoire de TVA de l'UE, traité comme export par Sage
            new[] { "RE", "REU", "La Réunion", "Réunion" },
            new[] { "GP", "GLP", "Guadeloupe" },
            new[] { "MQ", "MTQ", "Martinique" },
            new[] { "GF", "GUF", "Guyane", "Guyane française", "French Guiana" },
            new[] { "YT", "MYT", "Mayotte" },
            new[] { "PM", "SPM", "Saint-Pierre-et-Miquelon" },
            new[] { "BL", "BLM", "Saint-Barthélemy" },
            new[] { "MF", "MAF", "Saint-Martin" },
            new[] { "NC", "NCL", "Nouvelle-Calédonie", "New Caledonia" },
            new[] { "PF", "PYF", "Polynésie française", "French Polynesia", "Tahiti" },
            new[] { "WF", "WLF", "Wallis-et-Futuna" }
        };

        static readonly HashSet<string> EuMembers = new HashSet<string> {
            "AT", "BE", "BG", "HR", "CY", "CZ", "DK", "EE", "FI", "DE", "GR", "HU", "IE", "IT",
            "LV", "LT", "LU", "MT", "NL", "PL", "PT", "RO", "SK", "SI", "ES", "SE"
        };

        static readonly Dictionary<string, string[]> byKey = BuildIndex();

        static Dictionary<string, string[]> BuildIndex()
        {
            // Affectation par indexeur (et non initialiseur de collection) : une variante en double
            // ne peut plus faire planter le programme (erreur rencontrée en 1.1, cf. documentation).
            var d = new Dictionary<string, string[]>();
            foreach (string[] row in CountryTable)
                foreach (string name in row)
                {
                    string k = Text(name);
                    if (k.Length > 0 && !d.ContainsKey(k)) d[k] = row;
                }
            return d;
        }

        // Clé pays comparable : code ISO 2 lettres si reconnu. Vide = FR (dossier français).
        public static string CountryKey(string s)
        {
            string t = Text(s);
            if (t.Length == 0) return "FR";
            string[] row;
            if (byKey.TryGetValue(t, out row)) return row[0];
            if (t.Length == 2 && Regex.IsMatch(t, "^[a-z]{2}$")) return t.ToUpperInvariant();
            return t; // pays non reconnu : nom normalisé
        }

        // Résolution DOM/TOM par code postal français à 5 chiffres (97xxx / 98xxx).
        // Sert quand une boutique envoie un label de groupe (ex. « France DOM/TOM »,
        // « France Outre-Mer ») que ni Sage ni la table des pays ne reconnaissent :
        // sans ce repli, FacCtyCode reste vide et Sage refuse la fiche avec « La
        // rubrique Code Pays doit être remplie ». Incident 02/11/2026, commande
        // chez Boris, client « INTERIEUR D AUJOURD HUI » (pays « France DOM/TOM »).
        public static string PostalCodeToDomTomKey(string postalCode)
        {
            string d = Digits(postalCode);
            if (d.Length < 3) return "";
            // Cas spécifiques (chevauchent 971xx) :
            if (d.StartsWith("97133")) return "BL"; // Saint-Barthélemy
            if (d.StartsWith("97150")) return "MF"; // Saint-Martin
            string p3 = d.Substring(0, 3);
            if (p3 == "971") return "GP"; // Guadeloupe
            if (p3 == "972") return "MQ"; // Martinique
            if (p3 == "973") return "GF"; // Guyane
            if (p3 == "974") return "RE"; // La Réunion
            if (p3 == "975") return "PM"; // Saint-Pierre-et-Miquelon
            if (p3 == "976") return "YT"; // Mayotte
            if (p3 == "986") return "WF"; // Wallis-et-Futuna
            if (p3 == "987") return "PF"; // Polynésie française
            if (p3 == "988") return "NC"; // Nouvelle-Calédonie
            return "";
        }

        // Clé pays avec repli DOM/TOM par CP : si le pays n'est pas reconnu et que le
        // code postal pointe sur un territoire d'outre-mer français, on retourne le
        // code ISO 2 du territoire. Un pays reconnu n'est jamais écrasé (une facture
        // « France » + CP 97400 reste FR : si Boris a choisi « France », c'est voulu).
        public static string CountryKeyWithPostal(string country, string postalCode)
        {
            string key = CountryKey(country);
            if (IsKnownCountry(key)) return key;
            string fromCp = PostalCodeToDomTomKey(postalCode);
            return fromCp.Length > 0 ? fromCp : key;
        }

        public static bool IsKnownCountry(string key)
        {
            foreach (string[] row in CountryTable) if (row[0] == key) return true;
            return EuMembers.Contains(key);
        }

        public static string CountryIso3(string key)
        {
            foreach (string[] row in CountryTable) if (row[0] == key) return row[1];
            return "";
        }

        public static string CountryName(string key, string fallback)
        {
            foreach (string[] row in CountryTable) if (row[0] == key) return row[2];
            return fallback;
        }

        // Nom du pays sans accents, pour matcher la table pays interne de Sage.
        // Sage 50 stocke généralement les noms sans accents (ex. « Etats-Unis »,
        // pas « États-Unis »). Si on écrit un nom avec accents dans FacCtyName,
        // Sage ne le reconnaît pas et efface silencieusement FacCtyCode associé
        // → erreur « Code Pays doit être rempli » à Save(). Incident 30/10/2026
        // commande KJ25KJNK, client « Accent Piece LLC » (US).
        public static string CountryNameForSage(string key, string fallback)
        {
            string name = CountryName(key, fallback);
            if (string.IsNullOrEmpty(name)) return name;
            string d = name.Normalize(NormalizationForm.FormD);
            var sb = new StringBuilder(d.Length);
            foreach (char c in d)
                if (CharUnicodeInfo.GetUnicodeCategory(c) != UnicodeCategory.NonSpacingMark) sb.Append(c);
            return sb.ToString();
        }

        // Mode TVA Sage selon le pays de facturation : France et Monaco → Local ;
        // autres pays de l'UE → CEE ; reste du monde (dont DOM, Suisse, Royaume-Uni) → Hors CEE.
        public static string VatModeForCountry(string key)
        {
            if (key == "FR" || key == "MC") return "Local";
            if (EuMembers.Contains(key)) return "CEE";
            return "Hors CEE";
        }

        // Le COM de Sage stocke le RTF tel quel : on convertit en texte (retours à la ligne conservés).
        public static string RtfToText(string s)
        {
            if (string.IsNullOrEmpty(s) || !s.StartsWith("{\\rtf")) return s;
            var sb = new StringBuilder();
            int i = s.IndexOf(' ');
            i = i > 0 ? i + 1 : 0;
            while (i < s.Length)
            {
                char c = s[i];
                if (c == '{' || c == '}') { i++; continue; }
                if (c == '\\')
                {
                    if (i + 3 < s.Length && s[i + 1] == '\'')
                    {
                        int code;
                        if (int.TryParse(s.Substring(i + 2, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out code))
                            sb.Append(Encoding.GetEncoding(1252).GetString(new[] { (byte)code }));
                        i += 4; continue;
                    }
                    if (i + 1 < s.Length && (s[i + 1] == '\\' || s[i + 1] == '{' || s[i + 1] == '}')) { sb.Append(s[i + 1]); i += 2; continue; }
                    int j = i + 1;
                    while (j < s.Length && char.IsLetterOrDigit(s[j])) j++;
                    string word = s.Substring(i + 1, j - i - 1);
                    if (word == "line" || word == "par") sb.Append("\r\n");
                    i = j;
                    if (i < s.Length && s[i] == ' ') i++;
                    continue;
                }
                sb.Append(c); i++;
            }
            return sb.ToString().Trim();
        }
    }

    // ───────────────────────── Accès Sage 50 (COM, thread STA dédié) ─────────────────────────

    class SageWorker
    {
        // Moteurs (WGC.tlb, enum eEngineIds) — valeurs utilisables telles quelles.
        public const int ENG_PIECE_CLI = 97001, ENG_LIGNE_CLI = 97002, ENG_CLIENT = 97004;
        // Types de pièce (enum eWgcType) : 1 devis, 2 commande, 3 BL, 4 facture, 5 avoir, 6 acompte.
        public const int TYPE_FACTURE = 4;
        // Tables (enum eTableIds) : OpenTable attend la valeur sur 24 bits (valeur & 0xFFFFFF).
        // kWgcTableIdCli = 0x8C017AEB → 0x017AEB = 97003. La valeur brute ouvre une autre table.
        public const int TBL_CLIENTS = 0x017AEB;
        // Table articles : TableName vérifié live le 29/10/2026 sur Sage 50 v6.71.
        // L'id est caché derrière la bonne valeur ; la constante reste le secret
        // de bas niveau, ResolveArticleTableId() valide à l'ouverture.
        public const int TBL_ARTICLES = 97002;

        // Rubriques de la table clients : résolues par leur nom à l'ouverture (ITable.Fields),
        // avec ces valeurs de secours (enum eWgcTableCliFieldIds & 0xFFFFFF, vérifiées sur Sage 50 v6.71).
        static readonly Dictionary<string, int> CliFieldDefaults = new Dictionary<string, int> {
            { "Code", 1 }, { "Nom", 2 }, { "FacSociete", 3 }, { "FacAdresse1", 5 }, { "FacAdresse2", 6 },
            { "FacCP", 8 }, { "FacVille", 9 }, { "FacCtyCode", 10 }, { "FacCtyName", 11 },
            { "LivSociete", 12 }, { "LivAdresse1", 14 }, { "LivAdresse2", 15 }, { "LivCP", 17 }, { "LivVille", 18 },
            { "LivCtyCode", 19 }, { "LivCtyName", 20 }, { "FacTel", 25 }, { "ModeTVA", 32 }, { "NII", 33 },
            { "Email", 54 }, { "SIRET", 56 }, { "LivTel", 71 }, { "IsProspect", 129 }, { "TypeCli", 143 }
        };

        public Action<string> Log = delegate { };
        readonly Settings cfg;
        readonly BlockingCollection<Action> queue = new BlockingCollection<Action>();
        dynamic app, doc;
        Dictionary<string, int> cliFields;

        // Cache clients persistant (voir ReadClients). Clé = nom du dossier + NbN de la table.
        // En base locale un balayage de la table clients est quasi gratuit, en Données Mobiles
        // il coûte ~40 s pour 11000 fiches (chaque Next+GetVal est un aller-retour réseau).
        // Avec le cache, chaque facture économise ce balayage, sauf au premier démarrage sur
        // un dossier ou après qu'une fiche a été ajoutée/supprimée depuis Sage.
        List<ClientRow> cachedClients;
        string cachedClientsDossier;
        int cachedClientsNbN;
        const int CLIENTS_CACHE_VERSION = 1;

        public SageWorker(Settings settings)
        {
            cfg = settings;
            var th = new Thread(Loop);
            th.SetApartmentState(ApartmentState.STA); // les objets COM de Sage doivent rester sur un seul thread STA
            th.IsBackground = true;
            th.Name = "SageCOM";
            th.Start();
        }

        void Loop()
        {
            foreach (Action a in queue.GetConsumingEnumerable())
            {
                try { a(); } catch (Exception e) { Log("Erreur interne : " + e.Message); }
            }
        }

        // ── Processus Sage (WGC.exe) ──

        // Chemin de WGC.exe (serveur COM 32 bits : CLSID dans la vue 32 bits du registre).
        public static string SageExePath()
        {
            foreach (RegistryView view in new[] { RegistryView.Registry32, RegistryView.Registry64 })
            {
                try
                {
                    using (RegistryKey root = RegistryKey.OpenBaseKey(RegistryHive.ClassesRoot, view))
                    using (RegistryKey k = root.OpenSubKey(@"CielGestionCommerciale.Application\CLSID"))
                    {
                        string clsid = k == null ? null : (string)k.GetValue("");
                        if (clsid == null) continue;
                        using (RegistryKey ls = root.OpenSubKey(@"CLSID\" + clsid + @"\LocalServer32"))
                        {
                            string exe = ls == null ? null : ((string)ls.GetValue("") ?? "").Trim();
                            if (string.IsNullOrEmpty(exe)) continue;
                            if (exe.StartsWith("\"")) exe = exe.Substring(1, exe.IndexOf('"', 1) - 1);
                            else { int sp = exe.IndexOf(".exe ", StringComparison.OrdinalIgnoreCase); if (sp > 0) exe = exe.Substring(0, sp + 4); }
                            if (File.Exists(exe)) return exe;
                        }
                    }
                }
                catch { }
            }
            return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), @"Ciel\Gestion Commerciale\WGC.exe");
        }

        public static bool SageWindowOpen()
        {
            string t; return SageWindow(out t);
        }

        public static bool SageWindow(out string title)
        {
            title = "";
            try
            {
                foreach (Process p in Process.GetProcessesByName("WGC"))
                {
                    try { if (p.MainWindowHandle != IntPtr.Zero) { title = p.MainWindowTitle; return true; } }
                    catch { }
                    finally { p.Dispose(); }
                }
            }
            catch { }
            return false;
        }

        // Instances WGC.exe sans fenêtre, lancées par une activation COM alors que Sage était fermé.
        // Elles ne répondent jamais et empêchent Sage de s'ouvrir. On ne les ferme que si aucun Sage
        // avec fenêtre n'est ouvert et qu'elles ont plus d'une minute (pour ne pas tuer un Sage qui démarre).
        public static int KillOrphans(Action<string> log)
        {
            int killed = 0;
            if (SageWindowOpen()) return 0;
            try
            {
                foreach (Process p in Process.GetProcessesByName("WGC"))
                {
                    try
                    {
                        if (p.MainWindowHandle == IntPtr.Zero && (DateTime.Now - p.StartTime).TotalSeconds > 60)
                        {
                            log("Fermeture d'une instance invisible de Sage (COM orphelin, PID " + p.Id + ").");
                            p.Kill(); p.WaitForExit(5000); killed++;
                        }
                    }
                    catch (Exception e) { log("Instance Sage PID " + p.Id + " : " + e.Message); }
                    finally { p.Dispose(); }
                }
            }
            catch { }
            return killed;
        }

        // ── Exécution sur le thread COM ──

        T Run<T>(Func<T> f) { return Run<T>(f, 120000); }

        T Run<T>(Func<T> f, int timeoutMs)
        {
            if (!SageWindowOpen())
                throw new Exception("Sage 50 n'est pas ouvert sur ce PC : ouvrez Sage" +
                    (cfg.DossierName.Length > 0 ? " (dossier " + cfg.DossierName + ")" : "") + " puis réessayez.");
            T result = default(T);
            Exception error = null;
            var started = new ManualResetEvent(false);
            var done = new ManualResetEvent(false);
            queue.Add(delegate
            {
                started.Set();
                try { result = f(); } catch (Exception e) { error = e; }
                done.Set();
            });
            // Les 120 s sont comptées à partir du DÉBUT de l'action, pas de sa mise en file.
            // Incident ISSYMA 30/10/2026 : avec 3 factures en rafale, la 3ᵉ a vu son HTTP
            // expirer alors qu'elle attendait encore son tour ; l'action a continué, la
            // facture a été créée dans Sage sans journalisation ni historique (donc risque
            // de doublon à la retentative). Le garde-fou de 15 min sur l'attente en file
            // permet aux rafales longues de passer tout en évitant un blocage éternel.
            if (!started.WaitOne(900000))
                throw new Exception("Sage très sollicité : la demande est restée en file plus de 15 minutes. Vérifiez l'état de Sage et réessayez.");
            if (!done.WaitOne(timeoutMs))
            {
                queue.Add(delegate { Reset(); });
                throw new Exception("Sage ne répond pas (plus de " + (timeoutMs / 60000) + " minutes). Vérifiez qu'aucune fenêtre de Sage n'attend une réponse, puis réessayez.");
            }
            if (error != null) throw error;
            return result;
        }

        static string ComMessage(Exception e)
        {
            Exception x = e;
            while (x.InnerException != null) x = x.InnerException;
            return x.Message.Trim();
        }

        void Reset()
        {
            try { if (doc != null) Marshal.FinalReleaseComObject(doc); } catch { }
            try { if (app != null) Marshal.FinalReleaseComObject(app); } catch { }
            doc = null; app = null; cliFields = null;
        }

        void EnsureOpen()
        {
            if (doc != null) return;
            if (!Settings.DossierExists(cfg.DossierPath))
                throw new Exception("Dossier Sage introuvable : « " + cfg.DossierPath + " ». Choisissez-le dans Réglages.");
            Type t = Type.GetTypeFromProgID("CielGestionCommerciale.Application");
            if (t == null) throw new Exception("Sage 50 Gestion commerciale n'est pas installé sur ce PC (objet COM CielGestionCommerciale.Application absent).");
            // Toujours une NOUVELLE instance : l'objet obtenu par GetActiveObject ne répond pas (E_NOTIMPL).
            app = Activator.CreateInstance(t);
            string version = "?";
            try { version = (string)app.Name(0); } catch { }
            Log("Connexion à " + version);
            try { doc = app.OpenFile(cfg.DossierPath.Trim(), "", "", 0); }
            catch (Exception e)
            {
                Reset();
                throw new Exception(OpenFailureMessage(cfg.DossierPath, ComMessage(e)));
            }
            if (doc == null) { Reset(); throw new Exception("Sage a refusé d'ouvrir le dossier " + cfg.DossierPath); }
            Log("Dossier ouvert : " + (string)doc.Name);
            LoadClientFields();
        }

        // Explication adaptée au refus d'ouverture (avant 1.3.4, tout refus accusait à tort l'accès COM).
        static string OpenFailureMessage(string path, string sageMsg)
        {
            string head = "Sage a refusé d'ouvrir le dossier (" + sageMsg + "). ";
            string where = "Dans Sage : Dossier > Options > Informations, relevez le « Chemin d'accès » du dossier ouvert";
            string low = (sageMsg ?? "").ToLowerInvariant();
            if (Settings.IsOldCopy(path))
                return head + "Ce fichier « _OLD » est une ancienne copie laissée par Sage, pas le dossier en service. " + where + " et indiquez-le dans Réglages.";
            if (low.Contains("base locale") && !Settings.IsServerPath(path))
                return head + "Ce dossier est ouvert dans Sage par son serveur (mode Données Mobiles ou réseau). " + where + " (il commence par !Srv!) et collez-le tel quel dans Réglages > Dossier Sage 50.";
            return head + "Vérifiez que c'est bien le dossier ouvert dans Sage (" + where + ") et que l'autorisation d'accès COM est active (Utilitaires > Mode avancé > Divers > Autorisation accès COM).";
        }

        void LoadClientFields()
        {
            var map = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            dynamic t = doc.OpenTable(TBL_CLIENTS);
            try
            {
                string name = "";
                try { name = (string)t.TableName; } catch { }
                if (!string.Equals(name, "Cli", StringComparison.OrdinalIgnoreCase))
                    throw new Exception("La table des clients n'a pas été trouvée (table « " + name + " » à la place) : version de Sage 50 non prise en charge.");
                dynamic coll = t.Fields;
                int n = (int)coll.Count;
                for (int i = 1; i <= n; i++) // collection numérotée à partir de 1
                {
                    try { dynamic f = coll.Item(i); map[(string)f.Name] = (int)f.Id; } catch { }
                }
                int count = (int)t.NbN;
                Log("Table des clients : " + count + " fiche(s), " + map.Count + " rubrique(s).");
            }
            finally { try { Marshal.FinalReleaseComObject(t); } catch { } }
            cliFields = map;
        }

        int F(string name)
        {
            int id;
            if (cliFields != null && cliFields.TryGetValue(name, out id)) return id;
            return CliFieldDefaults[name];
        }

        public string Connect()
        {
            return Run<string>(delegate
            {
                try { EnsureOpen(); return (string)doc.Name; }
                catch (Exception e) { Reset(); throw new Exception(ComMessage(e)); }
            });
        }

        // Libère la connexion COM sans jamais bloquer plus de timeoutMs (la libération peut attendre Sage).
        public void Disconnect(int timeoutMs)
        {
            var done = new ManualResetEvent(false);
            queue.Add(delegate { try { Reset(); } catch { } done.Set(); });
            done.WaitOne(timeoutMs);
        }

        public InvoiceResult Create(InvoiceRequest req, Action<InvoiceResult> onCompleted)
        {
            return Run<InvoiceResult>(delegate
            {
                InvoiceResult r;
                try { r = CreateInner(req); }
                catch (Exception first)
                {
                    // Une reconnexion suffit souvent (Sage redémarré entre-temps, etc.).
                    Log("Échec (" + ComMessage(first) + "), nouvelle tentative après reconnexion…");
                    Reset();
                    try { r = CreateInner(req); }
                    catch (Exception second)
                    {
                        Reset();
                        r = new InvoiceResult();
                        r.Error = ComMessage(second);
                    }
                }
                // Appelé depuis le thread worker, AVANT que l'HTTP ne reprenne : l'historique
                // et le journal reflètent toujours ce que Sage a réellement fait, même si
                // l'HTTP a abandonné sur timeout. Sans ça, une facture créée dans Sage par
                // une rafale longue pouvait rester invisible côté SageBridge → doublon à la
                // retentative de l'extension (incident ISSYMA 30/10/2026, FG1733).
                if (onCompleted != null) { try { onCompleted(r); } catch { } }
                return r;
            });
        }

        // ── Clients ──

        class ClientRow
        {
            public string Code, NomBrut, Nom, ModeTVA, Nii, PaysBrut, VilleBrut;
            public string[] Fac, Liv; // [nom/société, adresse 1, CP, ville, pays]
            public bool Prospect;
        }

        static string[] Block(CustomerInfo c)
        {
            return new[] { Norm.Text(c.DisplayName), Norm.Text(c.Address1), Norm.Digits(c.PostalCode), Norm.Text(c.City), Norm.CountryKeyWithPostal(c.Country, c.PostalCode) };
        }

        static bool SameBlock(string[] want, string nomClient, string[] have)
        {
            if (want[1].Length == 0 || want[3].Length == 0) return false; // pas d'adresse exploitable
            bool nameOk = want[0].Length > 0 && (want[0] == have[0] || want[0] == nomClient);
            return nameOk && want[1] == have[1] && want[2] == have[2] && want[3] == have[3] && want[4] == have[4];
        }

        string V(dynamic t, string field)
        {
            try { object o = t.GetVal(F(field)); return o == null ? "" : Convert.ToString(o, CultureInfo.InvariantCulture).Trim(); }
            catch { return ""; }
        }

        static string FirstNonEmpty(string a, string b) { return string.IsNullOrEmpty(a) ? b : a; }

        // Point d'entrée : retourne la liste des clients. Essaie dans l'ordre : cache mémoire,
        // cache fichier, balayage complet de la table. La validité du cache repose sur NbN
        // (nombre de fiches côté Sage). NbN seul suffit pour détecter un ajout/suppression ;
        // une modification de fiche existante ne le fait pas bouger — supprimer manuellement
        // le fichier .cache force alors une reconstruction.
        List<ClientRow> ReadClients()
        {
            string dossier = "";
            try { dossier = (string)doc.Name; } catch { }

            int nbN = -1;
            dynamic probe = doc.OpenTable(TBL_CLIENTS);
            try { nbN = (int)probe.NbN; }
            finally { try { Marshal.FinalReleaseComObject(probe); } catch { } }

            if (cachedClients != null && cachedClientsDossier == dossier && cachedClientsNbN == nbN)
                return cachedClients;

            List<ClientRow> rows = LoadClientsCache(dossier, nbN);
            if (rows == null) rows = ScanClientsAndSaveCache(dossier, nbN);

            cachedClients = rows;
            cachedClientsDossier = dossier;
            cachedClientsNbN = nbN;
            return rows;
        }

        ClientRow BuildClientRow(dynamic t)
        {
            var r = new ClientRow();
            r.Code = V(t, "Code");
            r.NomBrut = V(t, "Nom");
            r.Nom = Norm.Text(r.NomBrut);
            r.ModeTVA = V(t, "ModeTVA");
            r.Nii = V(t, "NII");
            r.PaysBrut = V(t, "FacCtyName");
            r.VilleBrut = (V(t, "FacCP") + " " + V(t, "FacVille")).Trim();
            r.Fac = new string[] { Norm.Text(V(t, "FacSociete")), Norm.Text(V(t, "FacAdresse1")), Norm.Digits(V(t, "FacCP")), Norm.Text(V(t, "FacVille")), Norm.CountryKey(FirstNonEmpty(V(t, "FacCtyCode"), V(t, "FacCtyName"))) };
            r.Liv = new string[] { Norm.Text(V(t, "LivSociete")), Norm.Text(V(t, "LivAdresse1")), Norm.Digits(V(t, "LivCP")), Norm.Text(V(t, "LivVille")), Norm.CountryKey(FirstNonEmpty(V(t, "LivCtyCode"), V(t, "LivCtyName"))) };
            string pr = V(t, "IsProspect").ToLowerInvariant();
            r.Prospect = pr == "true" || pr == "1" || pr == "-1" || pr == "oui";
            return r;
        }

        List<ClientRow> ScanClientsAndSaveCache(string dossier, int nbN)
        {
            DateTime start = DateTime.Now;
            var rows = new List<ClientRow>();
            dynamic t = doc.OpenTable(TBL_CLIENTS);
            try
            {
                bool ok = (bool)t.First();
                while (ok)
                {
                    ClientRow r = BuildClientRow(t);
                    if (!string.IsNullOrEmpty(r.Code)) rows.Add(r);
                    ok = (bool)t.Next();
                }
            }
            finally { try { Marshal.FinalReleaseComObject(t); } catch { } }
            SaveClientsCache(dossier, nbN, rows);
            Log("Index clients construit : " + rows.Count + " fiche(s) en " + ((int)(DateTime.Now - start).TotalSeconds) + " s.");
            return rows;
        }

        static string ClientsCachePathFor(string dossier)
        {
            string d = string.IsNullOrEmpty(dossier) ? "default" : dossier;
            foreach (char ch in Path.GetInvalidFileNameChars()) d = d.Replace(ch, '_');
            return Path.Combine(Settings.Dir, "clients-" + d + ".cache");
        }

        List<ClientRow> LoadClientsCache(string dossier, int nbN)
        {
            string path = ClientsCachePathFor(dossier);
            if (!File.Exists(path)) return null;
            try
            {
                string[] lines = File.ReadAllLines(path, new UTF8Encoding(true));
                if (lines.Length < 4) return null;
                if (lines[0] != "# SageBridge cache clients v" + CLIENTS_CACHE_VERSION) return null;
                if (lines[1] != "# dossier=" + dossier) return null;
                if (lines[2] != "# nbN=" + nbN) return null;
                var rows = new List<ClientRow>(nbN);
                for (int i = 4; i < lines.Length; i++)
                {
                    ClientRow r = DeserializeClientRow(lines[i]);
                    if (r != null && !string.IsNullOrEmpty(r.Code)) rows.Add(r);
                }
                Log("Index clients chargé depuis le cache (" + rows.Count + " fiche(s)).");
                return rows;
            }
            catch (Exception e) { Log("Index clients : cache illisible (" + e.Message + "), reconstruction."); return null; }
        }

        void SaveClientsCache(string dossier, int nbN, List<ClientRow> rows)
        {
            string path = ClientsCachePathFor(dossier);
            try
            {
                using (var w = new StreamWriter(path, false, new UTF8Encoding(true)))
                {
                    w.WriteLine("# SageBridge cache clients v" + CLIENTS_CACHE_VERSION);
                    w.WriteLine("# dossier=" + dossier);
                    w.WriteLine("# nbN=" + nbN);
                    w.WriteLine("# built=" + DateTime.Now.ToString("s", CultureInfo.InvariantCulture));
                    foreach (ClientRow r in rows) w.WriteLine(SerializeClientRow(r));
                }
            }
            catch (Exception e) { Log("Index clients non enregistré : " + e.Message); }
        }

        static string Clean(string s)
        {
            if (s == null) return "";
            return s.Replace('\t', ' ').Replace('\r', ' ').Replace('\n', ' ');
        }

        static string SerializeClientRow(ClientRow r)
        {
            var cols = new string[17];
            cols[0] = Clean(r.Code);
            cols[1] = Clean(r.NomBrut);
            cols[2] = Clean(r.ModeTVA);
            cols[3] = Clean(r.Nii);
            cols[4] = Clean(r.PaysBrut);
            cols[5] = Clean(r.VilleBrut);
            cols[6] = r.Prospect ? "1" : "0";
            for (int i = 0; i < 5; i++) cols[7 + i] = Clean(r.Fac != null && i < r.Fac.Length ? r.Fac[i] : "");
            for (int i = 0; i < 5; i++) cols[12 + i] = Clean(r.Liv != null && i < r.Liv.Length ? r.Liv[i] : "");
            return string.Join("\t", cols);
        }

        static ClientRow DeserializeClientRow(string line)
        {
            if (line == null || line.Length == 0) return null;
            string[] cols = line.Split('\t');
            if (cols.Length < 17) return null;
            var r = new ClientRow();
            r.Code = cols[0];
            r.NomBrut = cols[1];
            r.Nom = Norm.Text(r.NomBrut);
            r.ModeTVA = cols[2];
            r.Nii = cols[3];
            r.PaysBrut = cols[4];
            r.VilleBrut = cols[5];
            r.Prospect = cols[6] == "1";
            r.Fac = new string[] { cols[7], cols[8], cols[9], cols[10], cols[11] };
            r.Liv = new string[] { cols[12], cols[13], cols[14], cols[15], cols[16] };
            return r;
        }

        // Appelée après la création réussie d'un client : ajoute la nouvelle fiche au cache
        // (mémoire + fichier) pour que la prochaine facture n'ait pas à rebalayer la table.
        void AddToClientsCache(string code, CustomerInfo c, CustomerInfo shipping, string modeTva, string vat, string ctyKey)
        {
            if (cachedClients == null) return; // ResolveClient a toujours chargé le cache avant d'arriver ici
            CustomerInfo lv = (shipping != null && !shipping.IsEmpty) ? shipping : c;
            string lvKey = Norm.CountryKeyWithPostal(lv.Country, lv.PostalCode);
            var r = new ClientRow();
            r.Code = code;
            r.NomBrut = c.DisplayName;
            r.Nom = Norm.Text(r.NomBrut);
            r.ModeTVA = modeTva;
            r.Nii = vat;
            r.PaysBrut = Norm.CountryName(ctyKey, c.Country);
            r.VilleBrut = ((c.PostalCode == null ? "" : c.PostalCode) + " " + (c.City == null ? "" : c.City)).Trim();
            r.Fac = new string[] { Norm.Text(c.DisplayName), Norm.Text(c.Address1), Norm.Digits(c.PostalCode), Norm.Text(c.City), ctyKey };
            r.Liv = new string[] { Norm.Text(lv.DisplayName), Norm.Text(lv.Address1), Norm.Digits(lv.PostalCode), Norm.Text(lv.City), lvKey };
            r.Prospect = false;
            cachedClients.Add(r);
            cachedClientsNbN++;
            SaveClientsCache(cachedClientsDossier, cachedClientsNbN, cachedClients);
        }

        public List<Dictionary<string, object>> ListClients()
        {
            return Run<List<Dictionary<string, object>>>(delegate
            {
                EnsureOpen();
                var outp = new List<Dictionary<string, object>>();
                foreach (ClientRow r in ReadClients())
                {
                    var d = new Dictionary<string, object>();
                    d["code"] = r.Code; d["nom"] = r.NomBrut; d["ville"] = r.VilleBrut; d["pays"] = r.PaysBrut;
                    d["modeTVA"] = r.ModeTVA; d["tva"] = r.Nii; d["prospect"] = r.Prospect;
                    outp.Add(d);
                }
                return outp;
            });
        }

        // Cache (par instance de worker) de l'id de la table articles — 0 tant
        // que non vérifié, >0 une fois ouvert avec succès et TableName confirmé.
        // Pas de valeur négative mise en cache : si la validation échoue sur
        // l'instance courante, chaque appel retentera la validation (c'est léger
        // tant que ça marche la 1ʳᵉ fois).
        int tblArticlesCache = 0;

        int ResolveArticleTableId()
        {
            if (tblArticlesCache > 0) return tblArticlesCache;
            dynamic t = null;
            try
            {
                t = doc.OpenTable(TBL_ARTICLES);
                string name = null;
                try { name = (string)t.TableName; } catch { }
                if (name == "Article")
                {
                    tblArticlesCache = TBL_ARTICLES;
                    return TBL_ARTICLES;
                }
                throw new InvalidOperationException("Table articles : nom inattendu « " + name + " » (attendu « Article »).");
            }
            finally { if (t != null) { try { Marshal.FinalReleaseComObject(t); } catch { } } }
        }

        // Liste les articles du dossier (code + nom court). Résout les rubriques
        // par nom (ITable.Fields) avec un repli sur les valeurs par défaut
        // observées sur Sage 50 v6.71. Mêmes garde-fous que `Create` : si Sage
        // a été redémarré entre-temps (RPC indisponible, HRESULT 0x800706BA),
        // on relâche la connexion et on retente une fois.
        public List<Dictionary<string, object>> ListArticles()
        {
            return Run<List<Dictionary<string, object>>>(delegate
            {
                try { return ListArticlesInner(); }
                catch (Exception first)
                {
                    Log("Échec (" + ComMessage(first) + "), nouvelle tentative après reconnexion…");
                    Reset();
                    return ListArticlesInner();
                }
            });
        }

        List<Dictionary<string, object>> ListArticlesInner()
        {
            EnsureOpen();
            int tableId = ResolveArticleTableId();
            var outp = new List<Dictionary<string, object>>();
            dynamic t = doc.OpenTable(tableId);
            try
            {
                int fidCode = 1, fidName = 2;
                try
                {
                    dynamic coll = t.Fields;
                    int n = (int)coll.Count;
                    for (int i = 1; i <= n; i++)
                    {
                        dynamic f = null;
                        try { f = coll.Item(i); } catch { continue; }
                        string nm = null;
                        try { nm = (string)f.Name; } catch { }
                        int id = 0;
                        try { id = (int)f.Id; } catch { }
                        if (nm == "Code") fidCode = id;
                        else if (nm == "DescShort" || nm == "Description" || nm == "Nom" || nm == "Intitule") fidName = id;
                    }
                }
                catch { }
                bool more = false;
                try { more = (bool)t.First(); } catch { }
                int safety = 50000;
                while (more && safety-- > 0)
                {
                    try
                    {
                        object oCode = t.GetVal(fidCode);
                        string code = oCode == null ? "" : Convert.ToString(oCode, CultureInfo.InvariantCulture);
                        if (!string.IsNullOrEmpty(code))
                        {
                            object oNom = null;
                            try { oNom = t.GetVal(fidName); } catch { }
                            string nom = oNom == null ? "" : Convert.ToString(oNom, CultureInfo.InvariantCulture);
                            var d = new Dictionary<string, object>();
                            d["code"] = code;
                            d["nom"] = nom;
                            outp.Add(d);
                        }
                    }
                    catch { }
                    try { more = (bool)t.Next(); } catch { break; }
                }
            }
            finally { try { Marshal.FinalReleaseComObject(t); } catch { } }
            return outp;
        }

        // ── Pièces clients (factures, avoirs…) — lecture seule ──
        // Table PieceHCli (kWgcTableIdPieceHCli = 97004), vérifiée live le 03/11/2026 sur Sage 50 v6.71.
        // Rubriques résolues par nom, avec ces valeurs de secours (eWgcTablePieceHCliFieldIds).
        public const int TBL_PIECES_CLI = 97004;
        static readonly Dictionary<string, int> PieceFieldDefaults = new Dictionary<string, int> {
            { "Type", 2 }, { "TypePiece", 3 }, { "Num", 4 }, { "DtPiece", 5 }, { "TotalHT", 22 }, { "TotalTTC", 24 },
            { "SoldeDu", 25 }, { "Client", 36 }, { "Societe", 37 }, { "MtRegle", 54 }, { "Valide", 56 }, { "Solde", 57 },
            { "CliName", 125 }, { "Statut", 160 }, { "ModeRglt", 50 }
        };

        // Toutes les pièces clients. Les pièces soldées ne sont lues qu'en partie (n°, nom, HT) :
        // elles ne servent qu'à signaler « déjà soldée ». Délai porté à 10 min (gros dossiers mobiles).
        public List<PieceRow> ListPieces(Action<int, int> progress)
        {
            return Run<List<PieceRow>>(delegate
            {
                try { return ListPiecesInner(progress); }
                catch (Exception first)
                {
                    Log("Échec (" + ComMessage(first) + "), nouvelle tentative après reconnexion…");
                    Reset();
                    return ListPiecesInner(progress);
                }
            }, 600000);
        }

        static string ValStr(dynamic t, int fid)
        {
            try { object o = t.GetVal(fid); return o == null ? "" : Convert.ToString(o, CultureInfo.InvariantCulture).Trim(); }
            catch { return ""; }
        }

        static double ValNum(dynamic t, int fid)
        {
            try { object o = t.GetVal(fid); return o == null ? 0 : Convert.ToDouble(o, CultureInfo.InvariantCulture); }
            catch { return 0; }
        }

        static bool ValBool(dynamic t, int fid)
        {
            string s = ValStr(t, fid).ToLowerInvariant();
            return s == "1" || s == "true" || s == "-1" || s == "oui";
        }

        List<PieceRow> ListPiecesInner(Action<int, int> progress)
        {
            EnsureOpen();
            var outp = new List<PieceRow>();
            dynamic t = doc.OpenTable(TBL_PIECES_CLI);
            try
            {
                string tname = "";
                try { tname = (string)t.TableName; } catch { }
                if (!string.Equals(tname, "PieceHCli", StringComparison.OrdinalIgnoreCase))
                    throw new Exception("La table des pièces clients n'a pas été trouvée (table « " + tname + " » à la place) : version de Sage 50 non prise en charge.");
                var f = new Dictionary<string, int>(PieceFieldDefaults, StringComparer.OrdinalIgnoreCase);
                try
                {
                    dynamic coll = t.Fields;
                    int n = (int)coll.Count;
                    for (int i = 1; i <= n; i++)
                    {
                        try { dynamic fd = coll.Item(i); string nm = (string)fd.Name; if (f.ContainsKey(nm)) f[nm] = (int)fd.Id; } catch { }
                    }
                }
                catch { }
                int total = 0;
                try { total = (int)t.NbN; } catch { }
                bool more = false;
                try { more = (bool)t.First(); } catch { }
                int done = 0, safety = 500000;
                while (more && safety-- > 0)
                {
                    done++;
                    if (progress != null && (done % 25 == 0 || done == total)) progress(done, total);
                    var p = new PieceRow();
                    p.Num = ValStr(t, f["Num"]);
                    if (p.Num.Length > 0)
                    {
                        p.TypePiece = ValStr(t, f["TypePiece"]);
                        p.Societe = ValStr(t, f["Societe"]);
                        p.CliName = ValStr(t, f["CliName"]);
                        p.TotalHT = ValNum(t, f["TotalHT"]);
                        p.Soldee = ValBool(t, f["Solde"]);
                        p.ModeRglt = ValStr(t, f["ModeRglt"]);
                        if (!p.Soldee)
                        {
                            p.Type = ValStr(t, f["Type"]);
                            p.Date = ValStr(t, f["DtPiece"]);
                            p.Client = ValStr(t, f["Client"]);
                            p.TotalTTC = ValNum(t, f["TotalTTC"]);
                            p.SoldeDu = ValNum(t, f["SoldeDu"]);
                            p.MtRegle = ValNum(t, f["MtRegle"]);
                            p.Valide = ValBool(t, f["Valide"]);
                            p.Statut = ValStr(t, f["Statut"]);
                        }
                        else p.Statut = "Soldé";
                        outp.Add(p);
                    }
                    try { more = (bool)t.Next(); } catch { break; }
                }
                Log("Pièces clients lues : " + outp.Count + ".");
            }
            finally { try { Marshal.FinalReleaseComObject(t); } catch { } }
            return outp;
        }

        // ── Modes de paiement (table Mop, kCtaTableIdMop = 105003, vérifiée live le 03/11/2026) ──
        public const int TBL_MODES_PAIEMENT = 105003;

        // [code, intitulé] de chaque mode de paiement du dossier.
        public List<string[]> ListModes()
        {
            return Run<List<string[]>>(delegate
            {
                EnsureOpen();
                var outp = new List<string[]>();
                dynamic t = doc.OpenTable(TBL_MODES_PAIEMENT);
                try
                {
                    string tname = "";
                    try { tname = (string)t.TableName; } catch { }
                    if (!string.Equals(tname, "Mop", StringComparison.OrdinalIgnoreCase))
                        throw new Exception("La table des modes de paiement n'a pas été trouvée (table « " + tname + " » à la place).");
                    int fCode = 1, fIntit = 2;
                    try
                    {
                        dynamic coll = t.Fields;
                        int n = (int)coll.Count;
                        for (int i = 1; i <= n; i++)
                        {
                            try { dynamic fd = coll.Item(i); string nm = (string)fd.Name; if (nm == "Code") fCode = (int)fd.Id; else if (nm == "Intit") fIntit = (int)fd.Id; } catch { }
                        }
                    }
                    catch { }
                    bool more = false;
                    try { more = (bool)t.First(); } catch { }
                    int safety = 10000;
                    while (more && safety-- > 0)
                    {
                        string code = ValStr(t, fCode);
                        string intit = ValStr(t, fIntit);
                        if (code.Length > 0) outp.Add(new string[] { code, intit });
                        try { more = (bool)t.Next(); } catch { break; }
                    }
                }
                finally { try { Marshal.FinalReleaseComObject(t); } catch { } }
                return outp;
            });
        }

        // ── Règlement client : solde une facture ──
        // Moteur WgcEngRegCli (eEngineIds.kWgcEngIdRegCli = 97003). Un règlement par facture, du montant
        // restant dû (TTC), date du jour. Mode de paiement = celui CHOISI PAR L'UTILISATEUR dans la liste
        // des modes du dossier (table Mop) juste avant de solder (demande de Boris, 03/11/2026) :
        // SageBridge n'en impose aucun.
        // simulate = tout sauf Save. Pas de nouvelle tentative automatique : un Save réussi suivi d'une
        // erreur de lecture ne doit pas produire un second règlement.
        public const int ENG_REGLEMENT_CLI = 97003;

        public SolderResult SolderFacture(string numFacture, string clientCode, double montant, string reference, string mode, bool simulate)
        {
            return Run<SolderResult>(delegate
            {
                EnsureOpen();
                var res = new SolderResult();
                dynamic reg = doc.CreateEngine(ENG_REGLEMENT_CLI);
                dynamic verif = app.CreateVerif();
                try
                {
                    if (!(bool)reg.Create()) throw new Exception("Sage n'a pas pu créer le règlement.");
                    reg.Tiers = clientCode;
                    reg.Date = DateTime.Today;
                    if (!string.IsNullOrEmpty(mode)) reg.ModePayment = mode;
                    if (!string.IsNullOrEmpty(reference)) { try { reg.reference = reference; } catch (Exception e) { res.Messages.Add("Référence non enregistrée : " + ComMessage(e)); } }
                    reg.Montant = montant;
                    // N° d'échéance : 1 refusé par Sage 50 v6.71 (« Impossible de lire l'échéance '1' ») → 0 d'abord, puis 1.
                    bool added = false; string addErr = "";
                    foreach (int eche in new[] { 0, 1 })
                    {
                        try { if ((bool)reg.AddFacture(numFacture, eche, montant)) { added = true; break; } }
                        catch (Exception e) { addErr = ComMessage(e); }
                    }
                    if (!added)
                        throw new Exception("Sage refuse d'imputer le règlement sur la facture " + numFacture + " (client " + clientCode + ")" + (addErr.Length > 0 ? " : " + addErr : "."));
                    try { res.Numero = (string)reg.Numero; } catch { }
                    try { res.Ecart = (double)reg.Ecart; } catch { }
                    try { res.TotalFac = (double)reg.TotalFac; } catch { }
                    try { res.Lignes = (int)reg.LineNb; } catch { }
                    try { res.Mode = (string)reg.ModePayment ?? ""; } catch { }
                    if (Math.Abs(res.Ecart) > 0.005)
                        throw new Exception("Écart de règlement de " + res.Ecart.ToString("0.00") + " € sur " + numFacture + " : rien n'est enregistré.");
                    if (simulate) { res.Ok = true; res.Simulated = true; return res; }
                    bool saved = (bool)reg.Save(0, verif);
                    try { int total = (int)verif.Total; for (int k = 1; k <= total; k++) { try { res.Messages.Add((string)verif.Desc(k)); } catch { } } } catch { }
                    if (!saved) throw new Exception("Sage a refusé le règlement : " + string.Join(" / ", res.Messages.ToArray()));
                    res.Ok = true;
                    Log("Règlement " + res.Numero + " enregistré : facture " + numFacture + " (client " + clientCode + ") soldée, " + montant.ToString("0.00") + " €" + (res.Mode.Length > 0 ? " (" + res.Mode + ")" : "") + (string.IsNullOrEmpty(reference) ? "" : ", réf. " + reference) + ".");
                    return res;
                }
                finally
                {
                    try { Marshal.FinalReleaseComObject(reg); } catch { }
                    try { Marshal.FinalReleaseComObject(verif); } catch { }
                }
            });
        }

        // Outil de développement : ouvre une table et liste ses rubriques (GET /diag?id=…&fids=…).
        public Dictionary<string, object> Diag(int tableId, int[] fids)
        {
            return Run<Dictionary<string, object>>(delegate
            {
                EnsureOpen();
                var d = new Dictionary<string, object>();
                d["tableId"] = tableId;
                dynamic t = doc.OpenTable(tableId);
                try
                {
                    try { d["tableName"] = (string)t.TableName; } catch (Exception e) { d["tableName"] = "ERR " + ComMessage(e); }
                    try { d["tableTitle"] = (string)t.TableTitle; } catch (Exception e) { d["tableTitle"] = "ERR " + ComMessage(e); }
                    try { d["nbN"] = (int)t.NbN; } catch (Exception e) { d["nbN"] = "ERR " + ComMessage(e); }
                    try { d["first"] = (bool)t.First(); } catch (Exception e) { d["first"] = "ERR " + ComMessage(e); }
                    try { d["recorName"] = (string)t.RecorName; } catch (Exception e) { d["recorName"] = "ERR " + ComMessage(e); }
                    var vals = new Dictionary<string, object>();
                    foreach (int fid in fids)
                    {
                        try { object o = t.GetVal(fid); vals[fid.ToString()] = o == null ? "<null>" : Convert.ToString(o, CultureInfo.InvariantCulture); }
                        catch (Exception e) { vals[fid.ToString()] = "ERR " + ComMessage(e); }
                    }
                    d["vals"] = vals;
                    var fields = new List<string>();
                    try
                    {
                        dynamic coll = t.Fields;
                        int n = (int)coll.Count;
                        d["fieldCount"] = n;
                        for (int i = 1; i <= n; i++)
                        {
                            try { dynamic f = coll.Item(i); fields.Add((int)f.Id + " | " + (string)f.Name + " | " + (string)f.Title); }
                            catch (Exception e) { fields.Add("ERR " + i + " : " + ComMessage(e)); }
                        }
                    }
                    catch (Exception e) { fields.Add("ERR Fields : " + ComMessage(e)); }
                    d["fields"] = fields;
                }
                finally { try { Marshal.FinalReleaseComObject(t); } catch { } }
                return d;
            });
        }

        // ── Export PDF des pièces (1.4.0) ──
        // Tâche « Impression pièce client » (eTaskIds.kWgcTaskIdPieceCliPrint, valeur COMPLÈTE : contrairement aux
        // tables, l'octet haut distingue les tâches). Params = ITable « PieceCli » ; rubriques vérifiées live le 03/11/2026 :
        // 1 Sel (sélection d'UD de PieceHCli : SEULE façon de désigner la pièce, PieceId/ManNum seuls → rien),
        // 2 NoAskConfirm, 4 PieceType, 20 PrnDst (liste : Demander, Imprimante, Ecran, Silencieux, « Ficher PDF »,
        // Envoyer par e-mail), 38 PrnReportName (nom du modèle). « Ficher PDF » → boîte « Paramètres incorrects »
        // (aucun champ pour le chemin) : on passe donc par l'aperçu « Ecran », puis Exporter en PDF (SagePdfExporter).
        public const int TASK_PIECE_CLI_PRINT = 0x12017AF5;

        // Factures et avoirs VALIDÉS dont la date de pièce tombe dans le mois donné (les provisoires n'ont pas
        // encore leur numéro définitif). Lecture seule.
        public List<PdfPiece> ListPiecesForMonth(int year, int month)
        {
            return Run<List<PdfPiece>>(delegate
            {
                try { return ListPiecesForMonthInner(year, month); }
                catch (Exception first)
                {
                    Log("Échec (" + ComMessage(first) + "), nouvelle tentative après reconnexion…");
                    Reset();
                    return ListPiecesForMonthInner(year, month);
                }
            }, 600000);
        }

        static DateTime ValDate(dynamic t, int fid)
        {
            try
            {
                object o = t.GetVal(fid);
                if (o is DateTime) return (DateTime)o;
                DateTime d; string s = Convert.ToString(o, CultureInfo.InvariantCulture);
                if (DateTime.TryParse(s, CultureInfo.InvariantCulture, DateTimeStyles.None, out d)) return d;
                if (DateTime.TryParse(s, CultureInfo.GetCultureInfo("fr-FR"), DateTimeStyles.None, out d)) return d;
            }
            catch { }
            return DateTime.MinValue;
        }

        List<PdfPiece> ListPiecesForMonthInner(int year, int month)
        {
            EnsureOpen();
            var outp = new List<PdfPiece>();
            dynamic t = doc.OpenTable(TBL_PIECES_CLI);
            try
            {
                string tname = "";
                try { tname = (string)t.TableName; } catch { }
                if (!string.Equals(tname, "PieceHCli", StringComparison.OrdinalIgnoreCase))
                    throw new Exception("La table des pièces clients n'a pas été trouvée (table « " + tname + " » à la place) : version de Sage 50 non prise en charge.");
                var f = new Dictionary<string, int>(PieceFieldDefaults, StringComparer.OrdinalIgnoreCase);
                try
                {
                    dynamic coll = t.Fields;
                    int n = (int)coll.Count;
                    for (int i = 1; i <= n; i++)
                    {
                        try { dynamic fd = coll.Item(i); string nm = (string)fd.Name; if (f.ContainsKey(nm)) f[nm] = (int)fd.Id; } catch { }
                    }
                }
                catch { }
                bool more = false;
                try { more = (bool)t.First(); } catch { }
                int safety = 500000;
                while (more && safety-- > 0)
                {
                    DateTime dt = ValDate(t, f["DtPiece"]);
                    if (dt.Year == year && dt.Month == month)
                    {
                        var p = new PdfPiece();
                        p.Num = ValStr(t, f["Num"]);
                        p.TypePiece = ValStr(t, f["TypePiece"]);
                        p.Date = dt;
                        string soc = ValStr(t, f["Societe"]);
                        p.Nom = soc.Length > 0 ? soc : ValStr(t, f["CliName"]);
                        p.TotalTTC = ValNum(t, f["TotalTTC"]);
                        p.Valide = ValBool(t, f["Valide"]);
                        try { p.UD = (int)t.CurUD; } catch { }
                        string k = Norm.Text(p.TypePiece);
                        p.Kind = k.Contains("avoir") ? "avoir" : (k.Contains("facture") && !k.Contains("acompte") ? "facture" : "");
                        if (p.Num.Length > 0 && p.Kind.Length > 0 && p.Valide && p.UD != 0) outp.Add(p);
                    }
                    try { more = (bool)t.Next(); } catch { break; }
                }
            }
            finally { try { Marshal.FinalReleaseComObject(t); } catch { } }
            outp.Sort(delegate(PdfPiece a, PdfPiece b) { int c = string.CompareOrdinal(a.Kind, b.Kind); return c != 0 ? -c : string.CompareOrdinal(a.Num, b.Num); });
            Log("Pièces validées de " + month.ToString("00") + "/" + year + " : " + outp.Count + ".");
            return outp;
        }

        // Ouvre l'aperçu écran de la pièce avec le modèle donné. Rend la main dès que l'aperçu est affiché
        // (la fenêtre d'aperçu ne bloque pas le thread COM : vérifié live, ~1 s).
        public void OpenPreview(PdfPiece p, string model)
        {
            Run<bool>(delegate
            {
                EnsureOpen();
                dynamic task = doc.CreateTask(TASK_PIECE_CLI_PRINT);
                if (task == null) throw new Exception("Sage n'a pas créé la tâche d'impression.");
                dynamic prm = null, sel = null;
                try
                {
                    prm = task.Params;
                    sel = app.CreateSelection();
                    sel.AddUD(p.UD);
                    prm.SetVal(1, sel);
                    prm.SetVal(2, 1);
                    prm.SetVal(4, p.Kind == "avoir" ? 5 : TYPE_FACTURE);
                    prm.SetVal(20, "Ecran");
                    if (!string.IsNullOrEmpty(model)) prm.SetVal(38, model);
                    task.Run();
                    return true;
                }
                finally
                {
                    try { if (sel != null) Marshal.FinalReleaseComObject(sel); } catch { }
                    try { if (prm != null) Marshal.FinalReleaseComObject(prm); } catch { }
                    try { Marshal.FinalReleaseComObject(task); } catch { }
                }
            });
        }

        // Code suivant en respectant le format dominant du dossier (ex. « 281 » → « 282 », « CLI0041 » → « CLI0042 »).
        static string NextClientCode(List<ClientRow> rows)
        {
            var re = new Regex(@"^(\D*)(\d+)$");
            var byPrefix = new Dictionary<string, int[]>(); // préfixe → [nombre, max, largeur]
            foreach (ClientRow r in rows)
            {
                Match m = re.Match(r.Code);
                if (!m.Success) continue;
                string pre = m.Groups[1].Value;
                long n; if (!long.TryParse(m.Groups[2].Value, out n) || n > int.MaxValue) continue;
                int[] st;
                if (!byPrefix.TryGetValue(pre, out st)) { st = new int[3]; byPrefix[pre] = st; }
                st[0]++; st[1] = Math.Max(st[1], (int)n); st[2] = Math.Max(st[2], m.Groups[2].Value.Length);
            }
            string best = ""; int[] bestSt = null;
            foreach (KeyValuePair<string, int[]> kv in byPrefix)
                if (bestSt == null || kv.Value[0] > bestSt[0]) { best = kv.Key; bestSt = kv.Value; }
            if (bestSt == null) return "001";
            return best + (bestSt[1] + 1).ToString(new string('0', Math.Max(3, bestSt[2])));
        }

        static string IncrementCode(string code)
        {
            Match m = Regex.Match(code, @"^(\D*)(\d+)$");
            if (!m.Success) return code + "1";
            long n = long.Parse(m.Groups[2].Value) + 1;
            return m.Groups[1].Value + n.ToString(new string('0', m.Groups[2].Value.Length));
        }

        void SetF(dynamic cli, string field, string value, InvoiceResult res, string label)
        {
            if (string.IsNullOrEmpty(value)) return;
            try
            {
                bool ok = (bool)cli.SetFieldIdValStr(F(field), value, 0, 0);
                if (!ok) res.Messages.Add("Client : rubrique « " + label + " » refusée par Sage.");
            }
            catch (Exception e) { res.Messages.Add("Client : rubrique « " + label + " » : " + ComMessage(e)); }
        }

        // SIRET : 14 chiffres + clé de Luhn.
        static bool SiretValid(string siret)
        {
            string d = Norm.Digits(siret);
            if (d.Length != 14) return false;
            int sum = 0;
            for (int i = 0; i < 14; i++)
            {
                if (d[i] < '0' || d[i] > '9') return false;
                int n = d[i] - '0';
                if (i % 2 == 0) { n *= 2; if (n > 9) n -= 9; }
                sum += n;
            }
            return sum % 10 == 0;
        }

        // N° TVA français : FR + clé (12 + 3 × (SIREN mod 97)) mod 97 + SIREN ; étranger : non contrôlé ici.
        static bool VatValid(string vat, string siret)
        {
            string v = Norm.Digits(vat);
            if (v.Length < 4) return false;
            if (!v.StartsWith("FR")) return true;
            if (v.Length != 13) return false;
            long siren; int key;
            if (!long.TryParse(v.Substring(4), out siren) || !int.TryParse(v.Substring(2, 2), out key)) return false;
            if ((12 + 3 * (siren % 97)) % 97 != key) return false;
            string sd = Norm.Digits(siret);
            return !(sd.Length == 14 && sd.Substring(0, 9) != v.Substring(4));
        }

        // Retrouve le client (nom + adresse + CP + ville + pays, sans accents ni casse) ou le crée.
        // En simulation, rien n'est créé : renvoie null si le client serait créé.
        string ResolveClient(InvoiceRequest req, InvoiceResult res)
        {
            CustomerInfo c = req.Customer;
            if (c == null || c.IsEmpty) c = (req.Shipping != null && !req.Shipping.IsEmpty) ? req.Shipping : null;
            List<ClientRow> rows = ReadClients();

            if (c == null)
            {
                if (string.IsNullOrEmpty(req.Client))
                    throw new Exception("La commande ne contient pas d'adresse exploitable et aucun client par défaut n'est réglé (bouton Réglages).");
                foreach (ClientRow r in rows)
                    if (string.Equals(r.Code, req.Client, StringComparison.OrdinalIgnoreCase)) { res.ClientModeTVA = r.ModeTVA; res.ClientName = r.NomBrut; }
                res.ClientCode = req.Client;
                res.Messages.Add("Pas d'adresse exploitable dans la commande : client par défaut " + req.Client + ".");
                return req.Client;
            }

            string[] wantFac = Block(c);
            string[] wantLiv = (req.Shipping != null && !req.Shipping.IsEmpty) ? Block(req.Shipping) : null;
            foreach (ClientRow r in rows)
            {
                if (r.Prospect) continue;
                if (SameBlock(wantFac, r.Nom, r.Fac) || SameBlock(wantFac, r.Nom, r.Liv) || (wantLiv != null && SameBlock(wantLiv, r.Nom, r.Liv)))
                {
                    res.ClientCode = r.Code; res.ClientName = r.NomBrut; res.ClientMatched = true; res.ClientModeTVA = r.ModeTVA;
                    Log("Client existant : " + r.Code + " (" + r.NomBrut + ", TVA " + r.ModeTVA + ")");
                    return r.Code;
                }
            }

            string ctyKey = Norm.CountryKey(c.Country);
            if (!Norm.IsKnownCountry(ctyKey))
            {
                string fromCp = Norm.PostalCodeToDomTomKey(c.PostalCode);
                if (fromCp.Length > 0)
                {
                    res.Messages.Add("Pays « " + c.Country + " » reconnu comme « " + Norm.CountryName(fromCp, fromCp) + " » via le code postal " + c.PostalCode + ".");
                    ctyKey = fromCp;
                }
            }
            string modeTva = Norm.VatModeForCountry(ctyKey);
            if (!Norm.IsKnownCountry(ctyKey))
                res.Messages.Add("Pays « " + c.Country + " » non reconnu : mode TVA « " + modeTva + " » par défaut, à vérifier sur la fiche.");
            res.ClientName = c.DisplayName; res.ClientModeTVA = modeTva;

            if (req.Simulation)
            {
                res.Messages.Add("Simulation : aucun client existant ne correspond, la fiche serait créée (mode TVA « " + modeTva + " »).");
                return null;
            }

            dynamic cli = doc.CreateEngine(ENG_CLIENT);
            dynamic verif = app.CreateVerif();
            try
            {
                if (!(bool)cli.Create()) throw new Exception("Sage n'a pas pu créer la fiche client.");
                string code = NextClientCode(rows);
                try { int guard = 0; while ((bool)cli.Exist(code) && guard++ < 1000) code = IncrementCode(code); } catch { }
                cli.Code = code;
                cli.Name = c.DisplayName;
                SetF(cli, "TypeCli", "Professionnel", res, "type de client");   // obligatoire
                SetF(cli, "ModeTVA", modeTva, res, "mode TVA");
                SetF(cli, "FacSociete", c.DisplayName, res, "société");
                SetF(cli, "FacAdresse1", c.Address1, res, "adresse");
                SetF(cli, "FacAdresse2", c.Address2, res, "complément d'adresse");
                SetF(cli, "FacCP", c.PostalCode, res, "code postal");
                SetF(cli, "FacVille", c.City, res, "ville");
                // Nom AVANT code : Sage recouple code↔nom quand on écrit le nom.
                // Un nom non reconnu (ex. « États-Unis » avec accent, alors que la
                // table pays Sage a « Etats-Unis » sans accent) efface le code.
                // On écrit donc d'abord le nom normalisé (sans accents) puis le code.
                SetF(cli, "FacCtyName", Norm.CountryNameForSage(ctyKey, c.Country), res, "pays");
                SetF(cli, "FacCtyCode", Norm.CountryIso3(ctyKey), res, "code pays");
                SetF(cli, "FacTel", c.Phone, res, "téléphone");
                SetF(cli, "Email", c.Email, res, "e-mail");
                string siret = SiretValid(c.Siret) ? Norm.Digits(c.Siret) : "";
                if (!string.IsNullOrEmpty(c.Siret) && siret.Length == 0) res.Messages.Add("Client : SIRET « " + c.Siret + " » invalide, non enregistré.");
                string vat = VatValid(c.Vat, siret) ? Norm.Digits(c.Vat) : "";
                if (!string.IsNullOrEmpty(c.Vat) && vat.Length == 0) res.Messages.Add("Client : N° TVA « " + c.Vat + " » invalide ou incohérent avec le SIRET, non enregistré.");
                SetF(cli, "SIRET", siret, res, "SIRET");
                SetF(cli, "NII", vat, res, "N° TVA intracom");

                // Adresse de livraison : Sage la pré-remplit avec la facturation, les champs vides
                // sont donc effacés explicitement avec une espace.
                CustomerInfo l = (req.Shipping != null && !req.Shipping.IsEmpty) ? req.Shipping : c;
                string lKey = Norm.CountryKeyWithPostal(l.Country, l.PostalCode);
                SetF(cli, "LivSociete", l.DisplayName, res, "société (livraison)");
                SetF(cli, "LivAdresse1", l.Address1, res, "adresse (livraison)");
                SetF(cli, "LivAdresse2", string.IsNullOrEmpty(l.Address2) ? " " : l.Address2, res, "complément (livraison)");
                SetF(cli, "LivCP", l.PostalCode, res, "code postal (livraison)");
                SetF(cli, "LivVille", l.City, res, "ville (livraison)");
                SetF(cli, "LivCtyName", Norm.CountryNameForSage(lKey, l.Country), res, "pays (livraison)");
                SetF(cli, "LivCtyCode", Norm.CountryIso3(lKey), res, "code pays (livraison)");
                SetF(cli, "LivTel", FirstNonEmpty(l.Phone, FirstNonEmpty(c.Phone, " ")), res, "téléphone (livraison)");

                string outCode = code;
                bool saved = (bool)cli.Save(verif, ref outCode);
                if (!saved && (siret.Length > 0 || vat.Length > 0))
                {
                    // Sage peut refuser un identifiant pourtant plausible : on réessaie sans plutôt que de bloquer la facture.
                    var first = new List<string>();
                    try { int tt = (int)verif.Total; for (int k = 1; k <= tt; k++) { try { first.Add((string)verif.Desc(k)); } catch { } } } catch { }
                    res.Messages.Add("Client : SIRET / N° TVA retirés après refus de Sage (" + string.Join(" / ", first.ToArray()) + ").");
                    try { verif.Reset(); } catch { }
                    SetF(cli, "SIRET", " ", res, "SIRET");
                    SetF(cli, "NII", " ", res, "N° TVA intracom");
                    outCode = code;
                    saved = (bool)cli.Save(verif, ref outCode);
                }
                try { int total = (int)verif.Total; for (int k = 1; k <= total; k++) { try { res.Messages.Add("Client : " + (string)verif.Desc(k)); } catch { } } } catch { }
                if (!saved) throw new Exception("Sage a refusé la création du client « " + c.DisplayName + " » : " + string.Join(" / ", res.Messages.ToArray()));
                if (!string.IsNullOrEmpty(outCode)) code = outCode;
                res.ClientCode = code; res.ClientCreated = true;
                Log("Client créé : " + code + " (" + c.DisplayName + ", " + Norm.CountryName(ctyKey, c.Country) + ", TVA " + modeTva + ")");
                AddToClientsCache(code, c, req.Shipping, modeTva, vat, ctyKey);
                return code;
            }
            finally
            {
                try { Marshal.FinalReleaseComObject(cli); } catch { }
                try { Marshal.FinalReleaseComObject(verif); } catch { }
            }
        }

        // ── Facture ──

        InvoiceResult CreateInner(InvoiceRequest req)
        {
            EnsureOpen();
            var res = new InvoiceResult();
            string clientCode = ResolveClient(req, res);

            string mode = (res.ClientModeTVA ?? "").Trim();
            bool exonere = mode.Length > 0 && !mode.Equals("Local", StringComparison.OrdinalIgnoreCase);
            if (exonere) res.Messages.Add("Client en mode TVA « " + mode + " » : lignes facturées à 0 % de TVA.");

            if (req.Simulation)
            {
                res.Ok = true; res.Simulated = true;
                if (clientCode != null) res.Messages.Add("Simulation : la facture serait créée pour le client " + clientCode + ".");
                res.NbLines = req.Lines.Count;
                foreach (InvoiceLine l in req.Lines) res.TotalHT += Math.Round(l.Qty * l.Puht, 2);
                return res;
            }

            dynamic piece = doc.CreateEngine(ENG_PIECE_CLI);
            dynamic line = doc.CreateEngine(ENG_LIGNE_CLI);
            dynamic verif = app.CreateVerif();
            try
            {
                if (!(bool)piece.Create(TYPE_FACTURE)) throw new Exception("Sage n'a pas pu créer la facture.");
                piece.Client = clientCode;
                string nomClient = "";
                try { nomClient = (string)piece.Name; } catch { }
                if (string.IsNullOrEmpty(nomClient)) throw new Exception("Client « " + clientCode + " » inconnu dans le dossier Sage.");
                if (string.IsNullOrEmpty(res.ClientName)) res.ClientName = nomClient;
                // La Référence de la facture reste vide sauf si l'option est activée dans Réglages (demande de Boris, 20/10/2026).
                if (cfg.WriteReference && !string.IsNullOrEmpty(req.Reference)) { try { piece.reference = req.Reference; } catch (Exception e) { res.Messages.Add("Référence non enregistrée : " + ComMessage(e)); } }
                if (!string.IsNullOrEmpty(req.Remarque)) { try { piece.Remarque = req.Remarque; } catch (Exception e) { res.Messages.Add("Remarque non enregistrée : " + ComMessage(e)); } }

                int n = 0;
                foreach (InvoiceLine l in req.Lines)
                {
                    line.RazLine();
                    line.Article = l.Code;
                    line.Quantite = l.Qty;
                    if (l.HasPuht) line.PUHT = l.Puht;   // AddLine garde ce prix (AddLineAutoCalc le remplacerait par le tarif)
                    double tva = exonere ? 0 : l.Tva;
                    if (tva >= 0) { try { line.TauxTVA = tva; } catch (Exception e) { res.Messages.Add("TVA ligne " + (n + 1) + " : " + ComMessage(e)); } }
                    string desc = Norm.RtfToText(l.Description);
                    if (!string.IsNullOrEmpty(desc)) { try { line.Description = desc; } catch (Exception e) { res.Messages.Add("Description ligne " + (n + 1) + " : " + ComMessage(e)); } }
                    if (!(bool)piece.AddLine(line)) throw new Exception("Ligne " + (n + 1) + " (article « " + l.Code + " ») refusée par Sage : l'article existe-t-il dans le dossier ?");
                    n++;
                }
                res.NbLines = n;
                // Lire numéro et totaux AVANT Save : après Save, le moteur est vidé
                // (« Aucune pièce client n'a été initialisée par le composant COM »).
                try { res.Numero = (string)piece.Numero; } catch { }
                try { res.TotalHT = (double)piece.TotalHT; } catch { }
                try { res.TotalTTC = (double)piece.TotalTTC; } catch { }

                bool saved = (bool)piece.Save(verif);
                try { int total = (int)verif.Total; for (int k = 1; k <= total; k++) { try { res.Messages.Add((string)verif.Desc(k)); } catch { } } } catch { }
                if (!saved) throw new Exception("Sage a refusé l'enregistrement de la facture : " + string.Join(" / ", res.Messages.ToArray()));
                res.Ok = true;
                return res;
            }
            finally
            {
                try { Marshal.FinalReleaseComObject(line); } catch { }
                try { Marshal.FinalReleaseComObject(piece); } catch { }
                try { Marshal.FinalReleaseComObject(verif); } catch { }
            }
        }
    }

    // ───────────────────────── Mises à jour (GitHub Releases) ─────────────────────────
    //  Chaque version publiée est une « release » GitHub du dépôt public (tag vX.Y.Z) portant
    //  l'installateur SageBridge-Setup.exe et son empreinte SageBridge-Setup.exe.sha256.
    //  SageBridge interroge l'API publique (sans compte) au démarrage puis toutes les heures.

    class UpdateInfo
    {
        public Version Version;
        public string Tag = "", Name = "", Notes = "", PageUrl = "", SetupUrl = "", ShaUrl = "";
        public long Size;
        public string VersionText { get { return Program.ShortVersion(Version); } }
    }

    static class Updater
    {
        public const string DefaultRepo = "Gnaris/SageBridge";
        public const string SetupAsset = "SageBridge-Setup.exe";
        public const string UninstallKey = @"Software\Microsoft\Windows\CurrentVersion\Uninstall\SageBridge";

        static Updater()
        {
            // .NET 4 n'active pas TLS 1.2 par défaut sur tous les Windows : GitHub l'exige.
            try { ServicePointManager.SecurityProtocol |= (SecurityProtocolType)3072; } catch { }
        }

        public static Version Current
        {
            get { Version v = Assembly.GetExecutingAssembly().GetName().Version; return new Version(v.Major, v.Minor, Math.Max(0, v.Build)); }
        }

        static HttpWebRequest Request(string url, string accept)
        {
            var r = (HttpWebRequest)WebRequest.Create(url);
            r.UserAgent = "SageBridge/" + Program.Version + " (+https://github.com/" + DefaultRepo + ")";
            r.Accept = accept;
            r.Timeout = 20000;
            r.ReadWriteTimeout = 60000;
            r.AllowAutoRedirect = true;
            try { if (r.Proxy != null) r.Proxy.Credentials = CredentialCache.DefaultCredentials; } catch { }
            return r;
        }

        public static Version ParseVersion(string tag)
        {
            Match m = Regex.Match(tag ?? "", @"(\d+)\.(\d+)(?:\.(\d+))?");
            if (!m.Success) return null;
            return new Version(int.Parse(m.Groups[1].Value), int.Parse(m.Groups[2].Value), m.Groups[3].Success ? int.Parse(m.Groups[3].Value) : 0);
        }

        // Dernière version publiée, ou null si le dépôt n'a encore aucune version.
        public static UpdateInfo Latest(string repo)
        {
            string body;
            try
            {
                using (var resp = (HttpWebResponse)Request("https://api.github.com/repos/" + repo + "/releases/latest", "application/vnd.github+json").GetResponse())
                using (var rd = new StreamReader(resp.GetResponseStream(), Encoding.UTF8)) body = rd.ReadToEnd();
            }
            catch (WebException e)
            {
                var hr = e.Response as HttpWebResponse;
                if (hr != null && hr.StatusCode == HttpStatusCode.NotFound) return null;
                if (hr != null && (int)hr.StatusCode == 403) throw new Exception("GitHub limite temporairement les vérifications (réessai plus tard).");
                throw new Exception("GitHub injoignable (" + e.Message + ").");
            }
            var d = new JavaScriptSerializer().Deserialize<Dictionary<string, object>>(body);
            var u = new UpdateInfo();
            u.Tag = Str(d, "tag_name"); u.Name = Str(d, "name"); u.Notes = Str(d, "body"); u.PageUrl = Str(d, "html_url");
            u.Version = ParseVersion(u.Tag);
            if (u.Version == null) return null;
            object assets;
            if (d.TryGetValue("assets", out assets) && assets is System.Collections.IEnumerable)
            {
                foreach (object o in (System.Collections.IEnumerable)assets)
                {
                    var a = o as Dictionary<string, object>;
                    if (a == null) continue;
                    string name = Str(a, "name");
                    if (string.Equals(name, SetupAsset, StringComparison.OrdinalIgnoreCase))
                    {
                        u.SetupUrl = Str(a, "browser_download_url");
                        long sz; long.TryParse(Str(a, "size"), out sz); u.Size = sz;
                    }
                    else if (string.Equals(name, SetupAsset + ".sha256", StringComparison.OrdinalIgnoreCase)) u.ShaUrl = Str(a, "browser_download_url");
                }
            }
            return u;
        }

        static string Str(Dictionary<string, object> d, string k)
        {
            object v; return d.TryGetValue(k, out v) && v != null ? Convert.ToString(v, CultureInfo.InvariantCulture) : "";
        }

        static bool TrustedUrl(string url)
        {
            Uri u;
            if (!Uri.TryCreate(url, UriKind.Absolute, out u) || u.Scheme != Uri.UriSchemeHttps) return false;
            string h = u.Host.ToLowerInvariant();
            return h == "github.com" || h.EndsWith(".github.com") || h.EndsWith(".githubusercontent.com");
        }

        // Télécharge l'installateur dans %TEMP%, contrôle taille, empreinte SHA-256 et identité du fichier.
        public static string Download(UpdateInfo u)
        {
            if (string.IsNullOrEmpty(u.SetupUrl) || !TrustedUrl(u.SetupUrl)) throw new Exception("La version " + u.VersionText + " ne contient pas d'installateur.");
            if (string.IsNullOrEmpty(u.ShaUrl) || !TrustedUrl(u.ShaUrl)) throw new Exception("La version " + u.VersionText + " ne contient pas d'empreinte SHA-256 : mise à jour ignorée par sécurité.");
            string dir = Path.Combine(Path.GetTempPath(), "SageBridge-maj");
            Directory.CreateDirectory(dir);
            string path = Path.Combine(dir, "SageBridge-Setup-" + u.VersionText + ".exe");

            string expected;
            using (var resp = Request(u.ShaUrl, "application/octet-stream").GetResponse())
            using (var rd = new StreamReader(resp.GetResponseStream())) expected = rd.ReadToEnd();
            Match m = Regex.Match(expected, "[0-9a-fA-F]{64}");
            if (!m.Success) throw new Exception("Empreinte SHA-256 illisible.");
            expected = m.Value.ToLowerInvariant();

            string part = path + ".part";
            using (var resp = Request(u.SetupUrl, "application/octet-stream").GetResponse())
            using (Stream src = resp.GetResponseStream())
            using (var dst = new FileStream(part, FileMode.Create, FileAccess.Write))
            {
                var buf = new byte[81920]; int n; long total = 0;
                while ((n = src.Read(buf, 0, buf.Length)) > 0)
                {
                    dst.Write(buf, 0, n); total += n;
                    if (total > 200L * 1024 * 1024) throw new Exception("Fichier anormalement volumineux.");
                }
            }
            string actual;
            using (var sha = System.Security.Cryptography.SHA256.Create())
            using (var fs = File.OpenRead(part))
                actual = BitConverter.ToString(sha.ComputeHash(fs)).Replace("-", "").ToLowerInvariant();
            if (actual != expected) { try { File.Delete(part); } catch { } throw new Exception("Empreinte SHA-256 différente de celle publiée : fichier rejeté."); }
            var fvi = FileVersionInfo.GetVersionInfo(part);
            if (fvi.ProductName != "SageBridge") { try { File.Delete(part); } catch { } throw new Exception("Le fichier téléchargé n'est pas un installateur SageBridge."); }
            if (File.Exists(path)) File.Delete(path);
            File.Move(part, path);
            return path;
        }

        // Dossier d'installation (inscrit par l'installateur), vide si SageBridge n'a pas été installé.
        public static string InstallDir
        {
            get
            {
                try
                {
                    using (RegistryKey k = Registry.CurrentUser.OpenSubKey(UninstallKey))
                        return k == null ? "" : ((string)k.GetValue("InstallLocation") ?? "").TrimEnd('\\');
                }
                catch { return ""; }
            }
        }

        // Vrai si ce programme est la copie installée (les mises à jour automatiques ne s'appliquent qu'à elle).
        public static bool IsInstalledCopy
        {
            get
            {
                string dir = InstallDir;
                return dir.Length > 0 && string.Equals(Path.GetFullPath(dir).TrimEnd('\\'), Path.GetDirectoryName(Application.ExecutablePath).TrimEnd('\\'), StringComparison.OrdinalIgnoreCase);
            }
        }
    }

    // ───────────────────────── Historique (anti-doublon) ─────────────────────────

    class History
    {
        readonly Dictionary<string, string> sent = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        readonly List<string[]> rows = new List<string[]>();
        readonly object sync = new object();

        public History()
        {
            if (!File.Exists(Settings.HistoryFile)) return;
            foreach (string l in File.ReadAllLines(Settings.HistoryFile, Encoding.UTF8))
            {
                string[] p = l.Split(';');
                if (p.Length < 7 || p[0] == "date") continue;
                sent[Key(p[1], p[2])] = p[3];
                rows.Add(p);
            }
        }

        static string Key(string site, string orderNo) { return (site ?? "") + "|" + (orderNo ?? ""); }

        public string Existing(string site, string orderNo)
        {
            lock (sync) { string v; return sent.TryGetValue(Key(site, orderNo), out v) ? v : null; }
        }

        // N° de facture créé pour ce n° de commande (tous sites confondus), ou null.
        public string FindByOrder(string orderNo)
        {
            if (string.IsNullOrEmpty(orderNo)) return null;
            lock (sync)
            {
                for (int i = rows.Count - 1; i >= 0; i--)
                    if (string.Equals(rows[i][2].Trim(), orderNo.Trim(), StringComparison.OrdinalIgnoreCase)) return rows[i][3];
            }
            return null;
        }

        public List<string[]> Recent(int max)
        {
            lock (sync)
            {
                var r = new List<string[]>();
                for (int i = rows.Count - 1; i >= 0 && r.Count < max; i--) r.Add(rows[i]);
                return r;
            }
        }

        public void Add(string site, string orderNo, string numero, string client, int nbLines, double totalHT)
        {
            lock (sync)
            {
                sent[Key(site, orderNo)] = numero;
                string[] p = { DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"), site, orderNo, numero, client, nbLines.ToString(), totalHT.ToString("0.00", CultureInfo.InvariantCulture) };
                rows.Add(p);
                bool fresh = !File.Exists(Settings.HistoryFile);
                using (var w = new StreamWriter(Settings.HistoryFile, true, Encoding.UTF8))
                {
                    if (fresh) w.WriteLine("date;site;commande;facture;client;lignes;totalHT");
                    w.WriteLine(string.Join(";", p));
                }
            }
        }
    }

    // ───────────────────────── Serveur HTTP local ─────────────────────────

    class Server
    {
        public Action<string> Log = delegate { };
        public Func<InvoiceRequest, InvoiceResult> OnInvoice;
        public Func<Dictionary<string, object>> OnStatus;
        public Func<int, int[], Dictionary<string, object>> OnDiag;
        public Func<List<Dictionary<string, object>>> OnClients;
        public Func<List<Dictionary<string, object>>> OnArticles;
        public Func<List<PieceRow>> OnPieces;
        public Func<List<string[]>> OnModes;
        public Func<string, string, double, string, string, bool, SolderResult> OnSolder;
        HttpListener listener;
        int port;
        readonly JavaScriptSerializer json = new JavaScriptSerializer();

        public void Start(int p)
        {
            port = p;
            listener = new HttpListener();
            listener.Prefixes.Add("http://127.0.0.1:" + port + "/");
            listener.Prefixes.Add("http://localhost:" + port + "/");
            listener.Start();
            listener.BeginGetContext(OnContext, null);
        }

        public void Stop()
        {
            try { if (listener != null) { listener.Stop(); listener.Close(); } } catch { }
            listener = null;
        }

        void OnContext(IAsyncResult ar)
        {
            HttpListener l = listener;
            if (l == null) return;
            HttpListenerContext ctx;
            try { ctx = l.EndGetContext(ar); } catch { return; }
            try { l.BeginGetContext(OnContext, null); } catch { }
            ThreadPool.QueueUserWorkItem(delegate { Handle(ctx); });
        }

        // Sécurité : seuls les extensions de navigateur, les pages de SageBridge lui-même et les outils
        // locaux (PowerShell, scripts : pas d'en-tête Origin) sont acceptés. Une page web quelconque
        // ouverte dans le navigateur ne peut ni créer de facture ni lire la liste des clients.
        // Les admins des boutiques tenant (BJ, Issyma) doivent pouvoir lire
        // /articles pour peupler le dropdown « Code article Sage » quand on
        // mappe une catégorie. Chaque admin tourne dans le navigateur de son
        // propriétaire, donc 127.0.0.1 reste local.
        static readonly string[] AllowedShopOrigins = new[] {
            "https://beliandjolie.com", "https://www.beliandjolie.com",
            "https://issyma.fr",        "https://www.issyma.fr",
            "http://localhost:3000"
        };

        bool OriginAllowed(HttpListenerRequest r, out string origin)
        {
            origin = r.Headers["Origin"];
            if (!string.IsNullOrEmpty(origin))
            {
                string o = origin.Trim().ToLowerInvariant();
                if (o.StartsWith("chrome-extension://") || o.StartsWith("moz-extension://") || o.StartsWith("safari-web-extension://")
                    || o == "http://127.0.0.1:" + port || o == "http://localhost:" + port) return true;
                foreach (string allowed in AllowedShopOrigins) if (o == allowed) return true;
                return false;
            }
            string site = (r.Headers["Sec-Fetch-Site"] ?? "").ToLowerInvariant();
            return !(site == "cross-site" || site == "same-site");
        }

        void Handle(HttpListenerContext ctx)
        {
            var resp = ctx.Response;
            resp.Headers["X-Content-Type-Options"] = "nosniff";
            resp.Headers["Cache-Control"] = "no-store";
            try
            {
                string origin;
                if (!OriginAllowed(ctx.Request, out origin))
                {
                    Log("Requête refusée (origine non autorisée : " + (string.IsNullOrEmpty(origin) ? ctx.Request.Headers["Sec-Fetch-Site"] : origin) + ").");
                    Send(resp, 403, "{\"ok\":false,\"error\":\"Origine non autorisée\"}");
                    return;
                }
                if (!string.IsNullOrEmpty(origin))
                {
                    resp.Headers["Access-Control-Allow-Origin"] = origin;
                    resp.Headers["Vary"] = "Origin";
                    resp.Headers["Access-Control-Allow-Methods"] = "GET, POST, OPTIONS";
                    resp.Headers["Access-Control-Allow-Headers"] = "Content-Type";
                    resp.Headers["Access-Control-Allow-Private-Network"] = "true";
                }
                string path = ctx.Request.Url.AbsolutePath.TrimEnd('/');
                string method = ctx.Request.HttpMethod;
                if (method == "OPTIONS") { Send(resp, 204, ""); return; }
                if (path == "" || path == "/ping")
                {
                    Dictionary<string, object> st = OnStatus != null ? OnStatus() : new Dictionary<string, object>();
                    st["ok"] = true; st["app"] = Program.AppName; st["version"] = Program.Version;
                    Send(resp, 200, json.Serialize(st)); return;
                }
                if (path == "/clients" && method == "GET")
                {
                    Send(resp, 200, json.Serialize(OnClients())); return;
                }
                if (path == "/articles" && method == "GET")
                {
                    Send(resp, 200, json.Serialize(OnArticles())); return;
                }
                if (path == "/factures" && method == "GET")
                {
                    // Lecture seule. Par défaut : pièces non soldées ; ?tout=1 pour toutes.
                    bool tout = ctx.Request.QueryString["tout"] == "1";
                    var outp = new List<Dictionary<string, object>>();
                    foreach (PieceRow p in OnPieces())
                    {
                        if (p.Soldee && !tout) continue;
                        var d = new Dictionary<string, object>();
                        d["numero"] = p.Num; d["type"] = p.Type; d["typePiece"] = p.TypePiece; d["nature"] = p.Kind;
                        d["date"] = p.Date; d["client"] = p.Client; d["societe"] = p.Societe; d["nomClient"] = p.CliName;
                        d["totalHT"] = p.TotalHT; d["totalTTC"] = p.TotalTTC; d["soldeDu"] = p.SoldeDu; d["mtRegle"] = p.MtRegle;
                        d["validee"] = p.Valide; d["soldee"] = p.Soldee; d["statut"] = p.Statut; d["modeReglement"] = p.ModeRglt;
                        outp.Add(d);
                    }
                    var js = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };
                    Send(resp, 200, js.Serialize(outp)); return;
                }
                if (path == "/modes" && method == "GET")
                {
                    var outp = new List<Dictionary<string, object>>();
                    foreach (string[] m in OnModes()) outp.Add(new Dictionary<string, object> { { "code", m[0] }, { "intitule", m[1] } });
                    Send(resp, 200, json.Serialize(outp)); return;
                }
                if (path == "/solder" && method == "POST")
                {
                    // { facture, client, montant, reference, simulation } — simulation par défaut (seul "simulation": false enregistre).
                    string body;
                    using (var r = new StreamReader(ctx.Request.InputStream, Encoding.UTF8)) body = r.ReadToEnd();
                    var d = (Dictionary<string, object>)json.DeserializeObject(body);
                    object sim; bool simulate = !(d.TryGetValue("simulation", out sim) && sim is bool && !(bool)sim);
                    var o = new Dictionary<string, object>();
                    try
                    {
                        SolderResult sr = OnSolder(S(d, "facture"), S(d, "client"), D(d, "montant", 0), S(d, "reference"), S(d, "mode"), simulate);
                        o["ok"] = sr.Ok; o["simulation"] = sr.Simulated; o["reglement"] = sr.Numero; o["mode"] = sr.Mode; o["ecart"] = sr.Ecart;
                        o["totalFactures"] = sr.TotalFac; o["lignes"] = sr.Lignes; o["messages"] = sr.Messages;
                        Send(resp, 200, json.Serialize(o));
                    }
                    catch (Exception e) { o["ok"] = false; o["error"] = e.Message; Send(resp, 500, json.Serialize(o)); }
                    return;
                }
                if (path == "/diag" && method == "GET")
                {
                    int id; if (!int.TryParse(ctx.Request.QueryString["id"], out id)) id = SageWorker.TBL_CLIENTS;
                    var fl = new List<int>();
                    foreach (string part in (ctx.Request.QueryString["fids"] ?? "1,2,3").Split(','))
                    {
                        int v; if (int.TryParse(part.Trim(), out v)) fl.Add(v);
                    }
                    Send(resp, 200, json.Serialize(OnDiag(id, fl.ToArray()))); return;
                }
                if (path == "/facture" && method == "POST")
                {
                    if (ctx.Request.ContentLength64 > 2 * 1024 * 1024) { Send(resp, 413, "{\"ok\":false,\"error\":\"Requête trop volumineuse\"}"); return; }
                    string body;
                    using (var r = new StreamReader(ctx.Request.InputStream, Encoding.UTF8)) body = r.ReadToEnd();
                    InvoiceResult res = OnInvoice(Parse(body));
                    var o = new Dictionary<string, object>();
                    o["ok"] = res.Ok; o["numero"] = res.Numero; o["totalHT"] = res.TotalHT; o["totalTTC"] = res.TotalTTC;
                    o["nbLines"] = res.NbLines; o["error"] = res.Error; o["duplicate"] = res.Duplicate; o["simulation"] = res.Simulated;
                    o["clientCode"] = res.ClientCode; o["clientName"] = res.ClientName; o["clientCreated"] = res.ClientCreated;
                    o["clientMatched"] = res.ClientMatched; o["clientModeTVA"] = res.ClientModeTVA; o["messages"] = res.Messages;
                    Send(resp, res.Ok ? 200 : (res.Duplicate ? 409 : 500), json.Serialize(o)); return;
                }
                Send(resp, 404, "{\"ok\":false,\"error\":\"Adresse inconnue\"}");
            }
            catch (Exception e)
            {
                Log("Requête en erreur : " + e.Message);
                try { Send(resp, 400, json.Serialize(new Dictionary<string, object> { { "ok", false }, { "error", e.Message } })); } catch { }
            }
        }

        static void Send(HttpListenerResponse resp, int status, string body)
        {
            byte[] b = Encoding.UTF8.GetBytes(body ?? "");
            resp.StatusCode = status;
            resp.ContentType = "application/json; charset=utf-8";
            resp.ContentLength64 = b.Length;
            resp.OutputStream.Write(b, 0, b.Length);
            resp.OutputStream.Close();
        }

        static string S(Dictionary<string, object> d, string k)
        {
            object v; return d.TryGetValue(k, out v) && v != null ? Convert.ToString(v, CultureInfo.InvariantCulture).Trim() : "";
        }

        static double D(Dictionary<string, object> d, string k, double def)
        {
            object v;
            if (!d.TryGetValue(k, out v) || v == null) return def;
            if (v is double) return (double)v;
            if (v is int) return (int)v;
            if (v is long) return (long)v;
            if (v is decimal) return (double)(decimal)v;
            double x;
            return double.TryParse(Convert.ToString(v, CultureInfo.InvariantCulture).Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out x) ? x : def;
        }

        static bool B(Dictionary<string, object> d, string k)
        {
            object v; return d.TryGetValue(k, out v) && v is bool && (bool)v;
        }

        static CustomerInfo ParseCustomer(Dictionary<string, object> d, string key)
        {
            object o; if (!d.TryGetValue(key, out o)) return null;
            var cd = o as Dictionary<string, object>; if (cd == null) return null;
            var c = new CustomerInfo();
            c.Company = S(cd, "company"); c.Contact = S(cd, "contact"); c.Address1 = S(cd, "address1"); c.Address2 = S(cd, "address2");
            c.PostalCode = S(cd, "postalCode"); c.City = S(cd, "city"); c.Country = S(cd, "country"); c.Phone = S(cd, "phone");
            c.Email = S(cd, "email"); c.Siret = S(cd, "siret"); c.Vat = S(cd, "vat");
            return c;
        }

        InvoiceRequest Parse(string body)
        {
            Dictionary<string, object> d;
            try { d = json.Deserialize<Dictionary<string, object>>(body); }
            catch (Exception e) { throw new Exception("JSON invalide : " + e.Message); }
            if (d == null) throw new Exception("JSON vide.");
            var req = new InvoiceRequest();
            req.Site = S(d, "site"); req.OrderNo = S(d, "orderNo"); req.Client = S(d, "client");
            req.Reference = S(d, "reference"); req.Remarque = S(d, "remarque");
            req.Force = B(d, "force"); req.Simulation = B(d, "simulation");
            req.Customer = ParseCustomer(d, "customer");
            req.Shipping = ParseCustomer(d, "shipping");
            object lines;
            if (!d.TryGetValue("lines", out lines) || !(lines is System.Collections.IEnumerable)) throw new Exception("Aucune ligne (champ « lines »).");
            foreach (object o in (System.Collections.IEnumerable)lines)
            {
                var ld = o as Dictionary<string, object>;
                if (ld == null) continue;
                var l = new InvoiceLine();
                l.Code = S(ld, "code"); l.Qty = D(ld, "qty", 0); l.Puht = D(ld, "puht", 0); l.HasPuht = ld.ContainsKey("puht");
                l.Tva = D(ld, "tva", 20); l.Description = S(ld, "description");
                if (string.IsNullOrEmpty(l.Code)) throw new Exception("Une ligne n'a pas de code article.");
                if (l.Qty <= 0) throw new Exception("Quantité invalide pour l'article « " + l.Code + " ».");
                req.Lines.Add(l);
            }
            if (req.Lines.Count == 0) throw new Exception("Aucune ligne.");
            return req;
        }
    }

    // ───────────────────────── Fenêtre principale ─────────────────────────

    class MainForm : Form
    {
        static readonly Color Green = Color.FromArgb(10, 122, 69), Orange = Color.FromArgb(180, 90, 0), Red = Color.Firebrick;

        Settings cfg;
        SageWorker sage;
        Server server;
        History history;
        bool reallyClose;
        float scale = 1f;

        Label lblStatus, lblSage, lblDossier, lblUrl, lblVersion;
        int busy, checking;
        volatile bool updating;
        string failedVersion = "", notifiedVersion = "", latestVersion = "";
        System.Windows.Forms.Timer updateTimer;
        ListView list;
        TextBox txtLog;
        CheckBox chkStartup;
        Button btnTest;
        NotifyIcon tray;
        System.Windows.Forms.Timer sageTimer;
        EventWaitHandle showEvent;

        public MainForm()
        {
            cfg = Settings.Load();
            history = new History();
            BuildUi();
            LogLine("SageBridge " + Program.Version + " démarré.");

            sage = new SageWorker(cfg);
            sage.Log = LogLine;
            server = new Server();
            server.Log = LogLine;
            server.OnInvoice = HandleInvoice;
            server.OnStatus = StatusDict;
            server.OnDiag = delegate(int id, int[] fids) { return sage.Diag(id, fids); };
            server.OnClients = delegate { return sage.ListClients(); };
            server.OnArticles = delegate { return sage.ListArticles(); };
            server.OnPieces = delegate { return sage.ListPieces(null); };
            server.OnModes = delegate { return sage.ListModes(); };
            server.OnSolder = delegate(string num, string cli, double mt, string reference, string mode, bool sim)
            {
                if (updating) throw new Exception("SageBridge installe une mise à jour : réessayez dans une minute.");
                Interlocked.Increment(ref busy);
                try { return sage.SolderFacture(num, cli, mt, reference, mode, sim); }
                finally { Interlocked.Decrement(ref busy); }
            };
            StartServer();

            showEvent = new EventWaitHandle(false, EventResetMode.AutoReset, Program.ShowEventName);
            var waiter = new Thread(delegate()
            {
                while (true) { try { showEvent.WaitOne(); Ui(ShowWindow); } catch { return; } }
            });
            waiter.IsBackground = true;
            waiter.Start();
        }

        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);
            EnsureDossier();
            int killed = SageWorker.KillOrphans(LogLine);
            if (killed > 0) LogLine(killed + " instance(s) invisible(s) de Sage fermée(s) : Sage peut de nouveau s'ouvrir normalement.");
            UpdateSageIndicator();
            ThreadPool.QueueUserWorkItem(delegate { TryConnect(false); });
            if (Program.PreviousVersion.Length > 0)
            {
                LogLine("Mise à jour installée : " + Program.PreviousVersion + " → " + Program.Version + ".");
                try { tray.ShowBalloonTip(5000, "SageBridge mis à jour", "Version " + Program.Version + " installée (avant : " + Program.PreviousVersion + ").", ToolTipIcon.Info); } catch { }
            }
            ScheduleUpdates();
        }

        // ── Mises à jour : au démarrage puis toutes les heures ──

        void ScheduleUpdates()
        {
            var first = new System.Windows.Forms.Timer { Interval = 8000 };
            first.Tick += delegate { first.Stop(); first.Dispose(); CheckForUpdates(false); };
            first.Start();
            updateTimer = new System.Windows.Forms.Timer { Interval = 60 * 60 * 1000 };
            updateTimer.Tick += delegate { CheckForUpdates(false); };
            updateTimer.Start();
        }

        void SetVersionText(string text, Color color) { Ui(delegate { lblVersion.Text = text; lblVersion.ForeColor = color; }); }

        void CheckForUpdates(bool manual)
        {
            if (updating) { if (manual) MessageBox.Show(this, "Une mise à jour est déjà en cours d'installation.", "SageBridge"); return; }
            if (Interlocked.Exchange(ref checking, 1) == 1) return;
            SetVersionText("Version " + Program.Version + " — recherche de mise à jour…", Color.DimGray);
            string repo = cfg.Repo;
            ThreadPool.QueueUserWorkItem(delegate
            {
                try
                {
                    UpdateInfo u = Updater.Latest(repo);
                    string when = DateTime.Now.ToString("HH:mm");
                    latestVersion = u == null ? "" : u.VersionText;
                    if (u == null || u.Version <= Updater.Current)
                    {
                        SetVersionText("Version " + Program.Version + (u == null ? " — aucune version publiée sur GitHub (vérifié à " : " — à jour (vérifié à ") + when + ")", Color.DimGray);
                        if (manual) Ui(delegate
                        {
                            MessageBox.Show(this, u == null ? "Aucune version n'est encore publiée sur GitHub (" + repo + ")." : "SageBridge " + Program.Version + " est la dernière version.",
                                "Mises à jour", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        });
                        return;
                    }
                    SetVersionText("Version " + Program.Version + " — mise à jour " + u.VersionText + " disponible (vérifié à " + when + ")", Orange);
                    if (manual) { Ui(delegate { AskInstall(u); }); return; }
                    if (cfg.AutoUpdate && Updater.IsInstalledCopy && u.VersionText != failedVersion) { InstallUpdate(u, false); return; }
                    if (notifiedVersion != u.VersionText)
                    {
                        notifiedVersion = u.VersionText;
                        LogLine("Mise à jour disponible : " + u.VersionText + " — bouton « Mises à jour… » pour l'installer.");
                        Ui(delegate { try { tray.ShowBalloonTip(5000, "Mise à jour de SageBridge disponible", "Version " + u.VersionText + " : ouvrez SageBridge puis « Mises à jour… ».", ToolTipIcon.Info); } catch { } });
                    }
                }
                catch (Exception e)
                {
                    SetVersionText("Version " + Program.Version + " — vérification des mises à jour impossible (" + e.Message + ")", Color.DimGray);
                    if (manual) Ui(delegate { MessageBox.Show(this, e.Message, "Mises à jour", MessageBoxButtons.OK, MessageBoxIcon.Warning); });
                }
                finally { Interlocked.Exchange(ref checking, 0); }
            });
        }

        void AskInstall(UpdateInfo u)
        {
            string notes = (u.Notes ?? "").Trim();
            if (notes.Length > 1200) notes = notes.Substring(0, 1200) + "…";
            DialogResult r = MessageBox.Show(this, "SageBridge " + u.VersionText + " est disponible (installé : " + Program.Version + ").\n\n" +
                (notes.Length > 0 ? notes + "\n\n" : "") + "Installer maintenant ? SageBridge se fermera puis redémarrera tout seul.",
                "Mise à jour disponible", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
            if (r == DialogResult.Yes) ThreadPool.QueueUserWorkItem(delegate { InstallUpdate(u, true); });
        }

        bool DialogOpen()
        {
            bool open = false;
            try { Invoke((Action)delegate { open = Application.OpenForms.Count > 1; }); } catch { }
            return open;
        }

        // Thread de fond : télécharge, attend que SageBridge soit libre, lance l'installateur puis se ferme.
        void InstallUpdate(UpdateInfo u, bool manual)
        {
            if (updating) return;
            updating = true;
            try
            {
                LogLine("Téléchargement de la mise à jour " + u.VersionText + "…");
                string setup = Updater.Download(u);
                LogLine("Mise à jour " + u.VersionText + " téléchargée et vérifiée (SHA-256).");
                DateTime limit = DateTime.Now.AddMinutes(10);
                while ((Thread.VolatileRead(ref busy) > 0 || DialogOpen()) && DateTime.Now < limit) Thread.Sleep(2000);
                if (Thread.VolatileRead(ref busy) > 0 || DialogOpen())
                {
                    updating = false;
                    LogLine("Mise à jour reportée : SageBridge est occupé (nouvel essai dans une heure).");
                    return;
                }
                LogLine("Installation de la version " + u.VersionText + " : SageBridge redémarre.");
                Process.Start(new ProcessStartInfo(setup, "/update /pid=" + Process.GetCurrentProcess().Id) { UseShellExecute = true });
                Ui(delegate { reallyClose = true; Close(); });
            }
            catch (Exception e)
            {
                updating = false;
                failedVersion = u.VersionText;
                LogLine("Mise à jour " + u.VersionText + " impossible : " + e.Message);
                SetVersionText("Version " + Program.Version + " — mise à jour " + u.VersionText + " impossible : " + e.Message, Red);
                if (manual) Ui(delegate { MessageBox.Show(this, "La mise à jour a échoué :\n" + e.Message, "Mises à jour", MessageBoxButtons.OK, MessageBoxIcon.Warning); });
            }
        }

        int S(int px) { return (int)Math.Round(px * scale); }

        void BuildUi()
        {
            AutoScaleMode = AutoScaleMode.Font;
            Font = new Font("Segoe UI", 9.5f);
            using (Graphics g = CreateGraphics()) scale = g.DpiX / 96f;
            Text = "SageBridge " + Program.Version + " — factures Sage 50";
            Icon = Program.AppIcon;
            StartPosition = FormStartPosition.CenterScreen;
            Size = new Size(S(940), S(640));
            MinimumSize = new Size(S(760), S(500));

            // En-tête
            var header = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 2, Padding = new Padding(S(14), S(10), S(14), S(6)), BackColor = Color.White };
            header.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            header.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            lblStatus = new Label { AutoSize = true, Font = new Font("Segoe UI Semibold", 12.5f), Text = "Démarrage…", Margin = new Padding(0, 0, 0, S(4)) };
            lblSage = new Label { AutoSize = true, Font = new Font("Segoe UI Semibold", 9.5f), Text = "", Margin = new Padding(0, 0, 0, S(2)) };
            lblDossier = new Label { AutoSize = true, Text = "", ForeColor = Color.FromArgb(60, 60, 60), Margin = new Padding(0, 0, 0, S(2)) };
            lblUrl = new Label { AutoSize = true, ForeColor = Color.DimGray, Text = "", Margin = new Padding(0, 0, 0, S(2)) };
            lblVersion = new Label { AutoSize = true, ForeColor = Color.DimGray, Text = "Version " + Program.Version };
            var btnConnect = MakeButton("Reconnecter", delegate { ThreadPool.QueueUserWorkItem(delegate { TryConnect(true); }); });
            var btnSettings = MakeButton("Réglages…", delegate { OpenSettings(); });
            var btnUpdates = MakeButton("Mises à jour…", delegate { CheckForUpdates(true); });
            var rightButtons = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.TopDown, WrapContents = false, Margin = new Padding(0) };
            rightButtons.Controls.Add(btnConnect); rightButtons.Controls.Add(btnSettings); rightButtons.Controls.Add(btnUpdates);
            header.Controls.Add(lblStatus, 0, 0);
            header.Controls.Add(lblSage, 0, 1);
            header.Controls.Add(lblDossier, 0, 2);
            header.Controls.Add(lblUrl, 0, 3);
            header.Controls.Add(lblVersion, 0, 4);
            header.Controls.Add(rightButtons, 1, 0);
            header.SetRowSpan(rightButtons, 5);

            var lblList = new Label { Dock = DockStyle.Top, AutoSize = false, Height = S(26), Text = "Dernières factures créées", Font = new Font("Segoe UI Semibold", 9.5f), Padding = new Padding(S(12), S(6), 0, 0) };

            list = new ListView { Dock = DockStyle.Fill, View = View.Details, FullRowSelect = true, GridLines = true, HideSelection = false };
            string[] cols = { "Date", "Origine", "Commande", "Facture", "Client", "Lignes", "Total HT", "Résultat" };
            int[] widths = { 135, 70, 130, 100, 85, 55, 85, 250 };
            for (int i = 0; i < cols.Length; i++) list.Columns.Add(cols[i], S(widths[i]));

            var buttons = new FlowLayoutPanel { Dock = DockStyle.Bottom, AutoSize = true, Padding = new Padding(S(8), S(6), S(8), S(6)), WrapContents = true };
            btnTest = MakeButton("Facture de test…", delegate { CreateTestInvoice(); });
            buttons.Controls.Add(btnTest);
            buttons.Controls.Add(MakeButton("Ouvrir Sage 50", delegate { OpenSage(); }));
            buttons.Controls.Add(MakeButton("Rapport de paiement…", delegate { OpenPaymentReport(); }));
            buttons.Controls.Add(MakeButton("Enregistrer les PDF…", delegate { OpenPdfExport(); }));
            buttons.Controls.Add(MakeButton("Historique", delegate { OpenFile(Settings.HistoryFile, "Aucune facture n'a encore été créée."); }));
            buttons.Controls.Add(MakeButton("Journal", delegate { OpenFile(Settings.LogFile, "Le journal est vide."); }));
            var btnDoc = MakeButton("Documentation", delegate { OpenFile(Program.DocPath, "La documentation n'est pas installée à côté du programme."); });
            btnDoc.Enabled = File.Exists(Program.DocPath);
            buttons.Controls.Add(btnDoc);
            chkStartup = new CheckBox { Text = "Démarrer avec Windows", AutoSize = true, Margin = new Padding(S(14), S(8), 0, 0), Checked = IsStartup() };
            chkStartup.CheckedChanged += delegate { SetStartup(chkStartup.Checked); };
            buttons.Controls.Add(chkStartup);

            txtLog = new TextBox { Dock = DockStyle.Bottom, Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Vertical, Height = S(120), BackColor = Color.White, Font = new Font("Consolas", 9f) };

            Controls.Add(list);
            Controls.Add(lblList);
            Controls.Add(header);
            Controls.Add(buttons);
            Controls.Add(txtLog);

            foreach (string[] p in history.Recent(200)) AddHistoryRow(p);

            tray = new NotifyIcon { Icon = Program.AppIcon, Text = "SageBridge", Visible = true };
            tray.DoubleClick += delegate { ShowWindow(); };
            var menu = new ContextMenuStrip();
            menu.Items.Add("Afficher SageBridge", null, delegate { ShowWindow(); });
            menu.Items.Add("Quitter", null, delegate { reallyClose = true; Close(); });
            tray.ContextMenuStrip = menu;

            sageTimer = new System.Windows.Forms.Timer { Interval = 5000 };
            sageTimer.Tick += delegate { UpdateSageIndicator(); };
            sageTimer.Start();
            RefreshDossierLabel();
        }

        Button MakeButton(string text, EventHandler onClick)
        {
            var b = new Button { Text = text, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Padding = new Padding(S(6), S(2), S(6), S(2)), Margin = new Padding(S(3)) };
            b.Click += onClick;
            return b;
        }

        // ── Aides d'interface ──

        void Ui(Action a)
        {
            if (IsDisposed) return;
            if (IsHandleCreated && InvokeRequired) { try { BeginInvoke(a); } catch { } }
            else a();
        }

        void ShowWindow() { Show(); if (WindowState == FormWindowState.Minimized) WindowState = FormWindowState.Normal; Activate(); BringToFront(); }

        void LogLine(string s)
        {
            Settings.AppendLog(s);
            string line = DateTime.Now.ToString("HH:mm:ss") + "  " + s;
            Ui(delegate { txtLog.AppendText(line + "\r\n"); });
        }

        void SetStatus(string text, Color color) { Ui(delegate { lblStatus.Text = text; lblStatus.ForeColor = color; }); }

        void RefreshDossierLabel()
        {
            lblDossier.Text = "Dossier Sage : " + (string.IsNullOrEmpty(cfg.DossierPath) ? "aucun — choisissez-le dans Réglages" : cfg.DossierPath)
                + (string.IsNullOrEmpty(cfg.DefaultClient) ? "" : "   •   Client par défaut : " + cfg.DefaultClient);
        }

        void UpdateSageIndicator()
        {
            string title;
            if (SageWorker.SageWindow(out title))
            {
                int i = title.IndexOf(':');
                lblSage.Text = "● Sage 50 ouvert" + (i >= 0 ? " — " + title.Substring(i + 1).Trim() : "");
                lblSage.ForeColor = Green;
            }
            else
            {
                lblSage.Text = "● Sage 50 fermé : ouvrez-le pour pouvoir créer des factures";
                lblSage.ForeColor = Orange;
            }
        }

        void OpenFile(string path, string missing)
        {
            if (!File.Exists(path)) { MessageBox.Show(this, missing, "SageBridge", MessageBoxButtons.OK, MessageBoxIcon.Information); return; }
            try { Process.Start(path); } catch (Exception e) { MessageBox.Show(this, e.Message, "SageBridge", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
        }

        void StartServer()
        {
            try
            {
                server.Start(cfg.Port);
                lblUrl.Text = "Adresse locale : http://127.0.0.1:" + cfg.Port + "/  (réservée à ce PC)";
                SetStatus("En attente de factures", Color.FromArgb(70, 70, 70));
            }
            catch (Exception e)
            {
                lblUrl.Text = "";
                SetStatus("Impossible d'écouter sur le port " + cfg.Port + " (déjà utilisé ?) — changez-le dans Réglages", Red);
                LogLine("Port " + cfg.Port + " : " + e.Message);
            }
        }

        Dictionary<string, object> StatusDict()
        {
            var d = new Dictionary<string, object>();
            string title;
            d["sageOuvert"] = SageWorker.SageWindow(out title);
            d["sageFenetre"] = title;
            d["dossier"] = cfg.DossierPath;
            d["clientParDefaut"] = cfg.DefaultClient;
            d["derniereVersionPubliee"] = latestVersion;
            d["majAuto"] = cfg.AutoUpdate;
            d["copieInstallee"] = Updater.IsInstalledCopy;
            return d;
        }

        // ── Dossier Sage ──

        void EnsureDossier()
        {
            if (Settings.DossierExists(cfg.DossierPath)) return;
            List<string> found = Settings.FindDossiers();
            string chosen = null;
            if (found.Count == 1) chosen = found[0];
            else if (found.Count > 1) chosen = DossierPicker.Pick(this, found, cfg.DossierPath);
            else
            {
                MessageBox.Show(this, "Aucun dossier Sage 50 Gestion commerciale n'a été trouvé automatiquement.\nIndiquez son fichier (.SgWgc) dans la fenêtre suivante.", "SageBridge", MessageBoxButtons.OK, MessageBoxIcon.Information);
                chosen = BrowseDossier(cfg.DossierPath);
            }
            if (!string.IsNullOrEmpty(chosen))
            {
                cfg.DossierPath = chosen; cfg.Save();
                LogLine("Dossier Sage choisi : " + chosen);
            }
            RefreshDossierLabel();
        }

        public static string BrowseDossier(string current)
        {
            using (var dlg = new OpenFileDialog { Title = "Choisir le dossier Sage 50 (fichier .SgWgc)", Filter = "Dossier Sage 50 Gestion commerciale (*.SgWgc)|*.SgWgc|Tous les fichiers (*.*)|*.*" })
            {
                try { if (!string.IsNullOrEmpty(current)) dlg.InitialDirectory = Path.GetDirectoryName(current); } catch { }
                return dlg.ShowDialog() == DialogResult.OK ? dlg.FileName : null;
            }
        }

        void OpenSettings()
        {
            using (var f = new SettingsForm(cfg))
            {
                if (f.ShowDialog(this) != DialogResult.OK) return;
                bool dossierChanged = f.Dossier != cfg.DossierPath;
                bool portChanged = f.Port != cfg.Port;
                cfg.DossierPath = f.Dossier; cfg.DefaultClient = f.Client; cfg.Port = f.Port; cfg.AutoUpdate = f.AutoUpdate; cfg.WriteReference = f.WriteReference;
                cfg.Save();
                RefreshDossierLabel();
                LogLine("Réglages enregistrés.");
                if (portChanged) { server.Stop(); StartServer(); }
                if (dossierChanged) ThreadPool.QueueUserWorkItem(delegate { sage.Disconnect(15000); TryConnect(true); });
            }
        }

        // ── Sage ──

        void TryConnect(bool verbose)
        {
            Ui(UpdateSageIndicator);
            if (!SageWorker.SageWindowOpen())
            {
                SetStatus("En attente : ouvrez Sage 50 pour pouvoir créer des factures", Orange);
                if (verbose) Ui(delegate { MessageBox.Show(this, "Ouvrez d'abord Sage 50" + (cfg.DossierName.Length > 0 ? " sur le dossier « " + cfg.DossierName + " »" : "") + ".", "SageBridge", MessageBoxButtons.OK, MessageBoxIcon.Information); });
                return;
            }
            try
            {
                string name = sage.Connect();
                SetStatus("Connecté au dossier « " + name + " » — prêt à recevoir des factures", Green);
            }
            catch (Exception e)
            {
                SetStatus("Sage non connecté : " + e.Message, Red);
                if (verbose) Ui(delegate { MessageBox.Show(this, e.Message, "SageBridge", MessageBoxButtons.OK, MessageBoxIcon.Warning); });
            }
        }

        InvoiceResult HandleInvoice(InvoiceRequest req)
        {
            if (updating) return new InvoiceResult { Error = "SageBridge installe une mise à jour : réessayez dans une minute." };
            Interlocked.Increment(ref busy);
            try { return HandleInvoiceInner(req); }
            finally { Interlocked.Decrement(ref busy); }
        }

        InvoiceResult HandleInvoiceInner(InvoiceRequest req)
        {
            if (string.IsNullOrEmpty(req.Client)) req.Client = cfg.DefaultClient;
            if (!SageWorker.SageWindowOpen())
            {
                SageWorker.KillOrphans(LogLine);
                var ko = new InvoiceResult { Error = "Sage 50 n'est pas ouvert sur ce PC : ouvrez Sage" + (cfg.DossierName.Length > 0 ? " (dossier " + cfg.DossierName + ")" : "") + " puis réessayez." };
                AddResultRow(req, ko);
                LogLine("Refusé : Sage 50 n'est pas ouvert.");
                SetStatus("Facture refusée : Sage 50 n'est pas ouvert", Red);
                return ko;
            }
            if (req.Customer != null && !req.Customer.IsEmpty)
                LogLine("Client demandé : " + req.Customer.DisplayName + ", " + req.Customer.Address1 + ", " + req.Customer.PostalCode + " " + req.Customer.City + " (" + req.Customer.Country + ")");

            if (!req.Simulation && !req.Force && !string.IsNullOrEmpty(req.OrderNo))
            {
                string existing = history.Existing(req.Site, req.OrderNo);
                if (existing != null)
                {
                    var dup = new InvoiceResult { Duplicate = true, Numero = existing, Error = "Commande déjà envoyée : facture " + existing };
                    AddResultRow(req, dup);
                    LogLine("Doublon refusé : " + req.Site + " " + req.OrderNo + " (déjà " + existing + ")");
                    return dup;
                }
            }
            LogLine((req.Simulation ? "Simulation : " : "Création : ") + req.Site + " " + req.OrderNo + " — " + req.Lines.Count + " ligne(s)");
            // Le callback est exécuté sur le thread worker avant même le retour de Create :
            // l'historique et le journal reflètent toujours ce que Sage a fait, y compris
            // quand l'HTTP a déjà renvoyé une erreur de timeout entre-temps.
            return sage.Create(req, delegate(InvoiceResult r) { RecordInvoiceOutcome(req, r); });
        }

        void RecordInvoiceOutcome(InvoiceRequest req, InvoiceResult res)
        {
            if (res.Ok && res.Simulated)
            {
                LogLine("Simulation terminée : client " + (string.IsNullOrEmpty(res.ClientCode) ? "à créer" : res.ClientCode) + ", mode TVA " + res.ClientModeTVA + ".");
                return;
            }
            if (res.Ok)
            {
                if (!string.IsNullOrEmpty(req.OrderNo)) history.Add(req.Site, req.OrderNo, res.Numero, res.ClientCode, res.NbLines, res.TotalHT);
                LogLine("Facture " + res.Numero + " créée (provisoire) — client " + res.ClientCode + (res.ClientCreated ? " (nouveau)" : "") + " — total HT " + res.TotalHT.ToString("0.00") + " €");
                SetStatus("Dernière facture : " + res.Numero + " à " + DateTime.Now.ToString("HH:mm") + " (provisoire, à valider dans Sage)", Green);
                Ui(delegate { try { tray.ShowBalloonTip(4000, "Facture créée dans Sage", res.Numero + " — " + res.ClientName + " (provisoire, à valider)", ToolTipIcon.Info); } catch { } });
            }
            else
            {
                LogLine("ÉCHEC : " + res.Error);
                SetStatus("Échec de la dernière facture : " + res.Error, Red);
            }
            AddResultRow(req, res);
        }

        void AddHistoryRow(string[] p)
        {
            var it = new ListViewItem(p[0]);
            it.SubItems.Add(p[1]); it.SubItems.Add(p[2]); it.SubItems.Add(p[3]); it.SubItems.Add(p[4]); it.SubItems.Add(p[5]);
            double ht; it.SubItems.Add(double.TryParse(p[6], NumberStyles.Any, CultureInfo.InvariantCulture, out ht) ? ht.ToString("0.00") + " €" : p[6]);
            it.SubItems.Add("Créée");
            it.ForeColor = Color.FromArgb(40, 40, 40);
            list.Items.Add(it);
        }

        void AddResultRow(InvoiceRequest req, InvoiceResult res)
        {
            Ui(delegate
            {
                var it = new ListViewItem(DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"));
                it.SubItems.Add(req.Site); it.SubItems.Add(req.OrderNo); it.SubItems.Add(res.Numero);
                it.SubItems.Add((string.IsNullOrEmpty(res.ClientCode) ? req.Client : res.ClientCode) + (res.ClientCreated ? " (nouveau)" : ""));
                it.SubItems.Add((res.NbLines > 0 ? res.NbLines : req.Lines.Count).ToString());
                it.SubItems.Add(res.Ok ? res.TotalHT.ToString("0.00") + " €" : "");
                it.SubItems.Add(res.Ok ? "Créée (provisoire)" : (res.Duplicate ? "Déjà envoyée" : "Échec : " + res.Error));
                it.ForeColor = res.Ok ? Green : (res.Duplicate ? Orange : Red);
                list.Items.Insert(0, it);
            });
        }

        void CreateTestInvoice()
        {
            using (var f = new TestInvoiceForm(cfg))
            {
                if (f.ShowDialog(this) != DialogResult.OK) return;
                cfg.TestArticle = f.Article; cfg.Save();
                var req = new InvoiceRequest { Site = "test", OrderNo = "TEST-" + DateTime.Now.ToString("yyyyMMdd-HHmmss"), Client = f.Client, Reference = "Test SageBridge", Remarque = "Facture de test SageBridge — à supprimer", Force = true };
                req.Lines.Add(new InvoiceLine { Code = f.Article, Qty = 1, Puht = f.Price, HasPuht = true, Tva = 20, Description = "Ligne de test SageBridge" });
                btnTest.Enabled = false;
                ThreadPool.QueueUserWorkItem(delegate
                {
                    InvoiceResult r = HandleInvoice(req);
                    Ui(delegate
                    {
                        btnTest.Enabled = true;
                        if (r.Ok) MessageBox.Show(this, "Facture " + r.Numero + " créée (provisoire) pour le client " + r.ClientCode + ".\nTotal HT " + r.TotalHT.ToString("0.00") + " € — TTC " + r.TotalTTC.ToString("0.00") + " €\n\nPensez à la supprimer dans Sage.", "Test réussi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        else MessageBox.Show(this, r.Error + (r.Messages.Count > 0 ? "\n\n" + string.Join("\n", r.Messages.ToArray()) : ""), "Test échoué", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    });
                });
            }
        }

        // Rapport de paiement (export PFS .xlsx) : rapproche chaque ligne d'une facture Sage, sans rien modifier dans Sage.
        void OpenPaymentReport()
        {
            if (!SageWorker.SageWindowOpen())
            {
                MessageBox.Show(this, "Ouvrez d'abord Sage 50" + (cfg.DossierName.Length > 0 ? " sur le dossier « " + cfg.DossierName + " »" : "") + ".", "SageBridge", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            string file;
            using (var dlg = new OpenFileDialog { Title = "Choisir le rapport de paiement (export Excel Paris Fashion Shop)", Filter = "Classeur Excel (*.xlsx)|*.xlsx" })
            {
                try { dlg.InitialDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads"); } catch { }
                if (dlg.ShowDialog(this) != DialogResult.OK) return;
                file = dlg.FileName;
            }
            Interlocked.Increment(ref busy);
            try
            {
                using (var f = new ReconcileForm(file, sage, history.FindByOrder, LogLine)) f.ShowDialog(this);
            }
            finally { Interlocked.Decrement(ref busy); }
        }

        // Enregistrement des factures/avoirs validés d'un mois en PDF (modèle choisi) dans Recap\année\mois\Facture|Avoir.
        void OpenPdfExport()
        {
            if (!SageWorker.SageWindowOpen())
            {
                MessageBox.Show(this, "Ouvrez d'abord Sage 50" + (cfg.DossierName.Length > 0 ? " sur le dossier « " + cfg.DossierName + " »" : "") + ".", "SageBridge", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            Interlocked.Increment(ref busy);
            try
            {
                using (var f = new PdfExportForm(cfg, sage, LogLine)) f.ShowDialog(this);
            }
            finally { Interlocked.Decrement(ref busy); }
        }

        void OpenSage()
        {
            string exe = SageWorker.SageExePath();
            if (string.IsNullOrEmpty(exe) || !File.Exists(exe)) exe = @"C:\Program Files (x86)\Ciel\Gestion Commerciale\WGC.exe";
            SageWorker.KillOrphans(LogLine);
            try { Process.Start(exe); } catch (Exception ex) { MessageBox.Show(this, "Impossible de lancer Sage : " + ex.Message, "SageBridge"); }
        }

        static bool IsStartup()
        {
            try { using (RegistryKey k = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Run")) return k != null && k.GetValue(Program.AppName) != null; }
            catch { return false; }
        }

        static void SetStartup(bool on)
        {
            try
            {
                using (RegistryKey k = Registry.CurrentUser.CreateSubKey(@"Software\Microsoft\Windows\CurrentVersion\Run"))
                {
                    if (on) k.SetValue(Program.AppName, "\"" + Application.ExecutablePath + "\"");
                    else k.DeleteValue(Program.AppName, false);
                }
            }
            catch (Exception e) { MessageBox.Show("Réglage impossible : " + e.Message, "SageBridge"); }
        }

        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);
            if (WindowState == FormWindowState.Minimized && tray != null)
            {
                Hide();
                try { tray.ShowBalloonTip(2500, "SageBridge reste actif", "Double-cliquez sur l'icône pour rouvrir la fenêtre.", ToolTipIcon.None); } catch { }
            }
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (!reallyClose && e.CloseReason == CloseReason.UserClosing)
            {
                DialogResult r = MessageBox.Show(this, "Quitter SageBridge ?\n\nPlus aucune facture ne pourra être créée dans Sage tant qu'il ne sera pas relancé.\n\n« Non » le garde actif dans la zone de notification.",
                    "SageBridge", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                if (r == DialogResult.Cancel) { e.Cancel = true; return; }
                if (r == DialogResult.No) { e.Cancel = true; WindowState = FormWindowState.Minimized; return; }
            }
            try { sageTimer.Stop(); } catch { }
            try { if (updateTimer != null) updateTimer.Stop(); } catch { }
            try { server.Stop(); } catch { }
            tray.Visible = false;
            LogLine("SageBridge arrêté.");
            base.OnFormClosing(e);
            // Ne JAMAIS relâcher explicitement les objets COM de Sage en quittant : libérer le dossier
            // obtenu par OpenFile peut figer le Sage ouvert par l'utilisateur (constaté le 20/10/2026).
            // Le processus est arrêté net, comme un « taskkill » (sans effet sur Sage, vérifié depuis la 1.0).
            try { Process.GetCurrentProcess().Kill(); } catch { Environment.Exit(0); }
        }
    }

    // ───────────────────────── Petites fenêtres ─────────────────────────

    // ───────────────────────── Excel (.xlsx) minimal : lecture + ajout de colonnes ─────────────────────────
    // Un .xlsx est un zip de fichiers XML. On lit la 1ʳᵉ feuille et on réécrit le même classeur
    // en ajoutant des colonnes à droite (le reste du fichier est conservé tel quel).

    class XlsxRow
    {
        public int Index;                                            // n° de ligne Excel (attribut r)
        public Dictionary<int, string> Cells = new Dictionary<int, string>(); // colonne (0 = A) → texte
        public string Get(int col) { string v; return col >= 0 && Cells.TryGetValue(col, out v) ? v : ""; }
    }

    class XlsxSheet
    {
        public string Entry = "";     // chemin de la feuille dans le zip
        public List<XlsxRow> Rows = new List<XlsxRow>();
        public int MaxCol = -1;
    }

    static class Xlsx
    {
        public static readonly System.Xml.Linq.XNamespace M = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        static readonly System.Xml.Linq.XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        static readonly System.Xml.Linq.XNamespace PR = "http://schemas.openxmlformats.org/package/2006/relationships";

        public static string ColName(int col)
        {
            string s = ""; col++;
            while (col > 0) { int m = (col - 1) % 26; s = (char)('A' + m) + s; col = (col - 1) / 26; }
            return s;
        }

        static int ColIndex(string cellRef)
        {
            int c = 0;
            foreach (char ch in cellRef) { if (ch < 'A' || ch > 'Z') break; c = c * 26 + (ch - 'A' + 1); }
            return c - 1;
        }

        static System.Xml.Linq.XDocument Load(System.IO.Compression.ZipArchive z, string name)
        {
            System.IO.Compression.ZipArchiveEntry e = Find(z, name);
            if (e == null) return null;
            using (Stream s = e.Open()) return System.Xml.Linq.XDocument.Load(s);
        }

        static System.IO.Compression.ZipArchiveEntry Find(System.IO.Compression.ZipArchive z, string name)
        {
            foreach (System.IO.Compression.ZipArchiveEntry e in z.Entries)
                if (string.Equals(e.FullName, name, StringComparison.OrdinalIgnoreCase)) return e;
            return null;
        }

        static string FirstSheetEntry(System.IO.Compression.ZipArchive z)
        {
            try
            {
                var wb = Load(z, "xl/workbook.xml");
                var rels = Load(z, "xl/_rels/workbook.xml.rels");
                var sheet = wb.Root.Element(M + "sheets").Element(M + "sheet");
                string rid = (string)sheet.Attribute(R + "id");
                foreach (var rel in rels.Root.Elements(PR + "Relationship"))
                {
                    if ((string)rel.Attribute("Id") != rid) continue;
                    string target = ((string)rel.Attribute("Target") ?? "").Replace('\\', '/');
                    return target.StartsWith("/") ? target.Substring(1) : "xl/" + target;
                }
            }
            catch { }
            return "xl/worksheets/sheet1.xml";
        }

        public static XlsxSheet Read(string path)
        {
            var sh = new XlsxSheet();
            using (var z = System.IO.Compression.ZipFile.OpenRead(path))
            {
                var shared = new List<string>();
                var sst = Load(z, "xl/sharedStrings.xml");
                if (sst != null)
                    foreach (var si in sst.Root.Elements(M + "si"))
                    {
                        var sb = new StringBuilder();
                        foreach (var t in si.Descendants(M + "t")) sb.Append(t.Value);
                        shared.Add(sb.ToString());
                    }
                sh.Entry = FirstSheetEntry(z);
                var doc = Load(z, sh.Entry);
                if (doc == null) throw new Exception("Feuille Excel introuvable dans le fichier.");
                int auto = 0;
                foreach (var row in doc.Root.Element(M + "sheetData").Elements(M + "row"))
                {
                    var xr = new XlsxRow();
                    int ri; xr.Index = int.TryParse((string)row.Attribute("r"), out ri) ? ri : auto + 1;
                    auto = xr.Index;
                    int nextCol = 0;
                    foreach (var c in row.Elements(M + "c"))
                    {
                        string r = (string)c.Attribute("r");
                        int col = string.IsNullOrEmpty(r) ? nextCol : ColIndex(r);
                        nextCol = col + 1;
                        string type = (string)c.Attribute("t") ?? "";
                        string val;
                        if (type == "inlineStr")
                        {
                            var sb = new StringBuilder();
                            foreach (var t in c.Descendants(M + "t")) sb.Append(t.Value);
                            val = sb.ToString();
                        }
                        else
                        {
                            var v = c.Element(M + "v");
                            val = v == null ? "" : v.Value;
                            int idx;
                            if (type == "s" && int.TryParse(val, out idx) && idx >= 0 && idx < shared.Count) val = shared[idx];
                        }
                        xr.Cells[col] = val;
                        if (col > sh.MaxCol) sh.MaxCol = col;
                    }
                    sh.Rows.Add(xr);
                }
            }
            return sh;
        }

        // Ajoute styles de remplissage (vert, orange, rouge, gris) ; renvoie les index de style ou null.
        static int[] AddFillStyles(System.IO.Compression.ZipArchive z)
        {
            try
            {
                System.IO.Compression.ZipArchiveEntry e = Find(z, "xl/styles.xml");
                if (e == null) return null;
                System.Xml.Linq.XDocument st;
                using (Stream s = e.Open()) st = System.Xml.Linq.XDocument.Load(s);
                var fills = st.Root.Element(M + "fills");
                var xfs = st.Root.Element(M + "cellXfs");
                if (fills == null || xfs == null) return null;
                string[] colors = { "FFC6EFCE", "FFFFEB9C", "FFFFC7CE", "FFE7E6E6" };
                var result = new int[colors.Length];
                for (int i = 0; i < colors.Length; i++)
                {
                    int fillId = fills.Elements(M + "fill").Count();
                    fills.Add(new System.Xml.Linq.XElement(M + "fill",
                        new System.Xml.Linq.XElement(M + "patternFill", new System.Xml.Linq.XAttribute("patternType", "solid"),
                            new System.Xml.Linq.XElement(M + "fgColor", new System.Xml.Linq.XAttribute("rgb", colors[i])),
                            new System.Xml.Linq.XElement(M + "bgColor", new System.Xml.Linq.XAttribute("indexed", "64")))));
                    result[i] = xfs.Elements(M + "xf").Count();
                    xfs.Add(new System.Xml.Linq.XElement(M + "xf",
                        new System.Xml.Linq.XAttribute("numFmtId", "0"), new System.Xml.Linq.XAttribute("fontId", "0"),
                        new System.Xml.Linq.XAttribute("fillId", fillId), new System.Xml.Linq.XAttribute("borderId", "0"),
                        new System.Xml.Linq.XAttribute("xfId", "0"), new System.Xml.Linq.XAttribute("applyFill", "1")));
                }
                fills.SetAttributeValue("count", fills.Elements(M + "fill").Count());
                xfs.SetAttributeValue("count", xfs.Elements(M + "xf").Count());
                Save(z, e, st);
                return result;
            }
            catch { return null; }
        }

        static void Save(System.IO.Compression.ZipArchive z, System.IO.Compression.ZipArchiveEntry e, System.Xml.Linq.XDocument d)
        {
            string name = e.FullName;
            e.Delete();
            var ne = z.CreateEntry(name);
            using (Stream s = ne.Open())
            using (var w = System.Xml.XmlWriter.Create(s, new System.Xml.XmlWriterSettings { Encoding = new UTF8Encoding(false) }))
                d.Save(w);
        }

        // Copie src vers dst en ajoutant des colonnes après la dernière colonne utilisée.
        // values : n° de ligne Excel → valeurs (string = texte, double = nombre, null = vide).
        // fillOf : n° de ligne Excel → couleur de la 1ʳᵉ colonne ajoutée (0 vert, 1 orange, 2 rouge, 3 gris, -1 aucune).
        public static void WriteWithColumns(string src, string dst, XlsxSheet sheet, int headerRow, string[] headers,
            Dictionary<int, object[]> values, Dictionary<int, int> fillOf, int[] widths)
        {
            File.Copy(src, dst, true);
            using (var z = System.IO.Compression.ZipFile.Open(dst, System.IO.Compression.ZipArchiveMode.Update))
            {
                int[] fills = AddFillStyles(z);
                System.IO.Compression.ZipArchiveEntry e = Find(z, sheet.Entry);
                System.Xml.Linq.XDocument doc;
                using (Stream s = e.Open()) doc = System.Xml.Linq.XDocument.Load(s);
                var root = doc.Root;
                var data = root.Element(M + "sheetData");
                int first = sheet.MaxCol + 1, lastRow = 0;
                string headerStyle = null;
                foreach (var row in data.Elements(M + "row"))
                {
                    int ri; if (!int.TryParse((string)row.Attribute("r"), out ri)) continue;
                    if (ri == headerRow)
                    {
                        var hc = row.Elements(M + "c").FirstOrDefault();
                        if (hc != null) headerStyle = (string)hc.Attribute("s");
                    }
                }
                foreach (var row in data.Elements(M + "row"))
                {
                    int ri; if (!int.TryParse((string)row.Attribute("r"), out ri)) continue;
                    if (ri > lastRow) lastRow = ri;
                    object[] vals = null;
                    string style = null;
                    if (ri == headerRow) { vals = headers; style = headerStyle; }
                    else if (values.TryGetValue(ri, out vals)) { }
                    if (vals == null) continue;
                    var spans = row.Attribute("spans"); if (spans != null) spans.Remove();
                    int fill;
                    for (int i = 0; i < vals.Length; i++)
                    {
                        object v = vals[i];
                        if (v == null) continue;
                        var c = new System.Xml.Linq.XElement(M + "c", new System.Xml.Linq.XAttribute("r", ColName(first + i) + ri));
                        string s = style;
                        if (i == 0 && ri != headerRow && fills != null && fillOf != null && fillOf.TryGetValue(ri, out fill) && fill >= 0 && fill < fills.Length)
                            s = fills[fill].ToString();
                        if (s != null) c.Add(new System.Xml.Linq.XAttribute("s", s));
                        if (v is double)
                            c.Add(new System.Xml.Linq.XElement(M + "v", ((double)v).ToString("0.##", CultureInfo.InvariantCulture)));
                        else
                        {
                            c.Add(new System.Xml.Linq.XAttribute("t", "inlineStr"));
                            c.Add(new System.Xml.Linq.XElement(M + "is", new System.Xml.Linq.XElement(M + "t",
                                new System.Xml.Linq.XAttribute(System.Xml.Linq.XNamespace.Xml + "space", "preserve"), Convert.ToString(v, CultureInfo.InvariantCulture))));
                        }
                        row.Add(c);
                    }
                }
                var dim = root.Element(M + "dimension");
                if (dim != null) dim.SetAttributeValue("ref", "A1:" + ColName(first + headers.Length - 1) + Math.Max(lastRow, 1));
                var cols = root.Element(M + "cols");
                if (cols == null) { cols = new System.Xml.Linq.XElement(M + "cols"); data.AddBeforeSelf(cols); }
                for (int i = 0; i < headers.Length; i++)
                {
                    int w = widths != null && i < widths.Length ? widths[i] : 20;
                    cols.Add(new System.Xml.Linq.XElement(M + "col",
                        new System.Xml.Linq.XAttribute("min", first + i + 1), new System.Xml.Linq.XAttribute("max", first + i + 1),
                        new System.Xml.Linq.XAttribute("width", w), new System.Xml.Linq.XAttribute("customWidth", "1")));
                }
                Save(z, e, doc);
            }
        }
    }

    // ───────────────────────── Rapprochement d'un rapport de paiement (PFS…) ─────────────────────────
    // Lecture seule côté Sage : retrouve pour chaque ligne du rapport la facture correspondante.

    class ReportLine
    {
        public int Row;                 // n° de ligne Excel
        public string Order = "", Client = "";
        public double HtValide, HtCommande;
        public bool HasHtCommande;
        // Résultat
        public string Resultat = "", Numero = "", ClientSage = "", Remarque = "";
        public string Reglement = "";   // après « Marquer en soldé » : n° du règlement ou message d'erreur
        public ListViewItem Item;
        public double HtFacture = double.NaN;
        public int Fill = -1;           // 0 vert, 1 orange, 2 rouge, 3 gris
    }

    class PaymentReconciler
    {
        public const string TROUVEE = "Trouvée", DEJA = "Déjà soldée", VERIF = "À vérifier", INTROUVABLE = "Introuvable";
        const double TOL = 0.011;

        readonly List<PieceRow> open = new List<PieceRow>();      // factures validées non soldées
        readonly List<PieceRow> provis = new List<PieceRow>();    // factures provisoires (ignorées)
        readonly List<PieceRow> avoirs = new List<PieceRow>();    // avoirs non soldés
        readonly List<PieceRow> soldees = new List<PieceRow>();   // factures déjà soldées
        readonly HashSet<string> used = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        readonly Func<string, string> historyLookup;

        public PaymentReconciler(List<PieceRow> pieces, Func<string, string> history)
        {
            historyLookup = history;
            foreach (PieceRow p in pieces)
            {
                string k = p.Kind;
                if (k == "facture")
                {
                    if (p.Soldee) soldees.Add(p);
                    else if (p.Valide) open.Add(p);
                    else provis.Add(p);
                }
                else if (k == "avoir" && !p.Soldee) avoirs.Add(p);
            }
            // Les plus anciennes d'abord (attribution des doublons dans l'ordre des numéros).
            Comparison<PieceRow> byNum = delegate(PieceRow a, PieceRow b) { return string.CompareOrdinal(a.Num, b.Num); };
            open.Sort(byNum); avoirs.Sort(byNum);
        }

        public int OpenCount { get { return open.Count; } }

        static readonly HashSet<string> LegalWords = new HashSet<string> {
            "sarl", "sas", "sasu", "eurl", "eirl", "snc", "sa", "sci", "ei", "ste", "societe", "ets", "inc", "ltd", "llc", "gmbh", "srl", "sl", "bv", "sprl", "srls", "spa", "the", "et", "and", "co"
        };

        static string Core(string s)
        {
            var parts = new List<string>();
            foreach (string w in Norm.Text(s).Split(' ')) if (w.Length > 0 && !LegalWords.Contains(w)) parts.Add(w);
            return string.Join(" ", parts.ToArray());
        }

        // 3 = nom identique, 2 = identique hors forme juridique, 1 = l'un contient l'autre, 0 = différent.
        static int NameLevel(string report, PieceRow p)
        {
            int best = 0;
            foreach (string n in new[] { p.Societe, p.CliName })
            {
                if (string.IsNullOrEmpty(n)) continue;
                string a = Norm.Text(report), b = Norm.Text(n);
                if (a.Length == 0 || b.Length == 0) continue;
                if (a == b) return 3;
                string ca = Core(report), cb = Core(n);
                if (ca.Length > 0 && ca == cb) { best = Math.Max(best, 2); continue; }
                string shorter = ca.Length <= cb.Length ? ca : cb, longer = ca.Length <= cb.Length ? cb : ca;
                if (shorter.Length >= 4 && (" " + longer + " ").Contains(" " + shorter + " ")) best = Math.Max(best, 1);
            }
            return best;
        }

        static bool Same(double a, double b) { return Math.Abs(a - b) <= TOL; }
        static string M(double v) { return v.ToString("#,##0.00", CultureInfo.GetCultureInfo("fr-FR")); }
        static string Desc(PieceRow p) { return p.Num + " (" + M(p.TotalHT) + " HT)"; }
        static string ClientOf(PieceRow p) { return (p.Client.Length > 0 ? p.Client + " – " : "") + p.NomAffiche; }

        void Assign(ReportLine l, PieceRow p, string resultat, int fill)
        {
            l.Resultat = resultat; l.Numero = p.Num; l.ClientSage = ClientOf(p); l.HtFacture = p.TotalHT; l.Fill = fill;
            if (resultat == TROUVEE) used.Add(p.Num);
        }

        void Note(ReportLine l, string s) { l.Remarque = l.Remarque.Length == 0 ? s : l.Remarque + " • " + s; }

        public void Match(ReportLine l)
        {
            // 1) Historique SageBridge : n° de commande → n° de facture (rapprochement exact).
            string hist = historyLookup == null || l.Order.Length == 0 ? null : historyLookup(l.Order);
            if (!string.IsNullOrEmpty(hist))
            {
                PieceRow p = open.Find(delegate(PieceRow x) { return string.Equals(x.Num, hist, StringComparison.OrdinalIgnoreCase); });
                if (p != null && !used.Contains(p.Num))
                {
                    if (Same(p.TotalHT, l.HtValide)) { Assign(l, p, TROUVEE, 0); Note(l, "par le n° de commande (historique SageBridge)"); return; }
                    Assign(l, p, VERIF, 1);
                    Note(l, "facture de cette commande (historique SageBridge) mais HT " + M(p.TotalHT) + " ≠ validé " + M(l.HtValide) + " (écart " + M(p.TotalHT - l.HtValide) + ") : avoir à faire ?");
                    AddAvoirHints(l, NamedAvoirs(l.Client), p);
                    return;
                }
                PieceRow s = soldees.Find(delegate(PieceRow x) { return string.Equals(x.Num, hist, StringComparison.OrdinalIgnoreCase); });
                if (s != null) { Assign(l, s, DEJA, 3); Note(l, "facture de cette commande déjà soldée dans Sage"); return; }
                PieceRow pv = provis.Find(delegate(PieceRow x) { return string.Equals(x.Num, hist, StringComparison.OrdinalIgnoreCase); });
                if (pv != null) Note(l, "facture " + Desc(pv) + " de cette commande encore provisoire (ignorée)");
            }

            // 2) Nom + montant HT validé.
            var named = new List<PieceRow>();
            int bestLevel = 0;
            foreach (PieceRow p in open)
            {
                int lv = NameLevel(l.Client, p);
                if (lv == 0) continue;
                if (lv > bestLevel) { bestLevel = lv; named.Clear(); }
                if (lv == bestLevel) named.Add(p);
            }
            var exact = named.FindAll(delegate(PieceRow p) { return !used.Contains(p.Num) && Same(p.TotalHT, l.HtValide); });
            if (exact.Count >= 1)
            {
                Assign(l, exact[0], TROUVEE, 0);
                if (bestLevel == 1) Note(l, "nom approchant : « " + exact[0].NomAffiche + " »");
                if (exact.Count > 1)
                {
                    var others = new List<string>();
                    for (int i = 1; i < exact.Count; i++) others.Add(exact[i].Num);
                    Note(l, exact.Count + " factures identiques (même nom, même HT) : la plus ancienne est retenue, autres : " + string.Join(", ", others.ToArray()));
                }
                return;
            }

            // 3) Déjà soldée ?
            PieceRow done = soldees.Find(delegate(PieceRow p) { return Same(p.TotalHT, l.HtValide) && NameLevel(l.Client, p) > 0; });
            if (done != null) { Assign(l, done, DEJA, 3); Note(l, "déjà soldée dans Sage : rien à faire"); return; }

            // 4) Introuvable : pistes.
            l.Resultat = INTROUVABLE; l.Fill = 2;
            var free = named.FindAll(delegate(PieceRow p) { return !used.Contains(p.Num); });
            if (free.Count > 0)
            {
                l.ClientSage = ClientOf(free[0]);
                var lst = new List<string>();
                foreach (PieceRow p in free) lst.Add(p.Num + " (" + M(p.TotalHT) + " HT, écart " + M(p.TotalHT - l.HtValide) + ")");
                Note(l, "autre(s) facture(s) non soldée(s) à ce nom, montant différent : " + string.Join(", ", lst.ToArray()));
                if (l.HasHtCommande && !Same(l.HtCommande, l.HtValide))
                {
                    PieceRow pc = free.Find(delegate(PieceRow p) { return Same(p.TotalHT, l.HtCommande); });
                    if (pc != null)
                    {
                        l.Resultat = VERIF; l.Fill = 1; l.Numero = pc.Num; l.HtFacture = pc.TotalHT;
                        Note(l, pc.Num + " correspond au montant COMMANDÉ (" + M(l.HtCommande) + ") mais PFS n'a validé que " + M(l.HtValide) + " : avoir de " + M(l.HtCommande - l.HtValide) + " HT à faire ?");
                    }
                }
            }
            AddAvoirHints(l, NamedAvoirs(l.Client), null);
            if (free.Count == 0)
            {
                PieceRow pv = provis.Find(delegate(PieceRow p) { return NameLevel(l.Client, p) > 0 && Same(p.TotalHT, l.HtValide); });
                if (pv != null) Note(l, "facture provisoire " + Desc(pv) + " à ce nom et ce montant (ignorée : à valider dans Sage)");
                var sameAmount = open.FindAll(delegate(PieceRow p) { return !used.Contains(p.Num) && Same(p.TotalHT, l.HtValide); });
                if (sameAmount.Count > 0 && sameAmount.Count <= 5)
                {
                    var lst = new List<string>();
                    foreach (PieceRow p in sameAmount) lst.Add(p.Num + " « " + p.NomAffiche + " »");
                    Note(l, "aucune facture non soldée à ce nom ; même montant HT sous un autre nom : " + string.Join(", ", lst.ToArray()));
                    if (l.Resultat == INTROUVABLE) { l.Resultat = VERIF; l.Fill = 1; }
                }
                else if (l.Remarque.Length == 0) Note(l, "aucune facture non soldée à ce nom");
            }
        }

        List<PieceRow> NamedAvoirs(string client)
        {
            return avoirs.FindAll(delegate(PieceRow a) { return NameLevel(client, a) > 0; });
        }

        void AddAvoirHints(ReportLine l, List<PieceRow> avs, PieceRow onlyInvoice)
        {
            if (avs.Count == 0) return;
            var lst = new List<string>();
            foreach (PieceRow a in avs) lst.Add(Desc(a));
            Note(l, "avoir(s) non soldé(s) à ce nom : " + string.Join(", ", lst.ToArray()));
            // Facture − avoir = montant validé ?
            foreach (PieceRow f in open)
            {
                if (used.Contains(f.Num)) continue;
                if (onlyInvoice != null && f != onlyInvoice) continue;
                if (NameLevel(l.Client, f) == 0) continue;
                foreach (PieceRow a in avs)
                {
                    if (Same(f.TotalHT - Math.Abs(a.TotalHT), l.HtValide))
                    {
                        l.Resultat = VERIF; l.Fill = 1; l.Numero = f.Num; l.HtFacture = f.TotalHT; l.ClientSage = ClientOf(f);
                        Note(l, f.Num + " − avoir " + a.Num + " = montant validé " + M(l.HtValide) + " : facture à solder avec l'avoir");
                        return;
                    }
                }
            }
        }
    }

    class ReconcileForm : Form
    {
        readonly string path;
        readonly SageWorker sage;
        readonly Func<string, string> history;
        readonly Action<string> log;
        readonly ProgressBar bar;
        readonly Label lblStep, lblSummary;
        readonly ListView list;
        readonly Button btnOpen, btnClose, btnSolder;
        string outPath;
        volatile bool closing;
        XlsxSheet sheet;
        int headerRowIndex;
        List<ReportLine> lines;
        bool solderDone;
        List<PieceRow> pieces0;
        List<string[]> modes;

        public ReconcileForm(string file, SageWorker worker, Func<string, string> hist, Action<string> logger)
        {
            path = file; sage = worker; history = hist; log = logger;
            Text = "Rapprochement des paiements — " + Path.GetFileName(file);
            Font = new Font("Segoe UI", 9.5f);
            Icon = Program.AppIcon;
            StartPosition = FormStartPosition.CenterParent;
            ShowInTaskbar = false;
            MinimizeBox = false;
            Size = new Size(Ui2.Px(1150), Ui2.Px(640));

            var top = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 1, Padding = new Padding(Ui2.Px(10)) };
            lblStep = new Label { AutoSize = true, Text = "Lecture du fichier…", Margin = new Padding(3, 3, 3, 6) };
            bar = new ProgressBar { Dock = DockStyle.Top, Height = Ui2.Px(20), Minimum = 0, Maximum = 100 };
            top.Controls.Add(lblStep); top.Controls.Add(bar);

            list = new ListView { Dock = DockStyle.Fill, View = View.Details, FullRowSelect = true, GridLines = true, HideSelection = false };
            string[] cols = { "Ligne", "Commande", "Client (rapport)", "HT validé", "Résultat", "Facture Sage", "Client Sage", "Règlement Sage", "Remarque" };
            int[] widths = { 50, 115, 200, 80, 95, 100, 170, 260, 480 };
            for (int i = 0; i < cols.Length; i++) list.Columns.Add(cols[i], Ui2.Px(widths[i]));

            var bottom = new FlowLayoutPanel { Dock = DockStyle.Bottom, AutoSize = true, FlowDirection = FlowDirection.RightToLeft, Padding = new Padding(Ui2.Px(8)) };
            btnClose = new Button { Text = "Fermer", AutoSize = true };
            btnClose.Click += delegate { Close(); };
            btnOpen = new Button { Text = "Ouvrir le fichier Excel", AutoSize = true, Enabled = false };
            btnOpen.Click += delegate { try { Process.Start(outPath); } catch (Exception e) { MessageBox.Show(this, e.Message, "SageBridge"); } };
            btnSolder = new Button { Text = "Marquer les factures trouvées en soldé", AutoSize = true, Enabled = false, Font = new Font("Segoe UI Semibold", 9.5f) };
            btnSolder.Click += delegate { AskSolder(); };
            lblSummary = new Label { AutoSize = true, Margin = new Padding(3, Ui2.Px(7), Ui2.Px(20), 3) };
            bottom.Controls.Add(btnClose); bottom.Controls.Add(btnOpen); bottom.Controls.Add(btnSolder); bottom.Controls.Add(lblSummary);

            Controls.Add(list); Controls.Add(top); Controls.Add(bottom);
            FormClosing += delegate(object s, FormClosingEventArgs e)
            {
                if (solderRunning && e.CloseReason == CloseReason.UserClosing)
                {
                    MessageBox.Show(this, "Les factures sont en train d'être soldées dans Sage : attendez la fin du traitement.", "SageBridge", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    e.Cancel = true; return;
                }
                closing = true;
            };
            Shown += delegate { var th = new Thread(Work) { IsBackground = true }; th.Start(); };
        }

        void Ui(Action a)
        {
            if (IsDisposed || closing) return;
            try { BeginInvoke(a); } catch { }
        }

        static double Num(string s)
        {
            s = (s ?? "").Replace("€", "").Replace(" ", "").Replace(" ", "").Trim();
            double d;
            if (double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out d)) return d;
            if (double.TryParse(s.Replace(',', '.'), NumberStyles.Float, CultureInfo.InvariantCulture, out d)) return d;
            return double.NaN;
        }

        void Step(string text, int value, int max)
        {
            Ui(delegate
            {
                lblStep.Text = text;
                bar.Maximum = Math.Max(1, max);
                bar.Value = Math.Max(0, Math.Min(bar.Maximum, value));
            });
        }

        void Work()
        {
            try { WorkInner(); }
            catch (Exception e)
            {
                log("Rapprochement : " + e.Message);
                Ui(delegate
                {
                    lblStep.Text = "Erreur : " + e.Message; lblStep.ForeColor = Color.Firebrick;
                    MessageBox.Show(this, e.Message, "Rapprochement impossible", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                });
            }
        }

        void WorkInner()
        {
            // 1) Fichier
            sheet = Xlsx.Read(path);
            int headerRow = -1, cOrder = -1, cClient = -1, cValide = -1, cCommande = -1;
            for (int r = 0; r < Math.Min(15, sheet.Rows.Count) && headerRow < 0; r++)
            {
                XlsxRow xr = sheet.Rows[r];
                int o = -1, c = -1, v = -1, cm = -1;
                foreach (KeyValuePair<int, string> kv in xr.Cells)
                {
                    string h = Norm.Text(kv.Value);
                    if (h.Contains("montant valide ht") || (h.Contains("valide") && h.Contains("ht") && h.Contains("montant"))) v = kv.Key;
                    else if (h.Contains("montant commande ht")) cm = kv.Key;
                    else if (h == "client" || (h.Contains("client") && !h.Contains("montant") && c < 0)) c = kv.Key;
                    else if (h.Contains("commande") && !h.Contains("montant") && o < 0) o = kv.Key;
                }
                if (c >= 0 && v >= 0) { headerRow = r; cOrder = o; cClient = c; cValide = v; cCommande = cm; }
            }
            if (headerRow < 0)
                throw new Exception("Colonnes « Client » et « Montant validé HT » introuvables dans la 1ʳᵉ feuille du fichier : est-ce bien un export de paiements Paris Fashion Shop ?");

            headerRowIndex = sheet.Rows[headerRow].Index;
            lines = new List<ReportLine>();
            for (int r = headerRow + 1; r < sheet.Rows.Count; r++)
            {
                XlsxRow xr = sheet.Rows[r];
                string client = xr.Get(cClient).Trim();
                double ht = Num(xr.Get(cValide));
                if (client.Length == 0 && double.IsNaN(ht)) continue;
                var l = new ReportLine { Row = xr.Index, Order = xr.Get(cOrder).Trim(), Client = client, HtValide = ht };
                double hc = cCommande >= 0 ? Num(xr.Get(cCommande)) : double.NaN;
                if (!double.IsNaN(hc)) { l.HtCommande = hc; l.HasHtCommande = true; }
                lines.Add(l);
            }
            if (lines.Count == 0) throw new Exception("Aucune ligne de paiement dans le fichier.");
            log("Rapprochement de « " + Path.GetFileName(path) + " » : " + lines.Count + " ligne(s).");

            // 2) Factures Sage (lecture seule)
            Step("Lecture des factures dans Sage…", 0, 1);
            List<PieceRow> pieces = sage.ListPieces(delegate(int done, int total)
            {
                Step("Lecture des factures dans Sage : " + done + " / " + total, done, total);
            });
            var rec = new PaymentReconciler(pieces, history);
            pieces0 = pieces;

            // 3) Ligne par ligne
            int ok = 0, verif = 0, ko = 0, deja = 0;
            for (int i = 0; i < lines.Count; i++)
            {
                if (closing) return;
                ReportLine l = lines[i];
                Step("Rapprochement ligne " + (i + 1) + " / " + lines.Count + " : " + l.Client, i, lines.Count);
                if (double.IsNaN(l.HtValide)) { l.Resultat = PaymentReconciler.INTROUVABLE; l.Fill = 2; l.Remarque = "montant validé HT illisible"; }
                else rec.Match(l);
                if (l.Resultat == PaymentReconciler.TROUVEE) ok++;
                else if (l.Resultat == PaymentReconciler.VERIF) verif++;
                else if (l.Resultat == PaymentReconciler.DEJA) deja++;
                else ko++;
                ReportLine cur = l;
                Ui(delegate { AddRow(cur); });
                Thread.Sleep(20);
            }

            // 4) Fichier résultat
            Step("Écriture du fichier Excel…", lines.Count, lines.Count);
            outPath = WriteExcel();
            string summary = lines.Count + " lignes : " + ok + " trouvée(s), " + verif + " à vérifier, " + ko + " introuvable(s)" + (deja > 0 ? ", " + deja + " déjà soldée(s)" : "")
                + " — " + rec.OpenCount + " facture(s) non soldée(s) dans Sage.";
            log("Rapprochement terminé : " + summary + " Fichier : " + outPath);
            if (ok > 0) { try { modes = sage.ListModes(); } catch (Exception e) { log("Modes de paiement : " + e.Message); } }
            int nTrouvees = ok;
            Ui(delegate
            {
                lblStep.Text = "Terminé. Fichier enregistré : " + outPath;
                bar.Value = bar.Maximum;
                lblSummary.Text = summary;
                btnOpen.Enabled = true;
                if (nTrouvees > 0)
                {
                    btnSolder.Text = "Marquer les " + nTrouvees + " facture(s) trouvée(s) en soldé";
                    btnSolder.Enabled = true;
                    btnSolder.Focus();
                }
                else btnOpen.Focus();
            });
        }

        static string Eur(double v) { return v.ToString("#,##0.00", CultureInfo.GetCultureInfo("fr-FR")) + " €"; }

        // Écrit (ou réécrit) la copie « … - rapproché.xlsx ». Après « Marquer en soldé », on réécrit
        // d'abord le même fichier ; s'il est ouvert dans Excel, on prend le nom libre suivant.
        string WriteExcel()
        {
            var values = new Dictionary<int, object[]>();
            var fills = new Dictionary<int, int>();
            foreach (ReportLine l in lines)
            {
                var v = new List<object> { l.Resultat, l.Numero, l.ClientSage, double.IsNaN(l.HtFacture) ? null : (object)l.HtFacture,
                    double.IsNaN(l.HtFacture) || double.IsNaN(l.HtValide) ? null : (object)Math.Round(l.HtFacture - l.HtValide, 2), l.Remarque };
                if (solderDone) v.Add(l.Reglement.Length > 0 ? l.Reglement : null);
                values[l.Row] = v.ToArray();
                fills[l.Row] = l.Fill;
            }
            var headers = new List<string> { "Rapprochement", "N° facture Sage", "Client Sage", "HT facture Sage", "Écart HT", "Remarque" };
            if (solderDone) headers.Add("Règlement Sage");
            int[] widths = { 16, 16, 30, 14, 10, 90, 45 };
            string dir = Path.GetDirectoryName(path), name = Path.GetFileNameWithoutExtension(path);
            var candidates = new List<string>();
            if (!string.IsNullOrEmpty(outPath)) candidates.Add(outPath);
            for (int n = 1; n < 50; n++) candidates.Add(Path.Combine(dir, name + " - rapproché" + (n > 1 ? " (" + n + ")" : "") + ".xlsx"));
            foreach (string dst in candidates)
            {
                try { Xlsx.WriteWithColumns(path, dst, sheet, headerRowIndex, headers.ToArray(), values, fills, widths); return dst; }
                catch (IOException) { /* fichier ouvert dans Excel : nom suivant */ }
                catch (UnauthorizedAccessException) { }
            }
            throw new Exception("Impossible d'écrire le fichier résultat à côté de « " + path + " ».");
        }

        // ── Marquer en soldé : un règlement par facture « Trouvée » ──

        volatile bool solderRunning;

        void AskSolder()
        {
            if (lines == null || solderRunning) return;
            var todo = lines.FindAll(delegate(ReportLine l) { return l.Resultat == PaymentReconciler.TROUVEE && l.Reglement.Length == 0 && l.Numero.Length > 0; });
            if (todo.Count == 0) return;
            if (modes == null || modes.Count == 0)
            {
                try { modes = sage.ListModes(); }
                catch (Exception e) { MessageBox.Show(this, "Impossible de lire les modes de paiement dans Sage : " + e.Message, "Marquer en soldé", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
                if (modes.Count == 0) { MessageBox.Show(this, "Aucun mode de paiement n'est défini dans Sage.", "Marquer en soldé", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
            }
            // Proposé par défaut : le mode le plus fréquent sur les factures à solder (l'utilisateur choisit).
            var count = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            if (pieces0 != null)
                foreach (ReportLine l in todo)
                {
                    PieceRow p = pieces0.Find(delegate(PieceRow x) { return string.Equals(x.Num, l.Numero, StringComparison.OrdinalIgnoreCase); });
                    if (p == null || p.ModeRglt.Length == 0) continue;
                    int c; count.TryGetValue(p.ModeRglt, out c); count[p.ModeRglt] = c + 1;
                }
            string suggested = null; int best = 0;
            foreach (KeyValuePair<string, int> kv in count) if (kv.Value > best) { best = kv.Value; suggested = kv.Key; }
            string exemple = todo[0].Order.Length > 0 ? " (ex. " + todo[0].Order + ")" : "";
            string msg = todo.Count + " facture(s) « Trouvée » vont être marquées SOLDÉES dans Sage.\n\n"
                + "Pour chacune, SageBridge enregistre un règlement client :\n"
                + "   • date du jour (" + DateTime.Today.ToString("dd/MM/yyyy") + ")\n"
                + "   • montant = reste dû TTC de la facture\n"
                + "   • référence = n° de commande" + exemple + "\n"
                + "   • mode de paiement choisi ci-dessous (le même pour toutes)\n\n"
                + "Les lignes « À vérifier », « Introuvable » et « Déjà soldée » ne sont pas touchées.";
            string mode;
            using (var dlg = new SolderDialog(msg, modes, suggested))
            {
                if (dlg.ShowDialog(this) != DialogResult.OK) return;
                mode = dlg.Mode;
            }
            if (string.IsNullOrEmpty(mode)) return;
            solderRunning = true;
            btnSolder.Enabled = false; btnClose.Enabled = false;
            var th = new Thread(delegate() { SolderWork(todo, mode); }) { IsBackground = true };
            th.Start();
        }

        void SolderWork(List<ReportLine> todo, string mode)
        {
            int ok = 0, skip = 0, err = 0, consecutive = 0;
            try
            {
                // Relecture : l'état a pu changer depuis le rapprochement (saisie dans Sage entre-temps).
                Step("Relecture des factures dans Sage…", 0, 1);
                List<PieceRow> fresh = sage.ListPieces(delegate(int done, int total) { Step("Relecture des factures dans Sage : " + done + " / " + total, done, total); });
                var byNum = new Dictionary<string, PieceRow>(StringComparer.OrdinalIgnoreCase);
                foreach (PieceRow p in fresh) if (p.Kind == "facture" && !byNum.ContainsKey(p.Num)) byNum[p.Num] = p;
                log("Marquage en soldé : " + todo.Count + " facture(s) à solder, mode de paiement " + mode + ".");

                for (int i = 0; i < todo.Count; i++)
                {
                    ReportLine l = todo[i];
                    Step("Solde de " + l.Numero + " (" + (i + 1) + " / " + todo.Count + ") : " + l.Client, i, todo.Count);
                    PieceRow p;
                    if (!byNum.TryGetValue(l.Numero, out p)) { l.Reglement = "Non soldée : facture introuvable à la relecture"; skip++; }
                    else if (p.Soldee) { l.Reglement = "Déjà soldée dans Sage"; skip++; }
                    else if (!p.Valide) { l.Reglement = "Non soldée : facture provisoire"; skip++; }
                    else if (p.SoldeDu <= 0.004) { l.Reglement = "Non soldée : solde dû nul"; skip++; }
                    else if (Math.Abs(p.TotalHT - l.HtValide) > 0.011) { l.Reglement = "Non soldée : le HT de la facture a changé (" + Eur(p.TotalHT) + ")"; skip++; }
                    else
                    {
                        double montant = Math.Round(p.SoldeDu, 2);
                        try
                        {
                            SolderResult r = sage.SolderFacture(p.Num, p.Client, montant, l.Order, mode, false);
                            l.Reglement = "Soldée – règlement " + r.Numero + " de " + Eur(montant) + (r.Mode.Length > 0 ? " (" + r.Mode + ")" : "");
                            ok++; consecutive = 0;
                        }
                        catch (Exception e)
                        {
                            l.Reglement = "Erreur : " + e.Message;
                            log("Solde de " + p.Num + " impossible : " + e.Message);
                            err++; consecutive++;
                        }
                    }
                    UpdateRow(l);
                    if (consecutive >= 3)
                    {
                        for (int k = i + 1; k < todo.Count; k++) { todo[k].Reglement = "Non traitée (arrêt après 3 erreurs de suite)"; UpdateRow(todo[k]); }
                        log("Marquage en soldé arrêté après 3 erreurs de suite.");
                        break;
                    }
                }

                // Vérification : le statut Sage doit être « Soldé ».
                Step("Vérification dans Sage…", todo.Count, todo.Count);
                var after = new Dictionary<string, PieceRow>(StringComparer.OrdinalIgnoreCase);
                foreach (PieceRow p in sage.ListPieces(null)) if (p.Kind == "facture" && !after.ContainsKey(p.Num)) after[p.Num] = p;
                int notSolde = 0;
                foreach (ReportLine l in todo)
                {
                    if (!l.Reglement.StartsWith("Soldée")) continue;
                    PieceRow p;
                    if (after.TryGetValue(l.Numero, out p) && !p.Soldee)
                    {
                        l.Reglement += " — ATTENTION : statut Sage « " + p.Statut + " » (reste dû " + Eur(p.SoldeDu) + ")";
                        notSolde++; UpdateRow(l);
                    }
                }

                solderDone = true;
                outPath = WriteExcel();
                string summary = ok + " facture(s) soldée(s)" + (skip > 0 ? ", " + skip + " ignorée(s)" : "") + (err > 0 ? ", " + err + " erreur(s)" : "")
                    + (notSolde > 0 ? ", " + notSolde + " sans statut Soldé" : "") + ".";
                log("Marquage en soldé terminé : " + summary + " Fichier : " + outPath);
                bool warn = err > 0 || notSolde > 0;
                Ui(delegate
                {
                    lblStep.Text = "Marquage terminé. Fichier mis à jour : " + outPath;
                    bar.Value = bar.Maximum;
                    lblSummary.Text = summary;
                    btnSolder.Text = "Factures soldées";
                    MessageBox.Show(this, summary + "\n\nLe détail est dans la colonne « Règlement Sage » du fichier Excel.", "Marquer en soldé",
                        MessageBoxButtons.OK, warn ? MessageBoxIcon.Warning : MessageBoxIcon.Information);
                });
            }
            catch (Exception e)
            {
                log("Marquage en soldé interrompu : " + e.Message);
                int done = ok;
                Ui(delegate
                {
                    lblStep.Text = "Erreur : " + e.Message; lblStep.ForeColor = Color.Firebrick;
                    MessageBox.Show(this, e.Message + (done > 0 ? "\n\n" + done + " facture(s) ont déjà été soldées (voir le journal)." : ""), "Marquer en soldé", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                });
            }
            finally
            {
                solderRunning = false;
                Ui(delegate { btnClose.Enabled = true; });
            }
        }

        void UpdateRow(ReportLine l)
        {
            ReportLine cur = l;
            Ui(delegate
            {
                if (cur.Item == null) return;
                ListViewItem.ListViewSubItem s = cur.Item.SubItems[7];
                s.Text = cur.Reglement;
                bool good = cur.Reglement.StartsWith("Soldée") && !cur.Reglement.Contains("ATTENTION");
                bool bad = cur.Reglement.StartsWith("Erreur") || cur.Reglement.Contains("ATTENTION");
                s.BackColor = good ? Color.FromArgb(198, 239, 206) : bad ? Color.FromArgb(255, 199, 206) : Color.FromArgb(231, 230, 230);
                cur.Item.EnsureVisible();
            });
        }

        void AddRow(ReportLine l)
        {
            var it = new ListViewItem(new[] {
                l.Row.ToString(), l.Order, l.Client, double.IsNaN(l.HtValide) ? "" : l.HtValide.ToString("0.00"),
                l.Resultat, l.Numero, l.ClientSage, l.Reglement, l.Remarque });
            Color[] bg = { Color.FromArgb(198, 239, 206), Color.FromArgb(255, 235, 156), Color.FromArgb(255, 199, 206), Color.FromArgb(231, 230, 230) };
            it.UseItemStyleForSubItems = false;
            if (l.Fill >= 0 && l.Fill < bg.Length) it.SubItems[4].BackColor = bg[l.Fill];
            l.Item = it;
            list.Items.Add(it);
            it.EnsureVisible();
        }
    }

    // Choix du mode de paiement des règlements, parmi les modes du dossier Sage.
    class SolderDialog : Form
    {
        readonly ComboBox box;
        readonly List<string[]> modes;
        public string Mode { get { return box.SelectedIndex >= 0 ? modes[box.SelectedIndex][0] : null; } }

        public SolderDialog(string message, List<string[]> list, string suggested)
        {
            modes = list;
            Ui2.Setup(this, "Marquer en soldé");
            var g = Ui2.Grid(2);
            g.Controls.Add(new Label { Text = message, AutoSize = true, MaximumSize = new Size(Ui2.Px(560), 0), Margin = new Padding(3, 3, 3, 12) }, 0, 0);
            g.SetColumnSpan(g.GetControlFromPosition(0, 0), 2);
            g.Controls.Add(Ui2.L("Mode de paiement :"), 0, 1);
            box = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = Ui2.Px(300) };
            foreach (string[] m in list) box.Items.Add(m[0] + (m[1].Length > 0 ? " – " + m[1] : ""));
            int sel = -1;
            if (!string.IsNullOrEmpty(suggested))
                for (int i = 0; i < list.Count; i++) if (string.Equals(list[i][0], suggested, StringComparison.OrdinalIgnoreCase)) sel = i;
            box.SelectedIndex = sel;   // rien de présélectionné si aucune suggestion : l'utilisateur choisit
            g.Controls.Add(box, 1, 1);
            if (!string.IsNullOrEmpty(suggested) && sel >= 0)
                g.Controls.Add(new Label { Text = "Proposé : le mode inscrit sur la plupart de ces factures.", AutoSize = true, ForeColor = Color.DimGray, Margin = new Padding(3, 4, 3, 3) }, 1, 2);
            Controls.Add(g);
            var buttons = Ui2.OkCancel(this, "Solder les factures");
            Controls.Add(buttons);
            FormClosing += delegate(object s, FormClosingEventArgs e)
            {
                if (DialogResult == DialogResult.OK && box.SelectedIndex < 0)
                {
                    MessageBox.Show(this, "Choisissez le mode de paiement.", "Marquer en soldé", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    e.Cancel = true;
                }
            };
        }
    }

    // ───────────────────────── Export PDF : pilotage de l'aperçu de Sage ─────────────────────────
    // Reproduit le geste de l'utilisateur : aperçu de la pièce (modèle choisi) > Exporter en PDF > Exporter tout
    // en PDF > Enregistrer sous (chemin complet) > fermeture de l'aperçu. Sage n'offre pas d'export PDF par COM
    // (« Ficher PDF » → « Paramètres incorrects », tâche EngPrn → « Fonction non implémentée ») : on passe par
    // des messages Windows envoyés aux fenêtres de WGC.exe, sans toucher à la souris ni au clavier réels.
    // Vérifié live le 03/11/2026 (FA10005913 : fichier identique à l'octet près à l'export manuel).
    static class SagePdfExporter
    {
        delegate bool EnumProc(IntPtr h, IntPtr l);
        [DllImport("user32.dll")] static extern bool EnumWindows(EnumProc f, IntPtr l);
        [DllImport("user32.dll")] static extern bool EnumChildWindows(IntPtr p, EnumProc f, IntPtr l);
        [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
        [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr h);
        [DllImport("user32.dll")] static extern bool IsWindow(IntPtr h);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetClassName(IntPtr h, StringBuilder s, int n);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetWindowText(IntPtr h, StringBuilder s, int n);
        [DllImport("user32.dll")] static extern IntPtr SendMessageTimeout(IntPtr h, uint m, IntPtr w, IntPtr l, uint flags, uint timeout, out IntPtr result);
        [DllImport("user32.dll", CharSet = CharSet.Unicode, EntryPoint = "SendMessageTimeoutW")] static extern IntPtr SendMessageTimeoutStr(IntPtr h, uint m, IntPtr w, string l, uint flags, uint timeout, out IntPtr result);
        [DllImport("user32.dll")] static extern bool PostMessage(IntPtr h, uint m, IntPtr w, IntPtr l);
        [DllImport("user32.dll")] static extern int GetDlgCtrlID(IntPtr h);
        [DllImport("user32.dll")] static extern uint GetDpiForWindow(IntPtr h);

        const uint WM_SETTEXT = 0x000C, WM_CLOSE = 0x0010, WM_COMMAND = 0x0111, WM_KEYDOWN = 0x0100, WM_KEYUP = 0x0101,
            WM_MOUSEMOVE = 0x0200, WM_LBUTTONDOWN = 0x0201, WM_LBUTTONUP = 0x0202, MN_SELECTITEM = 0x01E5, SMTO_ABORTIFHUNG = 2;
        const int VK_RETURN = 0x0D, VK_ESCAPE = 0x1B, IDOK = 1, IDCANCEL = 2;
        // Bouton « Exporter en PDF » dans le ruban de l'aperçu (coordonnées client du ruban à 96 ppp, Sage 50 v6.71).
        const int EXPORT_X = 232, EXPORT_Y = 18;

        class Win { public IntPtr H; public string Cls = "", Title = ""; }

        static string ClassOf(IntPtr h) { var sb = new StringBuilder(256); GetClassName(h, sb, 256); return sb.ToString(); }
        static string TitleOf(IntPtr h) { var sb = new StringBuilder(512); GetWindowText(h, sb, 512); return sb.ToString(); }

        static List<Win> TopWindows(int pid)
        {
            var r = new List<Win>();
            EnumWindows(delegate(IntPtr h, IntPtr l)
            {
                uint p; GetWindowThreadProcessId(h, out p);
                if (p == (uint)pid && IsWindowVisible(h)) r.Add(new Win { H = h, Cls = ClassOf(h), Title = TitleOf(h) });
                return true;
            }, IntPtr.Zero);
            return r;
        }

        static List<Win> Children(IntPtr parent)
        {
            var r = new List<Win>();
            EnumChildWindows(parent, delegate(IntPtr h, IntPtr l) { r.Add(new Win { H = h, Cls = ClassOf(h), Title = TitleOf(h) }); return true; }, IntPtr.Zero);
            return r;
        }

        static int SagePid()
        {
            foreach (Process p in Process.GetProcessesByName("WGC"))
            {
                try { if (p.MainWindowHandle != IntPtr.Zero) return p.Id; }
                catch { }
                finally { p.Dispose(); }
            }
            return 0;
        }

        static Win WaitFor(Func<Win> probe, int ms)
        {
            DateTime end = DateTime.Now.AddMilliseconds(ms);
            while (true)
            {
                Win w = probe();
                if (w != null || DateTime.Now >= end) return w;
                Thread.Sleep(200);
            }
        }

        static void Command(IntPtr dlg, int id, IntPtr ctl)
        {
            IntPtr r; SendMessageTimeout(dlg, WM_COMMAND, (IntPtr)id, ctl, SMTO_ABORTIFHUNG, 5000, out r);
        }

        static void Click(IntPtr h, int x, int y)
        {
            IntPtr lp = (IntPtr)((y << 16) | (x & 0xFFFF));
            PostMessage(h, WM_MOUSEMOVE, IntPtr.Zero, lp);
            PostMessage(h, WM_LBUTTONDOWN, (IntPtr)1, lp);
            Thread.Sleep(120);
            PostMessage(h, WM_LBUTTONUP, IntPtr.Zero, lp);
        }

        // Texte des libellés d'une boîte de message de Sage (pour l'afficher à l'utilisateur).
        static string DialogText(IntPtr dlg)
        {
            var parts = new List<string>();
            foreach (Win c in Children(dlg))
                if (c.Cls == "Static" && c.Title.Trim().Length > 0) parts.Add(c.Title.Trim());
            return string.Join(" ", parts.ToArray());
        }

        static bool HasOkButton(IntPtr dlg)
        {
            foreach (Win c in Children(dlg))
                if (c.Cls == "Button" && c.Title.Replace("&", "").Trim() == "OK" && IsWindowVisible(c.H)) return true;
            return false;
        }

        // Ferme proprement les fenêtres ouvertes par l'export : boîtes de message (OK), « Enregistrer sous » (Annuler), menus.
        static void CloseStray(int pid, HashSet<IntPtr> before)
        {
            foreach (Win w in TopWindows(pid))
            {
                if (before.Contains(w.H)) continue;
                if (w.Cls == "#32768") { PostMessage(w.H, WM_KEYDOWN, (IntPtr)VK_ESCAPE, IntPtr.Zero); PostMessage(w.H, WM_KEYUP, (IntPtr)VK_ESCAPE, IntPtr.Zero); }
                else if (w.Cls == "#32770") Command(w.H, IDCANCEL, IntPtr.Zero);
                else if (w.Cls == "CamDialogNC")
                {
                    foreach (Win c in Children(w.H))
                        if (c.Cls == "Button" && c.Title.Replace("&", "") == "OK") { Command(w.H, GetDlgCtrlID(c.H), c.H); break; }
                }
            }
        }

        public static void Export(SageWorker sage, PdfPiece p, string model, string target)
        {
            int pid = SagePid();
            if (pid == 0) throw new Exception("Sage 50 n'est pas ouvert.");
            var before = new HashSet<IntPtr>();
            foreach (Win w in TopWindows(pid)) before.Add(w.H);

            sage.OpenPreview(p, model);
            string suffix = " - " + p.Num;
            Win prev = WaitFor(delegate
            {
                foreach (Win w in TopWindows(pid))
                    if (!before.Contains(w.H) && w.Cls == "CamDialog" && w.Title.StartsWith("Aper") && w.Title.EndsWith(suffix, StringComparison.OrdinalIgnoreCase)) return w;
                return null;
            }, 30000);
            if (prev == null)
            {
                CloseStray(pid, before);
                throw new Exception("l'aperçu de la pièce ne s'est pas ouvert dans Sage (modèle « " + model + " » introuvable ?).");
            }
            try
            {
                Thread.Sleep(700); // le ruban et la page finissent de s'afficher
                IntPtr ribbon = IntPtr.Zero;
                foreach (Win c in Children(prev.H)) if (c.Cls == "{{ribbon}}") { ribbon = c.H; break; }
                if (ribbon == IntPtr.Zero) throw new Exception("barre d'outils de l'aperçu introuvable.");
                uint dpi = 96;
                try { dpi = GetDpiForWindow(ribbon); if (dpi == 0) dpi = 96; } catch { dpi = 96; }
                int x = (int)(EXPORT_X * dpi / 96), y = (int)(EXPORT_Y * dpi / 96);

                Win menu = null;
                for (int attempt = 0; attempt < 3 && menu == null; attempt++)
                {
                    Click(ribbon, x, y);
                    menu = WaitFor(delegate
                    {
                        foreach (Win w in TopWindows(pid)) if (!before.Contains(w.H) && w.Cls == "#32768") return w;
                        return null;
                    }, 3000);
                }
                if (menu == null) throw new Exception("le menu « Exporter en PDF » de l'aperçu ne s'est pas ouvert.");
                IntPtr r;
                SendMessageTimeout(menu.H, MN_SELECTITEM, IntPtr.Zero, IntPtr.Zero, SMTO_ABORTIFHUNG, 3000, out r); // « Exporter tout en PDF »
                Thread.Sleep(150);
                PostMessage(menu.H, WM_KEYDOWN, (IntPtr)VK_RETURN, IntPtr.Zero);
                PostMessage(menu.H, WM_KEYUP, (IntPtr)VK_RETURN, IntPtr.Zero);

                Win save = WaitFor(delegate
                {
                    foreach (Win w in TopWindows(pid)) if (!before.Contains(w.H) && w.Cls == "#32770") return w;
                    return null;
                }, 20000);
                if (save == null) throw new Exception("la fenêtre « Enregistrer sous » de Sage ne s'est pas ouverte.");
                Thread.Sleep(300);
                IntPtr edit = IntPtr.Zero, okBtn = IntPtr.Zero;
                foreach (Win c in Children(save.H))
                {
                    int id = GetDlgCtrlID(c.H);
                    if (c.Cls == "Edit" && (id == 1148 || id == 1152) && edit == IntPtr.Zero) edit = c.H;
                    if (c.Cls == "Button" && id == IDOK) okBtn = c.H;
                }
                if (edit == IntPtr.Zero) throw new Exception("zone « Nom du fichier » introuvable dans la fenêtre « Enregistrer sous ».");
                SendMessageTimeoutStr(edit, WM_SETTEXT, IntPtr.Zero, target, SMTO_ABORTIFHUNG, 5000, out r);
                Thread.Sleep(150);
                Command(save.H, IDOK, okBtn);

                // Attente du fichier (taille stable) ; toute nouvelle boîte de Sage est une erreur à remonter.
                DateTime end = DateTime.Now.AddSeconds(90);
                long lastLen = -1; int stable = 0;
                while (DateTime.Now < end)
                {
                    Thread.Sleep(400);
                    try
                    {
                        var fi = new FileInfo(target);
                        if (fi.Exists && fi.Length > 0)
                        {
                            if (fi.Length == lastLen) { if (++stable >= 2) return; }
                            else { lastLen = fi.Length; stable = 0; }
                            continue;
                        }
                    }
                    catch { }
                    foreach (Win w in TopWindows(pid))
                    {
                        if (before.Contains(w.H) || w.H == prev.H) continue;
                        // Erreur = boîte de message avec un bouton OK. La fenêtre de progression de Sage
                        // (« … calcul en cours … Traitement en cours … », CamDialogNC sans OK) est normale.
                        if ((w.Cls == "CamDialogNC" || (w.Cls == "#32770" && w.H != save.H)) && HasOkButton(w.H))
                            throw new Exception("Sage : " + DialogText(w.H));
                    }
                }
                throw new Exception("le fichier PDF n'a pas été créé dans le délai (90 s).");
            }
            finally
            {
                CloseStray(pid, before);
                PostMessage(prev.H, WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
                DateTime end = DateTime.Now.AddSeconds(8);
                while (IsWindow(prev.H) && IsWindowVisible(prev.H) && DateTime.Now < end) Thread.Sleep(200);
            }
        }
    }

    // Rangement des PDF : Recap\<année>\<n°>. <Mois>\Facture|Avoir\<numéro>.pdf. Un dossier existant est réutilisé
    // tel quel (« Novembre » ou « 11. Novembre », « Facture » ou « Factures ») ; sinon il est créé au format « 11. Novembre ».
    static class RecapFolders
    {
        public static readonly string[] Mois = { "Janvier", "Février", "Mars", "Avril", "Mai", "Juin", "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre" };

        static string FindSub(string dir, Func<string, bool> match)
        {
            if (!Directory.Exists(dir)) return null;
            try { foreach (string d in Directory.GetDirectories(dir)) if (match(Path.GetFileName(d))) return d; } catch { }
            return null;
        }

        public static string MonthDir(string root, int year, int month, bool create)
        {
            string yearDir = Path.Combine(root, year.ToString(CultureInfo.InvariantCulture));
            string want = Norm.Text(Mois[month - 1]);
            string found = FindSub(yearDir, delegate(string name) { return Regex.Replace(Norm.Text(name), @"^\d+\s*", "") == want; });
            if (found != null) return found;
            string dir = Path.Combine(yearDir, month + ". " + Mois[month - 1]);
            if (create) Directory.CreateDirectory(dir);
            return dir;
        }

        public static string TypeDir(string monthDir, string kind, bool create)
        {
            string a = kind == "avoir" ? "avoir" : "facture";
            string found = FindSub(monthDir, delegate(string name) { string n = Norm.Text(name); return n == a || n == a + "s"; });
            if (found != null) return found;
            string dir = Path.Combine(monthDir, kind == "avoir" ? "Avoir" : "Facture");
            if (create) Directory.CreateDirectory(dir);
            return dir;
        }

        // PDF déjà présents n'importe où sous Recap, indexés par le numéro de pièce en tête du nom :
        // « FA10005662 Reçu le 06-10-2026.pdf » compte pour FA10005662 (fichiers renommés ou déplacés à la main).
        public static Dictionary<string, string> IndexExisting(string root)
        {
            var idx = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var stack = new Stack<string>();
            if (Directory.Exists(root)) stack.Push(root);
            var re = new Regex(@"^([A-Za-z]*\d+)");
            while (stack.Count > 0)
            {
                string d = stack.Pop();
                try { foreach (string sub in Directory.GetDirectories(d)) stack.Push(sub); } catch { }
                try
                {
                    foreach (string f in Directory.GetFiles(d, "*.pdf"))
                    {
                        Match m = re.Match(Path.GetFileName(f));
                        if (m.Success && !idx.ContainsKey(m.Groups[1].Value)) idx[m.Groups[1].Value] = f;
                    }
                }
                catch { }
            }
            return idx;
        }
    }

    class PdfExportForm : Form
    {
        readonly Settings cfg;
        readonly SageWorker sage;
        readonly Action<string> log;
        readonly ComboBox cboMonth;
        readonly NumericUpDown numYear;
        readonly CheckBox chkFac, chkAvo;
        readonly TextBox txtRoot, txtModel;
        readonly Button btnSearch, btnExport, btnStop, btnOpenDir, btnClose;
        readonly Label lblStep, lblSummary;
        readonly ProgressBar bar;
        readonly ListView list;
        List<PdfPiece> pieces = new List<PdfPiece>();
        readonly Dictionary<PdfPiece, ListViewItem> rows = new Dictionary<PdfPiece, ListViewItem>();
        volatile bool running, stop, closing;
        int curYear, curMonth;

        public PdfExportForm(Settings settings, SageWorker worker, Action<string> logger)
        {
            cfg = settings; sage = worker; log = logger;
            Text = "Enregistrer les factures en PDF";
            Font = new Font("Segoe UI", 9.5f);
            Icon = Program.AppIcon;
            StartPosition = FormStartPosition.CenterParent;
            ShowInTaskbar = false;
            MinimizeBox = false;
            Size = new Size(Ui2.Px(1000), Ui2.Px(620));

            var top = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 1, Padding = new Padding(Ui2.Px(10)) };
            var row1 = new FlowLayoutPanel { AutoSize = true, WrapContents = false, Margin = new Padding(0) };
            cboMonth = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = Ui2.Px(120) };
            cboMonth.Items.AddRange(RecapFolders.Mois);
            cboMonth.SelectedIndex = DateTime.Today.Month - 1;
            numYear = new NumericUpDown { Minimum = 2000, Maximum = 2100, Value = DateTime.Today.Year, Width = Ui2.Px(70) };
            chkFac = new CheckBox { Text = "Factures", Checked = true, AutoSize = true, Margin = new Padding(Ui2.Px(14), Ui2.Px(5), 0, 0) };
            chkAvo = new CheckBox { Text = "Avoirs", Checked = true, AutoSize = true, Margin = new Padding(Ui2.Px(6), Ui2.Px(5), 0, 0) };
            btnSearch = new Button { Text = "Rechercher", AutoSize = true, Margin = new Padding(Ui2.Px(14), 0, 0, 0) };
            btnSearch.Click += delegate { Search(); };
            row1.Controls.Add(new Label { Text = "Mois :", AutoSize = true, Margin = new Padding(0, Ui2.Px(6), Ui2.Px(4), 0) });
            row1.Controls.Add(cboMonth);
            row1.Controls.Add(new Label { Text = "Année :", AutoSize = true, Margin = new Padding(Ui2.Px(12), Ui2.Px(6), Ui2.Px(4), 0) });
            row1.Controls.Add(numYear);
            row1.Controls.Add(chkFac); row1.Controls.Add(chkAvo); row1.Controls.Add(btnSearch);

            var row2 = new FlowLayoutPanel { AutoSize = true, WrapContents = false, Margin = new Padding(0, Ui2.Px(6), 0, 0) };
            txtRoot = new TextBox { Text = cfg.RecapDir, Width = Ui2.Px(420) };
            var browse = new Button { Text = "Choisir…", AutoSize = true };
            browse.Click += delegate { ChooseRoot(); };
            txtModel = new TextBox { Text = cfg.PdfModel, Width = Ui2.Px(170) };
            row2.Controls.Add(new Label { Text = "Dossier :", AutoSize = true, Margin = new Padding(0, Ui2.Px(6), Ui2.Px(4), 0) });
            row2.Controls.Add(txtRoot); row2.Controls.Add(browse);
            row2.Controls.Add(new Label { Text = "Modèle Sage :", AutoSize = true, Margin = new Padding(Ui2.Px(12), Ui2.Px(6), Ui2.Px(4), 0) });
            row2.Controls.Add(txtModel);

            lblStep = new Label { AutoSize = true, Text = "Choisissez le mois puis « Rechercher ».", Margin = new Padding(3, Ui2.Px(8), 3, Ui2.Px(4)) };
            bar = new ProgressBar { Dock = DockStyle.Top, Height = Ui2.Px(18), Minimum = 0, Maximum = 100 };
            top.Controls.Add(row1); top.Controls.Add(row2); top.Controls.Add(lblStep); top.Controls.Add(bar);

            list = new ListView { Dock = DockStyle.Fill, View = View.Details, FullRowSelect = true, GridLines = true, HideSelection = false };
            string[] cols = { "Pièce", "Type", "Date", "Client", "Total TTC", "État", "Fichier" };
            int[] widths = { 105, 65, 85, 190, 80, 150, 420 };
            for (int i = 0; i < cols.Length; i++) list.Columns.Add(cols[i], Ui2.Px(widths[i]));
            list.DoubleClick += delegate
            {
                if (list.SelectedItems.Count == 0) return;
                var p = (PdfPiece)list.SelectedItems[0].Tag;
                string f = p.Existing.Length > 0 ? p.Existing : p.Target;
                if (File.Exists(f)) { try { Process.Start(f); } catch { } }
            };

            var bottom = new FlowLayoutPanel { Dock = DockStyle.Bottom, AutoSize = true, FlowDirection = FlowDirection.RightToLeft, Padding = new Padding(Ui2.Px(8)) };
            btnClose = new Button { Text = "Fermer", AutoSize = true };
            btnClose.Click += delegate { Close(); };
            btnOpenDir = new Button { Text = "Ouvrir le dossier", AutoSize = true, Enabled = false };
            btnOpenDir.Click += delegate { OpenMonthDir(); };
            btnStop = new Button { Text = "Arrêter", AutoSize = true, Enabled = false };
            btnStop.Click += delegate { stop = true; btnStop.Enabled = false; lblStep.Text = "Arrêt après la pièce en cours…"; };
            btnExport = new Button { Text = "Enregistrer les PDF manquants", AutoSize = true, Enabled = false, Font = new Font("Segoe UI Semibold", 9.5f) };
            btnExport.Click += delegate { StartExport(); };
            lblSummary = new Label { AutoSize = true, Margin = new Padding(3, Ui2.Px(7), Ui2.Px(20), 3) };
            bottom.Controls.Add(btnClose); bottom.Controls.Add(btnOpenDir); bottom.Controls.Add(btnStop); bottom.Controls.Add(btnExport); bottom.Controls.Add(lblSummary);

            Controls.Add(list); Controls.Add(top); Controls.Add(bottom);
            AcceptButton = btnSearch;
            FormClosing += delegate(object s, FormClosingEventArgs e)
            {
                if (running && e.CloseReason == CloseReason.UserClosing)
                {
                    MessageBox.Show(this, "L'enregistrement est en cours : cliquez sur « Arrêter » et attendez la fin de la pièce en cours.", "SageBridge", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    e.Cancel = true; return;
                }
                closing = true;
            };
            // Aucun dossier imposé : au premier usage, on demande où enregistrer les factures (retenu ensuite).
            Shown += delegate
            {
                if (txtRoot.Text.Trim().Length == 0)
                {
                    lblStep.Text = "Choisissez d'abord le dossier où enregistrer les factures (bouton « Choisir… »).";
                    ChooseRoot();
                }
            };
        }

        // Dossier racine des factures (ex. Bureau\Recap) : choisi par l'utilisateur, enregistré dans les réglages.
        bool ChooseRoot()
        {
            using (var dlg = new FolderBrowserDialog { Description = "Dans quel dossier enregistrer les factures ? (les sous-dossiers année / mois / Facture y seront créés)", ShowNewFolderButton = true })
            {
                if (Directory.Exists(txtRoot.Text.Trim())) dlg.SelectedPath = txtRoot.Text.Trim();
                if (dlg.ShowDialog(this) != DialogResult.OK) return false;
                txtRoot.Text = dlg.SelectedPath;
            }
            cfg.RecapDir = txtRoot.Text.Trim();
            try { cfg.Save(); } catch { }
            lblStep.Text = "Choisissez le mois puis « Rechercher ».";
            return true;
        }

        void Ui(Action a)
        {
            if (IsDisposed || closing) return;
            try { BeginInvoke(a); } catch { }
        }

        void SetBusy(bool b)
        {
            running = b;
            btnSearch.Enabled = !b; btnExport.Enabled = !b && Missing().Count > 0; btnStop.Enabled = b;
            cboMonth.Enabled = numYear.Enabled = chkFac.Enabled = chkAvo.Enabled = txtRoot.Enabled = txtModel.Enabled = !b;
        }

        List<PdfPiece> Missing()
        {
            var m = new List<PdfPiece>();
            foreach (PdfPiece p in pieces) if (p.Existing.Length == 0 && p.Status != "ok") m.Add(p);
            return m;
        }

        void SaveChoices()
        {
            cfg.RecapDir = txtRoot.Text.Trim();
            if (txtModel.Text.Trim().Length > 0) cfg.PdfModel = txtModel.Text.Trim();
            try { cfg.Save(); } catch { }
        }

        void Search()
        {
            if (txtRoot.Text.Trim().Length == 0 && !ChooseRoot()) return;
            if (!Directory.Exists(txtRoot.Text.Trim()))
            {
                MessageBox.Show(this, "Le dossier « " + txtRoot.Text.Trim() + " » n'existe pas. Choisissez le dossier où enregistrer les factures.", "SageBridge", MessageBoxButtons.OK, MessageBoxIcon.Information);
                ChooseRoot();
                return;
            }
            if (!chkFac.Checked && !chkAvo.Checked) { MessageBox.Show(this, "Cochez « Factures » et/ou « Avoirs ».", "SageBridge"); return; }
            SaveChoices();
            curYear = (int)numYear.Value; curMonth = cboMonth.SelectedIndex + 1;
            bool fac = chkFac.Checked, avo = chkAvo.Checked;
            string root = cfg.RecapRoot;
            list.Items.Clear(); rows.Clear(); pieces = new List<PdfPiece>();
            lblSummary.Text = ""; btnOpenDir.Enabled = false;
            SetBusy(true);
            lblStep.Text = "Lecture des pièces de " + RecapFolders.Mois[curMonth - 1].ToLowerInvariant() + " " + curYear + " dans Sage…";
            bar.Style = ProgressBarStyle.Marquee;
            var th = new Thread(delegate()
            {
                try
                {
                    List<PdfPiece> all = sage.ListPiecesForMonth(curYear, curMonth);
                    Dictionary<string, string> idx = RecapFolders.IndexExisting(root);
                    string monthDir = RecapFolders.MonthDir(root, curYear, curMonth, false);
                    var sel = new List<PdfPiece>();
                    foreach (PdfPiece p in all)
                    {
                        if ((p.Kind == "facture" && !fac) || (p.Kind == "avoir" && !avo)) continue;
                        string ex; p.Existing = idx.TryGetValue(p.Num, out ex) ? ex : "";
                        p.Target = Path.Combine(RecapFolders.TypeDir(monthDir, p.Kind, false), p.Num + ".pdf");
                        sel.Add(p);
                    }
                    Ui(delegate
                    {
                        pieces = sel;
                        foreach (PdfPiece p in sel) AddRow(p);
                        bar.Style = ProgressBarStyle.Blocks; bar.Value = 0;
                        int miss = Missing().Count;
                        lblStep.Text = sel.Count == 0 ? "Aucune facture ni avoir validé à cette date dans Sage."
                            : sel.Count + " pièce(s) validée(s) : " + (sel.Count - miss) + " déjà enregistrée(s), " + miss + " manquante(s).";
                        btnOpenDir.Enabled = true;
                        SetBusy(false);
                        if (miss > 0) btnExport.Text = "Enregistrer les " + miss + " PDF manquants";
                    });
                }
                catch (Exception e)
                {
                    log("Export PDF : " + e.Message);
                    Ui(delegate { bar.Style = ProgressBarStyle.Blocks; lblStep.Text = "Erreur : " + e.Message; SetBusy(false); });
                }
            }) { IsBackground = true };
            th.Start();
        }

        void AddRow(PdfPiece p)
        {
            var it = new ListViewItem(new[] { p.Num, p.Kind == "avoir" ? "Avoir" : "Facture", p.Date.ToString("dd/MM/yyyy"), p.Nom,
                p.TotalTTC.ToString("0.00", CultureInfo.GetCultureInfo("fr-FR")), "", "" });
            it.Tag = p;
            rows[p] = it;
            list.Items.Add(it);
            RefreshRow(p);
        }

        void RefreshRow(PdfPiece p)
        {
            ListViewItem it;
            if (!rows.TryGetValue(p, out it)) return;
            if (p.Existing.Length > 0) { it.SubItems[5].Text = "Déjà enregistrée"; it.SubItems[6].Text = p.Existing; it.ForeColor = Color.DimGray; }
            else if (p.Status == "ok") { it.SubItems[5].Text = "Enregistrée"; it.SubItems[6].Text = p.Target; it.ForeColor = Color.DarkGreen; }
            else if (p.Status == "err") { it.SubItems[5].Text = "Erreur"; it.SubItems[6].Text = p.Error; it.ForeColor = Color.Firebrick; }
            else if (p.Status == "run") { it.SubItems[5].Text = "En cours…"; it.SubItems[6].Text = p.Target; it.ForeColor = Color.RoyalBlue; it.EnsureVisible(); }
            else { it.SubItems[5].Text = "Manquante"; it.SubItems[6].Text = p.Target; it.ForeColor = Color.Black; }
        }

        void OpenMonthDir()
        {
            string d = RecapFolders.MonthDir(cfg.RecapRoot, curYear, curMonth, false);
            if (!Directory.Exists(d)) d = Path.GetDirectoryName(d);
            if (!Directory.Exists(d)) d = cfg.RecapRoot;
            try { Process.Start("explorer.exe", "\"" + d + "\""); } catch { }
        }

        void StartExport()
        {
            List<PdfPiece> todo = Missing();
            if (todo.Count == 0) return;
            if (!SageWorker.SageWindowOpen()) { MessageBox.Show(this, "Ouvrez d'abord Sage 50.", "SageBridge", MessageBoxButtons.OK, MessageBoxIcon.Information); return; }
            if (MessageBox.Show(this, todo.Count + " PDF vont être créés au modèle « " + cfg.PdfModel + " » dans :\n" +
                RecapFolders.MonthDir(cfg.RecapRoot, curYear, curMonth, false) + "\n\n" +
                "Pendant l'enregistrement (2 à 3 s par pièce), n'utilisez pas Sage : ses aperçus vont s'ouvrir et se fermer tout seuls.\n" +
                "Fermez aussi toute fenêtre ouverte dans Sage (fiche, saisie, aperçu).\n\nContinuer ?",
                "Enregistrer les PDF", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) != DialogResult.OK) return;
            string model = cfg.PdfModel, root = cfg.RecapRoot;
            int year = curYear, month = curMonth;
            stop = false;
            SetBusy(true);
            bar.Maximum = todo.Count; bar.Value = 0;
            log("Export PDF de " + todo.Count + " pièce(s) de " + month.ToString("00") + "/" + year + " (modèle « " + model + " »).");
            var th = new Thread(delegate()
            {
                int ok = 0, ko = 0, consecutive = 0;
                for (int i = 0; i < todo.Count; i++)
                {
                    if (stop || closing) break;
                    PdfPiece p = todo[i];
                    int n = i;
                    p.Status = "run";
                    Ui(delegate { lblStep.Text = "Enregistrement " + (n + 1) + " / " + todo.Count + " : " + p.Num + " — " + p.Nom; RefreshRow(p); });
                    try
                    {
                        string dir = RecapFolders.TypeDir(RecapFolders.MonthDir(root, year, month, true), p.Kind, true);
                        p.Target = Path.Combine(dir, p.Num + ".pdf");
                        if (File.Exists(p.Target)) { p.Existing = p.Target; p.Status = ""; }
                        else
                        {
                            SagePdfExporter.Export(sage, p, model, p.Target);
                            p.Status = "ok"; ok++; consecutive = 0;
                            log("PDF enregistré : " + p.Target);
                        }
                    }
                    catch (Exception e)
                    {
                        p.Status = "err"; p.Error = e.Message; ko++; consecutive++;
                        log("PDF de " + p.Num + " : " + e.Message);
                    }
                    Ui(delegate { RefreshRow(p); bar.Value = Math.Min(bar.Maximum, n + 1); });
                    if (consecutive >= 3) { log("Export PDF interrompu après 3 erreurs de suite."); break; }
                }
                bool interrupted = consecutive >= 3;
                Ui(delegate
                {
                    SetBusy(false);
                    int miss = Missing().Count;
                    btnExport.Text = miss > 0 ? "Enregistrer les " + miss + " PDF manquants" : "Enregistrer les PDF manquants";
                    lblStep.Text = (interrupted ? "Interrompu après 3 erreurs de suite. " : (stop ? "Arrêté. " : "Terminé. ")) + ok + " PDF enregistré(s)" + (ko > 0 ? ", " + ko + " erreur(s)" : "") + ".";
                    lblSummary.Text = miss > 0 ? miss + " encore manquante(s)" : "Toutes les pièces du mois sont enregistrées.";
                });
            }) { IsBackground = true };
            th.Start();
        }
    }

    static class Ui2
    {
        static float factor;
        public static int Px(int px)
        {
            if (factor == 0) { try { using (Graphics g = Graphics.FromHwnd(IntPtr.Zero)) factor = g.DpiX / 96f; } catch { factor = 1; } }
            return (int)Math.Round(px * factor);
        }

        public static TableLayoutPanel Grid(int cols)
        {
            var g = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = cols, AutoSize = true, Padding = new Padding(12) };
            for (int i = 0; i < cols; i++) g.ColumnStyles.Add(new ColumnStyle(i == 1 ? SizeType.Percent : SizeType.AutoSize, 100));
            return g;
        }

        public static Label L(string text) { return new Label { Text = text, AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(3, 8, 8, 3) }; }

        public static FlowLayoutPanel OkCancel(Form f, string okText)
        {
            var p = new FlowLayoutPanel { Dock = DockStyle.Bottom, FlowDirection = FlowDirection.RightToLeft, AutoSize = true, Padding = new Padding(8) };
            var cancel = new Button { Text = "Annuler", DialogResult = DialogResult.Cancel, AutoSize = true };
            var ok = new Button { Text = okText, DialogResult = DialogResult.OK, AutoSize = true };
            p.Controls.Add(cancel); p.Controls.Add(ok);
            f.AcceptButton = ok; f.CancelButton = cancel;
            return p;
        }

        public static void Setup(Form f, string title)
        {
            f.Text = title; f.Font = new Font("Segoe UI", 9.5f); f.AutoScaleMode = AutoScaleMode.Font;
            f.FormBorderStyle = FormBorderStyle.FixedDialog; f.MaximizeBox = false; f.MinimizeBox = false;
            f.StartPosition = FormStartPosition.CenterParent; f.ShowInTaskbar = false; f.Icon = Program.AppIcon;
            f.AutoSize = true; f.AutoSizeMode = AutoSizeMode.GrowAndShrink;
        }
    }

    class DossierPicker : Form
    {
        readonly ListBox box;
        readonly List<string> paths;

        DossierPicker(List<string> found, string current)
        {
            paths = found;
            Ui2.Setup(this, "Choisir le dossier Sage 50");
            var g = Ui2.Grid(1);
            g.Controls.Add(new Label { Text = "Plusieurs dossiers Sage 50 ont été trouvés sur ce PC. Lequel SageBridge doit-il utiliser ?", AutoSize = true, MaximumSize = new Size(Ui2.Px(520), 0), Margin = new Padding(3, 3, 3, 8) });
            box = new ListBox { Width = Ui2.Px(520), Height = Ui2.Px(160), HorizontalScrollbar = true };
            foreach (string p in found) box.Items.Add(Path.GetFileNameWithoutExtension(p) + "   —   " + p);
            int idx = found.IndexOf(current);
            box.SelectedIndex = idx >= 0 ? idx : 0;
            box.DoubleClick += delegate { DialogResult = DialogResult.OK; };
            g.Controls.Add(box);
            Controls.Add(g);
            Controls.Add(Ui2.OkCancel(this, "Utiliser ce dossier"));
        }

        public static string Pick(IWin32Window owner, List<string> found, string current)
        {
            using (var f = new DossierPicker(found, current))
                return f.ShowDialog(owner) == DialogResult.OK && f.box.SelectedIndex >= 0 ? f.paths[f.box.SelectedIndex] : null;
        }
    }

    class SettingsForm : Form
    {
        readonly TextBox txtDossier, txtClient;
        readonly NumericUpDown numPort;
        readonly CheckBox chkAuto, chkRef;
        public bool AutoUpdate { get { return chkAuto.Checked; } }
        public bool WriteReference { get { return chkRef.Checked; } }
        public string Dossier { get { return txtDossier.Text.Trim(); } }
        public string Client { get { return txtClient.Text.Trim(); } }
        public int Port { get { return (int)numPort.Value; } }

        public SettingsForm(Settings cfg)
        {
            Ui2.Setup(this, "Réglages de SageBridge");
            var g = Ui2.Grid(3);
            txtDossier = new TextBox { Text = cfg.DossierPath, Width = Ui2.Px(420) };
            var browse = new Button { Text = "Parcourir…", AutoSize = true };
            browse.Click += delegate { string p = MainForm.BrowseDossier(txtDossier.Text); if (p != null) txtDossier.Text = p; };
            var find = new Button { Text = "Rechercher…", AutoSize = true };
            find.Click += delegate
            {
                List<string> found = Settings.FindDossiers();
                if (found.Count == 0) { MessageBox.Show(this, "Aucun dossier trouvé automatiquement : utilisez « Parcourir… ».", "SageBridge"); return; }
                string p = DossierPicker.Pick(this, found, txtDossier.Text);
                if (p != null) txtDossier.Text = p;
            };
            g.Controls.Add(Ui2.L("Dossier Sage 50 :"), 0, 0); g.Controls.Add(txtDossier, 1, 0);
            var fp = new FlowLayoutPanel { AutoSize = true, Margin = new Padding(0) }; fp.Controls.Add(browse); fp.Controls.Add(find);
            g.Controls.Add(fp, 2, 0);

            txtClient = new TextBox { Text = cfg.DefaultClient, Width = Ui2.Px(120) };
            g.Controls.Add(Ui2.L("Client par défaut :"), 0, 1); g.Controls.Add(txtClient, 1, 1);
            g.Controls.Add(new Label { Text = "Code client utilisé si la commande n'a pas d'adresse.", AutoSize = true, ForeColor = Color.DimGray, Anchor = AnchorStyles.Left }, 2, 1);

            numPort = new NumericUpDown { Minimum = 1024, Maximum = 65535, Value = cfg.Port, Width = Ui2.Px(90) };
            g.Controls.Add(Ui2.L("Port local :"), 0, 2); g.Controls.Add(numPort, 1, 2);
            g.Controls.Add(new Label { Text = "8765 par défaut (attendu par l'extension).", AutoSize = true, ForeColor = Color.DimGray, Anchor = AnchorStyles.Left }, 2, 2);

            chkAuto = new CheckBox { Text = "Installer automatiquement les mises à jour (vérification au démarrage puis toutes les heures)", AutoSize = true, Checked = cfg.AutoUpdate, Margin = new Padding(3, 10, 3, 3) };
            g.Controls.Add(chkAuto, 0, 3);
            g.SetColumnSpan(chkAuto, 3);
            chkRef = new CheckBox { Text = "Écrire la référence de la commande dans le champ « Référence » de la facture", AutoSize = true, Checked = cfg.WriteReference, Margin = new Padding(3, 4, 3, 3) };
            g.Controls.Add(chkRef, 0, 4);
            g.SetColumnSpan(chkRef, 3);
            g.Controls.Add(new Label { Text = "Source des mises à jour : github.com/" + cfg.Repo + "/releases", AutoSize = true, ForeColor = Color.DimGray, Margin = new Padding(3, 8, 3, 3) }, 0, 5);
            g.SetColumnSpan(g.GetControlFromPosition(0, 5), 3);
            g.Controls.Add(new Label { Text = "Fichiers de SageBridge (réglages, historique, journal) : " + Settings.Dir, AutoSize = true, ForeColor = Color.DimGray, Margin = new Padding(3, 4, 3, 3) }, 0, 6);
            g.SetColumnSpan(g.GetControlFromPosition(0, 6), 3);
            g.Controls.Add(new Label { Text = "Dossier en mode Données Mobiles ou sur un serveur : collez dans « Dossier Sage 50 » le chemin d'accès affiché par Sage\n(Dossier > Options > Informations), qui commence par !Srv!. Les fichiers « _OLD » sont d'anciennes copies : ne pas les choisir.", AutoSize = true, ForeColor = Color.DimGray, Margin = new Padding(3, 4, 3, 3) }, 0, 7);
            g.SetColumnSpan(g.GetControlFromPosition(0, 7), 3);
            Controls.Add(g);
            Controls.Add(Ui2.OkCancel(this, "Enregistrer"));
        }
    }

    class TestInvoiceForm : Form
    {
        readonly TextBox txtClient, txtArticle, txtPrice;
        public string Client { get { return txtClient.Text.Trim(); } }
        public string Article { get { return txtArticle.Text.Trim(); } }
        public double Price
        {
            get { double p; return double.TryParse(txtPrice.Text.Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out p) ? p : 0; }
        }

        public TestInvoiceForm(Settings cfg)
        {
            Ui2.Setup(this, "Facture de test");
            var g = Ui2.Grid(2);
            g.Controls.Add(new Label { Text = "Crée une facture PROVISOIRE (non validée) d'une ligne dans Sage, pour vérifier que tout fonctionne.\nVous pourrez la supprimer ensuite dans Sage.", AutoSize = true, Margin = new Padding(3, 3, 3, 10) }, 0, 0);
            g.SetColumnSpan(g.GetControlFromPosition(0, 0), 2);
            txtClient = new TextBox { Text = cfg.DefaultClient, Width = Ui2.Px(140) };
            txtArticle = new TextBox { Text = cfg.TestArticle, Width = Ui2.Px(140) };
            txtPrice = new TextBox { Text = "5,00", Width = Ui2.Px(80) };
            g.Controls.Add(Ui2.L("Code client existant :"), 0, 1); g.Controls.Add(txtClient, 1, 1);
            g.Controls.Add(Ui2.L("Code article existant :"), 0, 2); g.Controls.Add(txtArticle, 1, 2);
            g.Controls.Add(Ui2.L("Prix unitaire HT (€) :"), 0, 3); g.Controls.Add(txtPrice, 1, 3);
            Controls.Add(g);
            Controls.Add(Ui2.OkCancel(this, "Créer la facture de test"));
            FormClosing += delegate(object s, FormClosingEventArgs e)
            {
                if (DialogResult == DialogResult.OK && (Client.Length == 0 || Article.Length == 0))
                {
                    MessageBox.Show(this, "Indiquez un code client et un code article qui existent dans votre dossier Sage.", "Facture de test");
                    e.Cancel = true;
                }
            };
        }
    }
}

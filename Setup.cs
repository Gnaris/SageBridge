// ════════════════════════════════════════════════════════════════════════════
//  SageBridge-Setup.exe — installateur, mise à jour et désinstallation
// ════════════════════════════════════════════════════════════════════════════
//  Un seul exécutable, qui contient SageBridge.exe, la documentation PDF et le
//  logo (ressources incorporées par build.cmd).
//
//    SageBridge-Setup.exe                  assistant d'installation
//    SageBridge-Setup.exe /silent          installation sans questions (options par défaut ou précédentes)
//    SageBridge-Setup.exe /update /pid=N   mise à jour lancée par SageBridge : attend la fin du processus N,
//                                          remplace les fichiers et relance SageBridge
//    SageBridge-Uninstall.exe /uninstall   désinstallation (appelée par « Programmes et fonctionnalités »)
//
//  Installation par utilisateur, sans droits administrateur :
//    %LOCALAPPDATA%\Programs\SageBridge   programme, documentation, désinstalleur
//    HKCU\...\Uninstall\SageBridge        inscription dans « Programmes et fonctionnalités »
//    HKCU\...\Run\SageBridge              démarrage avec Windows (option)
//  Les données (réglages, historique, journal) restent dans %APPDATA%\SageBridge.
// ════════════════════════════════════════════════════════════════════════════

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using Microsoft.Win32;

[assembly: AssemblyTitle("Installation de SageBridge")]
[assembly: AssemblyDescription("Installateur de SageBridge (factures Sage 50 par COM)")]
[assembly: AssemblyCompany("Beli & Jolie")]
[assembly: AssemblyProduct("SageBridge")]
[assembly: AssemblyCopyright("© 2026 Beli & Jolie")]

namespace SageBridgeSetup
{
    class Options
    {
        public string Dir;
        public bool DesktopShortcut = true, StartMenu = true, Startup = true, Launch = true;
    }

    static class Setup
    {
        public const string App = "SageBridge";
        public const string Publisher = "Beli & Jolie";
        public const string RepoUrl = "https://github.com/Gnaris/SageBridge";
        public const string UninstallKey = @"Software\Microsoft\Windows\CurrentVersion\Uninstall\SageBridge";
        public const string RunKey = @"Software\Microsoft\Windows\CurrentVersion\Run";
        public const string ExeName = "SageBridge.exe";
        public const string UninstallerName = "SageBridge-Uninstall.exe";
        public const string DocName = "SageBridge - Documentation.pdf";

        static byte[] exeBytes;
        public static string Version = "?";
        public static readonly Color Navy = Color.FromArgb(16, 58, 92), Teal = Color.FromArgb(13, 148, 136), Amber = Color.FromArgb(245, 165, 36);

        [STAThread]
        static int Main(string[] args)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            bool silent = false, update = false, uninstall = false;
            int pid = 0;
            foreach (string a in args)
            {
                string l = a.ToLowerInvariant();
                if (l == "/silent" || l == "/s" || l == "/quiet") silent = true;
                else if (l == "/update") update = true;
                else if (l == "/uninstall") uninstall = true;
                else if (l.StartsWith("/pid=")) int.TryParse(l.Substring(5), out pid);
            }
            // Le désinstalleur installé (SageBridge-Uninstall.exe) ne sert qu'à désinstaller.
            if (Path.GetFileName(Application.ExecutablePath).Equals(UninstallerName, StringComparison.OrdinalIgnoreCase)) uninstall = true;
            try
            {
                if (uninstall) return Uninstaller.Run(silent);
                exeBytes = Resource("SageBridge.exe");
                if (exeBytes == null) { MessageBox.Show("Installateur incomplet : SageBridge.exe manquant.", App, MessageBoxButtons.OK, MessageBoxIcon.Error); return 2; }
                try { Version = ShortVersion(Assembly.ReflectionOnlyLoad(exeBytes).GetName().Version); } catch { }
                if (update) { Application.Run(new UpdateForm(pid)); return UpdateForm.ExitCode; }
                if (silent)
                {
                    Options o = DefaultOptions(true);
                    Install(o, delegate(int p, string s) { }, 0);
                    if (o.Launch) LaunchApp(o.Dir, null);
                    return 0;
                }
                Application.Run(new Wizard());
                return 0;
            }
            catch (Exception e)
            {
                Log("ERREUR : " + e);
                if (!silent) MessageBox.Show("Erreur : " + e.Message, App, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return 1;
            }
        }

        static int dpi;
        public static int Dpi
        {
            get
            {
                if (dpi == 0) { try { using (Graphics g = Graphics.FromHwnd(IntPtr.Zero)) dpi = (int)g.DpiX; } catch { dpi = 96; } }
                return dpi;
            }
        }

        public static string ShortVersion(Version v) { return v.Major + "." + v.Minor + "." + Math.Max(0, v.Build); }

        public static byte[] Resource(string name)
        {
            using (Stream s = Assembly.GetExecutingAssembly().GetManifestResourceStream(name))
            {
                if (s == null) return null;
                var ms = new MemoryStream();
                s.CopyTo(ms);
                return ms.ToArray();
            }
        }

        public static Image Logo(int size)
        {
            byte[] b = Resource("logo.png");
            if (b == null) return null;
            using (var src = Image.FromStream(new MemoryStream(b)))
            {
                var bmp = new Bitmap(size, size);
                using (Graphics g = Graphics.FromImage(bmp))
                {
                    g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                    g.SmoothingMode = SmoothingMode.HighQuality;
                    g.DrawImage(src, 0, 0, size, size);
                }
                return bmp;
            }
        }

        public static Icon AppIcon { get { try { return Icon.ExtractAssociatedIcon(Application.ExecutablePath); } catch { return SystemIcons.Application; } } }

        // ── Journal ──

        public static void Log(string msg)
        {
            try
            {
                string d = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), App);
                Directory.CreateDirectory(d);
                File.AppendAllText(Path.Combine(d, "installation.log"), DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss") + "  " + msg + "\r\n", Encoding.UTF8);
            }
            catch { }
        }

        // ── État actuel ──

        public static string DefaultDir
        {
            get { return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), @"Programs\SageBridge"); }
        }

        public static string RegValue(string name)
        {
            try { using (RegistryKey k = Registry.CurrentUser.OpenSubKey(UninstallKey)) return k == null ? null : Convert.ToString(k.GetValue(name)); }
            catch { return null; }
        }

        public static string InstalledDir { get { string d = RegValue("InstallLocation"); return string.IsNullOrEmpty(d) ? null : d.TrimEnd('\\'); } }
        public static string InstalledVersion { get { return RegValue("DisplayVersion"); } }

        public static bool RunKeyExists
        {
            get { try { using (RegistryKey k = Registry.CurrentUser.OpenSubKey(RunKey)) return k != null && k.GetValue(App) != null; } catch { return false; } }
        }

        // Options : celles de l'installation précédente si elle existe, sinon les valeurs par défaut.
        public static Options DefaultOptions(bool keepPrevious)
        {
            var o = new Options { Dir = InstalledDir ?? DefaultDir };
            if (keepPrevious && InstalledVersion != null)
            {
                o.DesktopShortcut = RegValue("SB_Bureau") != "0";
                o.StartMenu = RegValue("SB_Demarrer") != "0";
                o.Startup = RunKeyExists;
            }
            return o;
        }

        // Sage 50 Gestion commerciale : objet COM inscrit et exécutable présent.
        // WGC.exe est un serveur COM 32 bits : son CLSID est inscrit dans la vue 32 bits
        // du registre (WOW6432Node), invisible depuis un programme 64 bits sans RegistryView.
        public static string SageExe
        {
            get
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
                string def = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), @"Ciel\Gestion Commerciale\WGC.exe");
                return File.Exists(def) ? def : null;
            }
        }

        public static List<Process> RunningApps()
        {
            var l = new List<Process>();
            try { foreach (Process p in Process.GetProcessesByName("SageBridge")) l.Add(p); } catch { }
            return l;
        }

        // Ferme SageBridge (il est verrouillé tant qu'il tourne). waitPid : processus qui doit se terminer seul (mise à jour).
        public static void StopApp(int waitPid)
        {
            if (waitPid > 0)
            {
                try
                {
                    using (Process p = Process.GetProcessById(waitPid))
                    {
                        Log("Attente de la fermeture de SageBridge (PID " + waitPid + ")…");
                        if (!p.WaitForExit(15000)) { Log("SageBridge ne s'est pas fermé : arrêt forcé."); p.Kill(); p.WaitForExit(5000); }
                    }
                }
                catch { }
            }
            foreach (Process p in RunningApps())
            {
                try { Log("Fermeture de SageBridge (PID " + p.Id + ")."); p.Kill(); p.WaitForExit(5000); } catch { }
                finally { p.Dispose(); }
            }
        }

        static void WriteFile(string path, byte[] data)
        {
            for (int i = 0; ; i++)
            {
                try { File.WriteAllBytes(path, data); return; }
                catch (IOException) { if (i >= 20) throw; Thread.Sleep(500); }
                catch (UnauthorizedAccessException) { if (i >= 20) throw; Thread.Sleep(500); }
            }
        }

        static string ShortcutPath(Environment.SpecialFolder folder)
        {
            return Path.Combine(Environment.GetFolderPath(folder), App + ".lnk");
        }

        static void CreateShortcut(string lnk, string target, string description)
        {
            Type t = Type.GetTypeFromProgID("WScript.Shell");
            dynamic shell = Activator.CreateInstance(t);
            try
            {
                dynamic sc = shell.CreateShortcut(lnk);
                sc.TargetPath = target;
                sc.WorkingDirectory = Path.GetDirectoryName(target);
                sc.IconLocation = target + ",0";
                sc.Description = description;
                sc.Save();
            }
            finally { System.Runtime.InteropServices.Marshal.FinalReleaseComObject(shell); }
        }

        static void DeleteQuiet(string path) { try { if (File.Exists(path)) File.Delete(path); } catch { } }

        // ── Installation ──

        public static void Install(Options o, Action<int, string> progress, int waitPid)
        {
            string previous = InstalledVersion;
            Log("Installation de SageBridge " + Version + " dans " + o.Dir + (previous != null ? " (version précédente : " + previous + ")" : ""));

            progress(5, "Fermeture de SageBridge s'il est ouvert…");
            StopApp(waitPid);

            progress(20, "Copie du programme…");
            Directory.CreateDirectory(o.Dir);
            string exe = Path.Combine(o.Dir, ExeName);
            WriteFile(exe, exeBytes);

            progress(40, "Copie de la documentation…");
            byte[] doc = Resource("Documentation.pdf");
            if (doc != null) WriteFile(Path.Combine(o.Dir, DocName), doc);

            progress(50, "Copie du désinstalleur…");
            string self = Application.ExecutablePath;
            string uninst = Path.Combine(o.Dir, UninstallerName);
            if (!string.Equals(Path.GetFullPath(self), Path.GetFullPath(uninst), StringComparison.OrdinalIgnoreCase))
                WriteFile(uninst, File.ReadAllBytes(self));

            progress(60, "Création des raccourcis…");
            string desktopLnk = ShortcutPath(Environment.SpecialFolder.DesktopDirectory);
            string menuLnk = ShortcutPath(Environment.SpecialFolder.Programs);
            if (o.DesktopShortcut) CreateShortcut(desktopLnk, exe, "Création de factures Sage 50"); else DeleteQuiet(desktopLnk);
            if (o.StartMenu) CreateShortcut(menuLnk, exe, "Création de factures Sage 50"); else DeleteQuiet(menuLnk);

            progress(75, "Inscription dans « Programmes et fonctionnalités »…");
            long size = 0;
            foreach (string f in Directory.GetFiles(o.Dir)) size += new FileInfo(f).Length;
            using (RegistryKey k = Registry.CurrentUser.CreateSubKey(UninstallKey))
            {
                k.SetValue("DisplayName", App);
                k.SetValue("DisplayVersion", Version);
                k.SetValue("Publisher", Publisher);
                k.SetValue("DisplayIcon", "\"" + exe + "\",0");
                k.SetValue("InstallLocation", o.Dir);
                k.SetValue("UninstallString", "\"" + uninst + "\" /uninstall");
                k.SetValue("QuietUninstallString", "\"" + uninst + "\" /uninstall /silent");
                k.SetValue("URLInfoAbout", RepoUrl);
                k.SetValue("HelpLink", RepoUrl + "#readme");
                k.SetValue("URLUpdateInfo", RepoUrl + "/releases");
                k.SetValue("Comments", "Création de factures Sage 50 Gestion commerciale par l'interface COM");
                k.SetValue("InstallDate", DateTime.Now.ToString("yyyyMMdd"));
                k.SetValue("EstimatedSize", (int)Math.Max(1, size / 1024), RegistryValueKind.DWord);
                k.SetValue("NoModify", 1, RegistryValueKind.DWord);
                k.SetValue("NoRepair", 1, RegistryValueKind.DWord);
                k.SetValue("SB_Bureau", o.DesktopShortcut ? "1" : "0");
                k.SetValue("SB_Demarrer", o.StartMenu ? "1" : "0");
            }

            progress(85, "Démarrage avec Windows…");
            using (RegistryKey k = Registry.CurrentUser.CreateSubKey(RunKey))
            {
                if (o.Startup) k.SetValue(App, "\"" + exe + "\"");
                else k.DeleteValue(App, false);
            }

            progress(100, "Installation terminée.");
            Log("Installation terminée (" + Version + ").");
        }

        public static void LaunchApp(string dir, string previousVersion)
        {
            string exe = Path.Combine(dir, ExeName);
            var psi = new ProcessStartInfo(exe) { WorkingDirectory = dir, UseShellExecute = true };
            if (!string.IsNullOrEmpty(previousVersion)) psi.Arguments = "/apres-maj=" + previousVersion;
            Process.Start(psi);
        }
    }

    // ───────────────────────── Habillage commun ─────────────────────────

    class Banner : Panel
    {
        readonly Image logo;
        public string Title = "SageBridge", Subtitle = "";

        public Banner(int logoSize)
        {
            DoubleBuffered = true;
            logo = Setup.Logo(logoSize);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.SmoothingMode = SmoothingMode.HighQuality;
            g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;
            using (var br = new LinearGradientBrush(ClientRectangle, Setup.Navy, Setup.Teal, LinearGradientMode.ForwardDiagonal))
                g.FillRectangle(br, ClientRectangle);
            bool vertical = Height > Width;
            int pad = (int)(16 * Setup.Dpi / 96f);
            if (vertical)
            {
                if (logo != null) g.DrawImage(logo, (Width - logo.Width) / 2, pad * 2, logo.Width, logo.Height);
                int y = pad * 2 + (logo != null ? logo.Height : 0) + pad;
                using (var f = new Font("Segoe UI Semibold", 16f))
                using (var sf = new StringFormat { Alignment = StringAlignment.Center })
                {
                    g.DrawString(Title, f, Brushes.White, new RectangleF(0, y, Width, 40 * Setup.Dpi / 96f), sf);
                    using (var f2 = new Font("Segoe UI", 9.5f))
                    using (var b2 = new SolidBrush(Color.FromArgb(220, 255, 255, 255)))
                        g.DrawString(Subtitle, f2, b2, new RectangleF(pad, y + 36 * Setup.Dpi / 96f, Width - 2 * pad, 120 * Setup.Dpi / 96f), sf);
                }
            }
            else
            {
                int x = pad;
                if (logo != null) { g.DrawImage(logo, x, (Height - logo.Height) / 2, logo.Width, logo.Height); x += logo.Width + pad; }
                using (var f = new Font("Segoe UI Semibold", 13f))
                    g.DrawString(Title, f, Brushes.White, x, Height / 2f - 26 * Setup.Dpi / 96f);
                using (var f2 = new Font("Segoe UI", 9.5f))
                using (var b2 = new SolidBrush(Color.FromArgb(225, 255, 255, 255)))
                    g.DrawString(Subtitle, f2, b2, x, Height / 2f + 2 * Setup.Dpi / 96f);
            }
        }
    }

    // ───────────────────────── Assistant d'installation ─────────────────────────

    class Wizard : Form
    {
        readonly Panel content;
        readonly Button btnBack, btnNext, btnCancel;
        readonly Panel[] pages = new Panel[4];
        int page;
        bool installing, done;

        TextBox txtDir;
        CheckBox chkDesktop, chkMenu, chkStartup, chkLaunch;
        ProgressBar bar;
        Label lblStep;
        ListBox lstSteps;
        Label lblDone;
        Options opts;

        int S(int px) { return (int)Math.Round(px * Setup.Dpi / 96f); }

        public Wizard()
        {
            Text = "Installation de SageBridge " + Setup.Version;
            Icon = Setup.AppIcon;
            Font = new Font("Segoe UI", 9.5f);
            AutoScaleMode = AutoScaleMode.None;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            StartPosition = FormStartPosition.CenterScreen;
            ClientSize = new Size(S(700), S(460));
            opts = Setup.DefaultOptions(true);

            var banner = new Banner(S(112)) { Dock = DockStyle.Left, Width = S(210), Subtitle = "Version " + Setup.Version + "\nFactures Sage 50\nen un clic" };
            var bottom = new Panel { Dock = DockStyle.Bottom, Height = S(52) };
            bottom.Paint += delegate(object s, PaintEventArgs e) { e.Graphics.DrawLine(SystemPens.ControlLight, 0, 0, bottom.Width, 0); };
            btnCancel = new Button { Text = "Annuler", Size = new Size(S(96), S(30)), Anchor = AnchorStyles.Right | AnchorStyles.Top };
            btnNext = new Button { Text = "Suivant >", Size = new Size(S(110), S(30)), Anchor = AnchorStyles.Right | AnchorStyles.Top };
            btnBack = new Button { Text = "< Précédent", Size = new Size(S(110), S(30)), Anchor = AnchorStyles.Right | AnchorStyles.Top };
            bottom.Controls.Add(btnCancel); bottom.Controls.Add(btnNext); bottom.Controls.Add(btnBack);
            bottom.Layout += delegate
            {
                int y = S(11), x = bottom.ClientSize.Width - S(14);
                btnCancel.Location = new Point(x - btnCancel.Width, y); x -= btnCancel.Width + S(14);
                btnNext.Location = new Point(x - btnNext.Width, y); x -= btnNext.Width + S(6);
                btnBack.Location = new Point(x - btnBack.Width, y);
            };
            btnCancel.Click += delegate { Close(); };
            btnBack.Click += delegate { ShowPage(page - 1); };
            btnNext.Click += delegate { Next(); };
            AcceptButton = btnNext;

            content = new Panel { Dock = DockStyle.Fill, Padding = new Padding(S(24), S(20), S(24), S(10)), BackColor = Color.White };
            Controls.Add(content);
            Controls.Add(bottom);
            Controls.Add(banner);

            pages[0] = WelcomePage();
            pages[1] = OptionsPage();
            pages[2] = ProgressPage();
            pages[3] = DonePage();
            foreach (Panel p in pages) { p.Dock = DockStyle.Fill; p.Visible = false; content.Controls.Add(p); }
            ShowPage(0);
        }

        Label Title(string text) { return new Label { Text = text, AutoSize = true, Font = new Font("Segoe UI Semibold", 14f), ForeColor = Setup.Navy, Margin = new Padding(0, 0, 0, S(12)) }; }

        Label Para(string text) { return new Label { Text = text, AutoSize = true, MaximumSize = new Size(S(440), 0), Margin = new Padding(0, 0, 0, S(10)) }; }

        FlowLayoutPanel Column()
        {
            return new FlowLayoutPanel { FlowDirection = FlowDirection.TopDown, WrapContents = false, AutoScroll = true };
        }

        Panel WelcomePage()
        {
            var p = Column();
            p.Controls.Add(Title("Bienvenue"));
            p.Controls.Add(Para("Cet assistant installe SageBridge " + Setup.Version + " sur ce PC.\n\nSageBridge crée vos factures directement dans Sage 50 Gestion commerciale (en provisoire, à valider dans Sage), à partir des commandes envoyées par l'extension Chrome « Regroupement facture » ou par tout autre outil. Il retrouve ou crée automatiquement la fiche client et applique le bon régime de TVA (France, Union européenne, hors Union européenne)."));

            string sage = Setup.SageExe;
            var lblSage = Para(sage != null
                ? "✔  Sage 50 Gestion commerciale est installé sur ce PC."
                : "⚠  Sage 50 Gestion commerciale n'a pas été trouvé sur ce PC. SageBridge peut être installé, mais il ne pourra créer des factures qu'une fois Sage 50 installé.");
            lblSage.ForeColor = sage != null ? Color.FromArgb(10, 122, 69) : Color.FromArgb(180, 90, 0);
            lblSage.Font = new Font("Segoe UI Semibold", 9.5f);
            p.Controls.Add(lblSage);

            string prev = Setup.InstalledVersion;
            if (prev != null) p.Controls.Add(Para("SageBridge " + prev + " est déjà installé : il sera remplacé par la version " + Setup.Version + ". Vos réglages et l'historique des factures sont conservés."));
            if (Setup.RunningApps().Count > 0) p.Controls.Add(Para("SageBridge est ouvert : il sera fermé pendant l'installation puis relancé."));
            p.Controls.Add(Para("Installation pour votre compte Windows uniquement, sans droits administrateur."));
            return p;
        }

        Panel OptionsPage()
        {
            var p = Column();
            p.Controls.Add(Title("Options d'installation"));
            p.Controls.Add(new Label { Text = "Dossier d'installation :", AutoSize = true, Margin = new Padding(0, 0, 0, S(4)) });
            var row = new FlowLayoutPanel { AutoSize = true, WrapContents = false, Margin = new Padding(0, 0, 0, S(14)) };
            txtDir = new TextBox { Text = opts.Dir, Width = S(340) };
            var browse = new Button { Text = "Parcourir…", AutoSize = true, Margin = new Padding(S(6), 0, 0, 0) };
            browse.Click += delegate
            {
                using (var d = new FolderBrowserDialog { Description = "Dossier d'installation de SageBridge", SelectedPath = Directory.Exists(txtDir.Text) ? txtDir.Text : "" })
                    if (d.ShowDialog(this) == DialogResult.OK)
                        txtDir.Text = Path.GetFileName(d.SelectedPath).Equals("SageBridge", StringComparison.OrdinalIgnoreCase) ? d.SelectedPath : Path.Combine(d.SelectedPath, "SageBridge");
            };
            row.Controls.Add(txtDir); row.Controls.Add(browse);
            p.Controls.Add(row);
            chkDesktop = new CheckBox { Text = "Créer un raccourci sur le Bureau", AutoSize = true, Checked = opts.DesktopShortcut };
            chkMenu = new CheckBox { Text = "Créer un raccourci dans le menu Démarrer", AutoSize = true, Checked = opts.StartMenu };
            chkStartup = new CheckBox { Text = "Démarrer SageBridge avec Windows (recommandé)", AutoSize = true, Checked = opts.Startup };
            chkLaunch = new CheckBox { Text = "Lancer SageBridge à la fin de l'installation", AutoSize = true, Checked = true };
            foreach (CheckBox c in new[] { chkDesktop, chkMenu, chkStartup, chkLaunch }) { c.Margin = new Padding(0, S(3), 0, S(3)); p.Controls.Add(c); }
            var note = Para("\nMises à jour : SageBridge vérifie les nouvelles versions au démarrage puis toutes les heures, et les installe automatiquement (réglable dans SageBridge > Réglages).");
            note.ForeColor = Color.DimGray;
            p.Controls.Add(note);
            return p;
        }

        Panel ProgressPage()
        {
            var p = Column();
            p.Controls.Add(Title("Installation en cours"));
            lblStep = new Label { Text = "Préparation…", AutoSize = true, Margin = new Padding(0, 0, 0, S(8)) };
            bar = new ProgressBar { Width = S(440), Height = S(20), Minimum = 0, Maximum = 100, Margin = new Padding(0, 0, 0, S(12)) };
            lstSteps = new ListBox { Width = S(440), Height = S(200), BorderStyle = BorderStyle.FixedSingle, IntegralHeight = false };
            p.Controls.Add(lblStep); p.Controls.Add(bar); p.Controls.Add(lstSteps);
            return p;
        }

        Panel DonePage()
        {
            var p = Column();
            p.Controls.Add(Title("SageBridge est installé"));
            lblDone = Para("");
            p.Controls.Add(lblDone);
            var check = Para(
                "Avant la première facture :\n\n" +
                "1.  Ouvrez Sage 50 Gestion commerciale sur votre dossier (Sage doit rester ouvert pendant l'utilisation).\n\n" +
                "2.  Autorisez l'accès COM dans Sage : menu Utilitaires > Mode avancé > Divers > Autorisation accès COM (autorisation liée au compte Sage, à renouveler environ tous les 6 mois).\n\n" +
                "3.  Dans SageBridge, cliquez sur « Réglages… » pour vérifier le dossier Sage, puis « Facture de test… » avec un client et un article qui existent dans votre dossier.\n\n" +
                "La documentation complète est disponible depuis le bouton « Documentation » de SageBridge.");
            p.Controls.Add(check);
            return p;
        }

        void ShowPage(int i)
        {
            page = i;
            for (int k = 0; k < pages.Length; k++) pages[k].Visible = k == i;
            btnBack.Visible = i == 1;
            btnNext.Text = i == 1 ? "Installer" : (i == 3 ? "Terminer" : "Suivant >");
            btnNext.Enabled = i != 2;
            btnCancel.Visible = i < 2;
            // Le bouton principal garde le focus : Entrée fait toujours « Suivant / Installer / Terminer ».
            if (btnNext.Enabled) ActiveControl = btnNext;
        }

        void Next()
        {
            if (page == 0) { ShowPage(1); return; }
            if (page == 3) { done = true; Close(); return; }
            if (page != 1) return;

            string dir = txtDir.Text.Trim();
            try { dir = Path.GetFullPath(dir).TrimEnd('\\'); }
            catch { MessageBox.Show(this, "Dossier d'installation invalide.", Setup.App); return; }
            opts.Dir = dir;
            opts.DesktopShortcut = chkDesktop.Checked; opts.StartMenu = chkMenu.Checked;
            opts.Startup = chkStartup.Checked; opts.Launch = chkLaunch.Checked;
            ShowPage(2);
            installing = true;
            string previous = Setup.InstalledVersion;
            var th = new Thread(delegate()
            {
                Exception err = null;
                try
                {
                    Setup.Install(opts, delegate(int pct, string step)
                    {
                        BeginInvoke((Action)delegate { bar.Value = pct; lblStep.Text = step; lstSteps.Items.Add(step); lstSteps.TopIndex = lstSteps.Items.Count - 1; });
                    }, 0);
                }
                catch (Exception e) { err = e; Setup.Log("ERREUR installation : " + e); }
                BeginInvoke((Action)delegate
                {
                    installing = false;
                    if (err != null)
                    {
                        lblStep.Text = "L'installation a échoué.";
                        lstSteps.Items.Add("Erreur : " + err.Message);
                        MessageBox.Show(this, "L'installation a échoué :\n\n" + err.Message, Setup.App, MessageBoxButtons.OK, MessageBoxIcon.Error);
                        btnCancel.Visible = true; btnCancel.Text = "Fermer";
                        return;
                    }
                    lblDone.Text = (previous != null && previous != Setup.Version ? "SageBridge a été mis à jour (" + previous + " → " + Setup.Version + ")." : "SageBridge " + Setup.Version + " a été installé.")
                        + "\nDossier : " + opts.Dir + "\nIl apparaît dans Paramètres > Applications (« Programmes et fonctionnalités ») pour la désinstallation.";
                    ShowPage(3);
                    if (opts.Launch) { try { Setup.LaunchApp(opts.Dir, null); } catch (Exception e) { MessageBox.Show(this, "Impossible de lancer SageBridge : " + e.Message, Setup.App); } }
                });
            });
            th.IsBackground = true;
            th.Start();
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (installing) { e.Cancel = true; return; }
            if (!done && page < 2 && e.CloseReason == CloseReason.UserClosing)
            {
                if (MessageBox.Show(this, "Abandonner l'installation de SageBridge ?", Setup.App, MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) e.Cancel = true;
            }
            base.OnFormClosing(e);
        }
    }

    // ───────────────────────── Mise à jour (lancée par SageBridge) ─────────────────────────

    class UpdateForm : Form
    {
        public static int ExitCode;
        readonly int pid;
        readonly Label lbl;

        public UpdateForm(int waitPid)
        {
            pid = waitPid;
            Text = "Mise à jour de SageBridge";
            Icon = Setup.AppIcon;
            Font = new Font("Segoe UI", 9.5f);
            AutoScaleMode = AutoScaleMode.None;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false; MinimizeBox = false; ControlBox = false;
            StartPosition = FormStartPosition.CenterScreen;
            int s96 = Setup.Dpi;
            ClientSize = new Size(460 * s96 / 96, 170 * s96 / 96);
            var banner = new Banner(56 * s96 / 96) { Dock = DockStyle.Top, Height = 80 * s96 / 96, Title = "Mise à jour de SageBridge", Subtitle = "Installation de la version " + Setup.Version + "…" };
            lbl = new Label { Text = "Fermeture de SageBridge…", AutoSize = false, Dock = DockStyle.Top, Height = 34 * s96 / 96, Padding = new Padding(16 * s96 / 96, 10 * s96 / 96, 16 * s96 / 96, 0) };
            var bar = new ProgressBar { Style = ProgressBarStyle.Marquee, Dock = DockStyle.Top, Height = 18 * s96 / 96, MarqueeAnimationSpeed = 25 };
            var barHost = new Panel { Dock = DockStyle.Top, Height = 30 * s96 / 96, Padding = new Padding(16 * s96 / 96, 4 * s96 / 96, 16 * s96 / 96, 4 * s96 / 96) };
            barHost.Controls.Add(bar);
            Controls.Add(barHost); Controls.Add(lbl); Controls.Add(banner);
        }

        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);
            string previous = Setup.InstalledVersion;
            var th = new Thread(delegate()
            {
                Options o = Setup.DefaultOptions(true);
                try
                {
                    Setup.Install(o, delegate(int p, string s) { BeginInvoke((Action)delegate { lbl.Text = s; }); }, pid);
                    BeginInvoke((Action)delegate { lbl.Text = "Redémarrage de SageBridge…"; });
                    Setup.LaunchApp(o.Dir, previous);
                    ExitCode = 0;
                }
                catch (Exception ex)
                {
                    ExitCode = 1;
                    Setup.Log("ERREUR mise à jour : " + ex);
                    BeginInvoke((Action)delegate { MessageBox.Show(this, "La mise à jour de SageBridge a échoué :\n\n" + ex.Message + "\n\nL'ancienne version est relancée.", Setup.App, MessageBoxButtons.OK, MessageBoxIcon.Warning); });
                    try { Setup.LaunchApp(o.Dir, null); } catch { }
                }
                Thread.Sleep(1200);
                BeginInvoke((Action)delegate { Close(); });
            });
            th.IsBackground = true;
            th.Start();
        }
    }

    // ───────────────────────── Désinstallation ─────────────────────────

    static class Uninstaller
    {
        public static int Run(bool silent)
        {
            string dir = Setup.InstalledDir ?? Path.GetDirectoryName(Application.ExecutablePath);
            string data = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), Setup.App);
            bool deleteData = false;
            if (!silent)
            {
                using (var f = new UninstallForm(data))
                {
                    if (f.ShowDialog() != DialogResult.OK) return 1;
                    deleteData = f.DeleteData;
                }
            }
            Setup.Log("Désinstallation de SageBridge " + (Setup.InstalledVersion ?? "") + " (" + dir + ")" + (deleteData ? ", données supprimées" : ""));
            Setup.StopApp(0);

            foreach (string lnk in new[] {
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory), "SageBridge.lnk"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Programs), "SageBridge.lnk") })
                try { if (File.Exists(lnk)) File.Delete(lnk); } catch { }
            try
            {
                using (RegistryKey k = Registry.CurrentUser.OpenSubKey(Setup.RunKey, true))
                    if (k != null) k.DeleteValue(Setup.App, false);
            }
            catch { }
            try { Registry.CurrentUser.DeleteSubKeyTree(Setup.UninstallKey, false); } catch { }

            string self = Application.ExecutablePath;
            if (Directory.Exists(dir))
            {
                foreach (string f in new[] { Setup.ExeName, Setup.DocName })
                    try { File.Delete(Path.Combine(dir, f)); } catch { }
            }
            if (deleteData) { try { Directory.Delete(data, true); } catch { } }

            // Le désinstalleur ne peut pas se supprimer lui-même pendant qu'il tourne :
            // une commande différée efface le fichier puis le dossier s'il est vide.
            string selfDir = Path.GetDirectoryName(self);
            if (string.Equals(selfDir.TrimEnd('\\'), dir.TrimEnd('\\'), StringComparison.OrdinalIgnoreCase))
            {
                string cmd = "/c ping 127.0.0.1 -n 3 > nul & del /f /q \"" + self + "\" & rmdir \"" + dir + "\"";
                try { Process.Start(new ProcessStartInfo("cmd.exe", cmd) { CreateNoWindow = true, UseShellExecute = false, WindowStyle = ProcessWindowStyle.Hidden }); } catch { }
            }
            if (!silent) MessageBox.Show("SageBridge a été désinstallé." + (deleteData ? "" : "\n\nVos réglages et l'historique des factures sont conservés dans :\n" + data), "Désinstallation de SageBridge", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return 0;
        }
    }

    class UninstallForm : Form
    {
        readonly CheckBox chk;
        public bool DeleteData { get { return chk.Checked; } }

        public UninstallForm(string dataDir)
        {
            Text = "Désinstallation de SageBridge";
            Icon = Setup.AppIcon;
            Font = new Font("Segoe UI", 9.5f);
            AutoScaleMode = AutoScaleMode.None;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false; MinimizeBox = false;
            StartPosition = FormStartPosition.CenterScreen;
            int d = Setup.Dpi;
            ClientSize = new Size(480 * d / 96, 250 * d / 96);
            var banner = new Banner(56 * d / 96) { Dock = DockStyle.Top, Height = 80 * d / 96, Title = "Désinstaller SageBridge", Subtitle = "Version " + (Setup.InstalledVersion ?? "?") };
            var body = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, WrapContents = false, Padding = new Padding(16 * d / 96, 12 * d / 96, 16 * d / 96, 0) };
            body.Controls.Add(new Label { Text = "SageBridge va être fermé puis supprimé de ce PC (programme, raccourcis, démarrage automatique). Les factures déjà créées dans Sage ne sont pas touchées.", AutoSize = true, MaximumSize = new Size(440 * d / 96, 0), Margin = new Padding(0, 0, 0, 10 * d / 96) });
            chk = new CheckBox { Text = "Supprimer aussi mes données SageBridge (réglages, historique, journal)", AutoSize = true };
            body.Controls.Add(chk);
            var bottom = new FlowLayoutPanel { Dock = DockStyle.Bottom, FlowDirection = FlowDirection.RightToLeft, AutoSize = true, Padding = new Padding(8 * d / 96) };
            var cancel = new Button { Text = "Annuler", DialogResult = DialogResult.Cancel, AutoSize = true };
            var ok = new Button { Text = "Désinstaller", DialogResult = DialogResult.OK, AutoSize = true };
            bottom.Controls.Add(cancel); bottom.Controls.Add(ok);
            AcceptButton = ok; CancelButton = cancel;
            Controls.Add(body); Controls.Add(bottom); Controls.Add(banner);
        }
    }
}

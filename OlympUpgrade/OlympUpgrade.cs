using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Management;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;

namespace OlympUpgrade
{
    public partial class OlympUpgrade : Form
    {
        /// <summary>
        /// v akom kroku instalacie je formular - sluzi na to co maju tlacidla robit
        /// </summary>
        private int _Stav;

        private bool _BoliChyby;

        private double _PocetNaTik;
        private long _AktTik;

        private List<string> _Chybne = new List<string>();
        private List<string> _Spravne = new List<string>();
        private List<string> _Preskocene = new List<string>();
        private List<string> _PomCol = new List<string>();

        private Encoding _zipEncoding = Encoding.GetEncoding(437); //Z Creatora -> DotNetZip (Ionic.Zip) v1.9.1.8 default IBM437. Funguje aj s 852/850

        private DialogResult _resVysledok_Activated = DialogResult.Yes;

        public OlympUpgrade()
        {
            InitializeComponent();
        }

        #region Event handlers

        private void OlympUpgrade_Load(object sender, EventArgs e)
        {
            try
            {
                int i;
                List<string> poleArg;
                string pom = string.Empty;

                HideTabs();
                Application.DoEvents();

                tabControl1.SelectTab(0);
                _Stav = 0;

                Declare.AKT_ADRESAR = AppDomain.CurrentDomain.BaseDirectory;

                Declare.RESTARTUJ = false;
                Declare.KOPIRUJ_LICENCIU = 0;

                NastavRok();
                Application.DoEvents();

                // zistime argument, s ktorym bol program spustany

                poleArg = Environment.GetCommandLineArgs().ToList();
                poleArg.RemoveAt(0);
                //return;
                // prvy je parameter, potom je restartuj (true, false) a potom nasleduje instalacny adresar
                if (poleArg.Count - 1 >= 3)
                {
                    Declare.COMMAND_LINE_ARGUMENT = poleArg[0];

                    if (poleArg[1].ToUpper() == "TRUE")
                        Declare.RESTARTUJ = true;


                    int.TryParse(poleArg[2], out Declare.KOPIRUJ_LICENCIU);

                    pom = string.Empty;
                    for (i = 3; i < poleArg.Count; i++)
                        pom = pom = pom + poleArg[i] + " ";

                    pom = pom.TrimEnd();
                }

                // MsgBox "zacinam " & COMMAND_LINE_ARGUMENT
                // sustil installshield
                if (!string.IsNullOrEmpty(Declare.COMMAND_LINE_ARGUMENT)
                    && Declare.COMMAND_LINE_ARGUMENT.ToUpper() == Declare.PARAMETER_INSTAL.ToUpper())
                {
                    Declare.TTITLE = "Sprievodca inštaláciou OLYMP";
                    Declare.KTO_VOLAL = Declare.VOLAL_INSTALL;
                    RegistryFunctions.OdstranZRegistrovDoinstalovanie();
                    Declare.DEST_PATH = HelpFunctions.PridajLomitko(pom);

                    tabControl1.SelectTab(1);
                    Application.DoEvents();
                    lblUpgrade.Text = DajVerziuUpgrade();
                }
                else if (!string.IsNullOrEmpty(Declare.COMMAND_LINE_ARGUMENT)
                    && (Declare.COMMAND_LINE_ARGUMENT.ToUpper() == Declare.SUBOR_STARA_INSTALACIA_EXE.ToUpper()
                        || Declare.COMMAND_LINE_ARGUMENT.ToUpper() == Declare.SUBOR_EXE_STARY.ToUpper()))
                {
                    // spustil uzivatel
                    Declare.TTITLE = "OlympUpgrade";
                    Declare.KTO_VOLAL = Declare.VOLAL_OLYMP;
                    btnChangeDir.Enabled = false;      // ak to spustal OLYMP, tak nema co menit adresar
                    Application.DoEvents();
                    NastavCestuVerzie(HelpFunctions.PridajLomitko(pom));
                }
                else
                {
                    // spustil uzivatel
                    Declare.TTITLE = "OlympUpgrade";
                    Declare.KTO_VOLAL = Declare.VOLAL_UZIVATEL;
                    Application.DoEvents();
                    NastavCestuVerzie();
                }

                this.Text = "OLYMP " + lblUpgrade.Text + " - Sprievodca inštaláciou";
            }
            catch (Exception ex)
            {
                Declare.Errors.Add(ex.ToString());
                MessageBox.Show(ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void OlympUpgrade_Activated(object sender, EventArgs e)
        {
            try
            {
                // Skontroluj, kto zavolal funkciu (inštalácia)
                if (Declare.KTO_VOLAL == Declare.VOLAL_INSTALL)
                {
                    //aby sa dal dopredu, ale dal sa posunut aj dozadu
                    //SetWindowPos(this.Handle, (IntPtr)HWND_TOP, 0, 0, 0, 0, SWP_SHOWWINDOW);
                    this.TopMost = true;
                    Application.DoEvents();
                    this.BringToFront();
                    Application.DoEvents();
                    this.TopMost = false;


                    while (_resVysledok_Activated == DialogResult.Yes)
                    {
                        btnOk.Visible = false;
                        btnStorno.Enabled = false;

                        if (HelpFunctions.JeSpustenyProgram(Declare.DEST_PATH, Declare.SUBOR_EXE))
                        {
                            _resVysledok_Activated = MessageBox.Show(
                                $"Nie je možné inštalovať UPGRADE, pretože program OLYMP v adresári {Declare.DEST_PATH} je práve spustený." +
                                "\n\nMusíte najskôr ukončiť spustený program OLYMP." +
                                "\nSkúsiť znovu?",
                                "Inštalácia",
                                MessageBoxButtons.YesNo,
                                MessageBoxIcon.Exclamation);

                            if (_resVysledok_Activated == DialogResult.No)
                            {
                                this.Close(); //UkonciProgram();
                            }
                        }
                        else
                        {
                            Instaluj();
                            _resVysledok_Activated = DialogResult.No;
                        }
                    }
                }

                // Manipulácia s tlačidlami na formulári
                if (btnOk.Visible && btnOk.Enabled && btnOk.Text == "Inštaluj")
                {
                    // Očakávaná logika pri zobrazení tlačidla "Inštaluj"
                }
                else if (!btnOk.Enabled)
                {
                    btnChangeDir.Focus();
                }
                else if (!btnOk.Visible)
                {
                    // Iná logika, ak tlačidlo nie je viditeľné
                }
                else if (btnOk.Text == "Uložiť")
                {
                    btnOk.Focus();
                }
                else
                {
                    btnOk.Focus();
                }
            }
            catch (Exception ex)
            {
                Declare.Errors.Add(ex.ToString());
                MessageBox.Show(ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnDalej_Click(object sender, EventArgs e)
        {
            try
            {
                ZobrazVysledok();
            }
            catch (Exception ex)
            {
                Declare.Errors.Add(ex.ToString());
                MessageBox.Show(ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            int P_CRV2_MAJOR, P_CRV2_MINOR, P_CRV2_REVISION;
            bool splna = true;

            try
            {
                //spusti olymp
                if (_Stav == 2)
                {
                    var psi = new ProcessStartInfo
                    {
                        FileName = Path.Combine(Declare.DEST_PATH, Declare.SUBOR_EXE),
                        WorkingDirectory = Declare.DEST_PATH,
                        UseShellExecute = true,
                        Verb = "open"
                    };
                    Process.Start(psi);
                    this.Close();
                    return;
                }

                // ulozi vysledok instalacie
                if (_Stav == 1)
                {
                    UlozitVysledok();
                    btnStorno?.Focus();
                    return;
                }



                // zisti verziu win a daj varovanie ak je to XP alebo Server 2003                
                Version v = Environment.OSVersion.Version;
                if (v.Major == 5 && (v.Minor == 1 || v.Minor == 2))
                {
                    MessageBox.Show(
                        "Pre inštaláciu je požadovaný operačný systém Windows Vista alebo novší!" +
                        Environment.NewLine + Environment.NewLine +
                        "Tip:" + Environment.NewLine +
                        "Informácie o odporúčanej konfigurácii PC nájdete na " + Environment.NewLine +
                        "https://www.kros.sk/podpora/mzdy-a-personalistika/odporucana-konfiguracia/",
                        "Setup",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    return;
                }

                // zisti ci je adresar vhodny na instalovanie
                if (!HelpFunctions.JeAdresarSpravny(Declare.DEST_PATH))
                {
                    MessageBox.Show(
                        "Zadaný adresár " + Declare.DEST_PATH + " neexistuje, alebo nemáte právo do neho zapisovať.",
                        "Setup",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    return;
                }

                // toto znamena, ze v zadanom adresari nie je alfa32.exe -> Olymp.exe??? (SUBOR_EXE)
                if (Declare.P_MAJOR == 0)
                {
                    var res = MessageBox.Show(
                        "V zadanom adresári " + Declare.DEST_PATH + " sa nenachádza súbor " + Declare.SUBOR_EXE + "." +
                        Environment.NewLine + Environment.NewLine +
                        "Chcete aj napriek tomu inštalovať komponenty programu OLYMP do zadaného adresára?",
                        Declare.TTITLE,
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Question,
                        MessageBoxDefaultButton.Button2);

                    if (res == DialogResult.No) splna = false;
                }
                else
                {
                    // nie je splnena poziadavka minimalnej verzie
                    if (HelpFunctions.JePrvaVerziaStarsia(Declare.P_MAJOR, Declare.P_MINOR, Declare.P_REVISION, Declare.MIN_MAJOR, Declare.MIN_MINOR, Declare.MIN_REVISION) == -1)
                    {
                        MessageBox.Show(
                            "V zadanom adresári " + Declare.DEST_PATH + " nemáte nainštalovanú minimálnu požadovanú verziu programu OLYMP. " +
                            Environment.NewLine + Environment.NewLine +
                            "Pre duplicitnú inštaláciu programu do viacerých priečinkov použite DVD, alebo OlympInstall.exe.",
                            Declare.TTITLE,
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Exclamation);
                        splna = false;
                    }

                    // upgrade je starsi ako nainstalovana verzia
                    if (splna && HelpFunctions.JePrvaVerziaStarsia(Declare.MAJOR, Declare.MINOR, Declare.REVISION, Declare.P_MAJOR, Declare.P_MINOR, Declare.P_REVISION) == -1)
                    {
                        var res = MessageBox.Show(
                            "V zadanom adresári " + Declare.DEST_PATH + " je nainštalovaná novšia verzia programu OLYMP " +
                            "ako je verzia UPGRADE." + Environment.NewLine + Environment.NewLine +
                            "Chcete aj napriek tomu inštalovať komponenty programu OLYMP do zadaného adresára?",
                            Declare.TTITLE,
                            MessageBoxButtons.YesNo,
                            MessageBoxIcon.Question,
                            MessageBoxDefaultButton.Button2);

                        if (res == DialogResult.No) splna = false;
                    }

                    // nie je nahodou spustena Alfa32.exe  -> Olymp.exe??? 
                    if (splna && HelpFunctions.JeSpustenyProgram(Declare.DEST_PATH, Declare.SUBOR_EXE))
                    {
                        MessageBox.Show(
                            "Nie je možné inštalovať UPGRADE, pretože program OLYMP v adresári " + Declare.DEST_PATH + " je práve spustený." +
                            Environment.NewLine + Environment.NewLine +
                            "Musíte najskôr ukončiť spustený program OLYMP.",
                            Declare.TTITLE,
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Exclamation);
                        splna = false;
                    }
                    else //if (splna)
                    {
                        // porovnám pôvodnú verziu a novú
                        string crv2Path = Path.Combine(Declare.DEST_PATH, "CRV2Kros.exe");
                        HelpFunctions.DajVerziuExe(crv2Path, out P_CRV2_MAJOR, out P_CRV2_MINOR, out P_CRV2_REVISION);

                        string oldV = $"{P_CRV2_MAJOR}.{P_CRV2_MINOR}.{P_CRV2_REVISION}";
                        string newV = $"{Declare.N_CRV2_MAJOR}.{Declare.N_CRV2_MINOR}.{Declare.N_CRV2_REVISION}";

                        if (oldV != newV)
                        {
                            if (HelpFunctions.JeSpustenyProgram(Declare.DEST_PATH, "CRV2Kros.exe"))
                            {
                                MessageBox.Show(
                                    "Nie je možné inštalovať UPGRADE, pretože tlačový modul programu Olymp - program CRV2Kros.exe je stále spustený." +
                                    Environment.NewLine + Environment.NewLine +
                                    "Ukončite program CRV2Kros.exe cez Správcu úloh, alebo reštartuje počítač.",
                                    Declare.TTITLE,
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Exclamation);
                                splna = false;
                            }
                        }
                    }
                }

                if (splna)
                {
                    tabControl1.SelectTab(1);//Pnl(5)
                    Instaluj();
                }
            }
            catch (Exception ex)
            {
                Declare.Errors.Add(ex.ToString());
                MessageBox.Show(ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnChangeDir_Click(object sender, EventArgs e)
        {
            try
            {
                this.Tag = Declare.DEST_PATH;

                // Path.Show vbModal  -> FolderBrowserDialog
                using (var dlg = new FolderBrowserDialog())
                {
                    dlg.Description = "Vyberte cieľový priečinok pre inštaláciu OLYMP";
                    dlg.ShowNewFolderButton = true;
                    if (!string.IsNullOrWhiteSpace(Declare.DEST_PATH) && Directory.Exists(Declare.DEST_PATH))
                        dlg.SelectedPath = Declare.DEST_PATH;

                    if (dlg.ShowDialog(this) == DialogResult.OK)
                    {
                        // NastavCestu lblDest, Me.Tag, True  -> nastav label podľa vybraného priečinka
                        NastavCestu(lblDest, dlg.SelectedPath, true);
                    }
                    else
                    {
                        // ak zrušené, vráť sa na pôvodnú cestu z this.Tag
                        NastavCestu(lblDest, (this.Tag as string) ?? string.Empty, true);
                    }
                }

                Declare.DEST_PATH = lblDest.Text;

                if (!string.Equals(Declare.DEST_PATH, Declare.DEST_PATH_NEZNAMY, StringComparison.OrdinalIgnoreCase) && Declare.MAJOR == 0)
                {
                    lblUpgrade.Text = DajVerziuUpgrade();
                    lblExe.Text = DajVerziuProgramu();
                    btnOk.Enabled = true;
                    this.Text = $"OLYMP {lblUpgrade.Text} - Sprievodca inštaláciou";
                }
                else
                {
                    lblExe.Text = DajVerziuProgramu();
                }
            }
            catch (Exception ex)
            {
                Declare.Errors.Add(ex.ToString());
                MessageBox.Show(ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnStorno_Click(object sender, EventArgs e)
        {
            try
            {
                if (_Stav == 1)
                {
                    Koniec();
                    return;
                }

                if (_Stav == 2)
                {
                    CheckAcrobatAndInstall();

                    if (Declare.RESTARTUJ)
                    {
                        var res = MessageBox.Show(
                            "Boli vykonané systémové zmeny, ktoré vyžadujú reštart počítača.\r\n" +
                            "Kým nebude počítač reštartovaný, program OLYMP a niektoré ďalšie programy nemusia pracovať správne.\r\n\r\n" +
                            "Chcete počítač reštartovať teraz?",
                            Declare.TTITLE,
                            MessageBoxButtons.YesNo,
                            MessageBoxIcon.Question
                        );

                        if (res == DialogResult.Yes)
                        {
                            if (!HelpFunctions.Restart())
                            {
                                MessageBox.Show(
                                    "Nepodarilo sa reštartovať počítač. Skúste počítač reštartovať manuálne.",
                                    "Informácia",
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Information
                                );
                            }
                            else
                            {
                                this.Close();// UkonciProgram();
                                return;
                            }
                        }

                        this.Close();
                        return;
                    }
                    else
                    {
                        this.Close();
                        return;
                    }
                }

                if (MessageBox.Show(
                        "Inštalácia programu OLYMP nebola dokončená.\r\n\r\n" +
                        "Prajete si ukončiť inštalačný program, aj keď neboli nainštalované žiadne komponenty?",
                        Declare.TTITLE,
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Question,
                        MessageBoxDefaultButton.Button2
                    ) == DialogResult.Yes)
                {
                    this.Close();
                }
            }
            catch (Exception ex)
            {
                Declare.Errors.Add(ex.ToString());
                MessageBox.Show(ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void linkLabelKros_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = Declare.URL_KROS,
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                Declare.Errors.Add(ex.ToString());
                MessageBox.Show(ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void linkLabelEmail_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try
            {
                if (!(sender is LinkLabel lbl)) return;

                Process.Start(new ProcessStartInfo
                {
                    FileName = $"mailto:{lbl.Text}",
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                Declare.Errors.Add(ex.ToString());
                MessageBox.Show(ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnSpat_Click(object sender, EventArgs e)
        {

        }

        private void OlympUpgrade_FormClosed(object sender, FormClosedEventArgs e)
        {
            try
            {
                //TODO comment after testing
                SaveResults(GetResultFile());
            }
            catch { }
        }

        #endregion Event handlers

        #region Methods

        /// <summary>
        /// ulozi vysledok instalacie - obsah listboxu
        /// </summary>
        private void UlozitVysledok()
        {
            try
            {
                string subor = GetResultFile();
                SaveResults(subor);

                MessageBox.Show($"Výsledok inštalácie bol uložený do súboru {subor}", Declare.TTITLE, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                Declare.Errors.Add(ex.ToString());
                MessageBox.Show("Nepodarilo sa uložiť výsledok inštalácie.", Declare.TTITLE, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private string GetResultFile() => Path.Combine(Declare.DEST_PATH, $"Instal_Olymp_datum[{DateTime.Now.Day}-{DateTime.Now.Month}-{DateTime.Now.Year}]_cas[{DateTime.Now.Hour}-{DateTime.Now.Minute}-{DateTime.Now.Second}].log");

        private void SaveResults(string file)
        {

            // Attempt to delete the file if it already exists
            if (File.Exists(file))
            {
                File.Delete(file);
            }

            // Open file for writing
            using (StreamWriter sw = new StreamWriter(file, false))
            {
                if (lst.Items != null && lst.Items.Count > 0)
                {
                    foreach (var item in lst.Items)
                    {
                        sw.WriteLine(item.ToString());
                    }
                }

                if (Declare.Errors != null && Declare.Errors.Count > 0)
                {
                    sw.WriteLine();
                    sw.WriteLine();
                    sw.WriteLine("====== DETAIL CHÝB ======");
                    sw.WriteLine();
                    foreach (var error in Declare.Errors)
                    {
                        sw.WriteLine(error);
                        sw.WriteLine("------");
                    }
                }
            }
        }

        /// <summary>
        /// samotna instalacia
        /// </summary>
        private void Instaluj()
        {
            btnOk.Visible = false;
            btnStorno.Enabled = false;
            Application.DoEvents();

            OverLicenciu();

            OverVolneMiesto();
            Application.DoEvents();

            // OverArchiv         'toto s novym komponentom nevieme
            // VymazSuboryStarejAlfy

            KopirujSubory();
            Application.DoEvents();

            InstalujHotFixPreMapi();//TODO Win61xXX.msu sa nenachdza v Zdroje
            Application.DoEvents();

            RegistryFunctions.UlozDoRegistrovVerziu();
            Application.DoEvents();

            Thread.Sleep(500); //iba kvoli tomu, aby to neprefrcalo tak rychlo
            Application.DoEvents();

            ZobrazVysledok();
            Application.DoEvents();
        }

        private void ZobrazVysledok()
        {
            if (_BoliChyby)
            {
                //Pnl(5).Top = 6500
                //Pnl(3).Top = 0
                tabControl1.SelectTab(3);

                _Stav = 1;
                btnDalej.Visible = false;

                btnOk.Text = "&Uložiť";
                btnOk.Visible = true;
                btnOk.Focus();

                btnStorno.Text = "Ď&alej >";
                btnStorno.Enabled = true;
            }
            else
            {
                //Pnl(5).Top = 6500
                //Pnl(4).Top = 0
                tabControl1.SelectTab(2);

                _Stav = 2;
                btnDalej.Text = "< S&päť";
                btnDalej.Visible = false;

                btnOk.Text = "&Spusti";
                btnOk.Visible = true;
                btnOk.Focus();

                btnStorno.Enabled = true;
                if (Declare.RESTARTUJ)
                {
                    btnOk.Visible = false;
                    lblInfo.Visible = false;
                }

                //!!!NELOGICKE!!!
                //If BoliChyby Then
                //  CmdOk.Visible = False
                //  Lbl(7) = "Inštalácia prebehla s chybami. Pokúste sa tieto chyby odstráni a potom spustite program " & _
                //            PridajLomitko(AKT_ADRESAR) & MENO_EXE & " znova. " & _
                //           "Ak sa vyskytla chyba typu ""nedá sa prepísa"", príèinou môže by " & _
                //           "to, že je súbor otvorený alebo spustený. V tomto prípade musíte súbor zavrie alebo ukonèi."
                //End If

                btnStorno.Text = "&Koniec";
                if (btnOk.Visible) btnOk.Focus();
            }
        }

        private void InstalujHotFixPreMapi()
        {
            try
            {
                // Windows 7 / Server 2008 R2 == 6.1
                Version v = Environment.OSVersion.Version;
                if (v.Major == 6 && v.Minor == 1)
                {
                    // pick 32/64-bit package
                    string fileName = Environment.Is64BitOperatingSystem ? "Win61x64.msu" : "Win61x86.msu";
                    string msuPath = Path.Combine(Declare.DEST_PATH ?? string.Empty, "Zdroje", fileName);

                    if (File.Exists(msuPath))
                    {
                        var psi = new ProcessStartInfo
                        {
                            FileName = "wusa.exe",
                            Arguments = $"\"{msuPath}\" /quiet /norestart",
                            UseShellExecute = false,
                            CreateNoWindow = true,
                            WindowStyle = ProcessWindowStyle.Hidden
                        };

                        Process.Start(psi);
                    }
                }
            }
            catch (Exception ex) { Declare.Errors.Add(ex.ToString()); }
        }

        /// <summary>
        /// Prida subory, ktore nebolo mozne nakopirovat kvoli chybe
        /// </summary>
        public void KopirujSubory()
        {
            int pocet;
            string povodnaVerziaX;
            long povodnaVerzia, chyba;
            bool prepis;
            var neprepisovatDBS = new List<string>(); ;
            //Dim chilkatEntry As New ChilkatZipEntry2

            try
            {
                ProgresiaPreset(0);

                lblDestFile.Text = "Kopírovanie súborov ...";
                lblDestFile.Refresh();

                using (ZipArchive zip = ZipFile.Open(Path.Combine(Declare.AKT_ADRESAR, Declare.SUBOR_ZIP), ZipArchiveMode.Read, _zipEncoding))
                {
                    ProgresiaPreset(zip.Entries.Count);
                    var tempAdr = HelpFunctions.DajTemp(Declare.AKT_ADRESAR);
                    tempAdr = Path.Combine(tempAdr, "OlympUpgradeTempZip");
                    //zip.ExtractToDirectory(tempAdr);

                    _Chybne.Clear();
                    _Spravne.Clear();
                    _Preskocene.Clear();

                    _PomCol.Clear();

                    chyba = 1;

                    if (zip.Entries.Count > 0)
                    {
                        // toto plati pre adresarik data
                        neprepisovatDBS.Clear();

                        var dataEntrie = zip.Entries.Where(e => e.FullName.ToUpper().Contains(@"DATA/")
                                                                        && e.Length > 0);

                        foreach (var entry in dataEntrie)
                        {
                            var zipPath = entry.FullName.Replace('\\', '/');

                            // iba položky pod "DATA/"
                            if (zipPath.StartsWith("DATA/", StringComparison.OrdinalIgnoreCase))
                            {
                                // plná cieľová cesta na disku
                                var destPath = Path.Combine(Declare.DEST_PATH, zipPath.Replace('/', Path.DirectorySeparatorChar));

                                if (!File.Exists(destPath))
                                {
                                    prepis = true;
                                    foreach (var dbName in neprepisovatDBS)
                                    {
                                        var guardedPrefix = "data/pripojenedata/" + dbName;
                                        if (zipPath.StartsWith(guardedPrefix, StringComparison.OrdinalIgnoreCase))
                                        {
                                            prepis = false;
                                            break;
                                        }
                                    }

                                    if (prepis)
                                    {
                                        try
                                        {
                                            Directory.CreateDirectory(Path.GetDirectoryName(destPath));
                                            entry.ExtractToFile(destPath, overwrite: true);
                                        }
                                        catch (Exception ex)
                                        {
                                            Declare.Errors.Add(ex.ToString());
                                            _PomCol.Add(destPath);
                                            chyba++;
                                        }
                                    }
                                }
                                else
                                {
                                    // ak už existuje a je to MDB, pridaj jeho názov (bez prípony) do "neprepisovať"
                                    if (string.Equals(Path.GetExtension(zipPath), ".mdb", StringComparison.OrdinalIgnoreCase))
                                    {
                                        var name = Path.GetFileNameWithoutExtension(zipPath).ToLowerInvariant();
                                        neprepisovatDBS.Add(name);
                                    }
                                }
                            }
                            ProgresiaTik();
                        }
                    }

                    if (chyba != 1 || _PomCol.Count > 0)
                    {
                        PridajChybu(); // call your method

                        var target = Path.Combine(Declare.DEST_PATH, "DATA");

                        MessageBox.Show(
                            "Nastala chyba pri kopírovaní súborov do adresára " + target + Environment.NewLine + Environment.NewLine +
                            "Súbory, ktoré sa nepodarilo nainštalovať, budú uvedené na konci inštalácie aj s typom chyby.",
                            Declare.TTITLE,
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error
                        );
                    }

                    // zmazem vsetky REPORTY\*.RPT okrem REPORTY\OB_FA*.RPT
                    try
                    {
                        long parseVal = 0;
                        povodnaVerziaX = DajVerziuProgramu();
                        if (povodnaVerziaX.IndexOf(".") > -1)
                            long.TryParse(povodnaVerziaX.Substring(0, povodnaVerziaX.IndexOf(".")), out parseVal);
                        povodnaVerzia = parseVal * 100;

                        parseVal = 0;
                        if (povodnaVerziaX.IndexOf(".") > -1)
                        {
                            povodnaVerziaX = povodnaVerziaX.Substring(povodnaVerziaX.IndexOf(".") + 1);
                            long.TryParse(povodnaVerziaX.Substring(0, povodnaVerziaX.IndexOf(".")), out parseVal);
                        }
                        povodnaVerzia = povodnaVerzia + parseVal;

                        if (povodnaVerzia < 946)
                        {
                            var rep = Path.Combine(Declare.DEST_PATH, "REPORTY");
                            if (Directory.Exists(rep))
                            {
                                foreach (var file in Directory.EnumerateFiles(rep))
                                    File.Delete(file);
                            }
                        }

                        //vzdy mazem len systemove reporty a uzivatelske nechavam
                        //---RPT REPORTY---
                        VymazZostavy("Reporty\\", "rpt");
                        VymazZostavy("Reporty\\", "repx");
                        //---XLS REPORTY---
                        VymazZostavy("Reporty\\EXCEL\\", "xls");
                        VymazZostavy("Reporty\\EXCEL\\", "pdf"); //v zostavach pre excel bude aj navod v pdf
                                                                 //---PDF REPORTY-- -
                        VymazZostavy("Reporty\\PDF\\", "pdf");

                    }
                    catch (Exception ex) { Declare.Errors.Add(ex.ToString()); }

                    // --- rozbal systémové reporty podľa listov ---
                    RozbalZostavy(zip, "Reporty\\", "\\.rpt", Declare.FILE_REPORTY_TXT);
                    RozbalZostavy(zip, "Reporty\\", "\\.repx", Declare.FILE_REPORTY_TXT);
                    RozbalZostavy(zip, "Reporty\\Excel\\", "\\.xls", Declare.FILE_REPORTY_EXCEL_TXT);
                    RozbalZostavy(zip, "Reporty\\Excel\\", "\\.pdf", Declare.FILE_REPORTY_EXCEL_P_TXT);
                    RozbalZostavy(zip, "Reporty\\Pdf\\", "\\.pdf", Declare.FILE_REPORTY_PDF_TXT);


                    // --- GRAFIKA\*.*  ---
                    _PomCol.Clear();
                    ExtractFilesExtensionOnPath(zip, Declare.DEST_PATH, "Grafika", "\\.*");
                    CheckCopyException(Path.Combine(Declare.DEST_PATH, "Grafika"));

                    // --- SKRIPTY\CREATE\*.sql ---
                    _PomCol.Clear();
                    ExtractFilesExtensionOnPath(zip, Declare.DEST_PATH, "Skripty\\create\\", "\\.sql");
                    CheckCopyException(Path.Combine(Declare.DEST_PATH, "Skripty", "create"));

                    // --- SKRIPTY\DROP\*.sql ---
                    _PomCol.Clear();
                    ExtractFilesExtensionOnPath(zip, Declare.DEST_PATH, "Skripty\\drop\\", "\\.sql");
                    CheckCopyException(Path.Combine(Declare.DEST_PATH, "Skripty", "drop"));

                    // --- ZDROJE\*.* ---
                    _PomCol.Clear();
                    ExtractFilesExtensionOnPath(zip, Declare.DEST_PATH, "Zdroje", "\\.*");
                    CheckCopyException(Path.Combine(Declare.DEST_PATH, "Zdroje"));

                    // --- Vzory ---
                    var vzoryPath = Path.Combine(Declare.DEST_PATH, "Vzory");
                    if (Directory.Exists(vzoryPath))
                        HelpFunctions.NastavPrava(vzoryPath, "*.*", FileAttributes.Normal);

                    _PomCol.Clear();
                    ExtractFilesExtensionOnPath(zip, Declare.DEST_PATH, "Vzory", "\\.*");
                    CheckCopyException(vzoryPath);
                    if (_PomCol.Count > 0)
                        HelpFunctions.NastavPrava(Path.Combine(vzoryPath, "Vzory\\"), "*.*", FileAttributes.ReadOnly);

                    CopyFileIfExists(Declare.AKT_ADRESAR, Declare.DEST_PATH, "OLYMP.CHM");
                    CopyFileIfExists(Declare.AKT_ADRESAR, Declare.DEST_PATH, "CRV2Kros.exe");
                    CopyFileIfExists(Declare.AKT_ADRESAR, Declare.DEST_PATH, "ADOWrapper.dll");
                    CopyFileIfExists(Declare.AKT_ADRESAR, Declare.DEST_PATH, "itextsharp.dll");
                    CopyFileIfExists(Declare.AKT_ADRESAR, Declare.DEST_PATH, "Kros.BankTransfers.dll");

                    VymazStareSubory();

                    // --- SK\*.* ---
                    _PomCol.Clear();
                    ExtractFilesExtensionOnPath(zip, Declare.DEST_PATH, "sk", "\\.*");
                    CheckCopyException(Path.Combine(Declare.DEST_PATH, "sk"));

                    // --- Citajma*.txt, CoJeNove.txt, Dealers.txt a EXE - prepisuju sa ---
                    _PomCol.Clear();
                    ExtractFileOnPath(zip, Declare.DEST_PATH, "", "LEGISL.MDB");
                    ExtractFileOnPath(zip, Declare.DEST_PATH, "", "OLYMP.chm");
                    ExtractFileOnPath(zip, Declare.DEST_PATH, "", "CRV2Kros.exe");
                    ExtractFileOnPath(zip, Declare.DEST_PATH, "", Declare.SUBOR_EXE);
                    ExtractFileOnPath(zip, Declare.DEST_PATH, "", Declare.SUBOR_TeamViewer);
                    ExtractFileOnPath(zip, Declare.DEST_PATH, "", Declare.SUBOR_ClickYes);
                    ExtractFileOnPath(zip, Declare.DEST_PATH, "", "Kros FTP Uploader.exe");
                    ExtractFileOnPath(zip, Declare.DEST_PATH, "", "Downloader.exe");
                    ExtractFileOnPath(zip, Declare.DEST_PATH, "", "Aktivácia.exe");


                    //automaticke rozbalenie suborov podla zoznamSuborov.txt
                    ExtractFileOnPath(zip, Declare.DEST_PATH, "", "zoznamSuborov.txt");
                    var zoznamPath = Path.Combine(Declare.DEST_PATH, "zoznamSuborov.txt");
                    if (HelpFunctions.ExistujeSubor(zoznamPath))
                    {
                        var lines = File.ReadAllLines(zoznamPath);

                        foreach (var line in lines)
                        {
                            var file = line.Trim();
                            if (file.Length == 0) continue;

                            var path = Path.GetDirectoryName(file);
                            file = Path.GetFileName(file);

                            ExtractFileOnPath(zip, Path.Combine(Declare.DEST_PATH, path), path, file);
                        }
                    }

                    CheckCopyException(Declare.DEST_PATH);

                    // DevExpress.*.dll
                    _PomCol.Clear();
                    ExtractFilesExtensionOnPath(zip, Declare.DEST_PATH, "", "DevExpress\\..*\\.dll");
                    CheckCopyException(Declare.DEST_PATH);

                    // --- Close ZIP ---
                }

                // --- Presun ZIP/installer do UPGRADE pripadne zmaz ak existuje ---
                var upgradeZipFile = Path.Combine(Declare.DEST_PATH, "UPGRADE", Declare.SUBOR_ZIP);
                if (HelpFunctions.ExistujeSubor(upgradeZipFile))
                {
                    HelpFunctions.TryToDeleteFile(upgradeZipFile);
                }

                //Copy SUBOR_ZIP
                try
                {
                    var upgradeDir = Path.Combine(Declare.DEST_PATH, "UPGRADE");
                    Directory.CreateDirectory(upgradeDir);

                    var upgradeZipPath = Path.Combine(Declare.DEST_PATH, "UPGRADE", Declare.SUBOR_ZIP);

                    lblDestFile.Text = upgradeZipPath;
                    lblDestFile.Refresh();

                    bool copyOk = true;
                    try
                    {
                        var src = Path.Combine(Declare.AKT_ADRESAR, Declare.SUBOR_ZIP);
                        File.Copy(src, upgradeZipPath, overwrite: true);
                    }
                    catch (Exception ex) { Declare.Errors.Add(ex.ToString()); copyOk = false; }

                    if (copyOk)
                    {
                        _Spravne.Add(Path.Combine("UPGRADE", Declare.SUBOR_ZIP));
                        //ak so spustany z akt adresara, tak vymazem OlympUpgrade.zip
                        if (HelpFunctions.ExistujeSubor(Path.Combine(Declare.DEST_PATH, Declare.SUBOR_ZIP)))
                            HelpFunctions.TryToDeleteFile(Path.Combine(Declare.DEST_PATH, Declare.SUBOR_ZIP));
                    }
                }
                catch (Exception ex) { Declare.Errors.Add(ex.ToString()); }

                //Copy SUBOR_INST
                try
                {
                    //od 2009 je mozne stahovat aj kompletnu instalaciu, tak ju vymazem z cieloveho adresara, ak tam je
                    var upgradeINST_Path = Path.Combine(Declare.DEST_PATH, "UPGRADE", Declare.SUBOR_INST);
                    if (HelpFunctions.ExistujeSubor(upgradeINST_Path))
                        HelpFunctions.TryToDeleteFile(upgradeINST_Path);

                    Directory.CreateDirectory(Path.Combine(Declare.DEST_PATH, "UPGRADE"));

                    lblDestFile.Text = upgradeINST_Path;
                    lblDestFile.Refresh();

                    var destINST_Path = Path.Combine(Declare.DEST_PATH, Declare.SUBOR_INST);
                    if (HelpFunctions.ExistujeSubor(destINST_Path))
                    {
                        File.Copy(destINST_Path, upgradeINST_Path, overwrite: true);
                        HelpFunctions.TryToDeleteFile(destINST_Path);
                    }
                }
                catch (Exception ex) { Declare.Errors.Add(ex.ToString()); }

                //Copy SUBOR_INST_NEW
                try
                {
                    //od 2010 sa plna instalacka vola OlympInstall.exe kvoli trojanovi
                    var upgradeINST_NEW_Path = Path.Combine(Declare.DEST_PATH, "UPGRADE", Declare.SUBOR_INST_NEW);
                    var destINST_New_Path = Path.Combine(Declare.DEST_PATH, Declare.SUBOR_INST_NEW);

                    if (HelpFunctions.ExistujeSubor(Path.Combine(destINST_New_Path))
                        && HelpFunctions.ExistujeSubor(Path.Combine(upgradeINST_NEW_Path)))
                        HelpFunctions.TryToDeleteFile(Path.Combine(upgradeINST_NEW_Path));

                    Directory.CreateDirectory(Path.Combine(Declare.DEST_PATH, "UPGRADE"));

                    lblDestFile.Text = upgradeINST_NEW_Path;
                    lblDestFile.Refresh();

                    if (HelpFunctions.ExistujeSubor(destINST_New_Path))
                    {
                        //File.Copy(destINST_New_Path, upgradeINST_NEW_Path, overwrite: true);
                        HelpFunctions.TryToDeleteFile(upgradeINST_NEW_Path);
                        File.Move(destINST_New_Path, upgradeINST_NEW_Path);
                    }

                    pocet = 0;
                    //iba kvoli tomu, aby to neprefrcalo tak rychlo a naozaj ho stihol zmazat
                    //ak nestihne zmazat ani po 500 tak idem dalej, asi je naozaj problem
                    while (HelpFunctions.ExistujeSubor(destINST_New_Path) && pocet < 10)
                    {
                        Thread.Sleep(50);
                        pocet = pocet + 1;
                    }
                }
                catch (Exception ex) { Declare.Errors.Add(ex.ToString()); }

                ProgresiaDone();
                ZapisVysledokDoListu();
            }
            catch (Exception ex)
            {
                Declare.Errors.Add(ex.ToString());
                MessageBox.Show($"Nastala chyba pri rozzipovani Erl : {ex.ToString()}\r\n\r\n {VratZoznam()}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Zapise vysledok instalacie do listboxu
        /// </summary>
        private void ZapisVysledokDoListu()
        {
            lst.Items.Clear();
            lst.Items.Add(string.Empty);

            if (_Chybne == null) _Chybne = new List<string>();
            if (_Preskocene == null) _Preskocene = new List<string>();
            if (_Spravne == null) _Spravne = new List<string>();

            if (_Chybne.Count == 0)
            {
                lst.Items.Add("   Inštalácia prebehla bez chýb.");
                _BoliChyby = false;
            }
            else
            {
                lst.Items.Add("   Inštalácia prebehla s chybami!");
                _BoliChyby = true;
            }

            if (Declare.KTO_VOLAL == Declare.VOLAL_UZIVATEL || Declare.KTO_VOLAL == Declare.VOLAL_OLYMP)
            {
                lst.Items.Add(string.Empty);
                lst.Items.Add("   Verzia pred inštaláciou: \t" + lblExe.Text);
                lst.Items.Add("   Verzia aktuálna: \t\t" + lblUpgrade.Text);
            }

            lst.Items.Add(string.Empty);
            lst.Items.Add("   Počet neinštalovaných kvôli chybe: \t" + _Chybne.Count);
            lst.Items.Add("   Počet preskočených: \t\t" + _Preskocene.Count);
            lst.Items.Add("   Počet správne inštalovaných: \t" + _Spravne.Count);

            if (_Chybne.Count > 0)
            {
                lst.Items.Add(string.Empty);
                lst.Items.Add("== Neinštalované komponenty kvôli chybe ========================");
                lst.Items.Add(string.Empty);
                foreach (var s in _Chybne)
                    lst.Items.Add("   " + s);
            }

            if (_Preskocene.Count > 0)
            {
                lst.Items.Add(string.Empty);
                lst.Items.Add("== Preskočené komponenty ====================================");
                lst.Items.Add(string.Empty);
                foreach (var s in _Preskocene)
                    lst.Items.Add("   " + s);
            }

            if (_Spravne.Count > 0)
            {
                lst.Items.Add(string.Empty);
                lst.Items.Add("== Inštalované komponenty ====================================");
                lst.Items.Add(string.Empty);
                foreach (var s in _Spravne)
                    lst.Items.Add("   " + s);
            }

            _Chybne.Clear();
            _Preskocene.Clear();
            _Spravne.Clear();
        }

        private string VratZoznam()
        {
            var sb = new StringBuilder();
            sb.Append("\r\nXYZ\r\n");

            if (_PomCol != null && _PomCol.Count > 0)
            {
                foreach (var nieco in _PomCol)
                {
                    sb.Append(Environment.NewLine);
                    sb.Append("X").Append(nieco).Append("X");
                }
            }

            return sb.ToString();
        }

        private void VymazStareSubory()
        {
            DeleteMatching(Declare.DEST_PATH, "Mzdy.exe");

            DeleteMatching(Declare.DEST_PATH, "DevExpress.*");

            string skDir = Path.Combine(Declare.DEST_PATH, "sk");
            DeleteMatching(skDir, "DevExpress.*");

            DeleteMatching(Declare.DEST_PATH, "Kros.Licenses*");
        }

        private static void DeleteMatching(string directory, string searchPattern)
        {
            if (!Directory.Exists(directory)) return;

            try
            {
                foreach (var file in Directory.EnumerateFiles(directory, searchPattern, SearchOption.TopDirectoryOnly))
                {
                    try { File.SetAttributes(file, FileAttributes.Normal); } catch (Exception ex) { Declare.Errors.Add(ex.ToString()); }
                    try { File.Delete(file); } catch (Exception ex) { Declare.Errors.Add(ex.ToString()); }
                }
            }
            catch (Exception ex) { Declare.Errors.Add(ex.ToString()); }
        }

        private void CopyFileIfExists(string sourcePath, string destPath, string file)
        {
            var sourceFilePath = Path.Combine(sourcePath, file);
            var destFilePath = Path.Combine(destPath, file);

            //Ak sorceFile existuje a nejedna o ten isty subor
            if (File.Exists(sourceFilePath)
                && !string.Equals(sourcePath, destPath, StringComparison.OrdinalIgnoreCase))
            {
                lblDestFile.Text = destFilePath.ToUpperInvariant();
                lblDestFile.Refresh();

                try
                {
                    Directory.CreateDirectory(destPath);

                    File.Copy(sourceFilePath, destFilePath, overwrite: true);

                    _Spravne.Add(file);
                }
                catch (Exception ex)
                {
                    Declare.Errors.Add(ex.ToString());
                    MessageBox.Show(
                        $"Nastala chyba pri kopírovaní súborov do adresára {destFilePath}\r\n\r\n" +
                        "Súbory, ktoré sa nepodarilo nainštalovať, budú uvedené na konci inštalácie aj s typom chyby.",
                        Declare.TTITLE, MessageBoxButtons.OK, MessageBoxIcon.Error);

                    _Chybne.Add("súbor sa nedá prepísať" + "\t" + file);
                }
            }
        }

        private void CheckCopyException(string path)
        {
            if (_PomCol.Count > 0)
            {
                PridajChybu();
                MessageBox.Show(
                    $"Nastala chyba pri kopírovaní súborov do adresára {path}\r\n\r\n" +
                    "Súbory, ktoré sa nepodarilo nainštalovať, budú uvedené na konci inštalácie aj s typom chyby.",
                    Declare.TTITLE, MessageBoxButtons.OK, MessageBoxIcon.Error);

                _PomCol.Clear();
            }
        }

        public void RozbalZostavy(ZipArchive zip, string Cesta, string pripona, string subor)
        {
            //cesta v tvare: "REPORTY\"
            //pripona v tvare: "rpt"
            _PomCol.Clear();
            string destPath = ExtractFilesExtensionOnPath(zip, Declare.DEST_PATH, Cesta, pripona);


            // 2) Rozbaliť konkrétny súbor/subset (Cesta + subor)
            //_PomCol.Clear();
            ExtractFileOnPath(zip, destPath, Cesta, subor);

            if (_PomCol.Count > 0)
            {
                PridajChybu();
                MessageBox.Show(
                    "Nastala chyba pri kopírovaní súborov do adresára " +
                    Path.Combine(Declare.DEST_PATH, Cesta ?? "") + Environment.NewLine + Environment.NewLine +
                    "Súbory, ktoré sa nepodarilo nainštalovať, budú uvedené na konci inštalácie aj s typom chyby.",
                    Declare.TTITLE, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            // 3) Volanie tvojej existujúcej post-procedúry
            pripona = pripona.Replace("\\", string.Empty);//.Replace("/", string.Empty);
            NastavDlheNazvy(Cesta, pripona, 50, subor);
        }

        private void ExtractFileOnPath(ZipArchive zip, string destPath, string filePath, string subor)
        {
            var fileName = Path.Combine(filePath, subor).Replace('\\', '/');
            var entryFile = zip.Entries
                           .FirstOrDefault(e => string.Equals(e.FullName, fileName, StringComparison.OrdinalIgnoreCase)
                                            && e.Length > 0);
            if (entryFile != null)
                ExtractToFile(destPath, entryFile);
            else
                _PomCol.Add($"Nenasiel sa: '{fileName}'");
        }

        /// <summary>
        /// Rozbali subory na daje searchPath s lubovolnou priponou ("\\.dll", "\\.*")
        /// Ak je searchPath prazdny, pouzije sa extension ako regex vyraz ($@"(?i)^{extension}$")
        /// </summary>
        /// <param name="zip"></param>
        /// <param name="destFolder"></param>
        /// <param name="searchPath"></param>
        /// <param name="extension"></param>
        /// <returns></returns>
        private string ExtractFilesExtensionOnPath(ZipArchive zip, string destFolder, string searchPath, string extension)
        {
            //extension = "\\.rpt";
            extension = (extension ?? "").ToLowerInvariant();
            var regexMask = string.Empty;

            if (!string.IsNullOrWhiteSpace(searchPath))
            {
                searchPath = searchPath.Last() == '\\' ? searchPath.Substring(0, searchPath.Length - 1) : searchPath;
                searchPath = searchPath.Last() == '/' ? searchPath.Substring(0, searchPath.Length - 1) : searchPath;
                searchPath = searchPath.Replace('\\', '/');

                regexMask = $@"(?i)^{Regex.Escape(searchPath)}[\\/][^\\/]+{extension}$";
            }
            else
                regexMask = $@"(?i)^{extension}$";

            var destPath = Path.Combine(destFolder, searchPath);
            Directory.CreateDirectory(destPath);

            // 1) Rozbaliť všetky *.pripona v podsložke Cesta
            //int zipRes = UnzipMatching(Declare.DEST_PATH, mask1, PomCol);
            //string regexMask = $@"(?i)^{Regex.Escape(searchPath)}[\\/][^\\/]+\.{extension}$";

            var dataEntrie = zip.Entries.Where(e => Regex.IsMatch(e.FullName, regexMask) && e.Length > 0);
            if (dataEntrie.Count() > 0)
            {
                foreach (var entry in dataEntrie)
                {
                    ExtractToFile(destPath, entry);
                }
            }
            else
                _PomCol.Add($"Ziadna zhoda regex: '{regexMask}'");
            return destPath;
        }

        private void ExtractToFile(string destPath, ZipArchiveEntry entry)
        {
            try
            {
                Directory.CreateDirectory(destPath);
                entry.ExtractToFile(Path.Combine(destPath, entry.Name), overwrite: true);
                lblDestFile.Text = $"{entry.FullName}";
                lblDestFile.Refresh();

                //Thread.Sleep(10);//iba kvoli tomu, aby to neprefrcalo tak rychlo

                _Spravne.Add(entry.FullName);
            }
            catch (Exception ex)
            {
                Declare.Errors.Add(ex.ToString());
                _PomCol.Add($"{entry.FullName} ex: {ex.Message}");
            }
            ProgresiaTik();
        }

        private void NastavDlheNazvy(string cesta, string pripona, int hranica, string txtSubor)
        {
            // Build absolute target directory
            var destDir = Path.GetFullPath(Path.Combine(Declare.DEST_PATH, cesta ?? string.Empty));
            Directory.CreateDirectory(destDir);

            // Mapping file (e.g., REPORTY.TXT) lives in destDir
            var mappingPath = Path.Combine(destDir, txtSubor);
            var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            if (!File.Exists(mappingPath)) return;

            using (var sr = new StreamReader(mappingPath, Encoding.Default, detectEncodingFromByteOrderMarks: true))
            {
                string line;
                while ((line = sr.ReadLine()) != null)
                {
                    if (string.Equals(line, "EOF", StringComparison.OrdinalIgnoreCase)) break;
                    if (line.Length < 6) continue;

                    // VB6 Mid$(s,1,6) and Mid$(s,8)
                    var key = line.Substring(0, Math.Min(6, line.Length));
                    var value = line.Length >= 8 ? line.Substring(7) : string.Empty;
                    map[key] = (value ?? string.Empty).Trim();
                }
            }

            // Enumerate files matching *{pripona} (VB6 Dir$ with "*"+pripona)
            foreach (var filePath in Directory.EnumerateFiles(destDir, "*" + pripona, SearchOption.TopDirectoryOnly))
            {
                var fileName = Path.GetFileName(filePath);

                try { File.SetAttributes(filePath, FileAttributes.Normal); } catch (Exception ex) { Declare.Errors.Add(ex.ToString()); }

                if (fileName.Length >= 6 && fileName[3] == '-')
                {
                    if (int.TryParse(fileName.Substring(4, 2), out var num) && num < hranica)
                    {
                        var key = fileName.Substring(0, 6);
                        if (map.TryGetValue(key, out var longName))
                        {
                            var newNameCore = key + "-" + longName;
                            var newPath = Path.Combine(destDir, newNameCore + pripona);

                            try
                            {
                                //File.SetAttributes(newPath, FileAttributes.Normal);
                                if (File.Exists(newPath))
                                    File.Replace(filePath, newPath, null);
                                else
                                    File.Move(filePath, newPath);
                            }
                            catch (Exception ex) { Declare.Errors.Add(ex.ToString()); }
                        }
                    }
                }
            }
        }

        private void VymazZostavy(string CestaKRep, string pripona)
        {
            //cesta v tvare: "REPORTY\"
            //pripona v tvare: "rpt"
            var dir = Path.Combine(Declare.DEST_PATH, CestaKRep ?? string.Empty);

            // zmazem prvych 49 reportov
            for (int i = 0; i <= 4; i++)
            {
                // maska: "???-0?*.rpt", "???-1?*.rpt", ...
                string maska = $"???-{i}?*.{pripona}";

                HelpFunctions.NastavPrava(dir, maska, FileAttributes.Normal);
                try
                {
                    if (Directory.Exists(dir))
                        foreach (var path in Directory.GetFiles(dir, maska))
                        {
                            try
                            {
                                if (File.Exists(path))
                                    File.Delete(path);
                            }
                            catch (Exception ex) { Declare.Errors.Add(ex.ToString()); }
                        }
                }
                catch (Exception ex) { Declare.Errors.Add(ex.ToString()); }
            }
        }

        /// <summary>
        /// prida subory, ktore nebolo mozne nakopirovat kvoli chybe
        /// </summary>
        void PridajChybu()
        {
            if (_PomCol != null && _PomCol.Count > 0)
            {
                foreach (var v in _PomCol)//.Values)
                {
                    _Chybne.Add("nedá sa prepísať\t" + v);
                }
            }
        }
        public void OverVolneMiesto()
        {
            double VolneMiesto, PotrebnaVelkost;
            string prompt = string.Empty;

            ProgresiaPreset(0);

            lblDestFile.Text = "Kontrola voľného miesta ...";
            lblDestFile.Refresh();



            //TODO prec, kontroluje "Mzdy.lic"
            if (Declare.DEST_PATH != Declare.AKT_ADRESAR
                && HelpFunctions.ExistujeSubor(HelpFunctions.PridajLomitko(Declare.AKT_ADRESAR) + Declare.LICENCIA))
            {
                // PotrebnaVelkost = PotrebnaVelkost + FileLen(Declare.PridajLomitko(Declare.AKT_ADRESAR) & Declare.LICENCIA) / 1024;
            }

            VolneMiesto = HelpFunctions.DiskSpaceKB(Declare.DEST_PATH);
            PotrebnaVelkost = DajVelkostSuborovVZip(HelpFunctions.PridajLomitko(Declare.AKT_ADRESAR) + Declare.SUBOR_ZIP) / 1024d;

            if (VolneMiesto < 0 || PotrebnaVelkost > VolneMiesto)
            {
                var destDiskLetter = Path.GetPathRoot(Declare.DEST_PATH);
                var potrebnaVelkostMB = Math.Ceiling(PotrebnaVelkost / 1024d + 2);
                prompt = " Chcete pokračovať v inštalácii?\r\n\r\n" +
                               $"Zvoľte Áno, ak ste si istý, že máte dostatok voľného miesta ({potrebnaVelkostMB} MB) na disku {destDiskLetter}. \r\n";

                if (VolneMiesto < 0)
                {
                    prompt = $"Nastala chyba pri zisťovaní voľného miesta na cieľovom disku {destDiskLetter}." +
                               prompt;
                }
                else if (PotrebnaVelkost > VolneMiesto)// * 1024 TODO ???) //Je malo miesta
                {
                    prompt = $"Na cieľovom disku {destDiskLetter} je nedostatok voľného miesta." +
                                   prompt;
                }

                prompt += "V opačnom prípade zvoľte Nie a inštalácia bude ukončená. Uvoľnite požadované miesto na cieľovom disku " +
                                "a potom opäť spustite program ";

                if (Declare.KTO_VOLAL == Declare.VOLAL_INSTALL)
                {
                    prompt += HelpFunctions.PridajLomitko(Declare.AKT_ADRESAR) + Declare.MENO_EXE;
                }
                else if (Declare.KTO_VOLAL == Declare.VOLAL_UZIVATEL || Declare.KTO_VOLAL == Declare.VOLAL_OLYMP)
                {
                    prompt += Declare.MENO_EXE;
                }

                var res = MessageBox.Show(prompt, Declare.TTITLE, MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1);
                if (res == DialogResult.No)
                {
                    ProgresiaDone();
                    Declare.ExitProg(0);
                    return;
                }
            }

            ProgresiaDone();
        }

        /// <summary>
        /// vrati nekomprimovanu velkost suborov v zip v bytoch
        /// </summary>
        /// <param name="pathZip"></param>
        /// <returns></returns>
        public long DajVelkostSuborovVZip(string pathZip)
        {
            if (string.IsNullOrWhiteSpace(pathZip) || !File.Exists(pathZip))
                return 0;

            long sum = 0;

            using (ZipArchive zip = ZipFile.Open(pathZip, ZipArchiveMode.Read, _zipEncoding))
            {
                ProgresiaPreset(zip.Entries.Count);

                foreach (var e in zip.Entries)
                {
                    try
                    {
                        // Rátame iba súbory (nie adresáre)
                        bool isDirectory = e.Name.Length == 0 || e.FullName.EndsWith("/", StringComparison.Ordinal);
                        if (!isDirectory)
                        {
                            sum += e.Length; // uncompressed size/
                        }
                    }
                    catch (Exception ex) { Declare.Errors.Add(ex.ToString()); }

                    ProgresiaTik();
                }
            }
            return sum;
        }

        /// <summary>
        /// kontrola licencie
        /// </summary>
        private void OverLicenciu()
        {
            bool MaOstruVerziu = false;
            bool KopirujLICENCIU = false;
            string Lic = string.Empty; // koreň, odkiaľ sa kopíruje licencia (adresár inštalačiek)

            ProgresiaPreset(0);
            lblDestFile.Text = "Kontrola licencie ...";
            lblDestFile.Refresh();
            ProgresiaPreset(3);

            try
            {
                // 1) Zisti „ostrú verziu“ v PC (a pri tom načítaj dátumy z licenčného súboru)
                CitajRegistracnySubor(Declare.PCD_POCITAC, Declare.DEST_PATH);
                MaOstruVerziu = (Declare.VerziaPC == Declare.LIC_OSTRA);

                ProgresiaTik();

                // 5) Nemá licenciu na inštalačkách a chce „ostrú“ – skontroluj nárok
                if (MaOstruVerziu && !KopirujLICENCIU)
                {
                    Declare.DatumDisketa = (Declare.VNUTORNY_DATUM_ROK * 10000) + (Declare.VNUTORNY_DATUM_MESIAC * 100) + Declare.VNUTORNY_DATUM_DEN;

                    if (HelpFunctions.IntYyyymmddToDate(Declare.DatumPC) < HelpFunctions.IntYyyymmddToDate(Declare.DatumDisketa))
                    {
                        MessageBox.Show(
                            "Nie je možné spustiť inštaláciu, pretože Vám vypršala lehota pre bezplatný upgrade programu OLYMP." +
                            Environment.NewLine + Environment.NewLine +
                            "Vyžiadajte si obnovu Balíka služieb v Kros a.s. alebo si ho objednajte on-line cez Zónu pre klienta.",
                            "OLYMP – inštalácia",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information);

                        Declare.ExitProg(0);
                        return;
                    }
                    // inak „má nárok“ – pokračuje ďalej
                }

                ProgresiaDone();
            }
            catch (Exception ex)
            {
                Declare.Errors.Add(ex.ToString());
                Declare.ExitProg(Declare.ID_CHYBA_LICENCIA);
            }
        }
             
        public void ProgresiaPreset(long pocet)
        {
            if (pocet > 0)
            {
                _PocetNaTik = ((double)pb.Maximum + 1.0) / pocet;
            }
            pb.Value = pb.Minimum;
            _AktTik = 0;
        }

        public void ProgresiaTik()
        {
            //Thread.Sleep(1);//iba kvoli tomu, aby to neprefrcalo tak rychlo

            _AktTik++;
            var val = (int)Math.Floor(_AktTik * _PocetNaTik);

            if (val > pb.Value && val <= pb.Maximum)
            {
                pb.Value = val;
            }
        }

        public void ProgresiaDone()
        {
            pb.Value = pb.Maximum;
        }

        /// <summary>
        /// Nastavi cestu, kde je nainstalovany olymp, zisti verzie upgradu a nainstalovaneho execka
        /// </summary>
        /// <param name="cesta"></param>
        private void NastavCestuVerzie(string cesta = "")
        {
            // zistím potrebné verzie
            lblUpgrade.Text = DajVerziuUpgrade();

            if (!RegistryFunctions.MamDostatocnyFramework())
            {
                Declare.ExitProg(Declare.ID_CHYBA_CHYBA_FRAMEWORK);
                return;
            }

            //ak nie je, je to chyba koncim
            if (!RegistryFunctions.CheckRegIfOlympIsInstalled(out string path))
            {
                Declare.ExitProg(Declare.ID_CHYBA_NIE_JE_MIN_VERZIA);
                return;
            }
            else
            {
                // uložím pôvodný adresár inštalácie z registrov
                Declare.DEST_PATH_INSTALLSHIELD = HelpFunctions.PridajLomitko(path);

                // načítam info z licencie
                CitajRegistracnySubor(Declare.PCD_INSTALACKY, HelpFunctions.PridajLomitko(Declare.AKT_ADRESAR));
                path = RegistryFunctions.GetOlympFolderBaseOnLic();
            }

            if (!string.IsNullOrEmpty(cesta))
            {
                // je to spustane z Olympu a cestu mam zadanu
                path = cesta;
            }

            bool prazdnyString;
            if (string.IsNullOrWhiteSpace(path))
            {
                prazdnyString = true;
                Declare.DEST_PATH_INSTALLSHIELD = string.Empty;
            }
            else
            {
                prazdnyString = false;
                Declare.DEST_PATH = HelpFunctions.PridajLomitko(path);
            }

            // skontrolujem, či je možné do adresára inštalovať
            if (!prazdnyString && !HelpFunctions.JeAdresarSpravny(Declare.DEST_PATH))
            {
                MessageBox.Show(
                    "Adresár " + Declare.DEST_PATH + " v ktorom je nainštalovaný program OLYMP je neprístupný. " +
                    "Nemáte právo do neho zapisovať alebo neexistuje." + Environment.NewLine + Environment.NewLine +
                    "Skontrolujte, či máte právo zapisovať do tohto adresára a potom spustite " + Declare.MENO_EXE +
                    " znova alebo vyberte iný adresár, v ktorom je program OLYMP nainštalovaný.",
                    Declare.TTITLE, MessageBoxButtons.OK, MessageBoxIcon.Error);

                Declare.DEST_PATH = Declare.DEST_PATH_NEZNAMY;
                btnOk.Enabled = false;
                NastavCestu(lblDest, Declare.DEST_PATH, true);
                lblUpgrade.Text = Declare.VERZIA_NEZNAMA;
                lblExe.Text = Declare.VERZIA_NEZNAMA;
                Declare.MAJOR = 0;
            }
            else if (prazdnyString)
            {
                Declare.DEST_PATH = Declare.DEST_PATH_NEZNAMY;
                btnOk.Enabled = false;
                NastavCestu(lblDest, Declare.DEST_PATH, true);
                lblUpgrade.Text = Declare.VERZIA_NEZNAMA;
                lblExe.Text = Declare.VERZIA_NEZNAMA;
                Declare.MAJOR = 0;
            }
            else
            {
                NastavCestu(lblDest, Declare.DEST_PATH, true);
                lblExe.Text = DajVerziuProgramu();
            }
        }

        /// <summary>
        /// vrati verziu olymp.exe
        /// </summary>
        /// <returns></returns>
        private string DajVerziuProgramu()
        {
            var pathExe = Path.Combine(Declare.DEST_PATH, Declare.SUBOR_EXE);
            var pathExeStary = Path.Combine(Declare.DEST_PATH, Declare.SUBOR_EXE_STARY);

            if (File.Exists(pathExe))
            {
                HelpFunctions.DajVerziuExe(pathExe, out Declare.P_MAJOR, out Declare.P_MINOR, out Declare.P_REVISION);
                if (Declare.P_MAJOR == 0)
                    return Declare.VERZIA_NEZNAMA;
                else
                    return HelpFunctions.DajVerziuString(Declare.P_MAJOR, Declare.P_MINOR, Declare.P_REVISION);
            }
            else if (File.Exists(pathExeStary))
            {
                HelpFunctions.DajVerziuExe(pathExeStary, out Declare.P_MAJOR, out Declare.P_MINOR, out Declare.P_REVISION);
                if (Declare.P_MAJOR == 0)
                    return Declare.VERZIA_NEZNAMA;
                else
                    return HelpFunctions.DajVerziuString(Declare.P_MAJOR, Declare.P_MINOR, Declare.P_REVISION);
            }

            Declare.P_MAJOR = 0;
            return Declare.VERZIA_NEZNAMA;
        }

        private void NastavCestu(Label lblDest, string path, bool tooltp = false)
        {
            lblDest.AutoSize = false;
            lblDest.AutoEllipsis = true;
            lblDest.Text = path;

            if (tooltp)
                toolTip1.SetToolTip(lblDest, path);
        }

        private void CitajRegistracnySubor(int PcD, string cesta, string cestaInstalacie = "")
        {
            cestaInstalacie = string.Empty;

            // Zisti, ci registracny subor existuje
            if (HelpFunctions.ExistujeSubor(Path.Combine(cesta, Declare.LICENCIA_SW)))
            {
                NacitajNovuLicenciu(PcD, Path.Combine(cesta, Declare.LICENCIA_SW));
                return;
            }

            //NEBOL NAJDENY REGISTRACNY SUBOR
            if (PcD == Declare.PCD_POCITAC) Declare.VerziaPC = 0; else Declare.VerziaDisketa = 0;
            return;
        }

        private void NacitajNovuLicenciu(int PcD, string cesta)
        {
            long datumLicencie = 0;
            bool prenajom = false;

            using (var sr = new StreamReader(cesta))//, Encoding.GetEncoding(1250)))
            {
                string riadok;
                while ((riadok = sr.ReadLine()) != null)
                {
                    // nahradiť \" -> "
                    riadok = riadok.Replace("\"\"", "\"");

                    string riadokUpper = riadok.ToUpperInvariant();
                    int colon = riadok.IndexOf(':');

                    if (colon < 0) continue;

                    if (riadokUpper.Contains("DATUMLICENCIE"))
                    {
                        //datumLicencie = Val(Mid(riadok, InStr(1, riadok, ":") + 3, 4) + Mid(riadok, InStr(1, riadok, ":") + 8, 2) + Mid(riadok, InStr(1, riadok, ":") + 11, 2))
                        string y = HelpFunctions.SafeSubstring(riadok, colon + 4, 4);
                        string m = HelpFunctions.SafeSubstring(riadok, colon + 9, 2);
                        string d = HelpFunctions.SafeSubstring(riadok, colon + 12, 2);
                        long.TryParse(y + m + d, out datumLicencie);
                    }
                    else if (riadokUpper.Contains("PRENAJOMSOFTVERU"))
                    {
                        //Val(Mid(riadok, InStr(1, riadok, ":") + 1, 2)) <> 0
                        string v = HelpFunctions.SafeSubstring(riadok, colon + 2, 1);
                        prenajom = (HelpFunctions.ValLong(v) != 0);
                    }
                    else if (riadokUpper.Contains("PARTNERICO"))
                    {
                        //If PcD = PCD_INSTALACKY Then
                        //    ICO_Disketa = Replace(Replace(Mid(riadok, InStr(1, riadok, ":") + 3, 20), """", ""), ",", "")
                        //Else
                        //    ICO_PC = Replace(Replace(Mid(riadok, InStr(1, riadok, ":") + 3, 20), """", ""), ",", "")
                        string s = HelpFunctions.SafeSubstring(riadok, colon + 3, 20).Replace("\"", "").Replace(",", "").Replace("\\", "");
                        if (PcD == Declare.PCD_INSTALACKY)
                            Declare.ICO_Disketa = s;
                        else
                            Declare.ICO_PC = s;
                    }
                    else if (riadokUpper.Contains("PARTNERMENO"))
                    {
                        //If PcD = PCD_INSTALACKY Then
                        //    NazovFirmyDisketa = Mid(riadok, InStr(1, riadok, ":") + 3, Len(riadok) - 4 - InStr(1, riadok, ":"))
                        //Else
                        //    NazovFirmyPC = Mid(riadok, InStr(1, riadok, ":") + 3, Len(riadok) - 4 - InStr(1, riadok, ":"))
                        //End If
                        int len = Math.Max(0, riadok.Length - 4 - colon - 3);
                        string s = HelpFunctions.SafeSubstring(riadok, colon + 4, len);
                        if (PcD == Declare.PCD_INSTALACKY)
                            Declare.NazovFirmyDisketa = s;
                        else
                            Declare.NazovFirmyPC = s;
                    }
                }
            }

            if (datumLicencie != 0)
            {
                // ak to nie je prenájom, posuň platnosť o rok (yyyymmdd + 10000)
                if (!prenajom) datumLicencie += 10000;

                if (PcD == Declare.PCD_INSTALACKY)
                {
                    Declare.DatumDisketa = datumLicencie;
                    Declare.VerziaDisketa = Declare.LIC_OSTRA;
                }
                else
                {
                    Declare.DatumPC = datumLicencie;
                    Declare.VerziaPC = Declare.LIC_OSTRA;
                }
            }
        }

        /// <summary>
        /// vrati verziu upgrade a nastavi dalsie hodnoty zo suboru txt
        /// </summary>
        /// <returns></returns>
        public string DajVerziuUpgrade()
        {
            string tempAdr = string.Empty;

            // Skontrolujeme, či existuje ZIP súbor
            if (!File.Exists(Path.Combine(Declare.AKT_ADRESAR, Declare.SUBOR_ZIP)))
            {
                Declare.ExitProg(Declare.ID_CHYBA_NIE_JE_ZIP);
                return string.Empty;
            }

            tempAdr = HelpFunctions.DajTemp(Declare.AKT_ADRESAR);
            if (string.IsNullOrEmpty(tempAdr))
            {
                if (HelpFunctions.JeAdresarSpravny(Declare.DEST_PATH))
                    tempAdr = Declare.DEST_PATH;
            }

            if (string.IsNullOrEmpty(tempAdr))
                return string.Empty;

            //rozbalim VERZIA_TXT do tempu
            var verziaTxtPath = ExtractFileFromUpgradeZip(tempAdr, Declare.VERZIA_TXT);
            if (string.IsNullOrEmpty(verziaTxtPath))
            {
                Declare.ExitProg(Declare.ID_CHYBA_CHYBA_ZIP);
                return string.Empty;
            }

            // Čítanie VERZIA_TXT súboru a získavanie údajov
            if (!ReadVersionTxt(verziaTxtPath))
            {
                Declare.ExitProg(Declare.ID_CHYBA_CHYBA_ZIP);
                return string.Empty;
            }

            tempAdr = HelpFunctions.DajSystemTemp();
            if (string.IsNullOrEmpty(tempAdr))
                return string.Empty;

            // Extrahovanie ďalšieho súboru CRV2Kros.exe
            var cRV2KrosFile = ExtractFileFromUpgradeZip(tempAdr, "CRV2Kros.exe");
            if (string.IsNullOrEmpty(cRV2KrosFile))
            {
                Declare.ExitProg(Declare.ID_CHYBA_CHYBA_ZIP2);
                return string.Empty;
            }

            HelpFunctions.DajVerziuExe(Path.Combine(tempAdr, "CRV2Kros.exe"), out Declare.N_CRV2_MAJOR, out Declare.N_CRV2_MINOR, out Declare.N_CRV2_REVISION);
            HelpFunctions.TryToDeleteFile(cRV2KrosFile);

            return HelpFunctions.DajVerziuString(Declare.MAJOR, Declare.MINOR, Declare.REVISION);
        }

        private bool ReadVersionTxt(string verzieTxtPath)
        {
            try
            {
                string s;
                int i;
                using (StreamReader reader = new StreamReader(verzieTxtPath))
                {
                    s = reader.ReadLine();
                    i = s.IndexOf("=");
                    Declare.MAJOR = int.Parse(s.Substring(i + 1));

                    s = reader.ReadLine();
                    i = s.IndexOf("=");
                    Declare.MINOR = int.Parse(s.Substring(i + 1));

                    s = reader.ReadLine();
                    i = s.IndexOf("=");
                    Declare.REVISION = int.Parse(s.Substring(i + 1));

                    // Skipping 3 lines
                    for (int j = 0; j < 3; j++)
                        reader.ReadLine();

                    s = reader.ReadLine();
                    i = s.IndexOf("=");
                    Declare.MIN_MAJOR = int.Parse(s.Substring(i + 1));

                    s = reader.ReadLine();
                    i = s.IndexOf("=");
                    Declare.MIN_MINOR = int.Parse(s.Substring(i + 1));

                    s = reader.ReadLine();
                    i = s.IndexOf("=");
                    Declare.MIN_REVISION = int.Parse(s.Substring(i + 1));

                    // Skipping 3 lines
                    for (int j = 0; j < 3; j++)
                        reader.ReadLine();

                    s = reader.ReadLine();
                    i = s.IndexOf("=");
                    Declare.VNUTORNY_DATUM_DEN = int.Parse(s.Substring(i + 1));

                    s = reader.ReadLine();
                    i = s.IndexOf("=");
                    Declare.VNUTORNY_DATUM_MESIAC = int.Parse(s.Substring(i + 1));

                    s = reader.ReadLine();
                    i = s.IndexOf("=");
                    Declare.VNUTORNY_DATUM_ROK = int.Parse(s.Substring(i + 1));
                }

                return true;
            }
            catch (Exception ex)
            {
                Declare.Errors.Add(ex.ToString());
                return false;
            }
            finally { HelpFunctions.TryToDeleteFile(verzieTxtPath); }
        }

        private string ExtractFileFromUpgradeZip(string desAdr, string fileName)
        {
            string resFilePath = Path.Combine(desAdr, fileName);
            try
            {
                using (ZipArchive zip = ZipFile.Open(Path.Combine(Declare.AKT_ADRESAR, Declare.SUBOR_ZIP), ZipArchiveMode.Read, _zipEncoding))
                {
                    var entry = zip.Entries//zip.GetEntry(fileName);
                           .FirstOrDefault(e => string.Equals(e.FullName, fileName,
                                                                StringComparison.OrdinalIgnoreCase));

                    entry.ExtractToFile(resFilePath, overwrite: true);
                }

                if (!string.IsNullOrWhiteSpace(resFilePath)
                    && File.Exists(resFilePath))
                    return resFilePath;
                else
                    return string.Empty;
            }
            catch (Exception ex)
            {
                Declare.Errors.Add(ex.ToString());
                HelpFunctions.TryToDeleteFile(resFilePath);
                return string.Empty;
            }
        }

        private void HideTabs()
        {
            tabControl1.Appearance = TabAppearance.FlatButtons;
            tabControl1.ItemSize = new Size(0, 1);
            tabControl1.SizeMode = TabSizeMode.Fixed;
        }

        private void NastavRok()
        {
            lblCopyright.Text = "\u00A9 1996 - " + DateTime.Now.Year + " KROS a.s.";
        }

        private void Koniec()
        {
            tabControl1.SelectTab(2);//pnl[4]

            _Stav = 2;

            btnDalej.Text = "< S&päť";
            btnDalej.Visible = true;

            btnOk.Text = "&Spusti";

            if (Declare.RESTARTUJ)
            {
                btnOk.Visible = false;
                lblInfo.Visible = false;
            }

            if (_BoliChyby)
            {
                btnOk.Visible = false;
                lblInfo.Text =
                    "Inštalácia prebehla s chybami. Pokúste sa tieto chyby odstrániť a potom spustite program " +
                    /*Path.Combine(Declare.AKT_ADRESAR ?? string.Empty, Declare.MENO_EXE ?? string.Empty) +*/ " znova. " +
                    "Ak sa vyskytla chyba typu \"nedá sa prepísať\", príčinou môže byť to, že je súbor " +
                    "otvorený alebo spustený. V tomto prípade musíte súbor zavrieť alebo ukončiť.";
                lblInfo.Visible = true;
            }

            btnStorno.Text = "&Koniec";

            if (btnOk.Visible)
                btnOk.Focus();
        }

        /// <summary>
        /// skontroluje, ci je v systeme nainstalovany acrobat reader, ak nie je tak ho na otazku nainstaluje
        /// </summary>
        /// <returns></returns>
        private bool CheckAcrobatAndInstall()
        {
            if (Declare.KTO_VOLAL != Declare.VOLAL_INSTALL) return false;

            string basePath = HelpFunctions.DajCestuOXAdresarovVyssie(Declare.AKT_ADRESAR, 3);
            string cesta = Path.Combine(basePath, Declare.CestaAcrobat, Declare.SuborAcrobat);

            if (RegistryFunctions.IsAcrobatReaderInstalled() == 0 && File.Exists(cesta))
            {
                string prompt =
                    "Inštalátor zistil, že vo vašom počítači nie je nainštalovaný program Adobe Acrobat Reader, " +
                    "ktorý program OLYMP využíva na tlač originálnych tlačív daňového úradu. " +
                    "Chcete ho teraz nainštalovať?";

                if (MessageBox.Show(prompt, Declare.TTITLE, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    // Zruš TOPMOST (ekvivalent SetWindowPos Me.hwnd, -2, 0,0,0,0, 1 Or 2)
                    //SetWindowPos(this.Handle, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE);
                    this.TopMost = false;

                    var psi = new ProcessStartInfo
                    {
                        FileName = cesta,
                        WorkingDirectory = Path.GetDirectoryName(cesta),
                        UseShellExecute = true // ShellExecute "open"
                    };
                    Process.Start(psi);
                    return true;
                }
            }
            return false;
        }

        #endregion Methods
    }
}

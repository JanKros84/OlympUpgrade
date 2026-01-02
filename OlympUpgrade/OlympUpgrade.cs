using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.IO.Compression;
using System.Linq;
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
                        LicenseVersionFunctions.DajVerziuExe(crv2Path, out P_CRV2_MAJOR, out P_CRV2_MINOR, out P_CRV2_REVISION);

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
                    lblExe.Text = LicenseVersionFunctions.DajVerziuProgramu();
                    btnOk.Enabled = true;
                    this.Text = $"OLYMP {lblUpgrade.Text} - Sprievodca inštaláciou";
                }
                else
                {
                    lblExe.Text = LicenseVersionFunctions.DajVerziuProgramu();
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

            //kopiruj subory
            var copyUpgradeFiles = new CopyUpgradeFilesManager(ProgresiaPreset, ProgresiaTik, ProgresiaDone, SetLabelInfo);
            copyUpgradeFiles.KopirujSubory();
            ZapisVysledokDoListu(copyUpgradeFiles);
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
        /// Zapise vysledok instalacie do listboxu
        /// </summary>
        private void ZapisVysledokDoListu(CopyUpgradeFilesManager copyUpgradeFiles)
        {
            lst.Items.Clear();
            lst.Items.Add(string.Empty);

            if (copyUpgradeFiles.Chybne == null) return;
            if (copyUpgradeFiles.Preskocene == null) return;
            if (copyUpgradeFiles.Spravne == null) return;

            if (copyUpgradeFiles.Chybne.Count == 0)
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
            lst.Items.Add("   Počet neinštalovaných kvôli chybe: \t" + copyUpgradeFiles.Chybne.Count);
            lst.Items.Add("   Počet preskočených: \t\t" + copyUpgradeFiles.Preskocene.Count);
            lst.Items.Add("   Počet správne inštalovaných: \t" + copyUpgradeFiles.Spravne.Count);

            if (copyUpgradeFiles.Chybne.Count > 0)
            {
                lst.Items.Add(string.Empty);
                lst.Items.Add("== Neinštalované komponenty kvôli chybe ========================");
                lst.Items.Add(string.Empty);
                foreach (var s in copyUpgradeFiles.Chybne)
                    lst.Items.Add("   " + s);
            }

            if (copyUpgradeFiles.Preskocene.Count > 0)
            {
                lst.Items.Add(string.Empty);
                lst.Items.Add("== Preskočené komponenty ====================================");
                lst.Items.Add(string.Empty);
                foreach (var s in copyUpgradeFiles.Preskocene)
                    lst.Items.Add("   " + s);
            }

            if (copyUpgradeFiles.Spravne.Count > 0)
            {
                lst.Items.Add(string.Empty);
                lst.Items.Add("== Inštalované komponenty ====================================");
                lst.Items.Add(string.Empty);
                foreach (var s in copyUpgradeFiles.Spravne)
                    lst.Items.Add("   " + s);
            }
        }

        public void OverVolneMiesto()
        {
            double VolneMiesto, PotrebnaVelkost;
            string prompt = string.Empty;

            ProgresiaPreset(0);

            lblDestFile.Text = "Kontrola voľného miesta ...";
            lblDestFile.Refresh();

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

            using (ZipArchive zip = ZipFile.Open(pathZip, ZipArchiveMode.Read, Declare.ZipEncoding))
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
                LicenseVersionFunctions.CitajRegistracnySubor(Declare.PCD_POCITAC, Declare.DEST_PATH);
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

        public void SetLabelInfo(string text)
        {
            lblDestFile.Text = text;
            lblDestFile.Refresh();
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
                LicenseVersionFunctions.CitajRegistracnySubor(Declare.PCD_INSTALACKY, HelpFunctions.PridajLomitko(Declare.AKT_ADRESAR));
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
                lblExe.Text = LicenseVersionFunctions.DajVerziuProgramu();
            }
        }

        private void NastavCestu(Label lblDest, string path, bool tooltp = false)
        {
            lblDest.AutoSize = false;
            lblDest.AutoEllipsis = true;
            lblDest.Text = path;

            if (tooltp)
                toolTip1.SetToolTip(lblDest, path);
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
            if (!LicenseVersionFunctions.ReadVersionTxt(verziaTxtPath))
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

            LicenseVersionFunctions.DajVerziuExe(Path.Combine(tempAdr, "CRV2Kros.exe"), out Declare.N_CRV2_MAJOR, out Declare.N_CRV2_MINOR, out Declare.N_CRV2_REVISION);
            HelpFunctions.TryToDeleteFile(cRV2KrosFile);

            return LicenseVersionFunctions.DajVerziuString(Declare.MAJOR, Declare.MINOR, Declare.REVISION);
        }
           
        private string ExtractFileFromUpgradeZip(string desAdr, string fileName)
        {
            string resFilePath = Path.Combine(desAdr, fileName);
            try
            {
                using (ZipArchive zip = ZipFile.Open(Path.Combine(Declare.AKT_ADRESAR, Declare.SUBOR_ZIP), ZipArchiveMode.Read, Declare.ZipEncoding))
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

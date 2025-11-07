using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net.NetworkInformation;
using System.Text;
using System.Threading.Tasks;
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

        public OlympUpgrade()
        {
            InitializeComponent();
        }

        private void OlympUpgrade_Load(object sender, EventArgs e)
        {
            try
            {
                int i;
                List<string> poleArg;
                string pom = string.Empty;

                HideTabs();

                tabControl1.SelectTab(0);
                _Stav = 0;

                Declare.AKT_ADRESAR = AppDomain.CurrentDomain.BaseDirectory;

                Declare.RESTARTUJ = false;
                Declare.KOPIRUJ_LICENCIU = 0;

                //zaregistrujChilkat();

                //this.Height = FRM_HEIGHT;
                //this.Width = FRM_WIDTH;
                //this.Top = (Screen.Height - this.Height) / 2;
                //this.Left = (Screen.Width - this.Width) / 2;
                NastavRok();

                // zistime argument, s ktorym bol program spustany

                poleArg = Environment.GetCommandLineArgs().ToList();
                poleArg.RemoveAt(0);

                // prvy je parameter, potom je restartuj (true, false) a potom nasleduje instalacny adresar
                if (poleArg.Count >= 3)
                {
                    Declare.COMMAND_LINE_ARGUMENT = poleArg[0];

                    if (poleArg[1].ToUpper() == "TRUE")
                        Declare.RESTARTUJ = true;


                    int.TryParse(poleArg[2], out Declare.KOPIRUJ_LICENCIU);

                    pom = string.Empty;
                    for (i = 3; i < poleArg.Count; i++)
                        pom = pom = pom + poleArg[i] + " "; //string.Concat(pom, poleArg[i], " ");

                    pom = pom.TrimEnd();
                }

                // MsgBox "zacinam " & COMMAND_LINE_ARGUMENT
                // sustil installshield
                if (Declare.COMMAND_LINE_ARGUMENT == Declare.PARAMETER_INSTAL)
                {
                    Declare.TTITLE = "Sprievodca inštaláciou OLYMP";
                    Declare.KTO_VOLAL = Declare.VOLAL_INSTALL;
                    OdstranZRegistrovDoinstalovanie();
                    Declare.DEST_PATH = Declare.PridajLomitko(pom);

                    tabControl1.SelectTab(1);

                    lblUpgrade.Text = DajVerziuUpgrade();
                }

                //else if (COMMAND_LINE_ARGUMENT == SUBOR_STARA_INSTALACIA_EXE | COMMAND_LINE_ARGUMENT == SUBOR_EXE_STARY)
                //{
                //    // spustil uzivatel
                //    TTITLE = "OlympUpgrade";
                //    KTO_VOLAL = VOLAL_OLYMP;
                //    cmdChDir.Enabled = false;      // ak to spustal OLYMP, tak nema co menit adresar
                //    NastavCestuVerzie(PridajLomitko(pom));
                //}
                //else
                //{
                //    // spustil uzivatel
                //    TTITLE = "OlympUpgrade";
                //    KTO_VOLAL = VOLAL_UZIVATEL;
                //    NastavCestuVerzie();
                //}

                this.Text = "OLYMP " + lblUpgrade.Text + " - Sprievodca inštaláciou";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// vrati verziu upgrade a nastavi dalsie hodnoty zo suboru txt
        /// </summary>
        /// <returns></returns>
        public string DajVerziuUpgrade()
        {
            string tempAdr = string.Empty;
            string s;
            int i;
            long chyba = 0;

            // Skontrolujeme, či existuje ZIP súbor
            if (!File.Exists(Path.Combine(Declare.AKT_ADRESAR, Declare.SUBOR_ZIP)))
            {
                Declare.ExitProg(Declare.ID_CHYBA_NIE_JE_ZIP);
                return string.Empty;
            }

            tempAdr = Declare.DajTemp(Declare.AKT_ADRESAR);
            if (string.IsNullOrEmpty(tempAdr))
            {
                if (Declare.JeAdresarSpravny(Declare.DEST_PATH))
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

            tempAdr = Declare.DajSystemTemp();
            if (string.IsNullOrEmpty(tempAdr))
                return string.Empty;

            // Extrahovanie ďalšieho súboru CRV2Kros.exe
            var cRV2KrosFile = ExtractFileFromUpgradeZip(tempAdr, "CRV2Kros.exe");
            if (string.IsNullOrEmpty(cRV2KrosFile))
            {
                Declare.ExitProg(Declare.ID_CHYBA_CHYBA_ZIP2);
                return string.Empty;
            }

            Declare.DajVerziuExe(Path.Combine(tempAdr, "CRV2Kros.exe"), out Declare.N_CRV2_MAJOR, out Declare.N_CRV2_MINOR, out Declare.N_CRV2_REVISION);
            TryToDeleteFile(cRV2KrosFile);

            return Declare.DajVerziuString(Declare.MAJOR, Declare.MINOR, Declare.REVISION);
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
            catch (Exception) //TODO Logger
            {
                return false;
            }
            finally { TryToDeleteFile(verzieTxtPath); }
        }

        private string ExtractFileFromUpgradeZip(string desAdr, string fileName)
        {
            string resFilePath = Path.Combine(desAdr, fileName);
            try
            {
                using (ZipArchive zip = ZipFile.OpenRead(Path.Combine(Declare.AKT_ADRESAR, Declare.SUBOR_ZIP)))
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
            catch
            {
                TryToDeleteFile(resFilePath);
                return string.Empty;
            }
        }

        private static void TryToDeleteFile(string filePath)
        {
            try
            {
                if (!string.IsNullOrWhiteSpace(filePath)
                    && File.Exists(filePath))
                    File.Delete(filePath);
            }
            catch { /*TODO ???*/ }
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

        /// <summary>
        /// odstrani z registrov spustenie tohoto programu po restarte z instalacneho CD ak tam nic nie je tak sa nic nestane
        /// </summary>
        private void OdstranZRegistrovDoinstalovanie()
        {
            Declare.ZmazHodnotuReg(Registry.CurrentUser/*Declare.HKEY_CURRENT_USER*/, @"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "krosOlymp");

            if (Declare.RESTARTUJ)
                Declare.RESTARTUJ = false;
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
                            if (!Declare.Restart())
                            {
                                MessageBox.Show(
                                    "Nepodarilo sa reštartovať počítač. Skúste počítač reštartovať manuálne.",
                                    "Informácia",
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Information
                                );
                            }
                            else
                                this.Close();// UkonciProgram();
                            
                        }

                        this.Close();
                    }
                    else
                        this.Close();
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
                MessageBox.Show(ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void Koniec()
        {
            // 6500 twips posunie panel "mimo" – prepočet twips -> pixely
            //pnl[5].Top = TwipsToPixelsY(6500);
            //pnl[4].Top = 0;
            //pnl[4].BringToFront();

            _Stav = 2;

            btnDalej.Text = "< S&päť";
            btnDalej.Visible = true;

            btnOk.Text = "&Spusti";

            if (Declare.RESTARTUJ)
            {
                btnOk.Visible = false;
                lbl[7].Visible = false;
            }

            if (_BoliChyby)
            {
                btnOk.Visible = false;
                lbl[7].Text =
                    "Inštalácia prebehla s chybami. Pokúste sa tieto chyby odstrániť a potom spustite program " +
                    Path.Combine(Declare.AKT_ADRESAR ?? string.Empty, Declare.MENO_EXE ?? string.Empty) + " znova. " +
                    "Ak sa vyskytla chyba typu \"nedá sa prepísať\", príčinou môže byť to, že je súbor " +
                    "otvorený alebo spustený. V tomto prípade musíte súbor zavrieť alebo ukončiť.";
                lbl[7].Visible = true;
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
            Declare.RESTARTUJ = false;
            if (Declare.KTO_VOLAL != Declare.VOLAL_INSTALL) return false;

            string basePath = Declare.DajCestuOXAdresarovVyssie(Declare.AKT_ADRESAR, 3);
            string cesta = Path.Combine(basePath, Declare.CestaAcrobat, Declare.SuborAcrobat);

            if (Declare.IsAcrobatReaderInstalled() == 0 && File.Exists(cesta))
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
    }
}

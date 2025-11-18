using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Security.AccessControl;
using System.Windows.Forms;

namespace OlympUpgrade
{
    internal class Declare
    {
        public static List<string> Errors = new List<string>();

        public const string URL_KROS = "https://www.kros.sk/";

        public const string PRODUCT_CODE_1 = "{38BB475B-CE35-4D84-8261-7FF5705CC4F4}";  // 13.30
        //public const string REG_PRODUCT_CODE_1 = "B574BB8353EC48D42816F75F07C54C4F";
        public const string PRODUCT_CODE_2 = "{E632280B-3D7A-41B5-8F20-79895AFFAC3F}";  // 15.00
        public const string REG_PRODUCT_CODE_2 = "B082236EA7D35B14F8029798A5FFCAF3";

        /// <summary>
        /// Ci sa bude musiet restartovat, alebo nie, predava sa to ako druhy parameter pri spustani
        /// </summary>
        public static bool RESTARTUJ;

        /// <summary>
        /// parameter, s ktorym sa program pusta pri installshielde
        /// </summary>
        public const string PARAMETER_INSTAL = "zInstalu";

        //public const string PARAMETER_OLYMP = "zOlympu";                 // parameter, s ktorym sa upgrade spusta z Olympu

        /// <summary>
        /// program spustil installshield
        /// </summary>
        public const int VOLAL_INSTALL = 1;

        /// <summary>
        /// program spustil uzivatel
        /// </summary>
        public const int VOLAL_UZIVATEL = 2;

        /// <summary>
        /// program sa spustil automaticky
        /// </summary>
        public const int VOLAL_OLYMP = 3;

        public const string VERZIA_NEZNAMA = "(Nedefinované)";

        public const string DEST_PATH_NEZNAMY = "Vyberte inštalačný adresár.";

        public const string SUBOR_EXE = "Olymp.exe";

        public const string SUBOR_EXE_STARY = "Mzdy.exe";

        public const string SUBOR_TeamViewer = "TeamViewerQS.exe";

        public const string SUBOR_ClickYes = "ClickYes.exe";

        public const string SUBOR_STARA_INSTALACIA_EXE = "Olymp.exe";

        public const string VERZIA_TXT = "Version.txt";

        public const string SUBOR_ZIP = "OlympUpgrade.zip";

        public const string SUBOR_INST = "OlympInst.exe";

        public const string SUBOR_INST_NEW = "OlympInstall.exe";
        // Global Const TEMP = "TEMPINST"

        public const string LICENCIA = "Mzdy.lic";

        public const string LICENCIA_SW = "Olymp.license";

        public const string MENO_EXE = "OlympUpgrade.exe";
        // Global Const MENO_JU_LDB = "JU.LDB"


        public const string CestaAcrobat = @"!Servis\Acrobat Reader\";

        public const string SuborAcrobat = "AdbeRdr705_cze_full.exe";


        public const string FILE_REPORTY_TXT = "reporty.txt";

        //public const string FILE_REPORTY_DEVEXPRES_TXT = "reportyDX.txt";

        public const string FILE_REPORTY_PDF_TXT = "pdf.txt";

        public const string FILE_REPORTY_EXCEL_TXT = "excel.txt";

        public const string FILE_REPORTY_EXCEL_P_TXT = "excel.txt";

        public static string TTITLE;
        public static string AKT_ADRESAR;

        // verzia upgrade
        public static int MAJOR;
        public static int MINOR;
        public static int REVISION;

        // min verzia instalovanej alfy, aby mohol byt vykonany upgrade
        public static int MIN_MAJOR;
        public static int MIN_MINOR;
        public static int MIN_REVISION;

        // verzia nainstalovaneho programu
        public static int P_MAJOR;
        public static int P_MINOR;
        public static int P_REVISION;

        public static int N_CRV2_MAJOR;
        public static int N_CRV2_MINOR;
        public static int N_CRV2_REVISION;


        // vnutorny datum z verzie.txt
        public static int VNUTORNY_DATUM_DEN;
        public static int VNUTORNY_DATUM_MESIAC;
        public static int VNUTORNY_DATUM_ROK;

        /// <summary>
        /// O AKY TYP INSTALACIE SA JEDNA NA INSTALACKACH; 0=demo,  1=ostra, 2=upgrate
        /// </summary>
        public static int VerziaDisketa;

        /// <summary>
        /// AKY TYP INSTALACIE JE NA POCITACI; 0=demo,  1=ostra, -1=ostra chybny .reg, 2=nic
        /// </summary>
        public static int VerziaPC;

        /// <summary>
        /// zakodovany datum registracie z instalacnych diskiet /RRRRMMDD
        /// </summary>
        public static long DatumDisketa;

        /// <summary>
        /// zakodovany datum registracie z PC /RRRRMMDD
        /// </summary>
        public static long DatumPC;

        //public static int mojTag;

        public static string NazovFirmyDisketa;

        /// <summary>
        /// ico na krotu firmu je registrovana na diskete
        /// </summary>
        public static string ICO_Disketa;

        /// <summary>
        /// poradove cislo registracie na diskete _
        /// </summary>
        public static string PorCisloDisketa;

        public static string NazovFirmyPC;

        /// <summary>
        /// ico na krotu firmu je registrovana na diskete
        /// </summary>
        public static string ICO_PC;

        /// <summary>
        /// poradove cislo registracie na diskete _
        /// </summary>
        public static string PorCisloPC;


        // identifikatory chyb

        public const int ID_CHYBA_CHYBA_ZIP = 1;

        public const int ID_CHYBA_NIE_JE_MIN_VERZIA = 2;

        public const int ID_CHYBA_LICENCIA = 3;

        public const int ID_CHYBA_NIE_JE_ZIP = 4;

        public const int ID_CHYBA_CHYBA_ZIP2 = 5;

        public const int ID_CHYBA_CHYBA_FRAMEWORK = 6;

        /// <summary>
        /// kto volal tento porgram
        /// </summary>
        public static int KTO_VOLAL;

        /// <summary>
        /// cielovy adresar instalacie
        /// </summary>
        public static string DEST_PATH;

        /// <summary>
        /// adresar, kde bola nainstalovana prvy krat alfa z installshieldu
        /// </summary>
        public static string DEST_PATH_INSTALLSHIELD;

        /// <summary>
        /// parameter, s ktorym sa program spustil
        /// </summary>
        public static string COMMAND_LINE_ARGUMENT;

        /// <summary>
        /// parameter z instalshieldu, ci sa ma lic skopcit alebo nie, ak je daky konflikt
        /// </summary>
        public static int KOPIRUJ_LICENCIU;


        public const int PCD_POCITAC = 0;
        public const int PCD_INSTALACKY = 1;

        //public const int LIC_CHYBNA = -1;
        public const int LIC_NEEXISTUJE = 0;
        public const int LIC_OSTRA = 1;
        //public const int LIC_UPDATE = 2;

        public static bool ZmazHodnotuReg(RegistryKey skupina_kluca, string meno_kluca, string meno_polozky)
        {
            try
            {
                using (RegistryKey hKey = skupina_kluca.OpenSubKey(meno_kluca, true))//RegistryRights.Delete))//FullControl))
                {
                    if (hKey != null)
                    {
                        hKey.DeleteValue(meno_polozky);//, false);
                        return true;
                    }
                }
                return false;
            }
            catch (Exception ex)
            {
                Declare.Errors.Add(ex.ToString());
                return false;
            }
        }

        /// <summary>
        /// cita hodnotu z registrov - vysledna precitana hodnota je vo vystup
        /// </summary>
        /// <param name="skupina_kluca"></param>
        /// <param name="meno_kluca"></param>
        /// <param name="meno_polozky"></param>
        /// <param name="vystup">out value</param>
        /// <returns></returns>
        public static bool CitajHodnotuReg(RegistryKey skupina_kluca, string meno_kluca, string meno_polozky, out object vystup)
        {
            vystup = null;
            try
            {
                using (RegistryKey hKey = skupina_kluca.OpenSubKey(meno_kluca, RegistryRights.QueryValues))
                {
                    if (hKey != null)
                    {
                        vystup = hKey.GetValue(meno_polozky);
                        return true;
                    }
                }
                return false;
            }
            catch (Exception ex) { Declare.Errors.Add(ex.ToString()); return false; }
        }

        public static bool Vytvor_kluc(RegistryKey skupina_kluca, string meno_kluca)
        {
            try
            {
                using (RegistryKey hKey = skupina_kluca.CreateSubKey(meno_kluca, RegistryKeyPermissionCheck.ReadWriteSubTree))
                {
                    return hKey == null ? false : true;
                }
            }
           catch (Exception ex)
            {
                Declare.Errors.Add(ex.ToString());
                return false;
            }
        }

        public static void ZapisHodnotuReg(RegistryKey skupina_kluca, string meno_kluca, string meno_polozky, string hodnota_polozky)//object hodnota_polozky, int typ_polozky)
        {
            try
            {
                using (RegistryKey hKey = skupina_kluca.OpenSubKey(meno_kluca, true))
                {
                    if (hKey == null)
                        return;// ERROR_FILE_NOT_FOUND;

                    hKey.SetValue(meno_polozky, hodnota_polozky, RegistryValueKind.String);
                }
            }
           catch( Exception ex)
            {
                Declare.Errors.Add(ex.ToString());
            }
        }

        public static bool ExistujeKlucReg(RegistryKey skupina_kluca, string meno_kluca)
        {
            try
            {
                using (RegistryKey hKey = skupina_kluca.OpenSubKey(meno_kluca, RegistryRights.QueryValues))
                    return hKey != null;
            }
            catch (Exception ex) { Declare.Errors.Add(ex.ToString()); return false; }
        }

        public static bool MamDostatocnyFramework()
        {
            if (CitajHodnotuReg(Registry.LocalMachine,
                                    @"SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full",
                                    "Release",
                                    out object val))
            {
                if (val is int verzia
                    && verzia >= 528040) // 528040 -> .NET Framework 4.8
                    return true;
            }

            return false;
        }

        public static string PridajLomitko(string path)
        {
            if (!string.IsNullOrEmpty(path) && !path.EndsWith("\\"))
                return path + "\\";

            return path;
        }

        /// <summary>
        /// ci existuje subor
        /// </summary>
        /// <param name="meno"></param>
        /// <returns></returns>
        public static bool ExistujeSubor(string meno)
        {
            if (string.IsNullOrWhiteSpace(meno))
                return false;

            if (meno.EndsWith("\\"))
                meno = meno.Substring(0, meno.Length - 1);

            try
            {
                return File.Exists(meno);
            }
            catch (Exception ex) { Declare.Errors.Add(ex.ToString()); return false; }
        }

        /// <summary>
        /// Globalny handler na chyby, ak sa dostane sem, program vzdy konci, su to fatalne chyby
        /// </summary>
        /// <param name="errNumber"></param>
        public static void ExitProg(long errNumber)
        {
            switch (errNumber)
            {
                case ID_CHYBA_NIE_JE_ZIP:
                    MessageBox.Show($"Súbor {SUBOR_ZIP} neexistuje v aktuálnom adresári.\r\n\r\n" +
                                    "Nie je čo inštalovať.",
                        TTITLE,
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    break;

                case ID_CHYBA_CHYBA_ZIP:
                    MessageBox.Show($"Nastala chyba pri dekomprimovaní súboru {SUBOR_ZIP}.\r\n\r\n" +
                        $"Súbor môže byť poškodený alebo cieľový adresár {DEST_PATH} je neprístupný.",
                        TTITLE,
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    break;

                case ID_CHYBA_CHYBA_ZIP2:
                    MessageBox.Show($"Nastala chyba pri dekomprimovaní súboru {SUBOR_ZIP}.\r\n\r\n" +
                        $"Súbor môže byť poškodený, alebo cieľový adresár {DEST_PATH} je neprístupný.",
                        TTITLE,
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    break;

                case ID_CHYBA_NIE_JE_MIN_VERZIA:
                    MessageBox.Show("Inštalácia programu OLYMP nemôže pokračovať, pretože nemáte nainštalovanú minimálnu požadovanú verziu.\r\n" +
                        "Ak chcete, aby bol program OLYMP nainštalovaný, musíte si spustiť plnú inštaláciu.",
                        TTITLE,
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                    break;

                case ID_CHYBA_LICENCIA:
                    MessageBox.Show("Nastala chyba pri kontrole licencie.",
                        TTITLE,
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    break;

                case ID_CHYBA_CHYBA_FRAMEWORK:
                    MessageBox.Show("Inštalácia programu OLYMP nemôže pokračovať, pretože nemáte nainštalovaný .NET Framework 4.8 alebo vyšší.\r\n\r\n" +
                        "Ak chcete, aby bol program OLYMP nainštalovaný, musíte si spustiť plnú inštaláciu OLYMPu, alebo nainštalovať .NET Framework ručne.",
                        TTITLE,
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                    break;

                default:
                    Console.WriteLine("Neznáma chyba.");
                    break;
            }

            MessageBox.Show("Inštalácia komponentov programu OLYMP bude ukončená.\r\n\r\n" +
                        "Samotná inštalácia ešte neprebehla, takže program OLYMP nebol modifikovaný.",
                        TTITLE,
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);

            System.Windows.Forms.Application.Exit();
        }

        /// <summary>
        /// vrati adresar, ktory moze byt pouzity ako temp aj s lomitkom, ak ziadny nenajde, vrati ""
        /// </summary>
        /// <param name="aktAdresar"></param>
        /// <returns></returns>
        public static string DajTemp(string aktAdresar)
        {
            string res = DajSystemTemp();

            if (JeAdresarSpravny(PridajLomitko(aktAdresar)))
            {
                return PridajLomitko(aktAdresar);
            }
            else if (!string.IsNullOrEmpty(res))
            {
                return res;
            }
            else if (JeAdresarSpravny(GetWindowsDir()))
            {
                return GetWindowsDir();
            }
            else
            {
                return string.Empty;
            }
        }

        /// <summary>
        /// vrati systemovy temp (user) adresar aj s lomitkom, ak ho nanejde, alebo nie je spravny, tak vrati ""
        /// </summary>
        /// <returns></returns>
        public static string DajSystemTemp()
        {
            string path = Path.GetTempPath();

            if (Directory.Exists(path) && JeAdresarSpravny(path))
                return path;
            else
                return string.Empty;
        }

        /// <summary>
        /// zistuje, ci je spravny adresar - teda ci existuje a ci sa da don zapisovat
        /// </summary>
        /// <param name="path"></param>
        /// <param name="ajZapis"></param>
        /// <returns></returns>
        public static bool JeAdresarSpravny(string path, bool ajZapis = true)
        {
            string currentDirectory = Directory.GetCurrentDirectory();

            try
            {
                Directory.SetCurrentDirectory(path);

                if (ajZapis)
                {
                    string tempFilePath = Path.Combine(path, "XXXXXX.XXXXXX");

                    using (var fs = File.Create(tempFilePath)) { }
                    File.Delete(tempFilePath);
                }

                return true;
            }
            catch (Exception ex) 
            {
                Declare.Errors.Add(ex.ToString());
                return false;
            }
            finally
            {
                Directory.SetCurrentDirectory(currentDirectory);
            }
        }

        /// <summary>
        /// vrati windows dir
        /// </summary>
        /// <returns></returns>
        public static string GetWindowsDir()
        {
            string windowsDir = Environment.GetFolderPath(Environment.SpecialFolder.Windows); //Environment.GetEnvironmentVariable("windir");

            if (!windowsDir.EndsWith("\\"))
                windowsDir += "\\";

            return windowsDir;
        }

        /// <summary>
        /// vrati verziu alfa.exe ??? CRV2Kros.exe ???
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="MAJOR"></param>
        /// <param name="MINOR"></param>
        /// <param name="REVISION"></param>
        public static void DajVerziuExe(string fileName, out int MAJOR, out int MINOR, out int REVISION)
        {
            MAJOR = 0;
            MINOR = 0;
            REVISION = 0;

            var fvi = FileVersionInfo.GetVersionInfo(fileName);
            MAJOR = fvi.FileMajorPart;
            MINOR = fvi.FileMinorPart;
            REVISION = fvi.FilePrivatePart;
        }

        /// <summary>
        /// vrati verziu ako string
        /// </summary>
        /// <param name="MAJOR"></param>
        /// <param name="MINOR"></param>
        /// <param name="REVISION"></param>
        /// <param name="nepouzitRevision"></param>
        /// <returns></returns>
        public static string DajVerziuString(int MAJOR, int MINOR, int REVISION, bool nepouzitRevision = false)
        {
            string s = $"{MAJOR:0}.{MINOR:00}";
            if (!nepouzitRevision)
                s += $".{REVISION:00}";
            return s;
        }

        /// <summary>
        /// Vrati cestu o x adresarov vyssie aj s lomitkom. Ak je to hlupost tak vrati prazdny retazec
        /// </summary>
        /// <param name="cesta"></param>
        /// <param name="pocet"></param>
        /// <returns></returns>
        public static string DajCestuOXAdresarovVyssie(string cesta, int pocet)
        {
            if (string.IsNullOrEmpty(cesta) || pocet <= 0)
                return string.Empty;

            if (cesta.EndsWith("\\"))
                cesta = cesta.Substring(0, cesta.Length - 1);

            int j = 0;
            while (cesta.Length > 0)
            {
                if (cesta[cesta.Length - 1] == '\\')
                {
                    j++;
                    if (j == pocet)
                        break; // máme požadovaný počet úrovní
                }

                cesta = cesta.Substring(0, cesta.Length - 1);
            }

            return (j == pocet) ? cesta : string.Empty;
        }

        /// <summary>
        /// 0 - nie je nainstalovany
        ///-1 - je nainstalovany acrobat reader
        /// 1 - je nainstalovany acrobat (aj writer)
        /// </summary>
        /// <returns></returns>
        public static int IsAcrobatReaderInstalled()
        {
            if (CitajHodnotuReg(Registry.LocalMachine,
                                    @"Software\Microsoft\Windows\CurrentVersion\App Paths\ACRORD32.EXE",
                                    "",
                                    out object _))
                return -1;

            if (CitajHodnotuReg(Registry.LocalMachine,
                                    @"Software\Microsoft\Windows\CurrentVersion\App Paths\ACROBAT.EXE",
                                    "",
                                    out object _))
                return 1;

            return 0;
        }

        /// <summary>
        /// restartuje pocitac
        /// </summary>
        public static bool Restart()
        {
            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = "shutdown",
                    Arguments = "/r /t 0",
                    UseShellExecute = false,
                    CreateNoWindow = true
                });

                return true;
            }
            catch (Exception ex)
            {
                Declare.Errors.Add(ex.ToString());
                return false;
            }
        }

        public static long DiskSpaceKB(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
                return -1;

            string root = Path.GetPathRoot(Path.GetFullPath(path));
            if (string.IsNullOrEmpty(root))
                return -1;

            var di = new DriveInfo(root);

            return di.AvailableFreeSpace / 1024;
        }

        /// <summary>
        /// '-1 ak je prva starsia, 0 ak sa rovnaju a 1 ak je druha starsia
        /// </summary>
        /// <param name="major1"></param>
        /// <param name="minor1"></param>
        /// <param name="revision1"></param>
        /// <param name="major2"></param>
        /// <param name="minor2"></param>
        /// <param name="revision2"></param>
        /// <param name="porovnavatRevision"></param>
        /// <returns></returns>
        public static int JePrvaVerziaStarsia(
                            int major1, int minor1, int revision1,
                            int major2, int minor2, int revision2,
                            bool porovnavatRevision = true)
        {
            if (major1 == major2)
            {
                if (minor1 == minor2)
                {
                    if (porovnavatRevision)
                    {
                        if (revision1 == revision2) return 0;
                        return revision1 < revision2 ? -1 : 1;
                    }
                    return 0;
                }
                return minor1 < minor2 ? -1 : 1;
            }
            return major1 < major2 ? -1 : 1;
        }
    }
}

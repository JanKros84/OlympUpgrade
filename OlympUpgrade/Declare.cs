using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Security.AccessControl;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

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
    }
}

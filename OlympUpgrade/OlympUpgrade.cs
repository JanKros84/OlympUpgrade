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
        //Dictionary<string, object> _PomCol = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
        private List<string> _PomCol = new List<string>();

        private Encoding _zipEncoding = Encoding.GetEncoding(852); //TODO uistit sa 1250/852/850

        public OlympUpgrade()
        {
            InitializeComponent();
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
                MessageBox.Show(ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void OlympUpgrade_Activated(object sender, EventArgs e)
        {
            try
            {
                DialogResult resVysledok = DialogResult.Yes;

                // Skontroluj, kto zavolal funkciu (inštalácia)
                if (Declare.KTO_VOLAL == Declare.VOLAL_INSTALL)
                {
                    // Zobraziť okno v popredí
                    //SetWindowPos(this.Handle, (IntPtr)HWND_TOP, 0, 0, 0, 0, SWP_SHOWWINDOW);
                    this.TopMost = true;
                    Application.DoEvents();
                    this.BringToFront();
                    Application.DoEvents();
                    this.TopMost = false;


                    btnOk.Visible = false;
                    btnStorno.Enabled = false;

                    while (resVysledok == DialogResult.Yes)
                    {
                        if (JeSpustenyProgram(Declare.DEST_PATH, Declare.SUBOR_EXE))
                        {
                            resVysledok = MessageBox.Show(
                                $"Nie je možné inštalovať UPGRADE, pretože program OLYMP v adresári {Declare.DEST_PATH} je práve spustený." +
                                "\n\nMusíte najskôr ukončiť spustený program OLYMP." +
                                "\nSkúsiť znovu?",
                                "Inštalácia",
                                MessageBoxButtons.YesNo,
                                MessageBoxIcon.Exclamation);

                            if (resVysledok == DialogResult.No)
                            {
                                this.Close(); //UkonciProgram();
                            }
                        }
                        else
                        {
                            Instaluj();
                            resVysledok = DialogResult.No;
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
                MessageBox.Show(ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// samotna instalacia
        /// </summary>
        private void Instaluj()
        {
            btnOk.Visible = false;
            btnStorno.Enabled = false;
            OverLicenciu();
            OverVolneMiesto();
            // OverArchiv         'toto s novym komponentom nevieme
            // VymazSuboryStarejAlfy
            KopirujSubory();
            //InstalujHotFixPreMapi();
            //UlozDoRegistrovVerziu();
            //Thread.Sleep(500); //iba kvoli tomu, aby to neprefrcalo tak rychlo
            //ZobrazVysledok();
        }

        /// <summary>
        /// Prida subory, ktore nebolo mozne nakopirovat kvoli chybe
        /// </summary>
        public void KopirujSubory()
        {
            int i, zipRes, zipRes2, BolaChyba, pocet;
            string pom, pom1, povodnaVerziaX;
            long povodnaVerzia, chyba;
            bool prepis;
            var neprepisovatDBS = new List<string>(); ;
            //Dim chilkatEntry As New ChilkatZipEntry2

            ProgresiaPreset(0);

            lblDestFile.Text = "Kopírovanie súborov ...";
            lblDestFile.Refresh();
            
            using (ZipArchive zip = ZipFile.Open(Path.Combine(Declare.AKT_ADRESAR, Declare.SUBOR_ZIP), ZipArchiveMode.Read, _zipEncoding))
            {
                ProgresiaPreset(zip.Entries.Count);
                var tempAdr = Declare.DajTemp(Declare.AKT_ADRESAR);
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
                                        entry.ExtractToFile(destPath, overwrite: true);
                                    }
                                    catch
                                    {
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
                    long.TryParse(povodnaVerziaX.Substring(0, povodnaVerziaX.IndexOf(".")), out parseVal);
                    povodnaVerzia = parseVal * 100;

                    parseVal = 0;
                    povodnaVerziaX = povodnaVerziaX.Substring(povodnaVerziaX.IndexOf(".") + 1);
                    long.TryParse(povodnaVerziaX.Substring(0, povodnaVerziaX.IndexOf(".")), out parseVal);
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
                catch { }

                // --- rozbal systémové reporty podľa listov ---
                RozbalZostavy(zip, "Reporty\\", "RPT", Declare.FILE_REPORTY_TXT);
                RozbalZostavy(zip, "Reporty\\", "REPX", Declare.FILE_REPORTY_TXT);
                RozbalZostavy(zip, "Reporty\\Excel\\", "XLS", Declare.FILE_REPORTY_EXCEL_TXT);
                RozbalZostavy(zip, "Reporty\\Excel\\", "PDF", Declare.FILE_REPORTY_EXCEL_P_TXT);
                RozbalZostavy(zip, "Reporty\\Pdf\\", "PDF", Declare.FILE_REPORTY_PDF_TXT);


                // --- GRAFIKA\*.*  ---
                _PomCol.Clear();
                ExtractFilesExtensionOnPath(zip, Declare.DEST_PATH, "Grafika", "*");
                CheckCopyException(Path.Combine(Declare.DEST_PATH, "Grafika"));

                // --- SKRIPTY\CREATE\*.sql ---
                _PomCol.Clear();
                ExtractFilesExtensionOnPath(zip, Declare.DEST_PATH, "Skripty\\create\\", "sql");
                CheckCopyException(Path.Combine(Declare.DEST_PATH, "Skripty", "create"));

                // --- SKRIPTY\DROP\*.sql ---
                _PomCol.Clear();
                ExtractFilesExtensionOnPath(zip, Declare.DEST_PATH, "Skripty\\drop\\", "sql");
                CheckCopyException(Path.Combine(Declare.DEST_PATH, "Skripty", "drop"));

                // --- ZDROJE\*.* ---
                _PomCol.Clear();
                ExtractFilesExtensionOnPath(zip, Declare.DEST_PATH, "Zdroje", "*");
                CheckCopyException(Path.Combine(Declare.DEST_PATH, "Zdroje"));

                /* ---TEMPLATE---
                PripravPomCol();
                zipRes = zip.UnzipMatching(target, "ZDROJE\\*.*", 0);
                if (zipRes < 1 || PomCol.Count > 0)
                {
                    PridajChybu();
                    MessageBox.Show(
                        $"Nastala chyba pri kopírovaní súborov do adresára {target}ZDROJE\\\r\n\r\n" +
                        "Súbory, ktoré sa nepodarilo nainštalovať, budú uvedené na konci inštalácie aj s typom chyby.",
                        Title, MessageBoxButtons.OK, MessageBoxIcon.Error);
                }*/

                // --- Vzory ---
                var vzoryPath = Path.Combine(Declare.DEST_PATH, "Vzory");
                if (Directory.Exists(vzoryPath))
                    NastavPrava(vzoryPath, "*.*", FileAttributes.Normal);

                _PomCol.Clear();
                ExtractFilesExtensionOnPath(zip, Declare.DEST_PATH, "Zdroje", "*");
                CheckCopyException(vzoryPath);
                if (_PomCol.Count > 0)
                    NastavPrava(Path.Combine(vzoryPath, "Vzory\\"), "*.*", FileAttributes.ReadOnly);

                CopyFileIfExists(Declare.AKT_ADRESAR, Declare.DEST_PATH, "OLYMP.CHM");
                CopyFileIfExists(Declare.AKT_ADRESAR, Declare.DEST_PATH, "CRV2Kros.exe");
                CopyFileIfExists(Declare.AKT_ADRESAR, Declare.DEST_PATH, "ADOWrapper.dll");
                CopyFileIfExists(Declare.AKT_ADRESAR, Declare.DEST_PATH, "itextsharp.dll");
                CopyFileIfExists(Declare.AKT_ADRESAR, Declare.DEST_PATH, "Kros.BankTransfers.dll");

                VymazStareSubory();

                // --- SK\*.* ---
                _PomCol.Clear();
                ExtractFilesExtensionOnPath(zip, Declare.DEST_PATH, "sk", "*");
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
                if (Declare.ExistujeSubor(zoznamPath))
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
                /*zipRes = zip.UnzipMatching(target, "DevExpress.*.dll", 0);
                if (zipRes < 1 || PomCol.Count > 0)
                {
                    PridajChybu();
                    MessageBox.Show(
                        $"Nastala chyba pri kopírovaní súborov do adresára {target}\r\n\r\n" +
                        "Súbory, ktoré sa nepodarilo nainštalovať, budú uvedené na konci inštalácie aj s typom chyby.",
                        Title, MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

                // --- Close ZIP ---
                zip.CloseZip();

                // --- Presun ZIP/installer do UPGRADE ---
                try
                {
                    var upgDir = Path.Combine(target, "UPGRADE\\");
                    EnsureDir(upgDir);

                    // presuň ZIP do UPGRADE
                    var dstZip = Path.Combine(upgDir, SUBOR_ZIP);
                    if (FileExists(dstZip)) File.Delete(dstZip);

                    NastavCestu(Path.Combine(target, "UPGRADE\\" + SUBOR_ZIP));
                    FileCopyOverwrite(zipPath, dstZip);
                    Spravne.Add("UPGRADE\\" + SUBOR_ZIP);

                    // pokus odstrániť pôvodný ZIP v targete (VB6: Kill UCase$(DEST_PATH & SUBOR_ZIP))
                    var targetZip = Path.Combine(target, SUBOR_ZIP);
                    if (FileExists(targetZip)) File.Delete(targetZip);
                }
                catch { }

                try
                {
                    var upgDir = Path.Combine(target, "UPGRADE\\");
                    EnsureDir(upgDir);

                    // presun pôvodných inštalátorov
                    var instOld = Path.Combine(target, SUBOR_INST);
                    if (FileExists(instOld))
                    {
                        var dst = Path.Combine(upgDir, SUBOR_INST);
                        if (FileExists(dst)) File.Delete(dst);
                        FileCopyOverwrite(instOld, dst);
                        File.Delete(instOld);
                    }

                    var instNew = Path.Combine(target, SUBOR_INST_NEW);
                    if (FileExists(instNew))
                    {
                        var dst = Path.Combine(upgDir, SUBOR_INST_NEW);
                        if (FileExists(dst)) File.Delete(dst);

                        // VB6: Name ... As ... (move+rename)
                        File.Move(instNew, dst);

                        // počkaj krátko, kým sa systém „spamätá“
                        var tries = 0;
                        while (FileExists(instNew) && tries++ < 10)
                        {
                            Thread.Sleep(50);
                        }
                    }
                }
                catch { }
                */
                ProgresiaDone();

                // VB6 volalo ZapisVysledokDoListu – tu to nechám na vás
                // ZapisVysledokDoListu();
            }
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
                    try { File.SetAttributes(file, FileAttributes.Normal); } catch { }
                    try { File.Delete(file); } catch { }
                }
            }
            catch { }
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
                catch (Exception)
                {
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
            NastavDlheNazvy(Cesta, "." + pripona, 50, subor);
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

        private string ExtractFilesExtensionOnPath(ZipArchive zip, string destFolder, string searchPath, string extension)
        {
            extension = (extension ?? "").ToLowerInvariant();
            searchPath = searchPath.Last() == '\\' ? searchPath.Substring(0, searchPath.Length - 1) : searchPath;
            searchPath = searchPath.Last() == '/' ? searchPath.Substring(0, searchPath.Length - 1) : searchPath;
            searchPath = searchPath.Replace('\\', '/');
            var destPath = Path.Combine(destFolder, searchPath);
            Directory.CreateDirectory(destPath);

            // 1) Rozbaliť všetky *.pripona v podsložke Cesta
            //int zipRes = UnzipMatching(Declare.DEST_PATH, mask1, PomCol);
            string regexMask = $@"(?i)^{Regex.Escape(searchPath)}[\\/][^\\/]+\.{extension}$";

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
                entry.ExtractToFile(Path.Combine(destPath, entry.Name), overwrite: true);
                lblDestFile.Text = $"{entry.FullName}";
                lblDestFile.Refresh();

                _Spravne.Add(entry.FullName);
            }
            catch (Exception ex)
            {
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

                // VB6: SetAttr ... vbNormal (clear RO/hidden)
                try { File.SetAttributes(filePath, FileAttributes.Normal); } catch { }

                // VB6: Mid$(s,4,1) = "-" And Val(Mid$(s,5,2)) < Hranica
                if (fileName.Length >= 6 && fileName[3] == '-')
                {
                    if (int.TryParse(fileName.Substring(4, 2), out var num) && num < hranica)
                    {
                        var key = fileName.Substring(0, 6); // e.g., "ABC-12"
                        if (map.TryGetValue(key, out var longName))
                        {
                            var newNameCore = key + "-" + longName;
                            //newNameCore = RemoveInvalidFileNameChars(newNameCore);

                            var newPath = Path.Combine(destDir, newNameCore + pripona);

                            // Emulate overwrite: delete target if it exists
                            try
                            {
                                //File.SetAttributes(newPath, FileAttributes.Normal);
                                if (File.Exists(newPath))
                                    File.Replace(filePath, newPath, null);
                                else
                                    File.Move(filePath, newPath);
                            }
                            catch { }
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

                NastavPrava(dir, maska, FileAttributes.Normal);
                try
                {
                    foreach (var path in Directory.GetFiles(dir, maska))
                    {
                        try { File.Delete(path); } catch { }
                    }
                }
                catch { /* adresár neexistuje a pod. – pokračuj */ }
            }
        }

        // CestaVym: adresár; sablona: maska s * a ?; prava: FileAttributes.Normal atď.
        void NastavPrava(string CestaVym, string sablona, FileAttributes prava)
        {
            if (string.IsNullOrWhiteSpace(CestaVym) || !Directory.Exists(CestaVym))
                return;

            try
            {
                foreach (var path in Directory.GetFiles(CestaVym, sablona))
                {
                    try { File.SetAttributes(path, prava); } catch { }
                }
            }
            catch { }
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
            VolneMiesto = Declare.DiskSpaceKB(Declare.DEST_PATH);
            PotrebnaVelkost = DajVelkostSuborovVZip(Declare.PridajLomitko(Declare.AKT_ADRESAR) + Declare.SUBOR_ZIP) / 1024d;


            //TODO prec, kontroluje "Mzdy.lic"
            if (Declare.DEST_PATH != Declare.AKT_ADRESAR
                && Declare.ExistujeSubor(Declare.PridajLomitko(Declare.AKT_ADRESAR) + Declare.LICENCIA))
            {
                // PotrebnaVelkost = PotrebnaVelkost + FileLen(Declare.PridajLomitko(Declare.AKT_ADRESAR) & Declare.LICENCIA) / 1024;
            }

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
                    prompt += Declare.PridajLomitko(Declare.AKT_ADRESAR) + Declare.MENO_EXE;
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
        /// <param name="path"></param>
        /// <returns></returns>
        public long DajVelkostSuborovVZip(string path)
        {
            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
                return 0;

            long sum = 0;

            using (ZipArchive zip = ZipFile.Open(path, ZipArchiveMode.Read, _zipEncoding))
            {
                ProgresiaPreset(zip.Entries.Count);

                foreach (var e in zip.Entries)
                {
                    try
                    {
                        // Rátame iba súbory (nie adresáre)
                        bool isDirectory = e.Name.Length == 0 || e.FullName.EndsWith("/", StringComparison.Ordinal);
                        if (!isDirectory)
                            sum += e.Length; // uncompressed size
                    }
                    catch { }

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
            bool LicVSysteme = false;
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

                //TODO Tato cast nikdy nepojde -> Mzdy.lic
                // 2) Pozri licenciu na inštalačkách (ak je v akt. adresári a nie je to istý adresár ako cieľ)
                string aktLicPath = Declare.PridajLomitko(Declare.AKT_ADRESAR) + Declare.LICENCIA;
                if (Declare.ExistujeSubor(aktLicPath) &&
                    !string.Equals(Declare.PridajLomitko(Declare.AKT_ADRESAR), Declare.DEST_PATH, StringComparison.OrdinalIgnoreCase))
                {
                    Lic = Declare.PridajLomitko(Declare.AKT_ADRESAR); // budem kopírovať z tohto koreňa
                }

                //TODO Tato cast nikdy nepojde -> Mzdy.lic
                // 3) Ak je licencia na inštalačkách, ide sa inštalovať „ostrá“
                if (!string.IsNullOrEmpty(Lic))
                {
                    CitajRegistracnySubor(Declare.PCD_INSTALACKY, Lic);

                    if (Declare.VerziaPC != Declare.LIC_OSTRA)
                    {
                        // Nemá v PC ostrú verziu alebo na inštalačkách je novší licenčný súbor – kopíruj
                        KopirujLICENCIU = true;
                    }
                    else
                    {
                        // V cieli je ostrá verzia
                        if (Declare.ExistujeSubor(Path.Combine(Declare.DEST_PATH, Declare.LICENCIA_SW)))
                        {
                            // ak je v cieli nová licencia (SW), nekopíruj
                            KopirujLICENCIU = false;
                        }
                        else
                        {
                            // Porovnaj licencie PC vs. Disketa – ak iné ICO/PorCislo a nie Install, pýtaj sa
                            if ((Declare.ICO_Disketa != Declare.ICO_PC || Declare.PorCisloDisketa != Declare.PorCisloPC) &&
                                !LicVSysteme &&
                                Declare.KTO_VOLAL != Declare.VOLAL_INSTALL)
                            {
                                var res = MessageBox.Show(
                                    "V cieľovom adresári sa nachádza ostrá verzia registrovaná na firmu " + Declare.NazovFirmyPC + "." +
                                    Environment.NewLine + Environment.NewLine +
                                    "Inštalácia obsahuje iný registračný súbor registrovaný na firmu " + Declare.NazovFirmyDisketa + "." +
                                    Environment.NewLine + Environment.NewLine +
                                    "Naozaj chcete inštaláciu spustiť?",
                                    "OLYMP – inštalácia",
                                    MessageBoxButtons.YesNo,
                                    MessageBoxIcon.Question,
                                    MessageBoxDefaultButton.Button2);

                                if (res == DialogResult.No)
                                {
                                    Declare.ExitProg(0);
                                    return;
                                }
                                KopirujLICENCIU = true;
                            }

                            string destLic = Path.Combine(Declare.DEST_PATH, Declare.LICENCIA);
                            if (Declare.ExistujeSubor(destLic))
                            {
                                // Má licenciu v počítači – kopíruj ak je staršia alebo rovnaká, alebo ak to prikázal InstallShield
                                bool pcLeOrEqDisk =
                                    IntYyyymmddToDate(Declare.DatumPC) <= IntYyyymmddToDate(Declare.DatumDisketa);

                                if (pcLeOrEqDisk || Declare.KOPIRUJ_LICENCIU == 1)
                                {
                                    KopirujLICENCIU = true;
                                }
                                else if (Declare.KOPIRUJ_LICENCIU == 2)
                                {
                                    KopirujLICENCIU = false;
                                }
                                else
                                {
                                    // Na inštalačkách je starší lic. súbor – opýtať sa na prepis
                                    var res2 = MessageBox.Show(
                                        "Inštalačný program sa chystá skopírovať súbor " + Declare.LICENCIA + " z inštalačného média do počítača." +
                                        Environment.NewLine + Environment.NewLine +
                                        "Súbor " + Declare.LICENCIA + " na inštalačnom médiu je ale starší ako ten, ktorý máte v počítači." + Environment.NewLine +
                                        "Prepísať existujúci súbor " + Declare.LICENCIA + " (" + FormatDdMmYyyy(Declare.DatumPC) + ") " +
                                        "starším súborom " + Declare.LICENCIA + " (" + FormatDdMmYyyy(Declare.DatumDisketa) + ") z inštalačného média?",
                                        "OLYMP – inštalácia",
                                        MessageBoxButtons.YesNo,
                                        MessageBoxIcon.Question,
                                        MessageBoxDefaultButton.Button2);

                                    KopirujLICENCIU = (res2 == DialogResult.Yes);
                                }
                            }
                            else
                            {
                                // Licencia bola len v systéme → kopíruj „naisto“
                                KopirujLICENCIU = true;
                            }
                        }
                    }

                    // 4) Kopírovanie licencie (ak treba)
                    if (KopirujLICENCIU)
                    {
                        string destLic = Path.Combine(Declare.DEST_PATH, Declare.LICENCIA);
                        string srcLic = Path.Combine(Lic, Declare.LICENCIA);

                        // VB6: On Error Resume Next - zruš RO atribút, vymaž pôvodný, potom kopíruj a nastav RO
                        try
                        {
                            if (File.Exists(destLic))
                            {
                                File.SetAttributes(destLic, FileAttributes.Normal);
                                File.Delete(destLic);
                            }
                        }
                        catch { /* ignoruj ako VB6 */ }

                        // skopíruj novú
                        File.Copy(srcLic, destLic, overwrite: false);

                        try
                        {
                            File.SetAttributes(destLic, FileAttributes.ReadOnly);
                        }
                        catch { /* ignore */ }

                        MaOstruVerziu = true;
                        Declare.DatumPC = Declare.DatumDisketa; // už je tam nový lic. súbor – test bude nakoniec
                    }
                }

                ProgresiaTik();

                // 5) Nemá licenciu na inštalačkách a chce „ostrú“ – skontroluj nárok
                if (MaOstruVerziu && !KopirujLICENCIU)
                {
                    Declare.DatumDisketa = (Declare.VNUTORNY_DATUM_ROK * 10000) + (Declare.VNUTORNY_DATUM_MESIAC * 100) + Declare.VNUTORNY_DATUM_DEN;

                    if (IntYyyymmddToDate(Declare.DatumPC) < IntYyyymmddToDate(Declare.DatumDisketa))
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
            catch
            {
                // chyba:
                Declare.ExitProg(Declare.ID_CHYBA_LICENCIA);
            }
        }

        static DateTime IntYyyymmddToDate(long ymd) =>
        new DateTime((int)ymd / 10000, (int)(ymd / 100) % 100, (int)ymd % 100);

        static string FormatDdMmYyyy(long ymd) =>
            IntYyyymmddToDate(ymd).ToString("dd.MM.yyyy");

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
        /// Zistuje, ci je olymp, ktora sa upgraduje, spustena (WMI hlada 32 aj 64)
        /// </summary>
        /// <param name="destPath"></param>
        /// <param name="suborExe"></param>
        /// <returns></returns>
        private bool JeSpustenyProgram(string destPath, string suborExe)
        {
            //string exeName = Path.GetFileName(fullExePath);
            var fullExePath = Path.Combine(destPath, suborExe);
            string query = $"SELECT ProcessId, ExecutablePath FROM Win32_Process WHERE Name = '{suborExe.Replace("'", "''")}'";

            using (var searcher = new ManagementObjectSearcher(query))
            using (var results = searcher.Get())
            {
                foreach (ManagementObject mo in results)
                {
                    var path = mo["ExecutablePath"] as string;
                    if (!string.IsNullOrEmpty(path) &&
                        string.Equals(path, fullExePath, StringComparison.OrdinalIgnoreCase))
                        return true;
                }
            }
            return false;
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
                    OdstranZRegistrovDoinstalovanie();
                    Declare.DEST_PATH = Declare.PridajLomitko(pom);

                    tabControl1.SelectTab(1);

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
                    NastavCestuVerzie(Declare.PridajLomitko(pom));
                }
                else
                {
                    // spustil uzivatel
                    Declare.TTITLE = "OlympUpgrade";
                    Declare.KTO_VOLAL = Declare.VOLAL_UZIVATEL;
                    NastavCestuVerzie();
                }

                this.Text = "OLYMP " + lblUpgrade.Text + " - Sprievodca inštaláciou";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Nastavi cestu, kde je nainstalovany olymp, zisti verzie upgradu a nainstalovaneho execka
        /// </summary>
        /// <param name="cesta"></param>
        private void NastavCestuVerzie(string cesta = "")
        {
            // zistím potrebné verzie
            lblUpgrade.Text = DajVerziuUpgrade();

            if (!Declare.MamDostatocnyFramework())
            {
                Declare.ExitProg(Declare.ID_CHYBA_CHYBA_FRAMEWORK);
                return;
            }

            // zistím, či je nainštalovaný OLYMP -> čítam InstallLocation z HKLM\...\Uninstall\{ProductCode}
            string path = null;
            bool nacital = false;

            foreach (var productCode in new[] { Declare.PRODUCT_CODE_1, Declare.PRODUCT_CODE_2 })
            {
                if (Declare.CitajHodnotuReg(Registry.LocalMachine,
                                    $@"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{productCode}",
                                    "InstallLocation",
                                    out object val))
                {
                    if (val is string licPath
                        && !string.IsNullOrWhiteSpace(licPath))
                    {
                        path = licPath;
                        nacital = true;
                    }
                    break;
                }
            }

            //ak nie je, je to chyba koncim
            if (!nacital)
            {
                Declare.ExitProg(Declare.ID_CHYBA_NIE_JE_MIN_VERZIA);
                return;
            }
            else
            {
                // uložím pôvodný adresár inštalácie z registrov
                Declare.DEST_PATH_INSTALLSHIELD = Declare.PridajLomitko(path);

                // načítam info z licencie
                CitajRegistracnySubor(Declare.PCD_INSTALACKY, Declare.PridajLomitko(Declare.AKT_ADRESAR));

                // zistím, či daná licencia už bola na PC – ak áno, nasmerujem inštaláciu na jej adresár
                //string licPath;
                var postKey = Declare.VerziaDisketa == Declare.LIC_OSTRA
                    ? "Vyplaty" + kodovanie.VratIDRegistracky(Declare.ICO_Disketa, Declare.PorCisloDisketa)
                    : string.Empty;

                if (Declare.CitajHodnotuReg(Registry.LocalMachine,
                                   $@"Software\Kros\Olymp\{postKey}",
                                   "InstalacnyAdresar",
                                   out object val))
                {
                    if (val is string licPath
                        && !string.IsNullOrWhiteSpace(licPath))
                        path = licPath;
                }
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
                Declare.DEST_PATH = Declare.PridajLomitko(path);
            }

            // skontrolujem, či je možné do adresára inštalovať
            if (!prazdnyString && !Declare.JeAdresarSpravny(Declare.DEST_PATH))
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
                Declare.DajVerziuExe(pathExe, out Declare.P_MAJOR, out Declare.P_MINOR, out Declare.P_REVISION);
                if (Declare.P_MAJOR == 0)
                    return Declare.VERZIA_NEZNAMA;
                else
                    return Declare.DajVerziuString(Declare.P_MAJOR, Declare.P_MINOR, Declare.P_REVISION);
            }
            else if (File.Exists(pathExeStary))
            {
                Declare.DajVerziuExe(pathExeStary, out Declare.P_MAJOR, out Declare.P_MINOR, out Declare.P_REVISION);
                if (Declare.P_MAJOR == 0)
                    return Declare.VERZIA_NEZNAMA;
                else
                    return Declare.DajVerziuString(Declare.P_MAJOR, Declare.P_MINOR, Declare.P_REVISION);
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
            if (Declare.ExistujeSubor(Path.Combine(cesta, Declare.LICENCIA_SW)))
            {
                NacitajNovuLicenciu(PcD, Path.Combine(cesta, Declare.LICENCIA_SW));
                return;
            }


            //var starySubor = Path.Combine(cesta, Declare.LICENCIA);
            //if (!Declare.ExistujeSubor(starySubor))
            {
                //NEBOL NAJDENY REGISTRACNY SUBOR
                if (PcD == Declare.PCD_POCITAC) Declare.VerziaPC = 0; else Declare.VerziaDisketa = 0;
                return;
            }
            /*
            string kodovany;
            try
            {
                // VB6 pracoval v ANSI; pre SK/CZ prostredie je typicky Windows-1250
                // (ak máš inú kódovú stránku, uprav číslo)
                var bytes = File.ReadAllBytes(starySubor);
                kodovany = Encoding.GetEncoding(1250).GetString(bytes);
            }
            catch
            {
                if (PcD == Declare.PCD_POCITAC)
                    Declare.VerziaPC = Declare.LIC_CHYBNA;
                else
                    Declare.VerziaDisketa = Declare.LIC_CHYBNA;
                return;
            }

            // prázdny/chybný súbor
            if (kodovany.Length < 13)
            {
                if (PcD == Declare.PCD_POCITAC)
                    Declare.VerziaPC = Declare.LIC_CHYBNA;
                else
                    Declare.VerziaDisketa = Declare.LIC_CHYBNA;
                return;
            }

            // text = prvých 8 znakov (CRC)
            var crcText = SafeSubstring(kodovany, 1, 8);

            // odstráň hlavičku/koniec podľa VB: Mid(kodovany, 11, Len-13)
            var payload = SafeSubstring(kodovany, 11, kodovany.Length - 13);

            // over CRC
            if (!string.Equals(VypocitajCrc(payload, payload.Length, 0), crcText, StringComparison.Ordinal))
            {
                if (PcD == Declare.PCD_POCITAC)
                    Declare.VerziaPC = Declare.LIC_CHYBNA;
                else
                    Declare.VerziaDisketa = Declare.LIC_CHYBNA;
                return;
            }

            // odhesluj -> 'text'
            var text = Odhesluj(payload);

            // VB indexovanie: i začína na j + 2, pričom j je -1 => i = 1 (1-based).
            // V C# si to urobíme 0-based a budeme simulovať ďalšie InStr volania.
            int i = 0;
            int j = IndexOfCrLf(text, i);

            // 1) názov firmy
            var nazovFirmy = Sub(text, i, j);
            if (PcD == Declare.PCD_INSTALACKY)
                Declare.NazovFirmyDisketa = nazovFirmy;
            else
                Declare.NazovFirmyPC = nazovFirmy;

            // preskoč 3 riadky (partner, nastavenia…)
            for (int X = 2; X <= 4; X++)
            {
                i = j + 2; j = IndexOfCrLf(text, i);
            }

            // 2) ICO
            i = j + 2; j = IndexOfCrLf(text, i);
            var icoStr = Sub(text, i, j);
            if (PcD == Declare.PCD_INSTALACKY)
                Declare.ICO_Disketa = ValLong(icoStr);
            else
                Declare.ICO_PC = ValLong(icoStr);

            // 3) pozícia 6 preskočím (typ inštalácie)
            i = j + 2; j = IndexOfCrLf(text, i);
            int typInstalacie = 0;
            if (i < j) int.TryParse(Sub(text, i, j), out typInstalacie);

            int SAStypLicencie = 0;

            if (PcD == Declare.PCD_INSTALACKY)
            {
                // ostrá / update
                Declare.VerziaDisketa = (typInstalacie == 1 || typInstalacie == 2 || typInstalacie == 4) ? Declare.LIC_OSTRA : Declare.LIC_UPDATE;

                // preskoč do dátumu registrácie (7..10)
                for (int X = 7; X <= 10; X++) { i = j + 2; j = IndexOfCrLf(text, i); }
                if (j >= 0)
                {
                    i = j + 2; j = IndexOfCrLf(text, i);
                    var dStr = (i < j) ? Sub(text, i, j) : Sub(text, i);
                    Declare.DatumDisketa = ValLong(dStr);
                }

                // Poradové číslo (12..14)
                for (int X = 12; X <= 14; X++) { i = j + 2; j = IndexOfCrLf(text, i); }
                Declare.PorCisloDisketa = ValLong(Sub(text, i, j));

                // SAS typ licencie (16..22)
                for (int X = 16; X <= 22; X++) { i = j + 2; j = IndexOfCrLf(text, i); }
                SAStypLicencie = (j - i >= 0) ? (int)ValLong(Sub(text, i, j)) : 0;

                // ak nie je SAS, posuň dátum o rok (+10000 vo formáte yyyymmdd)
                if (SAStypLicencie == 0 && Declare.DatumDisketa != 0) Declare.DatumDisketa += 10000;
            }
            else // PC licencia
            {
                // preskoč (7..10)
                for (int X = 7; X <= 10; X++) { i = j + 2; j = IndexOfCrLf(text, i); }

                // cesta k dátam
                string pom = (j >= 0) ? Sub(text, i, j) : Sub(text, i);
                if (Declare.ExistujeSubor(pom)) cestaInstalacie = pom;

                // dátum registrácie (pozor na „bug“ – prázdny riadok)
                if (j >= 0)
                {
                    i = j + 2; j = IndexOfCrLf(text, i);
                    if (i == j) { i = j + 2; j = IndexOfCrLf(text, i); } // fix VB bugu
                    var dStr = (i < j) ? Sub(text, i, j) : Sub(text, i);
                    Declare.DatumPC = ValLong(dStr);
                }

                // Poradové číslo (12..14)
                for (int X = 12; X <= 14; X++) { i = j + 2; j = IndexOfCrLf(text, i); }
                Declare.PorCisloPC = ValLong(Sub(text, i, j));

                // SAS typ licencie (16..22)
                for (int X = 16; X <= 22; X++) { i = j + 2; j = IndexOfCrLf(text, i); }
                SAStypLicencie = (j - i >= 0) ? (int)ValLong(Sub(text, i, j)) : 0;

                if (SAStypLicencie == 0 && DatumPC != 0) DatumPC += 10000;

                // dobrý reg
                Declare.VerziaPC = Declare.LIC_OSTRA;
            }*/
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
                        string y = SafeSubstring(riadok, colon + 4, 4);
                        string m = SafeSubstring(riadok, colon + 9, 2);
                        string d = SafeSubstring(riadok, colon + 12, 2);
                        long.TryParse(y + m + d, out datumLicencie);
                    }
                    else if (riadokUpper.Contains("PRENAJOMSOFTVERU"))
                    {
                        //Val(Mid(riadok, InStr(1, riadok, ":") + 1, 2)) <> 0
                        string v = SafeSubstring(riadok, colon + 2, 1);
                        prenajom = (ValLong(v) != 0);
                    }
                    else if (riadokUpper.Contains("PARTNERICO"))
                    {
                        //If PcD = PCD_INSTALACKY Then
                        //    ICO_Disketa = Replace(Replace(Mid(riadok, InStr(1, riadok, ":") + 3, 20), """", ""), ",", "")
                        //Else
                        //    ICO_PC = Replace(Replace(Mid(riadok, InStr(1, riadok, ":") + 3, 20), """", ""), ",", "")
                        string s = SafeSubstring(riadok, colon + 3, 20).Replace("\"", "").Replace(",", "").Replace("\\", "");
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
                        string s = SafeSubstring(riadok, colon + 4, len);
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
        private string SafeSubstring(string s, int startZeroBased, int length)
        {
            if (string.IsNullOrEmpty(s) || length <= 0) return string.Empty;
            if (startZeroBased < 0) startZeroBased = 0;
            if (startZeroBased >= s.Length) return string.Empty;
            int len = Math.Min(length, s.Length - startZeroBased);
            return s.Substring(startZeroBased, len);
        }

        private long ValLong(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return 0;
            long.TryParse(s.Trim(), out var v);
            return v;
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
                    Path.Combine(Declare.AKT_ADRESAR ?? string.Empty, Declare.MENO_EXE ?? string.Empty) + " znova. " +
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

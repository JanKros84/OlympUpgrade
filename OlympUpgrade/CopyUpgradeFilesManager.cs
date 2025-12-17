using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;

namespace OlympUpgrade
{
    internal class CopyUpgradeFilesManager
    {
        private List<string> _Chybne = new List<string>();
        private List<string> _Spravne = new List<string>();
        private List<string> _Preskocene = new List<string>();
        private List<string> _PomCol = new List<string>();

        public List<string> Chybne { get => _Chybne;}
        public List<string> Spravne { get => _Spravne; }
        public List<string> Preskocene { get => _Preskocene; }

        private Action<long> _progresiaPreset;
        private Action _progresiaTik, _progresiaDone;
        private Action<string> _setLabelInfo;

        public CopyUpgradeFilesManager(Action<long> progresiaPreset, Action progresiaTik, Action progresiaDone, Action<string> setLabelInfo)
        {
            _progresiaPreset = progresiaPreset;
            _progresiaTik = progresiaTik;
            _setLabelInfo = setLabelInfo;
            _progresiaDone = progresiaDone;
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
                _progresiaPreset?.Invoke(0);

                _setLabelInfo?.Invoke("Kopírovanie súborov ...");

                using (ZipArchive zip = ZipFile.Open(Path.Combine(Declare.AKT_ADRESAR, Declare.SUBOR_ZIP), ZipArchiveMode.Read, Declare.ZipEncoding))
                {
                    _progresiaPreset(zip.Entries.Count);
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
                            _progresiaTik?.Invoke();
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
                        povodnaVerziaX = HelpFunctions.DajVerziuProgramu();
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

                    _setLabelInfo?.Invoke(upgradeZipPath);

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

                    _setLabelInfo?.Invoke(upgradeINST_Path);

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

                    _setLabelInfo?.Invoke(upgradeINST_NEW_Path);

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

                _progresiaDone();
                //ZapisVysledokDoListu();
            }
            catch (Exception ex)
            {
                Declare.Errors.Add(ex.ToString());
                MessageBox.Show($"Nastala chyba pri rozzipovani Erl : {ex.ToString()}\r\n\r\n {VratZoznam()}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
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

        private void ExtractToFile(string destPath, ZipArchiveEntry entry)
        {
            try
            {
                Directory.CreateDirectory(destPath);
                entry.ExtractToFile(Path.Combine(destPath, entry.Name), overwrite: true);

                _setLabelInfo?.Invoke($"{entry.FullName}");

                //Thread.Sleep(10);//iba kvoli tomu, aby to neprefrcalo tak rychlo

                _Spravne.Add(entry.FullName);
            }
            catch (Exception ex)
            {
                Declare.Errors.Add(ex.ToString());
                _PomCol.Add($"{entry.FullName} ex: {ex.Message}");
            }
            _progresiaTik?.Invoke();
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
                _setLabelInfo?.Invoke(destFilePath.ToUpperInvariant());

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

        private void RozbalZostavy(ZipArchive zip, string Cesta, string pripona, string subor)
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

        private void PridajChybu()
        {
            if (_PomCol != null && _PomCol.Count > 0)
            {
                foreach (var v in _PomCol)//.Values)
                {
                    _Chybne.Add("nedá sa prepísať\t" + v);
                }
            }
        }
    }
}

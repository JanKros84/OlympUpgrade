using Microsoft.Win32;
using System;
using System.Diagnostics;
using System.IO;
using System.Management;

namespace OlympUpgrade
{
    internal class HelpFunctions
    {
        /// <summary>
        /// 0 - nie je nainstalovany
        ///-1 - je nainstalovany acrobat reader
        /// 1 - je nainstalovany acrobat (aj writer)
        /// </summary>
        /// <returns></returns>
        public static int IsAcrobatReaderInstalled()
        {
            if (RegistryFunctions.CitajHodnotuReg(Registry.LocalMachine,
                                    @"Software\Microsoft\Windows\CurrentVersion\App Paths\ACRORD32.EXE",
                                    "",
                                    out object _))
                return -1;

            if (RegistryFunctions.CitajHodnotuReg(Registry.LocalMachine,
                                    @"Software\Microsoft\Windows\CurrentVersion\App Paths\ACROBAT.EXE",
                                    "",
                                    out object _))
                return 1;

            return 0;
        }

        /// <summary>
        /// Minimal 528040 -> .NET Framework 4.8
        /// </summary>
        /// <returns></returns>
        public static bool MamDostatocnyFramework()
        {
            if (RegistryFunctions.CitajHodnotuReg(Registry.LocalMachine,
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

        public static void InstalujHotFixPreMapi()
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

        public static DateTime IntYyyymmddToDate(long ymd) =>
            new DateTime((int)ymd / 10000, (int)(ymd / 100) % 100, (int)ymd % 100);

        public static string FormatDdMmYyyy(long ymd) =>
            IntYyyymmddToDate(ymd).ToString("dd.MM.yyyy");

        public static string SafeSubstring(string s, int startZeroBased, int length)
        {
            if (string.IsNullOrEmpty(s) || length <= 0) return string.Empty;
            if (startZeroBased < 0) startZeroBased = 0;
            if (startZeroBased >= s.Length) return string.Empty;
            int len = Math.Min(length, s.Length - startZeroBased);
            return s.Substring(startZeroBased, len);
        }

        public static long ValLong(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return 0;
            long.TryParse(s.Trim(), out var v);
            return v;
        }

        public static void TryToDeleteFile(string filePath)
        {
            try
            {
                if (!string.IsNullOrWhiteSpace(filePath)
                    && File.Exists(filePath))
                    File.Delete(filePath);
            }
            catch (Exception ex) { Declare.Errors.Add(ex.ToString()); }
        }

        /// <summary>
        /// Deletes all files in the specified directory that match the given search pattern.
        /// </summary>
        /// <remarks>File attributes are reset to normal before deletion. Errors encountered during file
        /// deletion or attribute modification are recorded in the Declare.Errors collection. Only files in the
        /// top-level of the specified directory are affected; subdirectories are not searched.</remarks>
        /// <param name="directory">The path to the directory in which to search for files to delete. If the directory does not exist, no action
        /// is taken.</param>
        /// <param name="searchPattern">The search string used to match the names of files to delete. Wildcard characters such as '*' and '?' are
        /// supported.</param>
        public static void DeleteMatching(string directory, string searchPattern)
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

        /// <param name="CestaVym">adresár</param>
        /// <param name="sablona">maska s * a ?</param>
        /// <param name="prava">FileAttributes.Normal atď.</param>
        public static void NastavPrava(string CestaVym, string sablona, FileAttributes prava)
        {
            if (string.IsNullOrWhiteSpace(CestaVym) || !Directory.Exists(CestaVym))
                return;

            try
            {
                foreach (var path in Directory.GetFiles(CestaVym, sablona))
                {
                    try { File.SetAttributes(path, prava); } catch (Exception ex) { Declare.Errors.Add(ex.ToString()); }
                }
            }
            catch (Exception ex) { Declare.Errors.Add(ex.ToString()); }
        }

        /// <summary>
        /// Zistuje, ci je olymp, ktory sa upgraduje, spusteny (WMI hlada 32 aj 64)
        /// </summary>
        /// <param name="destPath"></param>
        /// <param name="suborExe"></param>
        /// <returns></returns>
        public static bool JeSpustenyProgram(string destPath, string suborExe)
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
                Declare.Errors.Add(ex.ToString() + $"\r\npath: {path}");
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
            catch (Exception ex) { Declare.Errors.Add(ex.ToString() + $"\r\nmeno: {meno}"); return false; }
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
               
    }
}

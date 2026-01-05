using System;
using System.Diagnostics;
using System.IO;

namespace OlympUpgrade
{
    internal class LicenseVersionFunctions
    {
        /// <summary>
        /// vrati verziu upgrade a nastavi dalsie hodnoty zo suboru txt
        /// </summary>
        /// <returns></returns>
        public static string DajVerziuUpgrade()
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
            var verziaTxtPath = ZipFunctions.ExtractFileFromUpgradeZip(tempAdr, Declare.VERZIA_TXT);
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
            var cRV2KrosFile = ZipFunctions.ExtractFileFromUpgradeZip(tempAdr, "CRV2Kros.exe");
            if (string.IsNullOrEmpty(cRV2KrosFile))
            {
                Declare.ExitProg(Declare.ID_CHYBA_CHYBA_ZIP2);
                return string.Empty;
            }

            LicenseVersionFunctions.DajVerziuExe(Path.Combine(tempAdr, "CRV2Kros.exe"), out Declare.N_CRV2_MAJOR, out Declare.N_CRV2_MINOR, out Declare.N_CRV2_REVISION);
            HelpFunctions.TryToDeleteFile(cRV2KrosFile);

            return LicenseVersionFunctions.DajVerziuString(Declare.MAJOR, Declare.MINOR, Declare.REVISION);
        }

        public static bool ReadVersionTxt(string verzieTxtPath)
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
        /// Vrati verziu .exe
        /// </summary>
        /// <returns></returns>
        public static string DajVerziuProgramu()
        {
            var pathExe = Path.Combine(Declare.DEST_PATH, Declare.SUBOR_EXE);
            var pathExeStary = Path.Combine(Declare.DEST_PATH, Declare.SUBOR_EXE_STARY);

            if (File.Exists(pathExe))
            {
                LicenseVersionFunctions.DajVerziuExe(pathExe, out Declare.P_MAJOR, out Declare.P_MINOR, out Declare.P_REVISION);
                if (Declare.P_MAJOR == 0)
                    return Declare.VERZIA_NEZNAMA;
                else
                    return LicenseVersionFunctions.DajVerziuString(Declare.P_MAJOR, Declare.P_MINOR, Declare.P_REVISION);
            }
            else if (File.Exists(pathExeStary))
            {
                LicenseVersionFunctions.DajVerziuExe(pathExeStary, out Declare.P_MAJOR, out Declare.P_MINOR, out Declare.P_REVISION);
                if (Declare.P_MAJOR == 0)
                    return Declare.VERZIA_NEZNAMA;
                else
                    return LicenseVersionFunctions.DajVerziuString(Declare.P_MAJOR, Declare.P_MINOR, Declare.P_REVISION);
            }

            Declare.P_MAJOR = 0;
            return Declare.VERZIA_NEZNAMA;
        }

        public static void CitajRegistracnySubor(int PcD, string cesta, string cestaInstalacie = "")
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

        private static void NacitajNovuLicenciu(int PcD, string cesta)
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
    }
}

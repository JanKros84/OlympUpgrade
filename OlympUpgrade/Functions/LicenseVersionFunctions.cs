using System;
using System.IO;

namespace OlympUpgrade
{
    internal class LicenseVersionFunctions
    {

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

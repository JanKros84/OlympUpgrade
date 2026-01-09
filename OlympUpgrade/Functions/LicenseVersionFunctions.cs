using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.Remoting.Contexts;
using System.Text;

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
            else if (!HelpFunctions.ExistujeSubor(cesta + Declare.LICENCIA))
            {
                //NEBOL NAJDENY REGISTRACNY SUBOR
                if (PcD == Declare.PCD_POCITAC)
                    Declare.VerziaPC = 0;
                else
                    Declare.VerziaDisketa = 0;

                return;
            }

            cesta += Declare.LICENCIA;


            StringBuilder kodovanyStrB = new StringBuilder();
            using (FileStream fs = new FileStream(cesta, FileMode.Open, FileAccess.Read))
            {
                int b;
                while ((b = fs.ReadByte()) != -1)
                {
                    kodovanyStrB.Append((char)b);
                }
            }

            //prazdny mzdy.reg, == chybny
            if (kodovanyStrB.Length < 13)
            {
                if (PcD == Declare.PCD_POCITAC)
                    Declare.VerziaPC = Declare.LIC_CHYBNA;
                else
                    Declare.VerziaDisketa = Declare.LIC_CHYBNA;
                return;
            }

            string text = kodovanyStrB.ToString().Substring(0, 8);
            string kodovany = kodovanyStrB.ToString().Substring(10, kodovanyStrB.Length - 12); //bez konca riadku

            //zisti crc zakodovaneho textu, ci je dobry regoverenie crc
            if (CodingFunctions.VypocitajCrc(kodovany, kodovany.Length, 0) != text)
            {
                //CHYBNY REGISTRACNY SUBOR
                if (PcD == Declare.PCD_POCITAC)
                    Declare.VerziaPC = Declare.LIC_CHYBNA;
                else
                    Declare.VerziaDisketa = Declare.LIC_CHYBNA;
                return;
            }

            //spracuj rozkodovany text - nacitaj cestu k datam
            text = string.Empty;
            CodingFunctions.Odhesluj(kodovany, out text);

            int i = 0, j = -1;

            //Nacitaj nazov firmy
            i = j + 1;
            j = text.IndexOf("\r\n", i);
            string nazovFirmy = text.Substring(i, j - i);

            if (PcD == Declare.PCD_INSTALACKY)
                Declare.NazovFirmyDisketa = nazovFirmy;
            else
                Declare.NazovFirmyPC = nazovFirmy;

            //Preskoc udaje o partnerovi,nastavenia
            for (int x = 2; x <= 4; x++)
            {
                i = j + 2;
                j = text.IndexOf("\r\n", i);
            }

            //Nacitam ICO firmy
            i = j + 2;
            j = text.IndexOf("\r\n", i);
            string ico = /*HelpFunctions.ValLong*/(text.Substring(i, j - i));

            if (PcD == Declare.PCD_INSTALACKY)
                Declare.ICO_Disketa = ico;
            else
                Declare.ICO_PC = ico;

            //poziciu 6 preskocim
            i = j + 2;
            j = text.IndexOf("\r\n", i);
            int typInstalacieX = (i < j) ? (int)HelpFunctions.ValLong(text.Substring(i, j - i)) : 0; //typ instalacie

            int sasTypLicencie = 0;

            if (PcD == Declare.PCD_INSTALACKY)
            {
                Declare.VerziaDisketa = (typInstalacieX == 1 || typInstalacieX == 2 || typInstalacieX == 4)
                                ? Declare.LIC_OSTRA
                                : Declare.LIC_UPDATE;

                //precitaj datum registracie
                for (int x = 7; x <= 10; x++)
                {
                    i = j + 2;
                    j = text.IndexOf("\r\n", i);
                }

                if (j > 0)
                {
                    i = j + 2;
                    j = text.IndexOf("\r\n", i);
                    Declare.DatumDisketa = (i < j) ? HelpFunctions.ValLong(text.Substring(i, j - i)) : 0;
                }

                for (int x = 12; x <= 14; x++)
                {
                    i = j + 2;
                    j = text.IndexOf("\r\n", i);
                }
                Declare.PorCisloDisketa = /*HelpFunctions.ValLong*/(text.Substring(i, j - i));

                for (int x = 16; x <= 22; x++)
                {
                    i = j + 2;
                    j = text.IndexOf("\r\n", i);
                }

                if (j - i >= 0)
                    sasTypLicencie = (int)HelpFunctions.ValLong(text.Substring(i, j - i));

                //ak je sas licencia, tak datum neposuvam, inak ho posuniem o jeden rok
                if (sasTypLicencie == 0 && Declare.DatumDisketa != 0)
                    Declare.DatumDisketa += 10000;
            }
            else //licencia na pocitaci
            {
                //preskoc, co nie je potrebne citat
                for (int x = 7; x <= 10; x++)
                {
                    i = j + 2;
                    j = text.IndexOf("\r\n", i);
                }

                //precitaj cestu k datam
                string pom = (j > 0) ? text.Substring(i, j - i) : text.Substring(i);
                //skontroluj cestu, ci existuje
                if (HelpFunctions.ExistujeSubor(pom))
                    cestaInstalacie = pom;

                //precitaj datum registracie
                if (j > 0)
                {
                    i = j + 2;
                    j = text.IndexOf("\r\n", i);

                    // ++++++++++++++
                    //nasledujuci riadok je tu z dovodu chyby v programe OLYMP,
                    //ktora bola od verzie 2.05(april99) a ktora sposobovala, ze pri
                    //zapise cesty k datam do registracneho suboru v programe sa za cestu
                    //pridal prazdny riadok a teda sa pozicia ostatnych (datum) udajov posunula
                    //a zle sa nacitavali tieto udaje
                    if (i == j)
                    {
                        i = j + 2;
                        j = text.IndexOf("\r\n", i);
                    }
                    //koniec riadku doplneneho na zachranu registracky s posunom

                    Declare.DatumPC = (i < j) ? HelpFunctions.ValLong(text.Substring(i, j - i)) : 0;
                }

                for (int x = 12; x <= 14; x++)
                {
                    i = j + 2;
                    j = text.IndexOf("\r\n", i);
                }

                Declare.PorCisloPC = /*HelpFunctions.ValLong*/(text.Substring(i, j - i));

                for (int x = 16; x <= 22; x++)
                {
                    i = j + 2;
                    j = text.IndexOf("\r\n", i);
                }

                if (j - i >= 0)
                    sasTypLicencie = (int)HelpFunctions.ValLong(text.Substring(i, j - i));

                //ak je sas licencia, tak datum neposuvam, inak ho posuniem o jeden rok
                if (sasTypLicencie == 0 && Declare.DatumPC != 0)
                    Declare.DatumPC += 10000;

                //DOBRY REG. CESTA,KAM MOZE INSTALOVAT NASTAVENA
                Declare.VerziaPC = Declare.LIC_OSTRA;
            }
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

using Microsoft.Win32;
using System;
using System.Security.AccessControl;

namespace OlympUpgrade
{
    internal class RegistryFunctions
    {
        /// <summary>
        /// odstrani z registrov spustenie tohoto programu po restarte z instalacneho CD ak tam nic nie je tak sa nic nestane
        /// </summary>
        public static void OdstranZRegistrovDoinstalovanie()
        {
            ZmazHodnotuReg(Registry.CurrentUser/*Declare.HKEY_CURRENT_USER*/, @"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "krosOlymp");

            if (Declare.RESTARTUJ)
                Declare.RESTARTUJ = false;
        }

        /// <summary>
        /// Zistím z reg., či daná licencia už bola na PC – ak áno, nasmerujem inštaláciu na jej adresár
        /// </summary>
        /// <returns></returns>
        public static string GetOlympFolderBaseOnLic()
        {
            string path = string.Empty;
            var postKey = Declare.VerziaDisketa == Declare.LIC_OSTRA
                ? "Vyplaty" + CodingFunctions.VratIDRegistracky(Declare.ICO_Disketa, Declare.PorCisloDisketa)
                : string.Empty;

            if (CitajHodnotuReg(Registry.LocalMachine,
                               $@"Software\Kros\Olymp\{postKey}",
                               "InstalacnyAdresar",
                               out object val))
            {
                if (val is string licPath
                    && !string.IsNullOrWhiteSpace(licPath))
                    path = licPath;
            }

            return path;
        }

        /// <summary>
        /// Zistím z reg., či je nainštalovaný OLYMP -> čítam InstallLocation z HKLM\...\Uninstall\{ProductCode}
        /// </summary>
        /// <param name="path"></param>
        /// <param name="nacital"></param>
        public static bool CheckRegIfOlympIsInstalled(out string path)
        {
            path = null;
            var nacital = false;
            foreach (var productCode in new[] { Declare.PRODUCT_CODE_1, Declare.PRODUCT_CODE_2 })
            {
                if (CitajHodnotuReg(Registry.LocalMachine,
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
            return nacital;
        }

        public static void UlozDoRegistrovVerziu()
        {
            //iba ak sa instalovalo do povodneho adresara
            if (Declare.KTO_VOLAL == Declare.VOLAL_INSTALL
                || string.Equals(Declare.DEST_PATH, Declare.DEST_PATH_INSTALLSHIELD, StringComparison.OrdinalIgnoreCase))
            {
                string RegProductCode = Declare.REG_PRODUCT_CODE_2;
                string ProductCode = Declare.PRODUCT_CODE_2;

                //zobrazovane v add/remove
                if (ExistujeKlucReg(Registry.ClassesRoot, $@"Installer\Products\{RegProductCode}"))
                {
                    ZapisHodnotuReg(
                        Registry.ClassesRoot,
                        $@"Installer\Products\{RegProductCode}",
                        "ProductName",
                        $"OLYMP {LicenseVersionFunctions.DajVerziuString(Declare.MAJOR, Declare.MINOR, Declare.REVISION)}");
                }

                //MessageBox.Show("BEFORE reg OLYMP22 S-1-5-18");
                //zobrazovane v MSI pri instalacii vyssej verzie
                string installProps = $@"SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products\{RegProductCode}\InstallProperties";
                if (ExistujeKlucReg(Registry.LocalMachine, installProps))
                {
                    ZapisHodnotuReg(
                        Registry.LocalMachine,
                        installProps,
                        "DisplayName",
                        $"OLYMP {LicenseVersionFunctions.DajVerziuString(Declare.MAJOR, Declare.MINOR, Declare.REVISION)}");

                    //MessageBox.Show("AFTER reg OLYMP22");
                }
                //$@"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{ProductCode}";
                string uninstallKey = $@"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{ProductCode}";
                if (ExistujeKlucReg(Registry.LocalMachine, uninstallKey))
                {
                    ZapisHodnotuReg(
                        Registry.LocalMachine,
                        uninstallKey,
                        "DisplayName",
                        $"OLYMP {LicenseVersionFunctions.DajVerziuString(Declare.MAJOR, Declare.MINOR, Declare.REVISION)}");

                    ZapisHodnotuReg(
                       Registry.LocalMachine,
                        uninstallKey,
                        "DisplayVersion",
                        LicenseVersionFunctions.DajVerziuString(Declare.MAJOR, Declare.MINOR, Declare.REVISION));
                }
            }

            //este zapiseme adresar do ktoreho sme naposledy instalovali danu registracku
            if (Declare.VerziaDisketa == Declare.LIC_OSTRA)  // ak som instaloval ostru, tak sa mi zmenila ostra na PC
            {
                Declare.ICO_PC = Declare.ICO_Disketa;
                Declare.PorCisloPC = Declare.PorCisloDisketa;
            }

            Vytvor_kluc(Registry.LocalMachine, @"Software\Kros\Olymp");
            if (ExistujeKlucReg(Registry.LocalMachine, @"Software\Kros\Olymp"))
            {
                ZapisHodnotuReg(
                       Registry.LocalMachine,
                        @"Software\Kros\Olymp",
                        "InstalacnyAdresar",
                        Declare.DEST_PATH ?? string.Empty);
            }

            if ((Declare.VerziaPC == Declare.LIC_OSTRA || Declare.VerziaDisketa == Declare.LIC_OSTRA)
                && Declare.VerziaDisketa != Declare.LIC_NEEXISTUJE)
            {
                string pomCesta = $@"Software\Kros\Olymp\Vyplaty{CodingFunctions.VratIDRegistracky(Declare.ICO_PC, Declare.PorCisloPC)}";

                if (!ExistujeKlucReg(Registry.LocalMachine, pomCesta))
                    Vytvor_kluc(Registry.LocalMachine, pomCesta);

                if (ExistujeKlucReg(Registry.LocalMachine, pomCesta))
                {
                    ZapisHodnotuReg(
                      Registry.LocalMachine,
                        pomCesta,
                        "InstalacnyAdresar",
                        Declare.DEST_PATH ?? string.Empty);
                }
            }
        }


        #region RegistriOperations

        private static bool ZmazHodnotuReg(RegistryKey skupina_kluca, string meno_kluca, string meno_polozky)
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
                Declare.Errors.Add(ex.ToString() + $"\r\nskupina_kluca: {skupina_kluca.Name}\r\nmeno_kluca: {meno_kluca}\r\nmeno_polozky: {meno_polozky}");
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
            catch (Exception ex)
            {
                Declare.Errors.Add(ex.ToString() + $"\r\nskupina_kluca: {skupina_kluca.Name}\r\nmeno_kluca: {meno_kluca}\r\nmeno_polozky: {meno_polozky}");
                return false;
            }
        }

        private static bool Vytvor_kluc(RegistryKey skupina_kluca, string meno_kluca)
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
                Declare.Errors.Add(ex.ToString() + $"\r\nskupina_kluca: {skupina_kluca.Name}\r\nmeno_kluca: {meno_kluca}");
                return false;
            }
        }

        private static void ZapisHodnotuReg(RegistryKey skupina_kluca, string meno_kluca, string meno_polozky, string hodnota_polozky)//object hodnota_polozky, int typ_polozky)
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
            catch (Exception ex)
            {
                Declare.Errors.Add(ex.ToString() + $"\r\nskupina_kluca: {skupina_kluca.Name}\r\nmeno_kluca: {meno_kluca}\r\nmeno_polozky: {meno_polozky}\r\nhodnota_polozky: {hodnota_polozky}");
            }
        }

        private static bool ExistujeKlucReg(RegistryKey skupina_kluca, string meno_kluca)
        {
            try
            {
                using (RegistryKey hKey = skupina_kluca.OpenSubKey(meno_kluca, RegistryRights.QueryValues))
                    return hKey != null;
            }
            catch (Exception ex) { Declare.Errors.Add(ex.ToString() + $"\r\nskupina_kluca: {skupina_kluca.Name}\r\nmeno_kluca: {meno_kluca}"); return false; }
        }

        #endregion RegistriOperations
    }
}

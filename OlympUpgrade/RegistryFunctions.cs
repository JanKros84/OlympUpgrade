using Microsoft.Win32;
using System;

namespace OlympUpgrade
{
    internal class RegistryFunctions
    {
        /// <summary>
        /// odstrani z registrov spustenie tohoto programu po restarte z instalacneho CD ak tam nic nie je tak sa nic nestane
        /// </summary>
        public static void OdstranZRegistrovDoinstalovanie()
        {
            RegistryOperation.ZmazHodnotuReg(Registry.CurrentUser/*Declare.HKEY_CURRENT_USER*/, @"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "krosOlymp");

            if (Declare.RESTARTUJ)
                Declare.RESTARTUJ = false;
        }

        /// <summary>
        /// Zistím, či daná licencia už bola na PC – ak áno, nasmerujem inštaláciu na jej adresár
        /// </summary>
        /// <returns></returns>
        public static string GetOlympFolderBaseOnLic()
        {
            string path = string.Empty;
            var postKey = Declare.VerziaDisketa == Declare.LIC_OSTRA
                ? "Vyplaty" + kodovanie.VratIDRegistracky(Declare.ICO_Disketa, Declare.PorCisloDisketa)
                : string.Empty;

            if (RegistryOperation.CitajHodnotuReg(Registry.LocalMachine,
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
        /// Zistím, či je nainštalovaný OLYMP -> čítam InstallLocation z HKLM\...\Uninstall\{ProductCode}
        /// </summary>
        /// <param name="path"></param>
        /// <param name="nacital"></param>
        public static bool CheckRegIfOlympIsInstalled(out string path)
        {
            path = null;
            var nacital = false;
            foreach (var productCode in new[] { Declare.PRODUCT_CODE_1, Declare.PRODUCT_CODE_2 })
            {
                if (RegistryOperation.CitajHodnotuReg(Registry.LocalMachine,
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
                if (RegistryOperation.ExistujeKlucReg(Registry.ClassesRoot, $@"Installer\Products\{RegProductCode}"))
                {
                    RegistryOperation.ZapisHodnotuReg(
                        Registry.ClassesRoot,
                        $@"Installer\Products\{RegProductCode}",
                        "ProductName",
                        $"OLYMP {HelpFunctions.DajVerziuString(Declare.MAJOR, Declare.MINOR, Declare.REVISION)}");
                }

                //MessageBox.Show("BEFORE reg OLYMP22 S-1-5-18");
                //zobrazovane v MSI pri instalacii vyssej verzie
                string installProps = $@"SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products\{RegProductCode}\InstallProperties";
                if (RegistryOperation.ExistujeKlucReg(Registry.LocalMachine, installProps))
                {
                    RegistryOperation.ZapisHodnotuReg(
                        Registry.LocalMachine,
                        installProps,
                        "DisplayName",
                        $"OLYMP {HelpFunctions.DajVerziuString(Declare.MAJOR, Declare.MINOR, Declare.REVISION)}");

                    //MessageBox.Show("AFTER reg OLYMP22");
                }

                string uninstallKey = $@"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{ProductCode}";
                if (RegistryOperation.ExistujeKlucReg(Registry.LocalMachine, uninstallKey))
                {
                    RegistryOperation.ZapisHodnotuReg(
                        Registry.LocalMachine,
                        uninstallKey,
                        "DisplayName",
                        $"OLYMP {HelpFunctions.DajVerziuString(Declare.MAJOR, Declare.MINOR, Declare.REVISION)}");

                    RegistryOperation.ZapisHodnotuReg(
                       Registry.LocalMachine,
                        uninstallKey,
                        "DisplayVersion",
                        HelpFunctions.DajVerziuString(Declare.MAJOR, Declare.MINOR, Declare.REVISION));
                }
            }

            //este zapiseme adresar do ktoreho sme naposledy instalovali danu registracku
            if (Declare.VerziaDisketa == Declare.LIC_OSTRA)  // ak som instaloval ostru, tak sa mi zmenila ostra na PC
            {
                Declare.ICO_PC = Declare.ICO_Disketa;
                Declare.PorCisloPC = Declare.PorCisloDisketa;
            }

            RegistryOperation.Vytvor_kluc(Registry.LocalMachine, @"Software\Kros\Olymp");
            if (RegistryOperation.ExistujeKlucReg(Registry.LocalMachine, @"Software\Kros\Olymp"))
            {
                RegistryOperation.ZapisHodnotuReg(
                       Registry.LocalMachine,
                        @"Software\Kros\Olymp",
                        "InstalacnyAdresar",
                        Declare.DEST_PATH ?? string.Empty);
            }

            if ((Declare.VerziaPC == Declare.LIC_OSTRA || Declare.VerziaDisketa == Declare.LIC_OSTRA)
                && Declare.VerziaDisketa != Declare.LIC_NEEXISTUJE)
            {
                string pomCesta = $@"Software\Kros\Olymp\Vyplaty{kodovanie.VratIDRegistracky(Declare.ICO_PC, Declare.PorCisloPC)}";

                if (!RegistryOperation.ExistujeKlucReg(Registry.LocalMachine, pomCesta))
                    RegistryOperation.Vytvor_kluc(Registry.LocalMachine, pomCesta);

                if (RegistryOperation.ExistujeKlucReg(Registry.LocalMachine, pomCesta))
                {
                    RegistryOperation.ZapisHodnotuReg(
                      Registry.LocalMachine,
                        pomCesta,
                        "InstalacnyAdresar",
                        Declare.DEST_PATH ?? string.Empty);
                }
            }
        }

        /// <summary>
        /// 0 - nie je nainstalovany
        ///-1 - je nainstalovany acrobat reader
        /// 1 - je nainstalovany acrobat (aj writer)
        /// </summary>
        /// <returns></returns>
        public static int IsAcrobatReaderInstalled()
        {
            if (RegistryOperation.CitajHodnotuReg(Registry.LocalMachine,
                                    @"Software\Microsoft\Windows\CurrentVersion\App Paths\ACRORD32.EXE",
                                    "",
                                    out object _))
                return -1;

            if (RegistryOperation.CitajHodnotuReg(Registry.LocalMachine,
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
            if (RegistryOperation.CitajHodnotuReg(Registry.LocalMachine,
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
    }
}

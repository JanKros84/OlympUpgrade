using Microsoft.Win32;
using System;
using System.Security.AccessControl;

namespace OlympUpgrade
{
    internal class RegistryOperation
    {
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
                Declare.Errors.Add(ex.ToString() + $"\r\nskupina_kluca: {skupina_kluca.Name}\r\nmeno_kluca: {meno_kluca}");
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
            catch (Exception ex)
            {
                Declare.Errors.Add(ex.ToString() + $"\r\nskupina_kluca: {skupina_kluca.Name}\r\nmeno_kluca: {meno_kluca}\r\nmeno_polozky: {meno_polozky}\r\nhodnota_polozky: {hodnota_polozky}");
            }
        }

        public static bool ExistujeKlucReg(RegistryKey skupina_kluca, string meno_kluca)
        {
            try
            {
                using (RegistryKey hKey = skupina_kluca.OpenSubKey(meno_kluca, RegistryRights.QueryValues))
                    return hKey != null;
            }
            catch (Exception ex) { Declare.Errors.Add(ex.ToString() + $"\r\nskupina_kluca: {skupina_kluca.Name}\r\nmeno_kluca: {meno_kluca}"); return false; }
        }
    }
}

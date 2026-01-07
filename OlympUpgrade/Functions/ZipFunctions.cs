using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OlympUpgrade
{
    internal class ZipFunctions
    {
        public static string ExtractFileFromUpgradeZip(string desAdr, string fileName)
        {
            string resFilePath = Path.Combine(desAdr, fileName); 
            try
            {
                using (ZipArchive zip = ZipFile.Open(Path.Combine(Declare.AKT_ADRESAR, Declare.SUBOR_ZIP), ZipArchiveMode.Read, Declare.ZipEncoding))
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
            catch (Exception ex)
            {
                Declare.Errors.Add(ex.ToString());
                HelpFunctions.TryToDeleteFile(resFilePath);
                return string.Empty;
            }
        }

        /// <summary>
        /// vrati nekomprimovanu velkost suborov v zip v bytoch
        /// </summary>
        /// <param name="pathZip"></param>
        /// <returns></returns>
        public static long DajVelkostSuborovVZip(string pathZip, Action<long> progresiaPreset, Action progresiaTik)
        {
            if (string.IsNullOrWhiteSpace(pathZip) || !File.Exists(pathZip))
                return 0;

            long sum = 0;

            using (ZipArchive zip = ZipFile.Open(pathZip, ZipArchiveMode.Read, Declare.ZipEncoding))
            {
                progresiaPreset?.Invoke(zip.Entries.Count);

                foreach (var e in zip.Entries)
                {
                    try
                    {
                        // Rátame iba súbory (nie adresáre)
                        bool isDirectory = e.Name.Length == 0 || e.FullName.EndsWith("/", StringComparison.Ordinal);
                        if (!isDirectory)
                        {
                            sum += e.Length; // uncompressed size/
                        }
                    }
                    catch (Exception ex) { Declare.Errors.Add(ex.ToString()); }

                    progresiaTik?.Invoke();
                }
            }
            return sum;
        }
    }
}

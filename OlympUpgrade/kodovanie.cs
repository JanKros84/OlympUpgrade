using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OlympUpgrade
{
    internal class Kodovanie
    {
        public static string Zahesluj(string zdroj)
        {
            //'Funkcia zahesluje vstupny reazec
            //   Zahesluj = "": znak = &H71: sucet = 0
            //   For i = 1 To Len(Zdroj)
            //      znak = znak Xor(Asc(Mid(Zdroj, i, 1)))
            //      sucet = (sucet + znak) Mod 256
            //      Zahesluj = Zahesluj & byte2hexstr(znak Xor & H75)
            //   Next
            //   'Nakoniec reazca pripoj kontrolny súèet
            //   Zahesluj = Zahesluj & byte2hexstr(sucet Xor & H75)


            byte znak = 0x71;
            int sucet = 0;

            byte[] bytes = Encoding.UTF8.GetBytes(zdroj);

            var sb = new StringBuilder(bytes.Length * 2 + 2);

            foreach (byte b in bytes)
            {
                // znak = znak Xor Asc(char)
                znak = (byte)(znak ^ b);

                // sucet = (sucet + znak) Mod 256
                sucet = (sucet + znak) & 0xFF;

                // výstupný bajt: znak Xor &H75
                byte outb = (byte)(znak ^ 0x75);
                sb.Append(outb.ToString("X2"));
            }

            // na koniec pripoj kontrolný súčet (sucet Xor &H75)
            byte checksum = (byte)((sucet ^ 0x75) & 0xFF);
            sb.Append(checksum.ToString("X2"));

            return sb.ToString();
        }

        public static bool Odhesluj(string zdroj, out string ciel)
        {
            ciel = string.Empty;

            // Nemožno odheslovať, ak je počet znakov nepárny alebo 0
            if (string.IsNullOrEmpty(zdroj) || (zdroj.Length % 2) != 0)
                return false;

            // znak = &H71 Xor &H75
            byte znak = (byte)(0x71 ^ 0x75); // 0x04
            int sucet = 0;

            int pairCount = zdroj.Length / 2;

            // 1) spočítaj súčet všetkých okrem posledného páru (checksum)
            for (int i = 1; i <= pairCount - 1; i++)
            {
                byte b = HexStr2byte(zdroj.Substring(2 * i - 2, 2));
                sucet = (sucet + (b ^ 0x75)) & 0xFF; // Mod 256
            }

            // 2) validácia checksumu (posledný pár)
            byte last = HexStr2byte(zdroj.Substring(zdroj.Length - 2, 2));
            bool ok = sucet == ((last ^ 0x75) & 0xFF);
            if (!ok) return false;

            // 3) dekódovanie obsahu
            var bytes = new List<byte>(pairCount - 1);
            for (int i = 1; i <= pairCount - 1; i++)
            {
                byte cur = HexStr2byte(zdroj.Substring(2 * i - 2, 2));
                byte decoded = (byte)(znak ^ cur);
                bytes.Add(decoded);
                znak = cur;
            }

            // VB6 Chr() pracuje v ANSI – použijeme Windows-1250 (uprav ak máš inú kódovú stránku)
            ciel = Encoding.GetEncoding(1250).GetString(bytes.ToArray());
            return true;
        }

        public static byte HexStr2byte(string hexstr)
        {
            if (string.IsNullOrWhiteSpace(hexstr)) return 0;
            if (byte.TryParse(hexstr, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var b))
                return b;
            return 0; // pri chybe VB6 tiež vráti 0
        }

        public static string VratIDRegistracky(string ICO, string PorCislo)
        {
            //vrati retazec pod akym je v registroch jednoznacne identifikovana dana registracka
            //(pre pripad ze by som to dakedy chcel menit), pouzijem toto vsade
            return Zahesluj(ICO + PorCislo);
        }
    }
}


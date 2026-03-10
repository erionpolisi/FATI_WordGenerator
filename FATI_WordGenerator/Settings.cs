using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FATI_WordGenerator
{
    public class Settings
    {
        public bool AutoCloseTerminal { get; set; } = false;
        public bool OpenWordDocument { get; set; } = false;

        static public Settings LoadSettings()
        {
            var settings = new Settings();
            string path = "";
            bool isDebug = false;

#if DEBUG
            path = Path.Combine(
                AppDomain.CurrentDomain.BaseDirectory,
                @"..\..\..\settings.txt");
            isDebug = true;
#else
            path = Path.Combine(
                AppDomain.CurrentDomain.BaseDirectory,
                "settings.txt");
#endif

            if (!File.Exists(path))
            {
                Console.WriteLine("Keine settings.txt gefunden. Standardwerte werden verwendet.");
                return settings;
            }


            foreach (var line in File.ReadLines(path))
            {
                if (line.StartsWith("#") || string.IsNullOrWhiteSpace(line))
                    continue;

                var parts = line.Split('=');

                if (parts.Length != 2)
                    continue;

                bool value = bool.Parse(parts[1].Trim());

                if (isDebug)
                {
                    Console.WriteLine($"Einstellung geladen: {parts[0].Trim()} = {value}");
                }

                switch (parts[0].Trim())
                {
                    case "AutoCloseTerminal":
                        settings.AutoCloseTerminal = value;
                        break;

                    case "OpenWordDocument":
                        settings.OpenWordDocument = value;
                        break;
                }
            }

            return settings;
        }
    }


}

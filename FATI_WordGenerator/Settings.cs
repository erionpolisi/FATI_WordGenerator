namespace FATI_WordGenerator
{
    public class Settings
    {
        public bool AutoCloseTerminal { get; set; } = false;
        public bool OpenWordDocument { get; set; } = false;
        public bool ChangePath { get; set; } = false;
        public string InvoiceFolderPath { get; set; } = string.Empty;

        static public Settings LoadSettings()
        {
            var settings = new Settings();
            string path = GetSettingsFilePath();
            bool isDebug = false;

#if DEBUG
            isDebug = true;
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

                var parts = line.Split('=', 2);

                if (parts.Length != 2)
                    continue;

                var key = parts[0].Trim();
                var rawValue = parts[1].Trim();

                switch (key)
                {
                    case "AutoCloseTerminal":
                        if (bool.TryParse(rawValue, out var autoClose))
                            settings.AutoCloseTerminal = autoClose;
                        break;

                    case "OpenWordDocument":
                        if (bool.TryParse(rawValue, out var openWord))
                            settings.OpenWordDocument = openWord;
                        break;

                    case "ChangePath":
                        if (bool.TryParse(rawValue, out var changePath))
                            settings.ChangePath = changePath;
                        break;

                    case "InvoiceFolderPath":
                        settings.InvoiceFolderPath = rawValue;
                        break;

                    default:
                        if (isDebug)
                        {
                            Console.WriteLine($"Unbekannte Einstellung: {key}");
                        }
                        break;
                }

                if (isDebug)
                {
                    Console.WriteLine($"Einstellung geladen: {key} = {rawValue}");
                }
            }

            return settings;
        }

        public static void SaveSettings(Settings settings)
        {
            string path = GetSettingsFilePath();

            var lines = new List<string>
            {
                "# Einstellungen für FATI Generator",
                "",
                $"AutoCloseTerminal={settings.AutoCloseTerminal}",
                $"OpenWordDocument={settings.OpenWordDocument}",
                $"ChangePath={settings.ChangePath}",
                $"InvoiceFolderPath={settings.InvoiceFolderPath}"
            };

            File.WriteAllLines(path, lines);
        }

        public static void ResetSettings()
        {
            SaveSettings(new Settings());
        }

        public static string GetSettingsFilePath()
        {
#if DEBUG
            return Path.Combine(
                AppDomain.CurrentDomain.BaseDirectory,
                @"..\..\..\settings.txt");
#else
            return Path.Combine(
                AppDomain.CurrentDomain.BaseDirectory,
                "settings.txt");
#endif
        }
    }


}

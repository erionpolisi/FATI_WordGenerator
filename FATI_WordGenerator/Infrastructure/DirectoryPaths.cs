using System.Windows.Forms;

namespace FATI_WordGenerator.Infrastructure

{
    public class PathService
    {
        public string ResourcesPath { get; }
        public string CurrentDirectory { get; }
        public string InvoiceFolder { get; }
        public string InputExcelPath { get; }

        public PathService(Settings settings)
        {
#if DEBUG
            ResourcesPath = Path.Combine(
                AppDomain.CurrentDomain.BaseDirectory,
                @"..\..\..\Resources");
#else
    ResourcesPath = Path.Combine(
        AppDomain.CurrentDomain.BaseDirectory,
        "Resources");
#endif

            CurrentDirectory = AppDomain.CurrentDomain.BaseDirectory;

            if (settings.ChangePath)
            {
                if (string.IsNullOrWhiteSpace(settings.InvoiceFolderPath) ||
                    !Directory.Exists(settings.InvoiceFolderPath))
                {
                    Console.WriteLine("Bitte wähle einen bestehenden Rechnungsordner...");

                    using var dialog = new FolderBrowserDialog
                    {
                        Description = "Rechnungsordner auswählen",
                        UseDescriptionForTitle = true
                    };

                    if (dialog.ShowDialog() == DialogResult.OK)
                    {
                        InvoiceFolder = dialog.SelectedPath;

                        settings.InvoiceFolderPath = InvoiceFolder;
                        Settings.SaveSettings(settings);
                    }
                    else
                    {
                        throw new Exception("Kein Ordner ausgewählt!");
                    }
                }
                else
                {
                    InvoiceFolder = settings.InvoiceFolderPath;
                }
            }
            else
            {
                InvoiceFolder = Path.Combine(CurrentDirectory, "Rechnungen");
                Directory.CreateDirectory(InvoiceFolder);
            }

            InputExcelPath = Path.Combine(ResourcesPath, "Input.xlsx");
        }
    }
}
namespace FATI_WordGenerator.Infrastructure

{
    public class PathService
    {
        public string ResourcesPath { get; }
        public string CurrentDirectory { get; }
        public string InvoiceFolder { get; }
        public string InputExcelPath { get; }

        public PathService()
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

            InvoiceFolder = Path.Combine(CurrentDirectory, "Rechnungen");
            Directory.CreateDirectory(InvoiceFolder);

            InputExcelPath = Path.Combine(ResourcesPath, "Input.xlsx");
        }
    }
}
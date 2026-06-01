using FATI_WordGenerator.Domain;
using FATI_WordGenerator.Infrastructure;

namespace FATI_WordGenerator.Services
{
    internal class InvoiceService(Settings settings)
    {
        private readonly PathService _paths = new (settings);
        private readonly ExcelService _excelService = new();
        private readonly WordService _wordService = new();

        public readonly int _year = DateTime.Now.Year;
        public string FileName { get; private set; } = string.Empty;

        public int Increment { get; private set; }

        internal string Generate(Action<int, string> progress)
        {
            try
            {
                progress(1, "Excel wird gelesen...");
                var invoice = _excelService.ReadInvoice(_paths.InputExcelPath);

                progress(2, "Rechnungsnummer wird berechnet...");
                Increment = GetIncrement(invoice);

                progress(3, "Word Template wird kopiert...");
                var outputPath = CopyTemplate(invoice);

                progress(4, "Word Dokument wird generiert...");
                _wordService.Generate(outputPath, invoice, Increment);

                return outputPath;
            }
            catch
            {
                throw;
            }
        }

        private int GetInkrementUnique(Invoice invoice)
        {
            return Directory
                .GetFiles(_paths.InvoiceFolder, $"*_{_year}_{invoice.Company.Name}_*.docx")//Old
                .Length + 1;
        }

        private int GetIncrement(Invoice invoice)
        {
            var count = Directory
                .EnumerateFiles(_paths.InvoiceFolder, "*.docx")
                .Count();

            return count + 1;
        }

        public string CopyTemplate(Invoice invoice)
        {
            try
            {
                FileName = $"{Increment:D3}_{_year}_" +
                           $"{Clean(invoice.Company.Name)}_" +
                           $"{Clean(invoice.Company.Street)} {Clean(invoice.Company.Number)}.docx";

                string outputPath = Path.Combine(_paths.InvoiceFolder, FileName);

                string templatePath = string.Empty;

                if (settings.CalculateSkonto)
                {
                    templatePath = Path.Combine(_paths.ResourcesPath, "RechnungTemplate.docx");
                }
                else
                {
                    templatePath = Path.Combine(_paths.ResourcesPath, "RechnungTemplate2.docx");
                }

                File.Copy(templatePath, outputPath, true);

                return outputPath;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Template konnte nicht kopiert werden: {ex.Message}");
                return string.Empty;
            }
        }

        private string Clean(string input) //Against InvalidFileNameChars (e.g. `/´)
        {
            if (string.IsNullOrWhiteSpace(input))
                return string.Empty;

            foreach (var c in Path.GetInvalidFileNameChars())
            {
                input = input.Replace(c, ' ');
            }

            return input.Trim();
        }
    }
}

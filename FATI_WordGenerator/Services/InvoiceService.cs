using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using FATI_WordGenerator.Domain;
using FATI_WordGenerator.Infrastructure;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FATI_WordGenerator.Services
{
    internal class InvoiceService
    {
        private readonly PathService _paths = new();
        private readonly ExcelService _excelService = new();
        private readonly WordService _wordService = new();

        public readonly int _year = DateTime.Now.Year;

        public int Increment { get; private set; }

        internal string Generate()
        {
            try
            {
                var invoice = _excelService.ReadInvoice(_paths.InputExcelPath);

                Increment = GetInkrement(invoice);

                var outputPath = CopyTemplate(invoice);

                _wordService.Generate(outputPath, invoice, Increment);

                return outputPath;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Fehler beim Erstellen der Rechnung: {ex.Message}");
                throw;
            }
        }

        private int GetInkrement(Invoice invoice)
        {
            return Directory
                .GetFiles(_paths.InvoiceFolder, $"*_{_year}_{invoice.Company.Name}_*.docx")
                .Length + 1;
        }

        public string CopyTemplate(Invoice invoice)
        {
            try
            {
                string filename = $"{Increment}_{_year}_{invoice.Company.Name}_{invoice.Company.Street} {invoice.Company.Number}.docx";
                string outputPath = Path.Combine(_paths.InvoiceFolder, filename);
                string templatePath = Path.Combine(_paths.ResourcesPath, "RechnungTemplate.docx");
                File.Copy(templatePath, outputPath, true);

                return outputPath;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message, "Couldn't copy Template to TargetFolder");
                return string.Empty;
            }
        }
    }
}

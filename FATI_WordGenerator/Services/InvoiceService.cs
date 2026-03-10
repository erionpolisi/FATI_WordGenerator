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
        public string FileName { get; private set; } = string.Empty;

        public int Increment { get; private set; }

        internal string Generate(Action<int, string> progress)
        {
            try
            {
                progress(1, "Excel wird gelesen...");
                var invoice = _excelService.ReadInvoice(_paths.InputExcelPath);

                progress(2, "Rechnungsnummer wird berechnet...");
                Increment = GetInkrement(invoice);

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
                FileName = $"{Increment}_{_year}_{invoice.Company.Name}_{invoice.Company.Street} {invoice.Company.Number}.docx";
                string outputPath = Path.Combine(_paths.InvoiceFolder, FileName);
                string templatePath = Path.Combine(_paths.ResourcesPath, "RechnungTemplate.docx");
                File.Copy(templatePath, outputPath, true);

                return outputPath;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Template konnte nicht kopiert werden: {ex.Message}");
                return string.Empty;
            }
        }
    }
}

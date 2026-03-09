using ClosedXML.Excel;
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

        internal void Generate()
        {
            var invoice = _excelService.ReadInvoice(_paths.InputExcelPath);
            try
            {
                _wordService.Generate(invoice);
            }
            catch
            {

            }           
        }

        public int GetInkrement(Invoice invoice)
        {
            string jahr = DateTime.Now.Year.ToString();

            return Directory
                .GetFiles(_paths.InvoiceFolder, $"*_{jahr}_{invoice.Company.Name}_*.docx")
                .Length + 1;
        }
    }
}

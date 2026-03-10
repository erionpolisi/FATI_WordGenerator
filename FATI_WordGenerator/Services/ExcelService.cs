using ClosedXML.Excel;
using FATI_WordGenerator.Domain;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static Program;

namespace FATI_WordGenerator.Services
{
    public class ExcelService
    {
        public Invoice ReadInvoice(string path)
        {
            var invoice = new Invoice();
            using var workbook = new XLWorkbook(path);

            try
            {
                var wsAdresse = workbook.Worksheet("Adresse");
                var wsFirma = workbook.Worksheet("Firma");
                var wsPos = workbook.Worksheet("Positionen");

                invoice.Address = CreateAdressString(wsAdresse);
                invoice.Period = GetPeriod(wsAdresse);
                invoice.Company = GetCompany(wsFirma);
                invoice.Positions = ReadPositions(wsPos);
                invoice.Details = GetDetails(wsPos); //Details is included in the position sheet, because it is needed for the invoice template and should be easily editable by the user in the Excel file

                if (invoice == null)
                {
                    throw new ArgumentNullException("Es wurden keine Daten gefunden. Bitte füllen Sie die Excel Datei aus.");
                }

                if (invoice.Positions == null || invoice.Positions.Count == 0)
                    throw new ArgumentException("Es wurden keine Positionen gefunden. Bitte füllen Sie die Positionen Tabelle in der Excel Datei aus.");

                return invoice;

            }
            catch (Exception ex)
            {
                throw new Exception("Fehler beim Lesen der Excel Datei.", ex);
            }
        }

        private string CreateAdressString(IXLWorksheet wsAdresse)
        {
            string strasse = wsAdresse.Cell(2, 1).GetString();
            string nr = wsAdresse.Cell(2, 2).GetString();
            string bezirk = wsAdresse.Cell(2, 3).GetString();
            string ort = wsAdresse.Cell(2, 4).GetString();

            return $"BVH: {strasse} {nr}, {bezirk} {ort}";
        }

        private string GetPeriod(IXLWorksheet wsAdresse)
        {
            string zeitraumMonateRaw = wsAdresse.Cell(2, 5).GetString();
            string zeitraumJahr = wsAdresse.Cell(2, 6).GetString();

            string[] zeitraumMonate = zeitraumMonateRaw.Split("-");
            string zeitraum = string.Empty;
            int zeitraumMonatsZahl = 0;

            switch (zeitraumMonate.Length)
            {
                case 1:
                    zeitraumMonatsZahl = int.Parse(zeitraumMonate[0].Trim());

                    zeitraum = ((Months)zeitraumMonatsZahl) + " " + zeitraumJahr;
                    break;

                case 2:
                    zeitraumMonatsZahl = int.Parse(zeitraumMonate[0].Trim());
                    int zeitraumMonatsZahl2 = int.Parse(zeitraumMonate[1].Trim());

                    zeitraum = ((Months)zeitraumMonatsZahl) + " - " + ((Months)zeitraumMonatsZahl2) + " " + zeitraumJahr;
                    break;

                default:
                    throw new ArgumentException(
                        "Ungültiger Zeitraum. Bitte Input.xlsx korrekt ausfüllen.\n" +
                        "Beispiele:\n" +
                        "3        -> März\n" +
                        "3-4      -> März - April"
                    );
            }

            return zeitraum;
        }

        private Company GetCompany(IXLWorksheet wsFirma)
        {
            var company = new Company
            {
                Name = wsFirma.Cell(2, 1).GetString(),
                Street = wsFirma.Cell(2, 2).GetString(),
                Number = wsFirma.Cell(2, 3).GetString(),
                PLZ = wsFirma.Cell(2, 4).GetString(),
                City = wsFirma.Cell(2, 5).GetString(),
                ATU = wsFirma.Cell(2, 6).GetString()
            };

            company.StreetAndNumber = $"{company.Street} {company.Number}";
            company.PLZAndCity = $"{company.PLZ} {company.City}";

            return company;
        }

        private List<InvoicePosition> ReadPositions(IXLWorksheet wsPos)
        {
            var positions = new List<InvoicePosition>();

            foreach (var row in wsPos.RowsUsed().Skip(1))
            {
                if (!row.Cell(1).TryGetValue<int>(out int quantity))
                    throw new ArgumentException($"Ungültige Menge in Zeile {row.RowNumber()}");

                string unit = row.Cell(2).GetString();
                string description = row.Cell(3).GetString();

                if (!row.Cell(4).TryGetValue<double>(out double price))
                {
                    if (!double.TryParse(row.Cell(4).GetString(), out price))
                        throw new ArgumentException($"Ungültiger Preis in Zeile {row.RowNumber()}");
                }

                positions.Add(new InvoicePosition
                {
                    Quantity = quantity,
                    Unit = unit,
                    Description = description,
                    Price = price
                });
            }

            return positions;
        }

        private string? GetDetails(IXLWorksheet wsPos)
        {
            return wsPos.Cell(2, 5).GetString();
        }
    }
}

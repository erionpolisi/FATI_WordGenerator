using ClosedXML.Excel;
using FATI_WordGenerator.Domain;
using System;
using System.Collections.Generic;
using System.Globalization;
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
                    throw new ArgumentNullException("\nEs wurden keine Daten gefunden. Bitte füllen Sie die Excel Datei aus.");
                }

                if (invoice.Positions == null || invoice.Positions.Count == 0)
                    throw new ArgumentException("\nEs wurden keine Positionen gefunden. Bitte füllen Sie die Positionen Tabelle in der Excel Datei aus.");

                return invoice;

            }
            catch (Exception ex)
            {
                throw new Exception($"\nFehler beim Lesen der Excel Datei: {ex.Message}", ex);
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
                if (IsEmptyPositionRow(row))
                    continue;

                if (!TryParseDecimal(row.Cell(1), out double quantity))
                    return positions;

                string unit = row.Cell(2).GetString();
                string description = row.Cell(3).GetString();

                if (!TryParseDecimal(row.Cell(4), out double price))
                    throw new ArgumentException($"Ungültiger Preis in Zeile {row.RowNumber()}");

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

        private static bool IsEmptyPositionRow(IXLRow row)
        {
            return row.Cells(1, 4).All(cell => string.IsNullOrWhiteSpace(cell.GetString()));
        }

        private static bool TryParseDecimal(IXLCell cell, out double value)
        {
            if (cell.TryGetValue<double>(out value))
                return true;

            var rawValue = cell.GetString().Trim();
            if (string.IsNullOrWhiteSpace(rawValue))
            {
                value = 0;
                return false;
            }

            return double.TryParse(rawValue, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.GetCultureInfo("de-DE"), out value)
                || double.TryParse(rawValue, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out value);
        }

        private string? GetDetails(IXLWorksheet wsPos)
        {
            return wsPos.Cell(2, 5).GetString();
        }
    }
}

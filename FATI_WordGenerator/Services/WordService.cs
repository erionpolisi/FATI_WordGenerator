using ClosedXML.Excel;
using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using DocumentFormat.OpenXml.Packaging;
using FATI_WordGenerator.Domain;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using W = DocumentFormat.OpenXml.Wordprocessing;

namespace FATI_WordGenerator.Services
{
    public class WordService
    {
        public void Generate(string outputPath, Invoice invoice, int increment) 
        {
            try
            {
                using var doc = WordprocessingDocument.Open(outputPath, true);

                var body = doc.MainDocumentPart?.Document?.Body
                 ?? throw new NullReferenceException("Word Dokument hat keinen Body");

                ReplaceText(body, invoice, increment);

                doc.MainDocumentPart.Document.Save();
            }
            catch (Exception ex)
            {
                throw new Exception("Fehler beim Generieren des Word Dokuments.", ex);
            }

        }

        private static W.Run CreateRun(string text)
        {
            return new W.Run(
                new W.RunProperties(
                    new W.RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }
                ),
                new W.Text(text)
            );
        }

        static void InsertTableAtPlaceholder(W.Body body, string placeholder, W.Table table)
        {
            var textElement = body.Descendants<W.Text>()
                .FirstOrDefault(t => t.Text.Contains(placeholder));

            if (textElement == null)
                throw new NullReferenceException($"Placeholder {placeholder} nicht im Word Dokument gefunden.");

            var parent = textElement.Parent
                ?? throw new NullReferenceException("TextElement hat kein Parent.");

            var grandParent = parent.Parent
                ?? throw new NullReferenceException("Parent hat kein Parent.");

            textElement.Text = textElement.Text.Replace(placeholder, "");

            grandParent.InsertAfter(table, parent);
        }

        private void InsertTableDetails(W.Body body, string? details)
        {
            if (string.IsNullOrWhiteSpace(details))
                return;

            var paragraph = new W.Paragraph(
                new W.ParagraphProperties(
                    new W.Justification() { Val = W.JustificationValues.Center }
                ),
                new W.Run(
                    new W.RunProperties(
                        new W.Bold(),
                        new W.RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }
                    ),
                    new W.Text(details)
                )
            );

            body.InsertBefore(
                paragraph,
                body.Descendants<W.Paragraph>()
                    .First(p => p.InnerText.Contains("{TABELLE}"))
            );
        }

        static void InsertWarningText(W.Body body)
        {
            var paragraph = new W.Paragraph(
                new W.Run(
                    new W.RunProperties(
                        new W.Bold(),
                        new W.RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }
                    ),
                    new W.Text("ACHTUNG: Wir möchten auf unsere neue Bankverbindung hinweisen!!!")
                )
            );

            var lastParagraph = body.Elements<W.Paragraph>().Last();
            body.InsertBefore(paragraph, lastParagraph);
        }

        private void CreateTable(W.Body body, Invoice invoice)
        {
            W.Table table = new W.Table(
                new W.TableProperties(

                    new W.TableJustification()
                    {
                       Val = W.TableRowAlignmentValues.Center
                    },

                    new W.TableWidth()
                    {
                        Width = "5000",
                        Type = W.TableWidthUnitValues.Pct
                    },

                    new W.TableBorders(
                        new W.TopBorder { Val = W.BorderValues.Nil },
                        new W.BottomBorder { Val = W.BorderValues.Nil },
                        new W.LeftBorder { Val = W.BorderValues.Nil },
                        new W.RightBorder { Val = W.BorderValues.Nil },
                        new W.InsideHorizontalBorder { Val = W.BorderValues.Nil },
                        new W.InsideVerticalBorder { Val = W.BorderValues.Nil }
                    )
                )
            );

            var header = new W.TableRow(

                CreateHeaderCell("Position"),
                CreateHeaderCell("Menge"),
                CreateHeaderCell("Bezeichnung"),
                CreateHeaderCell("Einzelpreis"),
                CreateHeaderCell("Gesamtpreis")
            );

            table.Append(header);

            int pos = 1;

            foreach (var p in invoice.Positions)
            {
                var tr = new W.TableRow(

                    new W.TableCell(new W.Paragraph(CreateRun($"{pos}"))),

                    new W.TableCell(
                        new W.Paragraph(CreateRun($"{p.Quantity} {p.Unit}"))
                    ),

                    new W.TableCell(
                        new W.Paragraph(CreateRun(p.Description))
                    ),

                    new W.TableCell(
                        new W.Paragraph(CreateRun($"{p.Price:N2} €"))
                    ),

                    new W.TableCell(
                        new W.Paragraph(CreateRun($"{p.Total:N2} €"))
                    )
                );

                table.Append(tr);
                pos++;
            }

            InsertTableAtPlaceholder(body, "{TABELLE}", table);
        }

        private static W.TableCell CreateHeaderCell(string text)
        {
            return new W.TableCell(
                new W.TableCellProperties(
                    new W.TableCellBorders(
                        new W.BottomBorder { Val = W.BorderValues.Single, Size = 8 }
                    )
                ),
                new W.Paragraph(
                    new W.Run(
                        new W.RunProperties(
                            new W.Bold(),
                            new W.RunFonts { Ascii = "Arial", HighAnsi = "Arial" }
                        ),
                        new W.Text(text)
                    )
                )
            );
        }

        private void ReplaceText(W.Body body, Invoice invoice, int increment)
        {
            InsertTableDetails(body, invoice.Details);

            string date = DateTime.Now.ToString("dd. MMMM yyyy");
            string year = DateTime.Now.Year.ToString();

            ReplaceText(body, "{DATUM}", date);
            ReplaceText(body, "{RECHNUNGSNUMMER}", $"{increment:D3}/{year}");
            ReplaceText(body, "{ADRESSE}", invoice.Address);
            ReplaceText(body, "{ZEITRAUM}", invoice.Period);

            ReplaceText(body, "{FIRMA}", invoice.Company.Name);
            ReplaceText(body, "{FIRMA_ADRESSE}", invoice.Company.StreetAndNumber);
            ReplaceText(body, "{FIRMA_PLZ_ORT}", invoice.Company.PLZAndCity);
            ReplaceText(body, "{ATU}", invoice.Company.ATU);

            ReplaceText(body, "{SUMME}", $"{invoice.Sum:N2} €");
            ReplaceText(body, "{SKONTO}", $"{invoice.Discount:N2} €");
            ReplaceText(body, "{GESAMTSUMME}", $"{invoice.Total:N2} €");

            CreateTable(body, invoice);
            InsertWarningText(body);
        }

        static void ReplaceText(W.Body body, string placeholder, string newValue)
        {
            var texts = body.Descendants<W.Text>();

            foreach (var text in texts)
            {
                if (text.Text.Contains(placeholder))
                {
                    text.Text = text.Text.Replace(placeholder, newValue);
                }
            }

            var paragraphs = body.Descendants<W.Paragraph>();

            foreach (var p in paragraphs)
            {
                string fullText = string.Concat(p.Descendants<W.Text>().Select(t => t.Text));

                if (fullText.Contains(placeholder))
                {
                    fullText = fullText.Replace(placeholder, newValue);

                    p.RemoveAllChildren<W.Run>();
                    p.AppendChild(new W.Run(new W.Text(fullText)));
                }
            }
        }
    }
}

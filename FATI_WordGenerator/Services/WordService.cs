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
                    new W.RunFonts() { Ascii = "Arial", HighAnsi = "Arial" },
                    new W.FontSize() { Val = "22" },
                    new W.FontSizeComplexScript() { Val = "22" }
                ),
                new W.Text(text)
            );
        }

        static void InsertTableAtPlaceholder(W.Body body, string placeholder, W.Table table)
        {
            var placeholderParagraph = body.Descendants<W.Paragraph>()
                .FirstOrDefault(p => p.InnerText.Contains(placeholder));

            if (placeholderParagraph == null)
                throw new NullReferenceException($"Placeholder {placeholder} nicht im Word Dokument gefunden.");

            foreach (var text in placeholderParagraph.Descendants<W.Text>())
            {
                if (text.Text.Contains(placeholder))
                    text.Text = text.Text.Replace(placeholder, "");
            }

            body.InsertAfter(table, placeholderParagraph);
            //RemoveFollowingEmptyParagraphs(placeholderParagraph);
            placeholderParagraph.Remove();
        }

        private static void RemoveFollowingEmptyParagraphs(W.Paragraph startParagraph)
        {
            var next = startParagraph.NextSibling<W.Paragraph>();

            while (next != null && string.IsNullOrWhiteSpace(next.InnerText))
            {
                var paragraphToRemove = next;
                next = next.NextSibling<W.Paragraph>();
                paragraphToRemove.Remove();
            }
        }

        private void InsertTableDetails(W.Body body, string? details)
        {
            if (string.IsNullOrWhiteSpace(details))
                return;

            var paragraph = new W.Paragraph(
                new W.ParagraphProperties(
                    new W.Justification() { Val = W.JustificationValues.Center },
                    new W.SpacingBetweenLines() { Before = "0", After = "0", Line = "240", LineRule = W.LineSpacingRuleValues.Auto }
                ),
                new W.Run(
                    new W.RunProperties(
                        new W.Bold(),
                        new W.RunFonts() { Ascii = "Arial", HighAnsi = "Arial" },
                        new W.FontSize() { Val = "22" },
                        new W.FontSizeComplexScript() { Val = "22" }
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
                new W.ParagraphProperties(
                    new W.SpacingBetweenLines() { Before = "0", After = "0", Line = "240", LineRule = W.LineSpacingRuleValues.Auto }
                ),
                new W.Run(
                    new W.RunProperties(
                        new W.Bold(),
                        new W.RunFonts() { Ascii = "Arial", HighAnsi = "Arial" },
                        new W.FontSize() { Val = "22" },
                        new W.FontSizeComplexScript() { Val = "22" }
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
                    new W.TableCellMarginDefault(
                        new W.TopMargin { Width = "0", Type = W.TableWidthUnitValues.Dxa },
                        new W.BottomMargin { Width = "0", Type = W.TableWidthUnitValues.Dxa },
                        new W.TableCellLeftMargin { Width = 40, Type = W.TableWidthValues.Dxa },
                        new W.TableCellRightMargin { Width = 40, Type = W.TableWidthValues.Dxa }
                    ),
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
                    new W.TableCell(CreateCompactParagraph($"{pos}")),

                    new W.TableCell(
                        CreateCompactParagraph($"{p.Quantity:0.00} {p.Unit}")
                    ),

                    new W.TableCell(
                        CreateCompactParagraph(p.Description)
                    ),

                    new W.TableCell(
                        CreateRightAlignedParagraph($"{p.Price:N2} €")
                    ),

                    new W.TableCell(
                        CreateRightAlignedParagraph($"{p.Total:N2} €")
                    )
                );

                table.Append(tr);
                pos++;
            }

            InsertTableAtPlaceholder(body, "{TABELLE}", table);
        }

        private static W.Paragraph CreateRightAlignedParagraph(string text)
        {
            return new W.Paragraph(
                new W.ParagraphProperties(
                    new W.Justification() { Val = W.JustificationValues.Right },
                    new W.SpacingBetweenLines()
                    {
                        Before = "0",
                        After = "0",
                        Line = "240",
                        LineRule = W.LineSpacingRuleValues.Auto
                    }
                ),
                CreateRun(text)
            );
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
                    new W.ParagraphProperties(
                        new W.SpacingBetweenLines() { Before = "0", After = "0", Line = "240", LineRule = W.LineSpacingRuleValues.Auto }
                    ),
                    new W.Run(
                        new W.RunProperties(
                            new W.Bold(),
                            new W.RunFonts { Ascii = "Arial", HighAnsi = "Arial" },
                            new W.FontSize() { Val = "22" },
                            new W.FontSizeComplexScript() { Val = "22" }
                        ),
                        new W.Text(text)
                    )
                )
            );
        }

        private static W.Paragraph CreateCompactParagraph(string text)
        {
            return new W.Paragraph(
                new W.ParagraphProperties(
                    new W.SpacingBetweenLines() { Before = "0", After = "0", Line = "240", LineRule = W.LineSpacingRuleValues.Auto }
                ),
                CreateRun(text)
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
            CompactSectionLayout(body);
            InsertWarningText(body);
        }

        private static void CompactSectionLayout(W.Body body)
        {
            var sectionProperties = body.GetFirstChild<W.SectionProperties>();
            var pageMargin = sectionProperties?.GetFirstChild<W.PageMargin>();

            if (pageMargin == null)
                return;

            pageMargin.Top = 1000;
            pageMargin.Bottom = 850;
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

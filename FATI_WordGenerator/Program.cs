using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Reflection.Metadata;
using W = DocumentFormat.OpenXml.Wordprocessing;

class Program
{
    static void Main()
    {
        Console.ResetColor();
        try
        {
            string resourcesPath;
#if DEBUG
            resourcesPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"..\..\..\Resources");
#else
            resourcesPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Resources");
#endif
            string currentDir = AppDomain.CurrentDomain.BaseDirectory;

            string jahr = DateTime.Now.Year.ToString();

            var folderPath = CreateFolder(currentDir);

            string inputPath = Path.Combine(resourcesPath, "Input.xlsx");

            using var workbook = new XLWorkbook(inputPath);
            Console.WriteLine("Input.xlsx geöffnet...");

            var wsAdresse = workbook.Worksheet("Adresse");
            var wsPos = workbook.Worksheet("Positionen");
            var wsFirma = workbook.Worksheet("Firma");

            string strasse = wsAdresse.Cell(2, 1).GetString();
            string nr = wsAdresse.Cell(2, 2).GetString();
            string bezirk = wsAdresse.Cell(2, 3).GetString();
            string ort = wsAdresse.Cell(2, 4).GetString();
            string zeitraumMonateRaw = wsAdresse.Cell(2, 5).GetString();
            string zeitraumJahr = wsAdresse.Cell(2, 6).GetString();

            string adresse = $"BVH: {strasse} {nr}, {bezirk} {ort}";
            var zeitraum = GetZeitraum(zeitraumMonateRaw, zeitraumJahr);

            string firma = wsFirma.Cell(2, 1).GetString();
            string firmaStrasse = wsFirma.Cell(2, 2).GetString();
            string firmaNr = wsFirma.Cell(2, 3).GetString();
            string firmaPLZ = wsFirma.Cell(2, 4).GetString();
            string firmaOrt = wsFirma.Cell(2, 5).GetString();
            string atu = wsFirma.Cell(2, 6).GetString();

            string firmaAdresse = $"{firmaStrasse} {firmaNr}";
            string firmaPLZOrt = $"{firmaPLZ} {firmaOrt}";

            int inkrement = Directory
                .GetFiles(folderPath, $"*_{jahr}_{firma}_*.docx")
                .Length + 1;

            string filename = $"{inkrement}_{jahr}_{firma}_{strasse} {nr}.docx";
            string outputPath = Path.Combine(folderPath, filename);

            string templatePath = Path.Combine(resourcesPath, "RechnungTemplate.docx");
            File.Copy(templatePath, outputPath, true);

            using var doc = WordprocessingDocument.Open(outputPath, true);
            Console.WriteLine("Neue Rechnung geöffnet...");

            var body = doc.MainDocumentPart.Document.Body;

            ReplaceText(body, "{DATUM}", DateTime.Now.ToString("dd. MMMM yyyy"));
            ReplaceText(body, "{RECHNUNGSNUMMER}", $"{inkrement:D3}/{jahr}");
            ReplaceText(body, "{ADRESSE}", adresse);
            ReplaceText(body, "{ZEITRAUM}", zeitraum);

            ReplaceText(body, "{FIRMA}", firma);
            ReplaceText(body, "{FIRMA_ADRESSE}", firmaAdresse);
            ReplaceText(body, "{FIRMA_PLZ_ORT}", firmaPLZOrt);
            ReplaceText(body, "{ATU}", atu);

            var summe = CreateTableAndGetSum(wsPos, body);

            double skonto = summe * 0.03;
            double gesamtsumme = summe - skonto;

            ReplaceText(body, "{SUMME}", $"{summe:N2} €");
            ReplaceText(body, "{SKONTO}", $"{skonto:N2} €");
            ReplaceText(body, "{GESAMTSUMME}", $"{gesamtsumme:N2} €");

            InsertWarningText(body);

            doc.MainDocumentPart.Document.Save();

            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine($"Rechnung erstellt: {filename}");
            Console.ResetColor();

            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
            {
                FileName = outputPath,
                UseShellExecute = true
            });
        }
        catch (ArgumentException e)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine(e.Message);
            Console.ResetColor();
        }
        catch (IOException e)
        {
            Console.ForegroundColor = ConsoleColor.Yellow;
            if (e.Message.Contains("Input.xlsx"))
            {
                Console.WriteLine("Bitte Excel schließen");
            }
            else if (e.Message.Contains("RechnungTemplate.docx"))
            {
                Console.WriteLine("Bitte Word schließen");
            }
            else
            {
                Console.WriteLine(e.Message);
            }
            Console.ResetColor();
        }
        catch (Exception ex)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine(ex.Message);
            Console.ResetColor();
        }
    }

    private static double CreateTableAndGetSum(IXLWorksheet wsPos, W.Body body)
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

     new W.TableCell(
         new W.TableCellProperties(
             new W.TableCellBorders(
                 new W.BottomBorder { Val = W.BorderValues.Single, Size = 8 }
             )
         ),
         new W.Paragraph(
             new W.Run(
                 new W.RunProperties(new W.Bold(), new W.RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }),
                 new W.Text("Position")
             )
         )
     ),

     new W.TableCell(
         new W.TableCellProperties(
             new W.TableCellBorders(
                 new W.BottomBorder { Val = W.BorderValues.Single, Size = 8 }
             )
         ),
         new W.Paragraph(
             new W.Run(
                 new W.RunProperties(new W.Bold(), new W.RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }),
                 new W.Text("Menge")
             )
         )
     ),

     new W.TableCell(
         new W.TableCellProperties(
             new W.TableCellBorders(
                 new W.BottomBorder { Val = W.BorderValues.Single, Size = 8 }
             )
         ),
         new W.Paragraph(
             new W.Run(
                 new W.RunProperties(new W.Bold(), new W.RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }),
                 new W.Text("Bezeichnung")
             )
         )
     ),

     new W.TableCell(
         new W.TableCellProperties(
             new W.TableCellBorders(
                 new W.BottomBorder { Val = W.BorderValues.Single, Size = 8 }
             )
         ),
         new W.Paragraph(
             new W.Run(
                 new W.RunProperties(new W.Bold(), new W.RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }),
                 new W.Text("Einzelpreis")
             )
         )
     ),

     new W.TableCell(
         new W.TableCellProperties(
             new W.TableCellBorders(
                 new W.BottomBorder { Val = W.BorderValues.Single, Size = 8 }
             )
         ),
         new W.Paragraph(
             new W.Run(
                 new W.RunProperties(new W.Bold(), new W.RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }),
                 new W.Text("Gesamtpreis")
             )
         )
     )
 );

        table.Append(header);

        double summe = 0;
        int pos = 1;
        string tableTitle = string.Empty;

        foreach (var row in wsPos.RowsUsed().Skip(1))
        {
            int menge = row.Cell(1).GetValue<int>();
            string mengenbez = row.Cell(2).GetString();
            string bez = row.Cell(3).GetString();
            double preis = row.Cell(4).GetDouble();

            if (string.IsNullOrWhiteSpace(tableTitle)) 
                tableTitle = row.Cell(5).GetString();


            double gesamt = menge * preis;
            summe += gesamt;

            var tr = new W.TableRow(

                new W.TableCell(new W.Paragraph(CreateRun($"{pos}"))),

                new W.TableCell(new W.Paragraph(CreateRun($"{menge} {mengenbez}"))),
                new W.TableCell(new W.Paragraph(CreateRun(bez))),
                new W.TableCell(new W.Paragraph(CreateRun($"{preis:N2} €"))),
                new W.TableCell(new W.Paragraph(CreateRun($"{gesamt:N2} €")))
            );

            table.Append(tr);
            pos++;
        }

        if (!string.IsNullOrWhiteSpace(tableTitle))
        {
            var titleParagraph = new W.Paragraph(
                new W.ParagraphProperties(
                    new W.Justification() { Val = W.JustificationValues.Center }
                ),
                new W.Run(
                    new W.RunProperties(
                        new W.Bold(),
                        new W.RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }
                    ),
                    new W.Text(tableTitle)
                )
            );

            body.InsertBefore(
                titleParagraph,
                body.Descendants<W.Paragraph>()
                    .First(p => p.InnerText.Contains("{TABELLE}"))
            );
        }

        InsertTableAtPlaceholder(body, "{TABELLE}", table);
        return summe;
    }

    private static string GetZeitraum(string zeitraumMonateRaw, string zeitraumJahr)
    {
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

    private static string CreateFolder(string currentDir)
    {
        string folderPath = Path.Combine(currentDir, "Rechnungen");
        Directory.CreateDirectory(folderPath);
        return folderPath;
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

        // Fix für gesplittete Runs
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

    static void InsertTableAtPlaceholder(W.Body body, string placeholder, W.Table table)
    {
        var textElement = body.Descendants<W.Text>()
            .FirstOrDefault(t => t.Text.Contains(placeholder));

        if (textElement != null)
        {
            var parent = textElement.Parent;
            textElement.Text = textElement.Text.Replace(placeholder, "");
            parent.Parent.InsertAfter(table, parent);
        }
    }

    public enum Months
    {
        Unbekannt = 0,
        Januar = 1,
        Februar = 2,
        März = 3,
        April = 4,
        Mai = 5,
        Juni = 6,
        Juli = 7,
        August = 8,
        September = 9,
        Oktober = 10,
        November = 11,
        Dezember = 12
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
}
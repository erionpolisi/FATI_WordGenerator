using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using W = DocumentFormat.OpenXml.Wordprocessing;

class Program
{
    static void Main()
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
        int inkrement = Directory.GetFiles(folderPath, "*.docx").Length + 1;

        string inputPath = Path.Combine(resourcesPath, "Input.xlsx");

        using var workbook = new XLWorkbook(inputPath);
        Console.WriteLine("Input_Excel entered...");

        var wsAdresse = workbook.Worksheet("Adresse");
        var wsPos = workbook.Worksheet("Positionen");
        var wsFirma = workbook.Worksheet("Firma");

        string strasse = wsAdresse.Cell(2, 1).GetString();
        string nr = wsAdresse.Cell(2, 2).GetString();
        string bezirk = wsAdresse.Cell(2, 3).GetString();
        string ort = wsAdresse.Cell(2, 4).GetString();
        string adresse = $"{strasse} {nr}, {bezirk} {ort}";

        string firma = wsFirma.Cell(2, 1).GetString();
        string firmaStrasse = wsFirma.Cell(2, 2).GetString();
        string firmaNr = wsFirma.Cell(2, 3).GetString();
        string firmaPLZ = wsFirma.Cell(2, 4).GetString();
        string firmaOrt = wsFirma.Cell(2, 5).GetString();
        string atu = wsFirma.Cell(2, 6).GetString();

        string firmaAdresse = $"{firmaStrasse} {firmaNr}";
        string firmaPLZOrt = $"{firmaPLZ} {firmaOrt}";

        string outputPath = Path.Combine(folderPath, $"{inkrement}_{jahr}_ALLITECH_{strasse}.docx");

        string templatePath = Path.Combine(resourcesPath, "RechnungTemplate.docx");
        File.Copy(templatePath, outputPath, true);

        using var doc = WordprocessingDocument.Open(outputPath, true);

        var body = doc.MainDocumentPart.Document.Body;

        ReplaceText(body, "{DATUM}", DateTime.Now.ToString("dd. MMMM yyyy"));
        ReplaceText(body, "{RECHNUNGSNUMMER}", $"{inkrement:D3}/{jahr}");
        ReplaceText(body, "{ADRESSE}", adresse);
        ReplaceText(body, "{ZEITRAUM}", DateTime.Now.ToString("MMMM"));

        W.Table table = new W.Table();

        double summe = 0;
        int pos = 1;

        foreach (var row in wsPos.RowsUsed().Skip(1))
        {
            int menge = row.Cell(1).GetValue<int>();
            string mengenbez = row.Cell(2).GetString();
            string bez = row.Cell(3).GetString();
            double preis = row.Cell(4).GetDouble();

            double gesamt = menge * preis;
            summe += gesamt;

            var tr = new W.TableRow(
                new W.TableCell(new W.Paragraph(new W.Run(new W.Text($"Pos.{pos}")))),
                new W.TableCell(new W.Paragraph(new W.Run(new W.Text($"{menge} {mengenbez}")))),
                new W.TableCell(new W.Paragraph(new W.Run(new W.Text(bez)))),
                new W.TableCell(new W.Paragraph(new W.Run(new W.Text($"{preis:0.00} €")))),
                new W.TableCell(new W.Paragraph(new W.Run(new W.Text($"{gesamt:0.00} €"))))
            );

            table.Append(tr);
            pos++;
        }

        InsertTableAtPlaceholder(body, "{TABELLE}", table);

        double skonto = summe * 0.03;
        double gesamtsumme = summe - skonto;

        ReplaceText(body, "{SUMME}", $"{summe:0.00} €");
        ReplaceText(body, "{SKONTO}", $"{skonto:0.00} €");
        ReplaceText(body, "{GESAMTSUMME}", $"{gesamtsumme:0.00} €");

        InsertWarningText(body);

        doc.MainDocumentPart.Document.Save();

        Console.WriteLine($"Rechnung erstellt: {outputPath}");
    }

    private static string CreateFolder(string currentDir)
    {
        string folderPath = Path.Combine(currentDir, "Rechnungen");
        Directory.CreateDirectory(folderPath);
        return folderPath;
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
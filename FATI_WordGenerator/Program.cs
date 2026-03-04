using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using W = DocumentFormat.OpenXml.Wordprocessing;

class Program
{
    static void Main()
    {
        string resourcesPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"..\..\..\Resources");
        string currentDir = AppDomain.CurrentDomain.BaseDirectory;

        string jahr = DateTime.Now.Year.ToString();

        var folderPath = CreateFolder(currentDir);
        int inkrement = Directory.GetFiles(folderPath, "*.docx").Length + 1;

        string inputPath = Path.Combine(resourcesPath, "Input.xlsx");

        using var workbook = new XLWorkbook(inputPath);
        Console.WriteLine("Input_Excel entered...");

        var wsAdresse = workbook.Worksheet("Adresse");
        var wsPos = workbook.Worksheet("Positionen");

        string strasse = wsAdresse.Cell(2, 1).GetString();
        string nr = wsAdresse.Cell(2, 2).GetString();
        string bezirk = wsAdresse.Cell(2, 3).GetString();
        string ort = wsAdresse.Cell(2, 4).GetString();

        string adresse = $"{strasse} {nr}, {bezirk} {ort}";

        string outputPath = Path.Combine(folderPath, $"{inkrement}_{jahr}_ALLITECH_{strasse}.docx");

        string templatePath = Path.Combine(resourcesPath, "RechnungTemplate.docx");
        File.Copy(templatePath, outputPath, true);

        using var doc = WordprocessingDocument.Open(outputPath, true);

        var body = doc.MainDocumentPart.Document.Body;

        ReplaceText(body, "{DATUM}", DateTime.Now.ToString("dd. MMMM yyyy"));
        ReplaceText(body, "{JAHR}", jahr);
        ReplaceText(body, "{INKREMENT}", inkrement.ToString("D3"));
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
        foreach (var text in body.Descendants<W.Text>())
        {
            if (text.Text.Contains(placeholder))
            {
                text.Text = text.Text.Replace(placeholder, newValue);
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
}
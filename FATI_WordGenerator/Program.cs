using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using W = DocumentFormat.OpenXml.Wordprocessing;

class Program
{
    public class ExcelData
    {
        public string Strasse { get; set; }
        public string Nr { get; set; }
        public string Bezirk { get; set; }
        public string Ort { get; set; }

        public string Adresse => $"{Strasse} {Nr}, {Bezirk} {Ort}";
    }

    public class DocData
    {
        // später: Tabelle, Summe, Skonto, Gesamtsumme, etc.
    }

    static void Main()
    {

        string resourcesPath = "C:\\Users\\polise\\source\\repos\\FATI_WordGenerator\\FATI_WordGenerator\\Resources";
        string currentDir = AppDomain.CurrentDomain.BaseDirectory;
        string jahr = DateTime.Now.Year.ToString();


        var folderPath = CreateFolder(currentDir);
        int inkrement = Directory.GetFiles(folderPath, "*.docx").Length + 1;
        
        using var doc = CreateDoc(folderPath, resourcesPath, jahr, inkrement);

        
        var body = doc.MainDocumentPart.Document.Body;

        // Platzhalter ersetzen
        ReplaceText(body, "{DATUM}", DateTime.Now.ToString("dd. MMMM yyyy"));
        ReplaceText(body, "{JAHR}", jahr);
        ReplaceText(body, "{INKREMENT}", inkrement.ToString("D3"));
        ReplaceText(body, "{ADRESSE}", adresse);
        ReplaceText(body, "{ZEITRAUM}", DateTime.Now.ToString("MMMM"));

        // Tabelle erstellen
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
                new W.TableCell(new W.Paragraph(new W.Run(new W.Text($"{menge:0} {mengenbez}")))),
                new W.TableCell(new W.Paragraph(new W.Run(new W.Text(bez)))),
                new W.TableCell(new W.Paragraph(new W.Run(new W.Text($"{preis:0.00} €")))),
                new W.TableCell(new W.Paragraph(new W.Run(new W.Text($"{gesamt:0.00} €"))))
            );

            table.Append(tr);
            pos++;
        }

        // Tabelle an Platzhalter einsetzen
        InsertTableAtPlaceholder(body, "{TABELLE}", table);

        // Berechnungen
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

    private static WordprocessingDocument CreateDoc(string folderPath, string resourcesPath, string jahr, int inkrement)
    {
        var inputExcel = CreateExcel(folderPath, resourcesPath, jahr, inkrement);
        GetOutPutPath(inputExcel);

        string templatePath = Path.Combine(resourcesPath, "RechnungTemplate.docx");
        File.Copy(templatePath, outputPath, true);

        return WordprocessingDocument.Open(outputPath, true);
    }

    private static string GetOutPutPath()
    {
        return Path.Combine(folderPath, $"{inkrement}_{jahr}_ALLITECH_{strasse}.docx");
    }

    private static string CreateExcel(string folderPath, string resourcesPath, string jahr, int inkrement)
    {
        var inputExcel = new
        string inputPath = Path.Combine(resourcesPath, "Input.xlsx");

        using var workbook = new XLWorkbook(inputPath);
        var wsAdresse = workbook.Worksheet("Adresse");
        var wsPos = workbook.Worksheet("Positionen");

        // Adresse zusammenbauen
        string strasse = wsAdresse.Cell(2, 1).GetString();
        string nr = wsAdresse.Cell(2, 2).GetString();
        string bezirk = wsAdresse.Cell(2, 3).GetString();
        string ort = wsAdresse.Cell(2, 4).GetString();
        string adresse = $"{strasse} {nr}, {bezirk} {ort}";

        return outputPath = Path.Combine(folderPath, $"{inkrement}_{jahr}_ALLITECH_{strasse}.docx");
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
        var textElement = body.Descendants<W.Text>().FirstOrDefault(t => t.Text.Contains(placeholder));
        if (textElement != null)
        {
            var parent = textElement.Parent;
            textElement.Text = textElement.Text.Replace(placeholder, "");
            parent.Parent.InsertAfter(table, parent);
        }
    }
}

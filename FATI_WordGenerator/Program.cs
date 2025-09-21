using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

class Program
{
    static void Main()
    {
        string templatePath = "RechnungTemplate.docx";
        string outputPath = "Rechnung.docx";

        // Kopie der Vorlage erstellen
        File.Copy(templatePath, outputPath, true);

        using (WordprocessingDocument doc = WordprocessingDocument.Open(outputPath, true))
        {
            var body = doc.MainDocumentPart.Document.Body;

            // 1. Adresse ersetzen
            ReplaceText(body, "{ADRESSE}", "BVH Auhofstraße 196, 1130 Wien");

            // 2. Datum ersetzen
            ReplaceText(body, "{DATUM}", DateTime.Now.ToString("dd. MMMM yyyy"));

            // 3. Tabelle suchen und Positionen einfügen
            Table table = body.Elements<Table>().First();

            // Beispiel-Positionen
            var items = new List<(double Menge, string Bez, double Preis)>
            {
                (18, "Regie FA", 50),
                (18, "Regie HA", 45)
            };

            double summe = 0;
            foreach (var item in items)
            {
                double gesamt = item.Menge * item.Preis;
                summe += gesamt;

                TableRow row = new TableRow(
                    new TableCell(new Paragraph(new Run(new Text($"Pos.{items.IndexOf(item) + 1}")))),
                    new TableCell(new Paragraph(new Run(new Text(item.Menge.ToString("0.00"))))),
                    new TableCell(new Paragraph(new Run(new Text(item.Bez)))),
                    new TableCell(new Paragraph(new Run(new Text($"{item.Preis:0.00} €")))),
                    new TableCell(new Paragraph(new Run(new Text($"{gesamt:0.00} €"))))
                );
                table.Append(row);
            }

            // 4. Gesamtsumme einsetzen
            ReplaceText(body, "{SUMME}", $"{summe:0.00} €");
            ReplaceText(body, "{GESAMTSUMME}", $"{summe * 0.97:0.00} €"); // z.B. mit Skonto
        }
    }

    // Hilfsfunktion für Text-Ersetzung
    static void ReplaceText(Body body, string placeholder, string newValue)
    {
        foreach (var text in body.Descendants<Text>())
        {
            if (text.Text.Contains(placeholder))
            {
                text.Text = text.Text.Replace(placeholder, newValue);
            }
        }
    }
}

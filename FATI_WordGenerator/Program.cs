using FATI_WordGenerator.Services;
using System.ComponentModel;
using System.Diagnostics;

class Program
{
    static void Main()
    {
        bool errorOccurred = false;

        try
        {
            var invoiceService = new InvoiceService();

            string outputPath = invoiceService.Generate();

            WriteColored($"Rechnung erstellt: {invoiceService.FileName}", ConsoleColor.Green);
            WriteColored($"\nDateipfad: {outputPath}\n", ConsoleColor.Cyan);

            Process.Start(new ProcessStartInfo(outputPath) { UseShellExecute = true });
        }
        catch (Exception ex)
        {
            errorOccurred = true;

            switch (ex)
            {
                case IOException ioEx when ioEx.Message.Contains("Input.xlsx"):
                    WriteColored("Bitte Excel schließen", ConsoleColor.Yellow);
                    break;

                case IOException ioEx when ioEx.Message.Contains("RechnungTemplate.docx"):
                    WriteColored("Bitte Word schließen", ConsoleColor.Yellow);
                    break;

                case IOException:
                    WriteColored(ex.Message, ConsoleColor.Red);
                    break;

                case ArgumentException:
                    WriteColored(ex.Message, ConsoleColor.Red);
                    break;

                case UnauthorizedAccessException:
                    WriteColored("Schreibrechte fehlen", ConsoleColor.Red);
                    break;

                case FormatException:
                    WriteColored($"Falsches Format: {ex.Source}", ConsoleColor.Red);
                    break;

                case NullReferenceException:
                    WriteColored($"{ex.Message}: {ex.Source}", ConsoleColor.Red);
                    break;

                case Win32Exception:
                    WriteColored("Datei kann nicht geöffnet werden", ConsoleColor.Red);
                    break;

                default:
                    WriteColored(ex.Message, ConsoleColor.Red);
                    break;
            }
        }
        finally
        {
            if (errorOccurred || Debugger.IsAttached)
            {
                Console.WriteLine("Beliebige Taste pressen um Programm zu schließen...");
                Console.ReadKey();
            }
        }
    }

    static void WriteColored(string text, ConsoleColor color)
    {
        Console.ForegroundColor = color;
        Console.WriteLine(text);
        Console.ResetColor();
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
}

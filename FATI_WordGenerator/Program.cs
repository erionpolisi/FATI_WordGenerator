using FATI_WordGenerator.Services;
using System.ComponentModel;
using System.Diagnostics;
using FATI_WordGenerator;

class Program
{
    static void Main()
    {
        var settings = Settings.LoadSettings();
        bool errorOccurred = false; //True when exception is thrown
        var sw = Stopwatch.StartNew(); //Start measuring time for loading bar
        ConsoleFormattingAndBanner();

        try
        {
            var invoiceService = new InvoiceService();

            int totalSteps = 4;// Excel lesen, RN berechnen, Template kopieren, Word generieren
            string outputPath = invoiceService.Generate((step, message) =>
            {
                DrawProgress(step, totalSteps, message);
            });
            sw.Stop();

            WriteColored($"Rechnung erstellt: {invoiceService.FileName}", ConsoleColor.Green);
            WriteColored($"\nDateipfad: {outputPath}", ConsoleColor.Cyan);
            
            if (settings.OpenWordDocument)
            {
                Console.WriteLine("\nNeue Rechnung wird geöffnet...");
                Process.Start(new ProcessStartInfo(outputPath) { UseShellExecute = true });
                WriteColored("Done", ConsoleColor.Green);
            }
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
            if (errorOccurred || Debugger.IsAttached || !settings.AutoCloseTerminal)
            {
                Console.WriteLine($"\nFertig in {sw.ElapsedMilliseconds} ms");
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

    static void DrawProgress(int step, int totalSteps, string message)
    {
        int width = 20;
        double percent = (double)step / totalSteps;
        int filled = (int)(width * percent);

        Console.CursorLeft = 0;
        Console.Write($"[{new string('#', filled).PadRight(width)}] {(int)(percent * 100),3}%  {message}");

        if (step == totalSteps)
            Console.WriteLine();
    }

    static void ConsoleFormattingAndBanner()
    {
        Console.CursorVisible = false;
        Console.SetWindowSize(120, 30);
        WriteColored(@"
███████╗ █████╗ ████████╗██╗    ███████╗████████╗███████╗██╗███╗   ██╗███╗   ███╗███████╗████████╗███████╗
██╔════╝██╔══██╗╚══██╔══╝██║    ██╔════╝╚══██╔══╝██╔════╝██║████╗  ██║████╗ ████║██╔════╝╚══██╔══╝╚══███╔╝
█████╗  ███████║   ██║   ██║    ███████╗   ██║   █████╗  ██║██╔██╗ ██║██╔████╔██║█████╗     ██║     ███╔╝ 
██╔══╝  ██╔══██║   ██║   ██║    ╚════██║   ██║   ██╔══╝  ██║██║╚██╗██║██║╚██╔╝██║██╔══╝     ██║    ███╔╝  
██║     ██║  ██║   ██║   ██║    ███████║   ██║   ███████╗██║██║ ╚████║██║ ╚═╝ ██║███████╗   ██║   ███████╗
╚═╝     ╚═╝  ╚═╝   ╚═╝   ╚═╝    ╚══════╝   ╚═╝   ╚══════╝╚═╝╚═╝  ╚═══╝╚═╝     ╚═╝╚══════╝   ╚═╝   ╚══════╝
", ConsoleColor.Cyan);
    }
}

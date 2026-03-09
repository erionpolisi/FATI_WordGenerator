using FATI_WordGenerator.Services;

class Program
{
    static void Main()
    {
        Console.ResetColor();

        try
        {
            var invoiceService = new InvoiceService();

            string outputPath = invoiceService.Generate();

            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine($"Rechnung erstellt: {outputPath}");
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
        finally
        {
            Console.WriteLine("Press to close Console");
            Console.ReadKey();
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
}
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Frostybee.EpsonEscPos
{
    public class Program
    {
        public static void Main(string[] args)
        {
            string comPortName = "COM1";
            var tester = new PrinterManager();

            // Example 1: Using default configuration (backward compatibility).
            Console.WriteLine("=== USING DEFAULT CONFIGURATION ===");
            IPrinterStatusResult result = tester.GetPrinterStatus(comPortName);
            DisplayResults(result);

            // Example 2: Using custom configuration.
            Console.WriteLine("\n=== USING CUSTOM CONFIGURATION ===");
            var customConfig = new PrinterConfiguration(
                baudRate: 115200,          // Higher speed.
                onlineTimeout: 5000,       // Longer timeout.
                offlineTimeout: 2000,      // Longer offline timeout.
                initializationSleepMs: 200 // Longer initialization delay.
            );
            result = tester.GetPrinterStatus(comPortName, customConfig);
            DisplayResults(result);

            // Example 3: Using predefined high-speed configuration.
            Console.WriteLine("\n=== USING HIGH-SPEED CONFIGURATION ===");
            var highSpeedConfig = PrinterConfiguration.CreateHighSpeed();
            result = tester.GetPrinterStatus(comPortName, highSpeedConfig);
            DisplayResults(result);

            // Example 4: Using predefined reliable configuration.
            Console.WriteLine("\n=== USING RELIABLE CONFIGURATION ===");
            var reliableConfig = PrinterConfiguration.CreateReliable();
            result = tester.GetPrinterStatus(comPortName, reliableConfig);
            DisplayResults(result);

            Console.WriteLine("\n" + new string('=', 60));
            Console.WriteLine("DETAILED STATUS REPORT (Default Config):");
            Console.WriteLine(new string('=', 60));
            Console.WriteLine(tester.GetStatusReport(comPortName));

            Console.WriteLine("\n" + new string('=', 60));
            Console.WriteLine("DETAILED STATUS REPORT (Custom Config):");
            Console.WriteLine(new string('=', 60));
            Console.WriteLine(tester.GetStatusReport(comPortName, customConfig));

            Console.WriteLine(new string('=', 60));
            Console.WriteLine("Press any key to exit...");
            Console.ReadKey();
        }

        private static void DisplayResults(IPrinterStatusResult result)
        {
            Console.WriteLine("\n" + new string('=', 60));
            Console.WriteLine("PRINTER STATUS RESULT:");
            Console.WriteLine(new string('=', 60));

            // Paper Status Code.
            Console.WriteLine($"Paper Status Code: {result.PaperStatus}");
            string statusDescription;
            switch (result.PaperStatus)
            {
                case 0:
                    statusDescription = "Paper OK";
                    break;
                case 1:
                    statusDescription = "Paper Near End";
                    break;
                case 2:
                    statusDescription = "Paper Out";
                    break;
                default:
                    statusDescription = "Unknown Status";
                    break;
            }
            Console.WriteLine($"Paper Status: {statusDescription}");

            Console.WriteLine();
            // Print Printer Status.
            Console.WriteLine($"Printer Status Code: {result.PrinterStatus}");
            string printerStatusDescription;
            switch (result.PrinterStatus)
            {
                case 0:
                    printerStatusDescription = "Offline";
                    break;
                case 1:
                    printerStatusDescription = "Online";
                    break;
                default:
                    printerStatusDescription = "Unknown";
                    break;
            }
            
            Console.WriteLine($"Printer Status: {printerStatusDescription}");
            Console.WriteLine();

            // Print Cover Status.
            Console.WriteLine($"Cover Status Code: {result.CoverStatus}");
            string coverStatusDescription;
            switch (result.CoverStatus)
            {
                case 0:
                    coverStatusDescription = "Open";
                    break;
                case 1:
                    coverStatusDescription = "Closed";
                    break;
                default:
                    coverStatusDescription = "Unknown";
                    break;
            }
            
            Console.WriteLine($"Cover Status: {coverStatusDescription}");

            // Print Error Information.
            Console.WriteLine($"Has Error: {result.HasError}");
            Console.WriteLine($"Error Status: {(result.HasError == 1 ? "ERROR" : "NO ERROR")}");
            
            if (result.HasError == 1 && !string.IsNullOrEmpty(result.ErrorMessage))
            {
                Console.WriteLine($"Error Message: {result.ErrorMessage}");
            }
            Console.WriteLine();
        }    
    }
}

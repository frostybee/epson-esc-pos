using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Frostybee.EpsonEscPos
{
    internal class Program
    {
        public static void Main(string[] args)
        {
            string comPortName = "COM1";


            Console.WriteLine($"Testing Epson Printer on {comPortName}");
            Console.WriteLine("This test will specifically check for paper status detection improvements.");
            Console.WriteLine(new string('=', 80));


            IPrinterStatusResult result;
            using (var tester = new EpsonPrinterManager())
            {
                result = tester.GetPrinterStatus(comPortName);
            }

            Console.WriteLine("\n" + new string('=', 60));
            Console.WriteLine("PRINTER STATUS RESULT:");
            Console.WriteLine(new string('=', 60));

            // Paper Status Code
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
            // Print Printer Status
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
            // Print Cover Status
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

            // Print Error Information
            Console.WriteLine($"Has Error: {result.HasError}");
            Console.WriteLine($"Error Status: {(result.HasError == 1 ? "ERROR" : "NO ERROR")}");

            if (result.HasError == 1 && !string.IsNullOrEmpty(result.ErrorMessage))
            {
                Console.WriteLine($"Error Message: {result.ErrorMessage}");
            }

            // Print Can Print Status
            Console.WriteLine($"Can Print Code: {result.CanPrint}");
            Console.WriteLine($"Can Print Status: {(result.CanPrint == 1 ? "Can Print" : "Cannot Print")}");

            Console.WriteLine();
            Console.WriteLine(new string('=', 60));
            Console.WriteLine("STATUS REPORT:");
            Console.WriteLine(new string('=', 60));

            // Use the same result object instead of making another call
            Console.WriteLine(FormatStatusReport(result, comPortName));

            // Test specific paper out scenario
            Console.WriteLine();
            Console.WriteLine(new string('=', 60));
            Console.WriteLine("PAPER STATUS TEST RESULTS:");
            Console.WriteLine(new string('=', 60));

            if (result.PaperStatus == 2)
            {
                Console.WriteLine("SUCCESS: Paper OUT condition detected correctly!");
                Console.WriteLine("This indicates the fix for paper status detection is working.");
            }
            else if (result.PaperStatus == 1)
            {
                Console.WriteLine("WARNING: Paper is near end - replace soon.");
                Console.WriteLine("Paper status detection is working.");
            }
            else if (result.PaperStatus == 0)
            {
                Console.WriteLine("Paper is OK.");
                Console.WriteLine("If you expect paper to be out, please check the physical printer.");
            }
            else
            {
                Console.WriteLine("Paper status is unknown.");
                Console.WriteLine("This may indicate a communication issue or unsupported printer model.");
            }

            Console.WriteLine();
            Console.WriteLine("TEST INSTRUCTIONS:");
            Console.WriteLine("1. Run this test with paper installed - should show 'Paper OK'");
            Console.WriteLine("2. Remove paper and run again - should show 'Paper OUT'");
            Console.WriteLine("3. Test with low paper - should show 'Paper Near End'");
            Console.WriteLine();
            Console.WriteLine("Press any key to exit...");

            Console.ReadKey();
        }

        private static string FormatStatusReport(IPrinterStatusResult status, string comPort)
        {
            var report = new StringBuilder();

            // Helper method to get paper status text
            string GetPaperStatusText(int paperStatus)
            {
                switch (paperStatus)
                {
                    case 0: return "PAPER OK";
                    case 1: return "PAPER NEAR END";
                    case 2: return "PAPER EMPTY";
                    case 99: return "UNKNOWN";
                    default: return "UNKNOWN";
                }
            }

            // Build report
            report.AppendLine("=== EPSON PRINTER STATUS REPORT ===");
            report.AppendLine($"COM Port: {comPort}, Time: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            report.AppendLine($"Status: {(status.HasError == 1 ? "ERROR" : "OK")} - {status.ErrorMessage}");
            report.AppendLine($"Printer: {(status.PrinterStatus == 1 ? "ONLINE" : "OFFLINE")}, Cover: {(status.CoverStatus == 1 ? "CLOSED" : "OPEN")}");
            report.AppendLine($"Paper: {GetPaperStatusText(status.PaperStatus)}");
            report.AppendLine($"CanPrint: {status.CanPrint} ({(status.CanPrint == 1 ? "Can Print" : "Cannot Print")})");
            report.AppendLine($"Ready to Print: {(status.CanPrint == 1 ? "YES" : "NO")}");

            // Add warnings if needed
            if (status.PaperStatus == 1) report.AppendLine("WARNING: Replace paper soon");
            if (status.PaperStatus == 2) report.AppendLine("CRITICAL: Out of paper");
            if (status.CoverStatus == 0) report.AppendLine("WARNING: Printer cover is open");

            return report.ToString();
        }
    }
}

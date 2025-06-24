# EPSON ESC/POS .NET Library

A .NET Framework library for real-time Epson ESC/POS printer status monitoring and management via serial communication.

## Features

- Real-time printer status monitoring via serial communication,
- Paper status detection with multiple states (OK, Near End, Out),
- Cover status monitoring (Open/Closed),
- Printer connectivity status tracking (Online/Offline),
- Comprehensive error reporting and diagnostics,
- COM-visible interface,
- Support for multiple development environments (C#, VBA, Python),
- Configurable communication parameters (baud rate, timeouts, etc.),
- Automatic printer initialization and command handling.

## Requirements

- .NET Framework 4.8
- Can be ported to .NET 8 and .NET 9
- Serial port connection to Epson ESC/POS compatible printer

**Tested Hardware:**

- Epson TM-T20III thermal receipt printer.
  
## Usage Examples

### C# Usage

```csharp
using Frostybee.EpsonEscPos;

// Basic status check.
using (var printerManager = new EpsonPrinterManager())
{
    IPrinterStatusResult result = printerManager.GetPrinterStatus("COM1");
    
    // Check printer status.
    bool isOnline = result.PrinterStatus == 1;
    bool canPrint = result.CanPrint == 1;
    
    // Check paper status.
    switch (result.PaperStatus)
    {
        case 0: Console.WriteLine("Paper OK"); break;
        case 1: Console.WriteLine("Paper Near End"); break;
        case 2: Console.WriteLine("Paper Out"); break;
        default: Console.WriteLine("Unknown Status"); break;
    }
    
    // Check for errors.
    if (result.HasError == 1)
    {
        Console.WriteLine($"Error: {result.ErrorMessage}");
    }
}

// Using custom configuration.
var config = new PrinterConfiguration
{
    BaudRate = 9600,
    ReadTimeout = 1000,
    WriteTimeout = 500
};

using (var printerManager = new EpsonPrinterManager())
{
    IPrinterStatusResult result = printerManager.GetPrinterStatus("COM1", config);
    string statusReport = printerManager.GetStatusReport("COM1", config);
    Console.WriteLine(statusReport);
}
```

### VBA Usage

```vb
Dim printerManager As Object
Dim result As Object

Set printerManager = CreateObject("Frostybee.EpsonEscPos.EpsonPrinterManager")
Set result = printerManager.GetPrinterStatus("COM1")

If result.PrinterStatus = 1 Then
    MsgBox "Printer is Online"
Else
    MsgBox "Printer is Offline"
End If

If result.PaperStatus = 2 Then
    MsgBox "Paper is Out!"
End If

printerManager.Dispose
```

### Example Output

The `GetStatusReport` method provides detailed formatted status information:

**Normal Operation:**

```
============================================================
STATUS REPORT:
============================================================
=== EPSON PRINTER STATUS REPORT ===
COM Port: COM1, Time: 2025-06-24 15:23:01
Status: OK - Printer online - ready to print.
Printer: ONLINE, Cover: CLOSED
Paper: PAPER OK
CanPrint: 1 (Can Print)
Ready to Print: YES
```

**Cover Open Error:**

```
============================================================
STATUS REPORT:
============================================================
=== EPSON PRINTER STATUS REPORT ===
COM Port: COM1, Time: 2025-06-24 15:23:37
Status: ERROR - Printer offline - cover is OPEN! Close the cover to resume printing.
Printer: OFFLINE, Cover: OPEN
Paper: UNKNOWN
CanPrint: 0 (Cannot Print)
Ready to Print: NO
WARNING: Printer cover is open
```

**Paper Out Error:**

```
============================================================
STATUS REPORT:
============================================================
=== EPSON PRINTER STATUS REPORT ===
COM Port: COM1, Time: 2025-06-24 15:24:09
Status: ERROR - Paper is OUT! Insert paper roll to resume printing.
Printer: OFFLINE, Cover: CLOSED
Paper: PAPER EMPTY
CanPrint: 0 (Cannot Print)
Ready to Print: NO
CRITICAL: Out of paper
```

**Paper Near End Warning:**

```
============================================================
STATUS REPORT:
============================================================
=== EPSON PRINTER STATUS REPORT ===
COM Port: COM1, Time: 2025-06-24 15:25:39
Status: OK - Paper is NEAR END - replace soon to avoid interruption.
Printer: ONLINE, Cover: CLOSED
Paper: PAPER NEAR END
CanPrint: 1 (Can Print)
Ready to Print: YES
WARNING: Replace paper soon
```

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

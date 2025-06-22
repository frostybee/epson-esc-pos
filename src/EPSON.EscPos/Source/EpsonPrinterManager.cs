using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Frostybee.EPSON.EscPos
{
    using System;
    using System.IO.Ports;
    using System.Text;
    using System.Threading;
    using System.Runtime.InteropServices;

    [ComVisible(true)]
    [Guid("E10E9F93-9368-41B9-8859-1DF69E424CEB")]
    [ClassInterface(ClassInterfaceType.None)]
    public class EpsonPrinterManager : IEpsonPrinterManager, IDisposable
    {
        // --- Configuration Constants.
        private const int BAUD_RATE = 38400; // Baud rate for the SerialPort class.
        private const Parity PARITY = Parity.None; // Parity for the SerialPort class.
        private const int DATA_BITS = 8; // Data bits for the SerialPort class.
        private const StopBits STOP_BITS = StopBits.One; // Stop bits for the SerialPort class.
        private const int ONLINE_TIMEOUT = 3000;  // This is the timeout for the ReadTimeout property of the SerialPort class.
        private const int OFFLINE_TIMEOUT = 1000; // Shorter timeout for offline detection.
        private const int DETECTION_TIMEOUT = 500; // Very short timeout for detection.
        private const int WRITE_TIMEOUT = 1000; // Timeout for the WriteTimeout property of the SerialPort class.
        private const int INIT_SLEEP_MS = 100; // Sleep time after sending the INIT command.
        private const int CLEAR_SLEEP_MS = 50; // Sleep time after sending the CLEAR command.

        // --- Status Byte Masks.
        private const byte FIXED_BIT_0_MASK = 0b00000001; // Bit 0 must be 0
        private const byte FIXED_BIT_1_MASK = 0b00000010; // Bit 1 must be 1
        private const byte FIXED_BIT_4_MASK = 0b00010000; // Bit 4 must be 1
        private const byte FIXED_BIT_7_MASK = 0b10000000; // Bit 7 must be 0
        private const byte COVER_OPEN_MASK = 0b00001000;  // Bit 3
        private const byte PAPER_OUT_MASK = 0x60;         // Bits 5-6 = 11 for paper out
        private const byte PAPER_NEAR_END_MASK = 0x0C;    // Bits 2-3 = 11 for near end
        private const byte PAPER_PRESENT_MASK = 0x60;     // Bits 5-6 = 00 for paper present

        // --- Command Constants.
        private static readonly byte[] INIT_CMD = { 0x1B, 0x40 };
        private static readonly byte[] CLEAR_NV_GRAPHICS_CMD = { 0x1C, 0x71, 0x01 };
        private static readonly byte[] GENERAL_STATUS_CMD = { 0x10, 0x04, 0x01 };
        private static readonly byte[] PAPER_STATUS_CMD = { 0x10, 0x04, 0x04 };

        // SerialPort instance for reuse.
        private SerialPort _serialPort;
        private string _currentComPort;

        // Thread safety.
        private readonly object _lockObject = new object();

        // Disposal tracking.
        private bool _disposed = false;

        /// <summary>
        /// Gets printer status using improved timeout-based detection.
        /// </summary>
        public IPrinterStatusResult GetPrinterStatus(string comPortName)
        {
            if (_disposed)
                throw new ObjectDisposedException(nameof(EpsonPrinterManager));

            var result = CreateDefaultResult();

            lock (_lockObject)
            {
                try
                {
                    // Ensure we have the correct serial port connection.
                    EnsureSerialPortConnection(comPortName);

                    // Try to get general status first.
                    if (TryGetGeneralStatus(result))
                    {
                        // Printer is online, check paper status if cover is closed.
                        if (result.CoverStatus == 1) // Cover closed
                        {
                            TryGetPaperStatus(result);
                        }
                    }
                    else
                    {
                        // Printer appears offline, try to determine reason.
                        DetectOfflineReason(result);
                    }
                }
                catch (UnauthorizedAccessException ex)
                {
                    SetError(result, $"Access to COM port '{comPortName}' denied. It might be in use by another application. Details: {ex.Message}");
                }
                catch (InvalidOperationException ex)
                {
                    SetError(result, $"COM port '{comPortName}' is invalid or configuration error. Details: {ex.Message}");
                }
                catch (ArgumentException ex)
                {
                    SetError(result, $"Invalid COM port name or settings. Details: {ex.Message}");
                }
                catch (Exception ex)
                {
                    SetError(result, $"Unexpected error: {ex.Message}");
                }
                finally
                {
                    CloseConnection();
                }
            }

            UpdateCanPrintStatus(result);

            return result;
        }

        /// <summary>
        /// Ensures we have a properly configured serial port connection.
        /// </summary>
        private void EnsureSerialPortConnection(string comPortName)
        {
            try
            {
                // Close existing connection if port changed or not properly configured.
                if (_serialPort != null && (_currentComPort != comPortName || !_serialPort.IsOpen))
                {
                    CloseConnection();
                }

                // Create new connection if needed.
                if (_serialPort == null)
                {
                    _serialPort = new SerialPort(comPortName, BAUD_RATE, PARITY, DATA_BITS, STOP_BITS)
                    {
                        ReadTimeout = ONLINE_TIMEOUT,
                        WriteTimeout = WRITE_TIMEOUT
                    };
                    _currentComPort = comPortName;
                }

                // Open connection if not already open.
                if (!_serialPort.IsOpen)
                {
                    _serialPort.Open();

                    // Initialize printer.
                    InitializePrinter();
                }
            }
            catch
            {
                CloseConnection(); // Ensure cleanup on failure.
                throw;
            }
        }

        /// <summary>
        /// Initializes the printer with standard commands. This is a standard command for the printer.
        /// </summary>
        private void InitializePrinter()
        {
            try
            {
                // ESC @ initializes the printer.
                _serialPort.Write(INIT_CMD, 0, INIT_CMD.Length);
                Thread.Sleep(INIT_SLEEP_MS);

                // Clear input buffer and NV graphics.
                _serialPort.DiscardInBuffer();
                _serialPort.Write(CLEAR_NV_GRAPHICS_CMD, 0, CLEAR_NV_GRAPHICS_CMD.Length);
                Thread.Sleep(CLEAR_SLEEP_MS);
            }
            catch (Exception ex)
            {
                // Log initialization failure but don't throw - it's not critical for status checking.
                System.Diagnostics.Debug.WriteLine($"Printer initialization warning: {ex.Message}");
            }
        }

        /// <summary>
        /// Attempts to get general printer status.
        /// </summary>
        private bool TryGetGeneralStatus(PrinterStatusResult result)
        {
            try
            {
                _serialPort.ReadTimeout = OFFLINE_TIMEOUT; // Use shorter timeout for initial detection.

                // DLE EOT n=1 command for General Printer Status.
                _serialPort.Write(GENERAL_STATUS_CMD, 0, GENERAL_STATUS_CMD.Length);

                byte[] buffer = new byte[1];
                int bytesRead = _serialPort.Read(buffer, 0, buffer.Length);

                if (bytesRead == 1)
                {
                    InterpretGeneralStatusByte(buffer[0], result);
                    return true; // Successfully got status.
                }
            }
            catch (TimeoutException)
            {
                // Timeout indicates printer is offline.
                result.PrinterStatus = 0; // Ensure offline status is set.
            }
            catch (Exception ex)
            {
                SetErrorWithOfflineStatus(result, $"Error checking general status: {ex.Message}");
            }

            return false; // Failed to get status.
        }

        /// <summary>
        /// Attempts to get paper status when printer is online.
        /// </summary>
        private void TryGetPaperStatus(PrinterStatusResult result)
        {
            try
            {
                _serialPort.ReadTimeout = ONLINE_TIMEOUT; // Use longer timeout when we know printer is online.
                _serialPort.DiscardInBuffer();

                // DLE EOT n=4 command for Paper Status.
                _serialPort.Write(PAPER_STATUS_CMD, 0, PAPER_STATUS_CMD.Length);

                byte[] buffer = new byte[1];
                int bytesRead = _serialPort.Read(buffer, 0, buffer.Length);

                if (bytesRead > 0)
                {
                    InterpretPaperStatusByte(buffer[0], result);
                }
            }
            catch (TimeoutException)
            {
                // If we can't get paper status but general status worked, assume paper is OK.
                result.PaperStatus = 0; // Paper OK (default assumption when online).
            }
            catch (Exception ex)
            {
                SetError(result, $"Error checking paper status: {ex.Message}");
            }
        }

        /// <summary>
        /// Attempts to determine why printer is offline using alternative methods.
        /// </summary>
        private void DetectOfflineReason(PrinterStatusResult result)
        {
            // Set offline status.
            result.PrinterStatus = 0; // Offline.

            try
            {
                // Save original timeout and use very short timeout for detection.
                int originalTimeout = _serialPort.ReadTimeout;
                _serialPort.ReadTimeout = DETECTION_TIMEOUT;

                // Try initialization and status check again.
                InitializePrinter();

                _serialPort.Write(GENERAL_STATUS_CMD, 0, GENERAL_STATUS_CMD.Length);

                byte[] response = new byte[1];
                int bytesRead = _serialPort.Read(response, 0, 1);

                if (bytesRead > 0)
                {
                    // Got response after init - printer came online.
                    InterpretGeneralStatusByte(response[0], result);
                    if (result.CoverStatus == 1) // Cover closed.
                    {
                        TryGetPaperStatus(result);
                    }
                }
                else
                {
                    // Still no response.
                    SetErrorWithOfflineStatus(result, "Printer offline - possible causes: cover open, paper out, power off, or Memory Switch 1-3 not configured for status transmission.");
                }

                _serialPort.ReadTimeout = originalTimeout;
            }
            catch (TimeoutException)
            {
                SetErrorWithOfflineStatus(result, "Printer offline - no response. Check power, cable connections, and ensure Memory Switch 1-3 is ON for automatic status transmission.");
            }
            catch (Exception ex)
            {
                SetErrorWithOfflineStatus(result, $"Error during offline detection: {ex.Message}");
            }
        }

        /// <summary>
        /// Interprets the general status byte returned by the printer.
        /// </summary>
        private void InterpretGeneralStatusByte(byte statusByte, PrinterStatusResult result)
        {
            // Validate fixed identifier bits.
            bool bit0Fixed = (statusByte & FIXED_BIT_0_MASK) == 0; // Bit 0 must be 0
            bool bit1Fixed = (statusByte & FIXED_BIT_1_MASK) != 0; // Bit 1 must be 1
            bool bit4Fixed = (statusByte & FIXED_BIT_4_MASK) != 0; // Bit 4 must be 1
            bool bit7Fixed = (statusByte & FIXED_BIT_7_MASK) == 0; // Bit 7 must be 0

            if (!bit0Fixed || !bit1Fixed || !bit4Fixed || !bit7Fixed)
            {
                SetErrorWithOfflineStatus(result, $"Invalid status response (0x{statusByte:X2}) - fixed identifier bits don't match expected values.");
                return;
            }

            // Interpret status bits.
            bool isCoverOpen = (statusByte & COVER_OPEN_MASK) != 0; // Bit 3

            // Set status values consistently.
            result.CoverStatus = isCoverOpen ? 0 : 1; // Open = 0, Closed = 1
            result.PrinterStatus = isCoverOpen ? 0 : 1; // Offline if cover open, Online if cover closed

            // Set appropriate message.
            if (isCoverOpen)
            {
                SetError(result, "Printer offline - cover is OPEN! Close the cover to resume printing.");
            }
            else
            {
                result.PrinterStatus = 1; // Online
                result.ErrorMessage = "Printer online - cover closed.";
                result.HasError = 0; // Clear error state when online
            }
        }

        /// <summary>
        /// Interprets the paper status byte returned by the printer.
        /// </summary>
        private void InterpretPaperStatusByte(byte statusByte, PrinterStatusResult result)
        {
            // Correct interpretation for TM-T20III.
            bool paperOut = (statusByte & PAPER_OUT_MASK) == PAPER_OUT_MASK;      // Bits 5-6 = 11
            bool paperNearEnd = (statusByte & PAPER_NEAR_END_MASK) == PAPER_NEAR_END_MASK;  // Bits 2-3 = 11
            bool paperPresent = (statusByte & PAPER_PRESENT_MASK) == 0x00;  // Bits 5-6 = 00

            if (paperOut)
            {
                result.PaperStatus = 2; // Paper Out
                result.PrinterStatus = 0; // Offline due to paper out
                SetError(result, "Paper is OUT! Insert paper roll to resume printing.");
            }
            else if (paperNearEnd)
            {
                result.PaperStatus = 1; // Paper Near End
                result.ErrorMessage = "Paper is NEAR END - replace soon to avoid interruption.";
                // Keep printer online for near-end condition.
            }
            else if (paperPresent)
            {
                result.PaperStatus = 0; // Paper OK
                                        // Keep existing error message if any.
            }
            else
            {
                result.PaperStatus = 99; // Unknown
                result.ErrorMessage += " (Paper status unclear - check paper roll)";
            }
        }

        /// <summary>
        /// Creates a default result object with initial values.
        /// </summary>
        private PrinterStatusResult CreateDefaultResult()
        {
            return new PrinterStatusResult
            {
                PaperStatus = 99,    // Unknown
                PrinterStatus = 0,   // Offline (assume offline until proven otherwise)
                CoverStatus = 1,     // Closed (assume closed)
                HasError = 0,        // No Error
                ErrorMessage = string.Empty,
                CanPrint = 0
            };
        }

        /// <summary>
        /// Sets error state on result object.
        /// </summary>
        private void SetError(PrinterStatusResult result, string message)
        {
            result.HasError = 1;
            result.ErrorMessage = message;
        }

        /// <summary>
        /// Sets error state and ensures offline status.
        /// </summary>
        private void SetErrorWithOfflineStatus(PrinterStatusResult result, string message)
        {
            result.HasError = 1;
            result.ErrorMessage = message;
            result.PrinterStatus = 0; // Ensure offline status. 
        }

        /// <summary>
        /// Updates the CanPrint property based on current status conditions.
        /// </summary>
        private void UpdateCanPrintStatus(PrinterStatusResult result)
        {
            // Can print if:
            // - No errors (HasError == 0)
            // - Printer is online (PrinterStatus == 1)
            // - Cover is closed (CoverStatus == 1)
            // - Paper is OK or Near End (PaperStatus == 0 or 1) - Near end still allows printing.
            result.CanPrint = (result.HasError == 0 &&
                              result.PrinterStatus == 1 &&
                              result.CoverStatus == 1 &&
                              (result.PaperStatus == 0 || result.PaperStatus == 1)) ? 1 : 0;
        }

        /// <summary>
        /// Closes and disposes the serial port connection.
        /// </summary>
        private void CloseConnection()
        {
            if (_serialPort != null)
            {
                try
                {
                    if (_serialPort.IsOpen)
                    {
                        _serialPort.Close();
                    }
                    _serialPort.Dispose();
                }
                catch (Exception ex)
                {
                    // Log disposal errors but don't throw.
                    System.Diagnostics.Debug.WriteLine($"Serial port disposal warning: {ex.Message}");
                }
                finally
                {
                    _serialPort = null;
                    _currentComPort = null;
                }
            }
        }

        /// <summary>
        /// Gets a formatted string containing a comprehensive printer status report.
        /// </summary>
        public string GetStatusReport(string comPortName)
        {
            if (_disposed)
                throw new ObjectDisposedException(nameof(EpsonPrinterManager));

            var status = GetPrinterStatus(comPortName);
            var report = new StringBuilder();

            // Helper method to get paper status text.
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

            // Build report.
            report.AppendLine("=== EPSON PRINTER STATUS REPORT ===");
            report.AppendLine($"COM Port: {comPortName}, Time: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            report.AppendLine($"Status: {(status.HasError == 1 ? "ERROR" : "OK")} - {status.ErrorMessage}");
            report.AppendLine($"Printer: {(status.PrinterStatus == 1 ? "ONLINE" : "OFFLINE")}, Cover: {(status.CoverStatus == 1 ? "CLOSED" : "OPEN")}");
            report.AppendLine($"Paper: {GetPaperStatusText(status.PaperStatus)}");
            report.AppendLine($"CanPrint: {status.CanPrint} ({(status.CanPrint == 1 ? "Can Print" : "Cannot Print")})");
            report.AppendLine($"Ready to Print: {(status.CanPrint == 1 ? "YES" : "NO")}");

            // Add warnings if needed.
            if (status.PaperStatus == 1) report.AppendLine("WARNING: Replace paper soon");
            if (status.PaperStatus == 2) report.AppendLine("CRITICAL: Out of paper");
            if (status.CoverStatus == 0) report.AppendLine("WARNING: Printer cover is open");

            return report.ToString();
        }

        #region IDisposable Implementation

        /// <summary>
        /// Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.
        /// </summary>
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        /// <summary>
        /// Protected implementation of Dispose pattern.
        /// </summary>
        protected virtual void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                if (disposing)
                {
                    lock (_lockObject)
                    {
                        CloseConnection();
                    }
                }
                _disposed = true;
            }
        }

        /// <summary>
        /// Finalizer to ensure resources are cleaned up.
        /// </summary>
        ~EpsonPrinterManager()
        {
            Dispose(false);
        }

        #endregion
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Frostybee.EpsonEscPos
{
    using System;
    using System.IO.Ports;
    using System.Text;
    using System.Threading;
    using System.Runtime.InteropServices;

    [ComVisible(true)]
    [Guid("E10E9F93-9368-41B9-8859-1DF69E424CEB")]
    [ClassInterface(ClassInterfaceType.None)]
    public class EpsonPrinterManager : IPrinterManager, IDisposable
    {
        // --- Default configuration for backward compatibility.
        private static readonly PrinterConfiguration DefaultConfiguration = new PrinterConfiguration();

        // --- Status Byte Masks.
        private const byte FIXED_BIT_0_MASK = 0b00000001; // Bit 0 must be 0.
        private const byte FIXED_BIT_1_MASK = 0b00000010; // Bit 1 must be 1.
        private const byte FIXED_BIT_4_MASK = 0b00010000; // Bit 4 must be 1.
        private const byte FIXED_BIT_7_MASK = 0b10000000; // Bit 7 must be 0.
        private const byte COVER_OPEN_MASK = 0b00001000;  // Bit 3.
        private const byte PAPER_OUT_MASK = 0x60;         // Bits 5-6 = 11 for paper out.
        private const byte PAPER_NEAR_END_MASK = 0x0C;    // Bits 2-3 = 11 for near end.
        private const byte PAPER_PRESENT_MASK = 0x60;     // Bits 5-6 = 00 for paper present.

        // --- Command Constants.
        // @see: https://download4.epson.biz/sec_pubs/pos/reference_en/escpos/dle_eot.html
        private static readonly byte[] INIT_CMD = { 0x1B, 0x40 };
        private static readonly byte[] CLEAR_NV_GRAPHICS_CMD = { 0x1C, 0x71, 0x01 };
        // Printer status command (n = 1)
        private static readonly byte[] GENERAL_STATUS_CMD = { 0x10, 0x04, 0x01 }; 
        // Offline cause status command (n = 2)
        private static readonly byte[] OFFLINE_CAUSE_STATUS_CMD = { 0x10, 0x04, 0x02 }; 
        // Roll paper sensor status command (n = 4)
        private static readonly byte[] PAPER_STATUS_CMD = { 0x10, 0x04, 0x04 }; 

        // SerialPort instance for reuse.
        private SerialPort _serialPort;
        private string _currentComPort;


        // Disposal tracking.
        private bool _disposed = false;

        /// <summary>
        /// Gets printer status using improved timeout-based detection.
        /// </summary>        
        public IPrinterStatusResult GetPrinterStatus(string comPortName)
        {
            return GetPrinterStatus(comPortName, DefaultConfiguration);
        }

        /// <summary>
        /// Gets printer status using improved timeout-based detection with custom configuration.
        /// </summary>
        public IPrinterStatusResult GetPrinterStatus(string comPortName, PrinterConfiguration configuration)
        {
            if (_disposed)
                throw new ObjectDisposedException(nameof(EpsonPrinterManager));

            if (configuration == null)
                throw new ArgumentNullException(nameof(configuration));

            var result = CreateDefaultResult();

            try
            {
                // Ensure we have the correct serial port connection.
                EnsureSerialPortConnection(comPortName, configuration);

                // Check if printer is online.
                bool isOnline = ReadGeneralPrinterStatus(configuration, result);
                result.PrinterStatus = isOnline ? 1 : 0;

                if (isOnline)
                {
                    // Printer is online, check cover status.
                    bool isCoverOpen = ReadCoverStatus(configuration, result);
                    result.CoverStatus = isCoverOpen ? 0 : 1; // 0 = Open, 1 = Closed.

                    if (isCoverOpen)
                    {
                        SetError(result, "Printer offline - cover is OPEN! Close the cover to resume printing.");
                        result.PrinterStatus = 0; // Set offline due to cover open.
                    }
                    else
                    {
                        // Cover is closed, check paper status.
                        result.PaperStatus = ReadPaperRollStatus(configuration, result);
                    }
                }
                else
                {
                    // Printer is offline - attempt to determine the cause using OFFLINE_CAUSE_STATUS_CMD
                    bool isCoverOpen = ReadOfflineCoverStatus(configuration, result);
                    result.CoverStatus = isCoverOpen ? 0 : 1; // 0 = Open, 1 = Closed.

                    if (isCoverOpen)
                    {
                        SetError(result, "Printer offline - cover is OPEN! Close the cover to resume printing.");
                    }
                    else
                    {
                        SetErrorWithOfflineStatus(result, "Printer offline - no response. Check power, cable connections, and ensure Memory Switch 1-3 is ON for automatic status transmission.");
                    }

                    result.PaperStatus = 99; // Unknown when offline.
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

            UpdateCanPrintStatus(result);

            return result;
        }

        /// <summary>
        /// Ensures we have a properly configured serial port connection.
        /// </summary>
        private void EnsureSerialPortConnection(string comPortName, PrinterConfiguration configuration)
        {
            if (IsConnectionValid(comPortName))
                return;

            CloseConnection();
            CreateAndOpenConnection(comPortName, configuration);
        }

        /// <summary>
        /// Initializes the printer with standard commands. This is a standard command for the printer.
        /// </summary>
        private void InitializePrinter(PrinterConfiguration configuration)
        {
            try
            {
                // ESC @ initializes the printer.
                _serialPort.Write(INIT_CMD, 0, INIT_CMD.Length);

                // Clear input buffer and NV graphics.
                _serialPort.DiscardInBuffer();
                _serialPort.DiscardOutBuffer();
                _serialPort.Write(CLEAR_NV_GRAPHICS_CMD, 0, CLEAR_NV_GRAPHICS_CMD.Length);
            }
            catch (Exception ex)
            {
                // Log initialization failure but don't throw - it's not critical for status checking.
                System.Diagnostics.Debug.WriteLine($"Printer initialization warning: {ex.Message}");
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
        /// Reads the general printer status to determine if printer is online or offline.
        /// Returns true if printer is online, false if offline.
        /// Error Handling Strategy: All communication errors result in offline status.
        /// Timeouts are expected when printer is off/disconnected and are reported to the caller.
        /// </summary>
        private bool ReadGeneralPrinterStatus(PrinterConfiguration configuration, PrinterStatusResult result)
        {
            try
            {
                int originalTimeout = _serialPort.ReadTimeout;
                _serialPort.ReadTimeout = configuration.OfflineTimeout;

                try
                {
                    _serialPort.DiscardInBuffer();
                    _serialPort.DiscardOutBuffer();
                    _serialPort.Write(GENERAL_STATUS_CMD, 0, GENERAL_STATUS_CMD.Length);

                    byte[] buffer = new byte[1];
                    int bytesRead = _serialPort.Read(buffer, 0, buffer.Length);

                    if (bytesRead == 1)
                    {
                        // Validate fixed identifier bits.
                        byte statusByte = buffer[0];
                        bool bit0Fixed = (statusByte & FIXED_BIT_0_MASK) == 0;
                        bool bit1Fixed = (statusByte & FIXED_BIT_1_MASK) != 0;
                        bool bit4Fixed = (statusByte & FIXED_BIT_4_MASK) != 0;
                        bool bit7Fixed = (statusByte & FIXED_BIT_7_MASK) == 0;

                        if (bit0Fixed && bit1Fixed && bit4Fixed && bit7Fixed)
                        {
                            return true; // Printer is online.
                        }
                    }
                }
                finally
                {
                    _serialPort.ReadTimeout = originalTimeout;
                }
            }
            catch (TimeoutException)
            {
                // Timeout indicates printer is offline - this is expected behavior when printer is off/disconnected
                SetError(result, "Printer status check timed out - printer appears to be offline");
            }
            catch (InvalidOperationException ex)
            {
                // Serial port not open or in invalid state - indicates connection issues
                SetError(result, $"Serial port operation failed during status check: {ex.Message}");
            }
            catch (Exception ex)
            {
                // Unexpected error during status communication - log for diagnostics
                SetError(result, $"Unexpected error during printer status check: {ex.Message}");
            }

            return false; // Printer is offline.
        }

        /// <summary>
        /// Reads the cover status when printer is offline using offline cause status command (n=2).
        /// Returns true if cover is open, false if cover is closed or unknown.
        /// Error Handling Strategy: Communication failures default to 'cover closed' as the safe assumption.
        /// This prevents false 'cover open' alerts when communication is simply unreliable.
        /// </summary>
        private bool ReadOfflineCoverStatus(PrinterConfiguration configuration, PrinterStatusResult result)
        {
            try
            {
                int originalTimeout = _serialPort.ReadTimeout;
                _serialPort.ReadTimeout = configuration.OfflineTimeout; // Use shorter timeout for offline detection

                try
                {
                    _serialPort.DiscardInBuffer();
                    _serialPort.DiscardOutBuffer();
                    _serialPort.Write(OFFLINE_CAUSE_STATUS_CMD, 0, OFFLINE_CAUSE_STATUS_CMD.Length);

                    byte[] buffer = new byte[1];
                    int bytesRead = _serialPort.Read(buffer, 0, buffer.Length);

                    if (bytesRead == 1)
                    {
                        byte statusByte = buffer[0];
                        // Bit 2 of DLE EOT n=2: Cover status (0 = closed, 1 = open).
                        bool isCoverOpen = (statusByte & 0b00000100) != 0;  // Check bit 2 (0x04).
                        System.Diagnostics.Debug.WriteLine($"Offline cover status check successful: Cover {(isCoverOpen ? "OPEN" : "CLOSED")}");
                        return isCoverOpen;
                    }
                    else
                    {
                        SetError(result, "Offline cover status check: No response received");
                    }
                }
                finally
                {
                    _serialPort.ReadTimeout = originalTimeout;
                }
            }
            catch (TimeoutException)
            {
                // Timeout reading offline cover status - assume closed as default safe state
                SetError(result, "Offline cover status check timed out - defaulting to closed");
            }
            catch (InvalidOperationException ex)
            {
                // Serial port operation failed - log and default to closed
                SetError(result, $"Serial port error reading offline cover status: {ex.Message} - defaulting to closed");
            }
            catch (Exception ex)
            {
                // Unexpected error reading offline cover status - log and default to closed for safety
                SetError(result, $"Unexpected error reading offline cover status: {ex.Message} - defaulting to closed");
            }

            return false; // Default to cover closed when offline and cannot determine status.
        }

        /// <summary>
        /// Reads the cover status using offline cause status command (n=2).
        /// Returns true if cover is open, false if cover is closed.
        /// Error Handling Strategy: Communication failures default to 'cover closed' as the safe assumption.
        /// This prevents false 'cover open' alerts when communication is simply unreliable.
        /// </summary>
        private bool ReadCoverStatus(PrinterConfiguration configuration, PrinterStatusResult result)
        {
            try
            {
                int originalTimeout = _serialPort.ReadTimeout;
                _serialPort.ReadTimeout = configuration.OnlineTimeout;

                try
                {
                    _serialPort.DiscardInBuffer();
                    _serialPort.DiscardOutBuffer();
                    _serialPort.Write(OFFLINE_CAUSE_STATUS_CMD, 0, OFFLINE_CAUSE_STATUS_CMD.Length);

                    byte[] buffer = new byte[1];
                    int bytesRead = _serialPort.Read(buffer, 0, buffer.Length);

                    if (bytesRead == 1)
                    {
                        byte statusByte = buffer[0];
                        // Bit 2 of DLE EOT n=2: Cover status (0 = closed, 1 = open).
                        bool isCoverOpen = (statusByte & 0b00000100) != 0;  // Check bit 2 (0x04).
                        return isCoverOpen;
                    }
                }
                finally
                {
                    _serialPort.ReadTimeout = originalTimeout;
                }
            }
            catch (TimeoutException)
            {
                // Timeout reading cover status - assume closed as default safe state
                SetError(result, "Cover status check timed out - defaulting to closed");
            }
            catch (InvalidOperationException ex)
            {
                // Serial port operation failed - log and default to closed
                SetError(result, $"Serial port error reading cover status: {ex.Message} - defaulting to closed");
            }
            catch (Exception ex)
            {
                // Unexpected error reading cover status - log and default to closed for safety
                SetError(result, $"Unexpected error reading cover status: {ex.Message} - defaulting to closed");
            }

            return false; // Default to cover closed.
        }

        /// <summary>
        /// Reads the paper roll status and sets appropriate messages and printer status.
        /// Returns: 0 = Paper OK, 1 = Paper Near End, 2 = Paper Empty, 99 = Unknown
        /// Error Handling Strategy: Communication failures return 'Unknown' status (99).
        /// This allows the application to handle uncertain paper states appropriately.
        /// </summary>
        private int ReadPaperRollStatus(PrinterConfiguration configuration, PrinterStatusResult result)
        {
            int paperStatus = 99; // Default to unknown.

            try
            {
                int originalTimeout = _serialPort.ReadTimeout;
                _serialPort.ReadTimeout = configuration.OnlineTimeout;

                try
                {
                    _serialPort.DiscardInBuffer();
                    _serialPort.DiscardOutBuffer();
                    _serialPort.Write(PAPER_STATUS_CMD, 0, PAPER_STATUS_CMD.Length);

                    byte[] buffer = new byte[1];
                    int bytesRead = _serialPort.Read(buffer, 0, buffer.Length);

                    if (bytesRead == 1)
                    {
                        byte statusByte = buffer[0];

                        // Correct interpretation for TM-T20III
                        bool paperOut = (statusByte & PAPER_OUT_MASK) == PAPER_OUT_MASK;      // Bits 5-6 = 11
                        bool paperNearEnd = (statusByte & PAPER_NEAR_END_MASK) == PAPER_NEAR_END_MASK;  // Bits 2-3 = 11
                        bool paperPresent = (statusByte & PAPER_PRESENT_MASK) == 0x00;  // Bits 5-6 = 00

                        if (paperOut)
                        {
                            paperStatus = 2; // Paper Empty.
                        }
                        else if (paperNearEnd)
                        {
                            paperStatus = 1; // Paper Near End.
                        }
                        else if (paperPresent)
                        {
                            paperStatus = 0; // Paper OK.
                        }
                    }
                }
                finally
                {
                    _serialPort.ReadTimeout = originalTimeout;
                }
            }
            catch (TimeoutException)
            {
                // Timeout reading paper status - return unknown as we cannot determine state
                SetError(result, "Paper status check timed out - returning unknown status");
            }
            catch (InvalidOperationException ex)
            {
                // Serial port operation failed - log and return unknown
                SetError(result, $"Serial port error reading paper status: {ex.Message} - returning unknown");
            }
            catch (Exception ex)
            {
                // Unexpected error reading paper status - log and return unknown
                SetError(result, $"Unexpected error reading paper status: {ex.Message} - returning unknown");
            }

            // Set appropriate messages based on paper status
            switch (paperStatus)
            {
                case 0: // Paper OK
                    result.ErrorMessage = "Printer online - ready to print.";
                    result.HasError = 0;
                    break;
                case 1: // Paper Near End
                    result.ErrorMessage = "Paper is NEAR END - replace soon to avoid interruption.";
                    result.HasError = 0; // Not an error, just a warning
                    break;
                case 2: // Paper Empty
                    SetError(result, "Paper is OUT! Insert paper roll to resume printing.");
                    result.PrinterStatus = 0; // Set offline due to paper out
                    break;
                case 99: // Unknown
                    result.ErrorMessage = "Printer online - paper status unclear, check paper roll.";
                    result.HasError = 0;
                    break;
            }

            return paperStatus;
        }

        /// <summary>
        /// Checks if the current connection is valid for the specified COM port.
        /// </summary>
        private bool IsConnectionValid(string comPortName)
        {
            return _serialPort != null &&
                   _currentComPort == comPortName &&
                   _serialPort.IsOpen;
        }

        /// <summary>
        /// Creates and opens a new serial port connection.
        /// </summary>
        private void CreateAndOpenConnection(string comPortName, PrinterConfiguration configuration)
        {
            try
            {
                _serialPort = CreateSerialPort(comPortName, configuration);
                _currentComPort = comPortName;

                _serialPort.Open();
                InitializePrinter(configuration);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Failed to create connection to {comPortName}: {ex.Message}");
                CloseConnection();
                throw;
            }
        }

        /// <summary>
        /// Creates a configured SerialPort instance.
        /// </summary>
        private SerialPort CreateSerialPort(string comPortName, PrinterConfiguration configuration)
        {
            return new SerialPort(comPortName, configuration.BaudRate, configuration.Parity, configuration.DataBits, configuration.StopBits)
            {
                ReadTimeout = configuration.OnlineTimeout,
                WriteTimeout = configuration.WriteTimeout
            };
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
            return GetStatusReport(comPortName, DefaultConfiguration);
        }

        /// <summary>
        /// Gets a formatted string containing a comprehensive printer status report with custom configuration.
        /// </summary>
        public string GetStatusReport(string comPortName, PrinterConfiguration configuration)
        {
            if (_disposed)
                throw new ObjectDisposedException(nameof(EpsonPrinterManager));

            if (configuration == null)
                throw new ArgumentNullException(nameof(configuration));

            var status = GetPrinterStatus(comPortName, configuration);
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
                    CloseConnection();
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

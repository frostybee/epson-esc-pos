using System.IO.Ports;

namespace Frostybee.EpsonEscPos
{
    /// <summary>
    /// Interface defining configuration options for printer communication.
    /// </summary>
    public interface IPrinterConfiguration
    {
        /// <summary>
        /// Gets the baud rate for serial communication.
        /// </summary>
        int BaudRate { get; }

        /// <summary>
        /// Gets the parity setting for serial communication.
        /// </summary>
        Parity Parity { get; }

        /// <summary>
        /// Gets the data bits for serial communication.
        /// </summary>
        int DataBits { get; }

        /// <summary>
        /// Gets the stop bits for serial communication.
        /// </summary>
        StopBits StopBits { get; }

        /// <summary>
        /// Gets the timeout for operations when printer is online (milliseconds).
        /// </summary>
        int OnlineTimeout { get; }

        /// <summary>
        /// Gets the timeout for operations when printer is offline (milliseconds).
        /// </summary>
        int OfflineTimeout { get; }

        /// <summary>
        /// Gets the timeout for detection operations (milliseconds).
        /// </summary>
        int DetectionTimeout { get; }

        /// <summary>
        /// Gets the write timeout for serial communication (milliseconds).
        /// </summary>
        int WriteTimeout { get; }

        /// <summary>
        /// Gets the sleep time after sending initialization commands (milliseconds).
        /// </summary>
        int InitializationSleepMs { get; }

        /// <summary>
        /// Gets the sleep time after sending clear commands (milliseconds).
        /// </summary>
        int ClearSleepMs { get; }
    }
} 
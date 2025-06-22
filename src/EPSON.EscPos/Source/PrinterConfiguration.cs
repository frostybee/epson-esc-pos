using System.IO.Ports;

namespace Frostybee.EpsonEscPos
{
    /// <summary>
    /// Default implementation of printer configuration with standard EPSON ESC/POS settings.
    /// </summary>
    public class PrinterConfiguration : IPrinterConfiguration
    {
        /// <summary>
        /// Initializes a new instance of PrinterConfiguration with default values.
        /// </summary>
        public PrinterConfiguration()
        {
            BaudRate = 38400;
            Parity = Parity.None;
            DataBits = 8;
            StopBits = StopBits.One;
            OnlineTimeout = 3000;
            OfflineTimeout = 1000;
            DetectionTimeout = 500;
            WriteTimeout = 1000;
            InitializationSleepMs = 100;
            ClearSleepMs = 50;
        }

        /// <summary>
        /// Initializes a new instance of PrinterConfiguration with custom values.
        /// </summary>
        public PrinterConfiguration(
            int baudRate = 38400,
            Parity parity = Parity.None,
            int dataBits = 8,
            StopBits stopBits = StopBits.One,
            int onlineTimeout = 3000,
            int offlineTimeout = 1000,
            int detectionTimeout = 500,
            int writeTimeout = 1000,
            int initializationSleepMs = 100,
            int clearSleepMs = 50)
        {
            BaudRate = baudRate;
            Parity = parity;
            DataBits = dataBits;
            StopBits = stopBits;
            OnlineTimeout = onlineTimeout;
            OfflineTimeout = offlineTimeout;
            DetectionTimeout = detectionTimeout;
            WriteTimeout = writeTimeout;
            InitializationSleepMs = initializationSleepMs;
            ClearSleepMs = clearSleepMs;
        }

        /// <inheritdoc />
        public int BaudRate { get; }

        /// <inheritdoc />
        public Parity Parity { get; }

        /// <inheritdoc />
        public int DataBits { get; }

        /// <inheritdoc />
        public StopBits StopBits { get; }

        /// <inheritdoc />
        public int OnlineTimeout { get; }

        /// <inheritdoc />
        public int OfflineTimeout { get; }

        /// <inheritdoc />
        public int DetectionTimeout { get; }

        /// <inheritdoc />
        public int WriteTimeout { get; }

        /// <inheritdoc />
        public int InitializationSleepMs { get; }

        /// <inheritdoc />
        public int ClearSleepMs { get; }

        /// <summary>
        /// Creates a configuration optimized for high-speed printing.
        /// </summary>
        public static PrinterConfiguration CreateHighSpeed()
        {
            return new PrinterConfiguration(
                baudRate: 115200,
                onlineTimeout: 5000,
                writeTimeout: 2000
            );
        }

        /// <summary>
        /// Creates a configuration optimized for reliable communication.
        /// </summary>
        public static PrinterConfiguration CreateReliable()
        {
            return new PrinterConfiguration(
                baudRate: 9600,
                onlineTimeout: 5000,
                offlineTimeout: 2000,
                detectionTimeout: 1000,
                writeTimeout: 3000
            );
        }
    }
} 
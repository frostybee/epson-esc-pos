using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace Frostybee.EPSON.EscPos
{
    /// <summary>
    /// COM interface for EpsonPrinterManager. This is the class that implements the IPrinterStatusResult interface.    
    /// </summary>
    [ComVisible(true)]
    [Guid("34229201-9B5E-4C41-B6FB-CE1CDC05ADF3")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IEpsonPrinterManager
    {
        IPrinterStatusResult GetPrinterStatus(string comPortName);
        string GetStatusReport(string comPortName);
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace Frostybee.EpsonEscPos
{
    [ComVisible(true)]
    [Guid("237278CD-E18F-4C87-B740-14FE6A1C19CE")]
    [ClassInterface(ClassInterfaceType.None)]
    public class PrinterStatusResult : IPrinterStatusResult
    {
        public int PaperStatus { get; set; }  // 0 = Paper OK, 1 = Paper Near End, 2 = Paper Empty, 99 = Unknown
        public int PrinterStatus { get; set; }  // 0 = Offline, 1 = Online
        public int CoverStatus { get; set; }  // 0 = Open, 1 = Closed
        public int CanPrint { get; set; }    // 0 = Cannot Print, 1 = Can Print
        public int HasError { get; set; }    // 0 = No Error, 1 = Has Error
        public string ErrorMessage { get; set; } = string.Empty;
    }
}

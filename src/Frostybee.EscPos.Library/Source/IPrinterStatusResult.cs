using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace Frostybee.EpsonEscPos
{
    /// <summary>
    /// COM interface for PrinterStatusResult. This is the result of the GetPrinterStatus method.
    /// </summary>
    [ComVisible(true)]
    [Guid("A1B2C3D4-E5F6-7890-ABCD-123456789ABC")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IPrinterStatusResult
    {
        int PaperStatus { get; set; }
        int PrinterStatus { get; set; }
        int CoverStatus { get; set; }
        int CanPrint { get; set; }
        int HasError { get; set; }
        string ErrorMessage { get; set; }
    }
}

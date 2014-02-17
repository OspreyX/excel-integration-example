using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using ExcelDna.Integration;
using ExcelDna.Integration.Rtd;

namespace Openfin.RTDAddin
{
    /// <summary>
    ///     Export functions to Excel.
    /// </summary>
    public static class RtdFunction
    {

        /// <summary>
        ///     Export "FinDesktop" to Excel.
        /// </summary>
        [ExcelFunction(Name = "FinDesktop")]
        public static object FinDesktop(string topic1, string topic2, string topic3)
        {
            // "Openfin.RTDAddin.FinDesktopServer" is defined in RtdServer.cs
            // topics should have AppId, topic, and any additional params
            return XlCall.RTD("Openfin.RTDAddin.FinDesktopServer", null, topic1, topic2, topic3);
        }
            
    }
}

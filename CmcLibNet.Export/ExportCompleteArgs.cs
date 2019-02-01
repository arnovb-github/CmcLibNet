using System;
using System.Runtime.InteropServices;

namespace Vovin.CmcLibNet.Export
{

    /// <summary>
    /// Passed when export has completed.
    /// </summary>
    [ComVisible(true)]
    [Guid("A6BFDD24-1078-419E-9D37-5C338A459758")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IExportCompleteArgs))]
    public class ExportCompleteArgs : EventArgs, IExportCompleteArgs
    {
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="rows">Rows</param>
        public ExportCompleteArgs(int rows)
        {
            Rows = rows;
        }
        /// <summary>
        /// Number of rows
        /// </summary>
        public int Rows { get; private set; }
    }
}

using System;
using System.Runtime.InteropServices;

namespace Vovin.CmcLibNet.Export
{
    /// <summary>
    /// ExportProgressAsStringChangedArgs class.
    /// Reports export progress and data to outside the assembly.
    /// </summary>
    [ComVisible(true)]
    [Guid("40919A47-E618-4C5F-AF94-EAFCEA5B3F0D")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IExportProgressChangedArgs))]
    public class ExportProgressChangedArgs : EventArgs, IExportProgressChangedArgs
    {
        /// <summary>
        /// Constructor.
        /// </summary> 
        /// <param name="rowsProcessed">Current row.</param>
        /// <param name="totalRows">Total number of rows to read.</param>
        internal ExportProgressChangedArgs(int rowsProcessed, int totalRows)
        {
            this.RowsProcessed = rowsProcessed;
            this.RowsTotal = totalRows;
        }

        /// <summary>
        /// Overloaded constructor with data.
        /// </summary>
        /// <param name="rowsProcessed">Current row.</param>
        /// <param name="totalRows">Total number of rows to read.</param>
        /// <param name="data">String</param>
        internal ExportProgressChangedArgs(int rowsProcessed, int totalRows, string data)
        {
            this.RowsProcessed = rowsProcessed;
            this.RowsTotal = totalRows;
            this.RowValues = data;
        }

        /// <summary>
        /// Overloaded constructor with iterations.
        /// </summary>
        /// <param name="rowsProcessed">Current row.</param>
        /// <param name="totalRows">Total number of rows to read.</param>
        /// <param name="currentIteration">Current iteration.</param>
        /// <param name="iterationCount">Total number of iterations.</param>
        internal ExportProgressChangedArgs(int rowsProcessed, int totalRows, int currentIteration, int iterationCount)
        {
            this.RowsProcessed = rowsProcessed;
            this.RowsTotal = totalRows;
            this.CurrentIteration = currentIteration;
            this.IterationCount = iterationCount;
        }

        /// <inheritdoc />
        public int RowsProcessed { get; }
        /// <inheritdoc />
        public int RowsTotal { get;  }
        /// <inheritdoc />
        public string RowValues { get; } = "Only populated when used with ExportSettings.ExportFormat.Event.";
        /// <inheritdoc />
        public int CurrentIteration { get; }
        /// <inheritdoc />
        public int IterationCount { get; }
    }
}

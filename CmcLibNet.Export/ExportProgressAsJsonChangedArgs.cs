using System;
using System.Runtime.InteropServices;

namespace Vovin.CmcLibNet.Export
{
    /// <summary>
    /// Interface for ExportProgressChanged, primarily for use with COM Interop.
    /// </summary>
    [ComVisible(true)]
    [Guid("D8A5F198-6D37-401F-AAD8-59B49E84ECA7")]
    public interface IExportProgressAsStringChangedArgs
    {
        /// <summary>
        /// Rows processed.
        /// </summary>
        int RowsProcessed { get; }
        /// <summary>
        /// Total number of rows to process.
        /// </summary>
        int RowsTotal { get; }
        /// <summary>
        /// Rowdata in structured string representation.
        /// </summary>
        string RowValues { get; }
    }

    /// <summary>
    /// ExportProgressAsStringChangedArgs class.
    /// Reports export progress and data to outside the assembly.
    /// </summary>
    [ComVisible(true)]
    [Guid("40919A47-E618-4C5F-AF94-EAFCEA5B3F0D")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IExportProgressAsStringChangedArgs))]
    public class ExportProgressAsStringChangedArgs : EventArgs, IExportProgressAsStringChangedArgs
    {
        /// <summary>
        /// Constructor.
        /// </summary> 
        /// <param name="rowsProcessed">Current row.</param>
        /// <param name="totalRows">Total number of rows to read.</param>
        internal ExportProgressAsStringChangedArgs(int rowsProcessed, int totalRows)
        {
            this.RowsProcessed = rowsProcessed;
            this.RowsTotal = totalRows;
        }

        /// <summary>
        /// Overloaded constructor
        /// </summary>
        /// <param name="rowsProcessed">Current row.</param>
        /// <param name="totalRows">Total number of rows to read.</param>
        /// <param name="data">String</param>
        internal ExportProgressAsStringChangedArgs(int rowsProcessed, int totalRows, string data)
        {
            this.RowsProcessed = rowsProcessed;
            this.RowsTotal = totalRows;
            this.RowValues = data;
        }
        /// <inheritdoc />
        public int RowsProcessed { get; internal set; }
        /// <inheritdoc />
        public int RowsTotal { get; internal set; }
        /// <inheritdoc />
        public string RowValues { get; internal set; } = "Only populated when used with ExportSettings.ExportFormat.Event.";
    }
}

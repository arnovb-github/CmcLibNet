using System;
using System.Runtime.InteropServices;

namespace Vovin.CmcLibNet.Export
{
    /// <summary>
    /// Interface for ExportProgressChangedArgs for use with COM Interop.
    /// </summary>
    [ComVisible(true)]
    [GuidAttribute("0902E210-9D7D-4CFF-9CF6-0D402F63D304")]
    public interface IExportProgressChangedArgs
    {
        /// <summary>
        /// Rows processed.
        /// </summary>
        int RowsProcessed { get; }
        /// <summary>
        /// Total number of rows to process.
        /// </summary>
        int RowsTotal { get; }
    }

    /// <summary>
    /// Custom ExportProgressChangedArgs class that reports export progress.
    /// </summary>
    [ComVisible(true)]
    [GuidAttribute("4BCBA029-04FC-44CB-9A4D-2960AEC225CD")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IExportProgressChangedArgs))]
    public class ExportProgressChangedArgs : EventArgs, IExportProgressChangedArgs
    {
        /// <summary>
        /// Constructor.
        /// </summary>
        /// <param name="currentrow">Current row.</param>
        /// <param name="totalrows">Total number of rows to read.</param>
        internal ExportProgressChangedArgs(int currentrow, int totalrows)
        {
            this.RowsProcessed = currentrow;
            this.RowsTotal = totalrows;
        }
        /// <summary>
        /// Rows processed.
        /// </summary>
        public int RowsProcessed { get; private set; }
        /// <summary>
        /// Total number of rows to process.
        /// </summary>
        public int RowsTotal { get; private set; }
    }
}

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
}

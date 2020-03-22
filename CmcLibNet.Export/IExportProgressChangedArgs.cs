using System;
using System.Runtime.InteropServices;

namespace Vovin.CmcLibNet.Export
{
    /// <summary>
    /// Interface for ExportProgressChanged, primarily for use with COM Interop.
    /// </summary>
    [ComVisible(true)]
    [Guid("D8A5F198-6D37-401F-AAD8-59B49E84ECA7")]
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
        /// <summary>
        /// Rowdata in structured string representation.
        /// </summary>
        string RowValues { get; }
        /// <summary>
        /// Current iteration
        /// </summary>
        int CurrentIteration { get; }
        /// <summary>
        /// Total number of iterations
        /// </summary>
        /// <remarks>When a complex dataread with multiple cursors is performed,
        /// there are a number of iterations.</remarks>
        int IterationCount { get; }
    }
}

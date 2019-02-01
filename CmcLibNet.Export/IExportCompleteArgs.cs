using System;
using System.Runtime.InteropServices;

namespace Vovin.CmcLibNet.Export
{
    /// <summary>
    /// Interface for ExportCompleted, primarily for use with COM Interop.
    /// </summary>
    [ComVisible(true)]
    [Guid("8D262E04-9AE5-4C64-968F-7A7AEAA2B58C")]
    public interface IExportCompleteArgs
    {
        /// <summary>
        /// Rows processed.
        /// </summary>
        int Rows { get; }
    }
}

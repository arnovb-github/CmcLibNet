using System;
using System.Runtime.InteropServices;

namespace Vovin.CmcLibNet.Export
{
    /// <summary>
    /// Interface for DataRowsReadArgs for use with COM Interop.
    /// </summary>
    [ComVisible(true)]
    [GuidAttribute("D8A5F198-6D37-401F-AAD8-59B49E84ECA7")]
    public interface IDataRowsReadArgs
    {   /// <summary>
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
    /// Custom DataRowsReadArgs class that reports export progress and data.
    /// </summary>
    [ComVisible(true)]
    [GuidAttribute("40919A47-E618-4C5F-AF94-EAFCEA5B3F0D")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IDataRowsReadArgs))]
    public class DataRowsReadArgs : EventArgs, IDataRowsReadArgs
    {
        /// <summary>
        /// Constructor.
        /// </summary>
        /// <param name="currentrow">Current row.</param>
        /// <param name="totalrows">Total number of rows to read.</param>
        /// <param name="data">Data.</param>
        internal DataRowsReadArgs(int currentrow, int totalrows, string data)
        {
            this.RowsProcessed = currentrow;
            this.RowsTotal = totalrows;
            this.RowValues = data;
        }
        /// <inheritdoc />
        public int RowsProcessed { get; private set; }
        /// <inheritdoc />
        public int RowsTotal { get; private set; }
        /// <inheritdoc />
        public string RowValues { get; private set; }
    }
}

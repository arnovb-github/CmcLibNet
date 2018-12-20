using System;
using System.Runtime.InteropServices;

namespace Vovin.CmcLibNet.Export
{
    /// THIS CLASS IS SILLY
    /// it's goal it to report progress of a Commence data read
    /// but Commence always returns data in batches anyway
    /// That means that this event will be fired after the batch has been retrieved,
    /// defeating the purpose

    /// <summary>
    /// Interface for DataRowReadArgs for use with COM Interop.
    /// </summary>
    [ComVisible(true)]
    [GuidAttribute("0902E210-9D7D-4CFF-9CF6-0D402F63D304")]
    public interface IDataRowReadArgs
    {
        /// <summary>
        /// Rows processed.
        /// </summary>
        int CurrentRow { get; }
        /// <summary>
        /// Total number of rows to process.
        /// </summary>
        int RowsTotal { get; }
    }

    /// <summary>
    /// Custom DataRowReadArgs class that reports export progress.
    /// </summary>
    [ComVisible(true)]
    [GuidAttribute("4BCBA029-04FC-44CB-9A4D-2960AEC225CD")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IDataRowReadArgs))]
    public class DataRowReadArgs : EventArgs, IDataRowReadArgs
    {
        /// <summary>
        /// Constructor.
        /// </summary>
        /// <param name="currentrow">Current row.</param>
        /// <param name="totalrows">Total number of rows to read.</param>
        internal DataRowReadArgs(int currentrow, int totalrows)
        {
            this.CurrentRow = currentrow;
            this.RowsTotal = totalrows;
        }
        /// <summary>
        /// Rows processed.
        /// </summary>
        public int CurrentRow { get; private set; }
        /// <summary>
        /// Total number of rows to process.
        /// </summary>
        public int RowsTotal { get; private set; }
    }
}

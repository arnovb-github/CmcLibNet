using System;

namespace Vovin.CmcLibNet.Database
{
    /// <summary>
    /// EventArgs for cursor read progress tracking.
    /// </summary>
    public class CursorRowsReadArgs : EventArgs
    {
        internal CursorRowsReadArgs(int rowsProcessed, int rowsTotal)
        {
            RowsProcessed = rowsProcessed;
            RowsTotal = rowsTotal;
        }

        /// <summary>
        /// Rows processed
        /// </summary>
        public int RowsProcessed { get; }

        /// <summary>
        /// Total number of rows in cursor.
        /// </summary>
        public int RowsTotal { get; }
    }
}
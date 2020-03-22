using System;
using System.Collections.Generic;

namespace Vovin.CmcLibNet.Export
{
    /// <summary>
    /// Holds the values of an export event used internally
    /// </summary>
    internal class CursorDataReadProgressChangedArgs : EventArgs
    {
        internal CursorDataReadProgressChangedArgs(List<List<CommenceValue>> list, int rowsprocessed, int totalrows)
        {
            this.RowValues = list;
            this.RowsProcessed = rowsprocessed;
            this.RowsTotal = totalrows;
        }
        internal List<List<CommenceValue>> RowValues { get; private set; }
        internal int RowsProcessed { get; private set; }
        internal int RowsTotal { get; private set; }
    }
}

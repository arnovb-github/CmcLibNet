﻿using System.Runtime.InteropServices;

namespace Vovin.CmcLibNet.Database
{
    /// <summary>
    /// Interface for CommenceDeleteRowSet
    /// </summary>
    [ComVisible(true)]
    [Guid("72F4E7E6-2ED1-491D-851F-50EDA90CCF4D")]
    public interface ICommenceDeleteRowSet : IBaseRowSet
    {
        #region Redefined signatures from IBaseRowSet required for COM
        /// <inheritdoc />
        new int RowCount { get; }
        /// <inheritdoc />
        new int ColumnCount { get; }
        /// <inheritdoc />
        new string GetRowValue(int nRow, int nCol, CmcOptionFlags flags = CmcOptionFlags.Default);
        /// <inheritdoc />
        new string GetColumnLabel(int nCol, CmcOptionFlags flags = CmcOptionFlags.Default);
        /// <inheritdoc />
        new int GetColumnIndex(string pLabel, CmcOptionFlags flags = CmcOptionFlags.Default);
        /// <inheritdoc />
        new object[] GetRow(int nRow, CmcOptionFlags flags = CmcOptionFlags.Default);
        /// <inheritdoc />
        new object[] GetRow(int nRow, string delim, CmcOptionFlags flags = CmcOptionFlags.Default);
        /// <inheritdoc />
        new bool GetShared(int nRow);
        /// <inheritdoc />
        new void Close();
        #endregion
        /// <summary>
        /// Returns a unique identifier for a row.
        /// </summary>
        /// <param name="nRow">The (0-based) index of the row.</param>
        /// <param name="flags">Unused at present, must be 0.</param>
        /// <returns>Returns a unique ID string (less than 100 chars) on success, <c>null</c> on error.</returns>
        string GetRowID(int nRow, CmcOptionFlags flags = CmcOptionFlags.Default);
        /// <summary>
        /// Mark a row for deletion.
        /// </summary>
        /// <param name="nRow">The (0-based) index of the row.</param>
        /// /// <param name="flags">Unused, must be 0</param>
        /// <returns>0 on success, -1 on error.</returns>
        int DeleteRow(int nRow, CmcOptionFlags flags = CmcOptionFlags.Default);
        /// <summary>
        /// Commit changes to disk.
        /// </summary>
        /// <param name="flags">Unused, must be 0</param>
        /// <returns>0 on success, -1 on error.</returns>
        int Commit(CmcOptionFlags flags = CmcOptionFlags.Default);
    }
}

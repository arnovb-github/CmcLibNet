﻿using System.Runtime.InteropServices;

namespace Vovin.CmcLibNet.Database
{
    /// <summary>
    /// Interface for CommenceQueryRowSet
    /// </summary>
    [ComVisible(true)]
    [Guid("F083B422-05FC-4C23-A5DE-C4528DDC477A")]
    public interface ICommenceQueryRowSet : IBaseRowSet
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
        /// GetRow implementation for .NET that gets the the Commence row as string values.
        /// </summary>
        /// <param name="nRow">Row pointer.</param>
        /// <returns>String array.</returns>
        /// <remarks>The purpose of this method is primarily to avoid boxing in internal export methods.</remarks>
        [ComVisible(false)]
        string[] GetRow(int nRow);
        /// <summary>
        /// Save the field value at the given (row,column) to a file.
        /// </summary>
        /// <param name="nRow">The (0-based) index of the row.</param>
        /// <param name="nCol">The (0-based) index of the column.</param>
        /// <param name="filename">(Fully qualified) Filename where the field value is written.</param>
        /// <param name="flags">Logical OR of following option flags:
        /// <see cref="CmcOptionFlags.Canonical"/> - return field value in canonical form.</param>
        /// <returns>File size in bytes, -1 on error.</returns>
        int GetFieldToFile(int nRow, int nCol, string filename, CmcOptionFlags flags = CmcOptionFlags.Default);
        /// <summary>
        /// Returns a unique identifier for a row.
        /// </summary>
        /// <param name="nRow">The (0-based) index of the row.</param>
        /// <param name="flags">Unused at present, must be 0.</param>
        /// <returns>Returns a unique ID string (less than 100 chars) on success, <c>null</c> on error.</returns>
        string GetRowID(int nRow, CmcOptionFlags flags = CmcOptionFlags.Default);
        /// <summary>
        /// Undocumented method, most likely returns a timestamp :).
        /// </summary>
        /// <param name="nRow">The (0-based) index of the row.</param>
        /// <param name="flags">(Assumed) Logical OR of following option flags:
        /// CmcOptionFlags.Canonical - return field value in canonical form.</param>
        /// <returns>some string with a timestamp? Who knows? Is it blue?</returns>
        string GetRowTimeStamp(int nRow, CmcOptionFlags flags = CmcOptionFlags.Default); // undocumented by Commence
        /// <summary>
        /// Gets the sequence number that Commence uses internally. Note: this is not a workgroup-wide number!
        /// </summary>
        /// <param name="nRow">The (0-based) index of the row.</param>
        /// <returns>Sequence number that Commence uses internally.</returns>
        /// <remarks>Local categories do not have a THID. Local items in shared categories do have THIDs.
        /// If no there are THIDs this method will use the RowID instead.</remarks>
        int GetRowSequenceNumber(int nRow);
    }
}
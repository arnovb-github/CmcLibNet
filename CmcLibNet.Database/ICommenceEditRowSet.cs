using System.Runtime.InteropServices;
using Vovin.CmcLibNet;

namespace Vovin.CmcLibNet.Database
{
    /// <summary>
    /// interface for CommenceEditRowSet
    /// </summary>
    [ComVisible(true)]
    [Guid("24AED779-089D-4E9A-B143-0EE9BAC94E4E")]
    public interface ICommenceEditRowSet : IBaseRowSet
    {
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
        /// <summary>
        /// Returns a unique identifier for a row.
        /// </summary>
        /// <param name="nRow">The (0-based) index of the row.</param>
        /// <param name="flags">Unused at present, must be 0.</param>
        /// <returns>Returns a unique ID string (less than 100 chars) on success, <c>null</c> on error.</returns>
        string GetRowID(int nRow, CmcOptionFlags flags = CmcOptionFlags.Default);
        /// <inheritdoc />>
        new bool GetShared(int nRow);
        /// <inheritdoc />
        new void Close();
        /// <summary>
        /// Sets the shared status of the row. Shared items cannot be made unshared.
        /// </summary>
        /// <param name="nRow">The (0-based) index of the row.</param>
        /// <returns>True on success.</returns>
        bool SetShared(int nRow);
        /// <summary>
        /// Modify fieldvalue
        /// </summary>
        /// <param name="nRow">The (0-based) index of the row.</param>
        /// <param name="nCol">The (0-based) index of the column.</param>
        /// <param name="fieldValue">new fieldValue.</param>
        /// <param name="flags">Unused, must be 0</param>
        /// <returns>0 on success, -1 on error.</returns>
        int ModifyRow(int nRow, int nCol, string fieldValue, CmcOptionFlags flags = CmcOptionFlags.Default);
        /// <summary>
        /// Commit changes to disk.
        /// </summary>
        /// <param name="flags">Unused, must be 0</param>
        /// <returns>0 on success, -1 on error.</returns>
        int Commit(CmcOptionFlags flags = CmcOptionFlags.Default);
        /// <summary>
        /// Make row modifications permanent (commit to disk) and return a cursor for the newly added data
        /// </summary>
        /// <param name="flags">Unused, must be 0</param>
        /// <returns>CommenceCursor object reflecting the changes.</returns>
        CommenceCursor CommitGetCursor(CmcOptionFlags flags = CmcOptionFlags.Default);
    }
}

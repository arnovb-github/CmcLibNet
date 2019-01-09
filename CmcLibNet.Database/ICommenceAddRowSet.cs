using System.Runtime.InteropServices;
using Vovin.CmcLibNet;

namespace Vovin.CmcLibNet.Database
{
    /// <summary>
    /// Interface for CommenceAddRowSet.
    /// </summary>
    [ComVisible(true)]
    [Guid("68063DB8-A8C4-49F3-AFD7-4E8F7ED1A426")]
    public interface ICommenceAddRowSet : IBaseRowSet
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
        /// <inheritdoc />
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
        /// Make row modifications permanent (commit to disk) and return a cursor for the newly added data.
        /// The CommenceAddRowSet should now be discarded.
        /// </summary>
        /// <param name="flags">Unused, must be 0</param>
        /// <returns>CommenceCursor object reflecting the changes.</returns>
        CommenceCursor CommitGetCursor(CmcOptionFlags flags = CmcOptionFlags.Default);
    }
}

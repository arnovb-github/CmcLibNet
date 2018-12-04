using System;
using Vovin.CmcLibNet;

namespace Vovin.CmcLibNet.Database
{
    /// <summary>
    /// Exposes methods and properties shared by all rowset types.
    /// </summary>
    public interface IBaseRowSet : IDisposable
    {
        /// <summary>
        /// Number of items in rowset.
        /// </summary>
        int RowCount { get; }
        /// <summary>
        /// Number of columns in rowset
        /// </summary>
        int ColumnCount { get; }
        /// <summary>
        /// Get rowvalue for specified row/column
        /// </summary>
        /// <param name="nRow">The (0-based) index of the row</param>
        /// <param name="nCol">The (0-based) index of the column.</param>
        /// <param name="flags">Logical OR of following option flags:
        /// CmcOptionFlags.Canonical - return field value in canonical form</param>
        /// <returns>rowvalue, <c>null</c> on error.</returns>
        string GetRowValue(int nRow, int nCol, CmcOptionFlags flags=CmcOptionFlags.Default);
        /// <summary>
        /// Search the column set and return the column label or fieldname
        /// </summary>
        /// <param name="nCol">The (0-based) index of the column.</param>
        /// <param name="flags">Logical OR of following option flags:
        /// CmcOptionFlags.Fieldname - return field label (ignore view labels)</param>
        /// <returns>Column label on success, -1 on error.</returns>
        string GetColumnLabel(int nCol, CmcOptionFlags flags = CmcOptionFlags.Default);
        /// <summary>
        /// Search the column set and return the index of the column with the given label
        /// </summary>
        /// <param name="pLabel">The column label to map.</param>
        /// <param name="flags">Logical OR of following option flags:
        /// CmcOptionFlags.Fieldname - return field label (ignore view labels).</param>
        /// <returns>0-based column index on success, -1 on error.</returns>
        int GetColumnIndex(string pLabel, CmcOptionFlags flags = CmcOptionFlags.Default);
        /// <summary>
        /// Returns an entire row's field values. Note that this method returns an array, whereas the native Commence method returns a delimited string.
        /// </summary>
        /// <param name="nRow">The (0-based) index of the row.</param>
        /// <param name="flags">Logical OR of following option flags:
        /// CmcOptionFlags.Canonical - return field value in canonical form</param>
        /// <returns>Object array of strings containing the row's values, <c>null</c> on error.</returns>
		/// <remarks>The return type is an object array, not a string array, even though Commence always returns values as string.
		/// That may seem counterintuitive. The reason for this is COM. I have not found a way to marshal a string array as variant.
		/// It is possible to marshal a string[], but accessing it will raise a type mismatch error in VBScript.
        /// </remarks>
        object[] GetRow(int nRow, CmcOptionFlags flags = CmcOptionFlags.Default);
        /// <summary>
        /// Gets the shared/local status of the row
        /// </summary>
        /// <param name="nRow">The (0-based) index of the row.</param>
        /// <returns>true if shared.</returns>
        bool GetShared(int nRow);
        /// <summary>
        /// Close any references to Commence. The object should be disposed after this.
        /// </summary>
        /// <remarks>When used from within a Commence Form Script, failing to call the <c>Close</c> method will leave the commence.exe process running in the background when the user closes Commence. IMPORTANT: this also happens when an unhandled exception (a 'script error') occurs. The Commence process then has to be closed manually from the Windows Task Manager. Be careful to implement proper error handling.
        /// <para>When the assembly is called from a.NET application, there is rarely a need to call this method, unless you want to explicitly release COM references and/or release memory. It can be useful in some cases, because Commence may complain about running out of memory before the Garbage Collector has a chance to kick in.</para>
        /// <para>Technical details: calling this method tells the assembly to release all COM handles (called 'RCW' for 'runtime callable wrapper') to Commence that are open. This is needed because when the object reference to this assembly is set to Nothing (in VB), the .NET assembly may not be notified and will think they are still in use. Garbage Collection will therefore not release them, and the commence.exe process will not be terminated.</para>
        /// </remarks>
        void Close();
    }
}

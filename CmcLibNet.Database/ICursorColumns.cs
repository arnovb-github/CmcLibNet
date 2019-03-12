using System.Runtime.InteropServices;

namespace Vovin.CmcLibNet.Database
{
    /// <summary>
    /// Interface for CursorColumns, allows for a convenient way of setting direct and related columns.
    /// </summary>
    [ComVisible(true)]
    [Guid("72AC55CA-BCB6-4B96-95CA-78A5BE52441A")]
    public interface ICursorColumns
    {
        /// <summary>
        /// Adds single fieldname to the collection of direct fields.
        /// </summary>
        /// <param name="fieldName">Direct fieldname from the Cursor's category.</param>
        void AddDirectColumn(string fieldName);
        /// <summary>
        /// Adds an array of direct fields to the collection of direct fields.
        /// </summary>
        /// <param name="fieldNames">Object array of fieldnames.</param>
        void AddDirectColumns(params object[] fieldNames); // we cannot get away with Marshaling here; I could find no proper way to marshal string[]. COM insists on Variant.
        /// <summary>
        /// Adds an array of direct fields to the collection of direct fields.
        /// </summary>
        /// <param name="fieldNames">Array of fieldnames, or a comma-separated list of fieldnames.</param>
        /// <remarks>This method is only available for .NET applications.</remarks>
        [ComVisible(false)]
        void AddDirectColumns(params string[] fieldNames);
        /// <summary>
        /// Add a related field to the collection of related fields.
        /// </summary>
        /// <param name="connectionName">Commence connection name (case-sensitive!).</param>
        /// <param name="categoryName">Commence connected category name.</param>
        /// <param name="fieldName">Commence connected category fieldname.</param>
        void AddRelatedColumn(string connectionName, string categoryName, string fieldName);
        /// <summary>
        /// Set all direct and related columns. Duplicate columns are ignored.
        /// </summary>
        /// <returns>Number of columns set, -1 on error.</returns>
        int Apply();
    }
}

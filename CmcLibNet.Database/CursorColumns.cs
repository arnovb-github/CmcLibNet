using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Linq;

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
    /// <summary>
    /// Convenience class for setting multiple columns in a cursor. Should not be combined with the direct methods of setting columns in <see cref="ICommenceCursor"/>.
    /// </summary>
    /// <remarks>Use of this class is recommended over setting columns directly. You cannot instantiate this class, it is automatically created when you call the <see cref="ICommenceCursor.Columns"/> property.</remarks>
    [ComVisible(true)]
    [Guid("48AA02B3-4616-4C2D-B273-096F1E7C5F00")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(CmcLibNet.Database.ICursorColumns))]
    public class CursorColumns : ICursorColumns
    {
        private List<string> _columns = null;
        private List<IRelatedColumn> _relatedcolumns = null;
        private CommenceCursor _cur = null; // note that we do not use the interface, but the direct type.

        #region Constructors
        internal CursorColumns(CommenceCursor cur)
        {
            _cur = cur;
            _columns = new List<string>();
            _relatedcolumns = new List<IRelatedColumn>();
        }
        #endregion

        #region Methods

        /// <inheritdoc />
        public void AddDirectColumn(string fieldName)
        {
            _columns.Add(fieldName);
        }

        /// <inheritdoc />
        public void AddDirectColumns(params object[] fieldNames)
        {
            foreach (object o in fieldNames)
            {
                _columns.Add(o.ToString());
            }
        }
        /// <inheritdoc />
        public void AddDirectColumns(params string[] fieldNames)
        {
            _columns.AddRange(fieldNames);
        }
        /// <inheritdoc />
        public void AddRelatedColumn(string connectionName, string categoryName, string fieldName)
        {
            IRelatedColumn rc = new RelatedColumn(connectionName, categoryName, fieldName, RelatedColumnType.ConnectedField, null);
            _relatedcolumns.Add(rc);
        }

        /// <inheritdoc />
        public int Apply()
        {
            int retval = -1;

            try
            {
                // set the columns, ignoring duplicates
                _cur.SetColumns(_columns.Distinct().ToArray());
            }
            catch (Vovin.CmcLibNet.CommenceCOMException)
            {
                throw;
            }

            try
            {
                // set the related columns, ignoring duplicates
                 _cur.SetRelatedColumns(_relatedcolumns.Distinct().ToList());
                // System.Console.WriteLine(_relatedcolumns.Distinct().ToList().Count);
            }
            catch (Vovin.CmcLibNet.CommenceCOMException)
            {
                throw;
            }
            retval = _cur.ColumnCount;
            return retval;
        }
        #endregion
    }
}

using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Linq;

namespace Vovin.CmcLibNet.Database
{
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
            catch (CommenceCOMException)
            {
                throw;
            }

            try
            {
                // set the related columns, ignoring duplicates
                 _cur.SetRelatedColumns(_relatedcolumns.Distinct().ToList());
            }
            catch (CommenceCOMException)
            {
                throw;
            }
            retval = _cur.ColumnCount;
            return retval;
        }
        #endregion
    }
}

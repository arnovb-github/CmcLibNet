using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Vovin.CmcLibNet.Attributes;

namespace Vovin.CmcLibNet.Database
{
    /// <summary>
    /// Wrapper around the Commence FormOA.ICommenceCursor interface.
    /// </summary>
    /// 
    [ComVisible(true)]
    [Guid("E4970967-1D0D-41a8-81A1-7863A94EEE9C")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(ICommenceCursor))]
    public class CommenceCursor : ICommenceCursor
    {
        /* TODO rethink the setcolumns stuff, those methods are a mess */

        private FormOA.ICommenceCursor _cur = null; // 'raw' ICommenceCursor object; a COM object
        private IRcwReleasePublisher _rcwReleasePublisher = null;
        private CursorFilters _filters = null;
        internal bool _directColumnsWereSet = false;
        internal bool _relatedColumnsWereSet = false;
        private CursorColumns _columns = null;
        internal string _viewName = string.Empty;
        internal CommenceViewType _viewType;
        bool disposed = false;

        #region Constructors
        /// <summary>
        /// Constructor that takes the raw FormOA.ICommenceCursor as parameter
        /// For use with the *RowSet CommitGetCursor methods
        /// </summary>
        /// <param name="cur">FormOA.ICommenceCursor.</param>
        /// <param name="rcwReleasePublisher">RCWReleasePublisher object used for COM Interop object cleanup.</param>
        internal CommenceCursor(FormOA.ICommenceCursor cur, IRcwReleasePublisher rcwReleasePublisher)
        {
            _cur = cur; // can be used with CommitGetCursor, or use SetCursor to re-use this object
            _rcwReleasePublisher = rcwReleasePublisher;
            _rcwReleasePublisher.RCWRelease += this.RCWReleaseHandler;
        }
        /// <summary>
        /// Creates default cursor type on category with default options.
        /// </summary>
        /// <param name="cmc">Native FormOA (Commence) reference.</param>
        /// <param name="pName">Commence category name.</param>
        /// <param name="rcwReleasePublisher">RCWReleasePublisher object used for COM Interop object cleanup.</param>
        internal CommenceCursor(FormOA.ICommenceDB cmc, string pName, IRcwReleasePublisher rcwReleasePublisher)
        {
            // default cursor type, on category, default flag
             _cur = cmc.GetCursor(0, pName, 0);
            _rcwReleasePublisher = rcwReleasePublisher;
             _rcwReleasePublisher.RCWRelease += this.RCWReleaseHandler;
        }
        /// <summary>
        /// Creates custom View cursor.
        /// </summary>
        /// <param name="cmc">Native FormOA (Commence) reference.</param>
        /// <param name="pCursorType">CmcCursorType.</param>
        /// <param name="pName">Commence category or view name.</param>
        /// <param name="rcwReleasePublisher">RCWReleasePublisher object used for COM Interop object cleanup.</param>
        /// <param name="pCursorFlags">CmcOptionFlags.</param>
        /// <param name="viewType">Viewtype.</param>
        internal CommenceCursor(FormOA.ICommenceDB cmc, CmcCursorType pCursorType, string pName, IRcwReleasePublisher rcwReleasePublisher, CmcOptionFlags pCursorFlags, string viewType)
        {
            CursorType = pCursorType;
            Flags = pCursorFlags;
            if (CursorType == CmcCursorType.View) 
            {
                _viewName = pName;
                //_viewType = Utils.GetValueFromEnumDescription<CommenceViewType>(viewType);
                _viewType = Utils.EnumFromAttributeValue<CommenceViewType, StringValueAttribute>(nameof(StringValueAttribute.StringValue),viewType);
            }
            _cur = cmc.GetCursor((int)pCursorType, pName, (int)pCursorFlags); // notice the type conversion
            _rcwReleasePublisher = rcwReleasePublisher;
            _rcwReleasePublisher.RCWRelease += this.RCWReleaseHandler;
        }

        /// <summary>
        /// Creates custom Category cursor.
        /// </summary>
        /// <param name="cmc">Native FormOA (Commence) reference.</param>
        /// <param name="pCursorType">CmcCursorType.</param>
        /// <param name="pName">Commence category or view name.</param>
        /// <param name="rcwReleasePublisher">RCWReleasePublisher object used for COM Interop object cleanup.</param>
        /// <param name="pCursorFlags">CmcOptionFlags.</param>
        internal CommenceCursor(FormOA.ICommenceDB cmc, CmcCursorType pCursorType, string pName, IRcwReleasePublisher rcwReleasePublisher, CmcOptionFlags pCursorFlags)
        {
            CursorType = pCursorType;
            Flags = pCursorFlags;
            _cur = cmc.GetCursor((int)pCursorType, pName, (int)pCursorFlags); // notice the type conversion
            _rcwReleasePublisher = rcwReleasePublisher;
            _rcwReleasePublisher.RCWRelease += this.RCWReleaseHandler;
        }

        /// <summary>
        /// Finalizer
        /// </summary>
        ~CommenceCursor()
        {
            Dispose(false);
        }

        #endregion

        #region Properties
        /// <summary>
        /// Allows for explicit setting of the ICommenceCursor object, normally done in the constructor.
        /// For use with CommitGetCursor method of *RowSet objects.
        /// </summary>
        internal FormOA.ICommenceCursor SetCursor
        {
            set { if (value != null) { _cur = value; } }
        }
        /// <inheritdoc />
        public int RowCount
        {
            get { return _cur.RowCount; }
        }

        /// <inheritdoc />
        public string Category
        {
            get { return _cur.Category; }
        }

        /// <inheritdoc />
        public string View
        {
            get { return _viewName; }
        }

        /// <inheritdoc />
        public int ColumnCount
        {
            get { return _cur.ColumnCount; }
        }

        /// <inheritdoc />
        public ICursorFilters Filters
        {
            get
            {
                if (_filters == null)
                {
                    _filters = new CursorFilters(this);
                }
                return _filters;
            }
        }

        /// <inheritdoc />
        public int MaxFieldSize
        {
            get
            {
                return _cur.MaxFieldSize;
            }
            set
            {
                _cur.MaxFieldSize = value;
            }
        }

        /// <inheritdoc />
        public int MaxRows
        {
            get
            {
                return _cur.MaxRows;
            }
        }

        /// <inheritdoc />
        public bool Shared
        {
            get { return _cur.Shared; }
        }

        /// <summary>
        /// Returns the cursortype.
        /// </summary>
        internal CmcCursorType CursorType { get; } = CmcCursorType.Category;

        /// <inheritdoc />
        public ICursorColumns Columns
        {
            get
            {
                if (_columns == null)
                {
                    _columns = new CursorColumns(this);
                }
                return _columns;
            }
        }

        internal CommenceViewType ViewType
        {
            get { return _viewType; }
        }

        internal CmcOptionFlags Flags { get; } = CmcOptionFlags.Default;
        #endregion

        #region Methods

        /// <inheritdoc />
        public ICommenceAddRowSet GetAddRowSet(int nRows, CmcOptionFlags flags = CmcOptionFlags.Default)
        {
            ICommenceAddRowSet ars = null;
            try
            {
                ars = new CommenceAddRowSet(_cur, nRows, _rcwReleasePublisher, flags);
            }
            catch (COMException e)
            {
                throw new CommenceCOMException("Unable to get a AddRowSet object from Commence", e);
            }
            return ars;
        }

        /// <inheritdoc />
        public ICommenceDeleteRowSet GetDeleteRowSet(int nRows, CmcOptionFlags flags = CmcOptionFlags.Default)
        {
            ICommenceDeleteRowSet drs = null;
            try
            {
                drs = new CommenceDeleteRowSet(_cur, nRows, _rcwReleasePublisher, flags);
            }
            catch (COMException e)
            {
                throw new CommenceCOMException("Unable to get a DeleteRowSet object from Commence", e);
            }
            return drs;
        }

        /// <inheritdoc />
        public ICommenceDeleteRowSet GetDeleteRowSet(CmcOptionFlags flags = CmcOptionFlags.Default)
        {
            ICommenceDeleteRowSet drs = null;
            try
            {
                drs = new CommenceDeleteRowSet(_cur, _rcwReleasePublisher,flags);
            }
            catch (COMException e)
            {
                throw new CommenceCOMException("Unable to get a DeleteRowSet object from Commence", e);
            }
            return drs;
        }
        /// <inheritdoc />
        public ICommenceDeleteRowSet GetDeleteRowSetByID(string pRowID, CmcOptionFlags flags = CmcOptionFlags.Default)
        {
            ICommenceDeleteRowSet drs = null;
            try
            {
                drs = new CommenceDeleteRowSet(_cur, pRowID, _rcwReleasePublisher,flags);
            }
            catch (COMException e)
            {
                throw new CommenceCOMException("Unable to get a DeleteRowSet object from Commence", e);
            }
            return drs;
        }
        /// <inheritdoc />
        public ICommenceEditRowSet GetEditRowSet(int nRows, CmcOptionFlags flags = CmcOptionFlags.Default)
        {
            {
                ICommenceEditRowSet ers = null;
                try
                {
                    ers = new CommenceEditRowSet(_cur, nRows, _rcwReleasePublisher,flags);
                }
                catch (COMException e)
                {
                    throw new CommenceCOMException("Unable to get a EditRowSet object from Commence", e);
                }
                return ers;
            }
        }
        /// <inheritdoc />
        public ICommenceEditRowSet GetEditRowSet(CmcOptionFlags flags = CmcOptionFlags.Default)
        {
            {
                ICommenceEditRowSet ers = null;
                try
                {
                   ers = new CommenceEditRowSet(_cur, _rcwReleasePublisher, flags);
                }
                catch (COMException e)
                {
                    throw new CommenceCOMException("Unable to get a EditRowSet object from Commence", e);
                }
                return ers;
            }
        }
        /// <inheritdoc />
        public ICommenceEditRowSet GetEditRowSetByID(string pRowID, CmcOptionFlags flags = CmcOptionFlags.Default)
        {
            ICommenceEditRowSet ers = null;
            try
            {
                ers = new CommenceEditRowSet(_cur, pRowID, _rcwReleasePublisher,flags);
            }
            catch (COMException e)
            {
                throw new CommenceCOMException("Unable to get a EditRowSet object from Commence", e);
            }
            return ers;
        }
        /// <inheritdoc />
        public ICommenceQueryRowSet GetQueryRowSet(CmcOptionFlags flags = CmcOptionFlags.Default)
        {
            ICommenceQueryRowSet qrs = null;
            try
            {
                qrs = new CommenceQueryRowSet(_cur, _rcwReleasePublisher, flags);
            }
            catch (COMException e)
            {
                throw new CommenceCOMException("Unable to get a QueryRowSet object from Commence", e);
            }
            return qrs;
        }
        /// <inheritdoc />
        public ICommenceQueryRowSet GetQueryRowSet(int nRows, CmcOptionFlags flags = CmcOptionFlags.Default)
        {
            ICommenceQueryRowSet qrs = null;
            try
            {
                qrs =  new CommenceQueryRowSet(_cur, nRows, _rcwReleasePublisher, flags);
            }
            catch (COMException e)
            {
                // when very large amounts of data are requested, this may cause a "Couldn't get memory" error in Commence.
                throw new CommenceCOMException("Unable to get a QueryRowSet object from Commence", e);
            }
            return qrs;
        }
        /// <inheritdoc />
        public ICommenceQueryRowSet GetQueryRowSetByID(string pRowID, CmcOptionFlags flags = CmcOptionFlags.Default)
        {
            ICommenceQueryRowSet qrs = null;
            try
            {
                qrs = new CommenceQueryRowSet(_cur, pRowID, _rcwReleasePublisher, flags);
            }
            catch (COMException e)
            {
                throw new CommenceCOMException("Unable to get a QueryRowSet object from Commence", e);
            }
            return qrs;
        }

        /// <inheritdoc />
        // only works on shared databases; local databases have no thids
        // you cannot just pass a thid value you retrieved by GetRowID(), you must strip off the category (=first) sequence
        // e.g. if GetRowID returns 0C:80006901:94BD3402, you must strip off the leading 0C: part
        public ICommenceQueryRowSet GetQueryRowSetByThid(string pThid, CmcOptionFlags flags = CmcOptionFlags.UseThids)
        {
            ICommenceQueryRowSet qrs = null;
            try
            {
                qrs = new CommenceQueryRowSet(_cur, pThid, _rcwReleasePublisher, flags, RowSetIdentifier.Thid);
            }
            catch (COMException e)
            {
                throw new CommenceCOMException("Unable to get a QueryRowSet object from Commence", e);
            }
            return qrs;
        }

        /// <inheritdoc />
        public int SeekRow(CmcCursorBookmark bkOrigin, int nRows)
        {
            return _cur.SeekRow((int)bkOrigin, nRows);
        }

        /// <inheritdoc />
        public int SeekRowApprox(int nNumerator, int nDenominator)
        {
            return _cur.SeekRowApprox(nNumerator, nDenominator);
        }

        /// <inheritdoc />
        public bool SetActiveDate(string sDate, CmcOptionFlags flags = CmcOptionFlags.Default)
        {
            try
            {
                return _cur.SetActiveDate(sDate, (int)flags);
            }
            catch (System.Runtime.Remoting.RemotingException)
            {
                return false;
            }
        }

        /// <inheritdoc />
        public bool SetActiveDateRange(string startDate, string endDate, CmcOptionFlags flags = CmcOptionFlags.Default)
        {
            return _cur.SetActiveDateRange(startDate, endDate, (int)flags);
        }

        /// <inheritdoc />
        public bool SetActiveItem(string pCategoryName, string pRowID, CmcOptionFlags flags = CmcOptionFlags.Default)
        {
            try
            {
                return _cur.SetActiveItem(pCategoryName, pRowID, (int)flags);
            }
            catch (System.Runtime.Remoting.RemotingException)
            {
                return false;
            }
        }

        /// <inheritdoc />
        public bool SetColumn(int nColumn, string pName, CmcOptionFlags flags)
        {
            try
            {
                _directColumnsWereSet = true;
                return _cur.SetColumn(nColumn, pName, (int)flags);
            }
            catch (System.Runtime.Remoting.RemotingException)
            {
                return false;
            }
        }

        /// <inheritdoc />
        public bool SetColumns(object[] columnNames)
        {
            return SetColumns(Utils.ToStringArray(columnNames));
        }

        /// <summary>
        /// Set a number of columns, in the order supplied.
        /// </summary>
        /// <param name="columnNames">Array of direct columnames.</param>
        /// <returns><c>true</c> on succes, <c>false</c> on error.</returns>
        /// <remarks>This method cannot be used to set related columns.</remarks>
        internal bool SetColumns(params string[] columnNames)
        {
            int col = -1;

            if (columnNames == null) { return false; }

            /* Commence requires that direct columns be defined first,
            * then the related columns. We use flags to check what we set previously.
            */

            // related columns were already set
            if (_relatedColumnsWereSet)
            {
                throw new InvalidOperationException("Related columns have been defined. You cannot add direct columns after related columns were set.");
            }

            if (_directColumnsWereSet && !_relatedColumnsWereSet)
            {
                col = this.ColumnCount - 1;
            }
            // no columns were set at all
            if (!_directColumnsWereSet && !_relatedColumnsWereSet)
            {
                col = -1;
            }

            for (int i = 0; i < columnNames.Length; i++)
            {
                bool result = this.SetColumn(col + 1, columnNames[i], (int)CmcOptionFlags.Default);
                if (!result)
                {
                    throw new CommenceCOMException("An error occurred while trying to set direct column:\n" + columnNames[i] + " at position " + (col + 1).ToString());
                }
                col++;
            }
            _directColumnsWereSet = true;
            return true;
        }

        /// <inheritdoc />
        public bool SetRelatedColumn(int nColumn, string pConnName, string pCatName, string pName, CmcOptionFlags flags = CmcOptionFlags.Default)
        {
            try
            {
                _relatedColumnsWereSet = true; // logic error!!
                return _cur.SetRelatedColumn(nColumn, pConnName, pCatName, pName, (int)flags);
            }
            catch (System.Runtime.Remoting.RemotingException)
            {
                return false;
            }
        }
        /// <summary>
        /// Sets related columns.
        /// </summary>
        /// <param name="relatedColumns">List of <c>IRelatedColumn</c> objects</param>
        /// <returns><c>true on success</c>, <c>false</c> on error.</returns>
        internal bool SetRelatedColumns(List<IRelatedColumn> relatedColumns)
        {
            int col = -1;
            if (relatedColumns == null) { return false; }

            // a default cursor has all columns, and you cannot append to it.
            // therefore, we have to check if any columns were set previously.

            // No columns were explicitly set
            // TODO: if no columns were explicitly set, a default cursor still holds all fields.
            if (!_directColumnsWereSet && !_relatedColumnsWereSet)
            {
                col = -1; // -1 because we 1 is added and we want to start at 0.
            }

            // No direct columns were set, but we did add related columns
            if (!_directColumnsWereSet && _relatedColumnsWereSet)
            {
                col = this.ColumnCount -1;
            }

            // direct columns were set, but no related columns
            if (_directColumnsWereSet && !_relatedColumnsWereSet)
            {
                col = this.ColumnCount - 1;
            }

            for (int i = 0; i < relatedColumns.Count; i++)
            {
                bool result = this.SetRelatedColumn(col + 1, relatedColumns[i].Connection, relatedColumns[i].Category, relatedColumns[i].Field, CmcOptionFlags.Default);
                if (result == false)
                {
                    throw new CommenceCOMException("An error occurred while trying to set related column:\n" + relatedColumns[i].Connection + ", " + relatedColumns[i].Category + ", " + relatedColumns[i].Field);
                }
                col++;
            }
            _relatedColumnsWereSet = true;
            return true;
        }

        /// <inheritdoc />
        public bool SetFilter(string pFilter, CmcOptionFlags flags)
        {
            try
            {
                return _cur.SetFilter(pFilter, (int)flags);
            }
            /* Contrary to Commence documentation this method will not return false if it fails.
             * Instead, a System.Runtime.Remoting.RemotingException is thrown complaining about
             * 'ByRef value type parameter cannot be null.'
             */
            catch (System.Runtime.Remoting.RemotingException)
            {
                return false;
            }
        }

        ///// <inheritdoc />
        //[ObsoleteAttribute("Use the Filters collection.")]
        //// Note the int type flags parameter, it's only for backwards compatibility with scripting.
        //public bool SetFilter(string pFilter, int flags)
        //{
        //    return SetFilter(pFilter, (CmcOptionFlags)flags);
        //}

        /// <inheritdoc />
        public bool SetLogic(string pLogic, CmcOptionFlags flags)
        {
            try
            {
                return _cur.SetLogic(pLogic, (int)flags);
            }
            catch (System.Runtime.Remoting.RemotingException)
            {
                return false;
            }
        }

        ///// <inheritdoc />
        //[ObsoleteAttribute("Use the Filters collection.")]
        //public bool SetLogic(string pLogic, int flags) // used from COM
        //{
        //    return SetLogic(pLogic, (CmcOptionFlags)flags);
        //}

        /// <inheritdoc />
        public bool SetSort(string pSort, CmcOptionFlags flags)
        {
            try
            {
                return _cur.SetSort(pSort, (int)flags);
            }
            catch (System.Runtime.Remoting.RemotingException)
            {
                return false;
            }
        }
        /// <inheritdoc />
        public bool HasDuplicates(string columnName, bool caseSensitive = true)
        {
            if (this.RowCount == 0) { return false; } // nothing to compare

            // we could just filter the cursor, but that is tricky for a couple of reasons:
            // - cursor may already be filtered, adding a new is not straightforward due to number and logic.
            // - filtering would require we obtain the fieldtype, because different fieldtypes take different qualifiers.
            // therefore we will retrieve all values and compare them
            // this is a little slower of course.
            List<string> values = new List<string>();
            List<string> distinctvalues = null;
            // capture current row by moving rowpointer to start.
            // currentrow will contain the number of rows moved back (for instance -789)
            int currentrow = this.SeekRow(CmcCursorBookmark.Beginning, 0);
            int colindex = -1;
            for (int i = 0; i < this.RowCount; i += 100) // process 100 rows at a time. A row can hold 250 columns of 30.000 characters each; 100 is an arbitrary amount.
            {
                using (ICommenceQueryRowSet qrs = this.GetQueryRowSet(100))
                {
                    if (colindex == -1) { colindex = qrs.GetColumnIndex(columnName); }
                    if (colindex == -1)
                    {
                        throw new CommenceCOMException("Column '" + columnName + "' is not included in cursor."); 
                    }
                    for (int j = 0; j < qrs.RowCount; j++)
                    {
                        values.Add(qrs.GetRow(j)[colindex].ToString());
                    }
                }
            }
            // put rowpointer back where it started
            this.SeekRow(CmcCursorBookmark.Beginning, Math.Abs(currentrow)); // reset rowpointer
            distinctvalues = (caseSensitive) ? values.Distinct().ToList() : values.Distinct(StringComparer.CurrentCultureIgnoreCase).ToList();
            return (values.Count() != distinctvalues.Count());
        }

        /// <inheritdoc />
        // the marshaling is needed to allow the parameter in, without it we get an argument exception.
        //public void CursorToFile(string fileName, [MarshalAs(UnmanagedType.IDispatch)] Export.IExportSettings settings = null) // tools like PowerShell reference the class, not the interface, so any optional parameters defined in the interface must be optional in the class as well.
        // fuck marshaling
        public void ExportToFile(string fileName, Export.IExportSettings settings = null) // tools like PowerShell reference the class, not the interface, so any optional parameters defined in the interface must be optional in the class as well.
        {
            Export.ExportEngine exportEngine = new Export.ExportEngine();
            if (settings != null) { exportEngine.Settings = settings; } // store custom settings
            // override setting, we would lose filter (if any) on the cursor if it was on a category
            if (this.CursorType == CmcCursorType.Category)
            {
                exportEngine.Settings.PreserveAllConnections = false;
            }
            exportEngine.ExportCursor(this, fileName, exportEngine.Settings);
        }

        /// <inheritdoc />
        public List<string> ReadRow(int lRow)
        {
            List<string> retval = null;
            if (this.RowCount > 0)
            {
                try
                {
                    if (this.SeekRow(CmcCursorBookmark.Beginning, lRow) == -1) { return retval; }
                    // SeekRow succeeded, read data
                    string[][] buffer = this.GetRawData(1); // get just 1 row
                    retval = buffer[0].ToList<string>();
                    if (!this.Flags.HasFlag(CmcOptionFlags.UseThids))
                    {
                        retval.RemoveAt(0);
                    }
                }
                catch { }
            }
            return retval;
        }

        /// <inheritdoc />
        public List<List<string>> ReadAllRows(int batchRows = 1000)
        {
            List<List<string>> retval = null;
            // capture current rowpointer
            int rowpointer = this.SeekRow(CmcCursorBookmark.Current, 0);
            // move rowpointer to start
            this.SeekRow(CmcCursorBookmark.Beginning, 0);

            if (this.RowCount > 0) // do we need this?
            {
                retval = new List<List<string>>();
                for (int i = 0; i < this.RowCount; i += batchRows)
                {
                    string[][] buffer = this.GetRawData(batchRows);
                    // from this jagged array, we want to create lists of rowvalues and add them to the output list
                    for (int j = 0; j < buffer.GetLength(0); j++)
                    {
                        List<string> newlist = buffer[j].ToList<string>();
                        // if cursor has no thids flag, the first element of newlist will always be null
                        // remove it to avoid confusion
                        if (!this.Flags.HasFlag(CmcOptionFlags.UseThids))
                        {
                            newlist.RemoveAt(0);
                        }
                        retval.Add(newlist);
                    }
                }
            }
            // put back the pointer back to where it was
            this.SeekRow(CmcCursorBookmark.Beginning, rowpointer);
            return retval;
        }

        private void RCWReleaseHandler(object sender, EventArgs e)
        {
            if (_cur != null)
            {
                Marshal.ReleaseComObject(_cur);
            }
        }

        /// <summary>
        /// Public implementation of Dispose pattern callable by consumers.
        /// </summary>
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        /// <summary>
        /// Protected implementation of Dispose pattern.
        /// </summary>
        /// <param name="disposing">disposing.</param>
        protected virtual void Dispose(bool disposing)
        {
            if (disposed)
                return;

            if (disposing)
            {
                // Free any other managed objects here.
                //
            }

            // Free any unmanaged objects here.
            //
            if (_cur != null)
            {
                Marshal.ReleaseComObject(_cur);
            }
            disposed = true;
        }
        /// <inheritdoc />
        public void Close() // just an alias for Dispose intended for COM users.
        {
            this.Dispose();
        }

        /// <summary>
        /// Reads the Commence data (using API, not DDE). This is the fastest way to retrieve Commence data.
        /// </summary>
        /// <param name="nRows">Number of rows to read at a time.</param>
        /// <returns>string[][] (a 'jagged array').</returns>
        /// <remarks>This function has no knowledge of the rowpointer.</remarks>
        /// <exception cref="CommenceCOMException"></exception>
        internal string[][] GetRawData(int nRows)
        {
            /* Note that for connected items, Commence returns a linefeed-delimited string, OR a comma delimited string(!)
             * If the connected field has no data, an empty string is returned, again linefeed-delimited.
             * It is up to the consumer to deal with this.
             * The Headers property can be used to determine what fieldtype is being returned.
             * 
             * Also note, that by getting a RowSet, Commence automatically advances the cursor's rowpointer for us
             * This can lead to some confusion for those who are used to advance it manually, like in ADO.
             */

            // CommenceQueryRowSet implements IDisposable, so we can use the using directive
            // This is vitally important, because it's Dispose method releases the COM reference (CCW) to Commence's QueryRowSet
            // Without releasing the CCW, all data would be held in memory until I don't know when exactly and memory usage would skyrocket.
            // By implementing it like this we make sure that the Garbage Collector keeps memory in check,
            // and we don't have to worry about releasing RCWs.
            // However we still run into issues if the garbage collector kicks in *after* Commence runs out of memory
            // Commence is pretty finicky about that. This is why an explicit close is included.
            using (ICommenceQueryRowSet qrs = this.GetQueryRowSet(nRows))
            {
                // number of rows requested may be larger than number of available rows in rowset,
                // so make sure the return value is sized properly
                string[][] rowvalues = new string[qrs.RowCount][];
                object[] buffer = null;
                int numColumns = qrs.ColumnCount; // store number of columns so we only need 1 COM call; makes method much faster
                int rowpointer = this.SeekRow(CmcCursorBookmark.Current, 0); // determine the rowpointer we are currently at
                int numRows = qrs.RowCount; // store number of rows to be read so we need only 1 COM call
                for (int i = 0; i < numRows; i++)
                {
                    rowvalues[i] = new string[numColumns + 1]; // number of columns plus extra element for thid
                    if (this.Flags.HasFlag(CmcOptionFlags.UseThids)) // do not make the extra API call unless requested
                    {
                        string thid = qrs.GetRowID(i, CmcOptionFlags.Default); // GetRowID does not advance the rowpointer. Note that the flag must be 0.
                        rowvalues[i][0] = thid; // put thid in first column of row
                    }

                    buffer = qrs.GetRow(i, CmcOptionFlags.Default); // don't bother with canonical flag, it doesn't work properly anyway.
                    if (buffer == null)
                    {
						qrs.Close();
                        throw new CommenceCOMException("An error occurred while reading row" + (rowpointer + i).ToString());
                    }
                    for (int j = 0; j < numColumns; j++)
                    {
                        rowvalues[i][j + 1] = buffer[j].ToString(); // put rowvalue in 2nd and up column of row
                    } // j
                } // i
                qrs.Close(); // close COM reference explicitly. the 'using' directive will do this for us, but may not kick in in time.
                return rowvalues;
            } // using; qrs will be disposed now
        }
        #endregion
    }
}

using System;
using System.Linq;
using System.Runtime.InteropServices;

namespace Vovin.CmcLibNet.Database
{
    /// <summary>
    /// Represent a rowset of items to delete. Items are not deleted until Commit is called.
    /// </summary>
    [ComVisible(true)]
    [Guid("613CBDBF-C2EB-46F1-8743-BB4E512B00A3")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(ICommenceDeleteRowSet))]
    public sealed class CommenceDeleteRowSet : BaseRowSet, ICommenceDeleteRowSet
    {
        /// <summary>
        /// the 'raw' Commence DeleteRowSet object that this class wraps.
        /// </summary>
        private FormOA.ICommenceDeleteRowSet _drs = null;
        private IRcwReleasePublisher _rcwReleasePublisher = null;
        bool disposed = false;

        #region Constructors
        /// <summary>
        /// Constructor that creates DeleteRowSet with all items.
        /// </summary>
        /// <param name="cur">FormOA.ICommenceCursor reference.</param>
        /// <param name="rcwpub">RCWReleasePublisher object used for COM Interop object cleanup.</param>
        /// <param name="flags">Option flags, must be 0.</param>
        internal CommenceDeleteRowSet(FormOA.ICommenceCursor cur, IRcwReleasePublisher rcwpub, CmcOptionFlags flags)
        {
            // deleterowset with all rows
            _drs = cur.GetDeleteRowSet(cur.RowCount, (int)flags);
            if (_drs == null)
            {
                throw new CommenceCOMException("Unable to obtain a DeleteRowSet from Commence.");
            }
            _rcwReleasePublisher = rcwpub;
            _rcwReleasePublisher.RCWRelease += this.RCWReleaseHandler;
        }

        /// <summary>
        /// Constructor that creates DeleteRowSet with set number of items.
        /// </summary>
        /// <param name="cur">FormOA.ICommenceCursor reference.</param>
        /// <param name="nCount">Number of items to delete.</param>
        /// <param name="rcwpub">RCWReleasePublisher object used for COM Interop object cleanup.</param>
        /// <param name="flags">option flags, must be 0.</param>
        internal CommenceDeleteRowSet(FormOA.ICommenceCursor cur, int nCount, IRcwReleasePublisher rcwpub ,CmcOptionFlags flags)
        {
            // deleterowset with set number of rows
            _drs = cur.GetDeleteRowSet(nCount, (int)flags);
            if (_drs == null)
            {
                throw new CommenceCOMException("Unable to obtain a DeleteRowSet from Commence.");
            }
            _rcwReleasePublisher = rcwpub;
            _rcwReleasePublisher.RCWRelease += this.RCWReleaseHandler;
        }

        /// <summary>
        /// Constructor that creates DeleteRowSet for particular row identified by RowID.
        /// </summary>
        /// <param name="cur">FormOA.ICommenceCursor reference.</param>
        /// <param name="pRowID">row id.</param>
        /// <param name="rcwpub">RCWReleasePublisher object used for COM Interop object cleanup.</param>
        /// <param name="flags">option flags, must be 0.</param>
        internal CommenceDeleteRowSet(FormOA.ICommenceCursor cur, string pRowID, IRcwReleasePublisher rcwpub, CmcOptionFlags flags)
        {
            // deleterowset by ID
            _drs = cur.GetDeleteRowSetByID(pRowID, (int)flags);
            if (_drs == null)
            {
                throw new CommenceCOMException("Unable to obtain a DeleteRowSet from Commence.");
            }
            _rcwReleasePublisher = rcwpub;
            _rcwReleasePublisher.RCWRelease += this.RCWReleaseHandler;
        }
        /// <summary>
        /// Destructor.
        /// </summary>
        ~CommenceDeleteRowSet()
        {
            Dispose(false);
        }
        #endregion

        /// <inheritdoc />
        public override int RowCount
        {
            get { return _drs.RowCount; }
        }
        /// <inheritdoc />
        public override int ColumnCount
        {
            get { return _drs.ColumnCount; }
        }
        /// <inheritdoc />
        public override string GetRowValue(int nRow, int nCol, CmcOptionFlags flags = CmcOptionFlags.Default)
        {
            try
            {
                return _drs.GetRowValue(nRow, nCol, (int)flags);
            }
            catch (COMException)
            {
                return null;
            }
        }
        /// <inheritdoc />
        public override string GetColumnLabel(int nCol, CmcOptionFlags flags = CmcOptionFlags.Default)
        {
            try
            {
                return _drs.GetColumnLabel(nCol, (int)flags);
            }
            catch (COMException)
            {
                return null;
            }
        }
        /// <inheritdoc />
        public override int GetColumnIndex(string pLabel, CmcOptionFlags flags = CmcOptionFlags.Default)
        {
            try
            {
                return _drs.GetColumnIndex(pLabel, (int)flags);
            }
            catch (COMException)
            {
                return -1;
            }
        }
        /// <inheritdoc />
        public override object[] GetRow(int nRow, CmcOptionFlags flags = CmcOptionFlags.Default)
        {
            object[] retval = null;

            try
            {
                retval = _drs.GetRow(nRow, base.Delim, (int)flags)
                    .Split(base.Splitter, StringSplitOptions.None)
                    .ToArray<object>();
            }
            catch (COMException)
            {
                // TO-DO: return something meaningful...
            }

            return retval;
        }
        /// <inheritdoc />
        public override object[] GetRow(int nRow, string delim, CmcOptionFlags flags = CmcOptionFlags.Default)
        {
            object[] retval = null;

            try
            {
                retval = _drs.GetRow(nRow, delim, (int)flags)
                    .Split(new string[] { delim }, StringSplitOptions.None)
                    .ToArray<object>();
            }
            catch (COMException)
            {
                // TO-DO: return something meaningful...
            }

            return retval;
        }
        /// <inheritdoc />
        public override bool GetShared(int nRow)
        {
            try
            {
                return _drs.GetShared(nRow);
            }
            catch (COMException)
            {
                return false;
            }
        }
        /// <inheritdoc />
        public override void Close()
        {
            this.Dispose();
        }
        /// <inheritdoc />
        public string GetRowID(int nRow, CmcOptionFlags flags = CmcOptionFlags.Default)
        {
            try
            {
                return _drs.GetRowID(nRow, (int)flags);
            }
            catch (COMException)
            {
                return null;
            }
        }

        /// <inheritdoc />
        public int DeleteRow(int nRow, CmcOptionFlags flags = CmcOptionFlags.Default)
        {
            try
            {
                return _drs.DeleteRow(nRow, (int)flags);
            }
            catch (COMException)
            {
                return -1;
            }
        }

        /// <inheritdoc />
        public int Commit(CmcOptionFlags flags = CmcOptionFlags.Default)
        {
            try
            {
                return _drs.Commit((int)flags);
            }
            catch (COMException)
            {
                return -1;
            }
        }
        internal override void RCWReleaseHandler(object sender, EventArgs e)
        {
            if (_drs != null)
            {
                Marshal.FinalReleaseComObject(_drs);
            }
        }
        // Protected implementation of Dispose pattern.
        /// <summary>
        /// Dispose method.
        /// </summary>
        /// <param name="disposing">disposing.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposed)
                return;

            if (disposing)
            {
                // Free any other managed objects here.
                //
                _rcwReleasePublisher.RCWRelease -= this.RCWReleaseHandler;
            }

            // Free any unmanaged objects here.
            //
            if (_drs != null)
            {
                Marshal.ReleaseComObject(_drs);
            }
            disposed = true;

            // Call the base class implementation.
            base.Dispose(disposing);
        }
    }
}

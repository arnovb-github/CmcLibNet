using System;
using System.Runtime.InteropServices;
using System.Linq;
using Vovin.CmcLibNet.Extensions;

namespace Vovin.CmcLibNet.Database
{
    /// <summary>
    /// Represents a rowset of items to query.
    /// </summary>
    [ComVisible(true)]
    [Guid("A1EF7CED-7305-4371-AAD0-73EC91A22AD4")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(ICommenceQueryRowSet))]
    public sealed class CommenceQueryRowSet : BaseRowSet, ICommenceQueryRowSet
    {
        /// <summary>
        /// the 'raw' Commence QueryRowSet object that this class wraps.
        /// </summary>
        private readonly FormOA.ICommenceQueryRowSet _qrs; // COM object!
        private readonly IRcwReleasePublisher _rcwReleasePublisher;
        private bool disposed;

        #region Constructors
        /// <summary>
        /// Constructor that creates QueryRowSet with all items.
        /// </summary>
        /// <param name="cur">FormOA.ICommenceCursor reference.</param>
        /// <param name="rcwpub">RCWReleasePublisher object used for COM Interop object cleanup.</param>
        /// <param name="flags">option flags, must be 0.</param>
        internal CommenceQueryRowSet(FormOA.ICommenceCursor cur, IRcwReleasePublisher rcwpub, CmcOptionFlags flags)
        {
            // queryrowset with all rows
            try
            {
                _qrs = cur.GetQueryRowSet(cur.RowCount, (int)flags);
            }
            catch (Exception)
            {
                // swallow all exceptions and throw our own.
            }
            if (_qrs == null)
            {
                throw new CommenceCOMException("Unable to obtain a QueryRowSet from Commence.");
            }
            _rcwReleasePublisher = rcwpub;
            _rcwReleasePublisher.RCWRelease += this.RCWReleaseHandler;
        }

        /// <summary>
        /// Constructor that creates QueryRowSet with set number of items.
        /// </summary>
        /// <param name="cur">FormOA.ICommenceCursor reference.</param>
        /// <param name="nCount">Number of items to query.</param>
        /// <param name="rcwpub">RCWReleasePublisher object used for COM Interop object cleanup.</param>
        /// <param name="flags">option flags, must be 0.</param>
        internal CommenceQueryRowSet(FormOA.ICommenceCursor cur, int nCount, IRcwReleasePublisher rcwpub, CmcOptionFlags flags)
        {
            // queryrowset with set number of rows
            _qrs = cur.GetQueryRowSet(nCount, (int)flags);
            if (_qrs == null)
            {
                throw new CommenceCOMException("Unable to obtain a QueryRowSet from Commence.");
            }
            _rcwReleasePublisher = rcwpub;
            _rcwReleasePublisher.RCWRelease += this.RCWReleaseHandler;
        }

        /// <summary>
        /// Constructor that creates QueryRowSet for particular row identified by RowID.
        /// </summary>
        /// <param name="cur">FormOA.ICommenceCursor reference.</param>
        /// <param name="pRowID">row or thid  id.</param>
        /// <param name="rcwpub">RCWReleasePublisher object used for COM Interop object cleanup.</param>
        /// <param name="flags">option flags, must be 0.</param>
        /// <param name="identifier"><see cref="RowSetIdentifier"/></param>
        internal CommenceQueryRowSet(FormOA.ICommenceCursor cur, string pRowID, IRcwReleasePublisher rcwpub, CmcOptionFlags flags = CmcOptionFlags.Default, RowSetIdentifier identifier = RowSetIdentifier.RowId)
        {
            switch (identifier)
            {
                case RowSetIdentifier.RowId:
                    _qrs = cur.GetQueryRowSetByID(pRowID, (int)flags);
                    break;
                case RowSetIdentifier.Thid:
                    _qrs = cur.GetQueryRowSetByThid(pRowID, (int)flags);
                    break;
            }
            if (_qrs == null)
            {
                throw new CommenceCOMException("Unable to obtain a QueryRowSet from Commence.");
            }
            _rcwReleasePublisher = rcwpub;
            _rcwReleasePublisher.RCWRelease += this.RCWReleaseHandler;
        }

        /// <summary>
        /// Destructor.
        /// </summary>
        ~CommenceQueryRowSet()
        {
            Dispose(false);
        }
        #endregion

        /// <inheritdoc />
        public override int RowCount
        {
            get { return _qrs.RowCount; }
        }
        /// <inheritdoc />
        public override int ColumnCount
        {
            get { return _qrs.ColumnCount; }
        }

        /// <inheritdoc />            
        public override string GetRowValue(int nRow, int nCol, CmcOptionFlags flags = CmcOptionFlags.Default)
        {

            try
            {
                return _qrs.GetRowValue(nRow, nCol, (int)flags);
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
                return _qrs.GetColumnLabel(nCol, (int)flags);
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
                return _qrs.GetColumnIndex(pLabel, (int)flags);
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
                retval = _qrs.GetRow(nRow, base.Delim, (int)flags)
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
        public string[] GetRow(int nRow)
        {
            string[] retval = null;

            try
            {
                retval = _qrs.GetRow(nRow, base.Delim, (int)CmcOptionFlags.Default)
                    .Split(base.Splitter, StringSplitOptions.None);
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
                retval = _qrs.GetRow(nRow, delim, (int)flags)
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
                return _qrs.GetShared(nRow);
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
        public int GetFieldToFile(int nRow, int nCol, string filename, CmcOptionFlags flags = CmcOptionFlags.Default)
        {
            try
            {
                return _qrs.GetFieldToFile(nRow, nCol, filename, (int)flags);
            }
            catch (COMException)
            {
                return -1;
            }
        }

        /// <inheritdoc />
        public string GetRowID(int nRow, CmcOptionFlags flags = CmcOptionFlags.Default)
        {

            try
            {
                return _qrs.GetRowID(nRow, (int)flags);
            }
            catch (COMException)
            {
                return null;
            }
        }

        /// <inheritdoc />
        public string GetRowTimeStamp(int nRow, CmcOptionFlags flags = CmcOptionFlags.Default)
        {
            try
            {
                return _qrs.GetRowTimeStamp(nRow, (int)flags);
            }
            catch (COMException)
            {
                return null;
            }
        }
        /// <inheritdoc />
        public int GetRowSequenceNumber(int nRow)
        {
            string thid = GetRowID(nRow);
            string hexPart;
            if (thid.CountChar(rowIdDelim) == 3)
            {
                hexPart = thid.Split(rowIdDelim)[1];
            }
            else
            {
                hexPart = thid.Split(rowIdDelim).Last();
            }
            return int.Parse(hexPart, System.Globalization.NumberStyles.HexNumber); ;
        }

        internal override void RCWReleaseHandler(object sender, EventArgs e)
        {
            if (_qrs != null)
            {
                Marshal.ReleaseComObject(_qrs);
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
            if (_qrs != null)
            {
                Marshal.ReleaseComObject(_qrs);
            }
            disposed = true;

            // Call the base class implementation.
            base.Dispose(disposing);
        }
    }
}

using System;
using Vovin.CmcLibNet;
using System.Runtime.InteropServices;

namespace Vovin.CmcLibNet.Database
{
    /// <summary>
    /// Represents a rowset of items to query.
    /// </summary>
    [ComVisible(true)]
    [Guid("A1EF7CED-7305-4371-AAD0-73EC91A22AD4")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(Vovin.CmcLibNet.Database.ICommenceQueryRowSet))]
    public sealed class CommenceQueryRowSet : BaseRowSet, Vovin.CmcLibNet.Database.ICommenceQueryRowSet
    {
        /// <summary>
        /// the 'raw' Commence QueryRowSet object that this class wraps.
        /// </summary>
        private FormOA.ICommenceQueryRowSet _qrs = null; // COM object!
        private IRCWReleasePublisher _rcwReleasePublisher = null;
        bool disposed = false;

        #region Constructors
        /// <summary>
        /// Constructor that creates QueryRowSet with all items.
        /// </summary>
        /// <param name="cur">FormOA.ICommenceCursor reference.</param>
        /// <param name="rcwpub">RCWReleasePublisher object used for COM Interop object cleanup.</param>
        /// <param name="flags">option flags, must be 0.</param>
        internal CommenceQueryRowSet(FormOA.ICommenceCursor cur, IRCWReleasePublisher rcwpub,CmcOptionFlags flags)
        {
            // queryrowset with all rows
            try
            {
                _qrs = cur.GetQueryRowSet(cur.RowCount, (int)flags);
                //Console.WriteLine(flags.ToString()); // DEBUG
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
        internal CommenceQueryRowSet(FormOA.ICommenceCursor cur, int nCount, IRCWReleasePublisher rcwpub,CmcOptionFlags flags)
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
        /// <param name="pRowID">row id.</param>
        /// <param name="rcwpub">RCWReleasePublisher object used for COM Interop object cleanup.</param>
        /// <param name="flags">option flags, must be 0.</param>
        internal CommenceQueryRowSet(FormOA.ICommenceCursor cur, string pRowID, IRCWReleasePublisher rcwpub ,CmcOptionFlags flags)
        {
            // queryrowset by id
            _qrs = cur.GetQueryRowSetByID(pRowID, (int)flags);
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
            string[] result = null;

            try
            {
                result = _qrs.GetRow(nRow, base.Delim, (int)flags).Split(base._splitter, StringSplitOptions.None);
                retval = toObjectArray(result);
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
                return _qrs.GetFieldToFile(nRow,nCol,filename,(int)flags);
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

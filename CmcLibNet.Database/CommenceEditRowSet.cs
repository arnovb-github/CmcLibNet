using System;
using System.Runtime.InteropServices;
using Vovin.CmcLibNet;

namespace Vovin.CmcLibNet.Database
{
    /// <summary>
    /// Represents a rowset of items to edit.
    /// </summary>
    [ComVisible(true)]
    [GuidAttribute("B2B03D51-74A6-4D98-A1C0-A4B6FB0997D3")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(Vovin.CmcLibNet.Database.ICommenceEditRowSet))]
    public sealed class CommenceEditRowSet : BaseRowSet, Vovin.CmcLibNet.Database.ICommenceEditRowSet
    {
        /// <summary>
        /// the 'raw' Commence EditRowSet object that this class wraps.
        /// </summary>
        private FormOA.ICommenceEditRowSet _ers = null;
        private IRCWReleasePublisher _rcwReleasePublisher = null;
        bool disposed = false;

        #region Constructors
        /// <summary>
        /// Constructor that creates EditRowSet with all items.
        /// </summary>
        /// <param name="cur">FormOA.ICommenceCursor reference.</param>
        /// <param name="rcwpub">RCWReleasePublisher object used for COM Interop object cleanup.</param>
        /// <param name="flags">option flags, must be 0.</param>
        internal CommenceEditRowSet(FormOA.ICommenceCursor cur, IRCWReleasePublisher rcwpub, CmcOptionFlags flags)
        {
            // editrowset with all rows
            _ers = cur.GetEditRowSet(cur.RowCount, (int)flags);
            if (_ers == null)
            {
                throw new CommenceCOMException("Unable to obtain a EditRowSet from Commence.");
            }
            _rcwReleasePublisher = rcwpub;
            _rcwReleasePublisher.RCWRelease += this.RCWReleaseHandler;
        }

        /// <summary>
        /// Constructor that creates EditRowSet with set number of items.
        /// </summary>
        /// <param name="cur">FormOA.ICommenceCursor reference.</param>
        /// <param name="nCount">Number of items to edit.</param>
        /// <param name="rcwpub">RCWReleasePublisher object used for COM Interop object cleanup.</param>
        /// <param name="flags">option flags, must be 0.</param>
        internal CommenceEditRowSet(FormOA.ICommenceCursor cur, int nCount, IRCWReleasePublisher rcwpub ,CmcOptionFlags flags)
        {
            // editrowset with set number of rows
            _ers = cur.GetEditRowSet(nCount, (int)flags);
            if (_ers == null)
            {
                throw new CommenceCOMException("Unable to obtain a EditRowSet from Commence.");
            }
            _rcwReleasePublisher = rcwpub;
            _rcwReleasePublisher.RCWRelease += this.RCWReleaseHandler;
        }

        /// <summary>
        /// Constructor that creates EditRowSet for particular row identified by RowID.
        /// </summary>
        /// <param name="cur">FormOA.ICommenceCursor reference.</param>
        /// <param name="pRowID">row id.</param>
        /// <param name="rcwpub">RCWReleasePublisher object used for COM Interop object cleanup.</param>
        /// <param name="flags">option flags, must be 0.</param>
        internal CommenceEditRowSet(FormOA.ICommenceCursor cur, string pRowID, IRCWReleasePublisher rcwpub, CmcOptionFlags flags)
        {
            // editrowset by ID
            _ers = cur.GetEditRowSetByID(pRowID, (int)flags);
            if (_ers == null)
            {
                throw new CommenceCOMException("Unable to obtain a EditRowSet from Commence.");
            }
            _rcwReleasePublisher = rcwpub;
            _rcwReleasePublisher.RCWRelease += this.RCWReleaseHandler;
        }

        /// <summary>
        /// Destructor.
        /// </summary>
        ~CommenceEditRowSet()
        {
            Dispose(false);
        }
        #endregion
        /// <inheritdoc />
        public override int RowCount
        {
            get { return _ers.RowCount; }
        }
        /// <inheritdoc />
        public override int ColumnCount
        {
            get { return _ers.ColumnCount; }
        }
        /// <inheritdoc />
        public override string GetRowValue(int nRow, int nCol, CmcOptionFlags flags = CmcOptionFlags.Default)
        {
            try
            {
                return _ers.GetRowValue(nRow, nCol, (int)flags);
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
                return _ers.GetColumnLabel(nCol, (int)flags);
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
                return _ers.GetColumnIndex(pLabel, (int)flags);
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
                result = _ers.GetRow(nRow, base.Delim, (int)flags).Split(base._splitter, StringSplitOptions.None);
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
                return _ers.GetShared(nRow);
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
        public bool SetShared(int nRow)
        {
            try
            {
                return _ers.SetShared(nRow);
            }
            catch (COMException)
            {
                return false;
            }
        }
        /// <inheritdoc />
        public string GetRowID(int nRow, CmcOptionFlags flags = CmcOptionFlags.Default)
        {
            try
            {
                return _ers.GetRowID(nRow, (int)flags);
            }
            catch (COMException)
            {
                return null;
            }
        }
        /// <inheritdoc />
        public int Commit(CmcOptionFlags flags = CmcOptionFlags.Default)
        {
            try
            {
                return _ers.Commit((int)flags);
            }
            catch (COMException)
            {
                return -1;
            }
        }
        /// <inheritdoc />
        public CommenceCursor CommitGetCursor(CmcOptionFlags flags = CmcOptionFlags.Default)
        {
            CommenceCursor cc = null;
            FormOA.ICommenceCursor newcur = null;
            try
            {
                newcur = _ers.CommitGetCursor((int)flags);
                if (newcur != null)
                {
                    cc = new CommenceCursor(newcur, _rcwReleasePublisher);
                }
            }
            catch (COMException)
            {
            }
            return cc;
        }
        /// <inheritdoc />
        public int ModifyRow(int nRow, int nCol, string fieldValue, CmcOptionFlags flags = CmcOptionFlags.Default)
        {
            try
            {
                return _ers.ModifyRow(nRow, nCol, fieldValue, (int)flags);
            }
            catch (COMException)
            {
                return -1;
            }
        }

        internal override void RCWReleaseHandler(object sender, EventArgs e)
        {
            if (_ers != null)
            {
                Marshal.ReleaseComObject(_ers);
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
            if (_ers != null)
            {
                Marshal.ReleaseComObject(_ers);
            }
            disposed = true;

            // Call the base class implementation.
            base.Dispose(disposing);
        }
    }
}

using System;
using System.Linq;
using System.Runtime.InteropServices;

namespace Vovin.CmcLibNet.Database
{
    /// <summary>
    /// Represents a AddRowSet to add items to Commence.
    /// </summary>
    [ComVisible(true)]
    [Guid("B7E1F32E-D346-48F0-B0D6-8BDCC3D96565")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(ICommenceAddRowSet))]
    public sealed class CommenceAddRowSet : BaseRowSet, ICommenceAddRowSet
    {
        /// <summary>
        /// the 'raw' Commence AddRowSet object that this class wraps.
        /// </summary>
        private FormOA.ICommenceAddRowSet _ars = null;
        private IRcwReleasePublisher _rcwReleasePublisher = null;
        bool disposed = true;

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="cur">FormOA.ICommenceCursor reference.</param>
        /// <param name="nCount">Number of items to add.</param>
        /// <param name="rcwpub">RCWReleasePublisher object used for COM Interop object cleanup.</param>
        /// <param name="flags">CMC_OPTION flags.</param>
        internal CommenceAddRowSet(FormOA.ICommenceCursor cur, int nCount, IRcwReleasePublisher rcwpub, CmcOptionFlags flags)
        {
            // addrowset with set number of rows
            _ars = cur.GetAddRowSet(nCount, (int)flags);
            if (_ars == null)
            {
                throw new CommenceCOMException("Unable to obtain a AddRowSet from Commence.");
            }
            _rcwReleasePublisher = rcwpub;
            _rcwReleasePublisher.RCWRelease += this.RCWReleaseHandler;
        }
        /// <summary>
        /// Destructor.
        /// </summary>
        ~CommenceAddRowSet()
        {
            Dispose(false);
        }
        /// <inheritdoc />
        public override int RowCount
        {
            get { return _ars.RowCount; }
        }
        /// <inheritdoc />
        public override int ColumnCount
        {
            get { return _ars.ColumnCount; }
        }
        /// <inheritdoc />
        public override string GetRowValue(int nRow, int nCol, CmcOptionFlags flags = CmcOptionFlags.Default)
        {
            try
            {
                return _ars.GetRowValue(nRow, nCol, (int)flags);
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
                return _ars.GetColumnLabel(nCol, (int)flags);
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
                return _ars.GetColumnIndex(pLabel, (int)flags);
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
                retval = _ars.GetRow(nRow, base.Delim, (int)flags)
                    .Split(base._splitter, StringSplitOptions.None)
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
                return _ars.GetShared(nRow);
            }
            catch (COMException)
            {
                return false;
            }
        }

        /// <inheritdoc />
        public bool SetShared(int nRow)
        {
            try
            {
                return _ars.SetShared(nRow);
            }
            catch (COMException)
            {
                return false;
            }
        }
        /// <inheritdoc />
        public int ModifyRow(int nRow, int nCol, string fieldValue, CmcOptionFlags flags = CmcOptionFlags.Default)
        {
            try
            {
                return _ars.ModifyRow(nRow, nCol, fieldValue, (int)flags);
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
                return _ars.Commit((int)flags);
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
                newcur = _ars.CommitGetCursor((int)flags);
                if (newcur != null)
                {
                    cc = new CommenceCursor(newcur, _rcwReleasePublisher);
                }
            }
            catch (COMException)
            {
            }
            // we should Dispose ourselves now, the AddRowset is no longer valid.
            // ...but I don't know how.
            return cc;
        }

        /// <inheritdoc />
        public override void Close()
        {
            this.Dispose();
        }

        internal override void RCWReleaseHandler(object sender, EventArgs e)
        {
            if (_ars != null)
            {
                Marshal.ReleaseComObject(_ars);
            }
        }

        // Protected implementation of Dispose pattern.
        /// <summary>
        /// Dispose method
        /// </summary>
        /// <param name="disposing">disposing</param>
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
            if (_ars != null)
            {
                Marshal.ReleaseComObject(_ars);
            }
            disposed = true;

            // Call the base class implementation.
            base.Dispose(disposing);
        }
    }
}
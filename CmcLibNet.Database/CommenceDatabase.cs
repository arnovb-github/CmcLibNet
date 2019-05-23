using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Vovin.CmcLibNet.Database.Metadata;

namespace Vovin.CmcLibNet.Database
{
    /// <summary>
    /// .Net Wrapper around the Commence database API.
    /// <para>COM clients can instantiate this by using the ProgId <c>CmcLibNet.Database</c>.</para>
    /// So from VBScript you would do:
    /// <code language="vbscript">
    /// Dim db : Set db = CreateObject("CmcLibNet.Database")
    /// '.. do stuff ..
    /// db.Close
    /// Set db = Nothing</code>
    /// <para>When used fom a Commence Item Detail Form or Office VBA, be sure to read up on the <see cref="Close()"/> method.</para>
    /// </summary>
    /// <remarks>Most of the documentation on this interface was copied 'verbatim' from the Commence help files.</remarks>
    [ComVisible(true)]
    [Guid("028E932E-4B00-49eb-BBD6-C1C7F2C0B834")]
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId("CmcLibNet.Database")]
    [ComDefaultInterface(typeof(ICommenceDatabase))]
    public partial class CommenceDatabase : ICommenceDatabase
    {
        // this portion of the class contains the implementation of the ICommenceCursor interface of Commence.
        private FormOA.ICommenceDB _db;
        private CommenceApp _app; // notice we do not use the interface, but the class directly, because we want to access the RCWRelease stuff which is not part of the interface.
        private IRcwReleasePublisher _rcwReleasePublisher;

        #region Constructors

        /// <summary>
        /// Public constructor.
        /// </summary>
        public CommenceDatabase()
        {
            _app = new CommenceApp(); // when we call this first, the next line does not fail when called from COM.
            // we also want this call to CommenceApp because it contains the check to see if commence.exe is running.
            _db = _app.Db; // THIS FAILS WHEN CALLED FROM COM WHEN NO INSTANCE OF CommenceApp EXISTS. CommenceApp isn't instantiated. Works from .NET without a CommenceApp instance.
            /* We create a publisher object so any classes created from this instance can use it.
             * The idea is that this way, COM clients can create and close the several COM-visible components (CommenceApp, Export) safely.
             * It would be easier to use a static event notifying all classes that consume a COM resource to release it.
             * This leads to unwanted behaviour because all instances will be notified.
             * The down-side of the current approach is that we have to pass the event publisher to all classes that will subscribe to it.
             * There must be a better way?
             */
            _rcwReleasePublisher = new RcwReleasePublisher();
            _rcwReleasePublisher.RCWRelease += _app.RCWReleaseHandler;
            _rcwReleasePublisher.RCWRelease += this.RCWReleaseHandler;            
        }
        #endregion

        #region Methods

        /// <summary>
        /// Get a CommenceCursor object that wraps the native FormOA.ICommenceCursor interface.
        /// </summary>
        /// <exception cref="CommenceCOMException">Unable to create a cursor.</exception>
        /// <param name="categoryName">Commence category name.</param>
        /// <returns>CommenceCursor wrapping Commence ICommenceCursor with default settings</returns>
        public Database.ICommenceCursor GetCursor(string categoryName)
        {
            CommenceCursor _cur; // our own cursor object
            /* 
            * Contrary to what Commence documentation says,
            * GetCursor() does not return null on error,
            * but instead a COM error is raised
            */
            try
            {
                _cur = new CommenceCursor(_db, categoryName, _rcwReleasePublisher); // should not need dependency injection, or should it?
                return _cur;
            }
            catch (COMException e)
            {
                throw new CommenceCOMException("GetCursor failed to create a cursor on category or view '" + categoryName + "'", e);
            }
        }
        /// <summary>
        /// Get a CommenceCursor object that wraps the native FormOA.ICommenceCursor interface.
        /// </summary>
        /// <param name="pName">Commence category name.</param>
        /// <param name="pCursorType">Type of Commence data to access with this cursor.</param>
        /// <param name="pCursorFlags">Logical OR of Option flags.</param>
        /// <returns>CommenceCursor wrapping ICommenceCursor according to flags passed.</returns>
        /// <exception cref="NotSupportedException">Viewtype cannot be used in a CommenceCursor.</exception>
        /// <exception cref="CommenceCOMException">Commence could not create a cursor.</exception>
        public Database.ICommenceCursor GetCursor(string pName, CmcCursorType pCursorType, CmcOptionFlags pCursorFlags)
        {
            CommenceCursor cur; // our own cursor object
            /* 
            * Contrary to what Commence documentation says,
            * GetCursor() does not return null on error,
            * but instead a COM error is raised.
            */
            IViewDef vd = null;
            if (pCursorType == CmcCursorType.View)
            {
                List<string> unsupported = new List<string>(new string[] { "Add Item", "Item Detail", "Multi-View", "Report Viewer", "Document", "Gantt Chart", "Calendar" });
                vd = GetViewDefinition(pName);
                if (vd != null && unsupported.Contains(vd.Type))
                {
                    throw new NotSupportedException("View is of type: " + vd.Type + 
                        "\nNo cursor can be created on views of the following types:\n\n" + 
                        string.Join(", ", unsupported) + 
                        ".\nEither the view type is unsupported for Commence cursors altogether, or it may produce inconsistent results. Use a view of type Report or Grid instead.");
                }
                try
                {
                    cur = new CommenceCursor(_db, pCursorType, pName, _rcwReleasePublisher, pCursorFlags, vd.Type);
                }
                catch (COMException e)
                {
                    throw new CommenceCOMException("GetCursor failed to create a cursor on category or view '" + pName + "'.", e);
                }
            }
            else
            {
                try
                {
                    cur = new CommenceCursor(_db, pCursorType, pName, _rcwReleasePublisher, pCursorFlags);
                }
                catch (COMException e)
                {
                    throw new CommenceCOMException("GetCursor failed to create a cursor on category or view '" + pName + "'.", e);
                }
            }
            return cur;
        }

        private void RCWReleaseHandler(object sender, EventArgs e)
        {
            if (_db != null)
            {
                while (Marshal.ReleaseComObject(_db) > 0) { };
                _db = null;
            }
        }

        /// <inheritdoc />
        public void Close()
        {
            _rcwReleasePublisher.ReleaseRCWReferences(); // notify subscribers to release COM references.

            // If we don't force garbage collection,
            // the commence.exe will remain open at least when the devenv is running
            // AVB 2019-03-18 seems we can get away with it after implementing IDisposable
#if DEBUG
            GC.Collect(); // force garbage collection
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
#endif
        }

        #endregion

        #region Properties
        /// <inheritdoc />
        public string Path
        {
            get
            {
                return _app.Path;
            }

        }
        /// <inheritdoc />
        public string Name
        {
            get
            {
                return _app.Name;
            }

        }
        #endregion

        #region IDisposable Support
        private bool disposedValue = false; // To detect redundant calls

        /// <summary>
        /// Dispose
        /// </summary>
        /// <param name="disposing"></param>
        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    Close();
                    // TODO: dispose managed state (managed objects).
                    if (DDETimer != null)
                    {
                        DDETimer.Stop();
                        _conv.CloseConversation();
                    }
                    _conv = null;
                }

                // TODO: free unmanaged resources (unmanaged objects) and override a finalizer below.
                if (_db != null)
                {
                    while (Marshal.ReleaseComObject(_db) > 0) { };
                    _db = null;
                }

                // TODO: set large fields to null.

                disposedValue = true;
            }
        }

        // TODO: override a finalizer only if Dispose(bool disposing) above has code to free unmanaged resources.
        /// <summary>
        /// Finalizer
        /// </summary>
        ~CommenceDatabase() // overriding finalizer
        {
            // Do not change this code. Put cleanup code in Dispose(bool disposing) above.
            Dispose(false);
        }

        // This code added to correctly implement the disposable pattern.
        /// <summary>
        /// Dispose
        /// </summary>
        public void Dispose()
        {
            // Do not change this code. Put cleanup code in Dispose(bool disposing) above.
            Dispose(true);
            // TODO: uncomment the following line if the finalizer is overridden above.
            GC.SuppressFinalize(this);
        }
        #endregion
    }
}

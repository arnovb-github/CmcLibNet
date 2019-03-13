using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace Vovin.CmcLibNet.Database
{
    /// <summary>
    /// Wrapper around the Commence database API functionality, including all functionality provided by its DDE interface (ICommenceConversation).
    /// There is no need to use ICommenceConversation with its DDE syntax with CmcLibNet.
    /// All functionality provided by the ICommenceConversation interface are implemented as native functions of this class. All underlying DDE plumbing is handled internally.
    /// <para>In fact, you cannot even get a reference to FormOA.ICommenceConversation from CmcLibNet.</para>
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
    [ProgIdAttribute("CmcLibNet.Database")]
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
            _db = CommenceApp.DB; // THIS FAILS WHEN CALLED FROM COM WHEN NO INSTANCE OF CommenceApp EXISTS. CommenceApp isn't instantiated. Works from .NET without a CommenceApp instance.
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

        ///// <summary>
        ///// Destructor.
        ///// </summary>
        //~CommenceDatabase()
        //{
        //    Dispose(false);
        //}

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
                _cur = new CommenceCursor(categoryName, _rcwReleasePublisher); // should not need dependency injection, or should it?
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
                if (vd != null && unsupported.Contains(vd.TypeDescription))
                {
                    throw new NotSupportedException("View is of type: " + vd.TypeDescription + "\nNo cursor can be created on views of the following types:\n\n" + string.Join(", ", unsupported) + ".\nEither the view type is unsupported for Commence cursors altogether, or it may produce inconsistent results. Use a view of type Report or Grid instead.");
                }
                try
                {
                    cur = new CommenceCursor(pCursorType, pName, _rcwReleasePublisher, pCursorFlags, vd.TypeDescription);
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
                    cur = new CommenceCursor(pCursorType, pName, _rcwReleasePublisher, pCursorFlags);
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
                Marshal.ReleaseComObject(_db);
            }
        }

        /// <inheritdoc />
        public void Close()
        {
            _rcwReleasePublisher.ReleaseRCWReferences(); // notify subscribers to release COM references.
            GC.Collect(); // force garbage collection
            GC.WaitForPendingFinalizers();
            GC.Collect();
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
    }
}

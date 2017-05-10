using System;
using System.Runtime.InteropServices;
using Vovin.CmcLibNet.Core;

/*
 * The ICommenceConversation object is a COM wrapper around the DDE calls that can be made to Commence.
 * The DDE specification for Commence specifies the following topics:
 * -System
 * -GetData (deprecated)
 * -ViewData (deprecated)
 * -<database name>
 * -<database path>

 * Native DDE actually provides more flexibility that the ICommenceConversation object
 * ICommenceConversation can only be created from the running instance of the Commence application
 * and only pertains to that copy of the database.
 * This has repercussions on what Commence will accept as valid topics: 
 * whereas native DDE will accept valid database names of paths for databases that are opened,
 * ICommenceConversation only accepts the name or path of the database it was created from. 
 * The DDE specification also states that the ViewData and GetData topics are only for backwards compatibility
 * and that name or path should be used instead.
 */

 namespace Vovin.CmcLibNet.Database
 {
    /// <summary>
    /// Wrapper around the ICommenceConversation object.
    /// </summary>
    internal class CommenceConversation : IDisposable
    {
        private FormOA.ICommenceConversation _nativeConv = null; // COM object
        private readonly string _pszApplication = "Commence";
        //private readonly string _topic = "GetData";
        bool disposed = false;

        #region Constructors
        /// <summary>
        /// Allows for specifying a DDE topic name. Use 'System' or 'MergeItem', otherwise create default conversation.
        /// This constructor is not likely to be used; note the restrictions pertaining to GetConversation from the API (as opposed to native DDE).
        /// </summary>
        /// <param name="sTopic">Topic name.</param>
        internal CommenceConversation(string sTopic)
        {
            //_topic = sTopic;
            _nativeConv = CommenceApp.DB.GetConversation(_pszApplication, sTopic);
        }

        /// <summary>
        /// For future use; allows for specifying an application name.
        /// Currently Commence only accepts 'Commence'.
        /// </summary>
        /// <param name="sApplication">Application name.</param>
        /// <param name="sTopic">DDE Topic.</param>
        internal CommenceConversation(string sApplication, string sTopic)
        {
            //_pszApplication = sApplication;
            //_topic = sTopic;
            _nativeConv = CommenceApp.DB.GetConversation(sApplication, sTopic);
        }
        #endregion

        #region Methods
        /// <summary>
        /// Perform DDE Request command.
        /// </summary>
        /// <exception cref="CommenceDDEException">DDE exception.</exception>
        /// <param name="dde">DDE request command string. See Commence documentation.</param>
        /// <returns>String with request results.</returns>
        internal string DDERequest(string dde)
        {
            string retval = null;
            try
            {
                retval = _nativeConv.Request(dde);
                //in order to be able to check for a returned null from this method
                //we will include an extra check for that
                if (String.IsNullOrEmpty(retval))
                {
                    return null; // wise?
                }
                return retval;
            }

            catch (System.Runtime.InteropServices.COMException e)
            {
                //a generic COM exception occurred,
                //so we'll throw our own custom error here to provide some more info
                string DDEError = string.Empty;
                try
                {
                    DDEError = _nativeConv.Request("[GetLastError]");
                    if (DDEError.Length > 0)
                    {
                        throw new CommenceDDEException("Commence DDE request returned error: " + DDEError, e);
                    }
                    else
                    {
                        throw new CommenceDDEException("Commence failed to process DDE request. That's all we know.", e);
                    }
                }

                catch (CommenceDDEException)
                {
                    throw;
                }
                catch (System.Runtime.InteropServices.COMException) //useful??
                {
                    throw;
                }
            }
        }
        /// <summary>
        /// Performs DDE Execute command.
        /// </summary>
        /// <exception cref="CommenceDDEException">DDE exception.</exception>
        /// <param name="dde">DDEExecute command string. See Commence documentation</param>
        /// <returns>True on success, false on error</returns>
        internal bool DDEExecute(string dde)
        {
            try
            {
                return _nativeConv.Execute(dde);
            }

            catch (System.Runtime.InteropServices.COMException e)
            {
                // a generic COM exception occurred,
                // we'll throw our own custom error here to provide some more info
                string DDEError = string.Empty;
                try
                {
                    DDEError = _nativeConv.Request("[GetLastError]");
                    if (DDEError.Length > 0)
                    {
                        throw new CommenceDDEException("Commence DDE execute returned error: " + DDEError, e);
                    }
                    else
                    {
                        throw new CommenceDDEException("Commence failed to process DDE execute.", e);
                    }
                }
                // for some reason, we have to catch and rethrow to make it catchable in another method
                // it has something to do with preserving the stack trace
                catch (CommenceDDEException)
                {
                    throw;
                }
                catch (System.Runtime.InteropServices.COMException) // useful if GetLastError failed?
                {
                    throw;
                }
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
        /// <param name="disposing">disposing</param>
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
            if (_nativeConv != null)
            {
                Marshal.ReleaseComObject(_nativeConv);
                _nativeConv = null; // explicitly make _nativeConv eligible for garbage collection
            }
            disposed = true;
        }

        /// <summary>
        /// Conversations are kept open until a Timer elapses.
        /// This method should subscribe to the Elapsed event.
        /// </summary>
        /// <param name="sender">sender.</param>
        /// <param name="e">ElapsedEventArgs.</param>
        internal void OnTimedEvent(Object sender, System.Timers.ElapsedEventArgs e)
        {
            Console.WriteLine("The Elapsed event was raised at {0}", e.SignalTime + ". Disposing CommenceConversation object.");
            Dispose();
        }
        #endregion
    }
}

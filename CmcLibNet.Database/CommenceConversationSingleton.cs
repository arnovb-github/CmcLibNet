using System;
using System.Runtime.InteropServices;
using Vovin.CmcLibNet;

/*
 * The FormOA.ICommenceConversation object is a COM wrapper around the DDE calls that can be made to Commence.
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
 * FormOA.ICommenceConversation only accepts the name or path of the database it was created from. 
 * The DDE specification also states that the ViewData and GetData topics are only for backwards compatibility
 * and that name or path should be used instead.
 */

namespace Vovin.CmcLibNet.Database
{
    /// <summary>
    /// Singleton class wrapping the Commence FormOA.ICommenceConversation interface (its 'DDE engine').
    /// </summary>
    internal sealed class CommenceConversation : IDisposable // Implementing IDisposable is actually overkill but it doesn't hurt
    {
        private static readonly Lazy<CommenceConversation> lazy = new Lazy<CommenceConversation>(() => new CommenceConversation());
        private FormOA.ICommenceConversation _nativeConv = null; // COM object
        private string _pszApplication = "Commence"; // the  only supported application parameter at this time
        private string _topic = "System"; // rarely used topic!
        bool disposed = false;

        #region Constructors
        // Do not allow other classes to create an instance
        private CommenceConversation() {}

        ~CommenceConversation()
        {
            Dispose(false);
        }
        #endregion

        /// <summary>
        /// Return class instance.
        /// </summary>
        internal static CommenceConversation Instance
        {
            get
            {
                return lazy.Value;
            }
        }

        // if no topic is set, System is used.
        internal string Topic 
        {
            get
            {
                return _topic;
            }
            set
            {
                _topic = value;
            }
        }

        internal string Application 
        {
            get
            {
                return _pszApplication;
            }
            private set // currently Commence only allows 1 value, no reason to expose this
            {
                _pszApplication = value;
            }
        }

        /// <summary>
        /// Returns the actual FormOA.ICommenceConversation object.
        /// </summary>
        private FormOA.ICommenceConversation Conversation
        {
            get
            {
                // for now we only return a single topic (path), which equates to the obsolete GetData
                // There is no way to request the System topic (yet).
                if (_nativeConv == null)
                {
                    // This should be in a try/catch block.
                    _nativeConv = CommenceApp.DB.GetConversation(this.Application, this.Topic);
                }
                return _nativeConv;
            }
        }
        #region Methods
        /// <summary>
        /// Perform DDE Request command.
        /// </summary>
        /// <exception cref="CommenceDDEException">DDE exception.</exception>
        /// <param name="dde">DDE request command string. See Commence documentation.</param>
        /// <returns>String with request results.</returns>
        /// <exception cref="COMException"></exception>
        /// <exception cref="CommenceDDEException"></exception>
        internal string DDERequest(string dde)
        {
            try
            {
                string retval = Conversation.Request(dde);
                return retval;
            }
            catch (COMException e)
            {
                // A generic COM exception occurred,
                // We'll try to throw our own custom error here to provide some more info
                string DDEError = string.Empty;
                try
                {
                    DDEError = _nativeConv.Request("[GetLastError]"); // may throw it's own exception
                    if (DDEError.Length > 0)
                    {
                        throw new CommenceDDEException("Commence DDE request returned error: " + DDEError, e);
                    }
                    else
                    {
                        throw new CommenceDDEException("Commence failed to process DDE request. That's all we know.", e);
                    }
                }
                catch (COMException) // we got another exception
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
        /// <exception cref="COMException"></exception>
        /// <exception cref="CommenceDDEException"></exception>
        internal bool DDEExecute(string dde)
        {
            try
            {
                return Conversation.Execute(dde);
            }
            catch (COMException e)
            {
                // a generic COM exception occurred,
                // we'll throw our own custom error here to provide some more info
                string DDEError = string.Empty;
                try
                {
                    DDEError = _nativeConv.Request("[GetLastError]"); // may throw it's own exception
                    if (DDEError.Length > 0)
                    {
                        throw new CommenceDDEException("Commence DDE execute returned error: " + DDEError, e);
                    }
                    else
                    {
                        throw new CommenceDDEException("Commence failed to process DDE execute.", e);
                    }
                }
                catch (Exception)
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
        private void Dispose(bool disposing)
        {
            if (disposed) // [weird comment] PowerShell and VBA upon running >1 times regard this as disposed so no cleanup is done until it is closed. Strange.
            {
                return;
            }

            if (disposing)
            {
                // Free any other managed objects here.
                //
            }

            // Free any unmanaged objects here.
            //
            if (_nativeConv != null)
            {
                Marshal.ReleaseComObject(_nativeConv); // closes conversation or so it should
            }
            disposed = true;
        }

        /// <summary>
        /// Closes the DDE conversation.
        /// Multiple requests can made in a single conversation,
        /// so closing the conversation after every request would add considerable overhead.
        /// <para>Conversations should kept open until a Timer elapses.
        /// This method should subscribe to the Elapsed event of that Timer.</para>
        /// <para>This way, the conversation stays open and multiple requests can be made.
        /// If no more requests are received, the conversation is closed after the timer elapses.</para>
        /// </summary>
        /// <param name="sender">sender.</param>
        /// <param name="e">ElapsedEventArgs.</param>
        internal void HandleTimerElapsed(Object sender, System.Timers.ElapsedEventArgs e)
        {
            if (_nativeConv != null)
            {
                Marshal.ReleaseComObject(_nativeConv); // closes the actual conversation
                _nativeConv = null; // explicitly make _nativeConv eligible for garbage collection. Not needed.
            }
            Dispose(); // there is no need for this since we released the ComObject which was the whole point.
            // we could get rid of IDisposable altogether.
        }
        #endregion
    }
}

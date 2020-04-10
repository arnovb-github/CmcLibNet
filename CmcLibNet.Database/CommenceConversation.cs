using System;
using System.Runtime.InteropServices;

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
    /// There can be only a limited number of Commence DDE conversations (10), 
    /// so by making this a singleton, there should be just one at a time.
    /// </summary>
    internal sealed class CommenceConversation
    {
        private static readonly Lazy<CommenceConversation> lazy = new Lazy<CommenceConversation>(() => new CommenceConversation());
        private FormOA.ICommenceConversation _nativeConv = null; // COM object

        #region Constructors
        // Do not allow other classes to create an instance
        private CommenceConversation() {}
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
        internal string Topic { get; set; } = "System";

        internal string Application { get; private set; } = "Commence"; // currently Commence only allows 1 value, no reason to expose this

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
                    CommenceApp _app = new CommenceApp();
                   _nativeConv = _app.Db.GetConversation(this.Application, this.Topic);
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
        internal void HandleTimerElapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            CloseConversation();
        }

        internal void CloseConversation()
        {
            if (_nativeConv != null)
            {
                while (Marshal.ReleaseComObject(_nativeConv) > 0) { }; // closes the actual conversation
                _nativeConv = null; // explicitly make _nativeConv eligible for garbage collection. Not needed.
            }
        }
        #endregion
    }
}

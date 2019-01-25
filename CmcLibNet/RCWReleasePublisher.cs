using System;

namespace Vovin.CmcLibNet
{
    /// <summary>
    /// Interface for RCWReleasePublisher.
    /// </summary>
    internal interface IRCWReleasePublisher
    {
        event EventHandler RCWRelease;
        void ReleaseRCWReferences();
    }

    /// <summary>
    /// Raises the RCWReleased event to notify subscribed classes that use FormOA objects to release them.
    /// This is necessary when this assembly is run from a Commence Item Detail Form script.
    /// </summary>
    /// <remarks>Because of how VBScript support is implemented in Item Detail Form scripting,
    /// Commence and this assembly keep each other in deadlock;
    /// after setting the reference to this assembly to Nothing from within a Form Script,
    /// the assembly won't run it's finalizers because it thinks the FormOA COM objects (RCWs) are still in use.
    /// This not just a Form Script issue, it also happens when used in Office VBA.
    /// It has something to do with how the code calling the assembly does not tell mscoree.dll (a system dll) to release the COM references.
    /// <para>In the future, perhaps careful implementation of IDisposable could make this unnecessary. AVB 20170107 No it won't.</para> 
    /// <para>The issue could maybe be brute-forced with calling GC.Collect, but that is considered bad practice.
    /// Also, that still leaves us with the problem that there is no 'exit point',
    /// no way of detecting that the form script is done with the assembly. It still requires an explicit call. Sux.</para>
    /// </remarks>
    internal class RCWReleasePublisher : IRCWReleasePublisher
    {
        public event EventHandler RCWRelease; 

        // convenience method to raise the event, without data.
        public void ReleaseRCWReferences()
        {
            OnRCWRelease(EventArgs.Empty);
        }

        // raises the actual event.
        protected virtual void OnRCWRelease(EventArgs e)
        {
            RCWRelease?.Invoke(this, e);
        }
    }
}

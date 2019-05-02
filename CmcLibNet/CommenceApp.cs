using System;
using System.Diagnostics;
using System.Management;
using System.Runtime.InteropServices;
using FormOA;
using System.Linq;
using System.Collections.Generic;

namespace Vovin.CmcLibNet
{

    /* it is recommended to declare an Interface to expose this class to COM
     * that way we can set [ClassInterface(ClassInterfaceType.None)]
     * and explicitly tell the compiler what interface to expose.
     * We explicitly define GUIDs for the same reason.
     * Also note that the assembly is not marked to be COM-visible on the assembly level
     * This ensures fine-grained control over the exposed interfaces
     * and reduces issues with generated GUIDs which may cause versioning problems.
     * see http://blogs.msdn.com/b/mbend/archive/2007/04/17/classinterfacetype-none-is-my-recommended-option-over-autodispatch-autodual.aspx
     */

    /// <summary>
    /// CmcLibNet is a .NET assembly that wraps the Commence API. 
    /// The primary goal for developing this library was to make it easier to communicate with Commence from Powershell.
    /// </summary>
    /// <remarks>
    /// <para>
    /// COM clients can use it by calling it with ProgId <c>'CmcLibNet.CommenceApp'</c>.
    /// </para>
    /// <para>COM applications such as Commence Item Detail Form scripts that use so-called 'late binding' can call the assembly thus:</para>
    /// <para>VBscript:</para>
    /// <code language="vbscript">Dim obj : Set obj = CreateObject("CmcLibNet.CommenceApp")</code>
    /// <para>When used from Commence Item Detail Forms or Office VBA, be sure to read this: <see cref="Close"/>.</para>
    /// </remarks>
    [ComVisible(true)] // make COM visible (overrides AssemblyInfo setting)
    [Guid("2F0DD17C-C020-4898-924A-82F4593DD569")] // explicitly set GUID so compiler doesn't generate a new one upon every build
    [ProgId("CmcLibNet.CommenceApp")] // custom ProgID, this is the name used when creating the object through COM
    [ClassInterface(ClassInterfaceType.None)] // tells compiler to not create a class interface (we'll supply our own interfaces).
    [ComDefaultInterface(typeof(ICommenceApp))] // explicitly define interface to expose to COM
    public class CommenceApp : ICommenceApp
    {
        private IRcwReleasePublisher rw = null;
        private readonly string PROCESS_NAME = "commence";
        /* _cmc is marked static to ensure that only a single reference to Commence ever exists
         Most calls to this assembly will be from Vovin.CmcLibNet.Database.CommenceDatabase
         That class uses the DB property of this class that returns this static field to access the Commence database.
         No other classes in this assembly ever create a reference to CommenceDB directly.
         The only way multiple references can be created is by referencing this assembly multiple times,
         which never makes sense unless Commence ever makes it possible to talk to multiple instances 
         and in which case this whole assembly needs redesigning.
         */
        //private static ICommenceDB _cmc = null;
        // 20190319
        private FormOA.ICommenceDB _cmc = null; // non-static

        #region Constructors

        /// <summary>
        /// Constructor.
        /// </summary>
        /// <exception cref="CommenceNotRunningException">Commence is not running (COM error: 0x80131500).</exception>
        /// <exception cref="CommenceMultipleInstancesException">Commence is running more than 1 instance (COM error: 0x80131500).</exception>
        /// <remarks>If the constructor fails because either Commence is not running,
        /// or Commence is running multiple instances in the same userprofile, an exception is thrown.
        /// COM users (VBScript, VBA, etc.) will not get a useful error message.
        /// This is a bit of a problem that I am unsure how to solve.
        /// We could incorporate a lazy initialization routine and call that as soon as any of the properties are called, 
        /// then return a more useful error to a consumer, but that is very messy from a .NET point of view.
        /// <para>Note that this behaviour may change in future releases.</para>
        /// <para>Note that if you want to execute a PowerShell of VBScript script, and Commence is started as administrator (or vice-versa),
        /// the script will prompt to start another instance of Commence. The userlevel of the user of the assembly must match the userlevel of the Commence instance.
        /// </para>
        /// </remarks>
        public CommenceApp()
        {
            /* Commence does not register itself in the ROT (Running Object Table)
            // * Therefore, we can not tell which database to talk to from COM
            // * This is a serious problem that Commence refuses to fix.
            // * Most of the third-party utilities built for Commence do not take this into account,
            // * they will happily try to talk to the first instance they encounter.
            // * Calling CommenceDB when no commence.exe instance is running will fire up commence.exe and an arbitrary database, usually the last one opened.
            // * This can be very much unwanted, especially in TS-like environments, so I tried to eliminate that possibility as much as possible
            // * This assembly will simply not run if commence.exe is not running
            // * The downside of this is that COM users will get an unintelligible error when no or multiple instances are running.
            // * It will also not run if multiple commence.exe processes were started by the same user
            // * It *should* run when different users started the commence.exe process
            // * In that case, it *should* talk to the instance fired by the same user who calls the assembly
            // * This behaviour may be subject to errors when Commence.DB has a non-standard DCOM settings defined.
            // * This has yet to be tested.
            // */
            rw = new RcwReleasePublisher();
            rw.RCWRelease += RCWReleaseHandler;
        }
        #endregion

        #region Properties

        ///// <summary>
        ///// Exposes the 'raw' Commence database (FormOA.ICommenceDB) object.
        ///// For internal use only.
        ///// </summary>
        //internal static FormOA.ICommenceDB DB
        //{
        //    get
        //    {
        //        return _cmc;
        //    }
        //}

        internal FormOA.ICommenceDB Db
        {
            get
            {
                if (_cmc != null) { return _cmc; }
                else
                {
                    switch (ProcessCountInCurrentSession(PROCESS_NAME))
                    {
                        case 0:
                            // Sorry COM users, but you'll have to deal with a cryptic error when Commence is not running.
                            throw new CommenceNotRunningException("Commence is not running.");
                        case 1:
                            _cmc = new CommenceDB(); // Note: if the assembly is called as Administrator, this will create a new instance.
                            break;
                        default:
                            throw new CommenceMultipleInstancesException("Multiple instances of Commence are running in this session. Make sure only 1 instance is running.");
                    }
                    return _cmc;
                }
            }
        }

        /// <inheritdoc />
        public string ExePath
        {
            // We account for multiple processes even though this assembly won't work then.
            // This just to be sure should we ever change our mind.
            get
            {
                List<string> retval = ProcessPaths(PROCESS_NAME);
                switch (retval.Count())
                {
                    case 0: // no process
                        return string.Empty;
                    case 1: // running once
                        return retval[0];
                    default: // running multiple times
                        if (retval.Distinct().Count() != retval.Count()) // Commence was installed in multiple locations
                        {
                            return "Unable to determine";
                        }
                        else
                        {
                            return retval[0]; // commence is running same exe multiple times, just return first
                        }
                }
            }
        }
 
        /// <inheritdoc />
        public string Name
        {
            get { return _cmc.Name; }
        }
        /// <inheritdoc />
        public string Path
        {
            get { return _cmc.Path; }
        }
        /// <inheritdoc />
        public string Version
        { 
            get { return _cmc.Version; }
        }
        /// <inheritdoc />
        public string VersionExt
        {
            get { return _cmc.VersionExt; }
        }
        /// <inheritdoc />
        public string RegisteredUser
        {
            get { return _cmc.RegisteredUser; }
        }
        #endregion

        #region Methods
        /// <summary>
        /// Running process counter for processes in current session.
        /// </summary>
        /// <param name="processName">Name of the process.</param>
        /// <returns>Number of running process instances.</returns>
        private int ProcessCountInCurrentSession(string processName)
        {
            Process[] p = Process.GetProcessesByName(processName);
            int currentSessionID = Process.GetCurrentProcess().SessionId;
            Process[] sameAsthisSession = (from c in p where c.SessionId == currentSessionID select c).ToArray();
            return sameAsthisSession.Length;
        }

        /// <summary>
        /// Get paths of process(es).
        /// </summary>
        /// <param name="processName">processnames to list.</param>
        /// <returns>List of paths of processes of supplied name.</returns>
        private List<string> ProcessPaths(string processName)
        {
            List<string> retval = new List<string>();
            var wmiQueryString = "SELECT ProcessId, ExecutablePath, CommandLine FROM Win32_Process";
            using (var searcher = new ManagementObjectSearcher(wmiQueryString))
            using (var results = searcher.Get())
            {
                var query = from p in Process.GetProcessesByName(processName)
                            join mo in results.Cast<ManagementObject>()
                            on p.Id equals (int)(uint)mo["ProcessId"]
                            select new
                            {
                                Process = p,
                                Path = (string)mo["ExecutablePath"],
                                CommandLine = (string)mo["CommandLine"],
                            };
                foreach (var item in query)
                {
                    retval.Add(item.Path);
                }
            }
            return retval;
        }

        /// <inheritdoc />
        public void Close()
        {
            // releases the FormOA COM references.
            rw.ReleaseRCWReferences();
        }

        // event handling method
        internal void RCWReleaseHandler(object sender, EventArgs e)
        {
            if (_cmc != null)
            {
                //Marshal.FinalReleaseComObject(_cmc); // kill all COM references to Commence.DB. No calls to Commence can be made after this.
                while (Marshal.ReleaseComObject(_cmc) > 0) { }; // kill current references to Commence.DB. No calls to Commence can be made after this.
                _cmc = null;
            }
        }
        #endregion
    }
}
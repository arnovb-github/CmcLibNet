using System;
using System.Diagnostics;
using System.Management;
using System.Runtime.InteropServices;
using FormOA;
using System.Linq;
using System.Collections.Generic;

namespace Vovin.CmcLibNet
{

    #region Enumerations
    /// <summary>
    /// Commence option flags.
    /// </summary>
    [ComVisible(true)]
    [GuidAttribute("1382A6E0-19E5-41c1-9BE7-10545F538DBE")]
    [FlagsAttribute]
    public enum CmcOptionFlags
    {
        //If bitwise operators are making your head spin,
        //.NET 4.0 has introduced the method HasFlag which can be used as follows: .HasFlag(enum.value)
        /// <summary>
        /// Default flag.
        /// </summary>
        Default = 0x0,
        /// <summary>
        /// Indicates fieldnames should be used, columnnames ignored.
        /// </summary>
        Fieldname = 0x0001,
        /// <summary>
        /// Used to request all fields from a category.
        /// </summary>
        All = 0x0002,
        /// <summary>
        /// Mark item as shared when adding items.
        /// </summary>
        Shared = 0x0004,
        /// <summary>
        /// Changes from 3Com Palm Pilot.
        /// </summary>
        [Obsolete]
        PalmPilot = 0x0008,
        /// <summary>
        /// Make Commence return data in canonical (i.e. consistent) format.
        /// <remarks>
        /// <list type="table">
        /// <listheader><term>Datatype</term><description>Format notes</description></listheader>
        /// <item><term>Date</term><description>yyyymmdd</description></item>
        /// <item><term>Time</term><description>hh:mm military time, 24 hour clock.</description></item>
        /// <item><term>Number</term><description>123456.78, no thousand separator, period for decimal delimiter. Note that when a field is defined as 'Show as currency' in the Commence UI, numerical values are prepended with a '$' sign.</description></item>
        /// <item><term>CheckBox</term><description>TRUE or FALSE.</description></item>
        /// </list>
        /// </remarks>
        /// </summary>
        Canonical = 0x0010,
        /// <summary>
        /// Allows for the Agent subsystem to distinguish between Internet/Intranet database operations.
        /// </summary>
        Internet = 0x0020,
        /// <summary>
        /// (Undocumented by Commence) Make Commence return THIDs instead of Name field values.
        /// Use *RowSet.GetRowID() on a cursor defined with thids flag to get a row's THID
        /// </summary>
        UseThids = 0x0100,
        /// <summary>
        /// (Undocumented by Commence) Unknown.
        /// </summary>
        IgnoreSyncCondition= 0x0200
    }
    #endregion

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
    /// CmcLibNet is a .NET assembly that wraps the Commence API. The primary goal for developing this library was to make it easier to communicate with Commence from Powershell.
    /// </summary>
    /// <remarks>
    /// <para>
    /// COM clients can use it by calling it with ProgId <c>'CmcLibNet.CommenceApp'</c>.
    /// Some convenience methods can only be used from .NET applications, but all functionality as defined in the Commence API is available from COM.
    /// </para>
    /// <para>COM applications such as Commence Item Detail Form scripts that use so-called 'late binding' can call the assembly thus:</para>
    /// <para>VBscript:</para>
    /// <code language="vbscript">Dim obj : Set obj = CreateObject("CmcLibNet.CommenceApp")</code>
    /// <para>When used from Commence Item Detail Forms or Office VBA, be sure to read this: <see cref="Close"/>.</para>
    /// </remarks>
    [ComVisible(true)] // make COM visible (overrides AssemblyInfo setting)
    [GuidAttribute("2F0DD17C-C020-4898-924A-82F4593DD569")] // explicitly set GUID so compiler doesn't generate a new one upon every build
    [ProgId("CmcLibNet.CommenceApp")] // custom ProgID, this is the name used when creating the object through COM
    [ClassInterface(ClassInterfaceType.None)] // tells compiler to not create a class interface (we'll supply our own interfaces).
    [ComDefaultInterface(typeof(ICommenceApp))] // explicitly define interface to expose to COM
    public class CommenceApp : ICommenceApp
    {
        private IRCWReleasePublisher rw = null;
        private const string PROCESS_NAME = "commence";
        /* _cmc is marked static to ensure that only a single reference to Commence ever exists
         Most calls to this assembly will be from Vovin.CmcLibNet.Database.CommenceDatabase
         That class uses the DB property of this class that returns this static field to access the Commence database.
         No other classes in this assembly ever create a reference to CommenceDB directly.
         The only way multiple references can be created is by referencing this assembly multiple times,
         which never makes sense unless Commence ever makes it possible to talk to multiple instances 
         and in which case this whole assembly needs redesigning.
         */
        private static FormOA.ICommenceDB _cmc = null;
        private Database.ICommenceDatabase _database = null;
        private Export.IExportEngine _exportengine = null;
        private Services.IServices _services = null;

        #region Constructors

        /// <summary>
        /// Constructor.
        /// </summary>
        /// <exception cref="CommenceNotRunningException">Commence is not running (COM error: 0x80131500).</exception>
        /// <exception cref="CommenceMultipleInstancesException">Commence is running more than 1 instance (COM error: 0x80131500).</exception>
        /// <remarks>If the constructor fails because either Commence is not running,
        /// or Commence is running multiple instances, an exception is thrown.
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
            /* Commence does not register itself in the ROT (Remote Object Table)
             * Therefore, we can not tell which database to talk to from COM
             * This is a serious problem that Commence refuses to fix.
             * Most of the third-party utilities built for Commence do not take this into account,
             * they will happily try to talk to the first instance they encounter.
             * Calling CommenceDB when no commence.exe instance is running will fire up commence.exe and an arbitrary database, usually the last one opened.
             * This can be very much unwanted, especially in TS-like environments, so I tried to eliminate that possibility as much as possible
             * This assembly will simply not run if commence.exe is not running
             * The downside of this is that COM users will get an unintelligible error when no or multiple instances are running.
             * It will also not run if multiple commence.exe processes were started by the same user
             * It *should* run when different users started the commence.exe process
             * In that case, it *should* talk to the instance fired by the same user who calls the assembly
             * This behaviour may be subject to errors when Commence.DB has a non-standard DCOM settings defined.
             * This has yet to be tested.
             */
            switch (ProcessCountInCurrentSession(PROCESS_NAME))
            {
                case 0:
                    // showing a messagebox from a DLL is simply a Bad Idea.
                    // AutoClosingMessageBox.Show("Commence is not running.", "Initialization error", 3000);
                    // Sorry COM users, but you'll have to deal with a cryptic error when Commence is not running.
                    throw new CommenceNotRunningException("Commence is not running.");
                case 1:
                    _cmc = new CommenceDB(); // Note: if the assembly is called as Administrator, this will create a new instance.
                    break;
                default:
                    //_cmc = new CommenceDB(); // get whatever Commence reference Windows gives us.
                    throw new CommenceMultipleInstancesException("Multiple instances of Commence are running in this session. Make sure only 1 instance is running.");
            }
            rw = new RCWReleasePublisher();
            rw.RCWRelease += RCWReleaseHandler;
        }
        #endregion

        #region Properties

        /// <summary>
        /// Exposes the 'raw' Commence database (FormOA.ICommenceDB) object.
        /// For internal use only.
         /// </summary>
        protected internal static FormOA.ICommenceDB DB
        {
            get
            {
                return _cmc;
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
                        return "No process";
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

        // the next properties expose the 'subdivisions' of CmcLibNet.
        /// <inheritdoc />
        public Database.ICommenceDatabase Database
        {
            get 
            {
                if (_database == null)
                {
                    _database = new Database.CommenceDatabase();
                }
                    return _database;
            }
        }

        /// <inheritdoc />
        public Export.IExportEngine Export
        {
            get 
            {
                if (_exportengine == null)
                {
                    _exportengine = new Export.ExportEngine();
                }
                return _exportengine;
            }
        }

        /// <inheritdoc />
        public Services.IServices Services
        {
            get 
            {
                if (_services == null)
                {
                    _services = new Services.Services();
                }
                return _services;
            }
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
            _cmc = null; // overkill
        }

        // event handling method
        internal void RCWReleaseHandler(object sender, EventArgs e)
        {
            if (_cmc != null)
            {
                Marshal.FinalReleaseComObject(_cmc); // kill all COM references to Commence.DB. No calls to Commence can be made after this.
            }
        }
        #endregion
    }
}
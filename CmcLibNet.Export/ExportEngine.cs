﻿using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Vovin.CmcLibNet;
using Vovin.CmcLibNet.Database;

namespace Vovin.CmcLibNet.Export
{
    /// <summary>
    /// Delegate for the ExportProgressChanged event.
    /// </summary>
    /// <param name="sender">sender object.</param>
    /// <param name="e">ExportProgressChangedArgs.</param>
    [ComVisible(false)] // no need to expose to COM
    public delegate void ExportProgressChangedHandler(object sender, ExportProgressChangedArgs e);

    /// <summary>
    /// Delegate for the CommenceRowsRead event.
    /// </summary>
    /// <param name="sender">sender object.</param>
    /// <param name="e">CommenceRowsReadArgs.</param>
    [ComVisible(false)] // no need to expose to COM
    public delegate void CommenceRowsReadHandler(object sender, CommenceRowsReadArgs e);

    /// <summary>
    /// Export engine for exporting data from Commence.
    /// <remarks>
    /// .NET users can simply create an instance,
    /// COM clients can instantiate this by using the ProgId <c>CmcLibNet.Export</c>.
    /// So from VBScript you would do:
    /// <code language="vbscript">Dim export : Set export = CreateObject("CmcLibNet.Export")
    /// '.. do stuff ..
    /// export.Close
    /// Set export = Nothing</code>
    /// <para>When used fom a Commence Item Detail Form or Office VBA, be sure to read up on the <see cref="Close()"/> method.</para>
    /// </remarks>
    /// </summary>
    [ComVisible(true)]
    [GuidAttribute("298DA1F6-9A8D-4BB2-A7EC-F776013F440A")]
    [ProgId("CmcLibNet.Export")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IExportEngine))]
    // The ComSourceInterfacesAttribute is the key to
    // exposing COM events from a managed class.
    [ComSourceInterfaces(typeof(IExportEngineCOMEvents))]
    public class ExportEngine : IExportEngine //note that we do not inherit from IExportEngineEvents!
    {
        /// <summary>
        /// ExportProgressChanged event.
        /// </summary>
        public event ExportProgressChangedHandler ExportProgressChanged;
        /// <summary>
        /// CommenceRowsRead event.
        /// </summary>
        public event CommenceRowsReadHandler CommenceRowsRead;
        private BaseWriter _writer = null;
        private IExportSettings _settings = null;
        private readonly ICommenceDatabase _db = null;

        #region Constructors
        /// <summary>
        /// Constructor.
        /// </summary>
        public ExportEngine()
        {
            _db = new CommenceDatabase(); // CommenceDatabase takes care of getting the reference to Commence
        }
        /// <summary>
        /// Destructor
        /// </summary>
        ~ExportEngine()
        {
            // is unsubscribing needed or even useful?
            if (_writer != null) {
                _writer.ExportProgressChanged -= this.ExportProgressChanged;
                _writer.CommenceRowsRead -= this.CommenceRowsRead;
            }
        }
        #endregion

        #region Methods
        /// <summary>
        /// This method is called all exports methods, it takes care of all the actual exports.
        /// Other export methods are really just preparation routines for calling this method.
        /// </summary>
        /// <param name="cur">ICommenceCursor</param>
        /// <param name="fileName">file name</param>
        /// <param name="settings">IExportSettings</param>
        internal void ExportCursor(Database.ICommenceCursor cur, string fileName, IExportSettings settings)
        {

            cur.MaxFieldSize = this.Settings.MaxFieldSize;

            // exporting may take a long time, and the system may go into power saving mode
            // this is annoying, so we tell the system not to go into sleep/hibernation
            // this may or may not be a good idea...
            PowerSavings ps = new PowerSavings();

            try
            {
                ps.EnableConstantDisplayAndPower(true, "Performing lengthy Commence export");
                if (settings.NestConnectedItems)
                {
                    /* Nested exports only support XML and JSON for now.
                     * Internally, all returned Commence data are first put into a ADO.NET dataset.
                     * Then that dataset is use to generate the XML or JSON using native ADO.NET or NewtonSoft serialization methods.
                     * Transforming Commence data to an ADO.NET dataset takes a bit of time and is not 100% fail-safe yet.
                     * The time to create it however is negligible compared to the reading of Commence cursor.
                     * In future versions, it may be that all export functions use a dataset.
                     * However, creating an in-memory dataset on top of the Commence memory stuff impacts RAM usage.
                     */
                    if (settings.ExportFormat == ExportFormat.Xml || settings.ExportFormat == ExportFormat.Json)
                    {
                        // we're okay for now but we will not win coding prizes for this ;->
                    }
                    else
                    {
                        // commence.exe will be released even when this error is thrown when using compiled assembly, not in debugger.
                        throw new NotSupportedException("The requested export format cannot be nested.");
                    }
                    using (_writer = new AdoNetWriter(cur, settings))
                    {
                        _writer.ExportProgressChanged += this.HandleProgressChanged; // not entirely accurate because there is post-processing after this event has finished.
                        _writer.CommenceRowsRead += this.CommenceRowsRead;
                        _writer.WriteOut(fileName);
                    }
                }
                else
                {
                    using (_writer = this.GetExportWriter(cur, settings))
                    {
                        _writer.ExportProgressChanged += this.HandleProgressChanged; // note that we never unsubscribe, but the writer will be disposed of.
                        _writer.CommenceRowsRead += this.CommenceRowsRead;
                        _writer.WriteOut(fileName);
                    }
                }
            } // try
            finally
            {
                ps.EnableConstantDisplayAndPower(false);
            }
        }

        /// <inheritdoc />
        public void ExportView(string viewName, string fileName, IExportSettings settings = null)
        {
            string _viewName = string.Empty;

            if (!String.IsNullOrEmpty(viewName))
            {
                _viewName = viewName;
            }
            else
            {
                //_db = new CommenceDatabase();
                Database.IActiveViewInfo av = _db.GetActiveViewInfo();
                if (av != null && String.IsNullOrEmpty(av.Field)) // view is active and it is not an item detail form
                {
                    _viewName = av.Name;
                }
                else
                {
                    throw new CommenceCOMException("No active view could be exported. Either no data view is currently active, or it is an item detail form.");
                }
            }
            if (settings != null) { this.Settings = settings; } // store custom settings
            ICommenceCursor cur = _db.GetCursor(_viewName, CmcCursorType.View, (this.Settings.UseThids) ? CmcOptionFlags.UseThids : CmcOptionFlags.Default);
            ExportCursor(cur, fileName, this.Settings);
        }

        /// <inheritdoc />
        public void ExportCategory(string categoryName, string fileName, IExportSettings settings = null)
        {
            if (settings != null) { this.Settings = settings; } // use custom settings if supplied
            CmcOptionFlags flags = (this.Settings.UseThids) ? CmcOptionFlags.UseThids : CmcOptionFlags.Default;
            if (this.Settings.SkipConnectedItems && this.Settings.CustomHeaders == null)
            {
                // User requested we skip connections.
                // A default cursor on a category contains all fields including connections.
                // The data receiving routines will ignore them, but they will be read unless we do not include them in our cursor
                // We optimize here by only including direct fields in the cursor
                // WAIT...we cannot do that because of potential custom headers
                // fuck.
                // okay, we only optimize when no custom headers were passed
                flags = flags | CmcOptionFlags.IgnoreSyncCondition;
                using (ICommenceCursor cur = _db.GetCursor(categoryName, Database.CmcCursorType.Category, flags))
                {
                    string[] fieldNames = _db.GetFieldNames(categoryName).ToArray();
                    cur.Columns.AddDirectColumns(fieldNames);
                    cur.Columns.Apply();
                    ExportCursor(cur, fileName, this.Settings);
                }
            }
            else
            {
            flags = flags | CmcOptionFlags.All | CmcOptionFlags.IgnoreSyncCondition; // slap on some more flags
            using (ICommenceCursor cur = _db.GetCursor(categoryName, Database.CmcCursorType.Category,flags))
                {
                    ExportCursor(cur, fileName, this.Settings);
                }
            }
        }
        /// <summary>
        /// Factory method for creating the required export writer object for a cursor export.
        /// </summary>
        /// <param name="cursor">Database.ICommenceCursor reference.</param>
        /// <param name="settings">Settings object.</param>
        /// <returns>Derived BaseDataWriter object.</returns>
        internal BaseWriter GetExportWriter(Database.ICommenceCursor cursor, IExportSettings settings)
        {
            switch (settings.ExportFormat)
            {
                case ExportFormat.Text:
                    settings.SplitConnectedItems = false;
                    _writer = new TextWriter(cursor, settings);
                    break;
                case ExportFormat.Html:
                    settings.SplitConnectedItems = false;
                    _writer = new HTMLWriter(cursor, settings);
                    break;
                case ExportFormat.Xml:
                    _writer = new XMLWriter(cursor, settings);
                    break;
                case ExportFormat.Json:
                    _writer = new JSONWriter(cursor, settings);
                    break;
                case ExportFormat.Excel:
                    _writer = new ExcelWriter(cursor, settings);
                    break;
                case ExportFormat.Event:
                    _writer = new EventWriter(cursor, settings);
                    break;
                case ExportFormat.GoogleSheets:
                    // will probably always be too slow
                    throw new NotImplementedException();
            }
            return _writer;
        }

        // we need some mechanism to tell consumers about the progress of our export
        // the writer classes are not exposed, so we need to capture the events of those classes,
        // and then raise an event that consumers can subscribe to.
        /// <summary>
        /// Event handler for the CommenceRowsRead event.
        /// </summary>
        /// <param name="sender">sender object.</param>
        /// <param name="e">CommenceRowsReadArgs object.</param>
        public virtual void HandleCommenceRowsRead(object sender, CommenceRowsReadArgs e)
        {
            // just re-raise event
            OnCommenceRowsRead(e);
        }
        /// <summary>
        /// Raise the CommenceRowsRead event.
        /// </summary>
        /// <param name="e">CommenceRowsReadArgs.</param>
        protected virtual void OnCommenceRowsRead(CommenceRowsReadArgs e)
        {
            try
            {
                CommenceRowsReadHandler handler = CommenceRowsRead;
                Delegate[] eventHandlers = handler.GetInvocationList();
                foreach (Delegate currentHandler in eventHandlers)
                {
                    CommenceRowsReadHandler currentSubscriber = (CommenceRowsReadHandler)currentHandler;
                    try
                    {
                        currentSubscriber(this, e);
                    }
                    catch { }
                }
            }
            catch { }
        }

        /// <summary>
        /// Event handler for the ExportProgressChanged event.
        /// </summary>
        /// <param name="sender">sender object.</param>
        /// <param name="e">ExportProgressChangedArgs object.</param>
        public virtual void HandleProgressChanged(object sender, ExportProgressChangedArgs e)
        {
            // just re-raise event
            OnExportProgressChanged(e);
        }

        /// <summary>
        /// Raise the ProgressChanged event.
        /// </summary>
        /// <param name="e">ExportProgressChangedArgs.</param>
        protected virtual void OnExportProgressChanged(ExportProgressChangedArgs e)
        {
            try
            {
                ExportProgressChangedHandler handler = ExportProgressChanged;
                Delegate[] eventHandlers = handler.GetInvocationList();
                foreach (Delegate currentHandler in eventHandlers)
                {
                    ExportProgressChangedHandler currentSubscriber = (ExportProgressChangedHandler)currentHandler;
                    try
                    {
                        currentSubscriber(this, e);
                    }
                    catch { }
                }
            }
            catch { }
        }

        /// <inheritdoc />
        public void Close()
        {
            if (_db != null)
            {
                _db.Close();
            }
        }
        #endregion

        #region Properties

        /// <inheritdoc />
        public IExportSettings Settings
        {
            get
            {
                if (_settings == null) { _settings = new ExportSettings(); } // return default settings object if none was provided.
                return _settings;
            }
            internal set
            {
                _settings = value;
            }
        }

        #endregion
    }
}
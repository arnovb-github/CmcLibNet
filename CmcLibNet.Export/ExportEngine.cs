using System;
using System.Runtime.InteropServices;
using Vovin.CmcLibNet.Database;

namespace Vovin.CmcLibNet.Export
{
    /// <summary>
    /// Delegate for the ExportProgressChanged event to deal with batches of rows.
    /// For use by external assemblies.
    /// </summary>
    /// <param name="sender">sender object.</param>
    /// <param name="e">ExportProgressAsStringChangedArgs.</param>
    [ComVisible(false)] // there is a separate interface for COM
    public delegate void ExportProgressAsStringChangedHandler(object sender, ExportProgressAsStringChangedArgs e);

    /// <summary>
    /// Delegate for the ExportCompleted event to deal with batches of rows.
    /// For use by external assemblies.
    /// </summary>
    /// <param name="sender">sender object.</param>
    /// <param name="e">ExportCompleteArgs.</param>
    [ComVisible(false)] // there is a separate interface for COM
    public delegate void ExportCompletedHandler(object sender, ExportCompleteArgs e);

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
    [Guid("298DA1F6-9A8D-4BB2-A7EC-F776013F440A")]
    [ProgId("CmcLibNet.Export")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IExportEngine))]
    // The ComSourceInterfacesAttribute is the key to
    // exposing COM events from a managed class.
    [ComSourceInterfaces(typeof(IExportEngineCOMEvents))]
    public class ExportEngine : IExportEngine //note that we do not inherit from IExportEngineEvents!
    {
        /// <summary>
        /// ExportProgressChanged event for outside assemblies.
        /// </summary>
        public event ExportProgressAsStringChangedHandler ExportProgressChanged;
        /// <summary>
        /// ExportCompleted event for outside assemblies.
        /// </summary>
        public event ExportCompletedHandler ExportCompleted;
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
                UnsubscribeToWriterEvents(_writer);
            }
        }
        #endregion

        #region Export methods
        /// <summary>
        /// This method is called by all exports methods, it takes care of all the actual exports.
        /// Other export methods are really just preparation routines for calling this method.
        /// </summary>
        /// <param name="cur">ICommenceCursor</param>
        /// <param name="fileName">file name</param>
        /// <param name="settings">IExportSettings</param>
        internal void ExportCursor(ICommenceCursor cur, string fileName, IExportSettings settings)
        {
            // exporting may take a long time, and the system may go into power saving mode
            // this is annoying, so we tell the system not to go into sleep/hibernation
            // this may or may not be a good idea...
            PowerSavings ps = new PowerSavings();

            try
            {
                ps.EnableConstantDisplayAndPower(true, "Performing lengthy Commence export");

                cur.MaxFieldSize = this.Settings.MaxFieldSize; // remember setting this size greatly impacts memory usage!
                if (settings.NestConnectedItems && settings.ExportFormat == ExportFormat.Json)
                {
                    /* Nested exports for Json.
                     * Internally, all returned Commence data are first put into a ADO.NET dataset.
                     * Then that dataset is serialzed yo the JSON .
                     */
                    using (_writer = new AdoNetWriter(cur, settings))
                    {
                        SubscribeToWriterEvents(_writer);
                        _writer.WriteOut(fileName);
                    }
                }
                else
                {
                    using (_writer = this.GetExportWriter(cur, settings))
                    {
                        SubscribeToWriterEvents(_writer);
                        switch (settings.ExportFormat)
                        {
                            case ExportFormat.Excel:
                                _writer.WriteOut(fileName);
                                break;
                            default:
                                _writer.WriteOut(fileName);
                                break;
                        }
                    }
                }
            }
            finally
            {
                ps.EnableConstantDisplayAndPower(false);
            }
        }

        private void SubscribeToWriterEvents(BaseWriter w)
        {
            w.ExportProgressChanged += (s, e) => HandleExportProgressChanged(s, e);
            w.ExportCompleted += (s, e) => HandleExportCompleted(s, e);
        }

        private void UnsubscribeToWriterEvents(BaseWriter w)
        {
            w.ExportProgressChanged -= (s, e) => HandleExportProgressChanged(s, e);
            w.ExportCompleted -= (s, e) => HandleExportCompleted(s, e);
        }

        /// <inheritdoc />
        public void ExportView(string viewName, string fileName, IExportSettings settings = null)
        {
            string _viewName = string.IsNullOrEmpty(viewName) ? GetActiveViewName() : viewName;
            if (settings != null) { this.Settings = settings; } // store custom settings
            using (ICommenceCursor cur = _db.GetCursor(_viewName, CmcCursorType.View, (this.Settings.UseThids) ? CmcOptionFlags.UseThids : CmcOptionFlags.Default))
            {
                ExportCursor(cur, fileName, this.Settings);
            }
        }

        /// <inheritdoc />
        public void ExportCategory(string categoryName, string fileName, IExportSettings settings = null)
        {
            if (settings != null) { this.Settings = settings; } // use custom settings if supplied
            CmcOptionFlags flags = (this.Settings.UseThids) ? CmcOptionFlags.UseThids : CmcOptionFlags.Default | CmcOptionFlags.IgnoreSyncCondition;
            if (this.Settings.SkipConnectedItems && this.Settings.HeaderMode != HeaderMode.CustomLabel)
            {
                // User requested we skip connections.
                // A default cursor on a category contains all fields including connections.
                // The data receiving routines will ignore them, but they will be read unless we do not include them in our cursor
                // We optimize here by only including direct fields in the cursor
                using (ICommenceCursor cur = GetCursorWithJustDirectFields(categoryName, flags))
                {
                    ExportCursor(cur, fileName, this.Settings);
                }
            }
            else
            {
            flags = flags | CmcOptionFlags.All; // slap on some more flags
            using (ICommenceCursor cur = _db.GetCursor(categoryName, Database.CmcCursorType.Category,flags))
                {
                    ExportCursor(cur, fileName, this.Settings);
                }
            }
        }

        private string GetActiveViewName()
        {
            string retval = string.Empty;
            Database.IActiveViewInfo av = _db.GetActiveViewInfo();
            if (av != null && String.IsNullOrEmpty(av.Field)) // view is active and it is not an item detail form
            {
                retval = av.Name;
            }
            else
            {
                throw new CommenceCOMException("Could not determine what view is active in Commence.");
            }
            return retval;
        }

        private ICommenceCursor GetCursorWithJustDirectFields(string categoryName, CmcOptionFlags flags )
        {
            ICommenceCursor cur = _db.GetCursor(categoryName, Database.CmcCursorType.Category, flags);
            string[] fieldNames = _db.GetFieldNames(categoryName).ToArray();
            cur.Columns.AddDirectColumns(fieldNames);
            cur.Columns.Apply();
            return cur;
        }
        /// <summary>
        /// Factory method for creating the required export writer object for a cursor export.
        /// </summary>
        /// <param name="cursor">Database.ICommenceCursor .</param>
        /// <param name="settings">Settings object.</param>
        /// <returns>Derived BaseDataWriter object.</returns>
        /// <remarks>Defaults to XML.</remarks>
        internal BaseWriter GetExportWriter(ICommenceCursor cursor, IExportSettings settings)
        {
            switch (settings.ExportFormat)
            {
                case ExportFormat.Text:
                    settings.SplitConnectedItems = false;
                    return new TextWriter(cursor, settings);
                case ExportFormat.Html:
                    settings.SplitConnectedItems = false;
                    return new HtmlWriter(cursor, settings);
                case ExportFormat.Xml:
                    return new XmlWriter(cursor, settings);
                case ExportFormat.Json:
                    return new JsonWriter(cursor, settings);
                case ExportFormat.Excel:
                    //_writer = new ExcelWriterUsingXml(cursor, settings);
                    settings.XSDCompliant = true;
                    settings.SplitConnectedItems = false;
                    return new ExcelWriterUsingOpenXml(cursor, settings);
                    //return new ExcelWriterUsingOleDb(cursor, settings);
                case ExportFormat.Event:
                    return new EventWriter(cursor, settings);
                case ExportFormat.GoogleSheets:
                    // will probably always be too slow
                    throw new NotImplementedException();
                default:
                    return  new XmlWriter(cursor, settings);
            }
        }
#endregion

#region Event methods
        // we need some mechanism to tell consumers about the progress of our export
        // the writer classes are not exposed, so we need to capture the events of those classes,
        // and then raise an event that consumers can subscribe to.
        /// <summary>
        /// Event handler for the ExportProgressChanged event.
        /// </summary>
        /// <param name="sender">sender object.</param>
        /// <param name="e">ExportProgressAsJsonChangedArgs object.</param>
        public virtual void HandleExportProgressChanged(object sender, ExportProgressAsStringChangedArgs e)
        {
            // just re-raise event
            OnExportProgressChanged(e); // outside assemblies can subscribe to this
        }

        /// <summary>
        /// Event handler for the ExportCompleted event
        /// </summary>
        /// <param name="sender">sender object.</param>
        /// <param name="e">ExportCompleteArgs</param>
        public virtual void HandleExportCompleted(object sender, ExportCompleteArgs e)
        {
            OnExportCompleted(e);
        }
        /// <summary>
        /// Raise the ExportProgressChanged event.
        /// </summary>
        /// <param name="e">ExportProgressAsStringChangedArgs object.</param>
        protected virtual void OnExportProgressChanged(ExportProgressAsStringChangedArgs e)
        {
            ExportProgressChanged?.Invoke(this, e);
        }

        /// <summary>
        /// Raises the ExportCompleted event
        /// </summary>
        /// <param name="e">ExportCompleteArgs</param>
        protected virtual void OnExportCompleted(ExportCompleteArgs e)
        {
            ExportCompleted?.Invoke(this, e);
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

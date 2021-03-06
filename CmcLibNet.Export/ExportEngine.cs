﻿using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Vovin.CmcLibNet.Database;
using Vovin.CmcLibNet.Database.Metadata;

namespace Vovin.CmcLibNet.Export
{
    /// <summary>
    /// Delegate for the ExportProgressChanged event to deal with batches of rows.
    /// For use by external assemblies.
    /// </summary>
    /// <param name="sender">sender object.</param>
    /// <param name="e">ExportProgressChangedArgs.</param>
    [ComVisible(false)] // there is a separate interface for COM
    public delegate void ExportProgressChangedHandler(object sender, ExportProgressChangedArgs e);

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
    /// </summary>
    /// <remarks>
    /// <para>Please note: the export engine tries to treat connected items as separate entities where possible
    /// - it is one of its core features.
    /// It relies on delimiters baked into the Commence API (either newline <c>\n</c> (not <c>\r\n</c>!) or a comma, 
    /// depending on the situation.).
    /// The Commence Text field, when set to Large, and the URL field *can* contain embedded newline characters.
    /// When exporting those field-types, you may get unexpected results if they do.
    /// In that case you options are either to simply not include them in your export,
    /// or to disable the splitting of connected items.<seealso cref="IExportSettings.SplitConnectedItems"/></para>
    /// <para>Usage: .NET users can simply create an instance,
    /// COM clients can instantiate this by using the ProgId <c>CmcLibNet.Export</c>.
    /// VBScript example:
    /// <code language="vbscript">Dim export : Set export = CreateObject("CmcLibNet.Export")
    /// '.. do stuff ..</code>
    /// Notice that contrary to e.g. <see cref="ICommenceCursor"/> or <see cref="ICommenceDatabase"/>,
    /// the <see cref="IExportEngine"/> does not require explicit closing. <seealso cref="IExportEngine.Close"/></para>
    /// </remarks>
    [ComVisible(true)]
    [Guid("298DA1F6-9A8D-4BB2-A7EC-F776013F440A")]
    [ProgId("CmcLibNet.Export")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IExportEngine))]
    // The ComSourceInterfacesAttribute is the key to
    // exposing COM events from a managed class.
    [ComSourceInterfaces(typeof(IExportEngineCOMEvents))]
    public class ExportEngine : IExportEngine
    {
        /// <summary>
        /// ExportProgressChanged event raised when (batch of) Commence data has been read.
        /// </summary>
        public event ExportProgressChangedHandler ExportProgressChanged;
        /// <summary>
        /// ExportCompleted event raised when Commence data reading has completed.
        /// </summary>
        public event ExportCompletedHandler ExportCompleted;
        private BaseWriter _writer;
        private IExportSettings _settings;

        #region Constructors
        /// <summary>
        /// Constructor.
        /// </summary>
        public ExportEngine() { }
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
                ps.EnableConstantDisplayAndPower(true, "Performing time-consuming Commence export");
                cur.MaxFieldSize = this.Settings.MaxFieldSize; // remember setting this size greatly impacts memory usage!

                using (_writer = this.GetExportWriter(cur, settings))
                {
                    SubscribeToWriterEvents(_writer);
                    _writer.WriteOut(fileName);
                }

            }
            finally
            {
                ps.EnableConstantDisplayAndPower(false); // CODE SMELL, may overwrite things
                UnsubscribeToWriterEvents(_writer);
                _writer?.Dispose();
            }
        }

        /// <inheritdoc />
        public void ExportView(string viewName, string fileName, IExportSettings settings = null)
        {

            if (settings != null) { this.Settings = settings; } // use user-supplied settings
            
            using (var db = new CommenceDatabase())
            {
                string _viewName = string.IsNullOrEmpty(viewName) ? GetActiveViewName(db) : viewName;
                using (ICommenceCursor cur = db.GetCursor(_viewName, CmcCursorType.View, this.Settings.UseThids
                    ? CmcOptionFlags.UseThids
                    : CmcOptionFlags.Default))
                {
                    ExportCursor(cur, fileName, this.Settings);
                }
            }
        }

        /// <inheritdoc />
        public void ExportCategory(string categoryName, string fileName, IExportSettings settings = null)
        {
            if (settings != null) { this.Settings = settings; } // use custom settings if supplied
            CmcOptionFlags flags = this.Settings.UseThids 
                ? CmcOptionFlags.UseThids 
                : CmcOptionFlags.Default | CmcOptionFlags.IgnoreSyncCondition;

            // User requested we skip connections.
            // A default cursor on a category contains all fields *and* connections.
            // The data receiving routines will ignore them, but they will be read unless we do not include them in our cursor
            // We optimize here by only including direct fields in the cursor
            using (var db = new CommenceDatabase())
            {
                if (this.Settings.SkipConnectedItems && this.Settings.HeaderMode != HeaderMode.CustomLabel)
                {
                    using (ICommenceCursor cur = GetCategoryCursorFieldsOnly(db, categoryName, flags))
                    {
                        // we can limit MAX_FIELD_SIZE in this case
                        this.Settings.MaxFieldSize = (int)Math.Pow(2, 15); // 32.768‬, the built-in Commence max fieldlength (large text) is 30.000
                        ExportCursor(cur, fileName, this.Settings);
                    }
                }
                else
                {
                    using (ICommenceCursor cur = GetCategoryCursorAllFieldsAndConnections(db, categoryName, flags))
                    {
                        // You can create a cursor with all fields including connections by just
                        // supplying CmcOptionFlags.All.
                        // However, when the cursor is read, connected items are returned as comma-delimited strng,
                        // which, because Commence does not supply text-qualifiers, makes it impossible to split them.
                        // We therefore explicitly set the connections which deteriorates performance
                        // but gains us (more) reliability.
                        ExportCursor(cur, fileName, this.Settings);
                    }
                }
            }
        }
        
        /// <inheritdoc />
        public void ExportCategory(string categoryName, IEnumerable<ICursorFilter> filters, string fileName, IExportSettings settings = null)
        {
            if (settings != null) { this.Settings = settings; } // use custom settings if supplied
            CmcOptionFlags flags = this.Settings.UseThids
                ? CmcOptionFlags.UseThids
                : CmcOptionFlags.Default | CmcOptionFlags.IgnoreSyncCondition;

            using (var db = new CommenceDatabase())
            {
                if (this.Settings.SkipConnectedItems && this.Settings.HeaderMode != HeaderMode.CustomLabel)
                {
                    using (ICommenceCursor cur = GetCategoryCursorFieldsOnly(db, categoryName, flags))
                    {
                        ApplyFilters(cur, filters);
                        this.Settings.MaxFieldSize = (int)Math.Pow(2, 15); // 32.768‬, the built-in Commence max fieldlength (large text) is 30.000
                        ExportCursor(cur, fileName, this.Settings);
                    }
                }
                else
                {
                    using (ICommenceCursor cur = GetCategoryCursorAllFieldsAndConnections(db, categoryName, flags))
                    {
                        ApplyFilters(cur, filters);
                        ExportCursor(cur, fileName, this.Settings);
                    }
                }
            }
        }
        #endregion

        #region Helper methods

        private void ApplyFilters(ICommenceCursor cur, IEnumerable<ICursorFilter> filters)
        {
            ((CursorFilters)cur.Filters).AddRange(filters); // AddRange is not exposed by interface
            cur.Filters.Apply();
        }

        private void SubscribeToWriterEvents(BaseWriter w)
        {
            if (w != null)
            {
                w.WriterProgressChanged += (s, e) => HandleExportProgressChanged(s, e);
                w.WriterCompleted += (s, e) => HandleExportCompleted(s, e);
            }
        }

        private void UnsubscribeToWriterEvents(BaseWriter w)
        {
            if (w != null)
            {
                w.WriterProgressChanged -= (s, e) => HandleExportProgressChanged(s, e);
                w.WriterCompleted -= (s, e) => HandleExportCompleted(s, e);
            }
        }

        private ICommenceCursor GetCategoryCursorAllFieldsAndConnections(ICommenceDatabase db, string categoryName, CmcOptionFlags flags)
        {
            ICommenceCursor cur = db.GetCursor(categoryName, CmcCursorType.Category, flags);
            string[] fieldNames = db.GetFieldNames(categoryName).ToArray();
            cur.Columns.AddDirectColumns(fieldNames);
            var cons = db.GetConnectionNames(cur.Category);
            foreach (var c in cons)
            {
                //string nameField = db.GetNameField(c.ToCategory);
                //cur.Columns.AddRelatedColumn(c.Name, c.ToCategory, nameField); // this is bad. a related column loses the THID flag
                cur.Columns.AddDirectColumn(c.Name + ' ' + c.ToCategory); // will respect UseThids flag
            }
            cur.Columns.Apply();
            return cur;
        }

        private ICommenceCursor GetCategoryCursorFieldsOnly(ICommenceDatabase db, string categoryName, CmcOptionFlags flags)
        {
            ICommenceCursor cur = db.GetCursor(categoryName, CmcCursorType.Category, flags);
            string[] fieldNames = db.GetFieldNames(categoryName).ToArray();
            cur.Columns.AddDirectColumns(fieldNames);
            cur.Columns.Apply();
            return cur;
        }

        private string GetActiveViewName(ICommenceDatabase db)
        {
            string retval = string.Empty;
            IActiveViewInfo av = db.GetActiveViewInfo();
            if (av != null && string.IsNullOrEmpty(av.Field)) // view is active and it is not an item detail form
            {
                retval = av.Name;
            }
            else
            {
                throw new CommenceCOMException("Could not determine what view is active in Commence.");
            }
            return retval;
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
            if (!settings.PreserveAllConnections && settings.NestConnectedItems && settings.ExportFormat == ExportFormat.Json)
            {
                return new AdoNetWriter(cursor, settings); // i think this can be taken out entirely
            }
            if (settings.PreserveAllConnections)
            {
                return new Complex.SQLiteWriter(cursor, settings);
            }

            switch (settings.ExportFormat)
            {
                case ExportFormat.Text:
                    return new TextWriter(cursor, settings);
                case ExportFormat.Html:
                    return new HtmlWriter(cursor, settings);
                case ExportFormat.Xml:
                    return new XmlWriter(cursor, settings);
                case ExportFormat.Json:
                    return new JsonWriter(cursor, settings);
                case ExportFormat.Excel:
                    //_writer = new ExcelWriterUsingXml(cursor, settings);
                    //return new ExcelWriterUsingOleDb(cursor, settings);
                    //return new ExcelWriterUsingOpenXml(cursor, settings);
                    return new ExcelWriterUsingEPPlus(cursor, settings);
                case ExportFormat.Event:
                    return new EventWriter(cursor, settings);
                case ExportFormat.GoogleSheets:
                    // will probably always be too slow
                    throw new NotImplementedException("Exportformat not yet implemented.");
                default:
                    return new XmlWriter(cursor, settings);
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
        public virtual void HandleExportProgressChanged(object sender, ExportProgressChangedArgs e)
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
        protected virtual void OnExportProgressChanged(ExportProgressChangedArgs e)
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
            // nothing to do, just in here to not break interface
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
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using Vovin.CmcLibNet.Database;
using Vovin.CmcLibNet.Extensions;

namespace Vovin.CmcLibNet.Export
{
    /// <summary>
    /// Base class for all export classes. It's main purpose is to fire up the data reader,
    /// subscribe to its events and call the appropriate data reading method (i.e., API or DDE).
    /// Because Commence data is passed using events, IDisposable is implemented;
    /// this allows us to close and clean up any objects that require explicit closing (like streamwriters) in derived classes.
    /// </summary>
    internal abstract class BaseWriter : IDisposable
    {
        internal event ExportProgressAsStringChangedHandler ExportProgressChanged; // we want to bubble up this event
        internal event ExportCompletedHandler ExportCompleted; // we want to bubble up this event

        #region Fields
        ///// <summary>
        ///// File to export to.
        ///// </summary>
        // protected internal readonly string _fileName = null;
        /// <summary>
        /// Custom headers.
        /// </summary>
        protected internal string[] _customColumnHeaders = null; // change to object for use with COM?
        /// <summary>
        /// Cursor object to retrieve data from.
        /// </summary>
        protected internal ICommenceCursor _cursor = null;
        /// <summary>
        /// Settings object.
        /// </summary>
        protected internal IExportSettings _settings = null;
        /// <summary>
        /// Name of the datasource, will be either category- or view name.
        /// Used as root node in XML and JSON.
        /// </summary>
        protected internal readonly string _dataSourceName = null;
        /// <summary>
        /// Holds information on the columns in the cursor.
        /// </summary>
        protected internal List<ColumnDefinition> _columnInfo = null;
        /// <summary>
        /// Convenience field to expose what 'column' headers to use.
        /// </summary>
        protected internal List<string> _exportHeaders = null;
        /// <summary>
        /// Disposed flag for use with IDisposable.
        /// </summary>
        private bool disposed = false;
        /// <summary>
        /// Datareader object.
        /// </summary>
        private DataReader dr = null;
        #endregion

        #region Constructors
        /// <summary>
        /// Constructor used when exporting cursor.
        /// </summary>
        /// <param name="cursor">Database.ICommenceCursor.</param>
        /// <param name="settings">ExportSettings object.</param>
        /// <exception cref="ArgumentOutOfRangeException"></exception>
        /// <exception cref="ArgumentException"></exception>
        protected internal BaseWriter(ICommenceCursor cursor, 
            IExportSettings settings)
        {
            _cursor = cursor;
            _settings = settings;
            // this check is very important
            if (_settings.HeaderMode == HeaderMode.CustomLabel)
            {
                if (ValidCustomHeaders(_settings.CustomHeaders.Select(x => x.ToString()).ToArray()))
                {
                    _customColumnHeaders = _settings.CustomHeaders.Select(x => x.ToString()).ToArray();
                }
            }
            _dataSourceName = (string.IsNullOrEmpty(_cursor.View)) ? _cursor.Category : _cursor.View;
            _cursor.SeekRow(CmcCursorBookmark.Beginning, 0); // put rowpointer on first item
        }

        ~BaseWriter()
        {
            Dispose(false);
        }
        #endregion

        #region Methods
        /// <summary>
        /// Write the export file.
        /// </summary>
        /// <param name="fileName">File name.</param>
        /// <exception cref="System.IO.IOException">File in use.</exception>
        /// <exception cref="System.ArgumentNullException">Argument is null.</exception>
        protected internal abstract void WriteOut(string fileName);

        /// <summary>
        /// Method that deals with the data as it is being read.
        /// The minimum amount of expected data is a single list of CommenceValue objects representing a single item (row) in Commence,
        /// but it can also be multiple lists representing Commence items.
        /// It must NOT be a partial Commence item!
        /// </summary>
        /// <param name="sender">sender.</param>
        /// <param name="e">ExportProgressChangedArgs.</param>
        protected internal abstract void HandleProcessedDataRows(object sender, CursorDataReadProgressChangedArgs e);

        /// <summary>
        /// Method that deals with any finalization of the export,
        /// such as writing closing elements and closing streams.
        /// </summary>
        /// <param name="sender">sender.</param>
        /// <param name="e">ExportCompleteArgs.</param>
        protected internal abstract void HandleDataReadComplete(object sender, ExportCompleteArgs e);
        #endregion

        #region Data fetching methods

        /// <summary>
        /// Main data reading method.
        /// </summary>
        protected internal void ReadCommenceData()
        {
            dr = new DataReader(_cursor, _settings, this.ColumnDefinitions, _customColumnHeaders);
            // subscribe to the events the datareader throws
            //dr.DataProgressChanged += this.HandleProcessedDataRows;
            //dr.DataReadCompleted += this.HandleDataReadComplete;
            // subscribe in a way we can use Invoke (asynchronous)(?).
            // Should be okay for synchronous stuff as well
            dr.DataProgressChanged += (s, e) => HandleProcessedDataRows(s, e);
            dr.DataReadCompleted += (s, e) => HandleDataReadComplete(s, e);

            if (this._settings.UseDDE)
            {
                dr.GetDataByDDE(this.GetTableInfoFromCursorColumns());
            }
            else
            {
                if (_settings.ReadCommenceDataAsync)
                {
                    dr.GetDataByAPIAsync(); // not awaited!
                }
                else
                {
                    dr.GetDataByAPI();
                }
            }
        }

        protected internal bool IsFileLocked(FileInfo file)
        {
            // file does not exist so it cannot be locked
            if (!File.Exists(file.FullName)) { return false; }

            FileStream stream = null;

            try
            {
                stream = file.Open(FileMode.Open, FileAccess.Read, FileShare.None);
            }
            catch (IOException)
            {
                //the file is unavailable because it is:
                //still being written to
                //or being processed by another thread
                //or does not exist (has already been processed)
                return true;
            }
            finally
            {
                if (stream != null)
                    stream.Close();
            }

            //file is not locked
            return false;
        }
        #endregion

        #region Helper methods

        /// <summary>
        /// Checks to see if the number and values of the supplied custom headers are valid.
        /// Does not check characters.
        /// </summary>
        /// <exception cref="ArgumentOutOfRangeException">Invalid number of custom headers supplied.</exception>
        /// <exception cref="ArgumentException">Duplicate custom headers supplied.</exception>
        /// <param name="cHeaders">Custom headers list.</param>
        /// <returns><c>true</c> if valid.</returns>
        /// <remarks>Note that a header must be passed for every column even when SkipConnections setting was set to true.</remarks>
        private bool ValidCustomHeaders(string[] cHeaders)
        {
            // did we get the correct number?
            if (cHeaders.Length != _cursor.ColumnCount)
            {
                throw new ArgumentOutOfRangeException("Invalid number of custom headers. Should be " 
                    + _cursor.ColumnCount.ToString() + ", but received " + cHeaders.Length.ToString() + ".");
            }
            // all are headers unique?
            string[] unique = cHeaders.Distinct().ToArray();
            if (unique.Length != cHeaders.Length)
            {
                throw new ArgumentException("Custom headers cannot contain duplicates.");
            }
            return true;
        }

        /// <summary>
        /// Creates a 'mock-list' of table-information from the columns of a Commence cursor.
        /// This list can be used to construct real tables in a DataSet.
        /// </summary>
        /// <returns>List of table descriptions to be created in DataSet.</returns>
        protected internal List<TableDef> GetTableInfoFromCursorColumns()
        {
            List<ColumnDefinition> columns = this.ColumnDefinitions; // do we want this dependency? YES
            // get the list of all connections
            List<Tuple<string,string>> qualifiedConnections = null;
            List<string> fields = new List<string>(); // there is always at least 1 field
            List<TableDef> TableDefList = new List<TableDef>(); // there is always at least 1 table

            TableDef td = new TableDef(_cursor.Category, _dataSourceName, true);
            TableDefList.Add(td);

            // see if we have any connections for which we have to create additional tables
            foreach (ColumnDefinition cd in columns)
            {
                if (cd.IsConnection) // store connection name
                {
                    if (qualifiedConnections == null) { qualifiedConnections = new List<Tuple<string, string>>(); }
                    qualifiedConnections.Add(Tuple.Create(cd.QualifiedConnection, cd.Category));
                }
                else // add direct field
                {
                    td.ColumnDefinitions.Add(cd);
                }
            }

            if (qualifiedConnections != null)
            {
                // connections may contain multiple members with the same connection;
                // for instance when a view displays different fields from the same connection.
                // Those fields should go in the same table.
                // Get rid of duplicate connection names
                qualifiedConnections = qualifiedConnections.Distinct().ToList();

                foreach (Tuple<string, string> qc in qualifiedConnections)
                {
                    TableDef t = new TableDef(qc.Item1, qc.Item2); // s will always be unique, BUT only if case-sensitivity is taken into account. The DataTable.Name property is case-sensitive.
                    // get a list of ColumnDefinition objects for this particular connection, so we can extract fieldnames.
                    IEnumerable<ColumnDefinition> l = columns.Where(o => o.QualifiedConnection == qc.Item1);
                    foreach (ColumnDefinition cd in l)
                    {
                        // fieldnames should be unique,
                        // but there is a tricky way to get duplicates in a cursor:
                        // if user both added related name field explicitly AND defined the connection as direct field(!)
                        // So we add a check to make sure no duplicate fieldnames are defined.
                        if (t.ColumnDefinitions.Find(o => o.FieldName == cd.FieldName) == null) // if no matches were found, continue.
                        {
                            t.ColumnDefinitions.Add(cd); // TODO could we pair up wrong field with wrong columndefinition here? For we query a collection, a duplicate may be found, then we store the first?
                        }
                    }
                    TableDefList.Add(t);
                }
            }
            return TableDefList;
        }

        // only used with ADO writers
        // TODO this method can be greatly simplified
        protected internal DataSet CreateDataSetFromCursorColumns()
        { 
            DataSet retval = null;

            List<TableDef> tables = GetTableInfoFromCursorColumns();
            // we have now collected all table info,
            // we can create a dataset
            DataTable dt = null;
            DataTable primaryTable = null; //hack

            for (int i = 0; i< tables.Count; i++)
            {
                if (retval == null) { retval = new DataSet(this._dataSourceName); }
                
                TableDef td = tables[i];

                try
                {
                    dt = retval.Tables.Add(td.Name);
                }
                catch (DuplicateNameException)
                {
                    // TODO deal with this
                    // It is a rare exception, but possible. Commence connection names are case-sensitive.
                    // In that case this exception is not thrown, but data can still end up in the wrong datatable (or more likely fail).
                    // we should probably include some pre-check and throw something more meaningful
                    throw; // for now just rethrow to be safe
                }

                // define columns
                DataColumn dc = null;
                
                // we need to set some general fields
                if (td.Primary) // primary table columns
                {
                    dc = dt.Columns.Add("id", typeof(Int32)); // not auto-incremented!
                    dc.AllowDBNull = false;
                    // hack for setting relations later on.
                    primaryTable = dt; // capture that this is our primary table so we can set relationships later on
                    dt.ExtendedProperties.Add("IsPrimaryTable", true);
                }
                else // related table columns
                {
                    // create a primary key that autoincrements
                    dc = dt.Columns.Add("id", typeof(Int32));
                    dc.AllowDBNull = false;
                    dc.AutoIncrement = true;
                    // create a foreign key field to reference the connected item
                    dc = dt.Columns.Add("fkid", typeof(Int32));
                    dc.AllowDBNull = false; // cannot be null
                }

                // now process Commence fields
                // rest of related columns
                for (int j = 0; j < td.ColumnDefinitions.Count; j++)
                {
                    try
                    {
                        dc = dt.Columns.Add(td.ColumnDefinitions[j].FieldName);
                    }
                    catch (DuplicateNameException)
                    {
                        // appending some number would decouple the fieldname in the dataset
                        // from the fieldname in the Commence columnset.
                        // for now just rethrow
                        throw;
                    }
                    // set column properties
                    dc.DataType = td.ColumnDefinitions[j].CommenceFieldDefinition.Type.GetTypeForCommenceField();
                    dc.AllowDBNull = true; // this is default, but setting it explicitly makes it more clear.
                }
            }
            // we would like to set relationships as well.
            // as it is, we don't know what our 'primary' table is at this point in the code
            // all we know it is the only table without a fkid field, but that's a little awkward to check.
            // we hack that by having a 'mock' primary table variable.
            List<DataRelation> relations = new List<DataRelation>();
            foreach (DataTable d in retval.Tables)
            {
                if (!d.Equals(primaryTable))
                {
                    DataRelation r = new DataRelation("rel_" + d.TableName,
                        primaryTable.TableName,
                        d.TableName,
                        new string[] { "id" },
                        new string[] { "fkid" },
                        true); // setting nested to true allows for nested XML exports
                    relations.Add(r);
                }
            }
            // add the relations to the dataset
            // tricky snag found on https://msdn.microsoft.com/en-us/library/2z22c2sz.aspx
            // "Any DataRelation object created by using this constructor must be added to the collection
            // with the AddRange method inside of a BeginInit and EndInit block.
            if (relations.Any())
            {
                retval.BeginInit();
                retval.Relations.AddRange(relations.ToArray());
                retval.EndInit();
            }
            return retval;
        }

        /// <summary>
        /// Dispose method
        /// </summary>
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        /// <summary>
        /// Protected implementation of Dispose pattern.
        /// </summary>
        /// <param name="disposing">disposing.</param>
        protected virtual void Dispose(bool disposing)
        {
            if (disposed)
                return;

            if (disposing)
            {
                // Free any other managed objects here.
                //
                if (dr != null)
                {
                    dr.DataProgressChanged -= this.HandleProcessedDataRows;
                    dr.DataReadCompleted -= this.HandleDataReadComplete;
                }

                if (_cursor!= null)
                {
                    _cursor.Close();
                }
                //is this overkill?
                this.ExportProgressChanged = null;
            }

            // Free any unmanaged objects here.
            //
            disposed = true;
        }
        #endregion

        #region Event methods
        /// <summary>
        /// Handler used to bubble up the ExportProgressChanged event
        /// </summary>
        /// <param name="e">ExportProgressAsJsonChangedArgs.</param>
        protected virtual void OnExportProgressChanged(ExportProgressChangedArgs e)
        {
            ExportProgressAsStringChangedHandler handler = ExportProgressChanged;
            if (handler == null) { return; } // no subscriptions
            Delegate[] eventHandlers = handler.GetInvocationList();
            foreach (Delegate currentHandler in eventHandlers)
            {
                ExportProgressAsStringChangedHandler currentSubscriber = (ExportProgressAsStringChangedHandler)currentHandler;
                try
                {
                    currentSubscriber(this, e);
                }
                catch { }
            }
        }

        /// <summary>
        /// Used to bubble up the Export completed event
        /// </summary>
        /// <param name="e">ExportCompleteArgs</param>
        protected virtual void OnExportCompleted(ExportCompleteArgs e)
        {
            ExportCompletedHandler handler = ExportCompleted;
            if (handler == null) { return; } // no subscriptions
            Delegate[] eventHandlers = handler.GetInvocationList();
            foreach (Delegate currentHandler in eventHandlers)
            {
                ExportCompletedHandler currentSubscriber = (ExportCompletedHandler)currentHandler;
                try
                {
                    currentSubscriber(this, e);
                }
                catch { }
            }
        }

        /// <summary>
        /// Derived classes can use this method to bubble up the ExportProgressChanged event
        /// </summary>
        /// <param name="e">ExportProgressChangedArgs</param>
        protected void BubbleUpProgressEvent(CursorDataReadProgressChangedArgs e)
        {
            OnExportProgressChanged(new ExportProgressChangedArgs(e.RowsProcessed, e.RowsTotal));
        }

        protected void BubbleUpCompletedEvent(ExportCompleteArgs e)
        {
            OnExportCompleted(e);
        }
        #endregion

        #region Properties

        /// <summary>
        /// Creates and returns an instance of the HeaderLists class that contains information on the columns and fields of the cursor to be exported.
        /// </summary>
        protected internal List<ColumnDefinition> ColumnDefinitions
        {
            // TODO: this should probably better be a property of CommenceCursor
            // it means substantial rewriting

            // create _headerLists just once!
            get
            {
                if (_columnInfo == null)
                {
                    ColumnParser cp = null;
                    if (this._customColumnHeaders != null)
                    {
                        cp = new ColumnParser(_cursor, this._customColumnHeaders);
                    }
                    else
                    {
                        cp = new ColumnParser(_cursor);
                    }
                    _columnInfo = cp.ParseColumns();
                }
                return _columnInfo;
            }
            set // added after introduction of SqliteWriter, which recreates cursors on the fly
            {
                _columnInfo = value;
            }
        }

        /// <summary>
        /// Specifies the columnheaders or 'nodenames' to use when exporting data.
        /// </summary>
        protected internal List<string> ExportHeaders
        {
            get
            {
                if (_exportHeaders == null)
                {
                    _exportHeaders = new List<string>();
                    foreach (ColumnDefinition cd in this.ColumnDefinitions)
                    {
                        // selects what headers to use
                        switch (_settings.HeaderMode)
                        {
                            case HeaderMode.Columnlabel:
                                //string label = Utils.AddUniqueIdentifier(cd.ColumnLabel, _exportHeaders, 1, 244, 20); // make sure we have unique columnlabels, or XML and JSON may complain.
                                _exportHeaders.Add(cd.ColumnLabel);
                                break;
                            case HeaderMode.Fieldname:
                                _exportHeaders.Add(cd.ColumnName);
                                break;
                            case HeaderMode.CustomLabel:
                                _exportHeaders.Add(cd.CustomColumnLabel);
                                break;
                        } // switch
                    } // foreach
                } // if
                // make sure all exportheaders are unique
                // if columns are renamed that will obvioulsy break the ability to export back into Commence.
                // there is little we can do about that
                // keep in mind that fieldnames are always unique within a cursor, they can just look ugly :)
                _exportHeaders = Utils.RenameDuplicates(_exportHeaders).ToList();
                return _exportHeaders;
            }
        }
        #endregion
    }
}

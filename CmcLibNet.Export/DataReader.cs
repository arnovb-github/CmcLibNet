using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Data;
using Vovin.CmcLibNet.Database;

namespace Vovin.CmcLibNet.Export
{
    #region Enumerations
    /// <summary>
    /// Enum for data output formats.
    /// </summary>
    internal enum ValueFormatting
    {
        /// <summary>
        /// No formatting, use data as-is. This means formatted to whatever formatting Commence inherits from the system.
        /// </summary>
        None = 0,
        /// <summary>
        /// Return canonical format as defined by Commence. Unlike Commence, CmcLibNet returns connected data in the format as well. <seealso cref="CmcOptionFlags.Canonical"/>
        /// </summary>
        Canonical = 1,
        /// <summary>
        /// Return data compliant with ISO 8601 format. http://www.iso.org/iso/iso8601
        /// </summary>
        XSD_ISO8601 = 2,
    }
    #endregion
    // Performs the reading of the Commence database and returns the results as event arguments
    internal class DataReader
    {
        internal delegate void DataProgressChangedHandler(object sender, DataProgressChangedArgs e);
        internal delegate void DataReadCompleteHandler(object sender, DataReadCompleteArgs e);
        //internal delegate void ExportProgressChangedHandler(object sender, ExportProgressChangedArgs e); // already defined in ExportEngine
        internal event DataProgressChangedHandler DataProgressChanged;
        internal event DataReadCompleteHandler DataReadCompleted;
        internal event ExportProgressChangedHandler ExportProgressChanged;
        
        private readonly CommenceCursor _cursor = null;
        private readonly IExportSettings _settings = null;
        private readonly int _maxrows = 1000; // maximum number of rows to read per iteration
        private readonly List<ColumnDefinition> _columndefinitions = null;
        private readonly int _totalrows = 0;
        private readonly string _pattern = "(?<!\r)\n";
        private readonly Regex _regex = null;
        private readonly bool _useThids;

        #region Constructors
        internal DataReader(ICommenceCursor cursor, IExportSettings settings, List<ColumnDefinition> columndefinitions, string[] customColumnHeaders)
        {
            this._cursor = (CommenceCursor)cursor;
            _totalrows = cursor.RowCount;
            this._settings = settings;
            _maxrows = this._settings.NumRows;
            this._columndefinitions = columndefinitions;
            _regex = new Regex(_pattern);
            if (this._settings.XSDCompliant)
            {
                this.Formatting = ValueFormatting.XSD_ISO8601;
            }
            else if (this._settings.Canonical)
            {
                this.Formatting = ValueFormatting.Canonical;
            }
            else
            {
                this.Formatting = ValueFormatting.None; // default
            }
            if ((_cursor).Flags.HasFlag(CmcOptionFlags.UseThids))
            {
                _useThids = true;
            }
        }
        #endregion

        #region Event raising methods

        protected virtual void OnDataProgressChanged(DataProgressChangedArgs e)
        {
            try
            {
                DataProgressChangedHandler handler = DataProgressChanged;
                if (handler != null)
                    handler(this, e);
            }
            catch { }
        }

        protected virtual void OnDataReadCompleted(DataReadCompleteArgs e)
        {
            try
            {
                DataReadCompleteHandler handler = DataReadCompleted;
                if (handler != null)
                    handler(this, e);
            }
            catch { }
        }

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
            catch { } //rethrow
        }

        #endregion

        #region Data fetching methods

        // collect Commence rowvalues as jagged array,
        // then raises an event with that array
        internal void GetDataByAPI()
        {
            int counter = 0;
            for (int totalrows = 0; totalrows < this._cursor.RowCount; totalrows += _maxrows)
            {
                string[][] rawdata = _cursor.GetRawData(_maxrows); // first dimension is rows, second dimension is columns
                List<List<CommenceValue>> retval = new List<List<CommenceValue>>();
                CommenceValue cv = null;
                ColumnDefinition cd = null;

                // rawdata represents the actual database row values
                for (int i = 0; i < rawdata.GetLength(0); i++) // rows
                {
                    List<CommenceValue> rowdata = new List<CommenceValue>();
                    // for thids we can assume the first row of rawdata contains the thid
                    if (this._useThids)
                    {
                        cv = new CommenceValue(rawdata[i][0], this._columndefinitions.First()); // assumes thid column is first. This is an accident waiting to happen.
                        rowdata.Add(cv);
                    }

                    // process row
                    for (int j = 1; j < rawdata[i].Length; j++) // columns
                    {
                        // a column for the thid is only returned when a thid is requested
                        // therefore getting the right column is a little tricky
                        int colindex;
                        if (this._useThids)
                        {
                            colindex = j;
                        }
                        else
                        {
                            colindex = j - 1;
                        }
                        cd = this._columndefinitions[colindex];

                        string[] buffer = null;
                        if (cd.IsConnection)
                        {
                            if (String.IsNullOrEmpty(rawdata[i][j].Trim()))
                            {
                                cv = new CommenceValue(cd); // always create a CommenceValue for consistency
                            }
                            else
                            {
                                if (!_settings.SplitConnectedItems)
                                {
                                    buffer = new string[] { rawdata[i][j] };

                                } // if
                                else
                                {
                                    switch (cd.FieldType)
                                    {
                                        case Database.CommenceFieldType.Text:
                                            // we use a regex to split values at "\n" *but not* "\r\n"
                                            // this is not 100% fail-safe as a fieldvalue *can* contain just \n if it is a large text field.
                                            // in that case, your only option is to suppress the splitting in ExportSettings
                                            buffer = _regex.Split(rawdata[i][j]); // this may result in Commence values being split if they contain embedded delimiters
                                            break;
                                        default:
                                            buffer = rawdata[i][j].Split(new string[] { cd.Delimiter }, StringSplitOptions.None);
                                            break;
                                    } // switch

                                    // buffer now contains the connected values as array, do any formatting transformation
                                    buffer = FormatValues(buffer, this.Formatting, cd);
                                }
                                cv = new CommenceValue(buffer, cd);
                            } // if !String.IsNullOrEmpty
                        } // if IsConnection
                        else // single value
                        {
                            buffer = new string[] { rawdata[i][j] };
                            buffer = FormatValues(buffer, this.Formatting, cd);
                            cv = new CommenceValue(buffer[0], cd);
                        } // else IsConnection
                        if (cv != null) { rowdata.Add(cv); }
                    } // for j
                    counter++;
                    ExportProgressChangedArgs rowread_args = new ExportProgressChangedArgs(counter, _totalrows);
                    OnExportProgressChanged(rowread_args); // raise event for each row read
                    retval.Add(rowdata);
                } // for i
                // per batch of rows
                DataProgressChangedArgs args = new DataProgressChangedArgs(retval, counter);
                OnDataProgressChanged(args); // raise event after each batch of rows
            } // totalrows
            DataReadCompleteArgs e = new DataReadCompleteArgs(counter);
            OnDataReadCompleted(e); // done with reading data
        }

        internal void GetDataByDDE(List<TableDef> mocktables)
        {
            /* DDE requests are limited to a maximum length of 255 characters, 
             * which is easily exceeded. A workaround is splitting the requests.
             * Not pretty but the only way to get to many-many relationships that contain >93750 worth of connected characters
             * without setting the maxfieldsize higher.
             */

            List<List<CommenceValue>> rows = null;
            List<CommenceValue> rowvalues = null;
            ICommenceDatabase db = new CommenceDatabase();

            // determine if we are dealing with a view or category
            if (String.IsNullOrEmpty(_cursor.View))
            {
                db.ViewCategory(this._cursor.Category);
            }
            else
            {
                db.ViewView(this._cursor.View);
            }
            int itemCount = db.ViewItemCount();

            for (int i = 1; i <= itemCount; i++) // note that we use a 1-based iterator
            {
                rows = new List<List<CommenceValue>>();
                rowvalues = new List<CommenceValue>();
                foreach (TableDef td in mocktables)
                {
                    string[] DDEResult = null;
                    List<string> fieldNames = td.ColumnDefinitions.Select(o => o.FieldName).ToList<string>();
                    if (td.Primary)
                    {
                        // ViewFields and ViewConnectedFields have a limited capacity
                        // the total length of a DDE command cannot exceed 255 characters
                        // What we are going to do is limit the number of characters to a value of up to 150 chars,
                        // to be on the safe side (ViewConnectedFilter and two delimiters already make up 35 characters!)
                        ListChopper lcu = new ListChopper(fieldNames, 150);
                        foreach (List<string> l in lcu.Portions)
                        {
                            DDEResult = db.ViewFields(i, l);

                            // we have our results, we now have to create CommenceValue objects from it
                            // and we also have to match them up with their respective column
                            // this is a little tricky
                            for (int j = 0; j < DDEResult.Length; j++)
                            {
                                ColumnDefinition cd = td.ColumnDefinitions.Find(o => o.FieldName.Equals(l[j]));
                                string[] buffer = new string[] { DDEResult[j] };
                                buffer = FormatValues(buffer,this.Formatting, cd);
                                CommenceValue v = new CommenceValue(buffer[0], cd);
                                rowvalues.Add(v);
                            } // for
                        } // list l
                    }
                    else // we are dealing with a connection
                    {
                        int conItemCount = db.ViewConnectedCount(i, td.ColumnDefinitions[0].Connection, td.ColumnDefinitions[0].Category); // doesn't matter which one we use
                        // here's a nice challenge:
                        // upon every iteration we get a row of fieldvalues from the connection
                        // to make things worse, we chop them up so it aren't even complete rows.
                        // we must aggregate the values for each field.
                        // We'll construct a datatable to hack around that;
                        // we could have also used a dictionary I suppose.

                        //  using a datatable may be easiest
                        DataTable dt = new DataTable();
                        for (int c = 0; c < fieldNames.Count; c++)
                        {
                            dt.Columns.Add(fieldNames[c]); // add fields as columns, keeping everything default
                        }
                        
                        // loop all connected items
                        for (int citemcount = 1; citemcount <= conItemCount; citemcount++)
                        {
                            DataRow dr = dt.NewRow(); // create a row containing all columns
                            ListChopper lcu = new ListChopper(fieldNames, 150);
                            foreach (List<string> list in lcu.Portions)
                            {
                                DDEResult = db.ViewConnectedFields(i, td.ColumnDefinitions[0].Connection, td.ColumnDefinitions[0].Category, citemcount, list);
                                // populate colums for the fields we requested
                                for (int j = 0; j < DDEResult.Length; j++)
                                {
                                    dr[list[j]] = DDEResult[j];
                                }

                            } // list l
                            dt.Rows.Add(dr);
                            
                        } // citemcount

                        // create a CommenceValue from every column in the datatable
                        foreach (DataColumn dc in dt.Columns)
                        {
                            // this will also return columns that have no data, which is what we want.
                            string[] query =
                                (from r in dt.AsEnumerable()
                                 select r.Field<String>(dc.ColumnName)).ToArray();
                            ColumnDefinition cd = td.ColumnDefinitions.Find(o => o.FieldName.Equals(dc.ColumnName));
                            CommenceValue cv = null;
                            if (query.Length > 0) // only create value if there is one
                            {
                                query = FormatValues(query, this.Formatting, cd);
                                cv = new CommenceValue(query, cd);
                                
                            }
                            else
                            {
                                // create empty CommenceValue
                                cv = new CommenceValue(cd);
                            }
                            rowvalues.Add(cv);
                        }
                    } // if
                    
                } // foreach tabledef
                rows.Add(rowvalues);
                DataProgressChangedArgs args = new DataProgressChangedArgs(rows, i); // progress within the cursor
                OnDataProgressChanged(args);
                ExportProgressChangedArgs rowread_args = new ExportProgressChangedArgs(i, _totalrows); // total progress
                OnExportProgressChanged(rowread_args);
            } // i
            db = null;
            DataReadCompleteArgs a = new DataReadCompleteArgs(itemCount);
            OnDataReadCompleted(a);
        }

        #endregion

        #region Supporting methods

        private static string[] FormatValues(string[] values, ValueFormatting format, ColumnDefinition cd)
        {
            string[] retval = values;
            switch (format)
            {
                case ValueFormatting.Canonical:
                    for (int i = 0; i < retval.Length; i++)
                    {
                        if (!String.IsNullOrEmpty(retval[i]))
                        {
                            //retval[i] = GetCanonicalCommenceValue(retval[i], cd.FieldType);
                            retval[i] = CommenceValueConverter.ToCanonical(retval[i], cd.FieldType);
                        }
                    }
                    break;
                case ValueFormatting.XSD_ISO8601:
                    for (int i = 0; i < retval.Length; i++)
                    {
                        if (!String.IsNullOrEmpty(retval[i]))
                        {
                            //string canonical = GetCanonicalCommenceValue(retval[i], cd.FieldType);
                            string canonical = CommenceValueConverter.ToCanonical(retval[i], cd.FieldType);
                            retval[i] = CommenceValueConverter.toIso8601(canonical, cd.FieldType);
                        }
                    }
                    break;
            } // switch
            return retval;
        }
        #endregion

        #region Properties

        private ValueFormatting Formatting { get; set; }

        #endregion

    }

    #region Helper classes
    internal class DataProgressChangedArgs : EventArgs
    {
        internal DataProgressChangedArgs(List<List<CommenceValue>> list, int row)
        {
            this.Values = list;
            this.Row = row;
        }
        internal List<List<CommenceValue>> Values { get; private set; }
        internal int Row { get; private set; }
    }
    internal class DataReadCompleteArgs : EventArgs
    {
        internal DataReadCompleteArgs(int row)
        {
            this.Row = row;
        }
        internal int Row { get; private set; }
    }
    internal class DataRowReadEventArgs : EventArgs
    {
        internal DataRowReadEventArgs(int row)
        {
            this.Row = row;
        }
        internal int Row { get; private set; }
    }
    #endregion
}

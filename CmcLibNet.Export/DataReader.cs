﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Data;
using Vovin.CmcLibNet;
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
                if (DataReadCompleted != null)
                    DataReadCompleted(this, e);
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

        internal void GetDataByAPI2D() // no longer used
        {
            int counter = 0;
            for (int totalrows = 0; totalrows < this._cursor.RowCount; totalrows += _maxrows)
            {
                string[,] rawdata = this.GetRawData2D(_maxrows); // first dimension is rows, second dimension is columns
                List<List<CommenceValue>> retval = new List<List<CommenceValue>>();

                // for thids we can assume the first row of rawdata contais the thid

                // rest of rawdata represents the actual database row values
                for (int i = 0; i < rawdata.GetLength(0); i++) // rows
                {
                    List<CommenceValue> rowdata = new List<CommenceValue>();
                    CommenceValue cv = null;
                    // process row
                    for (int j = 0; j < rawdata.GetLength(1); j++) // columns
                    {
                        ColumnDefinition cd = this._columndefinitions[j];
                        string[] buffer = null;

                        if (cd.IsConnection)
                        {
                            if (String.IsNullOrEmpty(rawdata[i, j].Trim()))
                            {
                                cv = new CommenceValue(cd); // always create a CommenceValue for consistency
                            }
                            else
                            {
                                switch (cd.FieldType)
                                {
                                    case Database.CommenceFieldType.Text:
                                        // we use a regex to split values at "\n" *but not* "\r\n"
                                        // this is not 100% fail-safe as a fieldvalue *can* contain just \n
                                        buffer = _regex.Split(rawdata[i, j]); // this may result in Commence values being split if they contain embedded delimiters
                                        break;
                                    default:
                                        //buffer = rawdata[i, j].Split(Core.Utils.ConnectedItemValueDelimiter, StringSplitOptions.None);
                                        buffer = rawdata[i, j].Split(new string[] { cd.Delimiter }, StringSplitOptions.None);
                                        break;
                                } // switch
                                // buffer now contains the connected values as array, do any formatting transformation
                                buffer = FormatValues(buffer, this.Formatting, cd);
                                cv = new CommenceValue(buffer, cd);
                            } // if !String.IsNullOrEmpty
                        } // if IsConnection
                        else // single value
                        {
                            buffer = new string[] { rawdata[i, j] };
                            buffer = FormatValues(buffer, this.Formatting, cd);
                            cv = new CommenceValue(buffer[0], cd);
                        } // else IsConnection
                        if (cv != null) { rowdata.Add(cv);}
                    } // for j
                    counter++;
                    ExportProgressChangedArgs rowread_args = new ExportProgressChangedArgs(counter, _totalrows);
                    OnExportProgressChanged(rowread_args);
                    retval.Add(rowdata);
                } // for i
                // per batch of rows
                DataProgressChangedArgs args = new DataProgressChangedArgs(retval, counter);
                OnDataProgressChanged(args); // raise event
            } // totalrows
            //return retval;
            DataReadCompleteArgs e = new DataReadCompleteArgs(counter);
            OnDataReadCompleted(e); // done with reading data
        }

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
            //return retval;
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
                            // this is a little tricky...
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

        private string[,] GetRawData2D(int nRows) // no longer used
        {
            /* Note that for connected items, Commence returns a linefeed-delimited string, OR a comma delimited string(!)
             * If the connected field has no data, an empty string is returned, again linefeed-delimited.
             * It is up to the consumer to deal with this.
             * The Headers property can be used to determine what fieldtype is being returned.
             * 
             * Also note, that by getting a RowSet, Commence automatically advances the cursor's rowpointer for us
             * This can lead to some confusion for those who are used to advance it manually, like in ADO.
             * 
             * TODO: There is no proper implementation for the usethids flag yet.
             * It will return thids for related columns that were set directly, but not for the rows themselves yet.
             * You have to use GetRowID to get to them. I haven't decided whether to include the thid as extra column or simply replace the name field value with the thid.
             * 
             */

            // It is very important to dispose of our qrs object when we are done with it.
            // Because it wraps a COM object, simply setting it to null will *not* cut it, the finalizer will *not* run.
            // If we do not dispose of our object, our code is still valid, but commence.exe memory-use explodes,
            // because Commence will keep a reference to each created rowset in memory
            // until all items in a cursor are processed.
            using (ICommenceQueryRowSet qrs = _cursor.GetQueryRowSet(nRows))
            {
                // number of rows requested may be larger than number of available rows in rowset,
                // so make sure the return value is sized properly
                string[,] rowvalues = new string[qrs.RowCount, _cursor.ColumnCount];
                object[] buffer = null;

                for (int i = 0; i < qrs.RowCount; i++)
                {
                    /* Note that we do not bother with the CmcOptionFlags.Canonical flag,
                     * because for indirect fields it is ignored.
                     * We implement our own formatting transformation.
                     * 
                     * TODO: how do we handle UseThids?
                     */
                    // process rowset
                    buffer = qrs.GetRow(i, CmcOptionFlags.Default); // don't bother with canonical flag, it doesn't work properly anyway.

                    for (int j = 0; j < qrs.ColumnCount; j++)
                    {
                        rowvalues[i, j] = buffer[j].ToString();
                    } // j
                } // i
                return rowvalues;
            } // using; qrs will be disposed now
        }

        // moved to CommenecCursor class
        ///// <summary>
        ///// Returns a jagged array of rows with their values.
        ///// In case a THID was requested, the first element contains the thid, otherwise it is empty.
        ///// </summary>
        ///// <param name="nRows">Number of rows to retrieve.</param>
        ///// <param name="useThids">Request thid.</param>
        ///// <returns>
        ///// 'Jagged' string array (=array containing arrays). 
        ///// <list type="table">
        ///// <listheader>Example</listheader>
        ///// <listheader><term>Row</term><description>Contains</description></listheader>
        ///// <item><term>row0</term><description>[(thid0)][fieldvalue0][fieldvalue1][fieldvalueN]</description></item>
        ///// <item><term>row1</term><description>[(thid1)][fieldvalue0][fieldvalue1][fieldvalueN]</description></item>
        ///// <item><term>rowN</term><description>[(thidN)][fieldvalue0][fieldvalue1][fieldvalueN]</description></item>
        ///// </list>
        ///// </returns>
        //private string[][] GetRawDataJagged(int nRows, bool useThids)
        //{
        //    /* Note that for connected items, Commence returns a linefeed-delimited string, OR a comma delimited string(!)
        //     * If the connected field has no data, an empty string is returned, again linefeed-delimited.
        //     * It is up to the consumer to deal with this.
        //     * The Headers property can be used to determine what fieldtype is being returned.
        //     * 
        //     * Also note, that by getting a RowSet, Commence automatically advances the cursor's rowpointer for us
        //     * This can lead to some confusion for those who are used to advance it manually, like in ADO.
        //     * 
        //     */

        //    // It is very important to dispose of our qrs object when we are done with it.
        //    // Because it wraps a COM object, simply setting it to null will *not* cut it, the finalizer will *not* run.
        //    // If we do not dispose of our object, our code is still valid, but commence.exe memory-use explodes,
        //    // because Commence will keep a reference to each created rowset in memory
        //    // until all items in a cursor are processed.
        //    using (ICommenceQueryRowSet qrs = _cursor.GetQueryRowSet(nRows))
        //    {
        //        // number of rows requested may be larger than number of available rows in rowset,
        //        // so make sure the return value is sized properly
        //        string[][] rowvalues = new string[qrs.RowCount][];
        //        object[] buffer = null;

        //        for (int i = 0; i < qrs.RowCount; i++)
        //        {
        //            /* Note that we do not bother with the CmcOptionFlags.Canonical flag,
        //             * because for indirect fields it is ignored.
        //             * We implement our own formatting transformation.
        //             */
        //            rowvalues[i] = new string[qrs.ColumnCount+1]; // add extra element for thid
        //            if (useThids) // do not make the extra API call unless requested
        //            {
        //                string thid = qrs.GetRowID(i,CmcOptionFlags.Default); // GetRowID does not advance the rowpointer. Note that the flag must be 0.
        //                rowvalues[i][0] = thid; // put thid in first column of row
        //            }

        //            // process rowset
        //            buffer = qrs.GetRow(i, CmcOptionFlags.Default); // don't bother with canonical flag, it doesn't work properly anyway.
                    
        //            for (int j = 0; j < qrs.ColumnCount; j++)
        //            {
        //                rowvalues[i][j+1] = buffer[j].ToString(); // put rowvalue in 2nd and up column of row
        //            } // j
        //        } // i
        //        return rowvalues;
        //    } // using; qrs will be disposed now
        //}
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

        ///// <summary>
        ///// Tries to convert certain datatypes to the canonical format Commence *should* return but doesn't sometimes.
        ///// </summary>
        //private static string GetCanonicalCommenceValue(string commenceValue, CommenceFieldType fieldType)
        //{
        //    // Assumes we get a value from Commence, *not* some random string value.
        //    // The Commence value will be formatted according to system regional settings,
        //    // because that is the format Commence will display and return them.
        //    // We take advantage of that in TryParse to the proper .NET data type,
        //    // then we deconstruct that data type to return a string in canonical format.

        //    string retval = commenceValue;
        //    bool isCurrency = false;
        //    string curSymbol = CultureInfo.CurrentCulture.NumberFormat.CurrencySymbol;

        //    switch (fieldType)
        //    {
        //        case CommenceFieldType.Sequence:
        //        case CommenceFieldType.Calculation:
        //        case CommenceFieldType.Number: // should return xxxx.yy
        //            // if Commence was set to "Display as currency" value will be preceded by the system currency symbol
        //            // This setting cannot be gotten from Commence, we have to resolve at runtime.
        //            string numPart = commenceValue;

        //            if (numPart.StartsWith(curSymbol))
        //            {
        //                numPart = numPart.Remove(0, curSymbol.Length).Trim();
        //                isCurrency = true;
        //            }
        //            // try to cast
        //            Decimal number;
        //            if (Decimal.TryParse(numPart, out number))
        //            {
        //                long intPart;
        //                intPart = (long)Math.Truncate(number); // may fail if number is bigger than long.
        //                // d may be negative
        //                int decPart;
        //                decPart = (int)(((Math.Abs(number) - Math.Abs(intPart)) * 100));
        //                if (isCurrency)
        //                {
        //                    retval = '$' + intPart.ToString() + '.' + decPart.ToString("D2");
        //                }
        //                else
        //                {
        //                    retval = intPart.ToString() + '.' + decPart.ToString("D2");
        //                }
        //            }
        //            break;
        //        case CommenceFieldType.Date: // should return yyyymmdd
        //            DateTime date;
        //            if (DateTime.TryParse(commenceValue, out date))
        //            {
        //                retval = date.Year.ToString("D4") + date.Month.ToString("D2") + date.Day.ToString("D2");
        //            }
        //            break;
        //        case CommenceFieldType.Time: // should return hh:mm
        //            DateTime time;
        //            if (DateTime.TryParse(commenceValue, out time))
        //            {
        //                retval = time.Hour.ToString("D2") + ':' + time.Minute.ToString("D2");
        //            }
        //            break;
        //        case CommenceFieldType.Checkbox: // should return 'true' or 'false' (lowercase)
        //            // direct checkbox fields are returned as 'TRUE' or 'FALSE'
        //            if (commenceValue.Equals("TRUE") || commenceValue.Equals("FALSE"))
        //            {
        //                retval = commenceValue.ToLower(); // all OK
        //            }
        //            else
        //            {
        //                // related checkbox fields return 'Yes' or 'No'
        //                retval = (commenceValue.ToLower() == "yes").ToString().ToLower(); // Only returns true if 'yes'. Potential issue here?
        //            }
        //            break;
        //    }
        //    return retval;
        //}
        
        //private static string GetXSDCompliantValue(string canonicalValue, Vovin.CmcLibNet.Export.ColumnDefinition coldef)
        //{
        //    // assumes Commence data in value are already in canonical format!

        //    if (String.IsNullOrEmpty(canonicalValue)) { return canonicalValue; }

        //    string retval = canonicalValue;
        //    string currencySymbol = Utils.CanonicalCurrencySymbol; // when canonical, Commence always uses '$' regardless of regional settings

        //    switch (coldef.FieldType)
        //    {
        //        // Commence may return a currency symbol if the field is defined to display as currency.
        //        // There is no way of requesting that setting, we have to resolve it at runtime.
        //        // When in canonical mode, it is always a dollar sign.
        //        case CommenceFieldType.Number:
        //        case CommenceFieldType.Calculation:
        //            if (canonicalValue.StartsWith(currencySymbol))
        //            {
        //                retval = canonicalValue.Remove(0, currencySymbol.Length);
        //            }
        //            break;
        //        case CommenceFieldType.Date: // expects "yyyymmdd"
        //            retval = canonicalValue.Substring(0, 4) + "-" + canonicalValue.Substring(4, 2) + "-" + canonicalValue.Substring(6, 2);
        //            break;
        //        case CommenceFieldType.Time: // expects "hh:mm"
        //            string[] s = canonicalValue.Split(':');
        //            retval = s[0] + ":" + s[1] + ":00";
        //            break;
        //        case CommenceFieldType.Checkbox: // expects "TRUE" or "FALSE" (case-insensitive)
        //            retval = canonicalValue.ToLower();
        //            break;
        //    }
        //    return retval;
        //}
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
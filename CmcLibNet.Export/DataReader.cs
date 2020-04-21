using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Data;
using Vovin.CmcLibNet.Database;
using System.Threading.Tasks;
using System.Threading;
using System.Collections.Concurrent;

namespace Vovin.CmcLibNet.Export
{
    // Performs the reading of the Commence database and returns the results as event arguments
    internal class DataReader
    {
        // delegates
        internal delegate void DataProgressChangedHandler(object sender, CursorDataReadProgressChangedArgs e);
        internal delegate void DataReadCompleteHandler(object sender, ExportCompleteArgs e);
        
        // events
        internal event DataProgressChangedHandler DataProgressChanged;
        internal event DataReadCompleteHandler DataReadCompleted;
        
        // fields
        private readonly CommenceCursor cursor;
        private readonly IExportSettings settings;
        private readonly int numRows; // maximum number of rows to read per iteration
        private readonly List<ColumnDefinition> columnDefinitions;
        private readonly int totalRows = 0;
        private readonly string pattern = "(?<!\r)\n";
        private readonly Regex regex;
        private readonly bool useThids;

        #region Constructors
        internal DataReader(ICommenceCursor cursor, IExportSettings settings, List<ColumnDefinition> columndefinitions, string[] customColumnHeaders)
        {
            this.cursor = (CommenceCursor)cursor;
            totalRows = cursor.RowCount;
            this.settings = settings;
            numRows = (int)Math.Pow(2, BalanceNumRowsAndFieldSize(settings));
            if (settings.NumRows < numRows)
            {
                numRows = settings.NumRows;
            }
            this.columnDefinitions = columndefinitions;
            regex = new Regex(pattern);
            if (this.cursor.Flags.HasFlag(CmcOptionFlags.UseThids))
            {
                useThids = true;
            }
        }
        #endregion

        #region Event raising methods
        protected virtual void OnDataProgressChanged(CursorDataReadProgressChangedArgs e)
        {
            DataProgressChanged?.Invoke(this, e);
        }

        protected virtual void OnDataReadCompleted(ExportCompleteArgs e)
        {
            DataReadCompleted?.Invoke(this, e);
        }
        #endregion

        #region Data fetching methods
        // collect Commence rowvalues as jagged array,
        internal void GetDataByAPI()
        {
            int rowsProcessed = 0;
            string[][] rawdata =  null;
            for (int rows = 0; rows < totalRows; rows += numRows)
            {
                try
                {
                    rawdata = cursor.GetRawData(numRows); // first dimension is rows, second dimension is columns
                }
                catch (CommenceCOMException)
                {
                    throw;
                }
                rowsProcessed += numRows;
                var data = ProcessDataBatch(rawdata);
                // raise 'progress' event
                CursorDataReadProgressChangedArgs args = new CursorDataReadProgressChangedArgs(data, rowsProcessed > totalRows ? totalRows : rowsProcessed, totalRows);
                OnDataProgressChanged(args); // raise event after each batch of rows
            }
            // raise 'done' event
            ExportCompleteArgs e = new ExportCompleteArgs(totalRows);
            OnDataReadCompleted(e); // done with reading data
        }

        /// <summary>
        /// Reads data using DDE. This is extremely show and should only ever be used as a last resort.
        /// </summary>
        /// <param name="mocktables"></param>
        internal void GetDataByDDE(List<TableDef> mocktables) // needs fixing
        {
            /* DDE requests are limited to a maximum length of 255 characters, 
             * which is easily exceeded. A workaround is splitting the requests.
             * Not pretty but the only way to get to many-many relationships that contain >93750 worth of connected characters
             * without setting the maxfieldsize higher.
             */

            List<List<CommenceValue>> rows;
            List<CommenceValue> rowvalues;
            ICommenceDatabase db = new CommenceDatabase();

            // always define a category
            db.ViewCategory(this.cursor.Category);
            // are we dealing with a view?
            if (!string.IsNullOrEmpty(cursor.View))
            { 
                db.ViewView(this.cursor.View);
            }
            int itemCount = db.ViewItemCount();

            for (int i = 1; i <= itemCount; i++) // note that we use a 1-based iterator
            {
                rows = new List<List<CommenceValue>>();
                rowvalues = new List<CommenceValue>();
                foreach (TableDef td in mocktables)
                {
                    string[] DDEResult = null;
                    List<string> fieldNames = td.ColumnDefinitions.Select(o => o.FieldName).ToList();
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
                                //buffer = FormatValues(buffer,this.Formatting, cd);
                                buffer = FormatValues(buffer, cd);
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
                                //query = FormatValues(query, this.Formatting, cd);
                                query = FormatValues(query, cd);
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
                CursorDataReadProgressChangedArgs args = new CursorDataReadProgressChangedArgs(rows, i, totalRows); // progress within the cursor
                OnDataProgressChanged(args);
            } // i
            db = null;
            ExportCompleteArgs a = new ExportCompleteArgs(itemCount);
            OnDataReadCompleted(a);
        }

        // based on SO feedback: https://stackoverflow.com/questions/53788909/convert-expensive-call-to-async-while-keeping-event-system-intact
        internal CancellationTokenSource CTS = new CancellationTokenSource();
        /// <summary>
        /// Reads the Commence database in a asynchronous fashion
        /// The idea is that the reading of Commence data continues as the event consumers do their thing.
        /// </summary>
        internal void GetDataByAPIAsync() // a bad method name, suggestion async
        {
            int rowsProcessed = 0;
            var values = new BlockingCollection<CmcData>();
            var readTask = Task.Factory.StartNew(() =>
            {
                try
                {
                    for (int rows = 0; rows < totalRows; rows += numRows)
                    {
                        string[][] rawdata = cursor.GetRawData(numRows); // first dimension is rows, second dimension is columns
                        {
                            if (CTS.Token.IsCancellationRequested) { break; }
                            rowsProcessed += numRows;
                            CmcData rowdata = new CmcData()
                            {
                                Data = rawdata,
                                RowsProcessed = rowsProcessed > totalRows ? totalRows : rowsProcessed
                            };
                            values.Add(rowdata);
                        }
                    }
                }
                catch
                {
                    CTS.Cancel(); // cancel data read
                    throw; // rethrow the event. If we didn't do this, all errors would be swallowed
                }
                finally
                {
                    values.CompleteAdding();
                }

            }, TaskCreationOptions.LongRunning);

            var processTask = Task.Factory.StartNew(() =>
            {
                foreach (var value in values.GetConsumingEnumerable())
                {
                    if (CTS.Token.IsCancellationRequested) { break; }

                    var data = ProcessDataBatch(value.Data);
                    CursorDataReadProgressChangedArgs args = new CursorDataReadProgressChangedArgs(data, value.RowsProcessed, totalRows);
                    OnDataProgressChanged(args); // raise event after each batch of rows
                }
            }, TaskCreationOptions.LongRunning); // longrunning is probably overkill here

            Task.WaitAll(readTask, processTask); // we need to wait all before we give the 'done' signal.
            values.Dispose();
            // raise 'done' event
            ExportCompleteArgs e = new ExportCompleteArgs(totalRows);
            OnDataReadCompleted(e); // done with reading data
        }
        #endregion

        #region Helper methods
        /// <summary>
        /// Takes raw cursor data and returns a list of CommenceValue lists
        /// </summary>
        /// <param name="rawdata">Data</param>
        /// <returns>A list of CommenceValue lists</returns>
        private List<List<CommenceValue>> ProcessDataBatch(string[][] rawdata)
        {
            List<List<CommenceValue>> retval = new List<List<CommenceValue>>();
            CommenceValue cv;
            ColumnDefinition cd;

            // rawdata represents the actual database row values
            for (int i = 0; i < rawdata.GetLength(0); i++) // rows
            {
                List<CommenceValue> rowdata = new List<CommenceValue>();
                // for thids we can assume the first row of rawdata contains the thid
                if (this.useThids)
                {
                    cv = new CommenceValue(rawdata[i][0], this.columnDefinitions.First()); // assumes thid column is first. This is an accident waiting to happen.
                    rowdata.Add(cv);
                }

                // process row
                for (int j = 1; j < rawdata[i].Length; j++) // columns
                {
                    // a column for the thid is only returned when a thid is requested
                    // therefore getting the right column is a little tricky
                    int colindex;
                    if (this.useThids)
                    {
                        colindex = j;
                    }
                    else
                    {
                        colindex = j - 1;
                    }
                    cd = this.columnDefinitions[colindex];

                    string[] buffer;
                    if (cd.IsConnection)
                    {
                        if (string.IsNullOrEmpty(rawdata[i][j].Trim()))
                        {
                            cv = new CommenceValue(cd); // always create a CommenceValue for consistency
                        }
                        else
                        {
                            if (!settings.SplitConnectedItems)
                            {
                                buffer = new string[] { rawdata[i][j] };
                            }
                            // We assume here that connected values are newline separated,
                            // depending on the type of cursor, they may not be!
                            // They will be if the cursor:
                            // - is of type View
                            // - has connected fields set using the SetRelatedColumn method
                            //
                            // they will be comma-separated when they were requested as direct fields on a Category-type cursor
                            // in that case, there is no meaningful way of splitting them,
                            // because there is no way to distinguish between commas in the field and separator
                            // That is just one of the quirks of Commence.
                            // If a user just dumps a cursor obtained by GetCursor(<category>, <flag All>),
                            // he will get weird results.
                            // We include a check for this
                            else if (cursor.CursorType == CmcCursorType.Category && !cursor._relatedColumnsWereSet)
                            {
                                // there *is* a situation when we do want to split on commas:
                                // when the cursor contains thids in the requested related column.
                                // It would actually be very useful in the special case of 'complex' exports.
                                // We *could* isolate that particular situation,
                                // and then split anyway, but then the (rest of the) row of CommenceValue
                                // would contain inconsistent arrays for field ConnectedFieldValues.
                                // So we are not going to do that.
                                buffer = new string[] { rawdata[i][j] };
                            }
                            else
                            {
                                switch (cd.CommenceFieldDefinition.Type)
                                {
                                    case CommenceFieldType.Text:
                                    case CommenceFieldType.URL:
                                        // any non-Name text field in Commence will accept a \n
                                        // we use a regex to split values at "\n" *but not* "\r\n"
                                        // this is not 100% fail-safe as a fieldvalue *can* contain just \n if it is a large text field.
                                        // in that case, your only option is to suppress the splitting in ExportSettings
                                        // what we *should* do is change every instance of a single \n to '\r\n' first
                                        // that would be safer, but we cannot distinguish between them.
                                        buffer = regex.Split(rawdata[i][j]); // this may result in Commence values being split if they contain embedded delimiters
                                        break;
                                    default:
                                        buffer = rawdata[i][j].Split(new string[] { cd.Delimiter }, StringSplitOptions.None);
                                        break;
                                } // switch

                                // buffer now contains the connected values as array, do any formatting transformation
                                buffer = FormatValues(buffer, cd);
                            }
                            cv = new CommenceValue(buffer, cd);
                        } // if !String.IsNullOrEmpty
                    } // if IsConnection
                    else // single value
                    {
                        buffer = new string[] { rawdata[i][j] };
                        buffer = FormatValues(buffer, cd);
                        cv = new CommenceValue(buffer[0], cd);
                    } // else IsConnection
                    if (cv != null) { rowdata.Add(cv); }
                } // for j
                retval.Add(rowdata);
            } // for i
            return retval;
        }

   private string[] FormatValues(string[] values, ColumnDefinition cd)
        {
            string[] retval = values;
            if (settings.Canonical)
            {
                for (int i = 0; i < retval.Length; i++)
                {
                    if (!string.IsNullOrEmpty(retval[i]))
                    {
                        retval[i] = CommenceValueConverter.ToCanonical(retval[i], cd.CommenceFieldDefinition.Type, settings);
                    }
                }
            }
            return retval;
        }

        private int BalanceNumRowsAndFieldSize(IExportSettings settings)
        {
            /* If the maxfieldsize is very large, decrease the number of rows being read.
             * This is largely untested.
            */
            int maxpower = 20;
            int threshold = (int)Math.Pow(2, maxpower);
            int i = 10;
            if (threshold < settings.MaxFieldSize)
            {
                while (Math.Pow(2, maxpower) < settings.MaxFieldSize)
                {
                    i--;
                    maxpower++;
                    if (i == 0) { break; }
                }
            }
            return i;
        }
        #endregion

        /// <summary>
        /// Helper class to capture the correct row.
        /// If we use a variable shared between the tasks, it may not reflect the correct value 
        /// </summary>
        private class CmcData
        {
            internal string[][] Data { get; set; }
            internal int RowsProcessed { get; set; }
        }
    }
}
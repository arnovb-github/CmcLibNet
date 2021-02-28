using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.IO;
using System.Linq;
using System.Text;
using Vovin.CmcLibNet.Database;
using Vovin.CmcLibNet.Database.Metadata;
using Vovin.CmcLibNet.Extensions;

/* On the surface, this looks a lot like the AdoNet classes 
 * There difference is that those classes use an in-memory dataset
 * That is a problem when dealing with connections
 * Connected data size can be very large, much larger than we can handle,
 * even with a huge max_field_size on a cursor,
 * There is also another problem: when you use a very large max_field_size,
 * reading data from Commence slows to a crawl.
 * Therefore we use a different approach here:
 * -We create a DataSet to represent the 'database' structure of the cursor.
 * -We create a temporary SQLite database from the DataSet.
 *  Strictly speaking that is not necessary, we could do everything in-memory,
 *  but it may be nice to have in the future.
 * -We read the 'primary' cursor data (=the category on which cursor or view was set)
 *  and just read the direct fields.
 * -We then read the primary category again and retrieve just the connection id's,
 *  which we store in link tables.
 * -We then dump all requested connected categories to separate tables
 *  This read cannot be filtered unfortunately, because that would require a 'reverse' filter,
 *  and it is impossible to match (paired) connections that Commence returns.
 *  (These are simply returned in alphabetical order)
 *  In other words: you can request all connected thids for:
 *  CategoryA - Relates to - CategoryB
 *  but when you want to read CategoryB and filter it for a connection
 *  back to CategoryA, you are out of luck.
 * -Using the DataSet, we use SQL and serializers to produce output.
 * 
 * This may seem backwards and slow, and while it is slow,
 * it is probably still the fastest way to retain all connections upon export.
 * The Commence data reading and processing by default happens asynchronously
 * and since the reading of Commence is a very slow process anyway,
 * we can do other heavy-lifting in the meantime.
 * 
 * A very important assumption we make here:
 * 1. The THID field is always in the first column.
 * We make sure of that ourselves, but that is important to understand the code.
 */
namespace Vovin.CmcLibNet.Export.Complex
{
    /// <summary>
    /// Writes to and exports from a SQLite database.
    /// </summary>
    internal sealed class SQLiteWriter : BaseWriter
    {

        #region Fields
        private readonly string _fkParamName = "@fk";
        private readonly string _pkParamName = "@pk";
        private readonly string _databaseName = string.Empty;
        private readonly string _cs = string.Empty; // connection string
        private readonly char[] _thidDelimiter = new char[] { ':' };
        private int _totalIterations;
        private int _iteration;
        private bool disposed = false;
        // if a cursor isn't shared, we need to calculate the thid sequence differently
        // because local cursors do not contain thids, only RowId's
        // we could probably get away with just using RowID's for our purpose
        // an important thing to note is that in shared cursors local items still get a thid(!)
        private bool _cursorShared;
        private string _fileName;
        private readonly ColumnDefinition[] originalColumnDefinitions;
        private readonly OriginalCursorProperties _ocp;
        private DataSet _ds;
        #endregion

        #region Constructors
        internal SQLiteWriter(ICommenceCursor cursor, IExportSettings settings) 
            : base(cursor, settings)
        {
            // store original cursor details, used in serializers
            _ocp = new OriginalCursorProperties(cursor);

            // put data in %AppData%
            _databaseName = Path.GetTempFileName();
#if DEBUG
            _databaseName = @"E:\Temp\sqlite.db";
            File.Delete(_databaseName);
#endif

            // We could do some optimization to see if the database would fit in memory;
            // an in-memory SQLite database will obviously perform better
            // there is caveat in that a SQLite In-memory database only exists while the connection is open.
            // It does not matter much because bottleneck will always be the Commence data reading
            // except for tiny exports, which will be fast (enough) anyway.
            // Testing shows that performance penalty for using an on-disk Sqlite database
            // is negligible compared to the slow Commence reading operation.
            // Note that since it is a discardible database, we can turn off some settings and gain a little performance
            // It is Write once, Read once (multiple queries), discard
            // SYNCHRONOUS = OFF : synchronous is only useful when worrying about syste crash or power off
            // JOURNAL MODE = OFF : do not keep a journal to do things like roll-backs
            // FOREIGN_KEYS = ON : can make lokups faster if implemented. Requires the REFERENCES keyword to be useful.
            // It is probably better to have a factory pattern that returns the appropriate string.
            _cs = @"URI=file:" + _databaseName + ";PRAGMA SYNCHRONOUS = OFF; JOURNAL MODE = OFF; FOREIGN_KEYS = ON";

            // Back up the original ColumnDefinitions as it contains the information on
            // what fields were originally requested.
            // We will discard the original cursor, so we need to store that information.
            originalColumnDefinitions = new ColumnDefinition[ColumnDefinitions.Count];
            ColumnDefinitions.CopyTo(originalColumnDefinitions, 0);

            // in case of JSON exports we always want nested connected items
            if (base._settings.ExportFormat == ExportFormat.Json) { base._settings.NestConnectedItems = true; }
        }

        // TODO for future support
        internal SQLiteWriter(ICommenceCursor cursor, string MemocomConfigFile, IExportSettings settings)
            : base(cursor, settings)
        {
            // we can immediately close the cursor, we're not using it
            _cursor.Close();

            // let's establish some of the checks that we have to do in processing the config file
            // does it exist?
            // is it XML?
            // are the prerequiste parameters populated? Which ones are they?
            // do the categories exist?
            // do the fields exist?
            // do the connections exist?
            // (is the filter(s) valid?
            // are there nested queries involved?
            // if nested, are there circular references?

            // then create a dataset
            // then define cursors to read
        }

        #endregion

        #region Finalizers
        ~SQLiteWriter()
        {
            Dispose(false);
        }
        #endregion

        protected internal override void WriteOut(string fileName)
        {
            _fileName = fileName;
            // create ADO.NET dataset to represent our data
            _ds = CreateDataSetFromCursor();
            // build a SQLite database from dataset
            CreateSqliteDatabase(_ds);

            // we do:
            // - a run on the primary cursor with just direct fields
            // - a run to retrieve only thids and connected thids from primary cursor
            // - a run for every connected category

            // close the original cursor
            // if it was a cursor with no connections at all, this is actually unneeded
            // but this way we can keep consistent
            // there is another benefit: max_field_size will now be better optimized
            _cursor.Close();
            _totalIterations = this.CursorDescriptors.Count(); // used in progress reports
            foreach (var cur in CursorDescriptors)
            {
                CurrentCursorDescriptor = cur;
                using (var cf = new CursorFactory())
                {
                    base._cursor = cf.Create(cur);
                    // we need to re-establish the columndefinitions,
                    // because the base reader needs them
                    ColumnParser cp = new ColumnParser(_cursor);
                    base.ColumnDefinitions = cp.ParseColumns();
                    _cursorShared = _cursor.Shared;
                    _iteration++;
                    base.ReadCommenceData();
                }
            }
        }

        #region Data processing methods
        protected internal override void HandleProcessedDataRows(object sender, CursorDataReadProgressChangedArgs e)
        {
            // it is interesting to note that the order actually matters for SQLite performance
            // I think Commence will return stuff in an sequential fashion, but we'd need to check.
            // We already know there will be gaps in the linktables.
            using (var con = new SQLiteConnection(_cs))
            {
                con.Open();
                using (var transaction = con.BeginTransaction())
                {
                    using (var cmd = new SQLiteCommand(con))
                    {
                        if (CurrentCursorDescriptor.IsTableWithConnectedThids)
                        {
                            ProcessLinkTableData(cmd, e.RowValues);
                        }
                        else
                        {
                            cmd.CommandText = GetInsertQueryForDirectTable();
                            ProcessDirectTableData(cmd, e.RowValues);
                        }
                    } // using cmd
                    transaction.Commit();
                } // using transaction
            } // using con

            // report progress
            ExportProgressChangedArgs args = new ExportProgressChangedArgs(
                e.RowsProcessed,
                e.RowsTotal,
                _iteration,
                _totalIterations);
            //base.OnWriterProgressChanged(args); // cannot use use BubbleUp method in base
            BubbleUpProgressEvent(args);
        }

        private int cursorsProcessed = 0;
        protected internal override void HandleDataReadComplete(object sender, ExportCompleteArgs e)
        {
            base._cursor.Close(); // very important or we'll leave hanging COM references!
            // fires for every cursor,
            // but we want to wait until they are all processed
            cursorsProcessed++;
            if (cursorsProcessed < CursorDescriptors.Count()) { return; } // not done reading yet

            if (_settings.WriteSchema)
            {
                FillDataSet(); // may be too large
                DataSetSerializer dse = new DataSetSerializer(_ds, _fileName, _settings);
                dse.Export();
            }
            else
            {
                switch (_settings.ExportFormat)
                {
                    case ExportFormat.Xml:
                        var xw = new SQLiteToXmlSerializer(this._settings, _ds.Tables[_ocp.Name], _cs);
                        xw.Serialize(_fileName);
                        break;
                    case ExportFormat.Json:
                        var jw = new SQLiteToJsonSerializer(this._settings, _ocp, _ds.Tables[_ocp.Name], _cs);
                        jw.Serialize(_fileName);
                        break;
                    case ExportFormat.Excel: // largely untested
                        FillDataSet(); // may be too large
                        DataSetSerializer dse = new DataSetSerializer(_ds, _fileName, _settings);
                        dse.Export(); // TODO: fails if in use, does not respect Excel options yet
                        break;
                }
            }

            // bubble up event
            base.BubbleUpCompletedEvent(e);
        }

        private void ProcessDirectTableData(SQLiteCommand cmd, List<List<CommenceValue>> rowValues)
        {
            // this method only provides the parameters for the command
            cmd.Prepare();
            for (int i = 0; i < rowValues.Count(); i++)
            {
                for (int j = 0; j < rowValues[i].Count(); j++)
                {
                    string value = rowValues[i][j].DirectFieldValue;
                    // first column is special case
                    if (j == 0) { value = SequenceFromThid(value, _cursorShared).ToString(); }
                    cmd.Parameters.AddWithValue($"@p{j}", value);
                }
                cmd.ExecuteNonQuery();
            }
        }

        private void ProcessLinkTableData(SQLiteCommand cmd, List<List<CommenceValue>> rowValues)
        {
            // the input this expects is a list of one or more columns in which every column contains 
            // both the THID of the parent table
            // and a CommenceValue with the connected THIDS
            // think of it as a Commence view that only contains connected values

            // this method requires that we get/set the CommandText on the command
            //IList<string> commandTexts = GetInsertCommandTextsForLinkTables().ToArray();
            var parentTable = _ds.Tables[CurrentCursorDescriptor.CategoryOrView];
            int pairedColumnIndex = 0;
            // notice that we do this column-first, it is faster.
            for (int col = 0; col < rowValues[0].Count() - 1; col++) // notice the -1, because of the thid being 0.
            {
                pairedColumnIndex++;
                cmd.CommandText = parentTable
                                    .ChildRelations[col] // requires that they match up. Fragile.
                                    .ChildTable
                                    .ExtendedProperties[DataSetHelper.LinkTableInsertCommandTextExtProp]
                                    .ToString();
                cmd.Prepare();
                for (int row = 0; row < rowValues.Count(); row++)
                {
                    // check if there are values to process
                    if (rowValues[row][pairedColumnIndex].ConnectedFieldValues is null) { continue; } // fields linktable all must contain a value, skip row
                    // cache thid of primary table
                    int pkVal = SequenceFromThid(rowValues[row][0].DirectFieldValue, _cursorShared);
                    // add parameters
                    for (int i = 0; i < rowValues[row][pairedColumnIndex].ConnectedFieldValues.Length; i++)
                    {
                        string s = rowValues[row][pairedColumnIndex].ConnectedFieldValues[i];
                        int fkVal = SequenceFromThid(s, _cursorShared);
                        cmd.Parameters.AddWithValue(_pkParamName, pkVal);
                        cmd.Parameters.AddWithValue(_fkParamName, fkVal);
                        cmd.ExecuteNonQuery();
                    }
                }
            }
        }
        #endregion

        /// <summary>
        /// A pretty complex method. It translates the Commence cursor to a ADO.NET DataSet
        /// </summary>
        /// <returns>DataSet.</returns>
        private DataSet CreateDataSetFromCursor()
        {
            string dataSource = string.IsNullOrEmpty(_cursor.View) ? _cursor.Category : _cursor.View;
            DataSet retval = new DataSet(dataSource);
            DataTable dt;
            // construct primary table
            dt = new DataTable(dataSource);
            // create fields for primary column, if any
            var directFields = ColumnDefinitions.Where(f => !f.IsConnection)
                .ToList();
            foreach (var df in directFields)
            {
                var dc = new DataColumn
                {
                    ColumnName = df.FieldName, // *must* be the fieldName because we create cursors on based on it
                    DataType = df.CommenceFieldDefinition.Type.GetTypeForCommenceField(),
                    // if custom headers were supplied, store them in caption
                    // remember that columnLabel defaults to fieldName
                    Caption = string.IsNullOrEmpty(df.CustomColumnLabel) ? df.ColumnLabel : df.CustomColumnLabel,
                };
                OverrideDefaultThidDataType(dc);
                // store the Commence fieldtype so we can distinguish between Date end Time later on.
                dc.ExtendedProperties.Add(DataSetHelper.CommenceFieldTypeDescriptionExtProp, df.CommenceFieldDefinition.Type);
                dt.Columns.Add(dc);
            }
            dt.PrimaryKey = new DataColumn[] { dt.Columns[ColumnDefinitions[0].FieldName] };
            retval.Tables.Add(dt);
            PopulateCursorDescriptors(dt, dataSource, true);

            // we do not have connections, return
            if (!ColumnDefinitions.Where(w => w.IsConnection).Any()) { return retval; }

            // connections in Commence are case sensitive
            // but in ADO.NET a DateRelation's relationName is not
            // we do not support that at this moment
            var possibleDups = ColumnDefinitions.Select(s => new { s.Connection, s.Category });
            foreach (var d in possibleDups)
            {
                if (ColumnDefinitions.Where(w => w.Category.Equals(d.Category)).Select(s => s.Connection).Distinct(StringComparer.CurrentCultureIgnoreCase).Count() !=
                    ColumnDefinitions.Where(w => w.Category.Equals(d.Category)).Select(s => s.Connection).Distinct().Count())
                {
                    throw new DuplicateNameException("Commence contains case sensitive connections with same name. Exporting these is not supported.");
                }
            }

            // create the tables for the connected categories
            // but build just a single datatable per category even if multiple relations exist
            var connectedCategories = ColumnDefinitions.Where(w => w.IsConnection)
                .Select(s => s.Category)
                .Distinct()
                .ToArray();
            foreach (string cat in connectedCategories)
            {
                var connectedFields = ColumnDefinitions.Where(w => w.IsConnection && w.Category.Equals(cat))
                    .Select(s => s.FieldName)
                    .Distinct()
                    .ToArray();
                // we need a prefix or we'd get an error on single-connected categories
                // the alternative would be to make single connections a special case,
                // but that would hurt code consistency.
                dt = new DataTable(DataSetHelper.ConnectedCategoryPrefix + cat);
                // the end-user never sees the actual tablenames in the output
                // except in the Excel export.
                // That is why we will include a property that just holds the category name
                dt.ExtendedProperties.Add(DataSetHelper.CommenceCategoryNameExtProp, cat); // TODO fix for single connections
                InsertThidField(dt);
                dt.PrimaryKey = new DataColumn[] { dt.Columns[0] };
                foreach (var cf in connectedFields)
                {
                    var cd = ColumnDefinitions.First(s => s.FieldName.Equals(cf) && s.Category.Equals(cat));
                    var dc = new DataColumn()
                    {
                        ColumnName = cd.FieldName,
                        DataType = cd.CommenceFieldDefinition.Type.GetTypeForCommenceField(),
                        Caption = string.IsNullOrEmpty(cd.CustomColumnLabel) ? cd.ColumnLabel : cd.CustomColumnLabel
                    };
                    OverrideDefaultThidDataType(dc);
                    // store the Commence fieldtype so we can distinguish between Date end Time later on.
                    dc.ExtendedProperties.Add(DataSetHelper.CommenceFieldTypeDescriptionExtProp, cd.CommenceFieldDefinition.Type);
                    dt.Columns.Add(dc); // throws error on duplicate THID column
                }
                retval.Tables.Add(dt);
                PopulateCursorDescriptors(dt, cat, false);
            }

            // create link tables
            // it is important to understand that for any cursor, 
            // is is assumed that the THID is always the first element
            var linkTables = ColumnDefinitions.Where(w => w.IsConnection)
                //.Select(s => new { s.Connection, s.Category })
                .Select(s => new CommenceConnection()
                {
                    Name = s.Connection,
                    ToCategory = s.Category
                })
                .DistinctBy(d => d.FullName) // notice the use of DistinctBy(), not Distinct()
                .ToList();
            foreach (var lt in linkTables)
            {
                string name = DataSetHelper.TableName(_cursor.Category, lt.Name, lt.ToCategory);
                dt = new DataTable(name);
                string fkp = DataSetHelper.ForeignKeyOfPrimaryTable(_cursor.Category, DataSetHelper.PostFixId);
                var dc = new DataColumn(fkp, typeof(int))
                {
                    AllowDBNull = false
                };
                dt.Columns.Add(dc);
                string fkc = DataSetHelper.ForeignKeyOfConnectedTable(lt.Name, lt.ToCategory, DataSetHelper.PostFixId);
                dc = new DataColumn(fkc, typeof(int))
                {
                    AllowDBNull = false
                };
                dt.Columns.Add(dc);
                // set compound primary key
                dt.PrimaryKey = new DataColumn[2] { dt.Columns[0], dt.Columns[1] };
                retval.Tables.Add(dt);
                // include the INSERT query
                dt.ExtendedProperties.Add(DataSetHelper.LinkTableInsertCommandTextExtProp, GetInsertQueryForLinkTable(dt.TableName, fkp, fkc));

                // set up the relations
                string rel1 = lt.Name + lt.ToCategory + _cursor.Category;
                DataRelation relPrimaryTableToLinkTable = retval.Relations.Add(rel1,
                    retval.Tables[dataSource].Columns[0],
                    retval.Tables[dt.TableName].Columns[fkp],
                    false);
                relPrimaryTableToLinkTable.Nested = false;

                // it may be handy to include the connection information in it,
                // so that from the DataSet we can easily identify which Commence connection
                // the relationship describes
                relPrimaryTableToLinkTable.ExtendedProperties.Add(DataSetHelper.CommenceConnectionDescriptionExtProp, JsonConvert.SerializeObject(lt));

                // from link table to connected table
                string rel2 = lt.Name + lt.ToCategory + lt.ToCategory;
                // test for single connections (these are only possible on a category itself. Maybe these should be reversed?)
                if (!rel1.Equals(rel2))
                {
                    // note that we DO NOT set the reverse Commence connection
                    // That is because even though there may be a reverse connection (provided the connnection is paired),
                    // there is no way to obtain that information from the Commence API.
                    // In Commence, the name of the reverse connection can be any string,
                    // we do not know anything about it at this point.
                    DataRelation relConnectedTableToLinkTable = retval.Relations.Add(rel2,
                        retval.Tables[DataSetHelper.ConnectedCategoryPrefix + lt.ToCategory].Columns[0],
                        //retval.Tables[lt.ToCategory].Columns[0],
                        retval.Tables[dt.TableName].Columns[fkc],
                        false);
                    relConnectedTableToLinkTable.Nested = false;
                    relConnectedTableToLinkTable.ExtendedProperties.Add(DataSetHelper.CommenceConnectionDescriptionExtProp, JsonConvert.SerializeObject(lt));
                    // include the SELECT query
                    dt.ExtendedProperties.Add(DataSetHelper.LinkTableSelectCommandTextExtProp,
                        GetSelectQueryForLinkTable(relPrimaryTableToLinkTable, fkp, fkc, DataSetHelper.ConnectedCategoryPrefix + lt.ToCategory));
                }
                else
                {
                    // we have a single connection on the category itself,
                    // so we must alter the select query to reflect that we are in fact reading a connection
                    dt.ExtendedProperties.Add(DataSetHelper.LinkTableSelectCommandTextExtProp,
                        GetSelectQueryForLinkTable(relPrimaryTableToLinkTable, fkp, fkc, DataSetHelper.ConnectedCategoryPrefix + lt.ToCategory, true));
                }
            } // foreach
            return retval;
        }

        private void CreateSqliteDatabase(DataSet ds)
        {
            using (var con = new SQLiteConnection(_cs))
            {
                con.Open();
                using (var transaction = con.BeginTransaction())
                {
                    foreach (DataTable dt in ds.Tables)
                    {
                        string ct = GetSqlCreateTableQuery(dt);
                        using (var cmd = new SQLiteCommand(ct, con))
                        {
                            cmd.ExecuteNonQuery();
                        } // using cmd
                    } // foreach
                    transaction.Commit();
                } // using transaction
            } // using con
        }

        private void FillDataSet()
        {
            using (var con = new SQLiteConnection(_cs))
            {
                foreach (DataTable dt in _ds.Tables)
                {
                    string stm = GetSQLiteSelectQueryForTable(dt); // we need a function for this
                    using (var da = new SQLiteDataAdapter(stm, con))
                    {
                        da.Fill(_ds, dt.TableName); // FYI: Fill will add a new column if it does not exist
                    }
                }
            }
        }
        #region Query methods
        internal static string GetSQLiteSelectQueryForTable(DataTable dt)
        {
            StringBuilder sb = new StringBuilder($"SELECT ");
            List<string> cols = new List<string>();
            foreach (DataColumn dc in dt.Columns)
            {
                cols.Add(GetSQLiteSelectColumnSyntax(dc));
            }
            sb.Append(string.Join(",", cols));
            sb.Append($" FROM {SanitizeSqlIdentifier(dt.TableName)}");
            return sb.ToString();
        }

        private string GetSelectQueryForLinkTable(
                DataRelation dr,
                string fkPrimary,
                string fkConnected,
                string connectedTableName,
                bool isSingleConnection = false)
        {
            // would this work with two relations to same table? 
            // Yes, because they are distinct objects
            IDictionary<DataTable, SQLiteCommand> retval = new Dictionary<DataTable, SQLiteCommand>();
            StringBuilder sb = new StringBuilder("SELECT ");

            // get a datatable with only the requested connected fields from Commence
            // this method call will also deserialize to CommenceConnection
            var leftTable = GetDataTableWithOnlyRequestedFields(dr); // table with the connected data, not the linktable
            IList<string> colNames = new List<string>();

            /* Query should look like this:
                * SELECT b.Title FROM Books AS b
                JOIN AuthorBook AS ab ON ab.BookID = b.BookID
                JOIN Authors AS a ON ab.AuthorID = a.AuthorID
                WHERE a.AuthorID = @id
                */
            foreach (DataColumn dc in leftTable.Columns)
            {
                // SELECT queries rely on aliases
                // if we have a date of time column, there is no field name in the select query;
                // (it will be: function(columnname) AS alias)
                // not a problem, we can check for that.
                // However, if nested connections were requested,
                // we do not want the entire connection description in the fieldname
                string customAlias = string.Empty;
                if (dc.DataType == typeof(DateTime) && _settings.NestConnectedItems)
                {
                    customAlias = GetAliasForColumn(dc);
                }

                if (!isSingleConnection)
                {
                    colNames.Add(GetSQLiteSelectColumnSyntax("b.", dc, customAlias));
                }
                // for data requested if for a Commence category with a single connection to itself
                // this is typically very rare (I've never used one in 20 years)
                // in the output, the fields would look like belonging to the primary category (and they do);
                // we have to make it explicit that they are in fact coming from a connection
                else
                {
                    // no check if property exists
                    CommenceConnection cc = JsonConvert.DeserializeObject<CommenceConnection>(
                        dr.ExtendedProperties[DataSetHelper.CommenceConnectionDescriptionExtProp].ToString());
                    if (_settings.NestConnectedItems) // it is fine to have the real columnname in nested exports
                    {
                        customAlias = GetAliasForColumn(dc);
                    }
                    else
                    {
                        customAlias = cc.Name + cc.ToCategory + GetAliasForColumn(dc);
                    }
                    colNames.Add(GetSQLiteSelectColumnSyntax("b.", dc, customAlias));
                }
            }
            sb.Append(string.Join(",", colNames));
            sb.Append(" FROM ");
            sb.Append(SanitizeSqlIdentifier(leftTable.TableName));
            sb.Append(" AS b");
            // linktable to connected table
            sb.Append(" JOIN ");
            sb.Append(SanitizeSqlIdentifier(dr.ChildTable.TableName));
            sb.Append(" AS ab ON ab.");
            sb.Append(SanitizeSqlIdentifier(fkConnected));
            sb.Append("=b.");
            sb.Append(SanitizeSqlIdentifier(leftTable.PrimaryKey[0].ColumnName));
            // linktable to primary table
            sb.Append(" JOIN ");
            sb.Append(SanitizeSqlIdentifier(dr.ParentTable.TableName)); // should be primary table
            sb.Append(" AS a ON a.");
            sb.Append(SanitizeSqlIdentifier(dr.ParentTable.PrimaryKey[0].ColumnName));
            sb.Append("=ab.");
            sb.Append(SanitizeSqlIdentifier(fkPrimary));
            sb.Append($" WHERE a.{SanitizeSqlIdentifier(dr.ParentTable.PrimaryKey[0].ColumnName)} = @id");
            return sb.ToString();
        }

        internal static string GetSQLiteSelectColumnSyntax(DataColumn dc, string customAlias = "")
        {
            string colName = SanitizeSqlIdentifier(dc.ColumnName);
            string alias = string.IsNullOrEmpty(customAlias) 
                ? SanitizeSqlIdentifier(dc.Caption)
                : SanitizeSqlIdentifier(customAlias);
            // process special cases
            // refer: https://www.sqlitetutorial.net/sqlite-date-functions/sqlite-date-function/
            if (dc.DataType == typeof(DateTime)
                && dc.ExtendedProperties.ContainsKey(DataSetHelper.CommenceFieldTypeDescriptionExtProp)
                && dc.ExtendedProperties[DataSetHelper.CommenceFieldTypeDescriptionExtProp].Equals(CommenceFieldType.Date))
            {
                colName = $"date({colName})"; // returns YYYY-MM-DD
            }
            else if (dc.DataType == typeof(DateTime)
                && dc.ExtendedProperties.ContainsKey(DataSetHelper.CommenceFieldTypeDescriptionExtProp)
                && dc.ExtendedProperties[DataSetHelper.CommenceFieldTypeDescriptionExtProp].Equals(CommenceFieldType.Time))
            {
                colName = $"time({colName})"; // returns HH:MM:SS
            }
            colName = $"{colName} AS {alias}";
            return colName;
        }

        internal static string GetSQLiteSelectColumnSyntax(string prefix, DataColumn dc, string customAlias)
        {
            string colName = SanitizeSqlIdentifier(dc.ColumnName);
            string alias = string.IsNullOrEmpty(customAlias)
                ? SanitizeSqlIdentifier(dc.Caption)
                : SanitizeSqlIdentifier(customAlias);
            // process special cases
            // refer: https://www.sqlitetutorial.net/sqlite-date-functions/sqlite-date-function/
            if (dc.DataType == typeof(DateTime)
                && dc.ExtendedProperties.ContainsKey(DataSetHelper.CommenceFieldTypeDescriptionExtProp)
                && dc.ExtendedProperties[DataSetHelper.CommenceFieldTypeDescriptionExtProp].Equals(CommenceFieldType.Date))
            {
                return colName = $"date({prefix}{colName}) AS {alias}"; // returns YYYY-MM-DD
            }
            else if (dc.DataType == typeof(DateTime)
                && dc.ExtendedProperties.ContainsKey(DataSetHelper.CommenceFieldTypeDescriptionExtProp)
                && dc.ExtendedProperties[DataSetHelper.CommenceFieldTypeDescriptionExtProp].Equals(CommenceFieldType.Time))
            {
                return colName = $"time({prefix}{colName}) AS {alias}"; // returns HH:MM:SS
            }
            return colName = $"{prefix}{colName} AS {alias}";
        }

        private string GetInsertQueryForDirectTable()
        {
            StringBuilder sb = new StringBuilder("INSERT INTO ");
            sb.Append($"'{CurrentCursorDescriptor.SqlColumnMappings.First().Value.TableName}'");
            sb.Append('(');
            sb.Append(string.Join(",", CurrentCursorDescriptor.SqlColumnMappings.Select(s => $"'{s.Value.ColumnName}'")));
            sb.Append(") VALUES(");
            // just use numeric placeholders
            int[] elements = new int[CurrentCursorDescriptor.SqlColumnMappings.Count];
            for (int i = 0; i < elements.Length; i++)
            {
                elements[i] = i;
            }
            sb.Append(string.Join(",", elements.Select(s => $"@p{s}")));
            sb.Append(')');
            return sb.ToString();
        }
        
        //private IEnumerable<string> GetInsertQueriesForLinkTables()
        //{
        //    var linkTables = ColumnDefinitions.Where(w => w.IsConnection)
        //        .Select(s => new { s.Connection, s.Category })
        //        .Distinct()
        //        .ToList();
        //    foreach (var lt in linkTables)
        //    {
        //        // a link table always contains just 2 columns
        //        StringBuilder sb = new StringBuilder("INSERT INTO ");
        //        string s = DataSetHelper.TableName(_cursor.Category, lt.Connection, lt.Category);
        //        sb.Append(SanitizeSqlIdentifier(s)); // expects name of link table
        //        sb.Append('(');
        //        s = DataSetHelper.ForeignKeyOfPrimaryTable(_cursor.Category, DataSetHelper.PostFixId);
        //        sb.Append(SanitizeSqlIdentifier(s)); // primary key of primary category.
        //        sb.Append(',');
        //        s = DataSetHelper.ForeignKeyOfConnectedTable(lt.Connection, lt.Category, DataSetHelper.PostFixId);
        //        sb.Append(SanitizeSqlIdentifier(s)); // the foreign key.
        //        sb.Append(") VALUES (");
        //        sb.Append(_pkParamName);
        //        sb.Append(',');
        //        sb.Append(_fkParamName);
        //        sb.Append(')');
        //        yield return sb.ToString();
        //    }
        //}

        private string GetInsertQueryForLinkTable(
                string tableName,
                string fkPrimary,
                string fkCconnected)
        {
            // a link table always contains just 2 columns
            StringBuilder sb = new StringBuilder("INSERT INTO ");
            sb.Append(SanitizeSqlIdentifier(tableName));
            sb.Append('(');
            sb.Append(SanitizeSqlIdentifier(fkPrimary)); // primary key of primary category.
            sb.Append(',');
            sb.Append(SanitizeSqlIdentifier(fkCconnected)); // the foreign key.
            sb.Append(") VALUES (");
            sb.Append(_pkParamName);
            sb.Append(',');
            sb.Append(_fkParamName);
            sb.Append(')');
            return sb.ToString();
        }

        private void OverrideDefaultThidDataType(DataColumn dc)
        {
            if (dc.ColumnName.Equals(ColumnDefinition.ThidIdentifier))
            {
                dc.DataType = typeof(int);
                dc.AllowDBNull = false;
            } // we will use int as datatype even tho a thid is a string
        }

        private string GetSqlCreateTableQuery(DataTable dt)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("CREATE TABLE ");
            sb.Append(SanitizeSqlIdentifier(dt.TableName));
            sb.Append('(');
            var columnCommands = GetSqlCreateColumnCommands(dt.Columns).ToArray();
            sb.Append(string.Join(",", columnCommands));
            sb.Append(",PRIMARY KEY(");
            sb.Append(string.Join(",", dt.PrimaryKey.Select(s => SanitizeSqlIdentifier(s.ColumnName))).ToArray());
            // TODO should we include logic for including a 'REFERENCES' keyword?
            sb.Append(")) WITHOUT ROWID"); // we supply our own primary keys so no need for a rowid. Does not work nicely with LinqPad.
            return sb.ToString();
        }

        private IEnumerable<string> GetSqlCreateColumnCommands(DataColumnCollection columns)
        {
            StringBuilder sb = new StringBuilder();
            foreach (DataColumn c in columns)
            {
                sb.Clear();
                sb.Append($"'{c.ColumnName}'");
                sb.Append(' ');
                // type
                // nullable
                if (!c.AllowDBNull)
                {
                    sb.Append("INTEGER NOT NULL"); // only thid fields
                }
                else
                {
                    sb.Append(c.DataType.Name);
                }
                // we do not set primary key here!
                yield return sb.ToString();
            }
        }

        // When the dataset is constructed, fields from different connections
        // that belong to the same category are aggregated,
        // so we need a way to figure out what fields were originally requested for what connection.
        // The original cursor's columndefinition still holds that information.
        private DataTable GetDataTableWithOnlyRequestedFields(DataRelation relation)
        {
            
            DataTable dt = null;
            CommenceConnection cc = null;
            if (relation.ExtendedProperties.ContainsKey(DataSetHelper.CommenceConnectionDescriptionExtProp))
            {
                cc = JsonConvert.DeserializeObject<CommenceConnection>(relation.ExtendedProperties[DataSetHelper.CommenceConnectionDescriptionExtProp].ToString());
            }
            if (cc is null) { return dt; }

            IEnumerable<ColumnDefinition> columnDefinitions = originalColumnDefinitions.Where(w => w.IsConnection
                    && w.Connection.Equals(cc.Name)
                    && w.Category.Equals(cc.ToCategory)).ToArray();
            if (columnDefinitions.Any())
            {
                // the whole point of this method is selecting a subset of its fields
                // what we can do is use the existing columns in it more easily create our own columns
                // the connected table is not in the relationship, but it is in the connection.
                string tableName = DataSetHelper.ConnectedCategoryPrefix + cc.ToCategory; // fragile
                //string tableName = cc.ToCategory;
                DataTable fullTable = relation.DataSet.Tables[tableName];
                dt = fullTable.Clone();
                var requestedFields = columnDefinitions.Select(s => s.FieldName).ToArray();
                foreach (DataColumn column in fullTable.Columns)
                {
                    // remove unneeded columns
                    if (!requestedFields.Contains(column.ColumnName) && !fullTable.PrimaryKey.Contains(column))
                    {
                        dt.Columns.Remove(column.ColumnName);
                    }
                }
            }
            return dt;
        }
        #endregion

        #region Helper methods

        // Notice we make it an int
        // This is because that way we can make SQLite faster
        // it means that we must not store the entire thid (which is a string)
        // but only the sequence number bit
        private void InsertThidField(DataTable dt)
        {
            DataColumn dc = new DataColumn(ColumnDefinitions[0].FieldName, typeof(int))
            {
                AllowDBNull = false
            };
            dt.Columns.Add(dc);
        }

        // a thid comes in the form of a:b:c (shared item) or a:b:c:d (local item, technically this is a rowid)
        private int SequenceFromThid(string thid, bool shared)
        {
            if (!shared)
            {
                // when requesting a thid for a local category,
                // you get the rowid instead (q:x:y:z)
                // in a rowid for a shared field the sequence is in the second element of four
                // Commence documentation states: is valid across cursor sessions
                return FromHex(thid.Split(_thidDelimiter)[1]);
            }
            // a thid consists of 3 elements (x:y:z), the sequence number is the last
            // if the cursor is shared, we will get a thid even if the item itself is local
            return FromHex(thid.Split(_thidDelimiter).Last());
        }

        private int FromHex(string value)
        {
            //// strip the leading 0x (not needed here)
            //if (value.StartsWith("0x", StringComparison.OrdinalIgnoreCase))
            //{
            //    value = value.Substring(2);
            //}
            return int.Parse(value, System.Globalization.NumberStyles.HexNumber);
        }

        internal void PopulateCursorDescriptors(DataTable dt, string categoryName, bool isPrimary)
        {
            CursorDescriptor cp;
            CmcCursorType cursorType = ((CommenceCursor)_cursor).CursorType; // ICommenceCursor doesn't expose this property
            // for the creation of the filters, we need the name fields of all involved categories
            // we also must count if we do not exceed 8 filters
            // if we do, do not apply filtering at all
            if (isPrimary)
            {
                // this will lose the current filter if it is on a category!
                // create a cursor with all direct fields
                var directFields = dt.Columns
                    .Cast<DataColumn>()
                    .Skip(1) // skip 1 because of thid field
                    .Select(s => s.ColumnName)
                    .ToArray();

                cp = new CursorDescriptor(categoryName)
                {
                    Fields = directFields,
                    CursorType = cursorType,
                    MaxFieldSize = (int)Math.Pow(2, 15) // 32.768
                };
                cp.CreateSqlMapping(dt);
                CursorDescriptors.Add(cp);

                // create a second cursor with only the connected fields so we read less data.
                // You can SetColumn on a cursor just once. There is no UnsetColumn
                // Note that we do retain the cursor type,
                // if we would not do that, we'd read the entire category
                var distinctConnections = ColumnDefinitions.Where(w => w.IsConnection)
                    .Select(s => new { s.Connection, s.Category })
                    .Distinct()
                    .ToArray();
                
                if (distinctConnections.Any())
                {
                    IDictionary<int, SqlMap> dict = new Dictionary<int, SqlMap>();
                    // okay, we need to define cursor columns for these
                    IList<string> conFields = new List<string>();
                    for (int i = 0; i < distinctConnections.Count(); i++)
                    {
                        conFields.Add(distinctConnections[i].Connection + ' ' + distinctConnections[i].Category);
                        // we screate a custom SqlMap object here
                        // because there is no corresponding table in the dataset
                        dict.Add(i, new SqlMap(_cursor.Category + distinctConnections[i].Connection + distinctConnections[i].Category,
                            distinctConnections[i].Category + DataSetHelper.PostFixId,
                            true));
                    }

                    // capture the parameters
                    cp = new CursorDescriptor(categoryName)
                    {
                        Fields = conFields,
                        MaxFieldSize = CommenceLimits.MaxItems * (CommenceLimits.ThidLength + 2), // 11.500.000
                        CursorType = cursorType,
                        IsTableWithConnectedThids = true,
                        SqlColumnMappings = dict
                    };
                    CursorDescriptors.Add(cp);
                }
            }
            else // connected categories
            {
                // just direct fields
                // we also want to add filters here!
                // but we cannot because Commence exposes no mechanism of matching paired connections
                // Therefore, if a category has multiple connections to the same other category
                // you cannot determine which 'reverse' filter to apply
                // We could apply a filter if just a single connection exists
                // between the primary category and the connected category

                // what we will do instead is take all the connections and filter them all.
                // At least that way we stand a chance of retrieving fewer items
                IList<ICursorFilterTypeCTCF> filters = GetFilters(categoryName, _cursor.Category).ToArray();

                // capture the parameters
                cp = new CursorDescriptor(categoryName)
                {
                    Fields = dt.Columns
                        .Cast<DataColumn>()
                        .Where(w => w.ColumnName != ColumnDefinition.ThidIdentifier)
                        .Select(c => c.ColumnName)
                        .ToList(),
                    MaxFieldSize = (int)Math.Pow(2, 15), // 32.768‬
                    CursorType = CmcCursorType.Category,
                    Filters = filters
                };
                cp.CreateSqlMapping(dt);
                CursorDescriptors.Add(cp);
            }
        }

        private IEnumerable<ICursorFilterTypeCTCF> GetFilters(string fromCategory, string toCategory)
        {
            using (ICommenceDatabase db = new CommenceDatabase())
            {
                // we need only the ones to the primary table
                var cons = db.GetConnectionNames(fromCategory)
                    .Where(w => w.ToCategory.Equals(toCategory))
                    .ToArray();
                if (cons.Any() && cons.Count() <= CommenceLimits.MaxFilters)
                {
                    for (int i = 0; i < cons.Count(); i++)
                    {
                        string fn = db.GetNameField(toCategory);
                        if (string.IsNullOrEmpty(fn)) { yield break; };
                        ICursorFilterTypeCTCF f = new CursorFilterTypeCTCF(i + 1)
                        {
                            Connection = cons[i].Name,
                            Category = cons[i].ToCategory,
                            FieldName = fn,
                            FieldValue = "?",
                            OrFilter = true,
                            Qualifier = FilterQualifier.Contains,
                        };
                        yield return f;
                    }
                }
            }
        }

        internal static string SanitizeSqlIdentifier(string s)
        {
            if (s is null)
            {
                throw new ArgumentNullException();
            }
            // escape double quotes
            s = s.Replace("\"", "\"\"");
            // double quote return value
            return string.Format("\"{0}\"", s);
        }

        private string GetAliasForColumn(DataColumn dc)
        {
            if (dc.DataType == typeof(DateTime))
            {
                return dc.ColumnName;
            }
            else
            {
                return string.IsNullOrEmpty(dc.Caption)
                                ? dc.ColumnName
                                : dc.Caption;
            }
        }
        #endregion

        #region IDisposable
        // Protected implementation of Dispose pattern.
        /// <summary>
        /// Dispose method.
        /// </summary>
        /// <param name="disposing">disposing.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposed)
                return;

            if (disposing)
            {
                // Free any other managed objects here.
                //
                try
                {
#if DEBUG == false
                    File.Delete(_databaseName);
#endif
                }
                catch { }
            }

            // Free any unmanaged objects here.
            //
            disposed = true;

            // Call the base class implementation.
            base.Dispose(disposing);
        }
        #endregion

        #region Properties
        private IList<CursorDescriptor> CursorDescriptors { get; } = new List<CursorDescriptor>();
        // allows us to keep track of what CursorDescriptor we are currently processing
        private CursorDescriptor CurrentCursorDescriptor { get; set; }
        #endregion
    }
}
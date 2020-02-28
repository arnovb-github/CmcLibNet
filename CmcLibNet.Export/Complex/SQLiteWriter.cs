using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.IO;
using System.Linq;
using System.Text;
using Vovin.CmcLibNet.Database;
using Vovin.CmcLibNet.Extensions;

/* On the surface, this looks a lot like the AdoNet classes 
 * There difference is that those classes use an in-memory dataset
 * That is a problem when dealing with connections
 * Connected data size can be very large, much larger than we can handle 
 * even with a huge max_field_size on a cursor
 * Therefore we use a different approach here:
 * -We still create a DataSet to represent our database structure
 * -We create a temporary SQLite database from the DataSet
 *  Strictly speaking that is not necessary, we could do everything in-memory,
 *  but it may be nice to have in the future.
 * -We read the 'primary' cursor data (=the category on which cursor or view was set)
 *  and just read the direct fields.
 * -We then read the primary category again and retrieve just the connection id's
 *  which we store in link tables.
 * -We then dump all requested connected categories to separate tables
 *  This read cannot be filtered unfortunately, because that would require a 'reverse' filter,
 *  and it is impossible to match (paired) connections that Commence returns.
 *  (They are simply returned in alphabetical order)
 *  In other words: you can request all connected thids for:
 *  CategoryA - Relates to - CategoryB
 *  but when you want to read CategoryB and filter for a connection back to CategoryA,
 *  you are out of luck.
 * -Using the DataSet, we use SQL and serializers to produce output.
 * 
 * This may seem backwards and slow, and while it is slow,
 * it is probably still the fastest way to retain all connections upon export.
 * The Commence data reading and processing by default happens asynchronously
 * and since the reading of Commence is a very slow process,
 * we can do other heavy-lifting in the meantime.
 * 
 * A very important assumption we make here: the THID field is always in the first column.
 * We make sure of that ourselves, but that is vitally important to understand the code.
 */
namespace Vovin.CmcLibNet.Export.Complex
{
    /// <summary>
    /// Writes to and exports from a SQLite database.
    /// </summary>
    internal sealed class SQLiteWriter : BaseWriter
    {
        #region Fields
        private readonly string _commenceFieldType = "CommenceFieldType";
        private readonly string _postFixId = "_ID";
        private readonly string _conPrefix = "_con";
        private readonly string _fkParamName = "@fk";
        private readonly string _pkParamName = "@pk";
        private readonly string _databaseName = string.Empty;
        private readonly string cs = string.Empty; // connection string
        private readonly char[] _thidDelimiter = new char[] { ':' };
        private bool disposed = false;
        // if a cursor isn't shared, we need to calculate the thid sequence differently
        // we could probably get away with just using RowID's for our purpose
        private bool _cursorShared;
        private string _fileName;
        private readonly ColumnDefinition[] originalColumnDefinitions;
        // hold the Name field for involved categories, used in filtering
        private IDictionary<string, string> categoryNamefield = new Dictionary<string, string>();
        private DataSet ds;
        #endregion

        #region Constructors
        internal SQLiteWriter(ICommenceCursor cursor, IExportSettings settings) 
            : base(cursor, settings)
        {
            // put data in %AppData%
            _databaseName = Path.GetTempFileName();
#if DEBUG
            _databaseName = @"E:\Temp\sqlite.db";
            File.Delete(_databaseName);
#endif
            
            // We could do some optimization to see if the database would fit in memory;
            // an in-memory database will obviously perform better
            // there is caveat in that any connection.Open() to an in-memory database seems to return a different database.
            // I could not figure it out
            // It does not matter much because bottleneck will always be the Commence data reading
            // except for tiny exports, which will be fast (enough) anyway.
            // Testing shows that performace penalty for using an on-disk Sqlite database
            // is negligible compared to the slow Commence reading operation.
            // Note that since it is a discardible database, we can turn off some settings and gain a little performance
            cs = @"URI=file:" + _databaseName + ";PRAGMA synchronous = OFF; JOURNAL MODE = OFF";

            // we should probably backup the original ColumnDefinitions as it has vital info
            // when it comes to custom headers and so on.
            originalColumnDefinitions = new ColumnDefinition[ColumnDefinitions.Count];
            ColumnDefinitions.CopyTo(originalColumnDefinitions, 0);
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
            // grab the Name field name of the primary category while we have an open cursor
            // there is a method for that in ICommenceDatabase but the overhead is unneeded
            using (ICommenceQueryRowSet qrs = _cursor.GetQueryRowSet(0))
            {
                // a Name field is always the first
                categoryNamefield.Add(_cursor.Category, qrs.GetColumnLabel(0, CmcOptionFlags.Fieldname));
            }

            _fileName = fileName;

            // create ADO.NET dataset to represent our data
            ds = CreateDataSetFromCursor();

            // build a SQLite database from dataset
            CreateSqliteDatabase(ds);

            // we do:
            // - a run on the primary cursor with just direct fields
            // - a run to retrieve only thids and connected thids from primary cursor
            // - a run for every connected category

            // close the original cursor
            // if it was a cursor with no connections at all, this is actually unneeded
            // but this way we can keep consistent
            // there is another benefit: max_field_size will now be better optimized
            _cursor.Close();

            // a better way of doing this may be to use the dataset all along
            // we could put both the Commence cursor parameters
            // and the INSERT commands in it as Extended Properties
            // not sure yet.
            // when a table doesn't have cursor parameters it would not be processed
            // so what would we do with are connections-only cursor that is not in the dataset?
            // and how would we get to the INSERT syntax in that case?
            foreach (var map in CursorDescriptors)
            {
                CurrentCursorDescriptor = map;
                using (var cf = new CursorFactory())
                {
                    base._cursor = cf.Create(map);
                    // we need to re-establish the columndefinitions,
                    // because the base reader needs them
                    ColumnParser cp = new ColumnParser(_cursor);
                    base.ColumnDefinitions = cp.ParseColumns();
                    _cursorShared = _cursor.Shared;
                    base.ReadCommenceData();
                }
            }
        }

        #region Data processing methods
        protected internal override void HandleProcessedDataRows(object sender, ExportProgressChangedArgs e)
        {
            // create a command that can be prepared
            using (var con = new SQLiteConnection(cs))
            {
                con.Open();
                using (var transaction = con.BeginTransaction())
                {
                    using (var cmd = new SQLiteCommand(con))
                    {

                        // I think the entry point of siphoning off control to submethods should be here
                        if (CurrentCursorDescriptor.IsLinkTable)
                        {
                            ProcessLinkTableData(cmd, e.RowValues);
                        }
                        else
                        {
                            cmd.CommandText = GetCommandTextForDirectTable();
                            ProcessDirectTableData(cmd, e.RowValues);
                        }
                    } // using cmd
                    transaction.Commit();
                } // using transaction
            } // using con
        }

        private int round = 0;
        protected internal override void HandleDataReadComplete(object sender, ExportCompleteArgs e)
        {
            base._cursor.Close(); // very important or we'll leave hanging COM references!
            // fires once for every cursor,
            // but we do not want 3 reads
            // we could simply count...
            round++;
            if (round < CursorDescriptors.Count()) { return; }

            using (var con = new SQLiteConnection(cs))
            {
                foreach (DataTable dt in ds.Tables)
                {
                    string stm = GetSelectCommandForTable(dt); // we need a function for this
                    using (var da = new SQLiteDataAdapter(stm, con))
                    {
                        da.Fill(ds, dt.TableName); // Fill will add a new column if it does not exist
                    }
                }
                // TODO figure out a way of not including the sequence 'thid' if not requested by user.
                ds.WriteXml(_fileName, XmlWriteMode.WriteSchema); // seems to work
            }

            // this is where the actual export takes place
            base.BubbleUpCompletedEvent(e);
        }

        private void ProcessDirectTableData(SQLiteCommand cmd, List<List<CommenceValue>> rowValues)
        {
            // this method only provides the parameters for the command
            for (int i = 0; i < rowValues.Count(); i++)
            {
                for (int j = 0; j < rowValues[i].Count(); j++)
                {
                    string value = rowValues[i][j].DirectFieldValue;
                    // first column is special case
                    if (j == 0) { value = SequenceFromThid(value, _cursorShared).ToString(); }
                    cmd.Parameters.AddWithValue($"@{j}", value);
                }
                cmd.Prepare();
                cmd.ExecuteNonQuery();
            }
        }

        private void ProcessLinkTableData(SQLiteCommand cmd, List<List<CommenceValue>> rowValues)
        {
            IList<string> commandTexts = GetCommandTextsForLinkedTables().ToArray();
            // this method requires that we set the CommandText on the command
            CursorParameters cp = CurrentCursorDescriptor;
            int pairedColumnIndex = 0;
            // notice that we do this column-first, it is faster
            for (int col = 0; col < rowValues[0].Count() - 1; col++) // notice the -1, because of the thid being 0
            {
                pairedColumnIndex++;
                cmd.CommandText = commandTexts[col]; // Code smell; the order and/or count may be different
                for (int row = 0; row < rowValues.Count(); row++)
                {
                    // check if there are values to process
                    if (rowValues[row][col].ConnectedFieldValues?.Count() == 0 
                        || rowValues[row][pairedColumnIndex].IsEmpty) { continue; } // all fields must contain a value, skip row
                    // cache thid of primary table
                    int pkVal = SequenceFromThid(rowValues[row][0].DirectFieldValue, _cursorShared);
                    // add parameters
                    for (int i = 0; i < rowValues[row][pairedColumnIndex].ConnectedFieldValues.Count(); i++)
                    {
                        string s = rowValues[row][pairedColumnIndex].ConnectedFieldValues[i];
                        int fkVal = SequenceFromThid(s, _cursorShared);
                        cmd.Parameters.AddWithValue(_pkParamName, pkVal);
                        cmd.Parameters.AddWithValue(_fkParamName, fkVal);
                    }
                    cmd.Prepare();
                    cmd.ExecuteNonQuery();
                }
            }
        }
        #endregion

        #region Properties
        // it may be bettor to include the descriptors as part of the dataset
        // using an extended property and serialized json
        private IList<CursorParameters> CursorDescriptors { get; } = new List<CursorParameters>();
        // allows us to keep track of what CursorDescriptor we are currently processing
        private CursorParameters CurrentCursorDescriptor { get; set; }
        #endregion

        #region Helper Methods
        private string GetSelectCommandForTable(DataTable dt)
        {
            StringBuilder sb = new StringBuilder("SELECT ");
            List<string> cols = new List<string>();
            foreach (DataColumn dc in dt.Columns)
            {
                // escape double quotes within a columnname
                string colName = dc.ColumnName.Replace("\"", "\"\"");
                // put colName itself between double quotes
                colName = $"\"{colName}\"";
                // same for alias
                string alias = dc.Caption.Replace("\"", "\"\"");
                alias = $"\"{alias}\""; // double-quote it
                
                // process special cases
                if (dc.DataType == typeof(DateTime) && dc.ExtendedProperties.Count > 0
                    && dc.ExtendedProperties[_commenceFieldType].Equals(CommenceFieldType.Date))
                {
                    colName = $"datetime({colName})";
                }
                else if (dc.DataType == typeof(DateTime) && dc.ExtendedProperties.Count > 0
                    && dc.ExtendedProperties[_commenceFieldType].Equals(CommenceFieldType.Time))
                {
                    colName = $"datetime({colName})";
                }
                colName = $"{colName} AS {alias}" ;
                cols.Add(colName);
            }
            sb.Append(string.Join(",", cols));
            sb.Append($" FROM \"{dt.TableName}\"");
            return sb.ToString();
        }

        private string GetCommandTextForDirectTable()
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
            sb.Append(string.Join(",", elements.Select(s => $"@{s}")));
            sb.Append(')');
            return sb.ToString();
        }

        private IEnumerable<string> GetCommandTextsForLinkedTables()
        {
            var linkTables = ColumnDefinitions.Where(w => w.IsConnection)
                .Select(s => new { s.Connection, s.Category })
                .Distinct()
                .ToList();
            foreach (var lt in linkTables)
            {
                // a link table always contains just 2 columns
                StringBuilder sb = new StringBuilder("INSERT INTO ");
                sb.Append($"'{_cursor.Category + lt.Connection + lt.Category}'"); // expects name of link table
                sb.Append('(');
                sb.Append($"'{_cursor.Category + _postFixId}'"); // primary key of primary category.
                sb.Append(',');
                sb.Append($"'{lt.Connection + lt.Category + _postFixId}'"); // the foreign key. This will fail for single connections
                sb.Append(") VALUES (");
                sb.Append(_pkParamName);
                sb.Append(',');
                sb.Append(_fkParamName);
                sb.Append(')');
                yield return sb.ToString();
            }
        }

        /// <summary>
        /// A pretty complex method. It translates the Commence cursor to a ADO.NET DataSet
        /// </summary>
        /// <returns></returns>
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
                dc.ExtendedProperties.Add(_commenceFieldType, df.CommenceFieldDefinition.Type);
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
            if (ColumnDefinitions.Select(s => s.Connection).Distinct(StringComparer.CurrentCultureIgnoreCase).Count() !=
                ColumnDefinitions.Select(s => s.Connection).Distinct().Count())
            {
                throw new DuplicateNameException("Commence contains case sensitive connections with same name. Exporting these is not supported.");
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
                dt = new DataTable(_conPrefix + cat); // we need a prefix or we'd get an error on single-connected categories
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
                    dc.ExtendedProperties.Add(_commenceFieldType, cd.CommenceFieldDefinition.Type);
                    dt.Columns.Add(dc);
                }
                retval.Tables.Add(dt);
                PopulateCursorDescriptors(dt, cat, false);
            }

            // create link tables
            // it is important to understand that for any cursor, 
            // is is assumed that the THID is always the first element
            var linkTables = ColumnDefinitions.Where(w => w.IsConnection)
                .Select(s => new { s.Connection, s.Category })
                .Distinct()
                .ToList();
            foreach (var lt in linkTables)
            {
                dt = new DataTable(_cursor.Category + lt.Connection + lt.Category);
                var dc = new DataColumn(_cursor.Category + _postFixId, typeof(int))
                {
                    AllowDBNull = false
                };
                dt.Columns.Add(dc);
                dc = new DataColumn(lt.Connection + lt.Category + _postFixId, typeof(int))
                {
                    AllowDBNull = false
                };
                dt.Columns.Add(dc);
                // set compound primary key
                dt.PrimaryKey = new DataColumn[2] { dt.Columns[0], dt.Columns[1] };
                retval.Tables.Add(dt);
                // add relationships to tie the linktable to the main table
                // from lnk table to primary table
                string rel1 = lt.Connection + lt.Category + _cursor.Category;
                retval.Relations.Add(rel1,
                    retval.Tables[dataSource].Columns[0],
                    retval.Tables[dt.TableName].Columns[_cursor.Category + _postFixId],
                    false).Nested = false;
                // from link table to connected table
                string rel2 = lt.Connection + lt.Category + lt.Category;
                // test for single connections (these are only possible on a category itself)
                if (!rel1.Equals(rel2))
                {
                    retval.Relations.Add(rel2,
                        retval.Tables[dt.TableName].Columns[lt.Connection + lt.Category + _postFixId],
                        retval.Tables[_conPrefix + lt.Category].Columns[0],
                        false).Nested = false;
                }
            }
            return retval;
        }

        // Notice we make it an Int32
        // This is because that way we can make SQLite much faster
        // it means that we must not store the entire thid (which is a string)
        // but only the last portion
        private void InsertThidField(DataTable dt)
        {
            DataColumn dc = new DataColumn(ColumnDefinitions[0].FieldName, typeof(int))
            {
                AllowDBNull = false
            };
            dt.Columns.Add(dc);
        }

        private void OverrideDefaultThidDataType(DataColumn dc)
        {
            if (dc.ColumnName.Equals(ColumnParser.thidIdentifier))
            {
                dc.DataType = typeof(int);
                dc.AllowDBNull = false;
            } // we will use int as datatype even tho a thid is a string
        }

        private void CreateSqliteDatabase(DataSet ds)
        {
            using (var con = new SQLiteConnection(cs))
            {
                con.Open();
                using (var transaction = con.BeginTransaction())
                {
                    foreach (DataTable dt in ds.Tables)
                    {
                        string ct = GetSqlCreateTableCommand(dt);
                        using (var cmd = new SQLiteCommand(ct, con))
                        {
                            cmd.ExecuteNonQuery();
                        } // using cmd
                    } // foreach
                    transaction.Commit();
                } // using transaction
            } // using con
        }

        private string GetSqlCreateTableCommand(DataTable dt)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("CREATE TABLE "); // note the single quote to escape weird names
            sb.Append($"'{dt.TableName}'");
            sb.Append('(');
            var columnCommands = GetSqlCreateColumnCommands(dt.Columns).ToArray();
            sb.Append(string.Join(",", columnCommands));
            sb.Append(",PRIMARY KEY(");
            sb.Append(string.Join(",", dt.PrimaryKey.Select(s => $"'{s.ColumnName}'")).ToArray());
            sb.Append(")) WITHOUT ROWID"); // we supply our own primary keys so no need for a rowid
            return sb.ToString();
        }

        private IEnumerable<string> GetSqlCreateColumnCommands(DataColumnCollection columns)
        {
            StringBuilder sb = new StringBuilder();
            foreach (DataColumn c in columns)
            {
                sb.Clear();
                sb.Append($"'{c.ColumnName}'");
                //sb.Append(' ');
                // type
                //sb.Append(c.DataType.Name); // leave the translation to Sqlite. Not sure if this will work
                // nullable
                if (!c.AllowDBNull)
                {
                    sb.Append(" INTEGER NOT NULL");
                }
                // we do not set primary key here!
                yield return sb.ToString();
            }
        }

        // a thid comes in the form of a:b:c (shared item) or a:b:c:d (local item, technically this is a rowid)
        // the last element is always the sequence field, in hexadecimal notation
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
            // would it be wise to collect some metadata here?
            // specifically, the connection counts for all involved categories?
            // it could help speed up the reading process because in some cases, filters coud be applied
            // if we would ever allow nested exports, this may get a little complicated

            CursorParameters cp;
            CmcCursorType cursorType = string.IsNullOrEmpty(_cursor.View) ? CmcCursorType.Category : CmcCursorType.View;
            // for the creation of the filters, we need the name fields of all involved categories
            // we also must count if we do not exceed 8 filters
            // if we do, do not apply filtering at all
            if (isPrimary)
            {
                // this will lose the current filter if it is on a category!!
                // The question is: do we care? Probably only when exporting a Cursor to file from CommenceCursor
                // create a cursor with all direct fields
                var directFields = dt.Columns
                    .Cast<DataColumn>()
                    .Skip(1) // skip 1 because of thid field
                    .Select(s => s.ColumnName)
                    .ToArray();

                cp = new CursorParameters(categoryName)
                {
                    Fields = directFields,
                    CursorType = cursorType,
                    MaxFieldSize = (int)Math.Pow(2, 15) // 32.768
                };
                cp.CreateSqlMapping(dt);
                CursorDescriptors.Add(cp);

                // create a second cursor with the connected fields
                // so we read less data
                // but you can do SetColumn on a cursor just once. There is no UnsetColumn
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
                            distinctConnections[i].Category + _postFixId,
                            true));
                    }

                    // capture the parameters
                    cp = new CursorParameters(categoryName)
                    {
                        Fields = conFields,
                        MaxFieldSize = CommenceLimits.MaxItems * (CommenceLimits.ThidLength + 2), // 11.500.000
                        CursorType = cursorType,
                        IsLinkTable = true,
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
                // We can apply a filter if just a single connection exists
                // between the primary category and the connected category

                // what we will do instead is take all the connections and filter them all.
                // At least that way we stand a chance of retrieving fewer items
                IList<ICursorFilterTypeCTCF> filters = GetFilters(categoryName, _cursor.Category).ToArray();

                // capture the parameters
                cp = new CursorParameters(categoryName)
                {
                    Fields = dt.Columns
                        .Cast<DataColumn>()
                        .Where(w => w.ColumnName != ColumnParser.thidIdentifier)
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
    }
}

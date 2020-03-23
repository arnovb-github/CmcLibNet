using Newtonsoft.Json;
using System.Data;
using System.Data.SQLite;
using System.IO;
using System.Linq;
using Vovin.CmcLibNet.Database.Metadata;

namespace Vovin.CmcLibNet.Export.Complex
{
    class SQLiteToJsonSerializer
    {
        private readonly string _cs;
        private readonly IExportSettings _settings;
        private readonly DataTable _primaryTable;
        private readonly OriginalCursorProperties _ocp;

        internal SQLiteToJsonSerializer(IExportSettings settings,
                OriginalCursorProperties ocp,
                DataTable primaryTable,
                string connectionString)
        {
            _settings = settings;
            _ocp = ocp;
            _cs = connectionString;
            _primaryTable = primaryTable;
        }

        internal void Serialize(string fileName)
        {
            /* We are going to perform a series of queries: 
             * One for the primary category (which we assume is the first in the dataset)
             * and then then one for every connection.
             * 
             * A number of observations:
             * If we ever were to allow nested connections, we would need some recursion algorithm.
             * In the current setup, we only process the child relations of the primary table
             */

            // collect some information before the dataread to prevent unnecessary calls
            var subQueries = _primaryTable.ChildRelations.Cast<DataRelation>().Select(s => new ChildTableQuery(
               // no checks if keys exist
               s.ChildTable.ExtendedProperties[SQLiteWriter.LinkTableSelectCommandTextExtProp].ToString(),
               JsonConvert.DeserializeObject<CommenceConnection>(s.ExtendedProperties[SQLiteWriter.CommenceConnectionDescriptionExtProp].ToString())
            )).ToArray();

            using (var connection = new SQLiteConnection(_cs))
            {
                connection.Open();
                using (var transaction = connection.BeginTransaction())
                {
                    using (var command = new SQLiteCommand(connection))
                    {
                        command.Connection = connection;
                        command.CommandText = SQLiteWriter.GetSQLiteSelectQueryForTable(_primaryTable);
                        // start reading data
                        using (var reader = command.ExecuteReader())
                        {
                            using (StreamWriter sw = new StreamWriter(fileName))
                            {
                                using (Newtonsoft.Json.JsonWriter wr = new JsonTextWriter(sw))
                                {
                                    wr.Formatting = Formatting.Indented;
                                    wr.WriteStartObject();
                                    wr.WritePropertyName("CommenceDataSource");
                                    wr.WriteValue(string.IsNullOrEmpty(_settings.CustomRootNode)
                                    ? _primaryTable.TableName
                                    : _settings.CustomRootNode);
                                    wr.WritePropertyName("CommenceCategory");
                                    wr.WriteValue(_ocp.Category);
                                    wr.WritePropertyName("CommenceDataSourceType");
                                    wr.WriteValue(_ocp.Type);
                                    wr.WritePropertyName("Items");
                                    wr.WriteStartArray();
                                    bool includeThids = ((ExportSettings)_settings).UserRequestedThids;
                                    while (reader.Read())
                                    {
                                        wr.WriteStartObject();
                                        for (int col = 0; col < reader.FieldCount; col++)
                                        {
                                            // include thid as fieldvalue in output?
                                            if (col == 0 && !includeThids) { continue; }
                                            wr.WritePropertyName(reader.GetName(col));
                                            wr.WriteValue(reader[col]);
                                        } // for
                                        SQLiteCommand sqCmd = new SQLiteCommand(connection);
                                        foreach (var sq in subQueries)
                                        {
                                            sqCmd.CommandText = sq.CommandText;
                                            sqCmd.Transaction = transaction;
                                            sqCmd.Parameters.AddWithValue("@id", reader.GetInt64(0)); // fragile
                                            var subreader = sqCmd.ExecuteReader();
                                            WriteConnectedObjects(sq, subreader, wr, includeThids);
                                            sqCmd.Reset();
                                        } // foreach subqueries
                                        wr.WriteEndObject();
                                    } // while
                                    wr.WriteEndArray();
                                    wr.WriteEndObject();
                                } // using streamwriter
                            } // using jsonwriter
                        } // using reader
                    } // using cmd
                    transaction.Commit();
                } // using transaction
            } // using con
        } // method

        private void WriteConnectedObjects(ChildTableQuery ctq, SQLiteDataReader sdr, Newtonsoft.Json.JsonWriter wr, bool includeThids)
        {
            wr.WritePropertyName(ctq.Connection.FullName);
            wr.WriteStartArray();
            while (sdr.Read())
            {
                wr.WriteStartObject();
                if (_settings.IncludeConnectionInfo)
                {
                    wr.WritePropertyName("Connection");
                    wr.WriteValue(ctq.Connection.Name);
                    wr.WritePropertyName("ToCategory");
                    wr.WriteValue(ctq.Connection.ToCategory);
                }
                for (int col = 0; col < sdr.FieldCount; col++)
                {
                    if (col == 0 && !includeThids) { continue; } // skip first row
                    string elementName = sdr.GetOriginalName(col);
                    // time and date fields have no original name
                    if (string.IsNullOrEmpty(elementName)) { elementName = sdr.GetName(col); }
                    wr.WritePropertyName(elementName);
                    wr.WriteValue(sdr[col].ToString());
                } // for
                wr.WriteEndObject();
            }
            wr.WriteEndArray();
        }

    }
}

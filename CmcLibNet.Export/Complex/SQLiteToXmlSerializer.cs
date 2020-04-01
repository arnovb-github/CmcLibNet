using Newtonsoft.Json;
using System;
using System.Data;
using System.Data.SQLite;
using System.Linq;
using System.Text;
using System.Xml;
using Vovin.CmcLibNet.Database.Metadata;

namespace Vovin.CmcLibNet.Export.Complex
{
    internal class SQLiteToXmlSerializer
    {
        private readonly string _cs;
        private readonly IExportSettings _settings;
        private readonly DataTable _primaryTable;

        internal SQLiteToXmlSerializer(IExportSettings settings,
                DataTable primaryTable,
                string connectionString) 
        {
            _settings = settings;
            _cs = connectionString;
            _primaryTable = primaryTable;
        }

        // We could use async Read from SQLite and async XML writing, but there is no performance gain, I tested that.
        // Going async actually introduces the problem of ExportComplete firing too early,
        // but you will only notice that with huge exports.
        internal void Serialize(string fileName)
        {
            /* We are going to perform a series of queries: 
             * One for the primary category
             * and then then one for every connection.
             * 
             * A number of observations:
             * If we ever were to allow nested connections, we would need some recursion algorithm.
             * In the current setup, we only process the child relations of the primary table
             */

            // collect some information before the dataread to prevent unnecessary calls
            var subQueries = _primaryTable.ChildRelations.Cast<DataRelation>().Select(s => new ChildTableQuery(
                // no checks if keys exist
                s.ChildTable.ExtendedProperties[DataSetHelper.LinkTableSelectCommandTextExtProp].ToString(),
                JsonConvert.DeserializeObject<CommenceConnection>(s.ExtendedProperties[DataSetHelper.CommenceConnectionDescriptionExtProp].ToString()))
                ).ToArray();

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
                            XmlWriterSettings xmlSettings = new XmlWriterSettings
                            {
                                Async = false,
                                WriteEndDocumentOnClose = true,
                                Indent = true,
                                Encoding = Encoding.UTF8 // this is what SQLite uses
                            };
                            using (System.Xml.XmlWriter writer = System.Xml.XmlWriter.Create(fileName, xmlSettings))
                            {
                                writer.WriteStartDocument();
                                writer.WriteStartElement(string.IsNullOrEmpty(_settings.CustomRootNode)
                                    ? XmlConvert.EncodeLocalName(_primaryTable.TableName)
                                    : XmlConvert.EncodeLocalName(_settings.CustomRootNode));
                                writer.WriteStartElement("Items");
                                bool includeThids = ((ExportSettings)_settings).UserRequestedThids;
                                while (reader.Read())
                                {
                                    writer.WriteStartElement(null, "Item", null);
                                    WriteNodes(reader, writer, includeThids);
                                    // next step is to get the connected values
                                    // we use a separate query for that
                                    // that is probably way too convoluted
                                    // we should probably stick to using a more intelligent reader.
                                    // the problem is that we need to make sense of the lines that
                                    // the reader returns. The XmlWriter is forward only,
                                    // so botching together the nodes that belong together is a problem.
                                    // we could just use a fully filled dataset? What about size limitations? Performance?
                                    SQLiteCommand sqCmd = new SQLiteCommand(connection);
                                    foreach (var q in subQueries)
                                    {
                                        sqCmd.CommandText = q.CommandText;
                                        sqCmd.Transaction = transaction;
                                        sqCmd.Parameters.AddWithValue("@id", reader.GetInt64(0)); // fragile
                                        var subreader = sqCmd.ExecuteReader();
                                        while (subreader.Read())
                                        {
                                            
                                            if (_settings.NestConnectedItems)
                                            {
                                                WriteNestedNodes(subreader, writer, q.Connection, includeThids);
                                            }
                                            else
                                            {
                                                WriteNodes(subreader, writer, includeThids);
                                            }
                                        } // while rdr.Read
                                        sqCmd.Reset(); // make ready for next use
                                    } // foreach subqueries
                                    writer.WriteEndElement();
                                } // while
                            } // xmlwriter
                        } // reader
                    } // cmd
                    transaction.Commit();
                } // transaction
            } // con
        } // method

        private void WriteNodes(SQLiteDataReader dr, System.Xml.XmlWriter writer, bool includeThids)
        {
            string value;
            for (int col = 0; col < dr.FieldCount; col++)
            {
                if (col == 0 && !includeThids) { continue; } // skip first row
                value = dr[col].ToString();
                // time and date fields have no original name
                if (string.IsNullOrEmpty(dr.GetOriginalName(col))) // either a time or a date
                {
                    value = GetShortDateOrTime(value);
                }
                // we currently do not match up the datatypes in the DataTable with th returned values from SQLite
                writer.WriteStartElement(null, XmlConvert.EncodeLocalName(dr.GetName(col)), null);
                writer.WriteString(value); // WriteString will escape most (all?) invalid chars
                writer.WriteEndElement();
            } // for
        }

        private void WriteNestedNodes(SQLiteDataReader dr, System.Xml.XmlWriter writer, CommenceConnection cc, bool includeThids)
        {
            writer.WriteStartElement(XmlConvert.EncodeLocalName(cc.Name));
            writer.WriteStartElement(XmlConvert.EncodeLocalName(cc.ToCategory));
            WriteNodes(dr, writer, includeThids);
            writer.WriteEndElement();
            writer.WriteEndElement();
        }

        private string GetShortDateOrTime(string value)
        {
            return value.ToString().Contains(':')
                ? Convert.ToDateTime(value).ToShortTimeString()
                : Convert.ToDateTime(value).ToShortDateString();
        }
    }
}
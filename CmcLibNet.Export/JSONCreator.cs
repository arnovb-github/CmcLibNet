using Newtonsoft.Json.Linq;
using System.Collections.Generic;

namespace Vovin.CmcLibNet.Export
{
    /// <summary>
    /// Collect Commence data and convert them into a JSON object.
    /// Allows for consistent Json output format.
    /// Used by export methods that return Json.
    /// </summary>
    internal class JsonCreator
    {
        private readonly BaseWriter _wr = null;

        internal JsonCreator(BaseWriter writer)
        {
            _wr = writer;
            this.Category = _wr._cursor.Category;
            this.DataSource = string.IsNullOrEmpty(_wr._settings.CustomRootNode) ? _wr._dataSourceName : _wr._settings.CustomRootNode;
            this.DataSourceType = string.IsNullOrEmpty(_wr._cursor.View) ? CmcLibNet.Database.CmcCursorType.Category.ToString() : CmcLibNet.Database.CmcCursorType.View.ToString();
        }

        #region Properties

        private string Category { get; set; }
        private string DataSource { get; set; }
        private string DataSourceType { get; set; }
        internal JObject MetaData
        {
            get
            {
                string type = string.IsNullOrEmpty(_wr._cursor.View) ? CmcLibNet.Database.CmcCursorType.Category.ToString() : CmcLibNet.Database.CmcCursorType.View.ToString();
                return new JObject(
                             new JProperty("CommenceDataSource", DataSource),
                             new JProperty("CommenceCategory", Category),
                             new JProperty("CommenceDataSourceType", type)
                         );
            }
        }

        #endregion

        #region Helper methods
        /// <summary>
        /// Takes a batch of Commence rowvalues and turns it into Json
        /// </summary>
        /// <param name="rowvalues">List of list of Commence rowvalues</param>
        /// <returns>List of JObject</returns>
        internal List<JObject> SerializeRowValues(List<List<CommenceValue>> rowvalues)
        {
            List<JObject> retval = new List<JObject>();
            // Take the rowvalues and process them into a JSON object.
            foreach (List<CommenceValue> lrv in rowvalues) // process rows
            {
                var row = new JObject();
                foreach (CommenceValue v in lrv) // process columns
                {
                    if (v.IsEmpty) { continue; } // nothing to do, process next column

                    if (v.ColumnDefinition.IsConnection) // we have a connection
                    {
                        if (!_wr._settings.SkipConnectedItems)
                        {
                            string nodeName = _wr.ExportHeaders[v.ColumnDefinition.ColumnIndex];
                            var citems = new JArray();
                            foreach (string value in v.ConnectedFieldValues)
                            {
                                if (_wr._settings.IncludeConnectionInfo)
                                {
                                    dynamic citem = new JObject(); // create object for value plus additional details on the connection
                                    citem.Connection = v.ColumnDefinition.Connection;
                                    citem.Category = v.ColumnDefinition.Category;
                                    citem[v.ColumnDefinition.FieldName] = value;
                                    citems.Add(citem);
                                }
                                else
                                {
                                    citems.Add(value); // just add the value to array
                                } //if (_wr._settings.IncludeConnectionInfo)
                            } // foreach
                            row[nodeName] = citems;
                        } // if (!this._settings.SkipConnectedFields)
                    }
                    else // we have a direct field
                    {
                        row.Add(_wr.ExportHeaders[v.ColumnDefinition.ColumnIndex], v.DirectFieldValue);
                    } // if v.IsConnection
                } // foreach CommenceValue
                retval.Add(row); // add row
            } // foreach List<CommenceValue>
            return retval;
        }
        #endregion
    }
}
using System.Collections.Generic;
using Newtonsoft.Json.Linq;

namespace Vovin.CmcLibNet.Export
{
    /// <summary>
    /// Collect Commence data and convert them into a JSON object.
    /// Allows for consistent JSON output format.
    /// Used by export methods that return JSON.
    /// </summary>
    internal class JSONCreator
    {
        /* Using JObject and JArray may seem overly complex,
        * but it is needed because the input data is of type List<List<CommenceValue>>, 
        * not some POCO class.
        * You cannot serialize that to a meaningful object easily.
        * So, we create a JSON object 'by hand'.
        * Another approach could be to create a dynamic object class, then serialize that.
         * However, since the JObject is itself already a dynamic object especially suitable for JSON.Net,
         * we're not going to bother.
        */

        private JArray _rows = null;
        private BaseWriter _wr = null;

        internal JSONCreator(BaseWriter writer)
        {
            _wr = writer;
            this.Category = _wr._cursor.Category;
            this.DataSource = _wr._dataSourceName;
            this.DataSourceType = string.IsNullOrEmpty(_wr._cursor.View) ? CmcLibNet.Database.CmcCursorType.Category.ToString() : CmcLibNet.Database.CmcCursorType.View.ToString();
            _rows = new JArray();
        }

        #region Properties

        private string Category { get; set; }
        private string DataSource { get; set; }
        private string DataSourceType { get; set; }

        #endregion

        #region Methods

        internal void AppendRowValues(List<List<CommenceValue>> rowvalues)
        {
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
                _rows.Add(row); // add row
            } // foreach List<CommenceValue>
        }

        // add metadata headers and return populated JSON object
        internal JObject ToJObject()
        {
            JObject j = new JObject
            {
                { "CommenceDataSource", this.DataSource },
                { "CommenceDataSourceType", this.DataSourceType },
                { "CommenceCategory", this.Category },
                { "Items", this._rows }
            };
            return j;
        }
        #endregion
    }
}
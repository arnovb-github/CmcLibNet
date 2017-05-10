using System;
using System.IO;
using Newtonsoft.Json.Linq;
using System.Collections.Generic;

namespace Vovin.CmcLibNet.Export
{
    /// <summary>
    /// Write export data to JSON file.
    /// </summary>
    internal class JSONDataWriter : BaseDataWriter
    {
        internal JSONDataWriter(Database.ICommenceCursor cursor, IExportSettings settings)
            : base(cursor, settings) { }
        //internal JSONDataWriter(string dataName, Database.CMC_GETCURSOR_FLAGS cursorType, IExportSettings settings) : base(dataName, cursorType, settings) { }
        /// <summary>
        /// Write data to file in JSON format.
        /// </summary>
        /// <param name="fileName">(fully qualified) filename. Overwrites existing.</param>
        internal override void WriteOut(string fileName)
        {
            // the easiest way to write JSON is probably by using dynamic objects
            // they do not complain (too much) about property names, which in our case are unpredictable
            using (StreamWriter sw = new StreamWriter(fileName))
            {
                // note the use of dynamic objects
                var json = new JObject() as dynamic;
                json.Add("dataSource", base._dataSourceName);
                json.Add("Category", base._cursor.Category);
                var items = new JArray();
                // loop cursor and get data, maxRows rows at a time
                for (int rows = 0; rows < _cursor.RowCount; rows += MAX_ROWS)
                {
                    List<List<CommenceValue>> values = base.GetDataAsList(MAX_ROWS);
                    foreach (List<CommenceValue> rowvalues in values)
                    {
                        var item = new JObject();
                        foreach (CommenceValue v in rowvalues)
                        {
                            if (v.ColumnDefinition.IsConnection)
                            {
                                if (!base._settings.SkipConnectedItems)
                                {
                                    string nodeName = base.ExportHeaders[v.ColumnDefinition.ColumnIndex].Item1;
                                    var citems = new JArray(); 
                                    foreach (string s in v.ConnectedFieldValues)
                                    {
                                        dynamic citem = new JObject();
                                        //string cnode = v.ColumnDefinition.FieldName; // does not contain same amount, so cannot use j!
                                        citem.Connection = v.ColumnDefinition.Connection;
                                        citem.Category = v.ColumnDefinition.Category;
                                        citem[v.ColumnDefinition.FieldName] = s; // should we use the fieldname instead of 'fieldvalue'?
                                        citems.Add(citem);
                                    }
                                    item[nodeName] = citems;
                                } // if (!base._settings.SkipConnectedFields)
                            } // if (base.ExportHeaders[j].Item2 == LABEL_TYPE.LABEL_TYPE_CONNECTION)
                            else
                            {
                                item.Add(base.ExportHeaders[v.ColumnDefinition.ColumnIndex].Item1, v.DirectFieldValue);
                            }
                        } // j
                        items.Add(item);
                        this.CurrentRow++;
                    } // i
                    base.RaiseProgressChangedEvent(this.CurrentRow); // notify subscribers of our progress
                } // rows
                json.Add("items", items);
                sw.WriteLine(json.ToString());
                base.RaiseProgressChangedEvent(_cursor.RowCount); // we're done!
            }
        }

        protected internal override int CurrentRow { get; set; }

    }
}

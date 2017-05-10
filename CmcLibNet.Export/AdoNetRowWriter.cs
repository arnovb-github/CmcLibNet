using System;
using System.Collections.Generic;
using System.Linq;
using System.Data;
using Vovin.CmcLibNet;
using Vovin.CmcLibNet.Database;

namespace Vovin.CmcLibNet.Export
{
    /// <summary>
    /// Takes a complete cursor row of Commence data and puts the values into their corresponding DataSet tables.
    /// Assumes values passed in are in canonical format; this requires converting then into a format that ADO.NET understands.
    /// </summary>
    internal class AdoNetRowWriter
    {
        #region Fields
        private List<CommenceValue> _vpList;
        private List<string> _tableNames = new List<string>();
        private DataSet _ds = null;
        private int _pk = 0;
        #endregion

        #region Constructors
        internal AdoNetRowWriter(int primaryKey, List<CommenceValue> CommenceColumnValueList, DataSet dataset)
        {
            _vpList = CommenceColumnValueList;
            _ds = dataset;
            _pk = primaryKey;
        }
        #endregion

        #region Methods
        internal void ProcessRow()
        {
            // magic happens here

            // process direct fields
            // A cursor can contain only related columns!
            // in that case all columns are IsConnection
            // There will be a primary table, but it will only contain id's, not name field values.
            DataTable dt = null;
            dt = _ds.Tables[0]; // assume primary table is always first table!
            DataRow dr = dt.NewRow();
            dr["id"] = _pk;
            foreach (CommenceValue cv in _vpList)
            {
                if (!cv.ColumnDefinition.IsConnection) // we have a direct field
                {
                    string value = cv.DirectFieldValue;
                    try
                    {
                        dr[cv.ColumnDefinition.FieldName] = (String.IsNullOrEmpty(cv.DirectFieldValue)) ? DBNull.Value : (object)CommenceValueConverter.ToAdoNet(value, cv.ColumnDefinition.FieldType);
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e.Message);
                    }
                }
            }
            dt.Rows.Add(dr);

            // process related columns
            // this is way more tricky...
            // we need to open datatables based on columns, but columns can use the same data table
            List<string> connectedDataTableNames = new List<string>();
            foreach (CommenceValue cv in _vpList)  // loop all row elements
            {
                if (cv.ColumnDefinition.IsConnection)
                {
                    // compile a list of related table names
                    connectedDataTableNames.Add(cv.ColumnDefinition.AdoTableName);
                }
            }
            // connectedDataTableNames now contains all table names for connected items
            // these can be duplicate, so make them unique
            connectedDataTableNames = connectedDataTableNames.Distinct().ToList();
            // now we have a list of unique table names
            // we now need to create a row in each table with the values Commence gave us for that particular connection
            foreach (string connectedTableName in connectedDataTableNames)
            {
                // Create a second list of connected value/column pairs,
                // then pass that on to ParseConnectedRows method with the associated DataTable
                List<CommenceValue> cvpList = new List<CommenceValue>(); // new list for just the connected values
                foreach (CommenceValue cv in _vpList)
                {
                    if (cv.ColumnDefinition.AdoTableName.Equals(connectedTableName))
                    {
                        cvpList.Add(cv);
                    }
                }
                dt = _ds.Tables[connectedTableName];
                // now pass on the list of connected values
                this.ProcessConnectedValuesForItem(cvpList, dt, _pk);
            }
        }

        private void ProcessConnectedValuesForItem(List<CommenceValue> values, DataTable dt, int fk)
        {
            // values contains arrays of connected values
            // we need a new row for every array element
            // all arrays will have same number of elements *if populated*
            // ideally we should compare the number of elements against the number of actual connected items
            // but I do not see a reliable way to do that (not quickly, anyway).

            // create a number of new rows equal to the number of connected items
            int numrows = 0;
            DataRow[] newrowbuffer = null;
            // find the number of connected items
            // this assumes the number is correct, which in case of a large text field may not be true.
            IEnumerable<CommenceValue> x = values.Where(o => o.ConnectedFieldValues != null);
            if (x.Count() == 0) // x is the collection CommenceValue objects that have connected values
            {
                return; // there are no connected values
            }
            else
            {
                // get the highest number of connected items.
                // They *should* all be the same, except when they were split incorrectly, in which case we end up with more rows than needed.
                numrows = x.Select(o => o.ConnectedFieldValues.Length).ToArray<int>().Max();
            }

            for (int i = 0; i < numrows; i++)
            {
                if (newrowbuffer == null) { newrowbuffer = new DataRow[numrows]; }
                newrowbuffer[i] = dt.NewRow();
                newrowbuffer[i]["fkid"] = fk;
            }

            foreach (CommenceValue v in values)
            {
                if (v.ConnectedFieldValues != null) // check to see if we have connected values
                {
                    for (int i = 0; i < v.ConnectedFieldValues.Length; i++) // process connected items
                    {
                        string value = v.ConnectedFieldValues[i];
                        // note the importance of entering DBNull value to keep connected item integrity.
                        try
                        {
                            newrowbuffer[i][v.ColumnDefinition.FieldName] = (string.IsNullOrEmpty(value)) ? DBNull.Value : (object)CommenceValueConverter.ToAdoNet(value, v.ColumnDefinition.FieldType); //cast to object is needed to make DBNull.Value and the string value equatable.
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine(e.Message);
                        }
                    } // foreach
                } // if
            } // foreach

            // commit rows to datatable
            foreach (DataRow dr in newrowbuffer)
            {
                dt.Rows.Add(dr);
            }
        } // method
        #endregion
    }
}

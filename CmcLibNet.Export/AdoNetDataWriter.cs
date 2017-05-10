using System;
using System.Collections.Generic;
using System.Linq;
using System.Data;
using Vovin.CmcLibNet.Database;

namespace Vovin.CmcLibNet.Export
{
    internal class AdoNetDataWriter : BaseDataWriter
    {
        /* Builds an in-memory database based on the cursor
         * -Analyze the cursor, collecting tables and fields
         * -Define the dataset
         * -Read in data
         * 
         * We can export this dataset in a variety of formats.
         */
        DataSet _ds = null;

        #region Constructors
        internal AdoNetDataWriter(Database.ICommenceCursor cursor, IExportSettings settings)
            : base(cursor, settings) 
        {
        }
        //internal AdoNetDataWriter(string dataName, Database.CMC_GETCURSOR_FLAGS cursorType, IExportSettings settings)
        //    : base(dataName, cursorType, settings)
        //{
        //}
        #endregion

        #region Methods

        internal override void WriteOut(string fileName = null) // we don't write to file, but our base class demands it.
        {
            //throw new NotImplementedException();
            CreateDataSetFromCursorColumns(); // TODO move to base
            for (int rows = 0; rows < _cursor.RowCount; rows += MAX_ROWS) // items
            {
                List<List<CommenceValue>> values = base.GetDataAsList(MAX_ROWS); 
                foreach (List<CommenceValue> row in values)
                {
                    // pass on rowdata for RowParser
                    RowParser rp = new RowParser(this.CurrentRow, row, _ds);
                    rp.ParseRow();
                    this.CurrentRow++;
                } // foreach
            } // for
        } // WriteOut

        private void CreateDataSetFromCursorColumns()
        {
            List<TableDef> tables = base.GetTableInfoFromCursorColumns(base.ColumnInfo.ColumnDefinitions);
            // we have now collected all table info,
            // we can create a dataset
            DataTable dt = null;
            DataTable primaryTable = null; //hack

            for (int i = 0; i< tables.Count; i++)
            {
                if (_ds == null) { _ds = new DataSet(base._dataSourceName); }
                
                TableDef td = tables[i];

                try
                {
                    dt = _ds.Tables.Add(td.Name);
                }
                catch (DuplicateNameException)
                {
                    // TODO deal with this
                    // It is a rare exception, but possible. Commence connection names are case-sensitive.
                    // In that case this exception is not thrown, but data can still end up in the wrong datatable (or more likely fail).

                    //dt = ds.Tables.Add(td.AdoTableName + i.ToString());
                }

                // define columns
                DataColumn dc = null;
                
                // we need to set some general fields
                if (td.Primary) // primary table columns
                {
                    dc = dt.Columns.Add("id", typeof(Int32)); // not autoincremented!
                    dc.AllowDBNull = false;
                    // hack for setting relations later on. We could also have used ExtendedProperties.
                    primaryTable = dt; // capture that this is our primary table so we can set relationships later on
                    dt.ExtendedProperties.Add("IsPrimaryTable", true);
                }
                else // related table columns
                {
                    // create a primary key that autoincrements
                    dc = dt.Columns.Add("id", typeof(Int32));
                    dc.AllowDBNull = false;
                    dc.AutoIncrement = true;
                    // create a foreign key field to reference the connected item
                    dc = dt.Columns.Add("fkid", typeof(Int32));
                    dc.AllowDBNull = false; // cannot be null
                }

                // now process Commence fields
                // rest of related columns
                for (int j = 0; j < td.ColumnDefinitions.Count; j++)
                {
                    try
                    {
                        dc = dt.Columns.Add(td.ColumnDefinitions[j].FieldName);
                    }
                    catch (DuplicateNameException)
                    {
                        // appending some number would decouple the fieldname in the dataset
                        // from the fieldname in the Commence columnset.
                    }
                    // set column properties
                    dc.DataType = Core.Utils.GetSystemTypeForCommenceField(td.ColumnDefinitions[j].FieldType);
                    dc.AllowDBNull = true; // this is default, but setting it explicitly makes it more clear.
                }
            }
            // we would like to set relationships as well.
            // as it is, we don't know what our 'primary' table is at this point in the code
            // all we know it is the only table without a fkid field, but that's a little awkward to check.
            // we hack that by creating a 'mock' primary table
            List<DataRelation> relations = null;
            foreach (DataTable d in _ds.Tables)
            {
                if (!d.Equals(primaryTable))
                {
                    if (relations == null) { relations = new List<DataRelation>(); }
                    DataRelation r = new DataRelation("rel" + d.TableName,
                        primaryTable.TableName,
                        d.TableName,
                        new string[] { "id" },
                        new string[] { "fkid" },
                        true); // setting nested to true allows for nested XML exports
                    relations.Add(r);
                }
            }
            // add the relations to the dataset
            // nasty snag found on https://msdn.microsoft.com/en-us/library/2z22c2sz.aspx
            // "Any DataRelation object created by using this constructor must be added to the collection
            // with the AddRange method inside of a BeginInit and EndInit block.
            if (relations != null)
            {
                _ds.BeginInit();
                _ds.Relations.AddRange(relations.ToArray());
                _ds.EndInit();
            }
        }
        #endregion

        #region Properties
        protected internal override int CurrentRow { get; set; }
        internal DataSet DataSet
        {
            get
            {
                return _ds;
            }
        }
        #endregion

        #region Helper Classes
        /// <summary>
        /// Takes a complete cursor row of Commence data and puts the values into their proper DataSet tables.
        /// </summary>
        private class RowParser
        {
            private List<CommenceValue> vpList;
            private List<string> tableNames = new List<string>();
            private DataSet ds = null;
            private int pk = 0;

            internal RowParser(int primaryKey, List<CommenceValue> CommenceColumnValueList, DataSet dataset)
            {
                vpList = CommenceColumnValueList;
                ds = dataset;
                pk = primaryKey;
            }

            internal void ParseRow()
            {
                // magic happens here

                // process direct fields
                // A cursor can contain only related columns!
                // in that case all columns are IsConnection
                // There will be a primary table, but it will only contain id's, not name field values.
                DataTable dt = null;
                dt = ds.Tables[0]; // assume primary table is always first table!
                DataRow dr = dt.NewRow();
                dr["id"] = pk;
                foreach (CommenceValue cv in vpList)
                {
                    if (!cv.ColumnDefinition.IsConnection) // we have a direct field
                    {
                        string value = cv.DirectFieldValue;
                        if (cv.ColumnDefinition.FieldType == CommenceFieldTypes.Number || cv.ColumnDefinition.FieldType == CommenceFieldTypes.Calculation)
                        {
                            value = Vovin.CmcLibNet.Core.Utils.RemoveCurrencySymbol(cv.DirectFieldValue);
                        }
                        dr[cv.ColumnDefinition.FieldName] = (String.IsNullOrEmpty(cv.DirectFieldValue)) ? DBNull.Value : (object)value;
                    }
                }
                dt.Rows.Add(dr);

                // process related columns
                // this is way more tricky...
                // we need to open datatables based on columns, but columns can use the same data table
                List<string> connectedDataTableNames = new List<string>();
                foreach (CommenceValue cv in vpList)  // = loop thru all row elements
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
                    foreach (CommenceValue cv in vpList)
                    {
                        if (cv.ColumnDefinition.AdoTableName.Equals(connectedTableName))
                        {
                            cvpList.Add(cv);
                        }
                    }
                    dt = ds.Tables[connectedTableName];
                    // now pass on the list of connected values
                    this.ParseConnectedRows(cvpList, dt, pk);
                }
            }

            private void ParseConnectedRows(List<CommenceValue> values, DataTable dt, int fk)
            {
                // values contains arrays of connected values
                // we need a new row for every array element
                // all arrays will have same number of elements *if populated*
                // ideally we should compare the number of elements against the number of actual connected items
                // but I do not see a reliable way to do that (not quickly, anyway).
                DataRow dr = dt.NewRow(); // prepare new row
                dr["fkid"] = fk;
                bool commit = false; // if no values are present, there is no need to commit the row
                foreach (CommenceValue v in values)
                {
                    commit = false;
                    if (v.ConnectedFieldValues != null)
                    {
                        commit = true;
                        foreach (string s in v.ConnectedFieldValues)
                        {
                            string value = s;
                            if (v.ColumnDefinition.FieldType == CommenceFieldTypes.Number || v.ColumnDefinition.FieldType == CommenceFieldTypes.Calculation)
                            {
                                value = Vovin.CmcLibNet.Core.Utils.RemoveCurrencySymbol(s);
                            }
                            dr[v.ColumnDefinition.FieldName] = (String.IsNullOrEmpty(value)) ? DBNull.Value : (object)value; //cast to object is needed to make DBNull.Value and the string value equatable.
                            // TODO note there is a potential error here in that large text fields that contain embedded delimiters may get split into separate rows
                        } // foreach
                        
                    } // if
                } // foreach
                if (commit) { dt.Rows.Add(dr); }
            } // ParseConnectedRows method

            //private void ParseConnectedRows2(List<CommenceValue> values, DataTable dt, int fk) // OBSOLETE
            //{
            //    // values contains arrays of connected values
            //    // we need a new row for every array element
            //    // all arrays will have same number of elements
                
            //    //int i = values[0].CommenceValue.Length; // returns string length, not element number
            //    int numItems = values[0].ConnectedFieldValues.Length; // all arrays will have same number of elements
            //    //Console.WriteLine(numItems.ToString() + " elements in connected category " + values[0].ColumnDefinition.Category);
            //    //Console.WriteLine("values are: " + values[0].CommenceValue);
            //    for (int i = 0; i < numItems; i++)
            //    {
            //        DataRow dr = dt.NewRow();
            //        dr["fkid"] = fk;
            //        foreach (CommenceValue v in values)
            //        {
            //            string[] fieldValues = v.ConnectedFieldValues;
            //            if (fieldValues.Length == numItems) // check if we still have the correct number of elements
            //            {
            //                string value = fieldValues[i];
            //                if (v.ColumnDefinition.FieldType == CommenceFieldTypes.Number || v.ColumnDefinition.FieldType == CommenceFieldTypes.Calculation)
            //                {
            //                    value = RemoveCurrencySymbol(fieldValues[i]);
            //                }
            //                dr[v.ColumnDefinition.FieldName] = (String.IsNullOrEmpty(fieldValues[i])) ? DBNull.Value : (object)value; //cast to object is needed to make DBNull.Value and the string value equatable.
            //            }
            //            else
            //            {
            //                // assume that v.ColumnDefinition.FieldName is a string field!
            //                try
            //                {
            //                    dr[v.ColumnDefinition.FieldName] = "#ERROR# Fieldvalue contains embedded line-feed character."; // there was an error, most likely due to embedded \n delimiter
            //                }
            //                catch { }
            //            }
            //        }
            //        dt.Rows.Add(dr);
            //    }
            //}
        } // RowParser class
        #endregion
    } // class
} // namespace

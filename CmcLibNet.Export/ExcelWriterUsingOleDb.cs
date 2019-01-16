using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using Vovin.CmcLibNet.Database;
using Vovin.CmcLibNet.Extensions;

namespace Vovin.CmcLibNet.Export
{
    internal class ExcelWriterUsingOleDb : BaseWriter
    {
        private bool disposed = false;
        private string _fileName;
        private readonly int MaxExcelCellSize = (int)Math.Pow(2, 15) - 1; // as per Microsoft documentation
        private readonly int MaxExcelNewLines = 253; // as per Microsoft documentation
        private OleDbConnections Connection = null;
        private OleDbConnection cn = null;
        private OleDbCommand cmd = null;
        private Dictionary<string, string> columnMap = null;
        private string sheetName;
        private string insertCommandText;
        private List<ColumnDefinition> columnDefinitions = null;

        internal ExcelWriterUsingOleDb(ICommenceCursor cursor, IExportSettings settings) : base(cursor, settings)
        {
            columnDefinitions = new List<ColumnDefinition>(_settings.UseThids ? base.ColumnDefinitions.Skip(1) : base.ColumnDefinitions);
        }

        ~ExcelWriterUsingOleDb()
        {
            Dispose(false);
        }

        protected internal override void HandleDataReadComplete(object sender, ExportCompleteArgs e)
        {
            cn.Close();
            base.BubbleUpCompletedEvent(e);
        }

        protected internal override void HandleProcessedDataRows(object sender, ExportProgressChangedArgs e)
        {
            foreach (List<CommenceValue> row in e.RowValues)
            {
                InsertValues(row);
            }
            BubbleUpProgressEvent(e);
        }

        /// <summary>
        /// Starts the export
        /// </summary>
        /// <param name="fileName">Fully qualified filename.</param>
        /// <exception cref="System.IO.IOException"></exception>
        protected internal override void WriteOut(string fileName)
        {
            _fileName = fileName;
            if (_settings.DeleteExcelFileBeforeExport)
            {
                File.Delete(fileName);
            }
            // override any user settings
            // let the OleDb driver do it's magic
            base._settings.Canonical = false;
            base._settings.XSDCompliant = false;

            sheetName = Utils.GetOleDbFieldName(_dataSourceName, '_');
            Dictionary<string, string> columnMap = ColumnMap;
            Connection = new OleDbConnections();
            cn = new OleDbConnection { ConnectionString = Connection.HeaderConnectionString(_fileName, 0) };
            CreateOleDbTable(cn, sheetName);
            base.ReadCommenceData();
        }

        private void CreateOleDbTable(OleDbConnection cn, string tableName)
        {
            cmd = new OleDbCommand { Connection = cn };

            cn.Open();
            if (Connection.SheetExists(cn, tableName))
            {
                cn.Close();
                return;
            }
            cmd.CommandText = GetCommandTextForCreatingTable(sheetName);
            cmd.ExecuteNonQuery();
            List<OleDbParameter> oleDbParams = new List<OleDbParameter>();
            if (_settings.UseThids)
            {
                oleDbParams.Add(new OleDbParameter
                {
                    ParameterName = "@" + columnMap["thid"],
                    DbType = DbType.String
                });
            }
            foreach (var cd in this.columnDefinitions)
            {
                OleDbParameter p = new OleDbParameter();
                if (cd.IsConnection)
                {
                    p.ParameterName = "@" + columnMap[cd.ColumnName];
                    p.DbType = DbType.String;
                    p.OleDbType = OleDbType.LongVarWChar;
                }
                else
                {
                    p.ParameterName = "@" + columnMap[cd.ColumnName];
                    p.DbType = Utils.GetDbTypeForCommenceField(cd.FieldType);
                    p.OleDbType = Utils.GetOleDbTypeForCommenceField(cd.FieldType);
                }
                oleDbParams.Add(p);
            }
            cmd.Parameters.AddRange(oleDbParams.ToArray());
        }

        private string GetCommandTextForCreatingTable(string tableName)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("CREATE TABLE " + tableName + "(");

            IList<string> fielddefs = new List<string>();
            if (_settings.UseThids)
            {
                fielddefs.Add(ColumnMap["thid"] + " CHAR (50)");
            }
            foreach (var cd in this.columnDefinitions)
            {
                StringBuilder sb2 = new StringBuilder();
                sb2.Append(ColumnMap[cd.ColumnName]);
                sb2.Append(" ");

                if (cd.IsConnection)
                {
                    sb2.Append("MEMO");
                }
                else
                {
                    sb2.Append(Utils.GetOleDbTypeStringForCommenceField(cd.FieldType).ToUpper());
                }

                fielddefs.Add(sb2.ToString());
            }
            sb.Append(string.Join(",", fielddefs));
            sb.Append(")");

            return sb.ToString();
        }

        private int InsertValues(IList<CommenceValue> values)
        {
            foreach (var v in values)
            {
                string paramIdentifier = "@" + columnMap[v.ColumnDefinition.ColumnName];
                if (!v.ColumnDefinition.IsConnection)
                {
                    string value = string.Empty;
                    // remove currency symbol if present or we'll get a type cast error
                    if (v.ColumnDefinition.FieldType == CommenceFieldType.Number)
                    {
                        value = CommenceValueConverter.RemoveCurrencySymbol(v.DirectFieldValue);
                    }
                    else
                    {
                        value = v.DirectFieldValue;
                    }
                    // if we are dealing with large text fields, we may have more than allowed number of newlines
                    if (value.CountChar('\n') > MaxExcelNewLines)
                    {
                        value = value.Substring(0, value.IndexOfNthChar('\n', 0, MaxExcelNewLines));
                    }
                    cmd.Parameters[paramIdentifier].Value = value;
                }
                else
                {
                    string cValue = string.Empty;
                    // there can be many more values in a connection than an Excel cell can hold
                    if (v.ConnectedFieldValues != null)
                    {
                        // check maximum number of newlines
                        if (_settings.TextDelimiterConnections == "\n")
                        {
                            cValue = string.Join(_settings.TextDelimiterConnections, v.ConnectedFieldValues.Take(MaxExcelNewLines));
                        }
                        else
                        {
                            cValue = string.Join(_settings.TextDelimiterConnections, v.ConnectedFieldValues);
                        }
                        // if we are dealing with large text fields, we still may have more than allowed number of newlines
                        if (cValue.CountChar('\n') > MaxExcelNewLines)
                        {
                            cValue = cValue.Substring(0, cValue.IndexOfNthChar('\n', 0, MaxExcelNewLines));
                        }

                        // make sure length doesn't exceed Excel cell limit
                        cValue = cValue.Length > MaxExcelCellSize ? cValue.Substring(0, MaxExcelCellSize) : cValue;
                    }
                    int lengte = cValue.Length;
                    cmd.Parameters[paramIdentifier].Value = cValue;
                }
            }
            cmd.CommandText = InsertCommandText;
            return cmd.ExecuteNonQuery();
        }

        private string MakeListOfUniqueStrings(List<string> list, string value, int append = 0)
        {
            string retval = value;
            if (append > 255) { throw new ArgumentOutOfRangeException("Could not create a unique columname for OleDb."); }
            if (list.Contains(retval))
            {
                retval = retval.Substring(0, retval.Length - append.ToString().Length);
                retval = retval + append.ToString();
                append++;
                retval = MakeListOfUniqueStrings(list, retval, append); // recurse
            }
            return retval;
        }
        #region Dispose
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
                if (cn != null) { cn.Close(); }
            }

            // Free any unmanaged objects here.
            //
            disposed = true;

            // Call the base class implementation.
            base.Dispose(disposing);
        }
        #endregion

        #region Properties
        private Dictionary<string, string> ColumnMap
        {
            get
            {
                if (columnMap == null)
                {
                    columnMap = new Dictionary<string, string>();
                    List<string> oleDbFields = new List<string>();
                    foreach (var cd in base.ColumnDefinitions)
                    {
                        // ensure we have unique values or oledb may fail
                        string f = Utils.GetOleDbFieldName(cd.ColumnName, '_');
                        f = MakeListOfUniqueStrings(oleDbFields, f);
                        oleDbFields.Add(f);
                    }

                    for (int i = 0; i < base.ColumnDefinitions.Count() ;i++)
                    {
                       columnMap.Add(base.ColumnDefinitions[i].ColumnName, oleDbFields[i]);
                    }
                }
                return columnMap;
            }
        }

        private string InsertCommandText
        {
            get
            {
                if (string.IsNullOrEmpty(insertCommandText))
                {
                    StringBuilder sb = new StringBuilder();
                    sb.Append("INSERT INTO ");
                    sb.Append(sheetName);
                    sb.Append("(");
                    sb.Append(string.Join(",", columnMap.Values));
                    sb.Append(")");
                    sb.Append("VALUES (");
                    sb.Append(string.Join(",", columnMap.Values.Select(s => "@" + s)));
                    sb.Append(")");
                    insertCommandText = sb.ToString();
                }
                return insertCommandText;
            }
        }

        #endregion
    }
}

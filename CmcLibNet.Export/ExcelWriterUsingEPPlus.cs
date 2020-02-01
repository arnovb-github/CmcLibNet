using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using Vovin.CmcLibNet.Database;
using Vovin.CmcLibNet.Extensions;

namespace Vovin.CmcLibNet.Export
{
    internal class ExcelWriterUsingEPPlus : BaseWriter
    {
        private string _sheetName = string.Empty;
        private readonly int MaxSheetNameLength = 31;// as per Microsoft documentation
        private readonly int MaxExcelCellSize = (int)Math.Pow(2, 15) - 1; // as per Microsoft documentation
        private readonly List<ColumnDefinition> columnDefinitions = null;
        private readonly DataTable _dataTable = null;
        private FileInfo _fi = null;

        internal ExcelWriterUsingEPPlus(ICommenceCursor cursor, IExportSettings settings) 
            : base(cursor, settings)
        {
            columnDefinitions = new List<ColumnDefinition>(_settings.UseThids ? base.ColumnDefinitions.Skip(1) : base.ColumnDefinitions);
            string s = string.IsNullOrEmpty(settings.CustomRootNode) ? _dataSourceName : settings.CustomRootNode;
            _dataTable = PrepareDataTable(s, columnDefinitions);
            _sheetName = Utils.EscapeString(s, "_").Left(MaxSheetNameLength);
            settings.Canonical = true; // override custom setting
            settings.SplitConnectedItems = false; // override custom setting
            // When dealing with very large cursors,
            // the default number of rows read may lead to memory issues when dumping to a datatable
            // we may have to cap the number of items being read
            // This superceeds the check already performed in the datareader.
            // I am undecided on this.
            // What I do know is that writing 1000 rows of 250 columns of size 30.000
            // (i.e. 250 large text fields, fully populated) will fail.
            // Maybe we should have some mechanism that collects the size of the fields
            // then caps the NumRows count accordingly.
            // However, it would not solve issues with EPPlus running out of memory when dealing with huge workbooks.
        }

        protected internal override void HandleDataReadComplete(object sender, ExportCompleteArgs e) 
        {
            base.BubbleUpCompletedEvent(e);
        }

        protected internal override void HandleProcessedDataRows(object sender, ExportProgressChangedArgs e)
        {
            // at this point EPPlus has no idea what to do with e.RowValues
            // we need to translate our values to something it understands
            // let's evaluate what we have:
            // e.RowValues contains a List of a List of CommenceValue
            // A list of CommenceValue represents a single item (row)
            // EPPlus cannot use anonymous values in LoadFromCollection,
            // so we will translate our results to a datatable first.
            // we do not use the more advanced ADO functionality in CmcLibNet,
            // just a flat table.

            using (ExcelPackage xl = new ExcelPackage(_fi))
            {
                var ws = xl.Workbook.Worksheets.FirstOrDefault(f => f.Name.Equals(_sheetName));
                if (ws == null)
                {
                    ws = xl.Workbook.Worksheets.Add(_sheetName);
                }

                _dataTable.Rows.Clear();
                foreach (List<CommenceValue> list in e.RowValues) // process rows
                {
                    object[] data = GetDataRowValues(list);
                    _dataTable.Rows.Add(data);
                }

                if (ws.Dimension == null) // first iteration
                {
                    ws.Cells.LoadFromDataTable(_dataTable, _settings.HeadersOnFirstRow);
                    if (_settings.HeadersOnFirstRow)
                    {
                        ws.Cells[1, 1, 1, ws.Dimension.End.Column].Style.Font.Bold = true;
                    }
                }
                else
                {
                    var lastRow = ws.Dimension.End.Row;
                    ws.Cells[lastRow + 1, 1].LoadFromDataTable(_dataTable, false);
                }
                int firstDataRow = _settings.HeadersOnFirstRow ? 2 : 1;
                SetNumberFormatStyles(ws, firstDataRow);
                try
                {
                    ws.Cells.AutoFitColumns(10, 50); 
                }
                catch (Exception) { } //throws an error in EPPLus on long strings. https://github.com/JanKallman/EPPlus/issues/445
                xl.Save();
            }
            base.BubbleUpProgressEvent(e);
        }

        private void SetNumberFormatStyles(ExcelWorksheet ws, int startRow)
        {
            for (int j = 0; j < columnDefinitions.Count(); j++)
            {
                int column = j + 1;
                ColumnDefinition cd = columnDefinitions[j];
                ExcelRange range = ws.Cells[startRow, column, ws.Dimension.End.Row, column];
                switch (cd.CommenceFieldDefinition.Type)
                {
                    case CommenceFieldType.Number:
                    case CommenceFieldType.Calculation:
                        range.Style.Numberformat.Format = "#,##0.00";
                        break;
                    case CommenceFieldType.Sequence:
                        range.Style.Numberformat.Format = "0";
                        break;
                    case CommenceFieldType.Date:
                        range.Style.Numberformat.Format = "dd-MM-yyyy";
                        break;
                    case CommenceFieldType.Time:
                        range.Style.Numberformat.Format = "hh:mm";
                        break;
                }
            }
        }

        private IEnumerable<object> GetValues(IEnumerable<CommenceValue> values)
        {
            foreach (CommenceValue cv in values)
            {
                if (!cv.ColumnDefinition.IsConnection)
                {
                    yield return CommenceValueConverter.ToAdoNet(cv.DirectFieldValue, cv.ColumnDefinition.CommenceFieldDefinition.Type) ?? DBNull.Value;
                }
                else
                {
                    if (cv.ConnectedFieldValues != null
                        && cv.ConnectedFieldValues.Length > 0)
                    {
                        // notice the trimming, EPPlus doesn't do it for you
                        yield return cv.ConnectedFieldValues[0].Left(MaxExcelCellSize); // we turned off splitting, so all values are in first element
                    }
                    else
                    {
                        yield return string.Empty; // always return something, otherwise column order gets thrown off
                    }
                }
            }
        }

        private object[] GetDataRowValues(IEnumerable<CommenceValue> values)
        {
            IEnumerable<object> rowvalues = GetValues(values);
            return rowvalues.ToArray();
        }

        // todo: new worksheet option
        protected internal override void WriteOut(string fileName)
        {
            _fi = new FileInfo(fileName);
            if (string.IsNullOrEmpty(fileName))
            {
                throw new ArgumentNullException("No filename supplied.");
            }

            // if file is in use abort
            if (base.IsFileLocked(_fi))
            {
                throw new IOException(string.Format("File \"{0}\" cannot be opened because it is already in use.", fileName));
            }

            // we want a new worksheet, delete if already exists
            if (_settings.XlUpdateOptions == ExcelUpdateOptions.CreateNewWorkbook && File.Exists(fileName))
            {
                File.Delete(fileName);
            }

            // we want to replace worksheet, clear existing values
            if (_settings.XlUpdateOptions == ExcelUpdateOptions.ReplaceWorksheet && File.Exists(fileName))
            {             
                ClearSheet(_sheetName);
            }

            if (_settings.XlUpdateOptions == ExcelUpdateOptions.CreateNewWorksheet)
            {
                _sheetName = NewSheet();
            }

            // start the data retrieval process that fires the events we'll respond to
            base.ReadCommenceData();
        }

        private void ClearSheet(string sheetName)
        {
            using (ExcelPackage xl = new ExcelPackage(_fi))
            {
                ExcelWorksheet ws = xl.Workbook.Worksheets.FirstOrDefault(f => f.Name.Equals(sheetName));
                if (ws != null)
                {
                    ws.Cells.Clear();
                    xl.SaveAs(_fi);
                }
            }
        }

        private string NewSheet()
        {
            using (ExcelPackage xl = new ExcelPackage(_fi))
            {
                _sheetName = _sheetName + (xl.Workbook.Worksheets.Count() + 1);
                var ws = xl.Workbook.Worksheets.Add(_sheetName);
                xl.SaveAs(_fi);
                return ws.Name;                
            }
        }

        // create a flat table
        private DataTable PrepareDataTable(string tableName, IEnumerable<ColumnDefinition> columnDefinitions)
        {
            DataTable dt = new DataTable(tableName);
            foreach (ColumnDefinition cd in columnDefinitions)
            {
                DataColumn dc;
                string caption = GetColumnLabel(cd);
                if (cd.IsConnection)
                {
                    dc = dt.Columns.Add("ConnectedValue" + dt.Columns.Count.ToString(), typeof(string));
                }
                else
                {
                    dc = dt.Columns.Add("DirectValue" + dt.Columns.Count.ToString(), Utils.GetTypeForCommenceField(cd.CommenceFieldDefinition.Type));
                }
                dc.Caption = caption;
                dc.AllowDBNull = true;
            }
            return dt;
        }

        private string GetColumnLabel(ColumnDefinition cd)
        {
            switch (_settings.HeaderMode)
            {
                case HeaderMode.CustomLabel:
                    return cd.CustomColumnLabel;
                default:
                    return cd.ColumnLabel; // will contain fieldname if not specified
            }
        }

    }
}
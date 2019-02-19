using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Vovin.CmcLibNet.Database;
using Vovin.CmcLibNet.Extensions;

namespace Vovin.CmcLibNet.Export
{
    internal class ExcelWriterUsingOpenXml : BaseWriter
    {
        //private SpreadsheetDocument doc = null;
        //private Sheet sh = null;
        private string _sheetName = string.Empty;
        private readonly int MAX_SHEETNAME_LENGTH = 31;// as per Microsoft documentation
        //private bool disposed = false;
        private readonly int MaxExcelCellSize = (int)Math.Pow(2, 15) - 1; // as per Microsoft documentation
        private readonly int MaxExcelNewLines = 253; // as per Microsoft documentation
        private List<ColumnDefinition> columnDefinitions = null;
        private Dictionary<string, string> existingSheets = null;
        private string _filename = string.Empty;
        private uint _sheetId;

        #region Constructors
        public ExcelWriterUsingOpenXml(ICommenceCursor cursor, IExportSettings settings) : base(cursor, settings)
        {
            columnDefinitions = new List<ColumnDefinition>(_settings.UseThids ? base.ColumnDefinitions.Skip(1) : base.ColumnDefinitions);
        }

        //~ExcelWriterUsingOpenXml()
        //{
        //    Dispose(false);
        //}
        #endregion

        #region Event handlers
        protected internal override void HandleDataReadComplete(object sender, ExportCompleteArgs e)
        {
            //doc.Dispose();
            base.BubbleUpCompletedEvent(e);
        }

        protected internal override void HandleProcessedDataRows(object sender, ExportProgressChangedArgs e)
        {
            // we just always append here
            // we open the file every time
            // this has a penalty performance but it is negligible
            using (SpreadsheetDocument doc = SpreadsheetDocument.Open(_filename, true))
            {

            }
        }
        #endregion

        #region Methods
        protected internal override void WriteOut(string fileName) { }
        protected internal override void WriteOut(string fileName, string sheetName)
        {
            if (string.IsNullOrEmpty(fileName))
            {
                throw new ArgumentNullException("No filename supplied.");
            }

            // if file is in use abort
            if (base.IsFileLocked(new FileInfo(fileName)))
            {
                throw new IOException(string.Format("File \"{0}\" cannot be opened because it is already in use.", fileName));
            }
            _filename = fileName;
            _sheetName = Utils.EscapeString(sheetName, "_").Left(MAX_SHEETNAME_LENGTH);
            // if workbook does not exist, create one
            if (!File.Exists(fileName) || _settings.XlUpdateOptions == ExcelUpdateOptions.CreateNewWorkbook)
            {
                CreateSpreadSheetDocument(fileName, sheetName);
            }
            else // if file exists, get a spreadsheet reference
            {
                switch (_settings.XlUpdateOptions)
                {
                    case ExcelUpdateOptions.CreateNewWorksheet: // insert new worksheet in existing doc
                        // we have a file, get the sheetnames in it
                        //existingSheets = GetSheetNames();
                        // we know the sheetname, changes the desired sheetname to create/update
                        //sheetName = Utils.AddUniqueIdentifier(sheetName, existingSheets.Values.ToList(), 0, (uint)Math.Pow(2, 10), (uint)MAX_SHEETNAME_LENGTH);
                        InsertNewSheet(sheetName);
                        break;

                    case ExcelUpdateOptions.RefreshWorksheet:
                        ClearSheet(sheetName);
                        break;

                    case ExcelUpdateOptions.AppendToWorksheet:
                        break;
                }
            }
            base.ReadCommenceData();
        }

        private void CreateSpreadSheetDocument(string fileName, string sheetName)
        {
            using (SpreadsheetDocument sd = SpreadsheetDocument.Create(fileName, SpreadsheetDocumentType.Workbook))
            {
                // overwrites existing
                // Add a WorkbookPart to the document.
                WorkbookPart workbookpart = sd.AddWorkbookPart();
                workbookpart.Workbook = new Workbook();

                // Add a WorksheetPart to the WorkbookPart.
                WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet(new SheetData());

                // Add Sheets to the Workbook.
                Sheets sheets = sd.WorkbookPart.Workbook.
                    AppendChild<Sheets>(new Sheets());

                // Append a new worksheet and associate it with the workbook.
                Sheet sheet = new Sheet()
                {
                    Id = sd.WorkbookPart.GetIdOfPart(worksheetPart),
                    SheetId = 1,
                    Name = sheetName
                };
                sheets.Append(sheet);
                //return sheet.SheetId;
            }
        }

        private void InsertNewSheet(string sheetName)
        {
            // Add a blank WorksheetPart.
            using (SpreadsheetDocument doc = SpreadsheetDocument.Open(_filename, true))
            {
                WorksheetPart newWorksheetPart = doc.WorkbookPart.AddNewPart<WorksheetPart>();
                newWorksheetPart.Worksheet = new Worksheet(new SheetData());

                Sheets sheets = doc.WorkbookPart.Workbook.GetFirstChild<Sheets>();
                string relationshipId = doc.WorkbookPart.GetIdOfPart(newWorksheetPart);

                // Get a unique ID for the new worksheet.
                uint sheetId = 1;
                if (sheets.Elements<Sheet>().Count() > 0)
                {
                    sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
                }
                // Append the new worksheet and associate it with the workbook.
                Sheet sheet = new Sheet() { Id = relationshipId, SheetId = sheetId, Name = sheetName };
                sheets.Append(sheet);
                //return this._sheetId;
            }
        }

        private void ClearSheet(string name)
        {
            using (SpreadsheetDocument doc = SpreadsheetDocument.Open(_filename, true))
            {
                // get a reference to the existing worksheet.
                Sheet sheet = doc.WorkbookPart.Workbook.Descendants<Sheet>()
                            .Where(s => s.Name == name).FirstOrDefault();
                if (sheet == null) { return; }

                // sheets contain a worksheet object with a sheetData object
                // all cell data are in there
                // string values are in a SharedStringTable, other cell-related stuff are in the cells
                // so we must inspect the cells before deleting them
                // also will we just delete all cell-values? I think that is best.
                List<Cell> cells = sheet.Descendants<Cell>().ToList();

            }
        }

        ///// <summary>
        ///// Scans workbook for Sheets.
        ///// </summary>
        ///// <returns>Dictionary of sheets and their corresponding id's.</returns>
        //private Dictionary<string, string> GetSheetNames()
        //{
        //    Dictionary<string, string> retval = new Dictionary<string, string>();
        //    using (SpreadsheetDocument doc = SpreadsheetDocument.Open(_filename, false))
        //    { 
        //    WorkbookPart wbPart = doc.WorkbookPart;
        //    var sheets = wbPart.Workbook.Sheets;
        //    string x = string.Empty;
        //    foreach (Sheet s in sheets)
        //    {
        //        retval.Add(s.SheetId, s.Name);
        //    }
        //}
        //    return retval;
        //}
        #endregion



        //#region Dispose
        //// Protected implementation of Dispose pattern.
        ///// <summary>
        ///// Dispose method.
        ///// </summary>
        ///// <param name="disposing">disposing.</param>
        //protected override void Dispose(bool disposing)
        //{
        //    if (disposed)
        //        return;

        //    if (disposing)
        //    {
        //        // Free any other managed objects here.
        //        //
        //    }

        //    // Free any unmanaged objects here.
        //    //
        //    disposed = true;

        //    // Call the base class implementation.
        //    base.Dispose(disposing);
        //}
        //#endregion
    }
}

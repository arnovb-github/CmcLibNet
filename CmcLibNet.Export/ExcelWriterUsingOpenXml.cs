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
        private SpreadsheetDocument doc = null;
        private Sheet sh = null;
        private string sheetName = string.Empty;
        private const int MAX_SHEETNAME_LENGTH = 31;// as per Microsoft documentation
        private bool disposed = false;
        private readonly int MaxExcelCellSize = (int)Math.Pow(2, 15) - 1; // as per Microsoft documentation
        private readonly int MaxExcelNewLines = 253; // as per Microsoft documentation
        private List<ColumnDefinition> columnDefinitions = null;
        private Dictionary<string, string> existingSheets = null;

        #region Constructors
        public ExcelWriterUsingOpenXml(ICommenceCursor cursor, IExportSettings settings) : base(cursor, settings)
        {
            columnDefinitions = new List<ColumnDefinition>(_settings.UseThids ? base.ColumnDefinitions.Skip(1) : base.ColumnDefinitions);
            sheetName = Utils.EscapeString(base._dataSourceName, "_").Left(MAX_SHEETNAME_LENGTH);
        }

        ~ExcelWriterUsingOpenXml()
        {
            Dispose(false);
        }
        #endregion

        #region Methods
        protected internal override void WriteOut(string fileName)
        {
            // if file is in use abort
            if (IsFileLocked(new FileInfo(fileName)))
            {
                Dispose(false);
                throw new IOException(string.Format("File \"{0}\" cannot be opened because it is already in use.", fileName));
            }

            // if workbook does not exist, create one
            if (!File.Exists(fileName) || _settings.XlUpdateOptions == ExcelUpdateOptions.CreateNewWorkbook)
            {
                doc = CreateSpreadSheetDocument(fileName, sheetName);
            }
            else // if file exists, get a spreadsheet reference
            {
                // we have a file, get the sheetnames in it
                doc = SpreadsheetDocument.Open(fileName, true); // open for editing
                existingSheets = GetSheetNames(doc);
                if (_settings.XlUpdateOptions == ExcelUpdateOptions.CreateNewWorksheet)
                {
                    // insert new worksheet in existing doc
                    sheetName = Utils.AddUniqueIdentifier(sheetName, existingSheets.Values.ToList(), 0, (uint)Math.Pow(2, 10), MAX_SHEETNAME_LENGTH);
                    sh = InsertNewSheet(doc, sheetName);
                }

                if (_settings.XlUpdateOptions == ExcelUpdateOptions.UpdateWorksheet)
                {
                    // Retrieve a reference to the existing worksheet.
                    sh = doc.WorkbookPart.Workbook.Descendants<Sheet>()
                        .Where(s => s.Name == sheetName).FirstOrDefault();
                }
            }
            base.ReadCommenceData();
        }

        private SpreadsheetDocument CreateSpreadSheetDocument(string fileName, string sheetName)
        {
            SpreadsheetDocument sd = SpreadsheetDocument.Create(fileName, SpreadsheetDocumentType.Workbook); // overwrites existing
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
            sd.Save();
            return sd;
        }

        private Sheet InsertNewSheet(SpreadsheetDocument doc, string name)
        {
            // Add a blank WorksheetPart.
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

            // Give the new worksheet a name.
            string sheetName = name;

            // Append the new worksheet and associate it with the workbook.
            Sheet sheet = new Sheet() { Id = relationshipId, SheetId = sheetId, Name = sheetName };
            sheets.Append(sheet);
            return sheet;
        }

        private Dictionary<string, string> GetSheetNames(SpreadsheetDocument doc)
        {
            Dictionary<string, string> retval = new Dictionary<string, string>();
            WorkbookPart wbPart = doc.WorkbookPart;
            var sheets = wbPart.Workbook.Sheets;
            string x = string.Empty;
            foreach (Sheet s in sheets)
            {
                retval.Add(s.SheetId, s.Name);
            }
            return retval;
        }
        #endregion

        #region Event handlers
        protected internal override void HandleDataReadComplete(object sender, ExportCompleteArgs e)
        {
            doc.Dispose();
            base.BubbleUpCompletedEvent(e);
        }

        protected internal override void HandleProcessedDataRows(object sender, ExportProgressChangedArgs e)
        {
            return;
        }
        #endregion

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
            }

            // Free any unmanaged objects here.
            //
            disposed = true;

            // Call the base class implementation.
            base.Dispose(disposing);
        }
        #endregion
    }
}

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
    /// <summary>
    /// Writes a Microsoft Excel (xlsx) file using OpenXML.
    /// This class takes the DOM approach (as opposed to the SAX approach).
    /// This means that when dealing with very large datasets, memory may be an issue.
    /// </summary>
    [Obsolete("Replaced with EPPlus implementation")]
    internal class ExcelWriterUsingOpenXml : BaseWriter
    {
        #region Fields
        private SpreadsheetDocument _spreadSheetDocument = null;
        private string _sheetName = string.Empty;
        private readonly int MaxSheetNameLength = 31;// as per Microsoft documentation
        private readonly int MaxExcelCellSize = (int)Math.Pow(2, 15) - 1; // as per Microsoft documentation
        //private readonly int MaxExcelNewLines = 253; // as per Microsoft documentation
        private readonly List<ColumnDefinition> columnDefinitions = null;
        private Dictionary<string, string> existingSheets = null;
        private string _filename = string.Empty;
        #endregion

        #region Constructors
        public ExcelWriterUsingOpenXml(ICommenceCursor cursor, IExportSettings settings) : base(cursor, settings)
        {
            columnDefinitions = new List<ColumnDefinition>(_settings.UseThids ? base.ColumnDefinitions.Skip(1) : base.ColumnDefinitions);
            _sheetName = string.IsNullOrEmpty(settings.CustomRootNode) ? Utils.EscapeString(_dataSourceName, "_").Left(MaxSheetNameLength) : settings.CustomRootNode;
            settings.ISO8601Compliant = true; // override custom setting(s)
            settings.SplitConnectedItems = false; // override custom setting(s)
        }

        #endregion

        #region Event handlers
        protected internal override void HandleDataReadComplete(object sender, ExportCompleteArgs e)
        {
            if (_spreadSheetDocument != null) { _spreadSheetDocument.Dispose(); } // saves, flushes and disposes.
            base.BubbleUpCompletedEvent(e);
        }

        protected internal override void HandleProcessedDataRows(object sender, ExportProgressChangedArgs e)
        {
            // we just always append at this point
            AppendRows(e.RowValues);
            base.BubbleUpProgressEvent(e);
        }
        #endregion

        #region Methods
        protected internal override void WriteOut(string fileName)
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
            _sheetName = Utils.EscapeString(_sheetName, "_").Left(MaxSheetNameLength);
            // if workbook does not exist, create one
            if (!File.Exists(fileName) || _settings.XlUpdateOptions == ExcelUpdateOptions.CreateNewWorkbook)
            {
                CreateSpreadSheetDocument(fileName, _sheetName);
            }
            else // if file exists, get a spreadsheet reference
            {
                // we have a file, get the sheetnames in it
                existingSheets = GetSheetNames();
                switch (_settings.XlUpdateOptions)
                {
                    case ExcelUpdateOptions.CreateNewWorksheet: // insert new worksheet in existing doc
                        // sheetname must be unique
                        _sheetName = Utils.AddUniqueIdentifier(_sheetName, existingSheets.Values.ToList(), 0, (uint)Math.Pow(2, 10), (uint)MaxSheetNameLength);
                        InsertNewSheet(_sheetName);
                        break;
                    case ExcelUpdateOptions.ReplaceWorksheet:
                        if (!existingSheets.ContainsValue(_sheetName))
                        {
                            InsertNewSheet(_sheetName);
                        }
                        else
                        {
                            DeleteWorkSheet(fileName, _sheetName);
                            InsertNewSheet(_sheetName);
                        }
                        break;
                    case ExcelUpdateOptions.AppendToWorksheet:
                        // nothing here, just included for completeness
                        break;
                }
            }
            base.ReadCommenceData(); // starts the data retrieval process that fires the events we'll respond to
        }

        /// <summary>
        /// Appends the Commence data to the workbook.
        /// </summary>
        /// <param name="data">Commence data.</param>
        private void AppendRows(List<List<CommenceValue>> data)
        {
            if (_spreadSheetDocument == null)
            {
                _spreadSheetDocument = SpreadsheetDocument.Open(_filename, true);
            }
            // determine current rowindex
            uint lastRowIndex = GetLastRowIndex(_spreadSheetDocument);
            if (CurrentRowIndex == 0) { CurrentRowIndex = ++lastRowIndex; }

            WorkbookPart wbPart = _spreadSheetDocument.WorkbookPart;
            Sheet sheet = wbPart.Workbook.Descendants<Sheet>().Where(s => s.Name.Equals(_sheetName)).FirstOrDefault();
            WorksheetPart worksheetPart = (WorksheetPart)(wbPart.GetPartById(sheet.Id));
            SheetData sd = worksheetPart.Worksheet.Elements<SheetData>().FirstOrDefault();

            // append headers when indicated and not already there
            if (CurrentRowIndex == 1 && _settings.HeadersOnFirstRow)
            {
                AppendHeaders(_spreadSheetDocument);
                CurrentRowIndex++;
            }

            // append the data
            foreach (List<CommenceValue> row in data) // loop Commence rows
            {
                Row r = new Row() { RowIndex = CurrentRowIndex };
                foreach (CommenceValue v in row) // loop Commence row
                {
                    Cell c = new Cell();
                    if (v.ColumnDefinition.IsConnection)
                    {
                        // Note that we assume that the connected items are not split!
                        if (v.ConnectedFieldValues != null)
                        {
                            string value = Utils.RemoveControlCharacters(v.ConnectedFieldValues[0]);
                            value = FitExcelCellValue(value);
                            c.CellValue = new CellValue(value);
                            // inline strings cannot deal with line breaks
                            c.DataType = CellValues.String; // should we bother putting this in SharedString?
                            c.StyleIndex = GetStyleIndexForCommenceType(CommenceFieldType.Text);
                        }
                    }
                    else
                    {
                        string value = Utils.RemoveControlCharacters(v.DirectFieldValue);
                        value = FitExcelCellValue(value);
                        c.DataType = Utils.GetTypeForOpenXml(v.ColumnDefinition.CommenceFieldDefinition.Type); // inline strings cannot deal with line breaks
                        if (v.DirectFieldValue.Contains("\r\n"))
                        {
                            c.DataType = CellValues.String; // inline strings cannot deal with line breaks
                        }
                        c.CellValue = new CellValue(value);
                        c.StyleIndex = GetStyleIndexForCommenceType(v.ColumnDefinition.CommenceFieldDefinition.Type);
                    }
                    r.Append(c);
                }
                sd.Append(r);
                CurrentRowIndex++;
            }
            //worksheetPart.Worksheet.Save(); // very costly call, and seems to increase memory usage
        }

        /// <summary>
        /// Returns a string matching maximum Excell cell limits.
        /// </summary>
        /// <param name="value">input string.</param>
        /// <returns>Valid Excell cell value.</returns>
        private string FitExcelCellValue(string value)
        {
            if (string.IsNullOrEmpty(value) || !value.Contains("\n")) { return value; }

            string cr = "\r\n";
            // normalize line endings
            string v = value.Replace(cr, "\n").Replace("\n", cr);
            // Check if the number of newlines does't exceed maximum
            // It seems that while documentation says the maximum number is 253,
            // control characters are not counted at all?
            // See: https://support.office.com/en-us/article/excel-specifications-and-limits-1672b34d-7043-467e-8e27-269d656771c3
            //if (v.Count(c => c.Equals('\n')) > MaxExcelNewLines)
            //{
            // v = v.Substring(0, v.IndexOfNthChar('\n', 0, MaxExcelNewLines)); 
            //}
            // see if string doesn't exceed maximum length
            // Excel doesn't seem to count control characters,
            // but string.Length takes them into account
            // This means that this can return slightly fewer data than the Excel cell will actually accept
            //string controlChars = new string(v.Where(c => char.IsControl(c)).ToArray());
            v = v.Length > MaxExcelCellSize ? v.Substring(0, MaxExcelCellSize) : v;
            return v;
        }
        /// <summary>
        /// Appends headers to top-row.
        /// </summary>
        /// <param name="spreadSheetDocument">SpreadsheetDocument.</param>
        private void AppendHeaders(SpreadsheetDocument spreadSheetDocument)
        {
            // we want out headers to be bold

            // see if there is already a cellformat associated with bold font
            // it may have other properties, so you could be in for a surprise
            // if you use an existing spreadsheet
            var styleIndex = FindCellFormatWithFontFormat(spreadSheetDocument, s => s.Bold != null);

            if (styleIndex == null)
            {
                styleIndex = AddFontBoldToStyleSheet(spreadSheetDocument);
            }
            Row r = new Row() { RowIndex = CurrentRowIndex };
            foreach (string header in base.ExportHeaders)
            {
                r.Append(new Cell()
                {
                    CellValue = new CellValue(header),
                    DataType = CellValues.String,
                    StyleIndex = styleIndex
                });
            }
            WorkbookPart wbPart = _spreadSheetDocument.WorkbookPart;
            Sheet sheet = wbPart.Workbook.Descendants<Sheet>().Where(s => s.Name.Equals(_sheetName)).FirstOrDefault();
            WorksheetPart worksheetPart = (WorksheetPart)(wbPart.GetPartById(sheet.Id));
            SheetData sd = worksheetPart.Worksheet.Elements<SheetData>().FirstOrDefault();
            sd.Append(r);
        }

        /// <summary>
        /// Creates new spreadsheet document (with stylesheet).
        /// </summary>
        /// <param name="fileName">File Name.</param>
        /// <param name="sheetName">Sheet name.</param>
        private void CreateSpreadSheetDocument(string fileName, string sheetName)
        {
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(fileName, SpreadsheetDocumentType.Workbook))
            {
                // overwrites existing
                // Add a WorkbookPart to the document.
                WorkbookPart workbookpart = spreadsheetDocument.AddWorkbookPart();
                workbookpart.Workbook = new Workbook();

                // create stylesheet so we can apply formatting
                AddStyleSheet(spreadsheetDocument);

                // Add a WorksheetPart to the WorkbookPart.
                WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet(new SheetData());

                // Add Sheets to the Workbook.
                Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.
                AppendChild<Sheets>(new Sheets());

                // Append a new worksheet and associate it with the workbook.
                Sheet sheet = new Sheet()
                {
                    Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart),
                    SheetId = 1,
                    Name = sheetName
                };
                sheets.Append(sheet);
                //return sheet.SheetId;
            }
        }

        /// <summary>
        /// Adds minimal stylesheet to workbook
        /// </summary>
        /// <param name="spreadsheetDocument">workbook</param>
        private void AddStyleSheet(SpreadsheetDocument spreadsheetDocument)
        {
            
            WorkbookStylesPart stylesheet = spreadsheetDocument.WorkbookPart.AddNewPart<WorkbookStylesPart>();

            Stylesheet workbookstylesheet = new Stylesheet();

            // Create "fonts" node.
            Fonts fonts = new Fonts();
            fonts.Append(new Font() // default font
            {
                FontName = new FontName() { Val = "Calibri" },
                FontSize = new FontSize() { Val = 11 },
                FontFamilyNumbering = new FontFamilyNumbering() { Val = 2 },
            });

            // <Fills>
            Fills fills = new Fills();      // Default fill
            fills.Append(new Fill());

            // <Borders>
            Borders borders = new Borders(); // default border
            borders.Append(new Border()); 

            // <CellFormats>
            // First define a default CellFormat; this is mandatory
            CellFormat cellformat0 = new CellFormat() {
                BorderId = 0,
                FillId = 0,
                FontId = 0,
                NumberFormatId = (uint)OutputNumberFormat.General,
                FormatId = 0,
                ApplyNumberFormat = false
            }; // Default style : Mandatory; Style ID =0

            // Number
            CellFormat cellformat1 = new CellFormat()
            {
                BorderId = 0,
                FillId = 0,
                FontId = 0,
                NumberFormatId = (uint)OutputNumberFormat.Number,
                FormatId = 0,
                ApplyNumberFormat = true
            };  // Style ID = 1

            // Date
            CellFormat cellformat2 = new CellFormat()
            {
                BorderId = 0,
                FillId = 0,
                FontId = 0,
                NumberFormatId = (uint)OutputNumberFormat.Date,
                FormatId = 0,
                ApplyNumberFormat = true
            };  // Style ID = 2

            // Time
            CellFormat cellformat3 = new CellFormat()
            {
                BorderId = 0,
                FillId = 0,
                FontId = 0,
                NumberFormatId = (uint)OutputNumberFormat.Time,
                FormatId = 0,
                ApplyNumberFormat = true
            };  // Style ID = 3

            // append cellformats
            CellFormats cellformats = new CellFormats();
            cellformats.Append(cellformat0);
            cellformats.Append(cellformat1);
            cellformats.Append(cellformat2);
            cellformats.Append(cellformat3);

            // Append FONTS, FILLS , BORDERS & CellFormats to stylesheet <Preserve the ORDER>
            workbookstylesheet.Append(fonts);
            workbookstylesheet.Append(fills);
            workbookstylesheet.Append(borders);
            workbookstylesheet.Append(cellformats);

            // Finalize
            stylesheet.Stylesheet = workbookstylesheet;
            stylesheet.Stylesheet.Save();
        }

        /// <summary>
        /// Add a CellFormat with Font Bold to stylesheet
        /// </summary>
        /// <param name="spreadsheetDocument">SpreadsheetDocument</param>
        /// <returns>StyleIndex</returns>
        private uint AddFontBoldToStyleSheet(SpreadsheetDocument spreadsheetDocument)
        {
            Stylesheet workbookStylesheet = spreadsheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet;

            // new Bold font
            Font font = new Font()
            {
                Bold = new Bold()
            };

            workbookStylesheet.Fonts.Append(font);
            int fontId = workbookStylesheet.Fonts.ChildElements.Count - 1;

            // create a CellFormat with the new font
            CellFormat newCellFormat = new CellFormat()
            {
                // keep everything default except NumberFormatId
                BorderId = 0,
                FillId = 0,
                FontId = (uint)fontId,
                NumberFormatId = 0,
                FormatId = 0,
            };
            workbookStylesheet.CellFormats.Append(newCellFormat);

            // Finalize
            spreadsheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet.Save();
            // return the index of the new CellFormat
            return (uint)workbookStylesheet.CellFormats.ChildElements.Count - 1;
        }

        /// <summary>
        /// Create a CellFormat for the desired display format.
        /// OpenXML specs define 163 of them.
        /// Excel will render them according to locale.
        /// </summary>
        /// <param name="numberFormat"></param>
        /// <returns></returns>
        private uint AddNumberFormatToStyleSheet(OutputNumberFormat numberFormat)
        {
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(_filename, true))
            {
                WorkbookStylesPart stylesheet = spreadsheetDocument.WorkbookPart.AddNewPart<WorkbookStylesPart>();
                Stylesheet workbookstylesheet = stylesheet.Stylesheet;

                CellFormat newCellFormat = new CellFormat()
                {
                    // keep everything default except NumberFormatId
                    BorderId = 0,
                    FillId = 0,
                    FontId = 0,
                    NumberFormatId = (uint)numberFormat,
                    FormatId = 0,
                    ApplyNumberFormat = true
                };
                CellFormats cellformats = new CellFormats();
                cellformats.Append(newCellFormat);
                    
                workbookstylesheet.Append(cellformats);

                // Finalize
                stylesheet.Stylesheet = workbookstylesheet;
                stylesheet.Stylesheet.Save();
                // return the index of the new CellFormat
                //return (uint)workbookstylesheet.Elements<CellFormat>().ToList().IndexOf(newCellFormat); // can be made faster
                return workbookstylesheet.CellFormats.Count - 1;
            }
        }

        /// <summary>
        /// Inserts a new worksheet
        /// </summary>
        /// <param name="sheetName">Sheet name.</param>
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

        /// <summary>
        /// Deletes worksheet and its dependencies.
        /// </summary>
        /// <param name="fileName">Workbook to delete sheet from.</param>
        /// <param name="sheetToDelete">Worksheetto delete.</param>
        /// <remarks>Code taken and modified from <a href="https://blogs.msdn.microsoft.com/vsod/2010/02/05/how-to-delete-a-worksheet-from-excel-using-open-xml-sdk-2-0/">Ankush Bhatia</a>.</remarks>
        private void DeleteWorkSheet(string fileName, string sheetToDelete)
        {
            string Sheetid = string.Empty;
            //Open the workbook
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(fileName, true))
            {
                WorkbookPart wbPart = document.WorkbookPart;

                // Get the pivot Table Parts
                IEnumerable<PivotTableCacheDefinitionPart> pvtTableCacheParts = wbPart.PivotTableCacheDefinitionParts;
                Dictionary<PivotTableCacheDefinitionPart, string> pvtTableCacheDefinationPart = new Dictionary<PivotTableCacheDefinitionPart, string>();
                foreach (PivotTableCacheDefinitionPart Item in pvtTableCacheParts)
                {
                    PivotCacheDefinition pvtCacheDef = Item.PivotCacheDefinition;
                    //Check if this CacheSource is linked to SheetToDelete
                    var pvtCahce = pvtCacheDef.Descendants<CacheSource>().Where(s => s.WorksheetSource.Sheet == sheetToDelete);
                    if (pvtCahce.Count() > 0)
                    {
                        pvtTableCacheDefinationPart.Add(Item, Item.ToString());
                    }
                }
                foreach (var Item in pvtTableCacheDefinationPart)
                {
                    wbPart.DeletePart(Item.Key);
                }
                //Get the SheetToDelete from workbook.xml
                Sheet theSheet = wbPart.Workbook.Descendants<Sheet>().Where(s => s.Name == sheetToDelete).FirstOrDefault();
                if (theSheet == null)
                {
                    return;
                }
                //Store the SheetID for the reference
                Sheetid = theSheet.SheetId;

                // Remove the sheet reference from the workbook.
                WorksheetPart worksheetPart = (WorksheetPart)(wbPart.GetPartById(theSheet.Id));
                theSheet.Remove();

                // Delete the worksheet part.
                wbPart.DeletePart(worksheetPart);

                //Get the DefinedNames
                var definedNames = wbPart.Workbook.Descendants<DefinedNames>().FirstOrDefault();
                if (definedNames != null)
                {
                    List<DefinedName> defNamesToDelete = new List<DefinedName>();

                    foreach (DefinedName Item in definedNames)
                    {
                        // This condition checks to delete only those names which are part of Sheet in question
                        if (Item.Text.Contains(sheetToDelete + "!"))
                            defNamesToDelete.Add(Item);
                    }

                    foreach (DefinedName Item in defNamesToDelete)
                    {
                        Item.Remove();
                    }

                }
                // Get the CalculationChainPart 
                //Note: An instance of this part type contains an ordered set of references to all cells in all worksheets in the 
                //workbook whose value is calculated from any formula

                CalculationChainPart calChainPart;
                calChainPart = wbPart.CalculationChainPart;
                if (calChainPart != null)
                {
                    var calChainEntries = calChainPart.CalculationChain.Descendants<CalculationCell>().Where(c => c.SheetId == Sheetid);
                    List<CalculationCell> calcsToDelete = new List<CalculationCell>();
                    foreach (CalculationCell Item in calChainEntries)
                    {
                        calcsToDelete.Add(Item);
                    }

                    foreach (CalculationCell Item in calcsToDelete)
                    {
                        Item.Remove();
                    }

                    if (calChainPart.CalculationChain.Count() == 0)
                    {
                        wbPart.DeletePart(calChainPart);
                    }
                }
                // force recalculation if any
                var calcProps = document.WorkbookPart.Workbook.Elements<CalculationProperties>().FirstOrDefault();
                if (calcProps != null)
                {
                    document.WorkbookPart.Workbook.CalculationProperties.ForceFullCalculation = true;
                    document.WorkbookPart.Workbook.CalculationProperties.FullCalculationOnLoad = true;
                }
            } // using autosaves
        }

        /// <summary>
        /// Scans workbook for Sheets.
        /// </summary>
        /// <returns>Dictionary of sheets and their corresponding id's.</returns>
        private Dictionary<string, string> GetSheetNames()
        {
            Dictionary<string, string> retval = new Dictionary<string, string>();
            using (SpreadsheetDocument doc = SpreadsheetDocument.Open(_filename, false))
            {
                WorkbookPart wbPart = doc.WorkbookPart;
                var sheets = wbPart.Workbook.Sheets;
                string x = string.Empty;
                foreach (Sheet s in sheets)
                {
                    retval.Add(s.SheetId, s.Name);
                }
            }
            return retval;
        }

        //private uint? FindCellFormatInStyleSheet(SpreadsheetDocument spreadsheet, OutputNumberFormat numberFormatId)
        //{
        //    WorkbookStylesPart stylesheet = spreadsheet.WorkbookPart.WorkbookStylesPart;
        //    CellFormat cf = stylesheet.Stylesheet
        //    .CellFormats.Elements<CellFormat>()
        //    .Where(s => s.NumberFormatId == (uint)numberFormatId).FirstOrDefault();
        //    if (cf == null)
        //    {
        //        return null;
        //    }
        //    else
        //    {
        //        // there is a CellFormat defined
        //        // but what StyleIndex does it have?
        //        // that is a cryptic way of saying we need to know which of the CellFormat elements this is
        //        // they are identified simply by their order
        //        return (uint)stylesheet.Stylesheet
        //            .CellFormats.Elements<CellFormat>().ToList().IndexOf(cf);
        //    }
        //}

        /// <summary>
        /// Searches stylesheet for fonts with a certain property.
        /// </summary>
        /// <param name="spreadsheet">Spreadsheet document</param>
        /// <param name="func">Property test.</param>
        /// <returns>StyleIndexId or <c>null</c> if not found.</returns>
        private uint? FindCellFormatWithFontFormat(SpreadsheetDocument spreadsheet, Func<Font, bool> func)
        {
            WorkbookStylesPart stylesheet = spreadsheet.WorkbookPart.WorkbookStylesPart;
            // find first font that is defined as bold
            Font font = stylesheet.Stylesheet.Fonts.Elements<Font>().Where(func).FirstOrDefault();
            // get the FontId for this font
            if (font == null) { return null; }
            int fontId = stylesheet.Stylesheet.Fonts.Elements<Font>().ToList().IndexOf(font);
            // now find the CellFormat that points to that font, if any
            CellFormat cf = stylesheet.Stylesheet.Fonts.Elements<CellFormat>().Where(s => s.FontId == fontId).FirstOrDefault();
            // now find the CellFormatId
            int cellFormatId = stylesheet.Stylesheet.Elements<CellFormat>().ToList().IndexOf(cf); // returns -1 if not found
            if (cellFormatId != -1)
            {
                return (uint)cellFormatId;
            }
            return null;
        }

        /// <summary>
        /// Searches stylesheet for CellFormat with a certains property.
        /// </summary>
        /// <param name="spreadsheet">Spreadsheet document.</param>
        /// <param name="func">Property test.</param>
        /// <returns>StyleIndexId or <c>null</c> if not found.</returns>
        private uint? FindCellFormatInStyleSheet(SpreadsheetDocument spreadsheet, Func<CellFormat, bool> func)
        {
            WorkbookStylesPart stylesheet = spreadsheet.WorkbookPart.WorkbookStylesPart;
            CellFormat cf = stylesheet.Stylesheet
            .CellFormats.Elements<CellFormat>()
            .Where(func).FirstOrDefault();
            if (cf == null)
            {
                return null;
            }
            else
            {
                // there is a CellFormat defined
                // but what StyleIndex does it have?
                // that is a cryptic way of saying we need to know which of the CellFormat elements this is
                // they are identified simply by their order
                return (uint)stylesheet.Stylesheet
                    .CellFormats.Elements<CellFormat>().ToList().IndexOf(cf);
            }
        }

        /// <summary>
        /// Gets last rowindex.
        /// </summary>
        /// <returns></returns>
        private uint GetLastRowIndex(SpreadsheetDocument spreadSheetDocument)
        {
            uint retval = 0;
            WorkbookPart wbPart = spreadSheetDocument.WorkbookPart;
            Sheet sheet = wbPart.Workbook.Descendants<Sheet>().Where(s => s.Name.Equals(_sheetName)).FirstOrDefault();
            WorksheetPart worksheetPart = (WorksheetPart)(wbPart.GetPartById(sheet.Id));
            SheetData sd = worksheetPart.Worksheet.Elements<SheetData>().FirstOrDefault();
            if (sd == null) { return retval; }
            // we are responsible for supplying our own, correct rowindex
            // so we need to find out what the last one is
            Row lastRow = sd.Elements<Row>().OrderBy(o => o.RowIndex).LastOrDefault();
            retval = (lastRow == null) ? 0 : lastRow.RowIndex;
            return retval;
        }

        /// <summary>
        /// Translates the Commence Field Type to the property that returns the Style Index to use with it.
        /// </summary>
        /// <param name="cft">Commence fieldtype.</param>
        /// <returns>Property.</returns>
        private uint GetStyleIndexForCommenceType(CommenceFieldType cft)
        {
            switch (cft)
            {
                case CommenceFieldType.Number:
                case CommenceFieldType.Sequence:
                case CommenceFieldType.Calculation:
                    return this.CellFormatStyleIndexNumber;
                case CommenceFieldType.Date:
                    return this.CellFormatStyleIndexDate;
                case CommenceFieldType.Time:
                    return this.CellFormatStyleIndexTime;
                default:
                    return this.CellFormatStyleIndexGeneral;
            }
        }


        #endregion

        #region Properties
        private uint CurrentRowIndex { get; set; }
        private uint StyleSheetId { get; set; }

        private uint? cellFormatStyleIndexGeneral = null;
        private uint CellFormatStyleIndexGeneral
        {
            get
            {
                if (cellFormatStyleIndexGeneral == null)
                {
                    cellFormatStyleIndexGeneral = FindCellFormatInStyleSheet(_spreadSheetDocument, f => f.NumberFormatId == (uint)OutputNumberFormat.General);
                    if (cellFormatStyleIndexGeneral == null)
                    {
                        cellFormatStyleIndexGeneral = AddNumberFormatToStyleSheet(OutputNumberFormat.General);
                    }
                }
                return (uint)cellFormatStyleIndexGeneral;
            }
            set
            {
                cellFormatStyleIndexGeneral = value;
            }
        }

        private uint? cellFormatStyleIndexNumber = null;
        private uint CellFormatStyleIndexNumber
        {
            get
            {
                if (cellFormatStyleIndexNumber == null)
                {
                    cellFormatStyleIndexNumber = FindCellFormatInStyleSheet(_spreadSheetDocument, f => f.NumberFormatId == (uint)OutputNumberFormat.Number);
                    if (cellFormatStyleIndexNumber == null)
                    {
                        cellFormatStyleIndexNumber = AddNumberFormatToStyleSheet(OutputNumberFormat.Number);
                    }
                }
                return (uint)cellFormatStyleIndexNumber;
            }
            set
            {
                cellFormatStyleIndexNumber = value;
            }
        }

        private uint? cellFormatStyleIndexTime = null;
        private uint CellFormatStyleIndexTime
        {
            get
            {
                if (cellFormatStyleIndexTime == null)
                {
                    cellFormatStyleIndexTime = FindCellFormatInStyleSheet(_spreadSheetDocument, f => f.NumberFormatId == (uint)OutputNumberFormat.Time);
                    if (cellFormatStyleIndexTime == null)
                    {
                        cellFormatStyleIndexTime = AddNumberFormatToStyleSheet(OutputNumberFormat.Time);
                    }
                }
                return (uint)cellFormatStyleIndexTime;
            }
            set
            {
                cellFormatStyleIndexTime = value;
            }
        }

        private uint? cellFormatStyleIndexDate = null;
        private uint CellFormatStyleIndexDate
        {
            get
            {
                if (cellFormatStyleIndexDate == null)
                {
                    cellFormatStyleIndexDate = FindCellFormatInStyleSheet(_spreadSheetDocument, f => f.NumberFormatId == (uint)OutputNumberFormat.Date);
                    if (cellFormatStyleIndexDate == null)
                    {
                        cellFormatStyleIndexDate = AddNumberFormatToStyleSheet(OutputNumberFormat.Date);
                    }
                }
                return (uint)cellFormatStyleIndexDate;
            }
            set
            {
                cellFormatStyleIndexDate = value;
            }
        }
        #endregion

        #region Enumerations
        private enum OutputNumberFormat : uint
        {
            General = 0,
            Number = 4, // '#,##0.00';
            Date = 14, // 'm/d/yyyy'; Excel translates this to Locale
            Time = 20, // 'h:mm';
        }
        #endregion
    }
}

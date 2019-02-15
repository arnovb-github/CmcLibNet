using System;
using Vovin.CmcLibNet.Database;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.Generic;
using System.Linq;
using System.IO;

namespace Vovin.CmcLibNet.Export
{
    internal class ExcelWriterUsingOpenXml : BaseWriter
    {
        private SpreadsheetDocument doc = null;
        private Workbook wb = null;
        private Worksheet ws = null;
        private bool disposed = false;
        private string fileName;
        private readonly int MaxExcelCellSize = (int)Math.Pow(2, 15) - 1; // as per Microsoft documentation
        private readonly int MaxExcelNewLines = 253; // as per Microsoft documentation
        private List<ColumnDefinition> columnDefinitions = null;

        #region Constructors
        public ExcelWriterUsingOpenXml(ICommenceCursor cursor, IExportSettings settings) : base(cursor, settings)
        {
            columnDefinitions = new List<ColumnDefinition>(_settings.UseThids ? base.ColumnDefinitions.Skip(1) : base.ColumnDefinitions);
        }
        #endregion

        #region Methods
        protected internal override void WriteOut(string fileName)
        {
            if (_settings.DeleteExcelFileBeforeExport)
            {
                File.Delete(fileName);
            }
            doc = SpreadsheetDocument.Create(fileName, SpreadsheetDocumentType.Workbook, false);
        }

        
        #endregion

        #region Event handlers
        protected internal override void HandleDataReadComplete(object sender, ExportCompleteArgs e)
        {
            doc.Save();
            doc.Dispose();
        }

        protected internal override void HandleProcessedDataRows(object sender, ExportProgressChangedArgs e)
        {
            throw new NotImplementedException();
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

using Microsoft.Office.Interop.Excel;
using System;
using System.Runtime.InteropServices;
using Vovin.CmcLibNet.Database;

namespace Vovin.CmcLibNet.Export
{
    /// <summary>
    /// The purpose of this class is to not have the Excel writer creating it's own copy of XMLWriter,
    /// but simply inherit it.
    /// We do this because events will not be bubbled up from the Excelwriter if we create a second writer.
    /// Internally, the ExcelWriter is just like the XMLWriter anyway
    /// At least until we have created a better implementation
    /// </summary>
    internal class ExcelWriter : XMLWriter
    {
        bool disposed = false;
        readonly string dataFile;
        readonly string xsdFile;
        string fileName;

        #region Constructors
        internal ExcelWriter(ICommenceCursor cursor, IExportSettings settings)
            : base(cursor, settings)
        {
            // create temporary name for XML data file
            // add .xml extension or Excel may complain about trust
            dataFile = System.IO.Path.GetTempFileName() + ".xml";
            // create temporary filename for XSD Schema file
            xsdFile = System.IO.Path.GetTempFileName() + ".xsd";
        }
        ~ExcelWriter()
        {
            Dispose(false);
        }
        #endregion

        protected internal override void WriteOut(string fileName)
        {
            // capture fileName
            this.fileName = fileName;
            PrepareXmlFile(dataFile);
            PrepareXsdFile();
            base.ReadCommenceData();
        }

        private void PrepareXsdFile()
        {
            _settings.XSDCompliant = true; // when importing XML in Excel, data has to be ISO 8601 compliant for it to get the type right
            _settings.IncludeConnectionInfo = true; // XSD follows the fullname - connection - category - field layout pattern
            _settings.XSDFile = xsdFile;
            WriteXMLSchemaFile(xsdFile);
        }

        #region Event handlers
        protected internal override void HandleProcessedDataRows(object sender, ExportProgressChangedArgs e)
        {
            AppendToXml(e.RowValues);
            BubbleUpProgressEvent(e);
        }

        protected internal override void HandleDataReadComplete(object sender, ExportCompleteArgs e)
        {
            CloseXmlFile();
            // now save the data in Excel
            SaveFileToExcel();
            base.BubbleUpCompletedEvent(e);
        }

        private void SaveFileToExcel()
        {
            try
            {
                Microsoft.Office.Interop.Excel.Application xl = new Microsoft.Office.Interop.Excel.Application();
                // this line needs to go immediately after declaring the application or it won't work.
                xl.DisplayAlerts = false; // suppress Excel alerts. Notably the ones displaying that .tmp is untrusted as XML, as well as 'no mapping present in xml'.

                Workbooks wbs = xl.Workbooks; // avoid 2-dot rule thingie just to be sure.
                // TODO: figure this out.
                // possible help: https://msdn.microsoft.com/en-us/library/microsoft.office.tools.excel.workbookbase.xmlimportxml.aspx
                Workbook wb = wbs.OpenXML(dataFile, Type.Missing, XlXmlLoadOption.xlXmlLoadImportToList);
                xl.DisplayAlerts = true;
                if (String.IsNullOrEmpty(fileName))
                {
                    xl.Visible = true;
                }
                else
                {
                    xl.DisplayAlerts = false;
                    wb.SaveAs(Filename: fileName, AccessMode: XlSaveAsAccessMode.xlNoChange, ConflictResolution: XlSaveConflictResolution.xlLocalSessionChanges);
                    wb.Close();
                    xl.DisplayAlerts = true;
                    xl.Quit();
                }

            }
            catch (COMException e)
            {
                //System.Windows.Forms.MessageBox.Show(e.Message);
                string msg = e.Message;
                throw; // rethrow
            }
            finally
            {
                try
                {
                    System.IO.File.Delete(dataFile); // can't refresh dataset when removed
                    System.IO.File.Delete(xsdFile); // fails because it is in use.
                }
                catch { }
            }

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


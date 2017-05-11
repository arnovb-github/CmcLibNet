﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace Vovin.CmcLibNet.Export
{
    // TODO this class works but needs some careful rethinking
    // It would make much more sense to populate Excel from an ADO.NET DataSet, BUT doing so we lose type information.
    // We *want* to preserve type information, it is what sets this assembly apart from other export engines.
    internal class ExcelWriter : BaseWriter
    {
        bool disposed = false;
        XMLWriter _xw = null;

        #region Constructors
        internal ExcelWriter(Database.ICommenceCursor cursor, IExportSettings settings)
            : base(cursor, settings) 
        {
            _xw = new XMLWriter(cursor, settings);
        }

        ~ExcelWriter()
        {
            Dispose(false);
        }
        #endregion

        protected internal override void WriteOut(string fileName)
        {
            // create temporary filename for XSD Schema file
            string xsdFile = System.IO.Path.GetTempFileName() +".xsd";
            _xw._settings.XSDCompliant = true; // when importing XML in Excel, data has to be ISO 8601 compliant for it to get the type right
            _xw._settings.IncludeConnectionInfo = true; // XSD follows the fullname - connection - category - field layout pattern
            _xw._settings.XSDFile = xsdFile;
            _xw.WriteXMLSchemaFile(xsdFile);
            // create temporary name for XML data file
            string dataFile = System.IO.Path.GetTempFileName()+".xml"; // add .xml extension or Excel may complain about trust

            try
            {

                _xw.WriteOut(dataFile);

                // apparently using PIA's doesn't require that Office COM RCWs are released.
                // TODO: it might be a good idea to late-bind Excel;
                // that way we don't have to worry about Office versions.
                // Also, we wouldn't need the reference to the Office PIA (and embed it).
                // This would require the 'dynamic' keyword feature,
                // making coding and debugging a lot harder.
                // For now, we'll just use the PIA.
                // When we call this code from COM Interop, there is no need to release RCWs. Pfew.
                // ideally, this should run on its own thread, so we can start Excel and process Commence data at the same time.
                Microsoft.Office.Interop.Excel.Application xl = new Microsoft.Office.Interop.Excel.Application();
                // this line needs to go immediately after declaring the application or it won't work.
                xl.DisplayAlerts = false; // suppress Excel alerts. Notably the ones displaying that .tmp is untrusted as XML, as well as 'no mapping present in xml'.

                Workbooks wbs = xl.Workbooks; // avoid 2-dot rule thingie just to be sure.
                // using XMLImport fails because it requires a mapping (xsd).
                // using OpenXML works!
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
                // do we want to remove the temp files? I think we do.
                // But I haven't found proper code to uncouple sheet/data yet.
                // Until then, leave them be.
                //System.IO.File.Delete(dataFile); // can't refresh dataset when removed
                //System.IO.File.Delete(xsdFile); // fails because it is in use.
            }
        }

        protected internal override void ProcessDataRows(object sender, DataProgressChangedArgs e)
        {
            //Console.WriteLine("ProcessDataRows in ExcelWriter called");
            return;
        }

        protected internal override void DataReadComplete(object sender, DataReadCompleteArgs e)
        {
            //Console.WriteLine("DataReadComplete in ExcelWriter called");
            return;
        }

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
                if (_xw != null)
                    _xw.Dispose();
            }

            // Free any unmanaged objects here.
            //
            disposed = true;

            // Call the base class implementation.
            base.Dispose(disposing);
        }
    }
}
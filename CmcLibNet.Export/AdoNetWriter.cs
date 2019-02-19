using System.Collections.Generic;
using System.Data;

namespace Vovin.CmcLibNet.Export
{
    internal class AdoNetWriter : BaseWriter
    {
        bool disposed = false;
        DataSet _ds = null;
        string _filename = null;
        int _rows_processed = 0;

        #region Constructors
        internal AdoNetWriter(Database.ICommenceCursor cursor, IExportSettings settings)
            : base(cursor, settings) {}

        #endregion
        protected internal override void WriteOut(string fileName, string sheetName) { }
        protected internal override void WriteOut(string fileName)
        {
            _ds = base.CreateDataSetFromCursorColumns();
            _filename = fileName;
            base._settings.Canonical = true; // TODO fails on large numbers with . or ,
            base._settings.XSDCompliant = false; // TODO when set to true, ADO.NET doesn't get it
            base.ReadCommenceData();
        }

        protected internal override void HandleProcessedDataRows(object sender, ExportProgressChangedArgs e)
        {
            try
            {
                foreach (List<CommenceValue> datarow in e.RowValues)
                {
                    // pass on rowdata for RowParser
                    AdoNetRowWriter rp = new AdoNetRowWriter(_rows_processed, datarow, _ds); // currentrow contains only last row for loop
                    rp.ProcessRow();
                    _rows_processed++;
                }
                BubbleUpProgressEvent(e);
            }
            catch { }
        }

        protected internal override void HandleDataReadComplete(object sender, ExportCompleteArgs e)
        {
            DataSetExporter dse = new DataSetExporter(this._ds, this._filename, base._settings);
            try
            {
                dse.Export();
            }
            catch { }
            base.BubbleUpCompletedEvent(e);
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

            }

            // Free any unmanaged objects here.
            //
            disposed = true;

            // Call the base class implementation.
            base.Dispose(disposing);
        }
    }
}

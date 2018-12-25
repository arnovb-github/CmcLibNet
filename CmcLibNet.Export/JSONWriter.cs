using System.IO;
using Newtonsoft.Json.Linq;

namespace Vovin.CmcLibNet.Export
{
    /// <summary>
    /// Writes to JSON.
    /// </summary>
    // TODO: make this write incremental.
    // The current implementation creates a single JSON object and then writes it to string.
    // On large database, this is a problem.
    internal class JSONWriter : BaseWriter
    {
        private bool disposed = false;
        private StreamWriter _sw = null;
        private JSONCreator _jc = null;

        #region Contructors
        internal JSONWriter(Database.ICommenceCursor cursor, IExportSettings settings)
            : base(cursor, settings) { }

        ~JSONWriter()
        {
            Dispose(false);
        }
        #endregion

        #region Methods

        protected internal override void WriteOut(string fileName)
        {
            _jc = new JSONCreator(this);
            _sw = new StreamWriter(fileName);
            base.ReadCommenceData();
        }

        protected internal override void HandleProcessedDataRows(object sender, ExportProgressChangedArgs e)
        {
            _jc.AppendRowValues(e.RowValues); // we should write to filesteam here for better performance
            BubbleUpProgressEvent(e);
        }

        /// <summary>
        /// Writes the object to file in JSON format.
        /// </summary>
        /// <param name="sender">sender.</param>
        /// <param name="e"><see cref="ExportCompleteArgs"/>.</param>
        protected internal override void HandleDataReadComplete(object sender, ExportCompleteArgs e)
        {
            // ideally, we want to write the file incrementally, so as to avoid memory load.
            // not sure how to do that yet.
            // it could be as simple as adding a bottom and top portion to the file I suppose
            try
            {
                JObject o = _jc.ToJObject(); // may be HUGE. TODO: rethink this.
                _sw.Write(o.ToString());
            }
            finally
            {
                _sw.Flush();
                _sw.Close();
            }
            base.BubbleUpCompletedEvent(e);
        }

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
                if (_sw != null)
                {
                    _sw.Close();
                }
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
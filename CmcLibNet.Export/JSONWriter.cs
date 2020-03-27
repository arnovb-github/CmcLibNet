using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace Vovin.CmcLibNet.Export
{
    /// <summary>
    /// Writes to JSON.
    /// </summary>
    internal class JsonWriter : BaseWriter
    {
        private bool disposed = false;
        private StreamWriter _sw = null;
        private string _fileName;
        private readonly string _tempFile;
        private readonly JsonCreator _jc;

        #region Constructors
        internal JsonWriter(Database.ICommenceCursor cursor, IExportSettings settings)
            : base(cursor, settings)
        {
            _tempFile = Path.GetTempFileName();
            _jc = new JsonCreator(this);
        }

        ~JsonWriter()
        {
            Dispose(false);
        }
        #endregion

        #region Methods
        protected internal override void WriteOut(string fileName)
        {
            if (base.IsFileLocked(new FileInfo(fileName)))
            {
                throw new IOException("File '" + fileName + "' in use.");
            }
            _sw = new StreamWriter(_tempFile); //change to fileName to use 'old' way
            _fileName = fileName;
            base.ReadCommenceData();
        }

        /// <summary>
        /// Writes json to a temporary file
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected internal override void HandleProcessedDataRows(object sender, CursorDataReadProgressChangedArgs e)
        {
            StringBuilder sb = new StringBuilder();
            List<JObject> list = _jc.SerializeRowValues(e.RowValues);
            var jsonString = string.Join(",", list.Select(o => o.ToString()));
            _sw.Write(jsonString);
            BubbleUpProgressEvent(e);
        }

        /// <summary>
        /// Writes the object to file in JSON format.
        /// </summary>
        /// <param name="sender">sender.</param>
        /// <param name="e"><see cref="ExportCompleteArgs"/>.</param>
        protected internal override void HandleDataReadComplete(object sender, ExportCompleteArgs e)
        {
            _sw.Flush();
            _sw.Close();
            using (StreamWriter sw = new StreamWriter(_fileName))
            {
                using (JsonTextWriter jtw = new JsonTextWriter(sw))
                {
                    jtw.WriteStartObject();
                    foreach (var o in _jc.MetaData)
                    {
                        jtw.WritePropertyName(o.Key);
                        jtw.WriteValue(o.Value);
                    }
                    jtw.WritePropertyName("Items");
                    jtw.WriteStartArray();
                    using (StreamReader tr = new StreamReader(_tempFile))
                    {
                        jtw.WriteRaw(tr.ReadToEnd());
                    }
                    jtw.WriteEndArray();
                    jtw.WriteEndObject();
                }
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
                    try
                    {
                        if (_sw.BaseStream != null) // If the StreamWriter is closed, the BaseStream property will return null.
                        {
                            _sw.Flush();
                            _sw.Close();
                        }
                    }
                    catch { }
                    try
                    {
                        File.Delete(_tempFile);
                    }
                    catch { }
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
using System;
using System.Data;
using System.Xml;
using System.IO;
using Newtonsoft.Json;

namespace Vovin.CmcLibNet.Export
{
    /// <summary>
    /// Exports data from a DataSet.
    /// </summary>
    internal class DataSetExporter
    {
        DataSet _ds = null;
        IExportSettings _settings = null;
        readonly string _filename = string.Empty;

        #region Constructors
        internal DataSetExporter(DataSet dataset, string fileName, IExportSettings settings)
        {
            _ds = dataset;
            _settings = settings;
            _filename = fileName;
        }
        #endregion

        #region Methods
        internal void Export()
        {
            switch (_settings.ExportFormat)
            {
                case ExportFormat.Xml:
                    ExportXML(_ds);
                    break;
                case ExportFormat.Json:
                    ExportJson(_ds);
                    break;
            }
        }

        private void ExportXML(DataSet ds)
        {
            
            // interesting fact: even with a schema, Excel can't deal with this!
            XmlWriterSettings xws = new XmlWriterSettings();
            xws.Indent = true;
            using (XmlWriter xw = XmlWriter.Create(_filename, xws))
            {
                _ds.WriteXml(xw, XmlWriteMode.WriteSchema); // write XML schema as well.
            }
        }

        private void ExportJson(DataSet ds)
        {
            // we want a quick and easy way to get nested json.
            XmlDocument doc = new XmlDocument();
            doc.LoadXml(_ds.GetXml());
            // see: http://stackoverflow.com/questions/21727144/convert-dataset-with-multiple-datatables-to-json
            string jsonText = JsonConvert.SerializeObject(_ds, Newtonsoft.Json.Formatting.Indented);
            using (StreamWriter sw = new StreamWriter(_filename))
            {
                sw.Write(jsonText);
            }
        }
        #endregion
    }
}

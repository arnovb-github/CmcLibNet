using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml;

namespace Vovin.CmcLibNet.Export
{
    // Writes data to XML file.
    internal class XmlWriter : BaseWriter
    {
        XmlTextWriter _xtw = null; // the writer object.
        bool disposed = false;

        #region Constructors
        internal XmlWriter(Database.ICommenceCursor cursor, IExportSettings settings)
            : base(cursor, settings){}

        ~XmlWriter()
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
            PrepareXmlFile(fileName);
            base.ReadCommenceData(); // call data reading engine
        }

        private void PrepareXmlFile(string fileName)
        {
            _xtw = new XmlTextWriter(fileName, System.Text.Encoding.UTF8);
            _xtw.Formatting = Formatting.Indented;
            _xtw.WriteStartDocument();
            _xtw.WriteStartElement("dataroot");
        }

        protected internal override void HandleProcessedDataRows(object sender, ExportProgressChangedArgs e)
        {
            AppendToXml(e.RowValues);
            BubbleUpProgressEvent(e);
        }

        private void AppendToXml(List<List<CommenceValue>> rows)
        {
            // populate XMLWriter with data
            foreach (List<CommenceValue> row in rows) // assume that the minimum amount of data is a complete, single Commence item.
            {
                _xtw.WriteStartElement(XmlConvert.EncodeLocalName(_cursor.Category));
                var citems = row.Where(s => s.ColumnDefinition.IsConnection)
                    .OrderBy(o => o.ColumnDefinition.FieldName)
                    .GroupBy(g => g.ColumnDefinition.ColumnName)
                    .ToList();
                foreach (CommenceValue v in row)
                {
                    if (!v.ColumnDefinition.IsConnection) // direct field, i.e. not a connection
                    {
                        // only write if we have something
                        if (!String.IsNullOrEmpty(v.DirectFieldValue))
                        {
                            _xtw.WriteStartElement(XmlConvert.EncodeLocalName(base.ExportHeaders[v.ColumnDefinition.ColumnIndex]));
                            // can we get away with writing the value or do we need to use CData?
                            if (v.ColumnDefinition.CommenceFieldDefinition.MaxChars > 80)
                            {
                                _xtw.WriteCData(v.DirectFieldValue);
                            }
                            else
                            {
                                _xtw.WriteString(v.DirectFieldValue);
                            }
                            _xtw.WriteEndElement();
                        }
                    } // if IsConnection
                } // row
                if (!base._settings.SkipConnectedItems)
                {
                    WriteConnectedItems(citems);
                }
                _xtw.WriteEndElement();
            } // rows
        }

        protected internal override void HandleDataReadComplete(object sender, ExportCompleteArgs e)
        {
            CloseXmlFile();
            base.BubbleUpCompletedEvent(e);
        }

        protected internal void CloseXmlFile()
        {
            try
            {
                // write closing elements and close XMLWriter
                _xtw.WriteEndElement();
                _xtw.WriteEndDocument();
            }
            finally
            {
                _xtw.Flush();
                _xtw.Close();
            }
        }

        private void WriteConnectedItems(IList<IGrouping<string, CommenceValue>> list)
        {
            // a group contains all CommenceValues for a connection
            foreach (IGrouping<string, CommenceValue> group in list)
            {
                string connectionName = XmlConvert.EncodeLocalName(group.FirstOrDefault().ColumnDefinition.QualifiedConnection);
                _xtw.WriteStartElement(connectionName);
                // iterate over the connected value and grab every x-th connected value from the columns
                for (int i = 0; i < group.FirstOrDefault().ConnectedFieldValues?.Length; i++)
                {
                    string categoryName = group.FirstOrDefault().ColumnDefinition.Category;
                    _xtw.WriteStartElement(XmlConvert.EncodeLocalName(categoryName));
                    // now iterate over the different columns
                    for (int j = 0; j < group.Count(); j++)
                    {
                        string value = group.ElementAt(j).ConnectedFieldValues[i];
                        if (!string.IsNullOrEmpty(value))
                        {
                            string fieldName = group.ElementAt(j).ColumnDefinition.FieldName;
                            _xtw.WriteStartElement(XmlConvert.EncodeLocalName(fieldName));
                            // are we dealing with a large text field?
                            if (group.ElementAt(j).ColumnDefinition.CommenceFieldDefinition.MaxChars > 80)
                            {
                                _xtw.WriteCData(value);
                            }
                            else
                            {
                                _xtw.WriteString(value);
                            }
                            _xtw.WriteEndElement();
                        }
                    }
                    _xtw.WriteEndElement();
                }
                _xtw.WriteEndElement();
            }
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
                if (_xtw != null && _xtw.WriteState != WriteState.Closed)
                {
                    _xtw.Flush();
                    _xtw.Close();
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

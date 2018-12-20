using System;
using System.Xml;
using System.Xml.Schema;
using System.Collections.Generic;

namespace Vovin.CmcLibNet.Export
{
    // Writes data to XML file.
    // An important difference with the JSON writer is that this class writes to disk immediately,
    // whereas the JSON writer first aggregates all data.
    // That means that memory use for this class is considerably lower.
    internal class XMLWriter : BaseWriter
    {
        protected internal XmlWriter _xw = null; // the writer object.
        protected internal readonly string _defaultNS = "http://cmclibnet.vovin.nl/export";
        bool disposed = false;

        #region Constructors
        internal XMLWriter(Database.ICommenceCursor cursor, IExportSettings settings)
            : base(cursor, settings){}

        ~XMLWriter()
        {
            Dispose(false);
        }
        #endregion

        #region Methods
        protected internal override void WriteOut(string fileName)
        {
            PrepareXmlFile(fileName);
            base.ReadCommenceData(); // call data reading engine
        }

        protected internal void PrepareXmlFile(string fileName)
        {
            // create a new XMLWriterSettings and some starting elements
            // note that the state of the writer is left open.
            XmlWriterSettings xws = new XmlWriterSettings();
            xws.Indent = true;
            _xw = XmlWriter.Create(fileName, xws);
            _xw.WriteStartDocument();
            if (!String.IsNullOrEmpty(base._settings.XSDFile)) // include xsd declaration
            {
                // we need to load in our XSD file to get some settings from it
                XmlTextReader reader = new XmlTextReader(base._settings.XSDFile);
                XmlSchema myschema = XmlSchema.Read(reader, ValidationCallback);  // no idea what this does
                _xw.WriteStartElement(XmlConvert.EncodeLocalName(base._dataSourceName), myschema.TargetNamespace);
                _xw.WriteAttributeString("xsi", "schemaLocation", "http://www.w3.org/2001/XMLSchema-instance", myschema.TargetNamespace + " " + base._settings.XSDFile);
            }
            else
            {
                // node names have to be properly encoded or else Write*Element may complain.
                _xw.WriteStartElement(XmlConvert.EncodeLocalName(_cursor.Category)); // <-- Important: root element
            }
        }

        protected internal override void HandleProcessedDataRows(object sender, ExportProgressChangedArgs e)
        {
            AppendToXml(e.RowValues);
            BubbleUpProgressEvent(e);
        }

        protected internal void AppendToXml(List<List<CommenceValue>> rows)
        {
            // populate XMLWriter with data
            foreach (List<CommenceValue> row in rows) // assume that the minimum amount of data is a complete, single Commence item.
            {
                _xw.WriteStartElement("Item");
                foreach (CommenceValue v in row)
                {
                    if (v.ColumnDefinition.IsConnection) // connection
                    {
                        WriteConnectedValue(v);
                    }
                    else // direct field
                    {
                        // only write if we have something
                        if (!String.IsNullOrEmpty(v.DirectFieldValue))
                        {
                            _xw.WriteStartElement(XmlConvert.EncodeLocalName(base.ExportHeaders[v.ColumnDefinition.ColumnIndex])); // not good. there shouldn't be a separate array for the headers, they should be in columndefinition.
                            _xw.WriteString(v.DirectFieldValue);
                            _xw.WriteEndElement();
                        }
                    } // if IsConnection
                } // row
                _xw.WriteEndElement();
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
                _xw.WriteEndElement();
                _xw.WriteEndDocument();
            }
            finally
            {
                _xw.Flush();
                _xw.Close();
            }
        }

        // callback needed to read XSD. We don't use it.
        protected internal static void ValidationCallback(object sender, ValidationEventArgs args)
        {
            if (args.Severity == XmlSeverityType.Warning)
                Console.Write("WARNING: ");
            else if (args.Severity == XmlSeverityType.Error)
                Console.Write("ERROR: ");

            Console.WriteLine(args.Message);
        }

        protected internal void WriteConnectedValue(CommenceValue v)
        {
            if (base._settings.SkipConnectedItems || v.ConnectedFieldValues == null) { return; }

            if (base._settings.IncludeConnectionInfo)
            {
                _xw.WriteStartElement(XmlConvert.EncodeLocalName(base.ExportHeaders[v.ColumnDefinition.ColumnIndex]));
                _xw.WriteStartElement(XmlConvert.EncodeLocalName(v.ColumnDefinition.Connection));
                _xw.WriteStartElement(XmlConvert.EncodeLocalName(v.ColumnDefinition.Category));
                foreach (string s in v.ConnectedFieldValues)
                {
                    _xw.WriteStartElement(XmlConvert.EncodeLocalName(v.ColumnDefinition.FieldName));
                    _xw.WriteString(s);
                    _xw.WriteEndElement();
                } //foreach s
                _xw.WriteEndElement();
                _xw.WriteEndElement();
                _xw.WriteEndElement();
            }
            else
            {
                foreach (string s in v.ConnectedFieldValues)
                {
                    _xw.WriteStartElement(XmlConvert.EncodeLocalName(base.ExportHeaders[v.ColumnDefinition.ColumnIndex]));
                    _xw.WriteString(s);
                    _xw.WriteEndElement();
                } //foreach s
            }
        }

        internal void WriteXMLSchemaFile(string fileName)
        {
            string ns = _defaultNS + "/" + XmlConvert.EncodeLocalName(_cursor.Category);
            XmlSchema xsd = new XmlSchema();
            xsd.Namespaces.Add("cmc", ns);
            xsd.TargetNamespace = ns;
            xsd.ElementFormDefault = XmlSchemaForm.Qualified;

            // root
            XmlSchemaElement rel = new XmlSchemaElement();
            XmlSchemaComplexType rcomplexType = new XmlSchemaComplexType();
            rel.Name = XmlConvert.EncodeLocalName(base._dataSourceName);
            rel.SchemaType = rcomplexType;
            XmlSchemaSequence eseq = new XmlSchemaSequence();
            rcomplexType.Particle = eseq;

            // item
            XmlSchemaElement iel = new XmlSchemaElement();
            XmlSchemaComplexType icomplexType = new XmlSchemaComplexType();
            iel.Name = "Item";
            iel.SchemaType = icomplexType;
            iel.MinOccurs = 0;
            iel.MaxOccursString = "unbounded";
            XmlSchemaSequence iseq = new XmlSchemaSequence();
            icomplexType.Particle = iseq;
            // now we have to add our fields
            foreach (ColumnDefinition cd in base.ColumnDefinitions)
            {
                // skip connected columns
                if (!(cd.IsConnection && base._settings.SkipConnectedItems))
                {
                    iseq.Items.Add(cd.XmlSchemaElement);
                }
            }
            eseq.Items.Add(iel);
            xsd.Items.Add(rel);
            XmlWriterSettings xws = new XmlWriterSettings();
            xws.Indent = true;
            using (XmlWriter writer = XmlWriter.Create(fileName, xws))
            {
                xsd.Write(writer);
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
                if (_xw != null)
                {
                    _xw.Close();
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

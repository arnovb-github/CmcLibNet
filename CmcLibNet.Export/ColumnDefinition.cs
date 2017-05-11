﻿using Vovin.CmcLibNet.Database;
using System.Xml;
using System.Xml.Schema;

namespace Vovin.CmcLibNet.Export
{
    /// <summary>
    /// Holds the properties of a column in a cursor.
    /// </summary>
    internal class ColumnDefinition
    {
        /* The columndefinition is fetched once for every cursor we read. */

        private CommenceFieldType _fieldType = CommenceFieldType.Text;
        private bool _fieldTypeFetched;
        private readonly int _colIndex = 0;

        #region Constructors
        internal ColumnDefinition(int colindex, string columnName)
        {
            _colIndex = colindex;
            this.ColumnName = columnName;
        }
        #endregion

        #region Properties
        internal bool IsConnection { get; private set; }
        internal string Category { get; set; }
        internal string Connection { get; private set; }
        internal string ColumnName { get; set; }
        internal string FieldName { get; set; }
        internal string ColumnLabel { get; set; }
        internal string CustomColumnLabel { get; set; }
        internal string Delimiter { get; private set; }

        /// <summary>
        /// Contains detailed information on the related column.
        /// </summary>
        internal IRelatedColumn RelatedColumn
        { 
            set
            {
                this.Category = value.Category;
                this.FieldName = value.Field;
                this.Connection = value.Connection;
                this.Delimiter = value.Delimiter;
                this.IsConnection = true;
            }
        }

        internal XmlSchemaElement XmlSchemaElement 
        { 
            get
            {
                return GetXmlSchemaElementForColumn();
            }
        }

        internal CommenceFieldType FieldType
        {
            get
            {
                if (!_fieldTypeFetched)
                {
                    _fieldType = GetFieldType(this.Category, this.FieldName);
                    _fieldTypeFetched = true;
                }
                return _fieldType;
            }
        }

        internal int ColumnIndex
        {
            get { return _colIndex; }
        }

        /// <summary>
        /// Returns the connection name and the category; i.e. "ConnectionName ToCategory"
        /// </summary>
        internal string QualifiedConnection
        {
            get
            {
                if (this.IsConnection)
                {
                    return this.Connection + ' ' + this.Category;
                }
                else
                {
                    return null;
                }
            }
        }

        /// <summary>
        /// Tablename used when creating ADO.NET DataSet.
        /// <remarks>Identifies the name of the table so it can be matched to a table in the dataset.
        /// Depending on the direct and related fields in a cursor,
        /// any number of tables containing any number of fields can be created from a cursor.
        /// We can use this property to identify which table to use.</remarks>
        /// </summary>
        internal string AdoTableName
        {
            get
            {
                if (this.IsConnection)
                    return this.QualifiedConnection;
                else
                    return this.Category;
            }
        }
        #endregion

        #region Methods
        XmlSchemaElement GetXmlSchemaElementForColumn()
        {
            // first we get the type of the field
            XmlSchemaElement elField = new XmlSchemaElement();
            // note that almost all elements are set nillable;
            // we do not check this against non-null field restrictions in the database.
            elField.Name = XmlConvert.EncodeLocalName(this.FieldName);
            elField.IsNillable = true;
            switch (this.FieldType)
            {
                case CommenceFieldType.Name:
                    elField.SchemaTypeName = new System.Xml.XmlQualifiedName("string", "http://www.w3.org/2001/XMLSchema");
                    elField.IsNillable = false;
                    break;
                case CommenceFieldType.Text:
                case CommenceFieldType.Telephone:
                case CommenceFieldType.Datafile:
                case CommenceFieldType.ExcelCell:
                case CommenceFieldType.Selection:
                case CommenceFieldType.Email:
                case CommenceFieldType.URL:
                    elField.SchemaTypeName = new System.Xml.XmlQualifiedName("string", "http://www.w3.org/2001/XMLSchema");
                    break;
                case CommenceFieldType.Number: // note that Commence may return a currency symbol, but there is no way to find out in advance.
                case CommenceFieldType.Calculation:
                    elField.SchemaTypeName = new System.Xml.XmlQualifiedName("decimal", "http://www.w3.org/2001/XMLSchema");
                    break;
                case CommenceFieldType.Sequence:
                    elField.SchemaTypeName = new System.Xml.XmlQualifiedName("integer", "http://www.w3.org/2001/XMLSchema");
                    elField.IsNillable = false;
                    break;
                case CommenceFieldType.Checkbox:
                    elField.SchemaTypeName = new System.Xml.XmlQualifiedName("boolean", "http://www.w3.org/2001/XMLSchema");
                    break;
                case CommenceFieldType.Date:
                    elField.SchemaTypeName = new System.Xml.XmlQualifiedName("date", "http://www.w3.org/2001/XMLSchema");
                    break;
            }
            if (this.IsConnection) // if we are a connection, wrap field in elements
            {
                // when dealing with connections, additional nodes are created to ensure uniqueness.
                // <fullnodename />
                // <connectionName />
                // <connectedCategoryname/>
                // <connectedFieldName /> <- this is the sequence

                // category name
                XmlSchemaElement catNode = new XmlSchemaElement();
                XmlSchemaComplexType catComplexType = new XmlSchemaComplexType();
                catNode.Name = XmlConvert.EncodeLocalName(this.Category);
                catNode.SchemaType = catComplexType;
                XmlSchemaSequence catSeq = new XmlSchemaSequence();
                catComplexType.Particle = catSeq;
                elField.IsNillable = true;
                elField.MinOccurs = 0; // we can have no connections.
                elField.MaxOccursString = "unbounded";
                catSeq.Items.Add(elField);
                //return catNode; // abort and return early. DEBUG

                // connection name
                XmlSchemaElement conNode = new XmlSchemaElement();
                XmlSchemaComplexType conComplexType = new XmlSchemaComplexType();
                conNode.Name = XmlConvert.EncodeLocalName(this.Connection);
                conNode.SchemaType = conComplexType;
                XmlSchemaSequence conSeq = new XmlSchemaSequence();
                conComplexType.Particle = conSeq;
                conSeq.Items.Add(catNode);

                // raw field name (i.e. connection%%category%%field)
                XmlSchemaElement rawNode = new XmlSchemaElement();
                XmlSchemaComplexType rawComplexType = new XmlSchemaComplexType();
                XmlSchemaSequence rawSeq = new XmlSchemaSequence();
                rawNode.Name = XmlConvert.EncodeLocalName(this.ColumnName);
                //rawNode.MinOccurs = 0;
                //rawNode.MaxOccurs = 1;
                rawNode.SchemaType = rawComplexType;
                rawComplexType.Particle = rawSeq;
                rawSeq.Items.Add(conNode);
                return rawNode;
            }
            else
            {
                return elField;
            }
        }

        private CommenceFieldType GetFieldType(string categoryName, string fieldName)
        {
            CommenceFieldType retval = CommenceFieldType.Text; // default to text
            try
            {
                ICommenceDatabase db = new CommenceDatabase();
                IFieldDef fd = db.GetFieldDefinition(categoryName, fieldName);
                retval = fd.Type;
            }
            catch { } // ignore all errors
            return retval;
        }
        #endregion
    }
}
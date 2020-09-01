using System.Data;
using System.Data.OleDb;
using Vovin.CmcLibNet.Database;
using Vovin.CmcLibNet.Database.Metadata;
using Vovin.CmcLibNet.Extensions;

namespace Vovin.CmcLibNet.Export
{
    /// <summary>
    /// Holds the properties of a column in a cursor.
    /// </summary>
    internal class ColumnDefinition
    {
        /* The columndefinition is fetched once for every cursor we read. */
        private ICommenceFieldDefinition _fieldDefinition;
        private bool _fieldDefinitionFetched;
        ICommenceDatabase _db;

        #region Constructors
        // passing in the ICommenceDatabase object solved the problem of having to create a new one
        // for every field we examine. It may also mask a deeper problem
        internal ColumnDefinition(ICommenceDatabase db, int colindex, string columnName)
        {
            _db = db; // we want this reference because when we would new it up when needed, we unneccesarily open/close DDE channels.
            ColumnIndex = colindex;
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

        internal DbType DbType => this.CommenceFieldDefinition.Type.GetDbTypeForCommenceField();
        internal OleDbType OleDbType => this.CommenceFieldDefinition.Type.GetOleDbTypeForCommenceField();

        internal ICommenceFieldDefinition CommenceFieldDefinition
        {
            get
            {
                if (_fieldDefinition == null && !_fieldDefinitionFetched)
                {
                    _fieldDefinition = _db.GetFieldDefinition(this.Category, this.FieldName);
                    _fieldDefinitionFetched = true; // only to prevent this call again if it returns null
                }
                return _fieldDefinition;
            }
            set
            {
                _fieldDefinition = value; // introduced with ComplexWriter, allows for manual override
            }
        }

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

        internal int ColumnIndex { get; } = 0;

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

        internal static string ThidIdentifier { get; } = "THID"; // TODO is this a good place for this? It is a little obscure 
        #endregion

    }
}
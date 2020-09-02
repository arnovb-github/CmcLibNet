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
        #region Constructors
        internal ColumnDefinition(int colindex, string columnName)
        {
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

        internal ICommenceFieldDefinition CommenceFieldDefinition { get; set; }

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

        internal int ColumnIndex { get; }

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

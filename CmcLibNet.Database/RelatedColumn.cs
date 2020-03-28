using System;

namespace Vovin.CmcLibNet.Database
{
    /// <summary>
    /// POCO class for holding related columns. Created in <c>CommenceCursor.Columns.AddRelatedColum</c>.
    /// </summary>
    internal class RelatedColumn : IRelatedColumn, IEquatable<RelatedColumn> 
    {
        #region Constructors
        internal RelatedColumn(string connection, string category, string field, RelatedColumnType type, string delimiter)
        {
            Connection = connection;
            Category = category;
            Field = field;
            ColumnType = type;
            Delimiter = delimiter;
        }
        #endregion

        #region Properties
        /// <inheritdoc />
        public string Connection { get; set; }
        /// <inheritdoc />
        public string Category { get; set; }
        /// <inheritdoc />
        public string Field { get; set; }
        /// <inheritdoc />
        public RelatedColumnType ColumnType { get; set; }
        /// <inheritdoc />
        public string Delimiter { get; set; }
        #endregion

        #region Methods
        public bool Equals(RelatedColumn other)
        {
            // Check whether the compared object is null.
            if (ReferenceEquals(other, null)) return false;

            // Check whether the compared object references the same data.
            if (ReferenceEquals(this, other)) return true;

            // Check whether the objects’ properties are equal.
            return Connection.Equals(other.Connection) &&
                   Category.Equals(other.Category) && Field.Equals(other.Field);
        }

        public override bool Equals(object obj)
        {
            if (obj == null)
                return false;

            if (!(obj is RelatedColumn rc))
                return false;
            else
                return Equals(rc);
        }

        // If Equals returns true for a pair of objects,
        // GetHashCode must return the same value for these objects.
        public override int GetHashCode()
        {

            // Get the hash code for the Connection field if it is not null.
            int hashConnection = Connection == null ? 0 : Connection.GetHashCode();

            // Get the hash code for the Category field if it is not null.
            int hashCategory = Category == null ? 0 : Category.GetHashCode();

            // Get the hash code for the Field field if it is not null.
            int hashField = Field == null ? 0 : Field.GetHashCode();

            // Calculate the hash code for the object.
            return hashConnection ^ hashCategory ^ hashField;
        }
        #endregion
    }
}

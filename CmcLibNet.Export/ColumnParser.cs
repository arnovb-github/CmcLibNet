using System;
using System.Collections.Generic;
using System.Linq;
using Vovin.CmcLibNet.Database;
using Vovin.CmcLibNet.Database.Metadata;

namespace Vovin.CmcLibNet.Export
{
    /// <summary>
    /// Captures information on the fields and columns of a cursor.
    /// </summary>
    internal class ColumnParser
    {
        /// <summary>
        /// Delimiter Commence returns in columnlabels when requested from a view.
        /// </summary>
        private readonly string connDelim = "%%";

        /// <summary>
        /// Collection of <see cref="ColumnDefinition"/> objects that hold information on a column.
        /// </summary>
        private List<ColumnDefinition> _columnDefinitions = null;

        /// <summary>
        /// Internal dictionary holding the name fields for (connected) categories.
        /// </summary>
        private Dictionary<string, string> _connectedNameFields = null;
        /// <summary>
        /// Holds connection information of the category.
        /// </summary>
        private IEnumerable<ICommenceConnection> _connNames = null;
        /// <summary>
        /// Captures incoming list of custom headers, if any.
        /// </summary>
        private readonly string[] _customHeaders = null;
        /// <summary>
        /// Cursor object.
        /// </summary>
        private Database.ICommenceCursor _cursor = null;

        #region Constructors
        internal ColumnParser(Database.ICommenceCursor cur)
        {
            _cursor = cur;
        }

        internal ColumnParser(Database.ICommenceCursor cur, string[] customHeaders)
        {
            _cursor = cur;
            _customHeaders = customHeaders;
        }
        #endregion

        #region Methods
        /// <summary>
        /// Gets all the field and column information from Commence.
        /// </summary>
        protected internal List<ColumnDefinition> ParseColumns()
        {
            _columnDefinitions = new List<ColumnDefinition>();
            using (ICommenceDatabase db = new CommenceDatabase())
            {
                // can't use using here, because we would close the database prematurely and lose our cursor. Not sure why that happens, it is a new reference?.
                _connNames = db.GetConnectionNames(_cursor.Category); // retrieve all connections for current category. Used to check columns against.

                if (_connNames != null) // there are connections
                {
                    // retrieve name field names from connections
                    if (_connectedNameFields == null) { _connectedNameFields = GetNameFieldsFromConnectedCategories(_connNames); }
                }

                // inject extra columndefintion for thid
                // it should always be the first definition!
                // this is a little tricky
                if (((CommenceCursor)_cursor).Flags.HasFlag(CmcOptionFlags.UseThids))
                {
                    ColumnDefinition cd = new ColumnDefinition(0, ColumnDefinition.ThidIdentifier)
                    {
                        FieldName = ColumnDefinition.ThidIdentifier,
                        CustomColumnLabel = ColumnDefinition.ThidIdentifier,
                        ColumnLabel = ColumnDefinition.ThidIdentifier,
                        Category = _cursor.Category,
                        CommenceFieldDefinition = new CommenceFieldDefinition() // provide empty definition to prevent DDEException on GetFieldDefinition
                    };
                    _columnDefinitions.Add(cd);
                }

                // process actual columns
                using (CmcLibNet.Database.ICommenceQueryRowSet qrs = _cursor.GetQueryRowSet(0))
                {
                    // create a rowset of 0 items
                    for (int i = 0; i < qrs.ColumnCount; i++)
                    {
                        ColumnDefinition cd = new ColumnDefinition(_columnDefinitions.Count, qrs.GetColumnLabel(i, CmcOptionFlags.Fieldname));
                        cd.ColumnLabel = qrs.GetColumnLabel(i);
                        if (this._customHeaders != null)
                        {
                            cd.CustomColumnLabel = this._customHeaders[i];
                        }

                        if (ColumnIsConnection(cd.ColumnName)) // we have a connection
                        {
                            IRelatedColumn rc = GetRelatedColumn(cd.ColumnName);
                            cd.RelatedColumn = rc;
                            if (((CommenceCursor)_cursor).Flags.HasFlag(CmcOptionFlags.UseThids))
                            {
                                cd.CommenceFieldDefinition = new CommenceFieldDefinition()
                                {
                                    MaxChars = CommenceLimits.MaxNameFieldCapacity,
                                    Type = CommenceFieldType.Text
                                };
                            }
                            else
                            {
                                cd.CommenceFieldDefinition = db.GetFieldDefinition(rc.Category, rc.Field);
                            }
                        }
                        else // we have a direct field
                        {
                            cd.Category = _cursor.Category;
                            cd.FieldName = cd.ColumnName;
                            cd.CommenceFieldDefinition = db.GetFieldDefinition(cd.Category, cd.FieldName);
                        }
                        _columnDefinitions.Add(cd);
                    }
                }
            }
            return _columnDefinitions;
        }

        /// <summary>
        /// Returns true if column represents a connection.
        /// </summary>
        /// <param name="fieldName">Fieldname to be evaluated.</param>
        /// <returns>true is field is a connection, otherwise false.</returns>
        private bool ColumnIsConnection(string fieldName)
        {
            // only fieldnames returned from a cursortype of View get the delimiter "%%"
            // if it is present, the format of the fieldName is "Connection%%ToCategory%%ToField",
            // so we expect 2 connDelim.
            // We do not otherwise check for field- or categorynames that have embedded "%%".
            // If you have a database with %% in a field- or categoryname, your database design is flawed :)
            if (fieldName.Contains(connDelim) &&
                fieldName.Split(new string[] { connDelim }, StringSplitOptions.None).Length == 3)
            {
                return true;
            }

            /* If the fieldnames are obtained from a cursor of type Category, a different approach is needed.
             * In this case, we actually have to check against the connection names
             */
            // the only other way for a fieldname to be a connection is to have at least 1 space in it
            if (!fieldName.Contains(' ')) // no space means cannot be a connection
            {
                return false;
            }
            else // we *may* have a connection, but if so, which one?
            {
                return ColumnNameMatchesConnectionName(fieldName, _connNames);
            }
        }

        private bool ColumnNameMatchesConnectionName(string fieldName, IEnumerable<ICommenceConnection> connNames)
        {
            if (_connNames == null) { return false; }

            foreach (ICommenceConnection t in connNames)
            {
                if (fieldName.StartsWith(t.Name) && fieldName.EndsWith(t.ToCategory)
                    && fieldName.Length == t.Name.Length + t.ToCategory.Length + 1)
                {
                    return true; // return true on first match. 
                }
            }
            return false;
        }

        private Dictionary<string, string> GetNameFieldsFromConnectedCategories(IEnumerable<ICommenceConnection> connNames)
        {
            Dictionary<string, string> retval = null;
            if (connNames == null) { return retval; }

            ICommenceDatabase db = new Database.CommenceDatabase();
            retval = new Dictionary<string, string>();
            // collect list of connected category names
            List<string> cats = new List<string>();
            foreach (CommenceConnection c in connNames)
            {
                cats.Add(c.ToCategory);
            }
            // process only unique category names
            cats = cats.Distinct().ToList<string>();
            foreach (string cat in cats)
            {
                retval.Add(cat, db.GetNameField(cat));
            }
            return retval;
        }

        private RelatedColumn GetRelatedColumn(string connectedColumn)
        {
            RelatedColumn retval = null;
            /* Depending on the type of cursor,
             * the fieldname of a connected column may have different formats
             * It can be "Connection ToCategory", with a space
             * or it can be "Connection%%ToCategory%%Fieldname"
             * 
             * The delimiter is different for different viewtypes, so we have to take that into account as well.
             */
            string delim;
            if (connectedColumn.Contains(connDelim))
            {
                // we can split on the delimiter, the last element will be the fieldname in the connected category
                string[] s = connectedColumn.Split(new string[] { connDelim }, StringSplitOptions.None);
                delim = GetDelimiter(RelatedColumnType.ConnectedField);
                retval = new RelatedColumn(s[0], s[1], s[2], RelatedColumnType.ConnectedField, delim);
            }
            if (connectedColumn.Contains(' ')) // this can be shorter
                foreach (CommenceConnection c in _connNames)
                {
                    if (connectedColumn.StartsWith(c.Name) && connectedColumn.EndsWith(c.ToCategory)
                        && connectedColumn.Length == c.Name.Length + c.ToCategory.Length + 1)
                    {
                        delim = GetDelimiter(RelatedColumnType.Connection);
                        // when a connection is assigned to a direct field and the UseThids flag is true,
                        // we run into a special situation: you can create a cursor that has duplicate 'fieldnames'.
                        // example from the Tutorial database:
                        // cursor.Columns.AddDirectColumns("accountKey", "Relates to Contact") // direct field, but in fact a connection
                        // cursor.Columns.AddRelatedColumn("Relates to", "Contact", "contactKey") // related field
                        // this would end us up with two 'contactKey' nodes
                        // so in that case, do not use the actual fieldname
                        if (((CommenceCursor)_cursor).Flags.HasFlag(CmcOptionFlags.UseThids))
                        {
                            retval = new RelatedColumn(c.Name, c.ToCategory, ColumnDefinition.ThidIdentifier, RelatedColumnType.Connection, delim);
                        }
                        else
                        {
                            retval = new RelatedColumn(c.Name, c.ToCategory, this._connectedNameFields[c.ToCategory], RelatedColumnType.Connection, delim);
                        }
                    }
                }
            return retval;
        }

        /// <summary>
        /// Returns the delimiter Commence will use when returning column data.
        /// </summary>
        /// <param name="rct">RelatedColumnType enum value.</param>
        /// <returns>delimiter.</returns>
        // simplified implementation
        private string GetDelimiter(RelatedColumnType rct)
        {
            /* Delimiter depends on cursor type and underlying view (if cursor is on view)
            * Luckily, changing flags doesn't alter this behaviour, it is complex enough as it is.
            *
            * If you would redefine the columns, you get different results.
            * For example:
            * If you do SetColumn(n, "ConnectionName CategoryName, 0) 
            * then that column will be comma-separated
            * If you do SetRelatedColumn(n, "ConnectionName", "CategoryName", "FieldName", 0), 
            * then that column will be newline-separated
            * (i.o.w., it will then act just like a Category type cursor)
            * That means that within the same View, you can have different delimiters.
            * Commence is such a mess in that respect :(
            */
            switch (rct)
            {
                case RelatedColumnType.Connection:
                    return ", "; // note the trailing space
                case RelatedColumnType.ConnectedField:
                    return "\n";
                default:
                    return "\n";
            }
        }

        #endregion
    }
}

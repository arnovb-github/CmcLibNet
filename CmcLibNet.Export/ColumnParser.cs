﻿using System;
using System.Collections.Generic;
using System.Linq;
using Vovin.CmcLibNet.Database;

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
        private const string connDelim = "%%";

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
        private List<Tuple<string, string>> _connNames = null;
        /// <summary>
        /// Captures incoming list of custom headers, if any.
        /// </summary>
        private string[] _customHeaders = null;
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
            Database.ICommenceDatabase db = new Database.CommenceDatabase();
            _connNames = db.GetConnectionNames(_cursor.Category); // retrieve all connections for current category. Used to check columns against.
            _columnDefinitions = new List<ColumnDefinition>();

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
                ColumnDefinition cd = new ColumnDefinition(db, 0, "thid");
                cd.FieldName = "thid";
                cd.CustomColumnLabel = "thid";
                cd.ColumnLabel = "thid";
                cd.Category = _cursor.Category;
                _columnDefinitions.Add(cd);
            }

            // process actual columns
            using (CmcLibNet.Database.ICommenceQueryRowSet qrs = _cursor.GetQueryRowSet(0))
            {
                // create a rowset of 0 items
                for (int i = 0; i < qrs.ColumnCount; i++)
                {
                    ColumnDefinition cd = new ColumnDefinition(db, _columnDefinitions.Count, qrs.GetColumnLabel(i, CmcOptionFlags.Fieldname)); // thids
                    _columnDefinitions.Add(cd);
                    cd.ColumnLabel = qrs.GetColumnLabel(i);
                    if (this._customHeaders != null)
                    {
                        cd.CustomColumnLabel = this._customHeaders[i];
                    }

                    if (ColumnIsConnection(cd.ColumnName)) // we have a connection
                    {
                        cd.RelatedColumn = GetRelatedColumn(cd.ColumnName);
                    }
                    else // we have a direct field
                    {
                        cd.Category = _cursor.Category;
                        cd.FieldName = cd.ColumnName;
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
                return ColumnNameMatchesConnectionName(fieldName, this._connNames);
            }
        }

        private bool ColumnNameMatchesConnectionName(string fieldName, List<Tuple<string, string>> connNames)
        {
            if (_connNames == null) { return false; }

            foreach (Tuple<string, string> t in connNames)
            {
                if (fieldName.StartsWith(t.Item1) && fieldName.EndsWith(t.Item2)
                    && fieldName.Length == t.Item1.Length + t.Item2.Length + 1)
                {
                    return true; // return true on first match. 
                }
            }
            return false;
        }

        private Dictionary<string, string> GetNameFieldsFromConnectedCategories(List<Tuple<string, string>> connNames)
        {
            Dictionary<string, string> retval = null;
            if (connNames == null) { return retval; }

            Database.ICommenceDatabase db = new Database.CommenceDatabase();
            retval = new Dictionary<string, string>();
            // collect list of connected category names
            List<string> cats = new List<string>();
            foreach (Tuple<string, string> t in connNames)
            {
                cats.Add(t.Item2);
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
            string delim = string.Empty;
            if (connectedColumn.Contains(connDelim))
            {
                // we can split on the delimiter, the last element will be the fieldname in the connected category
                string[] s = connectedColumn.Split(new string[] { connDelim }, StringSplitOptions.None);
                delim = GetDelimiter((CommenceCursor)_cursor, RelatedColumnType.ConnectedField);
                retval = new RelatedColumn(s[0], s[1], s[2], RelatedColumnType.ConnectedField, delim);
            }
            if (connectedColumn.Contains(' ')) // this can be shorter
                foreach (Tuple<string, string> t in this._connNames)
                {
                    if (connectedColumn.StartsWith(t.Item1) && connectedColumn.EndsWith(t.Item2)
                        && connectedColumn.Length == t.Item1.Length + t.Item2.Length + 1)
                    {
                        delim = GetDelimiter((CommenceCursor)_cursor, RelatedColumnType.Connection);
                        retval = new RelatedColumn(t.Item1, t.Item2, this._connectedNameFields[t.Item2], RelatedColumnType.Connection, delim);
                    }
                }
            return retval;
        }
        /// <summary>
        /// Returns the delimiter Commence will use when returning column data.
        /// </summary>
        /// <param name="cur">CommenceCursor.</param>
        /// <param name="rct">RelatedColumnType enum value.</param>
        /// <returns>delimiter.</returns>
        private string GetDelimiter(CommenceCursor cur, RelatedColumnType rct)
        {
            /* Delimiter depends on cursor type and underlying view (if cursor is on view)
             * Luckily, it does not matter if the thids flag is used or not
             */ 
            string retval = "\n";
            // check cursor type
            switch (cur.CursorType)
            {
                case CmcCursorType.Category:
                    
                    switch (rct)
                    {
                        case RelatedColumnType.Connection:
                            retval = ", "; // note the trailing space
                            break;
                        case RelatedColumnType.ConnectedField:
                            retval = "\n";
                            break;
                    }
                    break;

                case CmcCursorType.View:
                    
                    switch (cur.ViewType)
                    {
                        // process only viewtypes on which a cursor can be created
                        case CommenceViewType.Book:
                            retval = ", ";
                            break;
                        case CommenceViewType.Grid:
                        case CommenceViewType.Report:
                            retval = "\n";
                            break;
                    }
                    break;
            }
            return retval;
        }
        #endregion
    }
}

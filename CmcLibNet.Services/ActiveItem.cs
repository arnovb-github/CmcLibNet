using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Vovin.CmcLibNet.Database;
using Vovin.CmcLibNet.Database.Metadata;
using Vovin.CmcLibNet.Export;

namespace Vovin.CmcLibNet.Services
{
    internal class ActiveItem : IDisposable
    {
        private string _itemName = string.Empty;
        private string _nameField = string.Empty;
        private readonly IActiveViewInfo _avi;
        private readonly ICommenceDatabase _db;
        private readonly string clarifyStatus;

        internal ActiveItem()
        {
            _db = new CommenceDatabase();
            _avi = _db.GetActiveViewInfo();
            clarifyStatus = _db.ClarifyItemNames();
            _db.ClarifyItemNames("false");
            _itemName = _db.GetActiveItemName();
            _db.ClarifyItemNames(clarifyStatus);
        }

        internal List<Field> GetValues()
        {
            // first determine what we have. There are 3 scenarios:
            // -no view is active
            // -a view is showing
            // -a detail form is showing

            // then we need to figure out what the best way is to get the item values
            // filter uniquely
            // direct/indirect fields

            // no view is active
            if (_avi == null) { return null; }

            // the view can be empty, or not support getting of data (like the Document view).
            if (_itemName == null) { return null; }

            List<Field> retval = null;
            // we have a categoryname, so we can figure out the Name field name.
            _nameField = _db.GetNameField(_avi.Category);

            switch (_avi.Type.ToLower())
            {
                case "item detail form":
                    // we know the fieldname of this detail form, so we can pass in its XML and find out what fields are showing.
                    // That is, if we can find out a way which form is showing...which we can't.|
                    // okay never mind then
                    return null;
                default:
                    // a view is showing, create a cursor based on it
                    // remember that Commence cannot create a cursor on all types of views.
                    ICommenceCursor cur;
                    try
                    {
                        retval = new List<Field>();
                        cur = _db.GetCursor(_avi.Name, CmcCursorType.View); // may fail
                    }
                    catch (ArgumentException)
                    {
                        Field f = new Field(_nameField, string.Empty, _itemName); 
                        retval.Add(f);
                        return retval;
                    }
                    // when the open view's name has quotes in it, the DDE request to retrieve info on it may fail
                    // in that case we get an argument exception in Utils.EnumFromAttributeValue
                    // that all gets convoluted, let's just swallow all errors and pretend we do not know that :)
                    catch { return retval; }
                    retval = GetValuesFromCursor(cur).ToList();
                    break;
            } // switch
            return retval;
        }

        private IEnumerable<Field> GetValuesFromCursor(ICommenceCursor cur)
        {
            IEnumerable<Field> retval = new List<Field>(); // a cursor cannot have no fields, so we should be safe.
            ColumnParser cp = new ColumnParser(cur);
            IList<ColumnDefinition> columnDefinitions = cp.ParseColumns();
            _db.ClarifyItemNames("true");
            _itemName = _db.GetActiveItemName(); // also marks the item for us

            IList<Field> fields = new List<Field> {
                new Field { 
                    Name = _nameField,
                    Value = _itemName,
                    Label = columnDefinitions
                        .Where(w => !w.IsConnection)
                        .SingleOrDefault(s => s.FieldName.Equals(_nameField))?.ColumnLabel
                }
            };
            // if the name field isn't in the view the view label will be empty, 
            // in that casechange it to fieldname instead
            // note that this will actually return more information than is showing in Commence
            if (string.IsNullOrEmpty(fields[0].Label))
            {
                fields[0].Label = _nameField;
            }

            // get a list of all the direct fields except the name field
            IEnumerable<ColumnDefinition> directFields = columnDefinitions
                .Where(w => !w.IsConnection && w.FieldName != _nameField)
                .ToArray();
            // get the direct field values using DDE
            IEnumerable<Field> directFieldValues = GetDirectFieldValues(directFields);
            retval = fields.Concat(directFieldValues);

            // get the indirect values
            IEnumerable<RelatedColumn> relatedColumns = columnDefinitions
                .Where(w => w.IsConnection)
                .Select(s => new RelatedColumn(s.Connection, 
                    s.Category,
                    s.FieldName, 
                    RelatedColumnType.ConnectedField,
                    s.QualifiedConnection + ' ' + s.FieldName)) // very dirty trick!!!!!!
                .ToArray();
            IEnumerable<Field> connectedFieldValues = GetConnectedFieldValues(relatedColumns);
            retval = retval.Concat(connectedFieldValues);
            _db.ClarifyItemNames(clarifyStatus); // restore original setting
            return retval;
        } // method

        private IEnumerable<Field> GetConnectedFieldValues(IEnumerable<RelatedColumn> relatedColumns)
        {
            foreach (var r in relatedColumns)
            {
                StringBuilder sb = new StringBuilder(r.Connection);
                sb.Append("%%");
                sb.Append(r.Category);
                sb.Append("%%");
                sb.Append(r.Field);
                Field retval = new Field()
                {
                    Name = sb.ToString(),
                    Label = r.Delimiter, // abused the field for something else. Very dirty trick!!!
                };
                try
                {
                    // If the Category and Item parameters are both blank,
                    // then Commence uses the item from the most recent AddItem/AddSharedItem command,
                    // MarkActiveItem or ViewMarkItem REQUEST.
                    // In this case, it is MarkActiveItem, which is what we want
                    int conItems = _db.GetConnectedItemCount(string.Empty, string.Empty, r.Connection, r.Category);
                    // we are going to return connected data.
                    // but only if there is a single connected item
                    // we could retrieve all connected data,
                    // but that could potentially take too long
                    switch (conItems)
                    {
                        case 0:
                            retval.Value = "(none)";
                            break;
                        case 1:
                            retval.Value = _db.GetConnectedItemField(string.Empty, string.Empty, r.Connection, r.Category, r.Field);
                            break;
                        default:
                            retval.Value = "(more)";
                            break;
                    }
                }
                catch { } // swallow all errors
                yield return retval;
            }
        }

        private IEnumerable<Field> GetDirectFieldValues(IEnumerable<ColumnDefinition> directFields)
        {
            
            foreach (var cd in directFields)
            {
                Field retval = null;
                try
                {
                    retval = new Field()
                    {
                        Name = cd.FieldName,
                        Label = cd.ColumnLabel,
                        // If the Category and Item parameters are both blank,
                        // then Commence uses the item from the most recent AddItem/AddSharedItem command,
                        // MarkActiveItem or ViewMarkItem REQUEST.
                        // In this case, it is MarkActiveItem, which is what we want
                        Value = _db.GetField(string.Empty, string.Empty, cd.FieldName)
                    };
                }
                catch { } // ignore all errors
                yield return retval;
            }
        }

        #region IDisposable Support
        private bool disposedValue = false; // To detect redundant calls

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    _db?.Close();
                }
                disposedValue = true;
            }
        }

        // This code added to correctly implement the disposable pattern.
        public void Dispose()
        {
            // Do not change this code. Put cleanup code in Dispose(bool disposing) above.
            Dispose(true);
        }
        #endregion
    }
}
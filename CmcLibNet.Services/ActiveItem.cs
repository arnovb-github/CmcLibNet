using System;
using System.Collections.Generic;
using Vovin.CmcLibNet;
using Vovin.CmcLibNet.Database;
using Vovin.CmcLibNet.Export;

namespace Vovin.CmcLibNet.Services
{
    internal class ActiveItem
    {
        string _itemName = string.Empty;
        string _nameField = string.Empty;
        //ICommenceApp _cmc = null;
        IActiveViewInfo _avi = null;
        ICommenceDatabase _db = null;

        internal ActiveItem()
        {
            //_cmc = new CommenceApp();
            _db = new CommenceDatabase();
            _avi = _db.GetActiveViewInfo();
            _db.ClarifyItemNames("false");
            _itemName = _db.GetActiveItemName();
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
                    _itemName = _avi.Item;
                    // we know the fieldname of this detail form, so we can pass in its XML and find out what fields are showing.
                    // That is, if we can find out a way which form is showing...which we can't.
                    throw new NotSupportedException("You cannot copy data from an Item Detail Form with " + System.Reflection.Assembly.GetExecutingAssembly().GetName().Name + ".");
                default:
                   // a view is showing, create a cursor based on it
                    // remember that Commence cannot create a cursor on all types of views.
                    ICommenceCursor cur = null;
                    try
                    {
                        retval = new List<Field>();
                        cur = _db.GetCursor(_avi.Name, CmcCursorType.View);
                    }
                    catch (NotSupportedException)
                    {
                        Field f = new Field(_nameField, "", _itemName + ". That's all we could get, sorry.");
                        retval.Add(f);
                        return retval;
                    }
                    retval = GetValuesFromCursor(cur);
                    break;
            } // switch
            return retval;
        } // method

        private List<Field> GetValuesFromCursor(ICommenceCursor cur)
        {
            List<Field> retval = new List<Field>(); // a cursor cannot have no fields, so we should be safe.
            // okay, we have a cursor. Ideally we'd like to use it to get all the desired data.
            // At least we can use it to figure out what fields are showing.
            ColumnParser cp = new ColumnParser(cur);
            List<ColumnDefinition> columnInfo = cp.ParseColumns();
            /* Let's try and see if the item is unique.
             * If we get only 1 item, we're in luck.
             * If we don't, we could try and filter for it,
             * but because the view can contain any kind of field, 
             * this would be exceedingly complex.
             *  we are going to try one last time and see if the namefield is unique.
             */
            ICursorFilterTypeF filter = (ICursorFilterTypeF)cur.Filters.Add(1, FilterType.Field);
            filter.FieldName = _nameField; // not pretty
            filter.Qualifier = FilterQualifier.EqualTo;
            filter.MatchCase = true;
            filter.FieldValue = _itemName; // not pretty
            /* notice we could set another filter on the clarify field had we requested one,
             * but because we do not know what type of field that is, that would also get messy very quickly.
             */
            cur.Filters.Apply();
            if (cur.RowCount == 1)
            {
                using (ICommenceQueryRowSet qrs = cur.GetQueryRowSet())
                {
                    for (int i = 0; i < columnInfo.Count; i++)
                    {
                        Field f = new Field();
                        f.Label = columnInfo[i].ColumnLabel;
                        f.Name = columnInfo[i].FieldName;
                        f.Value = qrs.GetRowValue(0, i); // just i? Is that always the correct ColumnIndex?
                        retval.Add(f);
                    } // for (int i = 0; i < columnInfo.FieldNames.Count; i++)
                }
            } // if (cur.RowCount == 1)
            else
            {
                // no other way than to use good old DDE
                // try to use clarified itemname
                _db.ClarifyItemNames("true");
                _itemName = _db.GetActiveItemName();
                List<string> directFieldList = null;
                List<string> directColumnList = null;
                for (int i = 0; i < columnInfo.Count; i++)
                {
                    if (!columnInfo[i].IsConnection)
                    {
                        // collect all direct fields
                        if (directFieldList == null) { directFieldList = new List<string>(); }
                        // keep track of corresponding column names
                        if (directColumnList == null) { directColumnList = new List<string>(); }
                        directFieldList.Add(columnInfo[i].FieldName);
                        directColumnList.Add(columnInfo[i].ColumnLabel);
                    }
                    else
                    {
                        // we are not going to collect the indirect fields,
                        // instead we'll request them immediately.
                        string[] s = null;
                        /* Book-type views are a special case, since they return the connection fields in it differently from other viewtypes.
                         * Specifically, Book views return (and indeed, only show) connected itemnames.
                         * This is reflected in the way the connected values are returned
                         * Instead of Connection%%Category%%Field, they return Connection<space>Category (no field).
                         * This is a problem when either the connection or category also contain spaces.
                         * Fortunately for us, the connection names are captured in the HeaderLists class
                         */
                        if (_avi.Type.ToLower() == "book")
                        {
                            s = new string[] { columnInfo[i].Connection, columnInfo[i].Category, columnInfo[i].FieldName };
                        }
                        else
                        {
                            s = columnInfo[i].ColumnName.Split(new string[] { "%%" }, StringSplitOptions.None);
                        }
                        if (s.Length == 3)
                        {
                            Field f = new Field();
                            f.Name = columnInfo[i].FieldName;
                            f.Label = columnInfo[i].ColumnLabel;
                            f.Value = _db.GetConnectedItemField(cur.Category, _itemName, s[0], s[1], s[2]);
                            retval.Add(f);
                        } // if (s.Length == 3)
                    } // else (s.Length == 3)
                } // for (int i = 0; i < columnInfo.FieldNames.Count; i++)
                // now process all direct fields in one go
                List<string> directFieldValues = _db.GetFields(cur.Category, _itemName, directFieldList);
                for (int i = 0; i < directFieldValues.Count; i++)
                {
                    Field f = new Field();
                    f.Name = directFieldList[i];
                    f.Label = directColumnList[i];
                    f.Value = directFieldValues[i];
                    retval.Add(f);
                } // for (int i = 0; i < directFieldValues.Count; i++)
            } // else (cur.RowCount == 1)

            return retval;
        } // method
    }
}

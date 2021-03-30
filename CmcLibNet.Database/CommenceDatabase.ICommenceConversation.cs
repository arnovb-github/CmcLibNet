using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Timers;
using Vovin.CmcLibNet.Attributes;
using Vovin.CmcLibNet.Database.Metadata;

namespace Vovin.CmcLibNet.Database
{

    // This portion of the class contains the implementation of the properties and methods exposed by Commence's ICommenceGetconversation interface.
    // That is a difficult way of saying all DDE stuff is in here.
    public partial class CommenceDatabase : ICommenceDatabase
    {
        #region Fields
        private Timer DDETimer;
        private readonly int DDETimeout = 1000; // milliseconds after which a DDE conversation is closed.
        private CommenceConversation _conv;
        #endregion

        #region Commence DDE Request methods
        /// <inheritdoc />
        public string GetActiveCategory()
        {
            // first try to mark the active item
            string result = MarkActiveItem(this.Delim);
            if (result != null) // use the active item
            {
                string[] buffer = result.Split(this.Splitter, StringSplitOptions.None);
                return buffer[0];
            }
            else // no active item, try getting info on the view (if any) instead
            {
                IActiveViewInfo ai = GetActiveViewInfo(); // will not work on all types of views, but at least we tried.
                return ai?.Category; // no view active
            }
        }

        /// <inheritdoc />
        public string GetActiveItemName()
        {
            string result = MarkActiveItem(this.Delim);
            // the view can be empty, or not support getting of data (like the Document view).
            if (result == null) { return null; }
            string[] buffer = result.Split(this.Splitter, StringSplitOptions.None);
            return buffer[1];
        }

        /// <inheritdoc />
        public string ClarifyItemNames(string bStatus = null)
        {
            return (bStatus == null)
                ? DDERequest(BuildDDERequestCommand(new string[] { "ClarifyItemNames" }))
                : DDERequest(BuildDDERequestCommand(new string[] { "ClarifyItemNames", bStatus }));
        }

        /// <inheritdoc />
        public IActiveViewInfo GetActiveViewInfo()
        {
            ActiveViewInfo avi = null;
            string dde = BuildDDERequestCommand(new string[] { "GetActiveViewInfo", this.Delim });
            string viewInfo = DDERequest(dde);
            if (viewInfo != null) // null means no view was active
            {
                // commence will return {ViewName}Delim {ViewType}Delim {CategoryName}Delim {ItemName}Delim {FieldName}
                string[] buffer = viewInfo.Split(this.Splitter, StringSplitOptions.None);
                avi = new ActiveViewInfo
                {
                    Name = buffer[0],
                    Type = (buffer[1] == "Add Item") ? "Item Detail Form" : buffer[1], // translate. This is a discrepancy in Commence documentation.
                    Category = buffer[2],
                    Item = buffer[3]
                };
                if (avi.Type.ToLower() == "item detail form")
                {
                    avi.Field = buffer[4];
                }
                else
                {
                    avi.Field = string.Empty; // maybe overkill but can't do any harm
                }
            }
            return avi;
        }

        /// <inheritdoc />
        public string GetCallerID(string categoryName, string phoneNumber, string delim = null)
        {
            return (delim == null)
                ? GetDDEValues(new string[] { "GetCallerID", categoryName, phoneNumber })
                : GetDDEValues(new string[] { "GetCallerID", categoryName, phoneNumber, delim });
        }

        /// <inheritdoc />
        public List<string> GetCallerID(string categoryName, string phoneNumber)
        {
            return GetDDEValuesAsList(new string[] { "GetCallerID", categoryName, phoneNumber, this.Delim });
        }

        /// <inheritdoc />
        public int GetCategoryCount() // this could be more elegant
        {
            return GetDDECount(new string[] { "GetCategoryCount" });
        }

        /// <inheritdoc />
        public ICategoryDef GetCategoryDefinition(string categoryName)
        {
            CategoryDef cd = new CategoryDef();
            string dde = BuildDDERequestCommand(new string[] { "GetCategoryDefinition", categoryName, this.Delim });
            string categoryInfo = DDERequest(dde);

            if (categoryInfo != null)
            {
                string[] buffer = categoryInfo.Split(this.Splitter, StringSplitOptions.None);
                cd.MaxItems = Convert.ToInt32(buffer[0]);
                string s = buffer[1];
                cd.Shared = s.Substring(6, 1) == "1";
                // note that we skip substring 7 - it has no meaning
                cd.Duplicates = s.Substring(8, 1) == "1";
                cd.Clarified = s.Substring(9, 1) == "1";
                if (cd.Clarified)
                {
                    cd.ClarifySeparator = buffer[2];
                    cd.ClarifyField = buffer[3];
                }
                else
                {
                    cd.ClarifySeparator = string.Empty;
                    cd.ClarifyField = string.Empty;
                }
                cd.CategoryID = this.GetCategoryID(categoryName);
            }
            return cd;
        }

        /// <inheritdoc />
        public string GetCategoryNames(string delim = null)
        {
            return (delim == null)
                ? GetDDEValues(new string[] { "GetCategoryNames" })
                : GetDDEValues(new string[] { "GetCategoryNames", delim });
        }

        /// <inheritdoc />
        public List<string> GetCategoryNames()
        {
            return GetDDEValuesAsList(new string[] { "GetCategoryNames", this.Delim });
        }

        /// <inheritdoc />
        // this method is not a Commence method
        public string GetClarifiedItemName(string itemName, string clarifySeparator, string clarifyValue)
        {
            return Utils.GetClarifiedItemName(itemName, clarifySeparator, clarifyValue);
        }

        /// <inheritdoc />
        public string GetClarifyField(string categoryName)
        {
            ICategoryDef cd = this.GetCategoryDefinition(categoryName);
            return (cd.ClarifyField == string.Empty)
                ? null
                : cd.ClarifyField;
        }

        /// <inheritdoc />
        public string GetClarifySeparator(string categoryName)
        {
            ICategoryDef cd = this.GetCategoryDefinition(categoryName);
            return (cd.ClarifySeparator == String.Empty)
                ? null
                : cd.ClarifySeparator;
        }

        /// <inheritdoc />
        public int GetConnectedItemCount(string categoryName, string itemName, string connectionName, string connCategory)
        {
            return GetDDECount(new string[] { "GetConnectedItemCount", categoryName, itemName, connectionName, connCategory });
        }

        /// <inheritdoc />
        public string GetConnectedItemField(string categoryName, string itemName, string connectionName, string connCategory, string fieldName, string delim = null)
        {
            return (delim == null)
                ? GetDDEValues(new string[] { "GetConnectedItemField", categoryName, itemName, connectionName, connCategory, fieldName })
                : GetDDEValues(new string[] { "GetConnectedItemField", categoryName, itemName, connectionName, connCategory, fieldName, delim });
        }

        /// <inheritdoc />
        public string GetConnectedItemNames(string categoryName, string itemName, string connectionName, string connCategory, string delim = null)
        {
            return (delim == null)
                ? GetDDEValues(new string[] { "GetConnectedItemNames", categoryName, itemName, connectionName, connCategory })
                : GetDDEValues(new string[] { "GetConnectedItemNames", categoryName, itemName, connectionName, connCategory, delim });
        }

        /// <inheritdoc />
        public List<string> GetConnectedItemNames(string categoryName, string itemName, string connectionName, string connCategory)
        {
            return GetDDEValuesAsList(new string[] { "GetConnectedItemNames", categoryName, itemName, connectionName, connCategory, this.Delim });
        }

        /// <inheritdoc />
        public int GetConnectionCount(string categoryName)
        {
            return GetDDECount(new string[] { "GetConnectionCount", categoryName });
        }

        /// <inheritdoc />
        public object GetConnectionNames(string categoryName, string delim1 = null, string delim2 = null)
        {
            return GetConnectionNames(categoryName).Cast<object>().ToArray();
        }

        /// <inheritdoc />
        [ComVisible(false)]
        public IEnumerable<ICommenceConnection> GetConnectionNames(string categoryName)
        {
            IList<ICommenceConnection> retval = new List<ICommenceConnection>();
            string buffer = GetDDEValues(new string[] { "GetConnectionNames", categoryName, this.Delim, this.Delim2 });
            if (!string.IsNullOrEmpty(buffer) && string.IsNullOrEmpty(this.GetLastError())) // this will swallow the error if a DDE error occurred. Hmm.
            {
                string[] pairs = buffer.Split(this.Splitter, StringSplitOptions.None);
                foreach (string p in pairs)
                {
                    string[] pair = p.Split(this.Splitter2, StringSplitOptions.None);
                    retval.Add(new CommenceConnection()
                    {
                        Name = pair[0],
                        ToCategory = pair[1]
                    });
                }
            }
            return retval.ToArray();
        }

        /// <inheritdoc />
        [Obsolete("Use CmcLibNet.CommenceApp.Name and/or CmcLibNet.CommenceApp.Path.")]
        public string GetDatabase()
        {
            return GetDDEValues(new string[] { "GetDatabase" }); 
        }

        /// <inheritdoc />
        public IDBDef GetDatabaseDefinition()
        {
            DBDef db = new DBDef();
            string dde = BuildDDERequestCommand(new string[] { "GetDatabaseDefinition", this.Delim });
            string dbInfo = DDERequest(dde);
            if (dbInfo != null)
            {
                string[] buffer = dbInfo.Split(this.Splitter, StringSplitOptions.None);
                // Commence will return:
                // {DatabaseName}Delim {DatabasePath}Delim 000000{A}{X}{S}{C}Delim {UserName}Delim {SpoolPath}
                db.Name = buffer[0];
                db.Path = buffer[1];
                string s = buffer[2];
                db.Attached = s.Substring(6, 1) == "1";
                db.Connected = s.Substring(7, 1) == "1";
                db.IsServer = s.Substring(8, 1) == "1";
                db.IsClient = s.Substring(9, 1) == "1";
                db.Username = buffer[3];
                db.Spoolpath = buffer[4];
            }
            return db;
        }

        /// <inheritdoc />
        public int GetDesktopCount()
        {
            return GetDDECount(new string[] { "GetDesktopCount" });
        }

        /// <inheritdoc />
        public string GetDesktopNames(string delim = null)
        {
            return (delim == null)
                ? GetDDEValues(new string[] { "GetDesktopNames" })
                : GetDDEValues(new string[] { "GetDesktopNames", delim });
        }

        /// <inheritdoc />
        public List<string> GetDesktopNames()
        {
            return GetDDEValuesAsList(new string[] { "GetDesktopNames", this.Delim });
        }

        /// <inheritdoc />
        public string GetField(string categoryName, string itemName, string fieldName)
        {
            return GetDDEValues(new string[] { "GetField", categoryName, itemName, fieldName });
        }

        /// <inheritdoc />
        public int GetFieldCount(string categoryName)
        {
            return GetDDECount(new string[] { "GetFieldCount", categoryName });
        }

        /// <inheritdoc />
        public ICommenceFieldDefinition GetFieldDefinition(string categoryName, string fieldName)
        {
            CommenceFieldDefinition fd = new CommenceFieldDefinition();
            string dde = BuildDDERequestCommand(new string[] { "GetFieldDefinition", categoryName, fieldName, this.Delim });
            string fieldInfo = DDERequest(dde);
            if (fieldInfo != null)
            {
                string[] buffer = fieldInfo.Split(this.Splitter, StringSplitOptions.None);
                fd.Type = (CommenceFieldType)int.Parse(buffer[0]); // is this dangerous? If all goes well, buffer always contains a number represented as string.
                string s = buffer[1];
                fd.Combobox = s.Substring(6,1) == "1";
                fd.Shared = s.Substring(7, 1) == "1";
                fd.Mandatory = s.Substring(8, 1) == "1";
                fd.Recurring = s.Substring(9, 1) == "1";
                fd.MaxChars = Convert.ToInt32(buffer[2]);
                fd.DefaultString = buffer[3];
            }
            return fd;
        }

        /// <inheritdoc />
        public string GetFieldNames(string categoryName, string delim = null)
        {
            return (delim == null)
                ? GetDDEValues(new string[] { "GetFieldNames", categoryName})
                : GetDDEValues(new string[] { "GetFieldNames", categoryName, delim });
        }

        /// <inheritdoc />
        public List<string> GetFieldNames(string categoryName)
        {
            return GetDDEValuesAsList(new string[] { "GetFieldNames", categoryName, this.Delim });
        }

        /* this method may prove a little problematic.
        * It takes a list of fieldnames; because we do not know in advance what fields will be specified,
        * you have to pass in an array as parameter, however,
        * for this to work from COM, this array has to be of type 'object'
        * note that this function is obsolete in the sense that this can be done much faster using CommenceQueryRowSet
        * It can be useful however for getting fieldvalues of the active item.
        * 
        * sample usage from VBA:
        * GetFields(categoryName, itemName, Array(fieldname1, fieldname2, fieldnameN))
        */
        /// <inheritdoc />
        public string GetFields(string categoryName, string itemName, object[] fields, string delim = null)
        {
            // we need fieldnames to be  comma delimited string, so we can feed it to a new string array
            // our fieldnames are trapped in object array args,
            // so we first cast that
            string[] fieldNames = Utils.ToStringArray(fields);
            // create a comma-delimited string from fieldnames and pass that into a new string[] array.
            // how is that for unnecessary overhead? :)
            // also note that we have to supply the number of fields we want. Intriguing tidbit.
            return (delim == null)
                ? GetDDEValues(new string[] { "GetFields", categoryName, itemName, fieldNames.Length.ToString(), string.Join(",", EncodeDdeArguments(fieldNames)) })
                : GetDDEValues(new string[] { "GetFields", categoryName, itemName, fieldNames.Length.ToString(), string.Join(",", EncodeDdeArguments(fieldNames)), delim });
        }

        /// <inheritdoc />
        public List<string> GetFields(string categoryName, string itemName, List<string> fieldNames)
        {
            return GetDDEValuesAsList(new string[] { "GetFields", categoryName, itemName, fieldNames.Count.ToString(), string.Join(",", EncodeDdeArguments(fieldNames)), this.Delim });
        }

        /// <inheritdoc />
        public bool GetFieldToFile(string categoryName, string itemName, string fieldName, string fileName)
        {
            object retval = DDERequest(BuildDDERequestCommand(new string[] {"GetFieldToFile", categoryName, itemName, fieldName, fileName }));
            return retval != null;
        }

        /// <inheritdoc />
        public int GetFormCount(string categoryName)
        {
            return GetDDECount(new string[] { "GetFormCount", categoryName });
        }

        /// <inheritdoc />
        public string GetFormNames(string categoryName, string delim = null)
        {
            return (delim == null)
                ? GetDDEValues(new string[] { "GetFormNames", categoryName })
                : GetDDEValues(new string[] { "GetFormNames", categoryName, delim });
        }

        /// <inheritdoc />
        public List<string> GetFormNames(string categoryName)
        {
            return GetDDEValuesAsList(new string[] { "GetFormNames", categoryName, this.Delim});
        }

        /// <inheritdoc />
        public int GetImageFieldCount(string categoryName)
        {
            return GetDDECount(new string[] { "GetImageFieldCount", categoryName });
        }

        /// <inheritdoc />
        public string GetImageFieldNames(string categoryName, string delim = null)
        {
            return (delim == null)
                ? GetDDEValues(new string[] { "GetImageFieldNames", categoryName })
                : GetDDEValues(new string[] { "GetImageFieldNames", categoryName, delim });
        }

        /// <inheritdoc />
        public List<string> GetImageFieldNames(string categoryName)
        {
            return GetDDEValuesAsList(new string[] { "GetImageFieldNames", categoryName, this.Delim });
        }

        /// <inheritdoc />
        public bool GetImageFieldToFile(string categoryName, string itemName, string fieldName, string fileName)
        {
            object retval = DDERequest(BuildDDERequestCommand(new string[] { "GetImageFieldToFile", categoryName, itemName, fieldName, fileName }));
            return retval != null;
        }

        /// <inheritdoc />
        public int GetItemCount(string categoryName)
        {
            return GetDDECount(new string[] { "GetItemCount", categoryName });
        }

        /// <inheritdoc />
        public string GetItemNames(string categoryName, string delim = null)
        {
            return (delim == null)
                ? GetDDEValues(new string[] { "GetItemNames", categoryName })
                : GetDDEValues(new string[] { "GetItemNames", categoryName, delim });
        }

        /// <inheritdoc />
        public List<string> GetItemNames(string categoryName)
        {
            return GetDDEValuesAsList(new string[] { "GetItemNames", categoryName, this.Delim });
        }

        /// <inheritdoc />
        public string GetLastError()
        {
            return this.LastError;
        }

        /// <inheritdoc />
        public string GetLetterViewInfo(string delim)
        {
            throw new NotImplementedException();
        }

        /// <inheritdoc />
        public bool GetMarkItem(string categoryNameOrKeyword, string itemName = "", string clarifyValue = "")  // expects subsequent requests
        {
            object retval = DDERequest(BuildDDERequestCommand(new string[] { "GetMarkItem", categoryNameOrKeyword, itemName, clarifyValue }));
            return retval != null;
        }

        /// <inheritdoc />
        public string GetMeField(string fieldName)
        {
            return GetDDEValues(new string[] { "GetMeField", fieldName });
        }

        /// <inheritdoc />
        public string GetNameField(string categoryName)
        {
            string retval = null;
            try
            {
                // GetCursor is a very expensive operation. We need to optimize this
                using (ICommenceCursor cur = this.GetCursor(categoryName))
                {
                    using (ICommenceQueryRowSet qrs = cur.GetQueryRowSet(0))
                    {
                        retval = qrs.GetColumnLabel(0);
                    }
                }
            }
            catch (CommenceCOMException) {}
            return retval;
        }

        /// <inheritdoc />
        public string GetPhoneNumber(string phoneNumber, string delim = null)
        {
            throw new NotImplementedException();
        }

        /// <inheritdoc />
        public string GetPreference(string preferenceSetting)
        {
            Preferences pref = this.GetPreferences();

            switch (preferenceSetting.ToLower())
            {
                case "me":
                    return pref.Me;
                case "mecategory":
                    return pref.MeCategory;
                case "letterlogdir":
                    return pref.LetterLogDir;
                case "externaldir":
                    return pref.ExternalDir;
                default:
                    return "Unrecognized preferences setting.";
            }
        }

        /// <summary>
        /// Returns populated Prefereces object
        /// </summary>
        /// <returns>Preferences</returns>
        [ComVisible(false)]
        public Preferences GetPreferences()
        {
            Preferences pref = new Preferences();
            string dde = BuildDDERequestCommand(new string[] { "GetPreference", "Me", this.Delim });
            var result = DDERequest(dde); // null when no Me-category was defined
            if (result != null)
            {
                string[] me = result.Split(this.Splitter, StringSplitOptions.None);
                pref.MeCategory = me[0];
                pref.Me = me[1];
            }
            pref.LetterLogDir = GetDDEValues(new string[] { "GetPreference", "LetterlogDir" });
            pref.ExternalDir = GetDDEValues(new string[] { "GetPreference", "ExternalDir" });
            return pref;
        }

        /// <inheritdoc />
        public string GetReverseName(string itemName, int prefFlag = 0)
        {
            return GetDDEValues(new string[] { "GetReverseName", itemName, prefFlag.ToString() });
        }

        /// <inheritdoc />
        public int GetTriggerCount()
        {
            return GetDDECount(new string[] { "GetTriggerCount" });
        }

        /// <inheritdoc />
        public string GetTriggerNames(string delim = null)
        {
            return (delim == null)
                ? GetDDEValues(new string[] { "GetTriggerNames" })
                : GetDDEValues(new string[] { "GetTriggerNames", delim });
        }

        /// <inheritdoc />
        public List<string> GetTriggerNames()
        {
            return GetDDEValuesAsList(new string[] { "GetTriggerNames", this.Delim });
        }

        /// <inheritdoc />
        public int GetViewCount(string categoryName)
        {
            return GetDDECount(new string[] { "GetViewCount", categoryName, this.Delim });
        }

        /// <inheritdoc />
        public IViewDef GetViewDefinition(string viewName)
        {
            ViewDef vd;
            // note that this request is undocumented by Commence!
            string dde = BuildDDERequestCommand(new string[] { "GetViewDefinition", viewName, this.Delim });
            string viewInfo = DDERequest(dde);
            if (viewInfo != null)
            {
                vd = new ViewDef();
                string[] buffer = viewInfo.Split(this.Splitter, StringSplitOptions.None);
                vd.Name = buffer[0];
                vd.Type = buffer[1];
                vd.Category = buffer[2];
                vd.FileName = buffer[3];
                vd.ViewType = Utils.EnumFromAttributeValue<CommenceViewType, StringValueAttribute>(nameof(StringValueAttribute.StringValue), vd.Type);
            }
            else
            {
                throw new CommenceCOMException($"Unable to get information on view '{viewName}'. It may not exist");
            }
            return vd;
        }

        /// <inheritdoc />
        public string GetViewNames(string categoryName, string delim = null)
        {
            return (delim == null)
                ? GetDDEValues(new string[] { "GetViewNames", categoryName })
                : GetDDEValues(new string[] { "GetViewNames", categoryName, delim });
        }

        /// <inheritdoc />
        public List<string> GetViewNames(string categoryName)
        {
            return GetDDEValuesAsList(new string[] { "GetViewNames", categoryName, this.Delim });
        }

        /// <inheritdoc />
        public bool GetViewToFile(string viewName, string fileName, CmcLinkMode linkMode = CmcLinkMode.None, string param1 = null, string param2 = null)
        {
            object retval = DDERequest(BuildDDERequestCommand(new string[] { "GetViewToFile", viewName, ((int)linkMode).ToString(), param1, param2, fileName }));
            return retval != null;
        }

        /// <inheritdoc />
        public IEnumerable<string> GetViewColumnNames(string viewName, CmcOptionFlags flags = CmcOptionFlags.Fieldname)
        {
            IList<string> retval = new List<string>();
            using (ICommenceCursor cur = this.GetCursor(viewName, CmcCursorType.View, CmcOptionFlags.Default))
            {
                using (ICommenceQueryRowSet qrs = cur.GetQueryRowSet(0))
                {
                    for (int i = 0; i < qrs.ColumnCount; i++)
                    {
                        retval.Add(qrs.GetColumnLabel(i, flags));
                    }
                }
            }
            return retval;
        }

        /// <inheritdoc />
        public string MarkActiveItem(string delim = null) // expects subsequent requests
        {
            // will fail if no item is active, for instance when no views are open
            // consumers should always check for null!
            return (delim != null)
                ? DDERequest(BuildDDERequestCommand(new string[] { "MarkActiveItem", delim }))
                : DDERequest(BuildDDERequestCommand(new string[] { "MarkActiveItem"}));
        }

        /// <inheritdoc />
        public string MergeTemplateCreate(string templateName, string categoryName, bool shared)
        {
            return GetDDEValues(new string[] { "MergeTemplateCreate", templateName, categoryName, (shared) ? "1" : "0" });
        }

        /// <inheritdoc />
        public string MergeTemplateSave(string templateName, bool shared)
        {
            return GetDDEValues(new string[] { "MergeTemplateSave", templateName, (shared) ? "1" : "0" });
        }

        /// <inheritdoc />
        public string ViewCategory(string categoryName)  // expects subsequent requests
        {
            return GetDDEValues(new string[] { "ViewCategory", categoryName });
        }

        /// <inheritdoc />
        public string ViewConjunction(string AndOr12, string AndOr13, string AndOr34, string AndOr48, string AndOr56, string AndOr57, string AndOr78)
        {
             return GetDDEValues(new string[] { "ViewConjunction", AndOr12, AndOr13, AndOr34, AndOr48, AndOr56, AndOr57, AndOr78 });
        }

        /// <inheritdoc />
        public int ViewConnectedCount(int index, string connectionName, string connCategory)
        {
            return GetDDECount(new string[] { "ViewConnectedCount", index.ToString(), connectionName, connCategory });
        }

        /// <inheritdoc />
        public string ViewConnectedField(int index, string connectionName, string connCategory, int connIndex, string fieldName)
        {
            return GetDDEValues(new string[] { "ViewConnectedField", index.ToString(), connectionName, connCategory, connIndex.ToString(), fieldName });
        }

        /// <inheritdoc />
        public string ViewConnectedFields(int index, string connectionName, string connCategory, int connIndex, object[] fields, string delim =  null)
        {
            IEnumerable<string> buffer;
            // we need fieldnames to be  comma delimited string, so we can feed it to a new string array
            // our fieldnames are trapped in object array args,
            // so we first cast that
            string[] fieldNames = Utils.ToStringArray(fields);
            // create a comma-delimited string from fieldnames and pass that into a new string[] array.
            // how is that for unnecessary overhead? :)
            // also note that we have to supply the number of fields we want. Intriguing tidbit.
            if (delim == null)
            {
                buffer = new string[] { "ViewConnectedFields", index.ToString(), connectionName, connCategory, connIndex.ToString(), fieldNames.Length.ToString() };
                buffer = buffer.Concat(fieldNames);
                //return GetDDEValues(new string[] { "ViewConnectedFields", index.ToString(), connectionName, connCategory, connIndex.ToString(), fieldNames.Length.ToString(), string.Join(",", EncodeDdeArguments(fieldNames)) });
            }
            else
            {
                buffer = new string[] { "ViewConnectedFields", index.ToString(), connectionName, connCategory, connIndex.ToString(), fieldNames.Length.ToString() };
                buffer = buffer.Concat(fieldNames);
                buffer = buffer.Concat(new string[] { delim });
                //return GetDDEValues(new string[] { "ViewConnectedFields", index.ToString(), connectionName, connCategory, connIndex.ToString(), fieldNames.Length.ToString(), string.Join(",", EncodeDdeArguments(fieldNames)), delim });
            }
            return GetDDEValues(buffer.ToArray());
        }

        /// <inheritdoc />
        [ComVisible(false)]
        public string[] ViewConnectedFields(int index, string connectionName, string connCategory, int connIndex, List<string> fields, string delim = null)
        {
            return ViewConnectedFields(index,connectionName, connCategory, connIndex, fields.ToArray<string>(), this.Delim).Split(this.Splitter, StringSplitOptions.None);
        }

        /// <inheritdoc />
        public string ViewConnectedItem(int index, string connectionName, string connCategory, int connIndex)
        {
            return GetDDEValues(new string[] { "ViewConnectedItem", index.ToString(), connectionName, connCategory, connIndex.ToString() });
        }

        /// <inheritdoc />
        public void ViewDeleteAllItems()
        {
            DDERequest(BuildDDERequestCommand(new string[] { "ViewDeleteAllItems" }));
        }

        /// <inheritdoc />
        public string ViewField(int index, string fieldName)
        {
            return GetDDEValues(new string[] { "ViewField", index.ToString(), fieldName });
        }

        /// <inheritdoc />
        public string ViewFields(int index, object[] fields, string delim = null)
        {
            // we need fieldnames to be  comma delimited string, so we can feed it to a new string array
            // our fieldnames are trapped in object array args,
            // so we first cast that
            string[] fieldNames = Utils.ToStringArray(fields);
            IEnumerable<string> buffer;
            if (delim == null)
            {
                buffer = new string[] { "ViewFields", index.ToString(), fieldNames.Length.ToString() };
                buffer = buffer.Concat(fieldNames);
                //return GetDDEValues(new string[] { "ViewFields", index.ToString(), fieldNames.Length.ToString(), string.Join(",", EncodeDdeArguments(fieldNames)) });
                
            }
            else
            {
                buffer = new string[] { "ViewFields", index.ToString(), fieldNames.Length.ToString() };
                buffer = buffer.Concat(fieldNames);
                buffer = buffer.Concat(new string[] { delim });
                //return GetDDEValues(new string[] { "ViewFields", index.ToString(), fieldNames.Length.ToString(), string.Join(",", EncodeDdeArguments(fieldNames)), delim });
            }
            return GetDDEValues(buffer.ToArray());
        }

        /// <inheritdoc />
        [ComVisible(false)]
        public string[] ViewFields(int index, List<string> fields, string delim = null)
        {
            return ViewFields(index, fields.ToArray<string>(),this.Delim).Split(this.Splitter,StringSplitOptions.None);
        }

        /// <inheritdoc />
        public bool ViewFieldToFile(int index, string fieldName, string fileName)
        {
            object retval = DDERequest(BuildDDERequestCommand(new string[] { "ViewFieldToFile", index.ToString(), fieldName, fileName }));
            return retval != null;
        }

        /// <inheritdoc />
        public bool ViewFilter(int clauseNumber, string filterType, bool notFlag, object args)
        {
            string[] fltParams = Utils.ToStringArray(args);
            object retval = DDERequest(BuildDDERequestCommand(new string[] { "ViewFilter", clauseNumber.ToString(), filterType, (notFlag) ? "NOT" : "", String.Join(",", EncodeDdeArguments(fltParams)) }));
            return retval != null;
        }

        /// <inheritdoc />
        public bool ViewImageFieldToFile(int index, string fieldName, string fileName)
        {
            object retval = DDERequest(BuildDDERequestCommand(new string[] { "ViewImageFieldToFile", index.ToString(), fieldName, fileName }));
            return retval != null;
        }

        /// <inheritdoc />
        public int ViewItemCount()
        {
            return GetDDECount(new string[] { "ViewItemCount" });
        }

        /// <inheritdoc />
        public int ViewItemIndex(string nameFieldValue)
        {
            return GetDDECount(new string[] { "ViewItemIndex", nameFieldValue });
        }

        /// <inheritdoc />
        public string ViewItemName(int index)
        {
            return GetDDEValues(new string[] { "ViewItemName", index.ToString() });
        }

        /// <inheritdoc />
        public bool ViewMarkItem(int index)
        {
            object retval = DDERequest(BuildDDERequestCommand(new string[] { "ViewMarkItem", index.ToString() }));
            return retval != null;
        }

        /// <inheritdoc />
        public string ViewReverseName(string itemName, int prefFlag = 0)
        {
            return GetDDEValues(new string[] { "ViewReverseName", itemName, prefFlag.ToString() });
        }

        /// <inheritdoc />
        public bool ViewSaveView(string newViewName, bool shared)
        {
            object retval = DDERequest(BuildDDERequestCommand(new string[] { "ViewSaveView", newViewName, (shared)?"yes":"no" }));
            return retval != null;
        }

        /// <inheritdoc />
        public bool ViewSort(string fieldName1, string sortOrder1, string fieldName2 = "", string sortOrder2 = "", string fieldName3 = "", string sortOrder3 = "", string fieldName4 = "", string sortOrder4 = "")
        {
            object retval = DDERequest(BuildDDERequestCommand(new string[] { "ViewSort", fieldName1, sortOrder1, fieldName2, sortOrder2, fieldName3, sortOrder3, fieldName4, sortOrder4 }));
            return retval != null;
        }

        /// <inheritdoc />
        public bool ViewView(string viewName = "")
        {
            object retval = DDERequest(BuildDDERequestCommand(new string[] { "ViewView", viewName }));
            return retval != null;
        }
        /// <summary>
        /// Perform DDE request and return results
        /// </summary>
        /// <param name="ddeRequestCommand">DDE request as defined by Commence DDE specifations in dde.chm</param>
        /// <returns>String array, <c>null</c> on error</returns>
        private string DDERequest(string ddeRequestCommand) // note the omission of topic
        {
            // When we open a DDE channel, we start a timer.
            // We then subscribe our conversation to the elapsed event of the timer
            // The conversation gets closed after a set timeout.
            // This allows us to make subsequent requests within the same DDE conversation,
            // provided they are made before the timeout.
            // Whenever a new request is made, the timer is restarted.
            // The reason for this is that it prevents
            // a) the overhead of opening a new conversation
            // b) too many conversations being opened (there can be only 10).
            string retval = null;

            if (_conv == null)
            {
                _conv = CommenceConversation.Instance; // CommenceConversation handles the actual call to commence
                _conv.Topic = this._db.Path;
            }

            if (DDETimer == null)
            {
                DDETimer = new Timer(DDETimeout)
                {
                    AutoReset = false
                };
                DDETimer.Elapsed += HandleTimerElapsed;
            }
            DDETimer.Interval = DDETimeout;
            DDETimer.Start(); // (re)start timer

            try
            {
                this.LastError = string.Empty; // clear GetLastError
                retval = _conv.DDERequest(ddeRequestCommand);
            }
            catch (CommenceDDEException e)
            {
                this.LastError = e.Message; // store the last error
                retval = null; // if null, we know an exception occurred, we will not (re)throw it!
            }
            return retval;
        }

        #endregion

        #region Commence DDE Execute methods

        /// <inheritdoc />
        public bool AddItem(string categoryName, string itemName, string clarifyValue = "")
        {
            return DDEExecute(BuildDDERequestCommand(new string[] { "AddItem", categoryName, itemName, clarifyValue }));
        }

        /// <inheritdoc />
        public bool AddSharedItem(string categoryName, string itemName, string clarifyValue = "")
        {
            return DDEExecute(BuildDDERequestCommand(new string[] { "AddSharedItem", categoryName, itemName, clarifyValue }));
        }

        /// <inheritdoc />
        public bool AppendText(string categoryName, string itemName, string fieldName, string text)
        {
            return DDEExecute(BuildDDERequestCommand(new string[] { "AppendText", categoryName, itemName, fieldName, text }));
        }

        /// <inheritdoc />
        public bool AssignConnection(string categoryName, string itemName, string connectionName, string connCategory, string connItem)
        {
            return DDEExecute(BuildDDERequestCommand(new string[] { "AssignConnection", categoryName, itemName, connectionName, connCategory, connItem }));
        }

        /// <inheritdoc />
        public bool CheckInFormScript(string categoryName, string formName, string fileName)
        {
            return DDEExecute(BuildDDERequestCommand(new string[] { "CheckInFormScript", categoryName, formName, fileName }));
        }

        /// <inheritdoc />
        public bool CheckOutFormScript(string categoryName, string formName, string fileName)
        {
            return DDEExecute(BuildDDERequestCommand(new string[] { "CheckOutFormScript", categoryName, formName, fileName }));
        }

        /// <inheritdoc />
        public bool DeleteItem(string categoryName, string itemName)
        {
            return DDEExecute(BuildDDERequestCommand(new string[] { "DeleteItem", categoryName, itemName }));
        }

        /// <inheritdoc />
        public bool DeleteView(string viewName)
        {
            return DDEExecute(BuildDDERequestCommand(new string[] { "DeleteView", viewName }));
        }

        /// <inheritdoc />
        public bool EditItem(string categoryName, string itemName, string fieldName, string fieldValue)
        {
            return DDEExecute(BuildDDERequestCommand(new string[] { "EditItem", categoryName, itemName, fieldName, fieldValue }));
        }

        /// <inheritdoc />
        public bool FireTrigger(string trigger, object[] args)
        {
            string[] trgParams = Utils.ToStringArray(args);
            if (trgParams.Length > 9)
            {
                this.LastError = "Too many parameters for FireTrigger command.";
                return false;
            }
            //return DDEExecute(buildDDERequestCommand(new string[] { "FireTrigger", trigger, string.Join("\",\"", trgParams) }));
            return DDEExecute(BuildDDERequestCommand(new string[] { "FireTrigger", trigger, string.Join(",", EncodeDdeArguments(trgParams)) }));
        }

        /// <inheritdoc />
        public bool LogPhoneCall(object[] args)
        {
            string[] ciPairs = Utils.ToStringArray(args);
            //return DDEExecute(buildDDERequestCommand(new string[] { "LogPhoneCall", string.Join("\",\"", ciPairs) }));
            return DDEExecute(BuildDDERequestCommand(new string[] { "LogPhoneCall", string.Join(",", EncodeDdeArguments(ciPairs)) }));
        }

        /// <inheritdoc />
        public bool PromoteItemToShared(string categoryName, string itemName)
        {
            return DDEExecute(BuildDDERequestCommand(new string[] { "PromoteItemToShared", categoryName, itemName }));
        }

        /// <inheritdoc />
        public bool ShowDesktop(string desktopName)
        {
            return DDEExecute(BuildDDERequestCommand(new string[] { "ShowDesktop", desktopName }));
        }

        /// <inheritdoc />
        public bool ShowItem(string categoryName, string itemName, string formName = null)
        {
            if (formName == null)
            {
                return DDEExecute(BuildDDERequestCommand(new string[] { "ShowItem", categoryName, itemName}));
            }
            else
            {
                return DDEExecute(BuildDDERequestCommand(new string[] { "ShowItem", categoryName, itemName, formName }));
            }
        }

        /// <inheritdoc />
        public bool ShowView(string viewName, bool newCopy)
        {
            return DDEExecute(BuildDDERequestCommand(new string[] { "ShowView", viewName, (newCopy) ? "1" : "0" }));
        }

        /// <inheritdoc />
        public bool UnassignConnection(string categoryName, string itemName, string connectionName, string connCategory, string connItem)
        {
            return DDEExecute(BuildDDERequestCommand(new string[] { "UnassignConnection", categoryName, itemName, connectionName, connCategory, connItem }));
        }


        /// <summary>
        /// Performs the DDE Excecute and stores last error, if any.
        /// </summary>
        /// <param name="DDEExecuteCommand">DDE Execute command</param>
        /// <returns>true on succes, false on failure. Inspect GetLastError on failure.</returns>
        private bool DDEExecute(string DDEExecuteCommand)
        {
            bool retval = false;
            if (_conv == null)
            {
                _conv = CommenceConversation.Instance;
                _conv.Topic = this._db.Path;
            }
            if (DDETimer == null)
            {
                DDETimer = new Timer(DDETimeout)
                {
                    AutoReset = false
                };
                DDETimer.Elapsed += HandleTimerElapsed;
            }
            DDETimer.Interval = DDETimeout; // (re)set interval
            DDETimer.Start(); // (re)start timer

            try
            {
                this.LastError = string.Empty; // clear GetLastError
                retval = _conv.DDEExecute(DDEExecuteCommand);
            }
            catch (CommenceDDEException e)
            {
                this.LastError = e.Message; // store the last error
            }
            return retval;
        }

        /// <inheritdoc />
        public bool FieldValueExists(string categoryName, string fieldName, string fieldValue, bool caseSensitive = true)
        {
            using (ICommenceCursor cur = this.GetCursor(categoryName))
            {
                ICursorFilters filters = new CursorFilters(cur);
                ICursorFilterTypeF f = filters.Create(1, FilterType.Field);
                f.FieldName = fieldName;
                f.FieldValue = fieldValue;
                f.Qualifier = FilterQualifier.EqualTo;
                if (caseSensitive) { f.MatchCase = true; }
                if (filters.Apply() > 0) { return true;}
            }
            return false;
        }
        #endregion

        #region Methods

        /// <inheritdoc />
        public bool HasDuplicates(string categoryName, string fieldName, bool caseSensitive = true)
        {
            using (ICommenceCursor cur = this.GetCursor(categoryName, CmcCursorType.Category, CmcOptionFlags.All))
            {
                return cur.HasDuplicates(fieldName, caseSensitive);
            }
        }

        /// <summary>
        /// Returns the ID of a category.
        /// </summary>
        /// <param name="categoryName">Commence categoryname.</param>
        /// <returns>category ID</returns>
        private int GetCategoryID(string categoryName)
        {
            /* <*RowSet>.GetRowID() returns a string of hexadecimal values delimited by a colon.
             * Shared databases return the string in a different format than Local databases.
             * We assume here that the first element is always the category ID, regardless of format.
             */
            int retval = -1;
            try
            {
                using (ICommenceCursor cur = this.GetCursor(categoryName,CmcCursorType.Category,CmcOptionFlags.UseThids))
                {
                    if (cur.RowCount > 0) // category contains items, we can just request one
                    {
                        using (ICommenceQueryRowSet qrs = cur.GetQueryRowSet(1))
                        {
                            string[] thid = qrs.GetRowID(0).Split(new char[] { ':' }); // returns string of hex values
                            retval = Convert.ToInt32(thid[0], 16);
                        }
                    }
                    else
                    {
                        //// Now we are in trouble. The category does not contain items.
                        //// We can create a local item.
                        //// This *should* always work, regardless of permissions on the category.
                        //// However, if the category is intentionally empty and
                        //// protected by a Commence Agent that deletes items immediately after they are created,
                        //// this will fail. There is no workaround in that case.
                        //// Commence may even crash.
                        //// This is why we swallow all errors.
                        //// In fact, just skip all this altogether.
                        //// The risk of triggering an item, as well as errors on missing mandatory fields is simply to great
                        //using (ICommenceAddRowSet ars = cur.GetAddRowSet(1))
                        //{
                        //    string guid = Guid.NewGuid().ToString();
                        //    ars.ModifyRow(0, 0, guid, CmcOptionFlags.Default);
                        //    using (ICommenceCursor newcur = ars.CommitGetCursor(CmcOptionFlags.UseThids)) // this SHOULD only return a cursor with the item we just added
                        //    {
                        //        using (ICommenceQueryRowSet qrs = newcur.GetQueryRowSet(newcur.RowCount))
                        //        {
                        //            // double-check we are only deleting the item we just created
                        //            if ((qrs.RowCount == 1) && (qrs.GetRowValue(0, 0, CmcOptionFlags.Default) == guid))
                        //            {
                        //                // calling GetRowValue moves the rowpointer.
                        //                // move back rowpointer or getting a DeleteRowSet will return an empty rowset!
                        //                newcur.SeekRow(CmcCursorBookmark.Beginning, 0);
                        //                using (ICommenceDeleteRowSet drs = newcur.GetDeleteRowSet(newcur.RowCount))
                        //                {
                        //                    string[] thid = drs.GetRowID(0).Split(new char[] { ':' });
                        //                    retval = Convert.ToInt32(thid[0], 16);
                        //                    drs.DeleteRow(0);
                        //                    drs.Commit();
                        //                } // using ICommenceDeleteRowSet drs
                        //            } // if
                        //        } // using ICommenceQueryRowSet qrs
                        //    } // using ICommenceCursor newcur
                        //} // using ICommenceAddRowSet ars
                    } // else
                } // using ICommenceCursor cur
            } // try
            catch (CommenceCOMException) 
            {
                /* ignore all errors */
            }
            return retval;
        }

        #endregion

        #region Helper methods
        /// <summary>
        /// Closes the DDE conversation.
        /// Multiple requests can made in a single conversation,
        /// so closing the conversation after every request would add considerable overhead.
        /// <para>Conversations should kept open until a Timer elapses.
        /// This method should subscribe to the Elapsed event of that Timer.</para>
        /// <para>This way, the conversation stays open and multiple requests can be made.
        /// If no more requests are received, the conversation is closed after the timer elapses.</para>
        /// </summary>
        /// <param name="sender">sender.</param>
        /// <param name="e">ElapsedEventArgs.</param>
        private void HandleTimerElapsed(object sender, ElapsedEventArgs e)
        {
            _conv?.CloseConversation();
        }

        private List<string> GetDDEValuesAsList(string[] args)
        {
            string values = DDERequest(BuildDDERequestCommand(args));
            // TODO code-smell, we should probably just return an empty list
            // for example, there is an edge-case scenario in which a category can have no views
            // in that case, there is also no last error and I think we end up with an unwanted empty element
            if (values == null) // an error occurred, return error value
            {
                List<string> retval = new List<string>
                {
                    this.GetLastError()
                };
                return retval;
            }
            // the delimiter is always the last element
            string[] splitter = new string[] { args.Last() };
            return values.Split(splitter, StringSplitOptions.None).ToList();
        }

        private string GetDDEValues(string[] args)
        {
            string retval;
            retval = DDERequest(BuildDDERequestCommand(args));
            if (retval == null) // an error occurred, return error value
            {
                retval = this.GetLastError();
            }
            return retval;
        }

        /// <summary>
        /// Returns the *Count of the object specified in DDE request command as specified by args
        /// </summary>
        /// <param name="args">Parameters of the desired request. YOU are responsible for the right order!</param>
        /// <returns>Count, -1 on error.</returns>
        private int GetDDECount(string[] args)
        {
            string result = DDERequest(BuildDDERequestCommand(args));
            return (int.TryParse(result, out int outval)) ? outval: -1;
        }

        /// <inheritdoc />
        public bool EncodeDDEArguments { get; set; } = true;

        /// <summary>
        /// Builds the DDE request string for Commence.
        /// </summary>.
        /// <param name="args">Request item plus optional parameters. The request item string must always be the first element.</param>
        /// <returns>string in format "[Request command(param1, param2, paramN)]".</returns>
        private string BuildDDERequestCommand(string[] args)
        {
            StringBuilder sb;
            sb = new StringBuilder("[" + args[0]);

            if (args.Length == 1)// request item without additional parameters
            {
                sb.Append("()]");
                return sb.ToString();
            }
            IEnumerable<string> arguments = args.Skip(1); // note the skip of the first argument
            arguments = EncodeDdeArguments(arguments);
            sb.Append("(");
            sb.Append(string.Join(",", arguments));
            sb.Append(")]"); 
            return sb.ToString();
        }

        /// <summary>
        /// Encode array of arguments.
        /// </summary>
        /// <param name="args">string array.</param>
        /// <returns>Encoded array of arguments.</returns>
        private IEnumerable<string> EncodeDdeArguments(IEnumerable<string> args) // move this to extension method
        {
            foreach (string arg in args)
            {
                if (EncodeDDEArguments)
                {
                    // enclose the entire argument in double quotes
                    // and escape embedded double quotes
                    yield return $"\"{arg.Replace("\"", "\"\"")}\"";
                }
                else // return args as is
                {
                    yield return arg; 
                }
            }
        }

        void ValidateDelimiter(string s)
        {
            if (string.IsNullOrEmpty(s)) 
            {
                throw new ArgumentNullException("Delimiter cannot be an empty string."); 
            }
            if (s.Length > CommenceLimits.MaxDelimiterLength)
            {
                throw new ArgumentOutOfRangeException($"Delimiter exceeds maximmum length of {CommenceLimits.MaxDelimiterLength} characters");
            }
        }
        #endregion

        #region Properties
        private string _delim = @"|$%`!^*|";
        /// <inheritdoc />
        public string Delim 
        { get
            {
                return _delim;
            }
            set 
            {
                ValidateDelimiter(value);
                _delim = value; 
            }
        }
        
        private string _delim2 = @"|@##~~&|";
        /// <inheritdoc />
        public string Delim2
        {
            get
            {
                return _delim2;
            }
            set
            {
                ValidateDelimiter(value);
                _delim2 = value;
            }
        }
        private string[] Splitter => new string[] { Delim };
        private string[] Splitter2 => new string[] { Delim2 };

        /// <summary>
        /// Last DDE error received from Commence.
        /// </summary>
        private string LastError { get; set; }
        #endregion
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Timers;
using Vovin.CmcLibNet.Extensions;

namespace Vovin.CmcLibNet.Database
{

    // TODO: it may be a good idea to implement IDisposable?
    // This portion of the class contains the implementation of the properties and methods exposed by Commence's ICommenceGetconversation interface.
    // That is a difficult way of saying all DDE stuff is in here.
    public partial class CommenceDatabase : ICommenceDatabase
    {
        /// <summary>
        /// Delimiters used in Commence DDE conversations
        /// By defining them here, consumers do not have to supply them with every 'DDE call' that uses them
        /// </summary>
        private const string CMC_DELIM = @"|$%`!^*|";
        private const string CMC_DELIM2 = @"|@##~~&|";
        private readonly string[] _splitter = new string[] { CMC_DELIM };
        private readonly string[] _splitter2 = new string[] { CMC_DELIM2 };
        private static Timer DDETimer = null; // static because we only use 1 timer
        private readonly int DDETimeout = 5000; // milliseconds after which a DDE conversation is closed.
        private CommenceConversation _conv = null;

        #region DDE Request methods
        /// <inheritdoc />
        public string GetActiveCategory()
        {
            // first try to mark the active item
            string result = MarkActiveItem(CMC_DELIM);
            if (result != null) // use the active item
            {
                string[] buffer = result.Split(new string[] { CMC_DELIM }, StringSplitOptions.None);
                return buffer[0];
            }
            else // no active item, try getting info on the view (if any) instead
            {
                IActiveViewInfo ai = GetActiveViewInfo(); // will not work on all types of views, but at least we tried.
                if (ai != null)
                {
                    return ai.Category;
                }
                else
                {
                    return null; // no view active
                }
            }
        }

        /// <inheritdoc />
        public string GetActiveItemName()
        {
            string result = MarkActiveItem(CMC_DELIM);
            // the view can be empty, or not support getting of data (like the Document view).
            if (result == null) { return null; }
            string[] buffer = result.Split(new string[] { CMC_DELIM }, StringSplitOptions.None);
            return buffer[1];
        }

        /// <inheritdoc />
        public string ClarifyItemNames(string bStatus = null)
        {
            if (bStatus == null)
            {
                return DDERequest(BuildDDERequestCommand(new string[] { "ClarifyItemNames" }));
            }
            else
            {
                return DDERequest(BuildDDERequestCommand(new string[] { "ClarifyItemNames", bStatus }));
            }
        }

        /// <inheritdoc />
        public IActiveViewInfo GetActiveViewInfo()
        {
            ActiveViewInfo avi = null;
            string dde = BuildDDERequestCommand(new string[] { "GetActiveViewInfo", CMC_DELIM });
            string viewInfo = DDERequest(dde);
            if (viewInfo != null) // null means no view was active
            {
                // commence will return {ViewName}Delim {ViewType}Delim {CategoryName}Delim {ItemName}Delim {FieldName}
                string[] buffer = viewInfo.Split(new string[] { CMC_DELIM },StringSplitOptions.None);
                avi = new ActiveViewInfo();
                avi.Name = buffer[0];
                avi.Type = (buffer[1] == "Add Item") ? "Item Detail Form" : buffer[1]; // translate. This is a discrepancy in Commence documentation.
                avi.Category = buffer[2];
                avi.Item = buffer[3];
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
            if (delim == null)
            {
                return GetDDEValues(new string[] { "GetCallerID", EncodeDdeArgument(categoryName), phoneNumber });
            }
            else
            {
                return GetDDEValues(new string[] { "GetCallerID", EncodeDdeArgument(categoryName), phoneNumber, delim });
            }
        }

        /// <inheritdoc />
        public List<string> GetCallerID(string categoryName, string phoneNumber)
        {
            return GetDDEValuesAsList(new string[] { "GetCallerID", EncodeDdeArgument(categoryName), phoneNumber, CMC_DELIM });
        }

        /// <inheritdoc />
        public int GetCategoryCount() // this could be more elegant
        {
            return GetDDECount(new string[] { "GetCategoryCount" });
        }

        /// <inheritdoc />
        public ICategoryDef GetCategoryDefinition(string categoryName)
        {
            CategoryDef cd = null;
            string dde = BuildDDERequestCommand(new string[] { "GetCategoryDefinition", EncodeDdeArgument(categoryName), CMC_DELIM });
            string categoryInfo = DDERequest(dde);

            if (categoryInfo != null)
            {
                cd = new CategoryDef();
                string[] buffer = categoryInfo.Split(new string[] { CMC_DELIM }, StringSplitOptions.None);
                cd.MaxItems = Convert.ToInt32(buffer[0]);
                string s = buffer[1];
                cd.Shared = (s.Substring(6, 1) == "1") ? true : false;
                // note that we skip substring 7 - it has no meaning
                cd.Duplicates = (s.Substring(8, 1) == "1") ? true : false;
                cd.Clarified = (s.Substring(9, 1) == "1") ? true : false;
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
            if (delim == null)
            {
                return GetDDEValues(new string[] { "GetCategoryNames"});
            }
            else
            {
                return GetDDEValues(new string[] { "GetCategoryNames", delim });
            }
        }

        /// <inheritdoc />
        public List<string> GetCategoryNames()
        {
            return GetDDEValuesAsList(new string[] { "GetCategoryNames", CMC_DELIM });
        }

        /// <inheritdoc />
        // this method is not a Commence method
        public string GetClarifiedItemName(string itemName, string clarifySeparator, string clarifyValue)
        {
            return itemName.PadRight(50) + clarifySeparator + clarifyValue.PadRight(40);
        }

        /// <inheritdoc />
        public string GetClarifyField(string categoryName)
        {
            ICategoryDef cd = this.GetCategoryDefinition(categoryName);
            if (cd.ClarifyField == String.Empty)
            {
                return null;
            }
            else
            {
                return cd.ClarifyField;
            }
        }

        /// <inheritdoc />
        public string GetClarifySeparator(string categoryName)
        {
            ICategoryDef cd = this.GetCategoryDefinition(categoryName);
            if (cd.ClarifySeparator == String.Empty)
            {
                return null;
            }
            else
            {
                return cd.ClarifySeparator;
            }
        }

        /// <inheritdoc />
        public int GetConnectedItemCount(string categoryName, string itemName, string connectionName, string connCategory)
        {
            return GetDDECount(new string[] { "GetConnectedItemCount", EncodeDdeArgument(categoryName), itemName, EncodeDdeArgument(connectionName), connCategory });
        }

        /// <inheritdoc />
        public string GetConnectedItemField(string categoryName, string itemName, string connectionName, string connCategory, string fieldName, string delim = null)
        {
            if (delim == null)
            {
                return GetDDEValues(new string[] { "GetConnectedItemField", EncodeDdeArgument(categoryName), itemName, EncodeDdeArgument(connectionName), connCategory, EncodeDdeArgument(fieldName) });
            }
            else
            {
                return GetDDEValues(new string[] { "GetConnectedItemField", EncodeDdeArgument(categoryName), itemName, EncodeDdeArgument(connectionName), connCategory, EncodeDdeArgument(fieldName), delim });
            }
        }

        /// <inheritdoc />
        public string GetConnectedItemNames(string categoryName, string itemName, string connectionName, string connCategory, string delim = null)
        {
            if (delim == null)
            {
                return GetDDEValues(new string[] { "GetConnectedItemNames", EncodeDdeArgument(categoryName), itemName, EncodeDdeArgument(connectionName), connCategory });
            }
            else
            {
                return GetDDEValues(new string[] { "GetConnectedItemNames", EncodeDdeArgument(categoryName), itemName, EncodeDdeArgument(connectionName), connCategory, delim });
            }
        }

        /// <inheritdoc />
        public List<string> GetConnectedItemNames(string categoryName, string itemName, string connectionName, string connCategory)
        {
            return GetDDEValuesAsList(new string[] { "GetConnectedItemNames", EncodeDdeArgument(categoryName), itemName, EncodeDdeArgument(connectionName), connCategory, CMC_DELIM });
        }

        /// <inheritdoc />
        public int GetConnectionCount(string categoryName)
        {
            return GetDDECount(new string[] { "GetConnectionCount", EncodeDdeArgument(categoryName) });
        }

        /// <inheritdoc />
        public string GetConnectionNames(string categoryName, string delim1 =  null, string delim2 = null)
        {
            string retval = string.Empty;
            // both delim1 and delim2 are optional, but they must be supplied if sepcified
            // we'll create a DDERequest command directly and perform the request,
            // it is simpler that repeatedly creating and unboxing arrays.
            List<string> list = new List<string>();
            list.Add("GetConnectionNames");
            list.Add(categoryName);
            if (delim1 != null)
            {
                list.Add(delim1);
            }
            if (delim2 != null)
            {
                list.Add(delim2);
            }
            retval = GetDDEValues(list.ToArray<string>());
            return retval;
        }

        /// <inheritdoc />
        public List<Tuple<string, string>> GetConnectionNames(string categoryName)
        {
            List<Tuple<string, string>> retval = null;
            string buffer = GetDDEValues(new string[] { "GetConnectionNames", EncodeDdeArgument(categoryName), CMC_DELIM, CMC_DELIM2 });
            if ((buffer != string.Empty) && (this.GetLastError() == string.Empty))
            {
                string[] pairs = buffer.Split(new string[] { CMC_DELIM }, StringSplitOptions.None);
                //if (pairs.Length <= 1) { return retval; } // a DDE error occurred, possibly because categoryName doesn't exist.
                retval = new List<Tuple<string, string>>();
                foreach (string s in pairs)
                {
                    string[] pair = s.Split(new string[] { CMC_DELIM2 }, StringSplitOptions.None);
                    retval.Add(Tuple.Create(pair[0], pair[1]));
                }
            }
            return retval;
        }

        /// <inheritdoc />
        [ObsoleteAttribute("Use CmcLibNet.CommenceApp.Name and/or CmcLibNet.CommenceApp.Path")]
        public string GetDatabase()
        {
            return GetDDEValues(new string[] { "GetDatabase" }); // works, but superfluous.
            //throw new NotImplementedException();
        }

        /// <inheritdoc />
        public IDBDef GetDatabaseDefinition()
        {
            DBDef db = null;
            string dde = BuildDDERequestCommand(new string[] { "GetDatabaseDefinition", CMC_DELIM });
            string dbInfo = DDERequest(dde);
            if (dbInfo != null)
            {
                string[] buffer = dbInfo.Split(new string[] { CMC_DELIM }, StringSplitOptions.None);
                // Commence will return:
                // {DatabaseName}Delim {DatabasePath}Delim 000000{A}{X}{S}{C}Delim {UserName}Delim {SpoolPath}
                db = new DBDef();
                db.Name = buffer[0];
                db.Path = buffer[1];
                string s = buffer[2];
                db.Attached = (s.Substring(6, 1) == "1") ? true : false;
                db.Connected = (s.Substring(7, 1) == "1") ? true : false;
                db.IsServer = (s.Substring(8, 1) == "1") ? true : false;
                db.IsClient = (s.Substring(9, 1) == "1") ? true : false;
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
            if (delim == null)
            {
                return GetDDEValues(new string[] { "GetDesktopNames" });
            }
            else
            {
                return GetDDEValues(new string[] { "GetDesktopNames", delim });
            }
        }

        /// <inheritdoc />
        public List<string> GetDesktopNames()
        {
            return GetDDEValuesAsList(new string[] { "GetDesktopNames", CMC_DELIM });
        }

        /// <inheritdoc />
        public string GetField(string categoryName, string itemName, string fieldName)
        {
            return GetDDEValues(new string[] { "GetField", EncodeDdeArgument(categoryName), itemName, EncodeDdeArgument(fieldName) });
        }

        /// <inheritdoc />
        public int GetFieldCount(string categoryName)
        {
            return GetDDECount(new string[] { "GetFieldCount", EncodeDdeArgument(categoryName) });
        }

        /// <inheritdoc />
        public ICommenceFieldDefinition GetFieldDefinition(string categoryName, string fieldName)
        {
            CommenceFieldDefinition fd = null;
            string dde = BuildDDERequestCommand(new string[] { "GetFieldDefinition", EncodeDdeArgument(categoryName), EncodeDdeArgument(fieldName), CMC_DELIM });
            string fieldInfo = DDERequest(dde);
            if (fieldInfo != null)
            {
                string[] buffer = fieldInfo.Split(new string[] { CMC_DELIM }, StringSplitOptions.None);
                fd = new CommenceFieldDefinition();
                fd.Type = (CommenceFieldType)int.Parse(buffer[0]); // is this dangerous? If all goes well, buffer always contains a number represented as string.
                string s = buffer[1];
                fd.Combobox = (s.Substring(6,1) == "1") ? true : false;
                fd.Shared = (s.Substring(7, 1) == "1") ? true : false;
                fd.Mandatory = (s.Substring(8, 1) == "1") ? true : false;
                fd.Recurring = (s.Substring(9, 1) == "1") ? true : false;
                fd.MaxChars = Convert.ToInt32(buffer[2]);
                fd.DefaultString = buffer[3];
            }
            return fd;
        }

        /// <inheritdoc />
        public string GetFieldNames(string categoryName, string delim = null)
        {
            if (delim == null)
            {
                return GetDDEValues(new string[] { "GetFieldNames", EncodeDdeArgument(categoryName)});
            }
            else
            {
                return GetDDEValues(new string[] { "GetFieldNames", EncodeDdeArgument(categoryName), delim });
            }

        }
        /// <inheritdoc />
        public List<string> GetFieldNames(string categoryName)
        {
            return GetDDEValuesAsList(new string[] { "GetFieldNames", EncodeDdeArgument(categoryName), CMC_DELIM });
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
            string[] fieldNames = ToStringArray(fields);
            // create a comma-delimited string from fieldnames and pass that into a new string[] array.
            // how is that for unnecessary overhead? :)
            // also note that we have to supply the number of fields we want. Intriguing tidbit.
            if (delim == null)
            {
                //return GetDDEValues(new string[] { "GetFields", EncodeDdeArgument(categoryName), itemName, fieldNames.Length.ToString(), string.Join("\",\"", fieldNames) });
                return GetDDEValues(new string[] { "GetFields", EncodeDdeArgument(categoryName), itemName, fieldNames.Length.ToString(), string.Join(",", EncodeDdeArguments(fieldNames)) });
            }
            else
            {
                //return GetDDEValues(new string[] { "GetFields", EncodeDdeArgument(categoryName), itemName, fieldNames.Length.ToString(), string.Join("\",\"", fieldNames), delim });
                return GetDDEValues(new string[] { "GetFields", EncodeDdeArgument(categoryName), itemName, fieldNames.Length.ToString(), string.Join(",", EncodeDdeArguments(fieldNames)), delim });
            }
        }

        /// <inheritdoc />
        public List<string> GetFields(string categoryName, string itemName, List<string> fieldNames)
        {
            //return GetDDEValuesAsList(new string[] { "GetFields", EncodeDdeArgument(categoryName), itemName, fieldNames.Count.ToString(), string.Join("\",\"", fieldNames), CMC_DELIM });
            return GetDDEValuesAsList(new string[] { "GetFields", EncodeDdeArgument(categoryName), itemName, fieldNames.Count.ToString(), string.Join(",", EncodeDdeArguments(fieldNames)), CMC_DELIM });
        }

        /// <inheritdoc />
        public bool GetFieldToFile(string categoryName, string itemName, string fieldName, string fileName)
        {
            object retval = DDERequest(BuildDDERequestCommand(new string[] {"GetFieldToFile", EncodeDdeArgument(categoryName), itemName, EncodeDdeArgument(fieldName), fileName }));
            return (retval == null) ? false : true;
        }

        /// <inheritdoc />
        public int GetFormCount(string categoryName)
        {
            return GetDDECount(new string[] { "GetFormCount", EncodeDdeArgument(categoryName) });
        }

        /// <inheritdoc />
        public string GetFormNames(string categoryName, string delim = null)
        {
            if (delim == null)
            {
                return GetDDEValues(new string[] { "GetFormNames", EncodeDdeArgument(categoryName) });
            }
            else
            {
                return GetDDEValues(new string[] { "GetFormNames", EncodeDdeArgument(categoryName), delim });
            }
        }

        /// <inheritdoc />
        public List<string> GetFormNames(string categoryName)
        {
            return GetDDEValuesAsList(new string[] { "GetFormNames", EncodeDdeArgument(categoryName), CMC_DELIM});
        }

        /// <inheritdoc />
        public int GetImageFieldCount(string categoryName)
        {
            return GetDDECount(new string[] { "GetImageFieldCount", EncodeDdeArgument(categoryName) });
        }

        /// <inheritdoc />
        public string GetImageFieldNames(string categoryName, string delim = null)
        {
            if (delim == null)
            {
                return GetDDEValues(new string[] { "GetImageFieldNames", EncodeDdeArgument(categoryName) });
            }
            else
            {
                return GetDDEValues(new string[] { "GetImageFieldNames", EncodeDdeArgument(categoryName), delim });
            }
        }
        /// <inheritdoc />
        public List<string> GetImageFieldNames(string categoryName)
        {
            return GetDDEValuesAsList(new string[] { "GetImageFieldNames", EncodeDdeArgument(categoryName), CMC_DELIM });
        }

        /// <inheritdoc />
        public bool GetImageFieldToFile(string categoryName, string itemName, string fieldName, string fileName)
        {
            object retval = DDERequest(BuildDDERequestCommand(new string[] { "GetImageFieldToFile", EncodeDdeArgument(categoryName), itemName, EncodeDdeArgument(fieldName), fileName }));
            return (retval == null) ? false : true;
        }

        /// <inheritdoc />
        public int GetItemCount(string categoryName)
        {
            return GetDDECount(new string[] { "GetItemCount", EncodeDdeArgument(categoryName) });
        }

        
        /// <inheritdoc />
        public string GetItemNames(string categoryName, string delim = null)
        {
            if (delim == null)
            {
                return GetDDEValues(new string[] { "GetItemNames", EncodeDdeArgument(categoryName) });
            }
            else
            {
                return GetDDEValues(new string[] { "GetItemNames", EncodeDdeArgument(categoryName), delim });
            }
        }
        /// <inheritdoc />
        public List<string> GetItemNames(string categoryName)
        {
            return GetDDEValuesAsList(new string[] { "GetItemNames", EncodeDdeArgument(categoryName), CMC_DELIM });
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
            return (retval == null) ? false : true;
        }

        /// <inheritdoc />
        public string GetMeField(string fieldName)
        {
            return GetDDEValues(new string[] { "GetMeField", EncodeDdeArgument(fieldName) });
        }

        /// <inheritdoc />
        public string GetNameField(string categoryName)
        {
            string retval = null;
            try
            {
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
            Preferences pref = new Preferences();
            string dde = BuildDDERequestCommand(new string[] { "GetPreference", "Me", CMC_DELIM });
            string[] me = DDERequest(dde).Split(new string[] { CMC_DELIM }, StringSplitOptions.None);
            if (me != null)
            {
                pref.MeCategory = me[0];
                pref.Me = me[1];
            }
            pref.LetterLogDir = GetDDEValues(new string[] { "GetPreference", "LetterlogDir" });
            pref.ExternalDir = GetDDEValues(new string[] { "GetPreference", "ExternalDir" });

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
            if (delim == null)
            {
                return GetDDEValues(new string[] { "GetTriggerNames" });
            }
            else
            {
                return GetDDEValues(new string[] { "GetTriggerNames", delim });
            }
        }

        /// <inheritdoc />
        public List<string> GetTriggerNames()
        {
            return GetDDEValuesAsList(new string[] { "GetTriggerNames", CMC_DELIM });
        }

        /// <inheritdoc />
        public int GetViewCount(string categoryName)
        {
            return GetDDECount(new string[] { "GetViewCount", EncodeDdeArgument(categoryName), CMC_DELIM });
        }

        /// <inheritdoc />
        public IViewDef GetViewDefinition(string viewName)
        {
            ViewDef vd = null;
            // note that this request is undocumented by Commence!
            string dde = BuildDDERequestCommand(new string[] { "GetViewDefinition", viewName, CMC_DELIM });
            string viewInfo = DDERequest(dde);
            if (viewInfo != null)
            {
                string[] buffer = viewInfo.Split(new string[] { CMC_DELIM }, StringSplitOptions.None);
                vd = new ViewDef();
                vd.Name = buffer[0];
                vd.TypeDescription = buffer[1];
                vd.Category = buffer[2];
                vd.FileName = buffer[3];
                vd.Type = Utils.GetValueFromEnumDescription<CommenceViewType>(vd.TypeDescription);
            }
            return vd;
        }

        /// <inheritdoc />
        public string GetViewNames(string categoryName, string delim = null)
        {
            if (delim == null)
            {
                return GetDDEValues(new string[] { "GetViewNames", EncodeDdeArgument(categoryName) });
            }
            else
            {
                return GetDDEValues(new string[] { "GetViewNames", EncodeDdeArgument(categoryName), delim });
            }
        }

        /// <inheritdoc />
        public List<string> GetViewNames(string categoryName)
        {
            return GetDDEValuesAsList(new string[] { "GetViewNames", EncodeDdeArgument(categoryName), CMC_DELIM });
        }

        /// <inheritdoc />
        public bool GetViewToFile(string viewName, string fileName, CmcLinkMode linkMode = CmcLinkMode.None, string param1 = null, string param2 = null)
        {
            object retval = DDERequest(BuildDDERequestCommand(new string[] { "GetViewToFile", viewName, ((int)linkMode).ToString(), param1, param2, fileName }));
            return (retval == null) ? false : true;
        }

        /// <inheritdoc />
        public string MarkActiveItem(string delim = null) // expects subsequent requests
        {
            // will fail if no item is active, for instance when no views are open
            // consumers should always check for null!
            if (delim != null)
            {
                return DDERequest(BuildDDERequestCommand(new string[] { "MarkActiveItem", delim }));
            }
            else
            {
                return DDERequest(BuildDDERequestCommand(new string[] { "MarkActiveItem"}));
            }
        }

        /// <inheritdoc />
        public string MergeTemplateCreate(string templateName, string categoryName, bool shared)
        {
            return GetDDEValues(new string[] { "MergeTemplateCreate", templateName, EncodeDdeArgument(categoryName), (shared) ? "1" : "0" });
        }

        /// <inheritdoc />
        public string MergeTemplateSave(string templateName, bool shared)
        {
            return GetDDEValues(new string[] { "MergeTemplateSave", templateName, (shared) ? "1" : "0" });
        }

        /// <inheritdoc />
        public string ViewCategory(string categoryName)  // expects subsequent requests
        {
            return GetDDEValues(new string[] { "ViewCategory", EncodeDdeArgument(categoryName) });
        }

        /// <inheritdoc />
        public string ViewConjunction(string AndOr12, string AndOr13, string AndOr34, string AndOr48, string AndOr56, string AndOr57, string AndOr78)
        {
             return GetDDEValues(new string[] { "ViewConjunction", AndOr12, AndOr13, AndOr34, AndOr48, AndOr56, AndOr57, AndOr78 });
        }

        /// <inheritdoc />
        public int ViewConnectedCount(int index, string connectionName, string connCategory)
        {
            return GetDDECount(new string[] { "ViewConnectedCount", index.ToString(), EncodeDdeArgument(connectionName), EncodeDdeArgument(connCategory) });
        }

        /// <inheritdoc />
        public string ViewConnectedField(int index, string connectionName, string connCategory, int connIndex, string fieldName)
        {
            return GetDDEValues(new string[] { "ViewConnectedField", index.ToString(), EncodeDdeArgument(connectionName), connCategory, connIndex.ToString(), EncodeDdeArgument(fieldName) });
        }

        /// <inheritdoc />
        public string ViewConnectedFields(int index, string connectionName, string connCategory, int connIndex, object[] fields, string delim =  null)
        {
            // we need fieldnames to be  comma delimited string, so we can feed it to a new string array
            // our fieldnames are trapped in object array args,
            // so we first cast that
            string[] fieldNames = ToStringArray(fields);
            // create a comma-delimited string from fieldnames and pass that into a new string[] array.
            // how is that for unnecessary overhead? :)
            // also note that we have to supply the number of fields we want. Intriguing tidbit.
            if (delim == null)
            {
                //return GetDDEValues(new string[] { "ViewConnectedFields", index.ToString(), EncodeDdeArgument(connectionName), connCategory, connIndex.ToString(), fieldNames.Length.ToString(), string.Join("\",\"", fieldNames) });
                return GetDDEValues(new string[] { "ViewConnectedFields", index.ToString(), EncodeDdeArgument(connectionName), EncodeDdeArgument(connCategory), connIndex.ToString(), fieldNames.Length.ToString(), string.Join(",", EncodeDdeArguments(fieldNames)) });
            }
            else
            {
                //return GetDDEValues(new string[] { "ViewConnectedFields", index.ToString(), EncodeDdeArgument(connectionName), connCategory, connIndex.ToString(), fieldNames.Length.ToString(), String.Join("\",\"", fieldNames), delim });
                return GetDDEValues(new string[] { "ViewConnectedFields", index.ToString(), EncodeDdeArgument(connectionName), EncodeDdeArgument(connCategory), connIndex.ToString(), fieldNames.Length.ToString(), string.Join(",", EncodeDdeArguments(fieldNames)), delim });
            }
        }

        /// <inheritdoc />
        [ComVisible(false)]
        public string[] ViewConnectedFields(int index, string connectionName, string connCategory, int connIndex, List<string> fields, string delim = null)
        {
            return ViewConnectedFields(index,connectionName, connCategory, connIndex, fields.ToArray<string>(), CMC_DELIM).Split(_splitter, StringSplitOptions.None);
        }

        /// <inheritdoc />
        public string ViewConnectedItem(int index, string connectionName, string connCategory, int connIndex)
        {
            return GetDDEValues(new string[] { "ViewConnectedItem", index.ToString(), EncodeDdeArgument(connectionName), connCategory, connIndex.ToString() });
        }

        /// <inheritdoc />
        public void ViewDeleteAllItems()
        {
            DDERequest(BuildDDERequestCommand(new string[] { "ViewDeleteAllItems" }));
        }

        /// <inheritdoc />
        public string ViewField(int index, string fieldName)
        {
            return GetDDEValues(new string[] { "ViewField", index.ToString(), EncodeDdeArgument(fieldName) });
        }

        /// <inheritdoc />
        public string ViewFields(int index, object[] fields, string delim = null)
        {
            // we need fieldnames to be  comma delimited string, so we can feed it to a new string array
            // our fieldnames are trapped in object array args,
            // so we first cast that
            string[] fieldNames = ToStringArray(fields);
            // create a comma-delimited string from fieldnames and pass that into a new string[] array.
            // how is that for unnecessary overhead? :)
            // also note that we have to supply the number of fields we want. Intriguing tidbit.
            if (delim == null)
            {
                //return GetDDEValues(new string[] { "ViewFields", index.ToString(), fieldNames.Length.ToString(), string.Join("\",\"", fieldNames) });
                return GetDDEValues(new string[] { "ViewFields", index.ToString(), fieldNames.Length.ToString(), string.Join(",", EncodeDdeArguments(fieldNames)) });
            }
            else
            {
                //return GetDDEValues(new string[] { "ViewFields", index.ToString(), fieldNames.Length.ToString(), string.Join("\",\"", fieldNames), delim });
                return GetDDEValues(new string[] { "ViewFields", index.ToString(), fieldNames.Length.ToString(), string.Join(",", EncodeDdeArguments(fieldNames)), delim });
            }
        }

        /// <inheritdoc />
        /// 
        [ComVisible(false)]
        public string[] ViewFields(int index, List<string> fields, string delim = null)
        {
            return ViewFields(index, fields.ToArray<string>(),CMC_DELIM).Split(_splitter,StringSplitOptions.None);
        }

        /// <inheritdoc />
        public bool ViewFieldToFile(int index, string fieldName, string fileName)
        {
            object retval = DDERequest(BuildDDERequestCommand(new string[] { "ViewFieldToFile", index.ToString(), EncodeDdeArgument(fieldName), fileName }));
            return (retval == null) ? false : true;
        }

        /// <inheritdoc />
        public bool ViewFilter(int clauseNumber, string filterType, bool notFlag, object args)
        {
            string[] fltParams = ToStringArray(args);
            //object retval = DDERequest(buildDDERequestCommand(new string[] { "ViewFilter", clauseNumber.ToString(), filterType, (notFlag) ? "NOT" : "", String.Join("\",\"", fltParams) }));
            object retval = DDERequest(BuildDDERequestCommand(new string[] { "ViewFilter", clauseNumber.ToString(), filterType, (notFlag) ? "NOT" : "", String.Join(",", EncodeDdeArguments(fltParams)) }));
            return (retval == null) ? false : true;
        }

        /// <inheritdoc />
        public bool ViewImageFieldToFile(int index, string fieldName, string fileName)
        {
            object retval = DDERequest(BuildDDERequestCommand(new string[] { "ViewImageFieldToFile", index.ToString(), EncodeDdeArgument(fieldName), fileName }));
            return (retval == null) ? false : true;
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
            return (retval == null) ? false : true;
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
            return (retval == null) ? false : true;
        }

        /// <inheritdoc />
        public bool ViewSort(string fieldName1, string sortOrder1, string fieldName2 = "", string sortOrder2 = "", string fieldName3 = "", string sortOrder3 = "", string fieldName4 = "", string sortOrder4 = "")
        {
            object retval = DDERequest(BuildDDERequestCommand(new string[] { "ViewSort", fieldName1, sortOrder1, fieldName2, sortOrder2, fieldName3, sortOrder3, fieldName4, sortOrder4 }));
            return (retval == null) ? false : true;
        }

        /// <inheritdoc />
        public bool ViewView(string viewName = "")
        {
            object retval = DDERequest(BuildDDERequestCommand(new string[] { "ViewView", viewName }));
            return (retval == null) ? false : true;
        }

        #endregion

        #region DDE Execute methods

        /// <inheritdoc />
        public bool AddItem(string categoryName, string itemName, string clarifyValue = "")
        {
            return DDEExecute(BuildDDERequestCommand(new string[] { "AddItem", EncodeDdeArgument(categoryName), itemName, clarifyValue }));
        }

        /// <inheritdoc />
        public bool AddSharedItem(string categoryName, string itemName, string clarifyValue = "")
        {
            return DDEExecute(BuildDDERequestCommand(new string[] { "AddSharedItem", EncodeDdeArgument(categoryName), itemName, clarifyValue }));
        }

        /// <inheritdoc />
        public bool AppendText(string categoryName, string itemName, string fieldName, string text)
        {
            return DDEExecute(BuildDDERequestCommand(new string[] { "AppendText", EncodeDdeArgument(categoryName), itemName, EncodeDdeArgument(fieldName), text }));
        }

        /// <inheritdoc />
        public bool AssignConnection(string categoryName, string itemName, string connectionName, string connCategory, string connItem)
        {
            return DDEExecute(BuildDDERequestCommand(new string[] { "AssignConnection", EncodeDdeArgument(categoryName), itemName, EncodeDdeArgument(connectionName), connCategory, connItem }));
        }

        /// <inheritdoc />
        public bool CheckInFormScript(string categoryName, string formName, string fileName)
        {
            return DDEExecute(BuildDDERequestCommand(new string[] { "CheckInFormScript", EncodeDdeArgument(categoryName), formName, fileName }));
        }

        /// <inheritdoc />
        public bool CheckOutFormScript(string categoryName, string formName, string fileName)
        {
            return DDEExecute(BuildDDERequestCommand(new string[] { "CheckOutFormScript", EncodeDdeArgument(categoryName), formName, fileName }));
        }

        /// <inheritdoc />
        public bool DeleteItem(string categoryName, string itemName)
        {
            return DDEExecute(BuildDDERequestCommand(new string[] { "DeleteItem", EncodeDdeArgument(categoryName), itemName }));
        }

        /// <inheritdoc />
        public bool DeleteView(string viewName)
        {
            return DDEExecute(BuildDDERequestCommand(new string[] { "DeleteView", viewName }));
        }

        /// <inheritdoc />
        public bool EditItem(string categoryName, string itemName, string fieldName, string fieldValue)
        {
            return DDEExecute(BuildDDERequestCommand(new string[] { "EditItem", EncodeDdeArgument(categoryName), itemName, EncodeDdeArgument(fieldName), fieldValue }));
        }

        /// <inheritdoc />
        public bool FireTrigger(string trigger, object[] args)
        {
            string[] trgParams = ToStringArray(args);
            if (trgParams.Length > 9)
            {
                this.LastError = "Too many parameters for FireTrigger.";
                return false;
            }
            //return DDEExecute(buildDDERequestCommand(new string[] { "FireTrigger", trigger, string.Join("\",\"", trgParams) }));
            return DDEExecute(BuildDDERequestCommand(new string[] { "FireTrigger", trigger, string.Join(",", EncodeDdeArguments(trgParams)) }));
        }

        /// <inheritdoc />
        public bool LogPhoneCall(object[] args)
        {
            string[] ciPairs = ToStringArray(args);
            //return DDEExecute(buildDDERequestCommand(new string[] { "LogPhoneCall", string.Join("\",\"", ciPairs) }));
            return DDEExecute(BuildDDERequestCommand(new string[] { "LogPhoneCall", string.Join(",", EncodeDdeArguments(ciPairs)) }));
        }

        /// <inheritdoc />
        public bool PromoteItemToShared(string categoryName, string itemName)
        {
            return DDEExecute(BuildDDERequestCommand(new string[] { "PromoteItemToShared", EncodeDdeArgument(categoryName), itemName }));
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
                return DDEExecute(BuildDDERequestCommand(new string[] { "ShowItem", EncodeDdeArgument(categoryName), itemName}));
            }
            else
            {
                return DDEExecute(BuildDDERequestCommand(new string[] { "ShowItem", EncodeDdeArgument(categoryName), itemName, formName }));
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
            return DDEExecute(BuildDDERequestCommand(new string[] { "UnassignConnection", EncodeDdeArgument(categoryName), itemName, EncodeDdeArgument(connectionName), connCategory, connItem }));
        }

        /// <inheritdoc />
        public bool HasDuplicates(string categoryName, string fieldName, bool caseSensitive = true)
        {
            using (ICommenceCursor cur = this.GetCursor(categoryName, CmcCursorType.Category, CmcOptionFlags.All))
            {
                return cur.HasDuplicates(fieldName, caseSensitive);
            }
        }

        /// <inheritdoc />
        public bool FieldValueExists(string categoryName, string fieldName, string fieldValue, bool caseSensitive = true)
        {
            using (ICommenceCursor cur = this.GetCursor(categoryName))
            {
                ICursorFilters filters = new CursorFilters(cur);
                ICursorFilterTypeF f = filters.Add(1, FilterType.Field);
                f.FieldName = fieldName;
                f.FieldValue = fieldValue;
                f.Qualifier = FilterQualifier.EqualTo;
                if (caseSensitive) { f.MatchCase = true; }
                if (filters.Apply() > 0) { return true;}
            }
            return false;
        }
        #endregion

        #region helper methods
        private List<string> GetDDEValuesAsList(string[] args)
        {
            string values = DDERequest(BuildDDERequestCommand(args));
            if (values == null) // an error occurred, return error value
            {
                List<string> retval = new List<string>();
                retval.Add(this.GetLastError());
                return retval;
            }
            // the delimiter is always the last element
            string[] splitter = new string[] { args.Last() };
            return values.Split(splitter, StringSplitOptions.None).ToList<string>();
        }

        private string GetDDEValues(string[] args)
        {
            string retval = string.Empty;
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
            return (result == null) ? -1 : Convert.ToInt32(result);
        }

        /// <summary>
        /// Builds the DDE request string for Commence.
        /// </summary>.
        /// <param name="args">Request item plus optional parameters. The request item string must always be the first element.</param>
        /// <returns>string in format "[Request command(param1, param2, paramN)]".</returns>
        private string BuildDDERequestCommand(string[] args)
        {
            StringBuilder sb = null;
            sb = new StringBuilder("[" + args[0]);

            if (args.Length == 1)// request item without additional parameters
            {
                sb.Append("()]");
                return sb.ToString();
            }
            //sb.Append("(\"" + String.Join("\",\"", args.Skip(1)) + "\")]"); // note the skip of the first argument
            sb.Append("(");
            sb.Append(String.Join(",",args.Skip(1)));
            sb.Append(")]"); // note the skip of the first argument
            return sb.ToString();
        }

        /// <summary>
        /// Encode array of arguments.
        /// </summary>
        /// <param name="args">string array.</param>
        /// <returns>Encoded array of arguments.</returns>
        private IEnumerable<string> EncodeDdeArguments(IEnumerable<string> args)
        {
            foreach (string arg in args)
            {
                string retval = EncodeDdeArgument(arg);
                yield return retval;
            }
        }

        /// <summary>
        /// Encodes the DDE arguments so that they will be processed correctly
        /// </summary>
        /// <param name="arg"></param>
        /// <returns>DDE-safe argument string.</returns>
        /// <remarks>This method is not exhaustive - some combinations cannot be handled.</remarks>
        private string EncodeDdeArgument(string arg)
        {
            // By default, arguments used in a DDE request will be double quoted
            // Commence does not require that they are, it is just a little safer
            // for example, when a fieldname contains a comma, the DDE request would trip unless the argument is double-quoted
            // however, if the argument contains an embedded 'control' character,
            // we need to make special arrangements
            string retval = arg.EncloseWithChar('\"');

            // when the argument itself is quoted we can get away by adding additional quotes
            if (arg.StartsWith("\"") && arg.EndsWith("\""))
            {

                return arg.EncloseWithChar('\"', 2);
            }
            // when the argument contains both a double-quote and a comma, we're in trouble
            if (arg.Contains('\"') && arg.Contains(','))
            {
                // I do not yet know how to deal with this, if at all possible
                // TODO even while this is very rare, it needs a solution
            }
            // when the argument contains a double-quote, forego double-quoting and just return the string
            if (arg.Contains('\"'))
            {
                return arg;
            }
            return retval;
        }

        /// <summary>
        /// Creates object array from string array.
        /// </summary>
        /// <param name="input">string array</param>
        /// <returns>object array that can be consumed by COM clients such as VBScript.</returns>
        private object[] toObjectArray(string[] input)
        {
            object[] objArray = new object[input.Length];
            input.CopyTo(objArray, 0);
            return objArray;
        }
        /// <summary>
        /// Creates string array from object.
        /// </summary>
        /// <param name="arg">object</param>
        /// <returns>string array</returns>
        private string[] ToStringArray(object arg)
        {
            var collection = arg as System.Collections.IEnumerable;
            if (collection != null)
            {
                return collection
                  .Cast<object>()
                  .Select(x => x.ToString())
                  .ToArray();
            }

            if (arg == null)
            {
                return new string[] { };
            }

            return new string[] { arg.ToString() };
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
                DDETimer = new Timer(DDETimeout);
                DDETimer.AutoReset = false;
                DDETimer.Elapsed += _conv.HandleTimerElapsed;
            }
            DDETimer.Interval = DDETimeout;
            DDETimer.Enabled = true; // (re)start timer

            try
            {
                this.LastError = string.Empty; // clear GetLastError
                retval = _conv.DDERequest(ddeRequestCommand);
            }
            catch (CommenceDDEException e)
            {
                this.LastError = e.Message; // store the last error
                //retval = null; // if null, we know an exception occurred without having to (re)throw it
                throw; // it is better to rethrow than to swallow
            }
            return retval;
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
                DDETimer = new Timer(DDETimeout);
                DDETimer.AutoReset = false;
                DDETimer.Elapsed += _conv.HandleTimerElapsed;
            }
            DDETimer.Interval = DDETimeout; // (re)set interval
            DDETimer.Enabled = true; // (re)start timer

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

        private string GetFieldTypeAsString(CommenceFieldType ft)
        {
            return ft.GetEnumDescription();
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

        #region Properties
        /// <summary>
        /// Last DDE error received from Commence.
        /// </summary>
        private string LastError { get; set; }

        #endregion
    }
}



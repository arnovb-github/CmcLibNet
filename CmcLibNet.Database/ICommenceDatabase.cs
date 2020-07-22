using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Vovin.CmcLibNet.Database.Metadata;

namespace Vovin.CmcLibNet.Database
{
    /// <summary>
    /// Exposes all 'native' Commence API database functionality, including all functionality provided by the DDE interface (ICommenceConversation).
    /// There is no need to use DDE syntax with CmcLibNet, all DDE plumbing is hidden from the end user.
    /// <para>You cannot get a reference to FormOA.ICommenceConversation from CmcLibNet.</para>
    /// <para>COM clients can reference this interface by using the ProgId <c>CmcLibNet.Database</c>.</para>
    /// VBScript:
    /// <code language="vbscript">
    /// Dim db : Set db = CreateObject("CmcLibNet.Database")
    /// '.. do stuff ..
    /// db.Close
    /// Set db = Nothing</code>
    /// <para>When used fom a Commence Item Detail Form or Office VBA, be sure to read up on the <see cref="Close()"/> method.</para>
    /// </summary>
    /// <remarks>A lot of the documentation on this interface comes 'verbatim' from the Commence help files.</remarks>
    [ComVisible(true)]
    [Guid("34A2D34C-8C7E-4fae-86CC-A7F75BBC8B45")]
    public interface ICommenceDatabase : IDisposable
    {
        /// <summary>
        /// Create a cursor.
        /// </summary>
        /// <param name="category">Commence category name.</param>
        /// <param name="pCursorType">Cursor type.</param>
        /// <param name="pOptionFlags">Cursor flags.</param>
        /// <returns>ICommenceCursor.</returns>
        /// <remarks>Note the optional parameters
        /// this allows COM clients to use just the category name to get a default cursor
        /// but also not that because of the optional paramaters, which must come at the end,
        /// the order is different from what Commence developers are used to.
        /// </remarks>
        Database.ICommenceCursor GetCursor(string category, CmcCursorType pCursorType = 0, CmcOptionFlags pOptionFlags = CmcOptionFlags.Default);

        /* for the methods that return object arrays, note that that is because
         * VBSCript, VB6 and VBA only accept Variant arrays
         * e.g. returning string[] getSomething() will not work in VBScript!
         * There is the MarshalAsAttribute, but I can't get it to work like I want to :/
         */

        #region DDE options
        /// <summary>
        /// Make assembly try to string-escape DDE arguments. Default is <c>true</c>.
        /// </summary>
        /// <remarks>By default, all arguments passed to DDE requests will be enclosed in double-quotes.
        /// Setting this to false allows you to have fine-grained control over the format of the parameters.</remarks>
        bool EncodeDDEArguments { get; set; }
        #endregion

        #region Commence DDE Request commands
        /// <summary>
        /// Specifies whether Commence should return clarified item names, or leave bStatus empty to get the current status.
        /// </summary>
        /// <param name="bStatus">"True" to set clarified itemnames, "False" for unclarified, leave empty to request current status.</param>
        /// <returns>"OK" if successful, inspect <see cref="GetLastError" /> on failure.
        /// <para>If bStatus was omitted this method returns "True" if clarified, "False" if not.</para></returns>
        /// <seealso cref="GetLastError"/>
        string ClarifyItemNames(string bStatus = null); // can't do bool? because COM doesn't like Nullable types

        /// <summary>
        /// Gets the category of the active item or view.
        /// </summary>
        /// <returns>Active category or <c>null</c> if no active item or view could be retrieved.</returns>
        string GetActiveCategory();

        /// <summary>
        /// Gets the itemname of the active item. Use <see cref="MarkActiveItem"/> to try and get a clariefied itemname.
        /// </summary>
        /// <returns>Active itemname, <c>null</c> if no active item could be retieved.</returns>
        string GetActiveItemName();

        /// <summary>
        /// Returms information on the currently active view (if any).
        /// </summary>
        /// <returns>ActiveViewInfo object holding information on the view, <c>null</c> if no view active.</returns>
        /// <remarks>Unlike it's Commence counterpart, this method returns a <see cref="IActiveViewInfo"/> object instead of a delimited string.
        /// Returns <c>null</c> if no view is active.</remarks>
		IActiveViewInfo GetActiveViewInfo();

        /// <summary>
        /// Returns a list of all items in the Category having the given phone number in a phone field.
        /// If Category is blank, searches all categories in the currently active database
        /// </summary>
        /// <param name="categoryName">Commence category name.</param>
        /// <param name="phoneNumber">Phonenumber to search.</param>
        /// <param name="delim">Delimiter, up to 8 chars.</param>
        /// <returns>String of matching items delimited by delim of CR/LF if no delim specified.</returns>
		string GetCallerID(string categoryName, string phoneNumber, string delim = null);

        /// <summary>
        /// Returns a list of all items in the Category having the given phone number in a phone field.
        /// If Category is blank, searches all categories in the currently active database
        /// </summary>
        /// <param name="categoryName">Commence category name.</param>
        /// <param name="phoneNumber">Phonenumber to search.</param>
        /// <returns>List of matching items.</returns>
        [ComVisible(false)]
        List<string> GetCallerID(string categoryName, string phoneNumber);

        /// <summary>
        /// Returns number of categories in the Commence database.
        /// </summary>
        /// <returns>Category count, -1 on error.</returns>
		int GetCategoryCount();

        /// <summary>
        /// Returns information on the definition of the specified category.
        /// </summary>
        /// <param name="categoryName">Commence category name.</param>
        /// <returns>CategoryDef object exposing properties of the category.</returns>
        /// <remarks>Unlike it's Commence counterpart, this method returns a <see cref="ICategoryDef"/> object instead of a delimited string. Returns <c>null</c> if category doesn't exist.</remarks>
		ICategoryDef GetCategoryDefinition(string categoryName);

        /// <summary>
        /// Gets a list of category names.
        /// </summary>
        /// <param name="delim">Delimiter, up to 8 chars.</param>
        /// <returns>String containing the category names separated by delim, delim defaults to CR/LF.</returns>
		string GetCategoryNames(string delim = null);

        /// <summary>
        /// Gets a list of categories in the current Commence database.
        /// </summary>
        /// <returns>List of categories</returns>
        /// <remarks>This method is only available to .Net consumers.</remarks>
        [ComVisible(false)]
        List<string> GetCategoryNames();

        /// <summary>
        /// Gets the clarifyfield for the specified category.
        /// </summary>
        /// <param name="categoryName">Commence category name.</param>
        /// <returns>Clarify field name, <c>null</c> if no clarify field.</returns>
        string GetClarifyField(string categoryName);

        /// <summary>
        /// Gets the clarify separator for the specified category.
        /// </summary>
        /// <param name="categoryName">Commence category name.</param>
        /// <returns>Clarify separator, <c>null</c> if no separator.</returns>
        string GetClarifySeparator(string categoryName);

        /// <summary>
        /// Gets the number of connected items in connected category for the specified item
        /// </summary>
        /// <param name="categoryName">Commence category name.</param>
        /// <param name="itemName">Commence item name to count from.</param>
        /// <param name="connectionName">Commence connection name (case-sensitive).</param>
        /// <param name="connCategory">Commence category name of connCategory</param>
        /// <returns>Number of connected items, -1 on error.</returns>
		int GetConnectedItemCount(string categoryName, string itemName, string connectionName, string connCategory);

        /// <summary>
        /// Returns the indirect field value(s) in connCategory that itemName of the categoryName is connected to via Connection.
        /// The order of the items in the list is arbitrary (i.e., not sorted).
        /// </summary>
        /// <param name="categoryName">categoryName.</param>
        /// <param name="itemName">itemName (use clarified if possible).</param>
        /// <param name="connectionName">Connection name (case-sensitive).</param>
        /// <param name="connCategory">connCategory.</param>
        /// <param name="fieldName">FieldName.</param>
        /// <param name="delim">Delimiter, up to 8 chars.</param>
        /// <returns>String of connected item fields. Text limitations imposed by Commence:
        /// <list type="bullet">
        /// <item><description>Limit to 256 chars.</description></item>
        /// <item><description>First line of multiline field.</description></item>
        /// <item><description>Returns string "(none)" if there are no connected items.</description></item>
        /// </list>
        /// </returns>
		string GetConnectedItemField(string categoryName, string itemName, string connectionName, string connCategory, string fieldName, string delim = null);

        /// <summary>
        /// Returns the list of items in category that named item is connected to. The order of the items in the list is arbitrary (i.e., not sorted).
        /// </summary>
        /// <param name="categoryName">Commence category containing item.</param>
        /// <param name="itemName">Commence item name to retrieve connected names from.</param>
        /// <param name="connectionName">Commence connection name (case-sensitive).</param>
        /// <param name="connCategory">Commence connected category.</param>
        /// <param name="delim">Delimiter, up to 8 chars.</param>
        /// <returns>Connected item names, "(none)" if no connected items.</returns>
		string GetConnectedItemNames(string categoryName, string itemName, string connectionName, string connCategory, string delim = null);

        /// <summary>
        /// Gets a list of items connected to the specified item.
        /// </summary>
        /// <param name="categoryName">Commence categoryname.</param>
        /// <param name="itemName">Itemname.</param>
        /// <param name="connectionName">Connection name (case-sensitive).</param>
        /// <param name="connCategory">Connected categoryname.</param>
        /// <returns>List of connected items.</returns>
        /// <remarks>This method is only available to .Net consumers.</remarks>
        /// <seealso cref="ClarifyItemNames"/>
        [ComVisible(false)]
        List<string> GetConnectedItemNames(string categoryName, string itemName, string connectionName, string connCategory);

        /// <summary>
        /// Gets the number of connections from the specified category.
        /// </summary>
        /// <param name="categoryName">Commence category to count connections from.</param>
        /// <returns>Number of connections.</returns>
		int GetConnectionCount(string categoryName);

        /// <summary>
        /// Gets the connection names for the specified category.
        /// </summary>
        /// <param name="categoryName">Commence category name.</param>
        /// <param name="delim1">First delimiter, up to 8 chars.</param>
        /// <param name="delim2">Second delimiter, up to 8 chars.</param>
        /// <returns>Array of <see cref="ICommenceConnection"/> objects</returns>
        //[return: MarshalAs(UnmanagedType.Struct, SafeArraySubType = VarEnum.VT_ARRAY)]
        object GetConnectionNames(string categoryName, string delim1 = null, string delim2 = null);

        /// <summary>
        /// Gets a list of connection names to the specified category.
        /// </summary>
        /// <param name="categoryName">Commence category name.</param>
        /// <returns>List of CommenceConnection objects.</returns>
        /// <remarks>This method is only available to .Net consumers.</remarks>
        [ComVisible(false)]
        IEnumerable<ICommenceConnection> GetConnectionNames(string categoryName);

        /// <summary>
        /// (Not implemented) Get the active Commence database name and path as delimited string. See <see cref="CommenceApp.Name"/> and/or <see cref="CommenceApp.Path"/>.
        /// </summary>
        /// <returns>NotImplementedException</returns>
        /// <seealso cref="CommenceApp.Name"/>
        /// <seealso cref="CommenceApp.Path"/>
        [Obsolete("Use CmcLibNet.CommenceApp.Name and/or CmcLibNet.CommenceApp.Path")]
		string GetDatabase();

        /// <summary>
        /// Queries Commence for information on the database definition.
        /// </summary>
        /// <returns>DBDef object that exposes Commence database definition properties, <c>null</c> on error.</returns>
        /// <remarks>Unlike it's Commence counterpart, this method returns a <see cref="IDBDef"/> object instead of a delimited string.</remarks>
		IDBDef GetDatabaseDefinition();

        /// <summary>
        /// Gets the number of Desktops in the Commence database.
        /// </summary>
        /// <returns>Number of desktops, -1 on error.</returns>
		int GetDesktopCount();

        /// <summary>
        /// Gets the desktop names in the Commence database.
        /// </summary>
        /// <param name="delim">Delimiter, up to 8 chars.</param>
        /// <returns>Strings containing the desktop names delimited by delim of CR/LF if no delim supplied.</returns>
		string GetDesktopNames(string delim = null);

        /// <summary>
        /// Get a list of desktopnames.
        /// </summary>
        /// <returns>List of desktopnames.</returns>
        /// <remarks>This method is only available to .Net consumers.</remarks>
        [ComVisible(false)]
        List<string> GetDesktopNames();

        /// <summary>
        /// Returns a fieldvalue from the specified Commence item.
        /// </summary>
        /// <param name="categoryName">Commence category name.</param>
        /// <param name="itemName">Commence item name.</param>
        /// <param name="fieldName">Commence field name.</param>
        /// <returns>FieldValue, "(Active item not found)" on error.</returns>
		string GetField(string categoryName, string itemName, string fieldName);

        /// <summary>
        /// Gets the number of fields in the specified category.
        /// </summary>
        /// <param name="categoryName">Commence category name</param>
        /// <returns>Number of fields, -1 on error.</returns>
		int GetFieldCount(string categoryName);

        /// <summary>
        /// Queries Commence for information on the specified field definition.
        /// </summary>
        /// <param name="categoryName">Commence category name.</param>
        /// <param name="fieldName">Commence field name.</param>
        /// <returns>FieldDef object exposing properties of the field, <c>null</c> on error.</returns>
        /// <remarks>Unlike it's Commence counterpart, this method returns a <see cref="ICommenceFieldDefinition"/> object instead of a delimited string.</remarks>
		ICommenceFieldDefinition GetFieldDefinition(string categoryName, string fieldName);

        /// <summary>
        /// Gets the fieldnames for the specified category.
        /// </summary>
        /// <param name="categoryName">Commence category name.</param>
        /// <param name="delim">Delimiter, up to 8 chars.</param>
        /// <remarks>Does not return Image field names, see <see cref="GetImageFieldNames(string)"/>.</remarks>
        /// <returns>string containing the field names delimited by delimiter delim, or CR/LF if delim is null.</returns>
		string GetFieldNames(string categoryName, string delim = null);

        /// <summary>
        /// Gets a list of fieldnames for the specified Commence category.
        /// </summary>
        /// <remarks>Does not return Image field names, see <see cref="GetImageFieldNames(string)"/>.</remarks>
        /// <param name="categoryName">Commence categoryname.</param>
        /// <returns>List of fieldnames.</returns>
        /// <remarks>This method is only available to .Net consumers.</remarks>
        [ComVisible(false)]
        List<string> GetFieldNames(string categoryName);

        /// <summary>
        /// Gets values of specified field(s) from specified item.
        /// </summary>
        /// <param name="categoryName">Commence category name.</param>
        /// <param name="itemName">Commence item name.</param>
        /// <param name="fields">object array containing fieldnames. .Net consumers can simply supply a string[], from VBScript/VBA and such use Array(field1, field2,...fieldN)</param>
        /// <param name="delim">Delimiter, up to 8 chars.</param>
        /// <returns>String with the fieldvalues of requested fields delimited by delim or CR/LF if delim not supplied.</returns>
        /// <remarks>This method may prove a little problematic.
        /// It takes a list of fieldnames; because we do not know in advance what fields will be specified,
        /// you have to pass in an array as parameter, however,
        /// for this to work from COM, this array has to be of type 'object'.
        /// Note that this function is obsolete in the sense that this can be done much faster using CommenceQueryRowSet
        /// It can be useful however for getting fieldvalues of the active item.
        /// 
        /// sample usage from VBA:
        /// <code>GetFields("categoryName", "itemName", Array(fieldname1, fieldname2,..., fieldnameN))</code>
        /// </remarks>
		string GetFields(string categoryName, string itemName, object[] fields, string delim = null);

        /// <summary>
        /// Gets values of specified field(s) from specified item.
        /// </summary>
        /// <param name="categoryName">Commence category name.</param>
        /// <param name="itemName">Commence item name.</param>
        /// <param name="fieldNames">List of fields to request values from.</param>
        /// <returns>List of fieldvalues.</returns>
        [ComVisible(false)]
        List<string> GetFields(string categoryName, string itemName, List<string> fieldNames);

        /// <summary>
        /// Writes a single fieldvalue to file.
        /// </summary>
        /// <param name="categoryName">Commence category name.</param>
        /// <param name="itemName">Commence item name.</param>
        /// <param name="fieldName">Commence field name.</param>
        /// <param name="fileName">fully qualified output file. If the file does not exist, Commence will create it; the drive and directory, however, must already exist. If the file exists, it will be overwritten.</param>
        /// <returns><c>true</c> on success, inspect <see cref="GetLastError" /> on failure.</returns>
		bool GetFieldToFile(string categoryName, string itemName, string fieldName, string fileName);

        /// <summary>
        /// Gets the number of Item Detail Forms for the specified category.
        /// </summary>
        /// <param name="categoryName">Commence category name.</param>
        /// <returns>Number of Item Detail Forms, -1 on error.</returns>
		int GetFormCount(string categoryName);

        /// <summary>
        /// Gets the Item detail Form names for the specified category
        /// </summary>
        /// <param name="categoryName">Commence category name.</param>
        /// <param name="delim">Delimiter, up to 8 chars.</param>
        /// <returns>String containing the Item Detail Form names delimited by delim, or CR/LF if not supplied.</returns>
        string GetFormNames(string categoryName, string delim = null);

        /// <summary>
        /// Gets a list of all Item Detail Form names in specified category.
        /// </summary>
        /// <param name="categoryName">Commence categoryname.</param>
        /// <returns>List of formnames.</returns>
        /// <remarks>This method is only available to .Net consumers.</remarks>
        [ComVisible(false)]
        List<string> GetFormNames(string categoryName);

        /// <summary>
        /// Gets the number of image fields in specified category.
        /// </summary>
        /// <param name="categoryName">Commence category name.</param>
        /// <returns>Number of image fields, -1 on error.</returns>
		int GetImageFieldCount(string categoryName);

        /// <summary>
        /// Gets the image-type field names for the specified category.
        /// </summary>
        /// <param name="categoryName">Commence category name.</param>
        /// <param name="delim">Delimiter, up to 8 chars.</param>
        /// <returns>String containing the image field names delimited by delim or CR/LF if not supplied.</returns>
  		string GetImageFieldNames(string categoryName, string delim = null);

        /// <summary>
        /// Gets a list of all field of type 'image' from the specified category.
        /// </summary>
        /// <param name="categoryName">Commence categoryname.</param>
        /// <returns>List of imagefield names</returns>
        /// <remarks>This method is only available to .Net consumers.</remarks>
        [ComVisible(false)]
        List<string> GetImageFieldNames(string categoryName);

        /// <summary>
        /// Copies the image field identified by the named Category, Item, and Field to the specified Filename.
        /// Only bitmap (.BMP) format image fields are currently supported.
        /// Filename must be a fully qualified path including the drive letter (e.g. c:\tmp\data.bmp).
        /// If the file does not exist, Commence will create it; the drive and directory, however, must already exist.
        /// If the file exists, it will be overwritten. It is the client’s responsibility to perform the necessary file cleanup.
        /// If the image field is blank, no file is written.
        /// </summary>
        /// <param name="categoryName">Commence category name</param>
        /// <param name="itemName">Commence item name</param>
        /// <param name="fieldName">Commence imagefield name</param>
        /// <param name="fileName">fully qualified filename</param>
        /// <returns><c>true</c> on success, inspect <see cref="GetLastError" /> on failure.</returns>
        bool GetImageFieldToFile(string categoryName, string itemName, string fieldName, string fileName);

        /// <summary>
        /// Gets the number of items in the specified category.
        /// </summary>
        /// <param name="categoryName">Commence category name.</param>
        /// <returns>Number of items, -1 on error.</returns>
		int GetItemCount(string categoryName);

        /// <summary>
        /// Gets the item names for the specified category
        /// </summary>
        /// <param name="categoryName">Commence category name.</param>
        /// <param name="delim">Delimiter, up to 8 chars.</param>
        /// <returns>String containing the item names delimited by delim or CR/LF if not supplied.</returns>
        /// <remarks>Use clarified Itemnames whenever you can.</remarks>
        string GetItemNames(string categoryName, string delim = null);

        /// <summary>
        /// Gets a list of all itemnames in the specified category.
        /// </summary>
        /// <param name="categoryName">Commence categoryname.</param>
        /// <returns>List of itemnames.</returns>
        /// <remarks>This method is only available to .Net consumers.</remarks>
        /// <seealso cref="ClarifyItemNames"/>
        [ComVisible(false)]
        List<string> GetItemNames(string categoryName);

        /// <summary>
        /// Gets the last error that occurred during DDE conversation.
        /// Inspect this value if a call failed.
        /// </summary>
        /// <returns>NACK code. If the NACK error code is less than 100, then that number indicates the position of the parameter that was invalid.
        /// For example, if the EXECUTE command <c>[EditItem("Person", "John Doe", "Work Phone", "(123)555-1234")]</c> returned a NACK error code of 2, then perhaps a John Doe does not exist in the Person category.
        /// <list type="table">
        /// <listheader><term>NACK codes</term><description>Error description.</description></listheader>
        /// <item><term>100 (0x64)</term><description>Out of memory.</description></item>
        /// <item><term>101 (0x65)</term><description> Internal error.</description></item>
        /// <item><term>102 (0x66)</term><description> Wrong number of parameters.</description></item>
        /// <item><term>103 (0x67)</term><description> Unknown conversation topic.</description></item>
        /// <item><term>104 (0x68)</term><description> No Phone Log category has been set; returned by LogPhoneCall EXECUTE command. Use Customize-Preferences-Event Logs to select a category.</description></item>
        /// <item><term>105 (0x69)</term><description> No category selected; returned by ViewData REQUESTS not preceded by a ViewCategory.</description></item>
        /// <item><term>106 (0x6A)</term><description> Unsupported clipboard format.</description></item>
        /// <item><term>107 (0x6B)</term><description> Unsupported field type.</description></item>
        /// <item><term>108 (0x6C)</term><description> Field/qualifier mismatch; returned by ViewFilter REQUEST.</description></item>
        /// <item><term>109 (0x6D)</term><description> Unknown EXECUTE command.</description></item>
        /// <item><term>110 (0x6E)</term><description> Parsing error. Check syntax of DDE command.</description></item>
        /// <item><term>111 (0x6F)</term><description> Unknown REQUEST item or parsing error.</description></item>
        /// <item><term>112 (0x70)</term><description> Cannot add new item, category full; returned by AddItem, AddSharedItem and LogPhoneCall EXECUTE commands.</description></item>
        /// <item><term>113 (0x71)</term><description> File I/O error. Is the disk full?</description></item>
        /// <item><term>114 (0x72)</term><description> Connection already exists; returned by the AssignConnection EXECUTE command.</description></item>
        /// <item><term>115 (0x73)</term><description> No such connection exists; returned by the UnassignConnection EXECUTE commands.</description></item>
        /// <item><term>116 (0x74)</term><description> The category has been invalidated.</description></item>
        /// <item><term>117 (0x75)</term><description> The item has been invalidated.</description></item>
        /// <item><term>118 (0x76)</term><description> An "every" date was received, or is the default value for a date field. Date ranges are not supported via DDE; returned by the AddItem or EditItem EXECUTE commands.</description></item>
        /// <item><term>119 (0x77)</term><description> Non-unique item name; returned by the AddItem or EditItem EXECUTE commands.</description></item>
        /// <item><term>120 (0x78)</term><description> Parameter too long; returned by the AddItem, EditItem and AppendText EXECUTE commands.</description></item>
        /// <item><term>121 (0x79)</term><description> No active child window or child window is unsupported; returned by the GetActiveViewInfo and GetLetterViewInfo REQUESTS.</description></item>
        /// <item><term>122 (0x80)</term><description> Agents are currently disabled. Returned by FireTrigger.</description></item>
        /// <item><term>123 (0x81)</term><description> No item has been marked. Use ViewMarkItem or AddItem to mark an item.</description></item>
        /// <item><term>124 (0x82)</term><description> Incompatible image type. Returned by GetImageFieldToFile or ViewImageFieldToFile.</description></item>
        /// <item><term>125 (0x83)</term><description> Permission denied. Workgroup client does not have correct permission level.</description></item>
        /// <item><term>126 (0x84) </term><description> No (-Me-) item has been defined. Use Customize-Preferences-Personal Info.</description></item>
        /// <item><term>127 (0x85)</term><description> TAPI error encountered.</description></item>
        /// <item><term>128 (0x86)</term><description> An agent exists but it is currently inactive. Returned by FireTrigger.</description></item>
        /// <item><term>201 (0xC9)</term><description> Filter 1 has been invalidated; this may occur if a connCategory item is deleted/modified.</description></item>
        /// <item><term>202 (0xCA)</term><description> Filter 2 has been invalidated.</description></item>
        /// <item><term>203 (0xCB)</term><description> Filter 3 has been invalidated.</description></item>
        /// <item><term>204 (0xCC)</term><description> Filter 4 has been invalidated.</description></item>
        /// </list>
        /// </returns>
		string GetLastError();

        /// <summary>
        /// Returns information about the active view and selected letter template at the time the Tools-Send Letter or Customize-Database-Letter Template command was executed.
        /// This call is used by the Commence letter macros and has not been implemented here.
        /// </summary>
        /// <param name="delim">Delimiter.</param>
        /// <returns>NotImplementedException.</returns>
        /// <remarks>This method is not implemented.</remarks>
		string GetLetterViewInfo(string delim = null);

        /// <summary>
        /// "Marks" the specified item and makes it the default category and item for subsequent calls to:
        /// EditItem, DeleteItem, AppendText, AssignConnection, UnassignConnection, PromoteItemToShared and ShowItem EXECUTE commands, and GetField, GetFieldToFile, GetFields, GetConnectedItemCount and GetConnectedItemNames REQUESTS.
        /// The category and item is marked only for the specific DDE conversation, not globally. that is, each DDE conversation can have a different marked item.
        /// Can be used to mark the (-Me-) item if one is defined.
        /// Syntax: [GetMarkItem("(-Me-)")]
        /// Note: If a Clarify Field is defined for the category, the value of the clarify field can be specified with an optional third paramter. Alternatively, the ItemName can be a clarified item name.
        /// </summary>
        /// <param name="categoryNameOrKeyword">Commence category name</param>
        /// <param name="itemName">Commence item name</param>
        /// <param name="clarifyValue">Commence clarify value</param>
        /// <returns><c>true</c> on success, inspect <see cref="GetLastError" /> on failure.</returns>
		bool GetMarkItem(string categoryNameOrKeyword, string itemName, string clarifyValue);

        /// <summary>
        /// Gets a fieldvalue from the item defined as (-Me-) item in Commence.
        /// </summary>
        /// <param name="fieldName">Commence field name.</param>
        /// <returns>FieldValue, "(Field not found)" if field doesn't exist, "(ME item not found)" if no (-Me-) item is set. Or NACK 126?</returns>
		string GetMeField(string fieldName);

        /// <summary>
        /// Get the name of the namefield for the specified category.
        /// </summary>
        /// <param name="categoryName">Commence category name.</param>
        /// <returns>Fieldname of the Name field, <c>null</c> on error. (Does the category exist?)</returns>
        string GetNameField(string categoryName);

        /// <summary>
        /// TAPI must be configured for GetPhoneNumber to work properly.
        /// TAPI is ancient, so this method is not implemented.
        /// </summary>
        /// <param name="phoneNumber">phone number.</param>
        /// <param name="delim">Delimiter, up to 8 chars.</param>
        /// <returns>NotImplementedException</returns>
		string GetPhoneNumber(string phoneNumber, string delim = null);

        /// <summary>
        /// Get preference settings.
        /// </summary>
        /// <param name="preferenceSetting">Setting to retrieve. Valid settings are "Me", "MeCategory", "LetterLogDir", "ExternalDir"</param>
        /// <returns>Commence Preferences setting.</returns>
		string GetPreference(string preferenceSetting);

        /// <summary>
        /// Uses Commence preference information and reverses the name if appropriate. 
        /// </summary>
        /// <param name="itemName">Commence item name.</param>
        /// <param name="prefFlag"> If the <c>PrefFlag</c> parameter is 1, 
        /// then the Customize / Preferences / Other / Name Field setting is used to determine if names should be reversed.
        /// If the PrefFlag parameter is left blank or is set to 0,
        /// then the command uses the Reverse Name checkbox value of
        /// the latest invocation of Tools-Send Letter and reverses the name if appropriate.</param>
        /// <returns>Reverse name, if applicable.</returns>
		string GetReverseName(string itemName, int prefFlag);

        /// <summary>
        /// Gets the number of DDE agent triggers defined in the database.
        /// </summary>
        /// <returns>Number of items, -1 on error.</returns>
		int GetTriggerCount();

        /// <summary>
        /// Gets the trigger names defined in the database.
        /// </summary>
        /// <param name="delim">Delimiter, up to 8 chars.</param>
        /// <returns>String containing the trigger names delimited by delim or CR/LF if not supplied.</returns>
		string GetTriggerNames(string delim = null);

        /// <summary>
        /// Get a list of Agent triggers.
        /// </summary>
        /// <returns>List of agent triggers.</returns>
        /// <remarks>This method is only available to .Net consumers.</remarks>
        [ComVisible(false)]
        List<string> GetTriggerNames();

        /// <summary>
        /// Gets the number of views in the specified category.
        /// </summary>
        /// <param name="categoryName">Commence category name.</param>
        /// <returns>Number of views, -1 on error.</returns>
		int GetViewCount(string categoryName);

        /// <summary>
        /// Get definition of a view
        /// </summary>
        /// <param name="viewName">Commence view name, case-sensitive.</param>
        /// <returns>ViewDef object that exposes properties of the view, <c>null</c> on error.</returns>
        /// <exception cref="CommenceCOMException">View not found.</exception>
        /// <remarks>Unlike it's (undocumented) Commence counterpart, this method returns a <see cref="IViewDef"/> object instead of a delimited string. Returns <c>null</c> on error. You cannot use this method on Multiviews.</remarks>
        IViewDef GetViewDefinition(string viewName);

        /// <summary>
        /// Gets a list of view names for the specified category.
        /// </summary>
        /// <param name="categoryName">Commence category name.</param>
        /// <param name="delim">Delimiter, up to 8 chars.</param>
        /// <returns>String containing the view names delimited by delim or CR/LF if not supplied.</returns>
		string GetViewNames(string categoryName, string delim = null);

        /// <summary>
        /// Gets a list of all views for the specified category. Returns all categories if category is not specified.
        /// </summary>
        /// <param name="categoryName">Commence category name. Set to <c>String.Empty</c> (or it's equivalent for your language) to get all views in database.</param>
        /// <returns>List of viewnames.</returns>
        /// <remarks>This method is only available to .Net consumers.</remarks>
        [ComVisible(false)]
        List<string> GetViewNames(string categoryName);

        /// <summary>
        /// Gets the columnnames from a view provided a cursor can be created on them.
        /// </summary>
        /// <param name="viewName">Commence view name (case-sensitive).</param>
        /// <param name="flags"></param>
        /// <returns>List of fieldnames or columnlabels for view.</returns>
        [ComVisible(false)]
        IEnumerable<string> GetViewColumnNames(string viewName, CmcOptionFlags flags = CmcOptionFlags.Fieldname);

        /// <summary>
        /// Saves the Commence View to a file in HTML format (similar to the File-Save As HTML menu command).
        /// If the file does not exist, Commence will create it; the drive and directory must exist.
        /// If the file exists, it will be overwritten.
        /// </summary>
        /// <param name="viewName">Commence view name (case-sensitive)</param>
        /// <param name="linkMode">
        /// <list type="list"><listheader>LINKMODE values</listheader>
        /// <item><term>0</term><description>No view linking (default).</description></item>
        /// <item><term>1</term><description>View linking on selected item; Param1 = category name; Param2 = Item name [clarify issues?]</description></item>
        /// <item><term>2</term><description>View linking on selected date; Param1 = date string (including AI dates); Param2 = unused</description></item>
        /// <item><term>3</term><description>View linking on selected date range; Param1 = start date string (including AI dates); Param2 = end date string (including AI dates)</description></item>
        /// </list>
        /// </param>
        /// <param name="param1">Used with view linking to identify the "active item/date/date range" when the view contains a "view linked" filter.</param>
        /// <param name="param2">Used with view linking to identify the "active item/date/date range" when the view contains a "view linked" filter.</param>
        /// <param name="fileName">(fully qualified) filename</param>
        /// <returns>true on success. Use GetLastError on failure</returns>
        bool GetViewToFile(string viewName, string fileName, CmcLinkMode linkMode = CmcLinkMode.None, string param1 = null, string param2 = null);

        /// <summary>
        /// Marks the active item for subsequent requests.
        /// </summary>
        /// <param name="delim">Delimiter, up to 8 characters.</param>
        /// <returns>String with categoryname and item name, delimited by delim or CR/LF if not supplied, <c>null</c> if no item is active.</returns>
        /// <remarks>Commence documentation does not specify a return value or delimiter for this request, but it does have a return value and can take a delimiter. Inspecting the return value can be useful for checking if the request was successful.</remarks>
		string MarkActiveItem(string delim = null);

        /// <summary>
        /// Creates a new template with the given parameters. If template exists, the call fails.
        /// This call returns the template filename assigned by Commence, including the full path and extension configured in Preferences-Letter dialog.
        /// </summary>
        /// <param name="templateName">Name for new template.</param>
        /// <param name="categoryName">Commence category name.</param>
        /// <param name="shared"><c>true</c> for shared template, ignored when database is stand-alone, fails if user has no author rights.</param>
        /// <returns>Template filename assigned by Commence, including the full path and extension.</returns>
		string MergeTemplateCreate(string templateName, string categoryName, bool shared);

        /// <summary>
        /// Updates an existing template identified by the name parameter. Template must exist, else the call fails.
        /// </summary>
        /// <param name="templateName">Commence template name.</param>
        /// <param name="shared"><c>true</c> for shared, ignored for stand-alone, fails when user is non-author.</param>
        /// <returns>Template filename assigned by Commence, including the full path and extension.</returns>
		string MergeTemplateSave(string templateName, bool shared);

        /// <summary>
        /// Defines the category to be used.
        /// This must be the first message sent to Commence for a ViewData conversation.
        /// ViewCategory may be sent at any time to reset the ViewData conversation state.
        /// ViewData conversation topic has been obsolete since Commence 3.0 (1995 or so),
        /// so it is extremely unlikely you will ever need this method.
        /// </summary>
        /// <param name="categoryName">Commence category name.</param>
        /// <returns>"OK" on success, NACK error code on failure, <see cref="GetLastError"/>.</returns>
		string ViewCategory(string categoryName);

        /// <summary>
        /// Defines the logical operations that link the filter clauses specified with the ViewFilter request.
        /// Each parameter can be either "And" or "Or"; the default value is "And". AndOr12 defines the relationship between the filter clauses 1 and 2, AndOr34 for clauses 3 and 4, and AndOr13 for (1 and 2) and (3 and 4), etc. Note the special significance of AndOr48, which signifies the relationship between the two groups of 4 filters.
        /// <para>
        /// The logical precedence is:
        /// (Filter1 AndOr12 Filter2) AndOr13 (Filter3 AndOr34 Filter4) AndOr4to8 [repeat 5 to 8]
        /// </para>
        /// </summary>
        /// <param name="AndOr12">string "And" or "Or".</param>
        /// <param name="AndOr13">string "And" or "Or".</param>
        /// <param name="AndOr34">string "And" or "Or".</param>
        /// <param name="AndOr48">string "And" or "Or".</param>
        /// <param name="AndOr56">string "And" or "Or".</param>
        /// <param name="AndOr57">string "And" or "Or".</param>
        /// <param name="AndOr78">string "And" or "Or".</param>
        /// <returns>"OK" on success, NACK error code on failure, <see cref="GetLastError"/></returns>
		string ViewConjunction(string AndOr12, string AndOr13, string AndOr34, string AndOr48, string AndOr56, string AndOr57, string AndOr78);

        /// <summary>
        /// Returns the number of items in the connCategory that the Indexth item is connected to via ConnectionName. 
        /// </summary>
        /// <param name="index">Index of Nth item (1-based).</param>
        /// <param name="connectionName">Commence connection name (case-sensitive).</param>
        /// <param name="connCategory">Commence connected category.</param>
        /// <returns>Connected item count, -1 on error.</returns>
        /// <seealso cref="GetConnectedItemCount"/>
		int ViewConnectedCount(int index, string connectionName, string connCategory);

        /// <summary>
        /// Returns the value of Field of the ConnIndexth item in the connCategory that is connected to the Indexth item via ConnectionName. Use the ViewConnectedCount REQUEST to determine the maximum ConnIndex value. The string "(deleted)" will be returned if the requested item is deleted after the ViewData snapshot is taken.
        /// </summary>
        /// <param name="index">Index of item to query.</param>
        /// <param name="connectionName">Commence connection name (case-sensitive).</param>
        /// <param name="connCategory">Commence connected category.</param>
        /// <param name="connIndex">Index of connected item.</param>
        /// <param name="fieldName">Commence fieldname to return from connCategory. Field cannot be of the field type Image.</param>
        /// <returns>Fieldvalue.</returns>
        /// <seealso cref="GetConnectedItemField"/>
		string ViewConnectedField(int index, string connectionName, string connCategory, int connIndex, string fieldName);

        /// <summary>
        /// Returns multiple field values (Field_1 through Field_n) in the ConnIndexth item in the connCategory that is connected to the Indexth item via ConnectionName. Use the ViewConnectedCount REQUEST to determine the maximum ConnIndex value. The string "(deleted)" will be returned if the requested item is deleted after the ViewData snapshot is taken.
        /// Field cannot be of the fieldtype Image.
        /// </summary>
        /// <param name="index">Index of Nth item (1-based).</param>
        /// <param name="connectionName">Commence connection name (case-sensitive).</param>
        /// <param name="connCategory">Commence connCategory name.</param>
        /// <param name="connIndex">Index of connected row.</param>
        /// <param name="fields">Object array containg fieldnames.</param>
        /// <param name="delim">Delimiter, up to max 8 chars.</param>
        /// <returns>String array of fieldvalues for the requested fields.</returns>
        string ViewConnectedFields(int index, string connectionName, string connCategory, int connIndex, object[] fields, string delim = null);

        /// <summary>
        /// Returns multiple field values (Field_1 through Field_n) in the ConnIndexth item in the connCategory that is connected to the Indexth item via ConnectionName. Use the ViewConnectedCount REQUEST to determine the maximum ConnIndex value. The string "(deleted)" will be returned if the requested item is deleted after the ViewData snapshot is taken.
        /// Field cannot be of the fieldtype Image.
        /// </summary>
        /// <param name="index">Index of Nth item (1-based).</param>
        /// <param name="connectionName">Commence connection name (case-sensitive).</param>
        /// <param name="connCategory">Commence connCategory name.</param>
        /// <param name="connIndex">Index of connected row.</param>
        /// <param name="fields">List of fieldnames.</param>
        /// <param name="delim">Delimiter, up to max 8 chars.</param>
        /// <returns>String array of fieldvalues for the requested fields.</returns>
        [ComVisible(false)]
        string[] ViewConnectedFields(int index, string connectionName, string connCategory, int connIndex, List<string> fields, string delim = null);

        /// <summary>
        /// Returns the ConnIndexth item in connCategory that the Indexth item is connected to via ConnectionName. Use the ViewConnectedCount REQUEST to determine the maximum ConnIndex value.
        /// </summary>
        /// <param name="index">Index of Nth item (1-based).</param>
        /// <param name="connectionName">Commence connection name (case-sensitive).</param>
        /// <param name="connCategory">Commence connCategory name.</param>
        /// <param name="connIndex">Index of connected row.</param>
        /// <returns>ConnIndexth item in connCategory that the Indexth item is connected to via ConnectionName.</returns>
        /// <seealso cref="GetConnectedItemNames(string, string, string, string, string)"/>
		string ViewConnectedItem(int index, string connectionName, string connCategory, int connIndex);

        /// <summary>
        /// USE WITH EXTREME CAUTION!
        /// Deletes all items that satisfy the currently defined filter.
        /// If no filter is defined, deletes all items in the previously named category, see <see cref="ViewCategory"/>.
        /// Fails without warning when no delete permissions were granted on the category.
        /// </summary>
		void ViewDeleteAllItems();

        /// <summary>
        /// Returns the value of Field of the Indexth item that satisfies the currently defined filter (if any).
        /// Use the ViewItemCount REQUEST to determine the maximum Index value.
        /// The string "(deleted)" will be returned if the requested item is deleted after the ViewData snapshot is taken.
        /// </summary>
        /// <param name="index">Index of Nth item (1-based).</param>
        /// <param name="fieldName">Commence field name.</param>
        /// <returns>FieldValue.</returns>
        /// <seealso cref="GetField"/>
        /// <remarks>If the Name field is requested, the returned value is NOT clarified. Use ViewItemNameand ClarifyItemNames) to retrieve the clarified item name.</remarks>
        string ViewField(int index, string fieldName);

        /// <summary>
        /// Returns multiple field values (n) for the Indexth item that satisfies the currently defined filter (if any). Use the ViewItemCount request to determine that maximum Index value. Use the GetFieldNames request to determine valid field names.
        /// The string "(deleted)" will be returned if the requested item is deleted after the ViewData snapshot is taken. Image type fields are not allowed in the Field_ parameters.
        /// </summary>
        /// <param name="index">Indexth item (1-based) that satisfies the currently defined filter (if any).</param>
        /// <param name="fields">Object array of fields to retrieve.</param>
        /// <param name="delim">Delimiter, up to max 8 chars.</param>
        /// <returns>String array of string representing the fieldvalues.</returns>
        /// <remarks>Using a <see cref="Vovin.CmcLibNet.Database.ICommenceQueryRowSet"/> is usually faster and easier.</remarks>
        /// <seealso cref="GetFields(string, string, object[], string)"/>
        string ViewFields(int index, object[] fields, string delim = null);
        /// <summary>
        /// Returns multiple field values (n) for the Indexth item that satisfies the currently defined filter (if any). Use the ViewItemCount request to determine that maximum Index value. Use the GetFieldNames request to determine valid field names.
        /// The string "(deleted)" will be returned if the requested item is deleted after the ViewData snapshot is taken. Image type fields are not allowed in the Field_ parameters.
        /// </summary>
        /// <param name="index">Indexth item (1-based) that satisfies the currently defined filter (if any).</param>
        /// <param name="fields">List of fields to retrieve.</param>
        /// <param name="delim">Delimiter, up to max 8 chars.</param>
        /// <returns>String array of string representing the fieldvalues.</returns>
        /// <remarks>Using a <see cref="Vovin.CmcLibNet.Database.ICommenceQueryRowSet"/> is usually faster and easier. This method is only available to .NET applications.</remarks>
        /// <seealso cref="GetFields(string, string, object[], string)"/>
        [ComVisible(false)]
        string[] ViewFields(int index, List<string> fields, string delim = null);
        /// <summary>
        /// Saves specified field to file
        /// </summary>
        /// <param name="index">Index of Nth item (1-based).</param>
        /// <param name="fieldName">Commence fieldName.</param>
        /// <param name="fileName">(fully qualified) filename. Existing files will be overwritten</param>
        /// <returns><c>true</c> on success, inspect <see cref="GetLastError" /> on failure.</returns>
        /// <seealso cref="GetFieldToFile"/>
		bool ViewFieldToFile(int index, string fieldName, string fileName);

        /// <summary>
        /// Defines the criteria for the multiple filter to be applied against the previously named category/view (see ViewCategory/ViewView).
        /// This is by far the most complicated API method to use. Unless you are actually creating views from code,
        /// you are strongly advised to use the Filters collection of the CommenceCursor class instead.
        /// </summary>
        /// <param name="clauseNumber">ClauseNumber defines which filter clause is being defined, where ClauseNumber is between 1 and 8.</param>
        /// <param name="filterType">FilterType sets the type of the filter to apply.
        /// Valid filterTypes are F, CTCF, CTI and CTCTI.</param>
        /// <param name="notFlag">NotFlag determines if a logical Not is applied against the entire clause.</param>
        /// <param name="args">String array of filtertype parameters. See Commence DDE documentation</param>
        /// <returns><c>true</c> on success, inspect <see cref="GetLastError" /> on failure.</returns>
        bool ViewFilter(int clauseNumber, string filterType, bool notFlag, object args);

        /// <summary>
        /// Copies the image field data in Field from the Indexth item that satisfies the currently defined filter (if any) to the specified Filename.
        /// </summary>
        /// <param name="index">Index of Nth item (1-based).</param>
        /// <param name="fieldName">Commence field name.</param>
        /// <param name="fileName">Fully qualified path including the drive letter (e.g. c:\tmp\data.bmp). If the file does not exist, Commence will create it; the drive and directory, however, must already exist. If the file exists, it will be overwritten. It is the client’s responsibility to perform the necessary file cleanup.</param>
        /// <returns><c>true</c> on success, inspect <see cref="GetLastError" /> on failure.</returns>
        /// <seealso cref="GetImageFieldToFile"/>
		bool ViewImageFieldToFile(int index, string fieldName, string fileName);

        /// <summary>
        /// Returns the number of items that satisfy the currently defined filter (if any).
        /// If no filter is defined, this request returns the total number of items in the category
        /// </summary>
        /// <returns>Item count.</returns>
        /// <seealso cref="GetItemCount"/>
		int ViewItemCount();

        /// <summary>
        /// Returns the index of the item with the indicated name field value. The item must also satisfy the currently defined filter (if any). Use a clarified itemname if possible.
        /// Note: If the NameFieldValue is left blank, the marked item will be used. Returns the index of the marked item (if any).
        /// </summary>
        /// <param name="nameFieldValue">Commence Name field value. Note: it must be the fieldvalue, NOT the clarified itemname.</param>
        /// <returns>Index (1-based).</returns>
		int ViewItemIndex(string nameFieldValue);

        /// <summary>
        /// Returns the itemname for the Indexth item.
        /// Different from ViewField(Index, Name) in that a clarified item name may be returned. The ClarifyItemNames request is used to enable or disable the use of clarified item names.
        /// </summary>
        /// <param name="index">Index of Nth item (1-based).</param>
        /// <returns>Item name.</returns>
		string ViewItemName(int index);

        /// <summary>
        /// Marks the Indexth item in the view and makes it the default category and item for the EditItem, AppendText, AssignConnection, UnassignConnection, PromoteItemToShared and ShowItem EXECUTE commands.
        /// </summary>
        /// <param name="index">Index of Nth item (1-based).</param>
        /// <returns>True on success, inspect <see cref="GetLastError" /> on failure</returns>
        /// <seealso cref="GetMarkItem"/>
		bool ViewMarkItem(int index);

        /// <summary>
        /// Uses Commence preference information and reverses the name if appropriate. 
        /// </summary>
        /// <param name="itemName">Commence item name.</param>
        /// <param name="prefFlag"> If the optional PrefFlag parameter is 1, 
        /// then the Customize-Preferences-Other-Name Field setting is used to determine if names should be reversed.
        /// If the PrefFlag parameter is left blank or is set to 0,
        /// then the command uses the Reverse Name checkbox value of
        /// the latest invocation of Tools-Send Letter and reverses the name if appropriate.</param>
        /// <returns>Reverse name, if applicable.</returns>
        /// <remarks>This method is identical to GetReverseName.</remarks>
        /// <seealso cref="GetReverseName"/>
        string ViewReverseName(string itemName, int prefFlag);

        /// <summary>
        /// Saves the virtual DDE view already created with the ViewCategory, ViewFilter, ViewSort, etc., as an actual view in Commence, accessible via the normal user interface. The view will be saved using the specified New View Name.
        /// If the view already exists, it will be overwritten. Use GetViewNames to determine if the view name conflicts with an existing view.
        /// </summary>
        /// <param name="newViewName">new View Name to save.</param>
        /// <param name="shared">Make view Shared. Ignored on stand-alone clients.</param>
        /// <returns><c>true</c> on success, inspect <see cref="GetLastError" /> on failure.</returns>
		bool ViewSaveView(string newViewName, bool shared);

        /// <summary>
        /// Defines the sort criteria for the view. Field1, Field2, Field3, and Field4 define the fields to sort by. Sort1, Sort2, Sort3, and Sort4 define the sort type and should be "Ascending" or "Descending". Unused sort pairs (i.e. FieldN, SortN) may be omitted.
        /// </summary>
        /// <param name="fieldName1">Commence sort field name1</param>
        /// <param name="sortOrder1">Order, should be "Ascending" or "Descending".</param>
        /// <param name="fieldName2">Commence sort field name2</param>
        /// <param name="sortOrder2">Order, should be "Ascending" or "Descending".</param>
        /// <param name="fieldName3">Commence sort field name3</param>
        /// <param name="sortOrder3">Order, should be "Ascending" or "Descending".</param>
        /// <param name="fieldName4">Commence sort field name4</param>
        /// <param name="sortOrder4">Order, should be "Ascending" or "Descending".</param>
        /// <returns><c>true</c> on success, inspect <see cref="GetLastError" /> on failure.</returns>
		bool ViewSort(string fieldName1, string sortOrder1, string fieldName2, string sortOrder2, string fieldName3, string sortOrder3, string fieldName4, string sortOrder4);

        /// <summary>
        /// Uses the category and filter criteria of a Commence View that has already be defined from the Commence menu interface.
        /// The ViewView REQUEST overrides any previously received ViewCategory, ViewConjunction, ViewFilter, and ViewSort REQUESTS.
        /// If the ViewName parameter is left blank, the active view will be used.
        /// </summary>
        /// <param name="viewName">Commence view name.</param>
        /// <returns><c>true</c> on success, inspect <see cref="GetLastError" /> on failure.</returns>
		bool ViewView(string viewName);
        #endregion

        #region Commence DDE Execute commands
        /// <summary>
        /// Adds the indicated Item to the Category. Fields other than the Name field are set to their default values. (Default values for date fields must not specify date ranges, e.g., "every day"). By default, the new item is local (i.e., non-shared)
        /// CAUTION: If the Category has mandatory fields, those fields must be filled in with subsequent EditItem commands. Use the GetFieldDefinition request to determine which fields are mandatory. Unpredictable results may follow if mandatory fields are left unfilled.
        /// Note: If a clarify field is defined for the category (but is not a Sequence Number field), the clarify field value can be specified using the optional third parameter, Clarify Value. Alternatively, the Item parameter can be a clarified item name.
        /// If the clarify field is a Sequence Number field, then the field value is determined automatically by Commence. An error results if the AddItem REQUEST tries to set a Sequence Number field. Use GetFieldDefinition to determine the Sequence Number field value so a clarified item name can be used with subsequent EditItems.
        /// </summary>
        /// <param name="categoryName">Commence category name.</param>
        /// <param name="itemName">Commence item name to add.</param>
        /// <param name="clarifyValue">Commence clarify value.</param>
        /// <returns><c>true</c> on success, inspect <see cref="GetLastError" /> on failure.</returns>
        bool AddItem(string categoryName, string itemName, string clarifyValue);

        /// <summary>
        /// Identical in function to AddItem except that it creates a shared item, provided that the database is connected and the category is shared. Otherwise a local item is created.
        /// CAUTION: If the Category has mandatory fields, those fields must be filled in with subsequent EditItem commands. Use the GetFieldDefinition request to determine which fields are mandatory. Unpredictable results may follow if mandatory fields are left unfilled.
        /// </summary>
        /// <param name="categoryName">Commence category name.</param>
        /// <param name="itemName">Commence item name to add.</param>
        /// <param name="clarifyValue">Commence clarify value.</param>
        /// <returns><c>true</c> on success, inspect <see cref="GetLastError" /> on failure.</returns>
		bool AddSharedItem(string categoryName, string itemName, string clarifyValue);

        /// <summary>
        /// Appends text to an existing text Field. Use this command to overcome the 256 character maximum string limitation of certain macro languages (such as WordBasic). Whenever possible, use a clarified item name.
        /// You are strongly advised to use a <see cref="ICommenceEditRowSet"/> for editing fields instead.
        /// <para>When an item was just created with <see cref="AddItem"/> or <see cref="AddSharedItem"/>, this is the item that will be used for appending.
        /// Otherwise, unless you specifically mark an item (using <see cref="GetMarkItem"/> or <see cref="MarkActiveItem"/>) the first match will be used.
        /// </para>
        /// </summary>
        /// <param name="categoryName">Commence category name.</param>
        /// <param name="itemName">Commence item name to edit.</param>
        /// <param name="fieldName">Commence field to edit.</param> 
        /// <param name="text">Textvalue to append.</param>
        /// <returns><c>true</c> on success, inspect <see cref="GetLastError" /> on failure.</returns>
		bool AppendText(string categoryName, string itemName, string fieldName, string text);

        /// <summary>
        /// Assigns a connection between two items.
        /// </summary>
        /// <param name="categoryName">Commence 'From' category.</param>
        /// <param name="itemName">Commence 'From' item. Use clarified names if possible.</param>
        /// <param name="connectionName">Commence connection name (case-sensitive).</param>
        /// <param name="connCategory">Commence 'To' category.</param>
        /// <param name="connItem">Commence 'To' item. Use clarified names if possible.</param>
        /// <returns><c>true</c> on success, inspect <see cref="GetLastError" /> on failure.</returns>
		bool AssignConnection(string categoryName, string itemName, string connectionName, string connCategory, string connItem);

        /// <summary>
        /// Updates the detail form script with the text in the specified Filename.
        /// </summary>
        /// <param name="categoryName">Commence category name.</param>
        /// <param name="formName">Commence Item Detail Form name.</param>
        /// <param name="fileName">Fully qualified path including the drive letter (e.g. c:\tmp\script.txt).</param>
        /// <returns><c>true</c> on success, inspect <see cref="GetLastError" /> on failure.</returns>
        /// <remarks>Warning: this function will return <c>true</c> even if the user has no Author rights on the database.
        /// This is a limitation of Commence. If you are not an Author, the script will not be checked in!</remarks>
		bool CheckInFormScript(string categoryName, string formName, string fileName);

        /// <summary>
        /// Saves the detail form script to file.
        /// </summary>
        /// <param name="categoryName">Commence category name.</param>
        /// <param name="formName">Commence Item Detail Form name.</param>
        /// <param name="fileName">Fully qualified path including the drive letter (e.g. c:\tmp\script.txt). If file already exists it is overwritten.</param>
        /// <returns><c>true</c> on success, inspect <see cref="GetLastError" /> on failure.</returns>
		bool CheckOutFormScript(string categoryName, string formName, string fileName);

        /// <summary>
        /// Deletes the indicated Item from the Category. This action cannot be undone.  Whenever possible, use a clarified item name.
        /// </summary>
        /// <param name="categoryName">Commence category name.</param>
        /// <param name="itemName">Commence item name to delete.</param>
        /// <returns><c>true</c> on success, inspect <see cref="GetLastError" /> on failure.</returns>
		bool DeleteItem(string categoryName, string itemName);

        /// <summary>
        /// Deletes the specified view. Use with caution! This action cannot be undone.
        /// </summary>
        /// <param name="viewName">Commence view name.</param>
        /// <returns><c>true</c> on success, inspect <see cref="GetLastError" /> on failure.</returns>
		bool DeleteView(string viewName);

        /// <summary>
        /// Sets the value of Field to Value for the item identified by Item in Category.
        /// Field must not specify a Calculation field.
        /// Date fields may not specify a range, (e.g., "every day").
        /// Whenever possible, use a clarified Item name.
        /// </summary>
        /// <param name="categoryName">Commence category name.</param>
        /// <param name="itemName">Commence item name to edit.</param>
        /// <param name="fieldName">Commence field to edit.</param> 
        /// <param name="fieldValue">New fieldvalue. 
        /// If Field is an image field, Value must specify the filename containing the image.
        /// Only .bmp files are supported.</param>
        /// <returns><c>true</c> on success, inspect <see cref="GetLastError" /> on failure.</returns>
		bool EditItem(string categoryName, string itemName, string fieldName, string fieldValue);

        /// <summary>
        /// Fires the indicated Agent Trigger.
        /// </summary>
        /// <param name="trigger">Trigger string.</param>
        /// <param name="args">.NET clients can pass in a string[], from VBScript, VB etc. use Array(param1, param2, ... ParamN). Up to 9 parameters can be used.
        /// See the Commence helpfile under "Receive DDE Trigger" for implementation details.</param>
        /// <returns><c>true</c> on success, inspect <see cref="GetLastError" /> on failure.</returns>
        /// <remarks>This is a very powerful command allowing you to make Commence perform agent actions.
        /// <para>Additional information from the Commence helpfile. Note that you do not have to worry about the syntax, it is taken care of internally in CmcLibNet. See examples below.</para>
        /// <para>Commence can pass data from the Send DDE Action of one agent to the Receive DDE Trigger of another agent, or an agent can be triggered by code. The data can then be used by the Receive DDE agent’s actions. The arguments passed via a Receive DDE trigger are accessible using the following keyword strings: (-0-), (-1-), (-2-), (-3-), (-4-), (-5-), (-6-), (-7-), (-8-) and (-9-). These keywords (i.e. arguments) may be used just as field codes would be used to display the data in a message box or status bar action, to set values in an add or edit item action, or any other place that keyword replacement is allowed in agents. The following example illustrates how the DDE trigger arguments are mapped to keywords.
        /// </para>
        /// <para>
        /// Usage:
        /// 
        /// <code>[FireTrigger(TrigStr, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9)]</code>
        /// </para>
        /// <list type="table">
        /// <item><term>(-0-)</term><description>Always "FireTrigger".</description></item>
        /// <item><term>(-1-)</term><description>Trigger String, e.g., TrigStr</description></item>
        /// <item><term>(-2-) through (-9-)</term><description>The string passed in that position, e.g. Arg2 through Arg9 respectively.</description></item>
        /// </list>
        /// <para>
        /// The Send DDE trigger also allows arguments 2 and 3 to be mapped to a category name and an item name respectively. Then If the "arg2 equals" check box is checked (in the DDE Trigger Definition dialog box of the Receive DDE agent), a category must be selected and Arg2 must equal that category name for that DDE trigger to fire. In addition, Arg3 must be the name of an item in the chosen category. If all these conditions are met, the trigger fires and values from that item will be available for use in the agent’s actions (using the field name keywords such as (%Name%) or using any of the arguments (-3-) through (-9-) that have been defined) 
        /// </para>
        /// <para>
        /// The corresponding agent with a Receive DDE trigger also allows you to define a filter for a DDE trigger using the item mapping option. If a filter is defined, not only do Arg2 and Arg3 have to match the category and an item name, but the item must also meet the filter criteria. If the item does not match the filter, the trigger will not fire.
        /// </para>
        /// <para>Examples:</para>
        /// <para>Execute: <c>[FireTrigger("test", "Person", "Wilson", "908-552-0466")]</c>
        /// </para>
        /// <para>
        /// This is a generic example where all arguments are typed literally.</para>
        /// <para>
        /// Execute: <c>[FireTrigger("test",(-Category-),(%Last Name%),(%Business Phone%))]</c>
        /// </para>
        /// <para>
        /// This is an equivalent example where the agent trigger uses the Person category, so field codes can be used as arguments whenever necessary.
        /// </para>
        /// </remarks>
        /// <example>(C#) <code>string[] args = new string[] { "Company", "012-3456" };
        /// FireTrigger("DoSomething", args);</code></example>
        /// <example>(VBScript) <code language="vbscript">Dim args : args = Array("Company", "012-3456")
        /// FireTrigger("DoSomething", args)</code></example>
        bool FireTrigger(string trigger, object[] args);

        /// <summary>
        /// Launches an add tem detail window for an item in the Phone Log category (as specified in the Commence Customize-Preferences-Event Logs dialog box). One or more Category/Item pairs may be given as arguments; connections will be set between the new Phone Log item and these items. If no connection exists between the two categories, the pair is ignored and no error is given. The connection is set via all connections found between the two categories.
        /// If no Phone Log category is set, NACK error code 104 (DDE_ERROR_NOLOGCLASS) is returned. If there is no room to add an item in the Phone Log category, NACK error code 112 (DDE_ERROR_FULL) is returned.
        /// </summary>
        /// <param name="args">Category1, Item1, ..., Category n, Item n</param>
        /// <returns><c>true</c> on success, inspect <see cref="GetLastError" /> on failure.</returns>
        /// <remarks>Note: if Commence prompts for an Item Detail Form to use, this will halt the call (and subsequent code) until a form was chosen or the form selection window was cancelled.</remarks>
        bool LogPhoneCall(object[] args);

        /// <summary>
        /// Promotes a local item to shared status, provided the database is connected and the category is shared. If the item is already shared, its status is unchanged. Note that there is no way to demote an item from shared to local status.
        /// Whenever possible, use a clarified Item name. 
        /// </summary>
        /// <param name="categoryName">Commence category name.</param>
        /// <param name="itemName">Commence item name.</param>
        /// <returns><c>true</c> on success, inspect <see cref="GetLastError" /> on failure.</returns>
		bool PromoteItemToShared(string categoryName, string itemName);

        /// <summary>
        /// Opens the specified desktop in Commence.
        /// </summary>
        /// <param name="desktopName">Commence desktop name.</param>
        /// <returns><c>true</c> on success, inspect <see cref="GetLastError" /> on failure.</returns>
		bool ShowDesktop(string desktopName);

        /// <summary>
        /// Opens an item detail window in Commence for the given Item in the given Category.
        /// If a detail is already open for the item, it will be brought to the top and given the focus.
        /// </summary>
        /// <param name="categoryName">Commence category name.</param>
        /// <param name="itemName">Commence item name.</param>
        /// <param name="formName">Commence detail form name.</param>
        /// <returns><c>true</c> on success, inspect <see cref="GetLastError" /> on failure.</returns>
		bool ShowItem(string categoryName, string itemName, string formName = null);

        /// <summary>
        /// Opens the specified view in Commence.
        /// If the view is already open, it will become the active view.
        /// </summary>
        /// <param name="viewName">Commence view name.</param>
        /// <param name="newCopy">Set to true to force a new copy of the view to be opened.</param>
        /// <returns><c>true</c> on success, inspect <see cref="GetLastError" /> on failure.</returns>
		bool ShowView(string viewName, bool newCopy);

        /// <summary>
        /// Unassigns a connection between two items.
        /// </summary>
        /// <param name="categoryName">Commence 'From' category.</param>
        /// <param name="itemName">Commence 'From' item. Use clarified names if possible.</param>
        /// <param name="connectionName">Commence connection name (case-sensitive)</param>
        /// <param name="connCategory">Commence 'To' category</param>
        /// <param name="connItem">Commence 'To' item. Use clarified names if possible.</param>
        /// <returns><c>true</c> on success, inspect <see cref="GetLastError" /> on failure.</returns>
		bool UnassignConnection(string categoryName, string itemName, string connectionName, string connCategory, string connItem);
        #endregion

        #region Custom commands

        /// <summary>
        /// Construct a clarified item name from the unclarified itemname, the clarify separator and the clarify value
        /// </summary>
        /// <param name="itemName">Unclarified itemname.</param>
        /// <param name="clarifySeparator">Clarify separator.</param>
        /// <param name="clarifyValue">Clarify value.</param>
        /// <returns>Clarified item name.</returns>
        /// <remarks>This method does not query Commence.</remarks>
        string GetClarifiedItemName(string itemName, string clarifySeparator, string clarifyValue);

        /// <summary>
        /// Checks if field contains duplicate values.
        /// </summary>
        /// <param name="categoryName">Categoryname.</param>
        /// <param name="fieldName">Fieldname.</param>
        /// <param name="caseSensitive">Perform case-sensitive check, default is <c>true</c>.</param>
        /// <returns><c>true</c> if duplicates are found.</returns>
        /// <remarks>Commence allows for filtering on duplicate Name field values only, this method allows for checking other fieldtypes as well.</remarks>
        bool HasDuplicates(string categoryName, string fieldName, bool caseSensitive = true);

        /// <summary>
        /// Checks to see if value is already present for given field.
        /// </summary>
        /// <param name="categoryName">Commence category name.</param>
        /// <param name="fieldName">Commence fieldname.</param>
        /// <param name="fieldValue">Value to check.</param>
        /// <param name="caseSensitive">Compare case-sensitive. Defaults to <c>true</c>.</param>
        /// <returns><c>true</c> if value exists.</returns>
        bool FieldValueExists(string categoryName, string fieldName, string fieldValue, bool caseSensitive = true);

        #endregion

        #region Methods

        /// <summary>
        /// Gets a serializable object containnig the Commence schema definition.
        /// </summary>
        /// <param name="options"><see cref="MetaDataOptions"/></param>
        /// <returns><see cref="DatabaseSchema"/></returns>
        [ComVisible(false)]
        IDatabaseSchema GetDatabaseSchema(MetaDataOptions options = null);

        /// <summary>
        /// Exports the schema information to file. Defaults to Json.
        /// </summary>
        /// <param name="fileName">Fully qualified filename.</param>
        /// <param name="options"><see cref="MetaDataOptions"/></param>
        void ExportDatabaseSchema(string fileName, MetaDataOptions options = null);

        /// <summary>
        /// Close any references to Commence. The object should be disposed after this.
        /// </summary>
        /// <remarks>When used from within a Commence Form Script, failing to call the <c>Close</c> method will leave the commence.exe process running in the background when the user closes Commence. IMPORTANT: this also happens when an unhandled exception (a 'script error') occurs. The Commence process then has to be closed manually from the Windows Task Manager. Be careful to implement proper error handling.
        /// <para>When the assembly is called from a.NET application, there is rarely a need to call this method, unless you want to explicitly release COM references and/or release memory. It can be useful in some cases, because Commence may complain about running out of memory before the Garbage Collector has a chance to kick in.</para>
        /// <para>Technical details: calling this method tells the assembly to release all COM handles (called 'RCW' for 'runtime callable wrapper') to Commence that are open. This is needed because when the object reference to this assembly is set to Nothing (in VB), the .NET assembly may not be notified and will think they are still in use. Garbage Collection will therefore not release them, and the commence.exe process will not be terminated.</para>
        /// </remarks>
        void Close();
        #endregion

        #region Properties
        /// <summary>
        /// Delimiter used in DDE requests. 
        /// Use this only if the default delimiter does not suffice. 
        /// This should be very rare. 
        /// Supply up to 8 characters.
        /// </summary>
        string Delim { get; set; }
        /// <summary>
        /// Secondary delimiter used in DDE requests that act on connections. 
        /// Use this only if the default delimiter does not suffice. 
        /// This should be very rare. 
        /// Supply up to 8 characters.
        /// </summary>
        string Delim2 { get; set; }
        /// <summary>
        /// Path of currently open Commence database.
        /// </summary>
        string Path { get; }
        /// <summary>
        /// Name of currently open Commence database.
        /// </summary>
        string Name { get; }
        #endregion
    }
}

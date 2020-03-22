using System;
using System.Runtime.InteropServices;

namespace Vovin.CmcLibNet.Export
{
    /// <summary>
    /// Interface for ExportSettings.
    /// </summary>
    /// <remarks>Some setttings will negate others or be ignored when not applicable or be manipulated internally.
    /// This is not always fully documented.
    /// <para>You should NOT create code that relies on the settings you provided. 
    /// For example: if you set both <see cref="MaxFieldSize"/> and <see cref="NumRows"/> to a huge number,
    /// the assembly will balance them out so as to prevent out-of-memory exceptions.</para>
    /// </remarks>
    [ComVisible(true)]
    [Guid("2C8B94A9-10AA-49F8-90C7-056492031E17")]
    public interface IExportSettings
    {
        /// <summary>
        /// Return data in a canonical format. Default is <c>false</c>.
        /// </summary>
        /// <remarks>
        /// <list type="table">
        /// <listheader>Canonical format</listheader>
        /// <item><term>Date</term><description>yyyymmdd</description></item>
        /// <item><term>Time</term><description>hh:mm military time, 24 hour clock</description></item>
        /// <item><term>Number</term><description>123456.78 no separator, period for decimal delimiter.  Note that when a field is defined as 'Show as currency' in the Commence UI, numerical values are prepended with a '$' sign.</description></item>
        /// <item><term>Checkbox</term><description>TRUE or FALSE (English)</description></item>
        /// </list>
        /// This option is ignored in when using <see cref="NestConnectedItems"/>, in that case ADO.NET data types are returned.
        /// </remarks>
        bool Canonical { get; set; }
        /// <summary>
        /// Include additional column holding the item's THID and return THIDS instead of Name field values for connected items. (XML and Json). For text-based exports, connected items are not returned as thids. Default is <c>false</c>.
        /// </summary>
        /// <remarks>Ignored for custom cursors; create them with the thid flag if you want thids.
        /// </remarks>
        bool UseThids { get; set; }
        /// <summary>
        /// Headermode, i.e. what to use for columnames (text, html, excel) or nodenames (xml, json). Default is <see cref="Export.HeaderMode.Fieldname"/>.
        /// </summary>
        HeaderMode HeaderMode { get; set;}
        /// <summary>
        /// Use custom columnheaders. They must be unique and match the number of fields in the cursor.
        /// </summary>
        /// <remarks>You cannot use custom headers in combination with <see cref="ExportSettings.NestConnectedItems"/>.
        /// <para>You must supply custom headers for all columns in the cursor, even when you have set <see cref="SkipConnectedItems"/> to <c>true</c>.</para>
        /// <para>Type is <c>Object</c> for compatibility with COM.</para>
        /// </remarks>
        object[] CustomHeaders { get; set;}
        /// <summary>
        /// Data format the export engine should generate. Default is XML.
        /// </summary>
        ExportFormat ExportFormat { get; set;}
        /// <summary>
        /// Ignore connected items.
        /// </summary>
        /// <remarks>Note that they may still be read from the database (e.g. when exporting views), just ignored in the output.
        /// <para>Not all exports respect this setting.</para>
        /// <para>The recommended way to ignore connected items is simply to create a custom cursor or view that does not include them.</para></remarks>
        bool SkipConnectedItems { get; set; }
        /// <summary>
        /// External CSS file to be associated with an HTML export. It will not be in-lined.
        /// Only applies to HTML exports. Default is empty (no CSSFile).
        /// </summary>
        string CSSFile { get; set; }
        /// <summary>
        /// Text-delimiter used in a Text export. Only applies to Text and HTML exports. Default is tab.
        /// </summary>
        string TextDelimiter { get; set; }
        /// <summary>
        /// Text-delimiter used for connected values. Only applies to Text and HTML exports. Default is new-line character (same as Commence uses).
        /// </summary>
        string TextDelimiterConnections { get; set; }
        /// <summary>
        /// Text-qualifier used in a Text export. Only applies to Text and HTML exports. Default is " (double-quote).
        /// </summary>
        string TextQualifier { get; set; } // we really want a char, but COM Interop will complain
        /// <summary>
        /// Use headers on the first row in table-like exports (like Text, HTML, Excel). Default is <c>true</c>.
        /// </summary>
        bool HeadersOnFirstRow { get; set; }
        /// <summary>
        /// Make values ISO8601-compliant. Overrides <see cref="Canonical"/>.
        /// </summary>
        bool ISO8601Compliant { get; set; }
        /// <summary>
        /// Treat connections as distinct nodes/elements. Formatting options may be ignored. Default is <code>false</code>.
        /// </summary>
        /// <remarks>
        /// Only applies to <see cref="ExportFormat.Json"/> and <see cref="ExportFormat.Xml"/>.
        /// </remarks>
        bool NestConnectedItems { get; set; }
        /// <summary>
        /// Make Commence use DDE calls to retrieve data. <b>WARNING:</b> extremely slow!
        /// </summary>
        /// <remarks>Should only be used as a last resort, as this can take a *very* long time.
        /// <para>You would use this in case you run into trouble retrieving all items from a connection. 
        /// This setting will request the connected items one by one. <seealso cref="PreserveAllConnections"/></para>
        /// </remarks>
        bool UseDDE { get; set; }
        /// <summary>
        /// Include all connected items.
        /// Overrides <see cref="Canonical"/>, <see cref="ISO8601Compliant"/>, <see cref="SkipConnectedItems"/>,
        /// <see cref="SplitConnectedItems"/>. 
        /// </summary>
        /// <remarks>Use this if connected data is truncated. Will be significantly slower due to multiple data reads.
        /// <para>Changes the order of columns in the cursor; direct columns will come first, then connected columns.</para>
        /// </remarks>
        bool PreserveAllConnections { get; set; }
        /// <summary>
        /// Get serialized XML with a XmlSchema (XSD).
        /// Only applies with <see cref="PreserveAllConnections"/> in combination with <see cref="ExportFormat.Xml"/>. 
        /// Default is <code>false</code>.
        /// </summary>
        /// <remarks>Uses ADO.NET built-in serialization. Allows for importing into SQL Server and performing your own query logic.</remarks>
        bool WriteSchema { get; set; }
        /// <summary>
        /// Include additional connection information. Only applies to <see cref="ExportFormat.Json"/>. Default is <c>true</c>.
        /// </summary>
        /// <remarks>
        /// When set to <c>true</c>, any connected value is an object containing information about the connection,
        /// when set to <c>false</c>, the values will be just an array..
        /// </remarks>
        // TODO: incorporate this in XML exports?
        bool IncludeConnectionInfo { get; set; }
        /// <summary>
        /// Split connected values into separate nodes/elements. Only applies to <see cref="ExportFormat.Json"/> and <see cref="ExportFormat.Xml"/>. Default is <c>true</c>.
        /// If set to <c>false</c>, connected values are not split and you will get the data as Commence returns them.
        /// </summary>
        /// <remarks>This setting may be overridden internally if there is no meaningful way of splitting the items.
        /// <list type="table">
        /// <listheader>It will be respected on a</listheader>
        /// <item><term><see cref="Vovin.CmcLibNet.Database.CmcCursorType.Category"/></term><description>If the connected field(s) were defined using 
        /// <see cref="Vovin.CmcLibNet.Database.ICommenceCursor.SetRelatedColumn(int, string, string, string, CmcOptionFlags)"/> 
        /// or <see cref="Vovin.CmcLibNet.Database.CursorColumns.AddRelatedColumn(string, string, string)"/>.</description></item>
        /// <item><term><see cref="Vovin.CmcLibNet.Database.CmcCursorType.View"/></term><description>Always</description></item>
        /// </list>
        /// <para>On a regular cursor, i.e. with the connections defined as 'direct columns' (the default behaviour),
        /// connected items are returned by Commence as comma delimited strings without a text-qualifier,
        /// making it impossible to split them meaningfully.</para>
        /// <para>When a cursor is created with the <see cref="UseThids"/> flag, thids will only be returned on 'direct' connection columns in a cursor.
        /// They can therefore not be split when the cursor is of type <see cref="Vovin.CmcLibNet.Database.CmcCursorType.Category"/>
        /// *unless* only the name field of the connection is present in the cursor.
        /// Including logic for that scenario unfortunately blows my mind.
        /// If you want thids *and* split them, create a view, or just split them yourselves.</para>
        /// </remarks>
        bool SplitConnectedItems { get; set; }
        /// <summary>
        /// Specify the number of rows to export at a time. Default is 1024.
        /// </summary>
        /// <remarks>This setting may be overriden internally for performance reasons.</remarks>
        int NumRows { get; set; }
        /// <summary>
        /// Maximum number of characters to retrieve from fields. This includes connected fields.
        /// The default when exporting from the export engine is ~500.000, which is about five times the Commence default.
        /// </summary>
        /// <remarks>Setting this can have severe impact on memory usage. Note that when set to >2^20, <see cref="NumRows"/> may be overridden to prevent your system from exploding.</remarks>
        int MaxFieldSize { get; set; }
        /// <summary>
        /// Delete and recreate Excel file when exporting. Only applies to <see cref="ExportFormat.Excel"/>. Default is <c>true</c>.
        /// </summary>
        [Obsolete]
        bool DeleteExcelFileBeforeExport { get; set; }
        /// <summary>
        /// Update options when exporting to Microsoft Excel
        /// </summary>
        ExcelUpdateOptions XlUpdateOptions { get; set; }
        /// <summary>
        /// Custom root node for Xml, Json, Excel.
        /// </summary>
        /// <remarks>For Text, Html and Event writers this property is ignored
        /// <list type="table">
        /// <listheader><term>Export format</term><description>Notes</description></listheader>
        /// <item><term><see cref="ExportFormat.Excel"/></term><description>Custom sheetname</description></item>
        /// <item><term><see cref="ExportFormat.Json"/></term><description>Custom top-level node name</description></item>
        /// <item><term><see cref="ExportFormat.Xml"/></term><description>Custom root element name</description></item>
        /// </list>
        /// <para>Defaults to the datasource name.
        /// For a view, this is the viewname,
        /// for a category or cursor this is the (primary) category name.</para>
        /// </remarks>
        string CustomRootNode { get; set; }
        /// <summary>
        /// Read Commence data in an asynchronous way. Async reads tend to be slightly faster, but it depends on a number of things,
        /// such as <see cref="NumRows"/> and <see cref="MaxFieldSize"/>.
        /// <para>When you export to <see cref="ExportFormat.Excel"/> or use <see cref="PreserveAllConnections"/> there is a performance gain.</para>
        /// <para>Default is <c>true</c>.</para>
        /// </summary>
        /// <remarks>When <c>true</c>, errors are collected in a <see cref="System.AggregateException"/>. They are harder to debug.
        /// Set this to <c>false</c> if you want the immediate exception.
        /// </remarks>
        bool ReadCommenceDataAsync { get; set; }
    }
}
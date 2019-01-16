using System.Runtime.InteropServices;

namespace Vovin.CmcLibNet.Export
{
    /// <summary>
    /// interface for ExportSettings.
    /// </summary>
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
        /// Headermode, i.e. what to use for columnames (text, html) or nodenames (xml, json). Default is <see cref="Export.HeaderMode.Fieldname"/>.
        /// </summary>
        HeaderMode HeaderMode { get; set;}
        /// <summary>
        /// Use custom columnheaders. Make sure they are all unique and match the number of exported fields.
        /// </summary>
        /// <remarks>You cannot use custom headers in a nested export.</remarks>
        /// <remarks>Supply custom headers for all columns, even when you have set <see cref="SkipConnectedItems"/> to <c>true</c>.</remarks>
        /// <remarks>Datatype is object for compatibility with COM.</remarks>
        object[] CustomHeaders { get; set;}
        /// <summary>
        /// Data format the export engine should generate. Default is XML.
        /// </summary>
        ExportFormat ExportFormat { get; set;}
        /// <summary>
        /// Ignore connected items.
        /// </summary>
        /// <remarks>Note that they will still be read from the database, just ignored in the output.</remarks>
        bool SkipConnectedItems { get; set; }
        /// <summary>
        /// CSS file to be associated with an HTML export.
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
        /// Use headers on the first row in table-like exports (like Text, HTML). Default is <c>true</c>.
        /// </summary>
        bool HeadersOnFirstRow { get; set; }
        /// <summary>
        /// Make values ISO8601-compliant. This allows for use of so-called 'Simple data types' such as "date", "time" and "boolean" in XSD Schemas.
        /// </summary>
        bool XSDCompliant { get; set; }
        /// <summary>
        /// XML Schema Definition file associated with XML export. This property is not intended for use in your code. It is exposed publicly to fulfill COM Interop requirements. Setting this property has no effect.
        /// </summary>
        string XSDFile { get; set; }
        /// <summary>
        /// Adds an id value to every exported item and all it's connection data. Formatting options are ignored.
        /// </summary>
        /// <remarks>
        /// When set to <c>false</c>, values from connections are exported as individual nodes.
        /// If a field has no value, an empty node is returned.
        /// <para>The relationship between fields belonging to a connected item is preserved by the position of the node.
        /// For example, every 2-nd node for field A in connection X and every 2-nd node for field B belong to the same connected item in category X.</para>
        /// <para>That is fine for XPath style querying, but if you want more SQL-style querying, 
        /// setting this property to <c>true</c> will include an (arbitrary) id value for every item and for every connected item.</para>
        /// <para>Internally, this uses a completely different way of exporting, which is why formatting settings are ignored.</para>
        /// </remarks>
        bool NestConnectedItems { get; set; }
        /// <summary>
        /// Make Commence use DDE calls to retrieve data. WARNING: extremely slow! I am not kidding.
        /// </summary>
        /// <remarks>Should only be used as a last resort, as this can take a very long time (up to days!).
        /// <para>You would use this in case you run into trouble retrieving all items from a connection. 
        /// The Commence API only returns a limited number of characters by default; note that the limit is settable <see cref="Vovin.CmcLibNet.Database.CommenceCursor.MaxFieldSize"/>.
        /// This setting will request the connected items one by one.</para></remarks>
        bool PreserveAllConnections { get; set; }
        /// <summary>
        /// Include additional connection information. Applies to XML and JSON only. Default is <c>true</c>.
        /// </summary>
        /// <remarks>
        /// <para>When set to <c>true</c>, connection values are exported as follows:</para>
        /// <code language="xml">
        /// <Item>
        ///    <Name>ItemX in CategoryB</Name>
        ///    <ConnectedItem>
        ///      <ConnectionName>
        ///        <ConnectedCategory>
        ///          <ConnectedField>Item1 in CategoryA</ConnectedField>
        ///          <ConnectedField>Item2 in CategoryA</ConnectedField>
        ///          <ConnectedField>ItemN in CategoryA</ConnectedField>
        ///        </ConnectedCategory>
        ///      </ConnectionName>
        ///    </ConnectedItem>
        ///  </Item>
        ///  </code>
        ///  <para>When set to <c>false</c>, connection values are exported as follows:</para>
        /// <code language="xml">
        /// <Item>
        ///   <Name>ItemX in CategoryB</Name>
        ///   <ConnectedField>Item1 in CategoryA</ConnectedField>
        ///   <ConnectedField>Item2 in CategoryA</ConnectedField>
        ///   <ConnectedField>ItemN in CategoryA</ConnectedField>
        /// </Item>
        /// </code>
        /// <para>For Json, when set to <c>true</c>, any connected value is an object containing information about the connection,
        /// when set to <c>false</c>, the values will be just an array.</para>
        /// </remarks>
        bool IncludeConnectionInfo { get; set; }
        /// <summary>
        /// Split connected values into separate nodes (XML, JSON). Default is <c>true</c>.
        /// If set to <c>false</c>, connected values are not split and you will get the data as Commence returns them.
        /// </summary>
        bool SplitConnectedItems { get; set; }
        /// <summary>
        /// Specify the number of rows to export at a time. Default is 1000.
        /// </summary>
        int NumRows { get; set; }
        /// <summary>
        /// Maximum number of characters to retreive from fields. This includes related fields. The default when exporting from the export engine is 500.000.
        /// <see cref="Database.ICommenceCursor.MaxFieldSize"/>
        /// </summary>
        int MaxFieldSize { get; set; }
        /// <summary>
        /// Delete and recreate Excel file when exporting. Only applies to <seealso cref="ExportFormat.Excel"/>
        /// </summary>
        bool DeleteExcelFileBeforeToExport { get; set; }
    }
}
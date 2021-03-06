﻿using System;
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
        /// </remarks>
        bool Canonical { get; set; }
        /// <summary>
        /// Used in conjunction with <see cref="Canonical"/>, 
        /// allows removal of the currency symbol that Commence will include when numeric fields were set to 'Display as currency'
        /// Default is <c>false</c>.
        /// </summary>
        bool RemoveCurrencySymbol { get; set; }
        /// <summary>
        /// Include additional column holding the item's THID and return THIDs instead of Name field values for connected items. (XML and Json). For text-based exports, connected items are not returned as thids. Default is <c>false</c>.
        /// </summary>
        /// <remarks>Note that custom cursors must be created with the <see cref="Vovin.CmcLibNet.Database.CmcOptionFlags.UseThids"/> flag for this settings to take effect.
        /// </remarks>
        bool UseThids { get; set; }
        /// <summary>
        /// Headermode, i.e. what to use for columnames (<see cref="ExportFormat.Text"/>, <see cref="ExportFormat.Html"/>, <see cref="ExportFormat.Excel"/>) 
        /// or nodenames (<see cref="ExportFormat.Xml"/>, <see cref="ExportFormat.Json"/>). Default is <see cref="Export.HeaderMode.Fieldname"/>.
        /// </summary>
        HeaderMode HeaderMode { get; set;}
        /// <summary>
        /// Use custom columnheaders.
        /// </summary>
        /// <remarks>You cannot use custom headers in combination with <see cref="ExportSettings.NestConnectedItems"/>.
        /// <para>You must supply custom headers for all columns in the cursor, even when you have set <see cref="SkipConnectedItems"/> to <c>true</c>.</para>
        /// <para>Type is <c>object[]</c> for compatibility with COM.</para>
        /// </remarks>
        object[] CustomHeaders { get; set;}
        /// <summary>
        /// Data format the export engine should generate. Default is <see cref="ExportFormat.Xml"/>.
        /// </summary>
        ExportFormat ExportFormat { get; set;}
        /// <summary>
        /// Ignore connected items. 
        /// </summary>
        /// <remarks>Note that connected data will still be read if defined, just ignored in the output.
        /// <para>Not all exports respect this setting.</para>
        /// <para>For best performance create a custom cursor or view that does not include connected columns.</para></remarks>
        bool SkipConnectedItems { get; set; }
        /// <summary>
        /// External CSS file to be linked with an HTML export. Location is the output path.
        /// Only applies to HTML exports. Default is empty (no CSSFile).
        /// </summary>
        string CSSFile { get; set; }
        /// <summary>
        /// Text-delimiter used in a Text export. Only applies to <see cref="ExportFormat.Text"/> and <see cref="ExportFormat.Html"/> exports.
        /// Default is tab character.
        /// </summary>
        string TextDelimiter { get; set; }
        /// <summary>
        /// Text-delimiter used for connected values. Only applies to <see cref="ExportFormat.Text"/> and <see cref="ExportFormat.Html"/> exports. 
        /// Default is newline character (same as Commence uses).
        /// </summary>
        string TextDelimiterConnections { get; set; }
        /// <summary>
        /// Text-qualifier. Only applies to <see cref="ExportFormat.Text"/> and <see cref="ExportFormat.Html"/> exports.
        /// Default is " (double-quote).
        /// </summary>
        string TextQualifier { get; set; } // we really want a char, but COM Interop will complain
        /// <summary>
        /// Use headers on the first row in table-like exports (like <see cref="ExportFormat.Text"/>, <see cref="ExportFormat.Html"/>, <see cref="ExportFormat.Excel"/>). 
        /// Default is <c>true</c>.
        /// </summary>
        bool HeadersOnFirstRow { get; set; }
        /// <summary>
        /// Make dates return in ISO8601 format.
        /// </summary>
        /// <remarks>Dates are returned as "yyyy-mm-dd", times are returned as "hh-mm-ss".
        /// Commence does not store seconds, so "ss" will always be "00".
        /// <para>Setting this option to <c>true</c> will return numeric values in <see cref="Canonical"/> format.</para></remarks>
        bool ISO8601Format { get; set; }
        /// <summary>
        /// Only applies to <see cref="ExportFormat.Json"/>. Splits connection information into distinct elements/nodes. 
        /// </summary>
        /// <remarks>This is not the same as <see cref="ExportSettings.SplitConnectedItems"/>; that has to do with the data.</remarks>
        bool NestConnectedItems { get; set; }
        /// <summary>
        /// Make Commence use DDE calls to retrieve data. <b>WARNING:</b> extremely slow!
        /// </summary>
        /// <remarks>Should only be used as a last resort. Note that Fieldnames cannot contain a comma.
        /// <para>You would use this in case you run into trouble retrieving all items from a connection. 
        /// This setting will request connected items one by one. <seealso cref="PreserveAllConnections"/></para>
        /// </remarks>
        bool UseDDE { get; set; }
        /// <summary>
        /// Include all connected items.
        /// Supported only for <see cref="ExportFormat.Xml"/>, <see cref="ExportFormat.Json"/>, <see cref="ExportFormat.Excel"/>
        /// Overrides <see cref="Canonical"/>, <see cref="ISO8601Format"/>, <see cref="SkipConnectedItems"/>,
        /// <see cref="SplitConnectedItems"/>, <see cref="ExcelUpdateOptions"/>.
        /// </summary>
        /// <remarks>The Commence API returns connected field values as a string, subject to <see cref="ExportSettings.MaxFieldSize"/> . 
        /// This means that if there are a large number of connected items, not all data may be retrieved.
        /// <see cref="PreserveAllConnections"/> will ensure all connected data are read. 
        /// There is a performance penalty involved because it will perform multiple reads.
        /// <para>Alternatively, you can try adjust the <see cref="MaxFieldSize"/> parameter. 
        /// Which setting is faster depends on the data. A high <see cref="MaxFieldSize"/> also significantly degrades Commence performance.
        /// </para>
        /// <para>There is no golden rule to this, but for production environment exports you probably want to set this to <c>true</c> just to be safe.</para>
        /// </remarks>
        bool PreserveAllConnections { get; set; }
        /// <summary>
        /// Get serialized XML with a XmlSchema (XSD).
        /// Only applies to <see cref="PreserveAllConnections"/> in combination with <see cref="ExportFormat.Xml"/>. 
        /// Default is <c>false</c>.
        /// </summary>
        /// <remarks>Uses ADO.NET built-in serialization. 
        /// Allows for importing into SQL Server as well as performing your own query logic on the result.</remarks>
        bool WriteSchema { get; set; }
        /// <summary>
        /// Include additional connection information. Only applies to <see cref="ExportFormat.Json"/>. Default is <c>true</c>.
        /// </summary>
        bool IncludeConnectionInfo { get; set; }
        /// <summary>
        /// Split connected values into separate nodes/elements. Only applies to <see cref="ExportFormat.Json"/> and <see cref="ExportFormat.Xml"/>.
        /// Default is <c>true</c>.
        /// If set to <c>false</c>, connected values are not split and you will get the data as Commence returns them.
        /// </summary>
        /// <remarks>This setting may be overridden internally if there is no meaningful way of splitting the items.
        /// <list type="table">
        /// <listheader>It will be respected on a</listheader>
        /// <item><term><see cref="Database.CmcCursorType.Category"/></term><description>If the connected field(s) were defined using 
        /// <see cref="Database.ICommenceCursor.SetRelatedColumn(int, string, string, string, Database.CmcOptionFlags)"/> 
        /// or <see cref="Database.CursorColumns.AddRelatedColumn(string, string, string)"/>.</description></item>
        /// <item><term><see cref="Database.CmcCursorType.View"/></term><description>Always</description></item>
        /// </list>
        /// <para>On a cursor with the connections defined as 'direct columns',
        /// connected items are returned by Commence as comma-delimited strings without a text-qualifier,
        /// making it impossible to split them meaningfully.</para>
        /// </remarks>
        bool SplitConnectedItems { get; set; }
        /// <summary>
        /// Specify the number of rows to export at a time. Default is 1024.
        /// </summary>
        /// <remarks>This setting may be overridden internally for performance reasons.</remarks>
        int NumRows { get; set; }
        /// <summary>
        /// Maximum number of characters to retrieve from fields. This includes connected fields.
        /// The default is ~500.000, skabout five times the Commence default.
        /// </summary>
        /// <remarks>Severely impacts memory usage. When set to >2^20, <see cref="NumRows"/> may be overridden to prevent your system from exploding.</remarks>
        int MaxFieldSize { get; set; }
        /// <summary>
        /// Delete and recreate Excel file when exporting. Only applies to <see cref="ExportFormat.Excel"/>. Default is <c>true</c>.
        /// </summary>
        [Obsolete]
        bool DeleteExcelFileBeforeExport { get; set; }
        /// <summary>
        /// Update options when exporting to Microsoft Excel. See <see cref="ExcelUpdateOptions"/>.
        /// </summary>
        ExcelUpdateOptions XlUpdateOptions { get; set; }
        /// <summary>
        /// Custom root node for <see cref="ExportFormat.Xml"/>, <see cref="ExportFormat.Json"/>, <see cref="ExportFormat.Excel"/>.
        /// </summary>
        /// <remarks>For <see cref="ExportFormat.Text"/>, <see cref="ExportFormat.Html"/> and <see cref="ExportFormat.Event"/> writers this property is ignored
        /// <list type="table">
        /// <listheader><term>Export format</term><description>Notes</description></listheader>
        /// <item><term><see cref="ExportFormat.Excel"/></term><description>Custom sheetname</description></item>
        /// <item><term><see cref="ExportFormat.Json"/></term><description>Custom top-level node name</description></item>
        /// <item><term><see cref="ExportFormat.Xml"/></term><description>Custom root element name</description></item>
        /// </list>
        /// <para>Defaults to the datasource name.
        /// For a view that is the viewname,
        /// for a category or custom cursor that is the category name.</para>
        /// </remarks>
        string CustomRootNode { get; set; }
        /// <summary>
        /// Read Commence data in an asynchronous way. Async can be slightly faster, but it depends on a number of things,
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

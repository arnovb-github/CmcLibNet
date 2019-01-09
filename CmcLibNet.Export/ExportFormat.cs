using System.Runtime.InteropServices;

namespace Vovin.CmcLibNet.Export
{
    /// <summary>
    /// Enum specifying the export format.
    /// </summary>
    [ComVisible(true)]
    [Guid("57BB00BE-5BD8-4F9C-A9CE-34D22BEC4528")]
    public enum ExportFormat
    {
        /// <summary>
        /// Export to XML format (default). Fastest and most reliable.
        /// </summary>
        Xml = 0,
        /// <summary>
        /// Export to JSON format. Uses JSON.NET.
        /// <remarks>If you intend to use this assembly as a standalone dll AND you want to use JSON export, make sure to put the Newtonsoft.Json.dll version 6.0.x in the same folder. In the future, I will forget to update this comment if that changes.
        /// <para>Exporting to JSON uses a lot more memory than exporting to XML.</para></remarks>
        /// </summary>
        Json = 1,
        /// <summary>
        /// Export to HTML format. Note that the resulting HTML will be decorated with class attributes to allow custom formatting via CSS.
        /// </summary>
        Html = 2,
        /// <summary>
        /// Export to plain text format.
        /// <remarks>Note that Commence can do this much faster by using its built-in export functionality.
        /// <para>However, CmcLibNet offers more flexibility, such as custom text qualifiers and delimiters as well as the ability to export related columns other than just the Name field.</para>
        /// </remarks>
        /// </summary>
        Text = 3,
        /// <summary>
        /// Export to Microsoft Excel. No checks are performed to see if the Excel sheet can actually hold all data. Especially when there are connected items involved, you may quickly hit Excel worksheet limits.
        /// <remarks>Pass <c>null</c> or an empty string as filename to the export method to create and open a new Workbook instead of writing values to a file.</remarks>
        /// <para>In its current implementation, integrity of connected items is preserved. That means that every connection and every connected item gets its own row, with the parent item values repeated. CmcLibNet is the only export tool I know that does not treat connected values as a single string, which would likely not fit an Excel cell anyway.</para>
        /// <para>This makes the Excel sheet very suitable for importing into other systems.</para>
        /// <para>However, including connections makes the Excel sheet useless for meaningful calculations.</para>
        /// <para>If you want to do calculations, and most of you will want to, you have the following option(s):
        /// <list type="table">
        ///     <item>  
        ///         <term>Skip connections</term>  
        ///         <description>Simply tell the export to ignore connected items <see cref=" IExportSettings.SkipConnectedItems"/>. Or simply create and export a view that does not contain any.</description>  
        ///     </item>
        /// </list></para>
        /// <para>While the current implementation of Excel exports makes them very well suited for importing into other systems, it makes them rather limited for calculations, say for invoicing. Truth be told it is probably easier to use the Report Writer for that, but data in the Report Writer are fixed after generating the report.</para>
        /// <para>Future implementations of the export engine will have additional parameters controlling Excel exports to include connected data as simple field values.</para>
        /// <para>You can of course always export to another format first, then import into Excel.</para>
        /// </summary>
        Excel = 4,
        /// <summary>
        /// Export to Google Sheets. Requires Google account. Not implemented yet, may well be too slow anyway.
        /// </summary>
        GoogleSheets = 5,
        /// <summary>
        /// Does not export to file, but instead reads the data and emits a <see cref="IExportEngineEvents.ExportProgressChanged"/> event that you can subscribe to.
        /// <para>The <see cref="ExportProgressAsStringChangedArgs.RowValues"/> property will contain the Commence data in a JSON representation.</para>
        /// <para>The number of items contained in the JSON depends on the number of rows you request.</para>
        /// <para>The filename argument passed to any Export* methods is simply ignored when using this setting.</para>
        /// </summary>
        Event = 6,
    }
}

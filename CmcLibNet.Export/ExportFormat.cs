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
        /// Export to Microsoft Excel (version 2007 or higher, the .xslx extension is mandatory).
        /// <remarks>Excel does not have to be installed.
        /// <para>If you choose an existing Excel file, a sheet will be inserted, <seealso cref="ExportSettings.XlUpdateOptions"/>.</para>
        /// <para>If an existing file is specified, CmcLibNet will try to use it's styles for formatting data.</para>
        /// <para>The recommended way of using Excel exports is to start with a non-existing spreadsheet document,
        /// to which you can add spreadseets (and references like formulas etc. ) later.</para>
        /// <para>Note that the (default) <see cref="ExcelUpdateOptions.ReplaceWorksheet"/> option will recreate the worksheet
        /// containing the Commence data upon every eport. Do not place references like formulas directly in it for they will be overwritten.</para>
        /// </remarks>
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

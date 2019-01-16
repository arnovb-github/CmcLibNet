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
        /// Export to Microsoft Excel (xslx extension is mandatory).
        /// <remarks>No checks are performed to see if the Excel sheet can actually hold all rows.
        /// <para>Excel does not have to be installed on the system. CmcLibNet uses the <c>Microsoft.ACE.OLEDB.12.0</c> driver to create the file.
        /// This driver is installed with Microsoft Office, but is also available as separate component. There are two flavors (32 and 64 bit), CmcLibNet was only tested with the 32-bit version.</para>
        /// <para>Because CmcLibNet does not communicate with Excel at all, the export is much faster.
        /// There is a downside to this: columns cannot be formatted.
        /// Excel does a pretty good job at guessing the data type, but things like columnwidths remain default.</para>
        /// <para>The filename has to have the <c>.xslx</c> extension. CmcLibNet cannot export to an open file.
        /// It can also not write data to a sheet with the same name as the exported Commence category.
        /// If you want an unattended export, set <see cref="IExportSettings.DeleteExcelFileBeforeExport"/> to <c>true</c> and the file will simply be recreated.
        /// </para></remarks>
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

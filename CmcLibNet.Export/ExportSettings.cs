using System.Runtime.InteropServices;

namespace Vovin.CmcLibNet.Export
{
    #region Enumerations
    /// <summary>
    /// Enum specifying the desired export format.
    /// </summary>
    [ComVisible(true)]
    [GuidAttribute("57BB00BE-5BD8-4F9C-A9CE-34D22BEC4528")]
    public enum ExportFormat
    {
        /// <summary>
        /// Export to XML format (default). Fastest and safest.
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
        /// <remarks>Pass <c>null</c> or an empty string to the export method to create and open a new Workbook instead of writing values to a file.</remarks>
        /// <para>In its current implementation, integrity of connected items is preserved. That means that every connected item gets its own row, with the parents item values repeated. CmcLibNet is the only export tool that does not treat connected values as a single string, which would likely not fit an Excel cell anyway.</para>
        /// <para>This makes the Excel sheet very suitable for importing into other systems.</para>
        /// <para>It makes the Excel sheet useless for meaningful calculations.</para>
        /// <para>If you want to do calculations, and most of you will want to, you have two options:
        /// <list type="number">
        ///     <item>  
        ///         <term>No connections</term>  
        ///         <description>Simply tell the export to ignore connected items <see cref=" IExportSettings.SkipConnectedItems"/>. Or simply create and export a view that does not contain any.</description>  
        ///     </item>
        ///     <item>  
        ///         <term>Just single connections</term>  
        ///         <description>Make sure you only export connections with just 1 connected item at most.</description>  
        ///     </item>
        /// </list></para>
        /// <para>That may seem overly complex, but in reality there is rarely the need to export 1-N or N-N connections to an Excel sheet yet still perform calculations on them. Most often calculations will be performed on items pertaining to invoices or timesheets, that almost invariably are connected to a single value like a sender or recipient. There may be many notes attached, but these aren't very useful in an Excel sheet.</para>
        /// <para>If you want to build some -say- invoicing system in Commence and have a bunch of notes attached, using the Report Viewer feature in Commence is probably a better choice.</para>
        /// <para>Future implementations of the export engine may have additional parametets controlling Excel exports.</para>
        /// </summary>
        Excel = 4,
        /// <summary>
        /// Export to Google Sheets. Requires Google account. Not implemented yet, may well be too slow anyway.
        /// </summary>
        GoogleSheets = 5,
        /// <summary>
        /// Does not export to file, but instead reads the data and emits a <see cref="IExportEngineEvents.CommenceRowsRead"/> event that you can subscribe to.
        /// <para>The <see cref="CommenceRowsReadArgs.RowValues"/> property will contain the Commence data in a JSON representation.</para>
        /// <para>The number of items contained in the JSON depends on the number of rows you request.</para>
        /// <para>The filename argument passed to the Export* methods is simply ignored when using this setting.</para>
        /// </summary>
        Event = 6,
    }

    /// <summary>
    /// Enum specifying what headers (or nodenames) to use.
    /// </summary>
    [ComVisible(true)]
    [GuidAttribute("4793B385-DCE9-4A87-B557-20423EC0F1BB")]
    public enum HeaderMode
    {
        /// <summary>
        /// Use fieldnames as headers.
        /// </summary>
        Fieldname = 0,
        /// <summary>
        /// Use columnnames as headers. Only applies when exporting views.
        /// If columns have the same label, a sequence number is added. If no columlabel is defined, the underlying fieldname is used.
        /// </summary>
        Columnlabel = 1, // remember the columnname can be empty, Commence then defaults to the fieldname.
        /// <summary>
        /// Use custom headers. The number of supplied headers must be equal to the number of columns to export and they must be unique, regardless of whether SkipConnectedColumns is used.
        /// </summary>
        CustomLabel = 2 // must be correct number AND unique.
    }
    #endregion

    /// <summary>
    /// Class to hold export settings for export engine.
    /// </summary>
    [ComVisible(true)]
    [GuidAttribute("241F0C10-05F2-4282-A1C8-0550989C885A")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IExportSettings))]
    public class ExportSettings : IExportSettings
    {
        private bool _useThids = false;
        private bool _useCanonical = false;
        private HeaderMode _headerMode = HeaderMode.Fieldname;
        private object[] _customHeaders = null;
        private ExportFormat _exportType = ExportFormat.Xml;
        //private EXPORT_TARGETS _exportTarget = EXPORT_TARGETS.EXPORT_FILE; // not used (yet)
        private string _textDelim = "\t";
        private string _textDelim2 = "\n";
        private bool _headersOnFirstRow = true;
        private bool _skipConnectedItems = false;
        private string _textQualifier = "\"";
        private string _xsdfile = null;
        private bool _xsdcompliant =  false;
        private bool _splitConnectedItems = true;
        private int _maxrows = 1000;
        private int _maxfieldsize = 500000;

        /// <inheritdoc />
        public bool Canonical
        {
            get
            {
                return _useCanonical;
            }
            set
            {
                _useCanonical = value;
            }
        }
        /// <inheritdoc />
        public bool UseThids
        {
            get
            {
                if (this.PreserveAllConnections) { _useThids = false; } // override
                return _useThids;
            }
            set
            {
                _useThids = value;
            }
        }
        /// <inheritdoc />
        public HeaderMode HeaderMode
        {
            get
            {
                return _headerMode;
            }
            set
            {
                _headerMode = value;
            }
        }
        /// <inheritdoc />
        public object[] CustomHeaders
        {
            get
            {
                return _customHeaders;
            }
            set
            {
                _customHeaders = value;
                this.HeaderMode = Export.HeaderMode.CustomLabel;
            }
        }
        /// <inheritdoc />
        public ExportFormat ExportFormat
        {
            get
            {
                return _exportType;
            }
            set
            {
                _exportType = value;
            }
        }
        /// <inheritdoc />
        public string TextDelimiter
        {
            get
            {
                return _textDelim;
            }
            set
            {
                _textDelim = value;
            }
        }
        /// <inheritdoc />
        public string TextDelimiterConnections
        {
            get
            {
                return _textDelim2;
            }
            set
            {
                _textDelim2 = value;
            }
        }
        /// <inheritdoc />
        public bool HeadersOnFirstRow
        {
            get
            {
                return _headersOnFirstRow;
            }
            set
            {
                _headersOnFirstRow = value;
            }
        }
        /// <inheritdoc />
        public string CSSFile { get; set; }
        /// <inheritdoc />
        public bool SkipConnectedItems
        {
            get
            {
                return _skipConnectedItems;
            }
            set
            {
                _skipConnectedItems = value;
            }
        }
        /// <inheritdoc />
        public string TextQualifier
        {
            get
            {
                return _textQualifier;
            }
            set
            {
                _textQualifier = value;
            }
        }
        /// <inheritdoc />
        public bool XSDCompliant
        {
            get { return _xsdcompliant; }
            set
            {
                _xsdcompliant = value;
                if (_xsdcompliant) { this.Canonical = true; }
            }
        }

        /// <inheritdoc />
        public string XSDFile
        {
            get { return _xsdfile; }
            set { _xsdfile = value; }
        }

        /// <inheritdoc />
        public bool NestConnectedItems { get; set; }

        /// <inheritdoc />
        public bool PreserveAllConnections { get; set; }

        /// <inheritdoc />
        public bool IncludeConnectionInfo { get; set; }

        /// <inheritdoc />
        public bool SplitConnectedItems
        {
            get
            {
                return _splitConnectedItems;
            }
            set
            {
                _splitConnectedItems = value;
            }
        }
        /// <inheritdoc />
        public int NumRows
        {
            get
            {
                return _maxrows;
            }
            set
            {
                if (value < 1) { throw new System.ArgumentOutOfRangeException("MaxRows", value, "Argument must be positive non-zero number."); }
                _maxrows = value;
            }
        }
        /// <inheritdoc />
        public int MaxFieldSize 
        { get
            {
                return _maxfieldsize;
            }
            set
            {
                if (value != _maxfieldsize)
                {
                    _maxfieldsize = value;
                }
            }
        }
        ///// <inheritdoc />
        //public string EventContentType 
        //    {
        //     get
        //    { return _eventcontenttype; }
        //        set
        //    {
        //        switch (value.ToLower())
        //        {
        //            case "application/json":
        //            case "application/xml":
        //                _eventcontenttype = value;
        //                break;
        //            default:
        //                throw new System.FormatException("Content-Type invalid or not supported.");
        //        }
        //    }
        //}
    }
}

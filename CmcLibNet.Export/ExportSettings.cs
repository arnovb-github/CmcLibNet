using System;
using System.Runtime.InteropServices;

namespace Vovin.CmcLibNet.Export
{
    /// <summary>
    /// Class to hold export settings for export engine.
    /// </summary>
    [ComVisible(true)]
    [Guid("241F0C10-05F2-4282-A1C8-0550989C885A")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IExportSettings))]
    public class ExportSettings : IExportSettings
    {
        private bool _useThids = false;
        private object[] _customHeaders = null;
        private bool _xsdcompliant =  false;
        private int _maxrows = (int)Math.Pow(2, 10); // 1024
        private int _maxfieldsize = (int)Math.Pow(2, 19); // that is ~500.000

        /// <inheritdoc />
        public bool Canonical { get; set; } = false;
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
        public HeaderMode HeaderMode { get; set; } = HeaderMode.Fieldname;
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
        public ExportFormat ExportFormat { get; set; } = ExportFormat.Xml;
        /// <inheritdoc />
        public string TextDelimiter { get; set; } = "\t";
        /// <inheritdoc />
        public string TextDelimiterConnections { get; set; } = "\n";
        /// <inheritdoc />
        public bool HeadersOnFirstRow { get; set; } = true;
        /// <inheritdoc />
        public string CSSFile { get; set; }
        /// <inheritdoc />
        public bool SkipConnectedItems { get; set; } = false;
        /// <inheritdoc />
        public string TextQualifier { get; set; } = "\"";
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
        public string XSDFile { get; set; } = null;

        /// <inheritdoc />
        public bool NestConnectedItems { get; set; }

        /// <inheritdoc />
        public bool PreserveAllConnections { get; set; }

        /// <inheritdoc />
        public bool IncludeConnectionInfo { get; set; }

        /// <inheritdoc />
        public bool SplitConnectedItems { get; set; } = true;
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
        /// <inheritdoc />
        [Obsolete]
        public bool DeleteExcelFileBeforeExport { get; set; } = true;
        /// <inheritdoc />
        public ExcelUpdateOptions XlUpdateOptions { get; set; } = ExcelUpdateOptions.RefreshWorksheet;
        /// <inheritdoc />
        public bool ReadCommenceDataAsync { get; set; } = true;
    }
}

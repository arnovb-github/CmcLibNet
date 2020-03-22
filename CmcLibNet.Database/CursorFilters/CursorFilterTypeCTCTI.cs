using System.Runtime.InteropServices;

namespace Vovin.CmcLibNet.Database
{
    /// <summary>
    /// Represents a filter of type 'Connection To Connected Item' (CTCTI).
    /// </summary>
    [ComVisible(true)]
    [Guid("1866675F-7D57-4142-9A91-600C753EAEAB")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(ICursorFilterTypeCTCTI))]
    public sealed class CursorFilterTypeCTCTI : BaseCursorFilter, ICursorFilterTypeCTCTI
    {
        private string _itemName = string.Empty;
        private string _clarifyValue = string.Empty;

        /// <summary>
        /// constructor.
        /// </summary>
        /// <param name="clauseNumber">filter clause, must be a number between 1 and 8.</param>
        internal CursorFilterTypeCTCTI(int clauseNumber) : base(clauseNumber) { }

        /// <inheritdoc />
        public string Connection { get; set; }
        /// <inheritdoc />
        public string Category { get; set; }
        /// <inheritdoc />
        public string Connection2 { get; set; }
        /// <inheritdoc />
        public string Category2 { get; set; }
        /// <inheritdoc />
        public string ClarifySeparator { get; set; }
        /// <inheritdoc />
        // Notice the padding; this is required for CTI and CTCTI filters on clarified categories.
        public string Item
        {
            get
            {
                return _itemName;
            }
            set
            {
                if (!string.IsNullOrEmpty(value))
                {
                    _itemName = value.PadRight(50);
                }
            }
        }

        /// <inheritdoc />
        // Notice the padding; this is required by Commence for CTI and CTCTI filters on clarified categories.
        public string ClarifyValue
        {
            get
            {
                return _clarifyValue;
            }
            set
            {
                if (!string.IsNullOrEmpty(value))
                {
                    _clarifyValue = value.PadRight(40);
                }
            }
        }

        /// <inheritdoc />
        public override string FiltertypeIdentifier => "CTCTI";

        /// <summary>
        /// Creates the filter string.
        /// </summary>
        /// <returns>Filter string.</returns>
        public override string ToString() // It would be better to overload ToString() in the base class and pass a delegate to it
        {
            return base.ToString(FilterFormatters.FormatCTCTIFilter);
        }
    }
}
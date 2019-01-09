using System.Text;
using System.Runtime.InteropServices;
using Vovin.CmcLibNet;

namespace Vovin.CmcLibNet.Database
{
    /// <summary>
    /// Represents a filter of type 'Connection To Connected Item' (CTCTI).
    /// </summary>
    [ComVisible(true)]
    [Guid("1866675F-7D57-4142-9A91-600C753EAEAB")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(ICursorFilterTypeCTCTI))]
    public sealed class CursorFilterTypeCTCTI : CursorFilter, ICursorFilterTypeCTCTI
    {
        private const string _filterType = "CTCTI";
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
                _itemName = value.ToString().PadRight(50, ' ');
            }
        }

        /// <inheritdoc />
        // Notice the padding, this is required when filtering on clarified categories
        public string ClarifyValue
        {
            get
            {
                return _clarifyValue;
            }
            set
            {
                _clarifyValue = value.ToString().PadRight(40, ' ');
            }
        }
        /// <summary>
        /// Creates the filter string.
        /// </summary>
        /// <returns>Filter string.</returns>
        public override string ToString()
        {
            StringBuilder sb = new StringBuilder("[ViewFilter(");
            sb.Append(base.ClauseNumber.ToString() + ',');
            sb.Append(_filterType + ',');
            sb.Append((base.Except) ? "NOT" : ",");
            sb.Append(Utils.dq(this.Connection) + ',');
            sb.Append(Utils.dq(this.Category) + ',');
            sb.Append(Utils.dq(this.Connection2) + ',');
            sb.Append(Utils.dq(this.Category2) + ',');
            sb.Append(Utils.dq(this.Item + this.ClarifySeparator + this.ClarifyValue));
            sb.Append(")]");
            return sb.ToString();
        }
    }

}

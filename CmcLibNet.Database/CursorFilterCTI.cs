using System.Text;
using CmcLibNet.Core;
using System.Runtime.InteropServices;

namespace CmcLibNet.Database 
{
    /// <summary>
    /// Represents a filter of type (Item) Connected To Item (CTI)
    /// </summary>
    [ComVisible(true)]
    [GuidAttribute("57BE7683-A6EF-4f41-9AD3-717126AB7563")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(ICursorFilterTypeCTI))]
    public class CursorFilterTypeCTI : CursorFilter, ICursorFilterTypeCTI
    {
        private const string _filterType = "CTI";
        private string _itemName = string.Empty;
        private string _clarifyValue = string.Empty;

        /// <summary>
        /// constructor.
        /// </summary>
        /// <param name="clauseNumber">filter clause, must be a number between 1 and 8.</param>
        public CursorFilterTypeCTI(int clauseNumber) : base(clauseNumber) { }

        /// <inheritdoc />
        public string ConnectionName { get; set; }
        /// <inheritdoc />
        public string ConnectedCategory { get; set; }
        /// <inheritdoc />
        public string ClarifySeparator { get; set; }
        /// <inheritdoc />
        // Notice the padding; this is required for CTI and CTCTI filters on clarified categories.
        public string ItemName
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
            sb.Append(base.Except ? "NOT," : ",");
            sb.Append(Utils.dq(this.ConnectionName) + ',');
            sb.Append(Utils.dq(this.ConnectedCategory) + ',');
            sb.Append(Utils.dq(this.ItemName + this.ClarifySeparator + this.ClarifyValue));
            sb.Append(")]");
            return sb.ToString();
        }
    }
}

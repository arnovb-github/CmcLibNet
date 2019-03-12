using System;
using System.Text;
using Vovin.CmcLibNet;
using System.Runtime.InteropServices;
using Vovin.CmcLibNet.Extensions;

namespace Vovin.CmcLibNet.Database
{
    /// <summary>
    /// Represents a filter of type 'Field' (F).
    /// </summary>
    [ComVisible(true)]
    [Guid("5B58E064-675F-4d91-A34E-ED3824118033")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(ICursorFilterTypeF))]
    public sealed class CursorFilterTypeF : CursorFilter, ICursorFilterTypeF
    {
        private const string _filterType = "F"; 
        private bool? _Shared;
        private string _filterQualifierString = string.Empty;

        /// <summary>
        /// constructor.
        /// </summary>
        /// <param name="clauseNumber">filter clause, must be a number between 1 and 8.</param>
        internal CursorFilterTypeF(int clauseNumber) : base(clauseNumber) { }

        /// <inheritdoc />
        public string FieldName { get; set; }
        /// <inheritdoc />
        public string FieldValue { get; set; }
        /// <inheritdoc />
        public string FilterBetweenStartValue { get; set; }
        /// <inheritdoc />
        public string FilterBetweenEndValue { get; set; }
        /// <inheritdoc />
        public bool MatchCase { get; set; }
        /// <inheritdoc />
        public bool Shared
        {
            get
            {
                return (bool)_Shared;
            }
            set
            {
                _Shared = value;
            }
        }

        /// <inheritdoc />
        public string QualifierString
        {
            get
            { 
                return _filterQualifierString;
            }
            set
            {
                _filterQualifierString = string.Empty;
                foreach (FilterQualifier fq in Enum.GetValues(typeof(FilterQualifier)))
                {
                    if (string.Compare(value, fq.GetEnumDescription(), true) == 0)
                    {
                        _filterQualifierString = fq.GetEnumDescription();
                        break;
                    }
                }
                if (string.IsNullOrEmpty(_filterQualifierString))
                {
                    _filterQualifierString = "<INVALID QUALIFIER: '" + value + "'>";
                }
            }
        }

        /// <inheritdoc />
        public FilterQualifier Qualifier
        {
            set
            {
                this.QualifierString = value.GetEnumDescription();
            }
        }

        /// <summary>
        /// Sets filter syntax according to direct and 'between' filter qualifiers
        /// </summary>
        /// <param name="sb">StringBuilder</param>
        private void SetFilterValue(StringBuilder sb)
        {
            if (String.Compare(this.QualifierString, FilterQualifier.Between.GetEnumDescription(), true) == 0)
            {
                sb.Append(Utils.dq(this.FilterBetweenStartValue) + ',' + Utils.dq(this.FilterBetweenEndValue) + ',');
            }
            else
            {
                sb.Append(Utils.dq(this.FieldValue) + ',');
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
            sb.Append((base.Except) ? "NOT," : ",");
            if (_Shared.HasValue)
            {
                sb.Append("," + Utils.dq((this.Shared) ? "Shared" : "Local") + ",,");
            }
            else
            {
                sb.Append(Utils.dq(this.FieldName) + ',');
                sb.Append(Utils.dq(this.QualifierString) + ',');
                SetFilterValue(sb);
                sb.Append((this.MatchCase) ? "1" : "0");
            }
            sb.Append(")]");
            return sb.ToString();
        }
    }
}

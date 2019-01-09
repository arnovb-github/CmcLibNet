using System;
using System.Text;
using Vovin.CmcLibNet;
using System.Runtime.InteropServices;

namespace Vovin.CmcLibNet.Database
{
    /// <summary>
    /// Represents a filter of type 'Connection To Category Field' (CTCF).
    /// </summary>
    [ComVisible(true)]
    [Guid("AEADD9F3-002F-4f93-AA98-A8CF730AE1F3")]
    [ClassInterface(ClassInterfaceType.None)] // tells compiler to use first interface instead of AutoDispatch (generates IDispatch based on class, late-binding only) or AutoDual (also creates vtable, AutoDual is strongly recommended against)
    [ComDefaultInterface(typeof(ICursorFilterTypeCTCF))]
    // ideally, this should inherit from ICursorFilterF, but that would break COM
    public sealed class CursorFilterTypeCTCF : CursorFilter, ICursorFilterTypeCTCF
    {
        private const string _filterType = "CTCF"; 
        private bool? _Shared; // nullable field??
        private string _filterQualifierString = String.Empty;

        /// <summary>
        /// constructor.
        /// </summary>
        /// <param name="clauseNumber">filter clause, must be a number between 1 and 8.</param>
        internal CursorFilterTypeCTCF(int clauseNumber) : base(clauseNumber) { }

         /// <inheritdoc />
        public string Connection { get; set; }
        /// <inheritdoc />
        public string Category { get; set; }
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
                    if (String.Compare(value, fq.GetEnumDescription(), true) == 0)
                    {
                        _filterQualifierString = fq.GetEnumDescription();
                        break;
                    }
                }
                if (_filterQualifierString == string.Empty) { _filterQualifierString = "<INVALID QUALIFIER: '" + value + "'>"; }
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

        //private string GetQualifier()
        //{
        //    return String.Empty;
        //}
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
            sb.Append(base.Except ? "NOT," : ",");
            sb.Append(Utils.dq(this.Connection) + ',');
            sb.Append(Utils.dq(this.Category) + ',');
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

using System;
using System.Runtime.InteropServices;
using Vovin.CmcLibNet.Extensions;

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
    public sealed class CursorFilterTypeCTCF : BaseCursorFilter, ICursorFilterTypeCTCF
    {
        private string _filterQualifierString = string.Empty;
        private FilterQualifier _filterQualifier;

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
        // see also CursorFilterTypeF
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
                        _filterQualifier = fq;
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
        [ComVisible(false)]
        public FilterQualifier Qualifier
        {
            get { return _filterQualifier; }
            set
            {
                _filterQualifier = value;
                _filterQualifierString = value.GetEnumDescription();
            }
        }

        /// <inheritdoc />
        public override string FiltertypeIdentifier => "CTCF";

        private bool _shared = false;
        /// <inheritdoc />
        public bool Shared
        {
            get
            {
                return _shared;
            }
            set
            {
                _shared = value;
                SharedOptionSet = true;
            }
        }

        /// <summary>
        /// Keep track of whether the 'shared/local' option was set
        /// </summary>
        public bool SharedOptionSet { get; private set; }

        /// <summary>
        /// Creates the filter string.
        /// </summary>
        /// <returns>Filter string.</returns>
        public override string ToString() // It would be better to overload ToString() in the base class and pass a delegate to it
        {
            return base.ToString(FilterFormatters.FormatCTCFFilter);
            // before the base class ToString(Func<>) overload
            //StringBuilder sb = new StringBuilder("[ViewFilter(");
            //sb.Append(base.ClauseNumber.ToString() + ',');
            //sb.Append(this.FiltertypeIdentifier + ',');
            //sb.Append(base.Except ? "NOT," : ",");
            //sb.Append(Utils.dq(this.Connection) + ',');
            //sb.Append(Utils.dq(this.Category) + ',');
            //if (this.SharedOptionSet)
            //{
            //    sb.Append("," + Utils.dq((this.Shared) ? "Shared" : "Local") + ",,");
            //}
            //else
            //{
            //    sb.Append(Utils.dq(this.FieldName) + ',');
            //    sb.Append(Utils.dq(this.QualifierString) + ',');
            //    SetFilterValue(sb);
            //    sb.Append((this.MatchCase) ? "1" : "0");
            //}
            //sb.Append(")]");
            //return sb.ToString();
        }
    }
}

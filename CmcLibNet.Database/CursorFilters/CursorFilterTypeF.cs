using System;
using System.Text;
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
    public sealed class CursorFilterTypeF : BaseCursorFilter, ICursorFilterTypeF
    {
        private string _filterQualifierString = string.Empty;
        private FilterQualifier _filterQualifier;

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
        // TODO get rid of this property entirely
        // we need to make the Qualifier enum COM-visible and force COM users to use that.
        // tough luck for them, or is it? We'd lose the 'regular' syntax they are used to.
        // but we could then lose all the stuff using the description attribute
        // I'm undecided on this
        // at the very least, shouldn't we set the qualifier property here?
        // NO. circular reference
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
            get { return _filterQualifier; }
            set
            {
                _filterQualifier = value;
                this.QualifierString = value.GetEnumDescription();
            }
        }

        /// <inheritdoc />
        public override string FiltertypeIdentifier => "F";

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
        public override string ToString()
        {
            return base.ToString(FilterFormatters.FormatFFilter);
            // before the base class ToString(Func<>) overload
            //StringBuilder sb = new StringBuilder("[ViewFilter(");
            //sb.Append(base.ClauseNumber.ToString() + ',');
            //sb.Append(this.FiltertypeIdentifier + ',');
            //sb.Append((base.Except) ? "NOT," : ",");
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

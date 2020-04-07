using System;
using System.Runtime.InteropServices;
using Vovin.CmcLibNet.Attributes;
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
        public CursorFilterTypeCTCF(int clauseNumber) : base(clauseNumber) { }

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
                    string description = fq.GetAttributePropertyValue<string,
                        StringValueAttribute>(nameof(StringValueAttribute.StringValue));
                    if (string.Compare(value, description, true) == 0)
                    {
                        _filterQualifierString = description;
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
                _filterQualifierString = value.GetAttributePropertyValue<string,
                        StringValueAttribute>(nameof(StringValueAttribute.StringValue));
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
        }
    }
}

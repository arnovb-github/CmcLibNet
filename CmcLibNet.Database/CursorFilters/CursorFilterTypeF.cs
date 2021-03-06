﻿using System;
using System.Runtime.InteropServices;
using Vovin.CmcLibNet.Attributes;
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
        public CursorFilterTypeF(int clauseNumber) : base(clauseNumber) { }

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
        // This property is the one that COM-clients use to set the Qualifier.
        // Technically there is no reason why they couldn't also just use the Qualifier property itself.
        // However, they would then have to supply one of the enum values.
        // That would be impractical because:
        // a) the enum has lots of values, some of which are interchangeable, 
        //    meaning they mean the same in the context of a filter, such as True, 1 and Checked for a checkbox.
        // b) they will already be used to using strings when composing a filter in Item Detail Form scripting.
        public string QualifierString
        {
            
            get
            { 
                return _filterQualifierString;
            }
            set
            {
                _filterQualifierString = string.Empty;
                string description = string.Empty;
                foreach (FilterQualifier fq in Enum.GetValues(typeof(FilterQualifier)))
                {
                    description = fq.GetAttributePropertyValue<string,
                        StringValueAttribute>(nameof(StringValueAttribute.StringValue));
                    if (string.Compare(value, description, true) == 0)
                    {
                        _filterQualifierString = description;
                        _filterQualifier = fq; // also set Qualifier property
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
        }
    }
}

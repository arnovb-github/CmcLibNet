﻿using System.Runtime.InteropServices;

namespace Vovin.CmcLibNet.Database
{
    /// <summary>
    /// Represents a filter of type 'Item Connected To Item' (CTI).
    /// </summary>
    [ComVisible(true)]
    [Guid("57BE7683-A6EF-4f41-9AD3-717126AB7563")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(ICursorFilterTypeCTI))]
    public sealed class CursorFilterTypeCTI : BaseCursorFilter, ICursorFilterTypeCTI
    {
        private string _itemName = string.Empty;
        private string _clarifyValue = string.Empty;

        /// <summary>
        /// constructor.
        /// </summary>
        /// <param name="clauseNumber">filter clause, must be a number between 1 and 8.</param>
        public CursorFilterTypeCTI(int clauseNumber) : base(clauseNumber) { }

        /// <inheritdoc />
        public string Connection { get; set; }
        /// <inheritdoc />
        public string Category { get; set; }
        /// <inheritdoc />
        public string ClarifySeparator { get; set; }
        /// <inheritdoc />
        // Notice the padding; this is required by Commence for CTI and CTCTI filters on clarified categories.
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
        // Notice the padding, this is required when filtering on clarified categories
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
        public override string FiltertypeIdentifier => "CTI";

        /// <summary>
        /// Creates the filter string.
        /// </summary>
        /// <returns>Filter string.</returns>
        public override string ToString() // It would be better to overload ToString() in the base class and pass a delegate to it
        {
            return base.ToString(FilterFormatters.FormatCTIFilter);
        }

    }
}
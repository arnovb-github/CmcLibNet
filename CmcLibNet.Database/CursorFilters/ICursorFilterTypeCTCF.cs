using System.Runtime.InteropServices;

namespace Vovin.CmcLibNet.Database
{
    /// <summary>
    /// Interface for CursorFilterTypeCTCF.
    /// </summary>
    [ComVisible(true)]
    [Guid("41375E56-2540-4619-9C2A-E0E84D586393")]
    public interface ICursorFilterTypeCTCF : IBaseCursorFilter
    {
        // Simply inheriting from ICursorFilter would be cleaner, but break COM Interop.
        // Instead we just replicate ICursorFilter and slap on the 'new' keyword
        // This is obviously ugly and should be looked into in the future
        #region Replicated for COM Interop
        /// <summary>
        /// Specify Not-filter. Equivalent to the 'Except' chekbox in a filtyer dialog
        /// </summary>
        new bool Except { get; set; }
        /// <summary>
        /// Specify if this filter must be treated as an Or filter. Default is false.
        /// </summary>
        new bool OrFilter { get; set; }
        /// <summary>
        /// Returns the ViewFilter string, useful for debugging.
        /// </summary>
        /// <returns>fully formatted ViewFilter DDE command constructed by the ToString</returns>
        new string GetViewFilterString();
        /// <summary>
        /// Filter clause number. Should be a number between 1 and 8.
        /// </summary>
        new int ClauseNumber { get; set; }
        /// <summary>
        /// String representing the filtertype identifier for use in a Commence DDE request.
        /// </summary>
        new string FiltertypeIdentifier { get; }
        #endregion

        /// <summary>
        /// Connection name; case-sensitive!
        /// </summary>
        string Connection { get; set; }
        /// <summary>
        /// Connected category name; case-sensitive!
        /// </summary>
        string Category { get; set; }
        /// <summary>
        /// See <see cref="ICursorFilterTypeF.QualifierString"/>
        /// </summary>
        string QualifierString { get; set; }
        /// <summary>
        /// Convenience property for the qualifier for in .NET applications. This property takes precedence over <see cref=" QualifierString"/>.
        /// </summary>
        [ComVisible(false)]
        FilterQualifier Qualifier { get; set; }
        /// <summary>
        /// Field to evaluate.
        /// </summary>
        string FieldName { get; set; }
        /// <summary>
        /// Fieldvalue to evaluate.
        /// Ignored when Between qualifier is used.
        /// </summary>
        string FieldValue { get; set; }
        /// <summary>
        /// Start value (inclusive) for filters with Between qualifier. Used only with Between qualifier on F and CTCF filters, otherwise ignored.
        /// When this property is set, you must also set FilterBetweenEndValue.
        /// </summary>
        string FilterBetweenStartValue { get; set; }
        /// <summary>
        /// End value (inclusive) for filters with Between qualifier. Used only with Between qualifier on F and CTCF filters, otherwise ignored.
        /// When this property is set, you must also set FilterStartEndValue.
        /// </summary>
        string FilterBetweenEndValue { get; set; }
        /// <summary>
        /// Case-sensitive parameter.
        /// Set to true for case-sensitive filter.
        /// </summary>
        bool MatchCase { get; set; }
        /// <summary>
        /// Keep track of whether the 'shared/local' option was set
        /// </summary>
        bool SharedOptionSet { get; }
        /// <summary>
        /// Filter for Local/Shared items.
        /// </summary>
        /// <remarks>When set, fieldvalues are ignored.</remarks>
        bool Shared { get; set; }
    }
}

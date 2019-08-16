using System.Runtime.InteropServices;

namespace Vovin.CmcLibNet.Database
{
    /// <summary>
    /// Interface for CursorFilterTypeF.
    /// </summary>
    [ComVisible(true)]
    [Guid("1C3049E6-9BE6-4041-83CB-1E321DFA25D2")]
    public interface ICursorFilterTypeF : ICursorFilter
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
        /// String value specifying how to evaluate the filtervalue.
        /// Not all qualifiers can be used on all field-types. See Commence ViewFilter documentation in DDE helpfile.
        /// <para>If you use COM, either use this member with the strings listed below, or supply the <see cref="FilterQualifier"/> enum value to <see cref="Qualifier"/>.
        /// .NET users are strongly encouraged to use <see cref="Qualifier"/>.</para>
        /// <para>List of valid qualifier strings and the fields they can be applied to.</para>
        /// <list type="table">
        /// <listheader><term>Qualifier string</term><description>Applies to fieldtype(s)</description></listheader>
        /// <item><term>Equal To</term><description>Name, Calculation, E-mail, URL, Number, Selection, Sequence, Telephone, Text fields.</description></item>
        /// <item><term>Not Equal To</term><description>Name, Calculation, E-mail, URL, Number, Selection, Sequence, Telephone, Text fields.</description></item>
        /// <item><term>Less Than</term><description>Calculation, Number, Sequence number fields.</description></item>
        /// <item><term>Greater Than</term><description>Calculation, Number, Sequence number fields.</description></item>
        /// <item><term>Between</term><description>Name, Calculation, Date, E-mail, URL, Number, Sequence number, Time fields.</description></item>
        /// <item><term>True</term><description>Checkbox fields.</description></item>
        /// <item><term>False</term><description>Checkbox fields.</description></item>
        /// <item><term>Checked</term><description>Checkbox fields.</description></item>
        /// <item><term>Not Checked</term><description>Checkbox fields.</description></item>
        /// <item><term>Yes</term><description>Checkbox fields.</description></item>
        /// <item><term>No</term><description>Checkbox fields.</description></item>
        /// <item><term>Before</term><description>Date, Time fields.</description></item>
        /// <item><term>On</term><description>Date fields.</description></item>
        /// <item><term>At</term><description>Time fields.</description></item>
        /// <item><term>After</term><description>Date, Time fields.</description></item>
        /// <item><term>Contains</term><description>Name, E-mail, URL, Telephone, Text fields.</description></item>
        /// <item><term>Doesn't Contain</term><description>Name, E-mail, URL, Telephone, Text fields.</description></item>
        /// <item><term>Shared</term><description>N/A when used with fields.</description></item>
        /// <item><term>Local</term><description>N/A when used with fields.</description></item>
        /// <item><term>1</term><description>Checkbox fields.</description></item>
        /// <item><term>0</term><description>Checkbox fields.</description></item>
        /// </list>
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
        /// Fieldvalue to evaluate. Ignored when Between qualifier is used.
        /// </summary>
        string FieldValue { get; set; }
        /// <summary>
        /// Start value (inclusive) for filters with Between qualifier. Used only with Between qualifier.
        /// When this property is set, you must also set FilterBetweenEndValue.
        /// </summary>
        string FilterBetweenStartValue { get; set; }
        /// <summary>
        /// End value (inclusive) for filters with Between qualifier. Used only with Between qualifier.
        /// When this property is set, you must also set FilterStartEndValue.
        /// </summary>
        string FilterBetweenEndValue { get; set; }
        /// <summary>
        /// Case-sensitive parameter. Set to true for case-sensitive filter. Defaults to false.
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

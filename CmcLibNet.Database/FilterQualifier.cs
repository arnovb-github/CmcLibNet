using System.ComponentModel;

namespace Vovin.CmcLibNet.Database
{
    /// <summary>
    /// Valid qualifiers used in a Commence ViewFilter request.
    /// </summary>
    public enum FilterQualifier
    {
        /// <summary>
        /// "Contains" filter. Applies to: Name, E-mail, URL, Telephone, Text fields.
        /// </summary>
        [FilterValues(1)]
        Contains,
        /// <summary>
        /// "Does Not Contain" filter. Applies to: Name, E-mail, URL, Telephone, Text fields.
        /// </summary>
        [Description("Does not Contain")]
        [FilterValues(1)]
        DoesNotContain,
        /// <summary>
        /// "On" filter. Applies to: Date fields.
        /// </summary>
        [FilterValues(1)]
        On,
        /// <summary>
        /// "At" filter. Applies to: Time fields.
        /// </summary>
        [FilterValues(1)]
        At,
        /// <summary>
        /// "Equal To" filter. Applies to: Name, Calculation, E-mail, URL, Number, Selection, Sequence, Telephone, Text fields.
        /// </summary>
        [Description("Equal To")]
        [FilterValues(1)]
        EqualTo,
        /// <summary>
        /// "Not Equal To" filter. Applies to: Name, Calculation, E-mail, URL, Number, Selection, Sequence, Telephone, Text fields.
        /// </summary>
        [Description("Not Equal To")]
        [FilterValues(1)]
        NotEqualTo,
        /// <summary>
        /// "Less Than" filter. Applies to: Calculation, Number, Sequence number fields.
        /// </summary>
        [Description("Less Than")]
        [FilterValues(1)]
        LessThan,
        /// <summary>
        /// "Greater Than" filter. Applies to: Calculation, Number, Sequence number fields.
        /// </summary>
        [Description("Greater Than")]
        [FilterValues(1)]
        GreaterThan,
        /// <summary>
        /// "Between" filter. Applies to: Name, Calculation, Date, E-mail, URL, Number, Sequence number, Time fields.
        /// </summary>
        [FilterValues(2)]
        Between,
        /// <summary>
        /// "True" filter. Applies to: Checkbox fields.
        /// </summary>
        [FilterValues(0)]
        True,
        /// <summary>
        /// "False" filter. Applies to: Checkbox fields.
        /// </summary>
        [FilterValues(0)]
        False,
        /// <summary>
        /// "True" filter. Applies to: Checkbox fields.
        /// </summary>
        [FilterValues(0)]
        Checked,
        /// <summary>
        /// "False" filter. Applies to: Checkbox fields.
        /// </summary>
        [Description("Not Checked")]
        [FilterValues(0)]
        NotChecked,
        /// <summary>
        /// "True" filter. Applies to: Checkbox fields.
        /// </summary>
        [FilterValues(0)]
        Yes,
        /// <summary>
        /// "False" filter. Applies to: Checkbox fields.
        /// </summary>
        [FilterValues(0)]
        No,
        /// <summary>
        /// "Before" filter. Applies to: Date, Time fields.
        /// </summary>
        [FilterValues(1)]
        Before,
        /// <summary>
        /// "After" filter. Applies to: Date, Time fields.
        /// </summary>
        [FilterValues(1)]
        After,
        /// <summary>
        /// "Blank" filter. Applies to Name fields, Date fields, E-mail fields, Telephone fields, URL fields, Time fields.
        /// </summary>
        [FilterValues(0)]
        Blank,
        /// <summary>
        /// "Shared" filter. Applies to: N/A when used with fields.
        /// </summary>
        [FilterValues(0)]
        Shared,
        /// <summary>
        /// "Local" filter. Applies to: N/A when used with fields.
        /// </summary>
        [FilterValues(0)]
        Local,
        /// <summary>
        /// "True" filter. Applies to: Checkbox fields.
        /// </summary>
        [Description("1")]
        [FilterValues(0)]
        One,
        /// <summary>
        /// "False" filter. Applies to: Checkbox fields.
        /// </summary>
        [Description("0")]
        [FilterValues(0)]
        Zero
    }
}

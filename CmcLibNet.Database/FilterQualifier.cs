using System.ComponentModel;

namespace Vovin.CmcLibNet.Database
{
    /// <summary>
    /// Valid qualifiers used in a Commence ViewFilter request.
    /// We use this enum to check against the values passed in.
    /// This enum is not exposed to COM because enums cannot be strings,
    /// we just trick ourselves here.
    /// </summary>
    public enum FilterQualifier
    {
        /// <summary>
        /// "Contains" filter. Applies to: Name, E-mail, URL, Telephone, Text fields.
        /// </summary>
        Contains,
        /// <summary>
        /// "Does Not Contain" filter. Applies to: Name, E-mail, URL, Telephone, Text fields.
        /// </summary>
        [Description("Does not Contain")]
        DoesNotContain,
        /// <summary>
        /// "On" filter. Applies to: Date fields.
        /// </summary>
        On,
        /// <summary>
        /// "At" filter. Applies to: Time fields.
        /// </summary>
        At,
        /// <summary>
        /// "Equal To" filter. Applies to: Name, Calculation, E-mail, URL, Number, Selection, Sequence, Telephone, Text fields.
        /// </summary>
        [Description("Equal To")]
        EqualTo,
        /// <summary>
        /// "Not Equal To" filter. Applies to: Name, Calculation, E-mail, URL, Number, Selection, Sequence, Telephone, Text fields.
        /// </summary>
        [Description("Not Equal To")]
        NotEqualTo,
        /// <summary>
        /// "Less Than" filter. Applies to: Calculation, Number, Sequence number fields.
        /// </summary>
        [Description("Less Than")]
        LessThan,
        /// <summary>
        /// "Greater Than" filter. Applies to: Calculation, Number, Sequence number fields.
        /// </summary>
        [Description("Greater Than")]
        GreaterThan,
        /// <summary>
        /// "Between" filter. Applies to: Name, Calculation, Date, E-mail, URL, Number, Sequence number, Time fields.
        /// </summary>
        Between,
        /// <summary>
        /// "True" filter. Applies to: Checkbox fields.
        /// </summary>
        True,
        /// <summary>
        /// "False" filter. Applies to: Checkbox fields.
        /// </summary>
        False,
        /// <summary>
        /// "True" filter. Applies to: Checkbox fields.
        /// </summary>
        Checked,
        /// <summary>
        /// "False" filter. Applies to: Checkbox fields.
        /// </summary>
        [Description("Not Checked")]
        NotChecked,
        /// <summary>
        /// "True" filter. Applies to: Checkbox fields.
        /// </summary>
        Yes,
        /// <summary>
        /// "False" filter. Applies to: Checkbox fields.
        /// </summary>
        No,
        /// <summary>
        /// "Before" filter. Applies to: Date, Time fields.
        /// </summary>
        Before,
        /// <summary>
        /// "After" filter. Applies to: Date, Time fields.
        /// </summary>
        After,
        /// <summary>
        /// "Blank" filter. Applies to Name fields, Date fields, E-mail fields, Telephone fields, URL fields, Time fields.
        /// </summary>
        Blank,
        /// <summary>
        /// "Shared" filter. Applies to: N/A when used with fields.
        /// </summary>
        Shared,
        /// <summary>
        /// "Local" filter. Applies to: N/A when used with fields.
        /// </summary>
        Local,
        /// <summary>
        /// "True" filter. Applies to: Checkbox fields.
        /// </summary>
        [Description("1")]
        One,
        /// <summary>
        /// "False" filter. Applies to: Checkbox fields.
        /// </summary>
        [Description("0")]
        Zero
    }
}

using Vovin.CmcLibNet.Attributes;

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
        [StringValue("Contains")]
        [FilterValues(1)]
        Contains,
        /// <summary>
        /// "Does Not Contain" filter. Applies to: Name, E-mail, URL, Telephone, Text fields.
        /// </summary>
        [StringValue("Does not Contain")]
        [FilterValues(1)]
        DoesNotContain,
        /// <summary>
        /// "On" filter. Applies to: Date fields.
        /// </summary>
        [StringValue("On")]
        [FilterValues(1)]
        On,
        /// <summary>
        /// "At" filter. Applies to: Time fields.
        /// </summary>
        [StringValue("At")]
        [FilterValues(1)]
        At,
        /// <summary>
        /// "Equal To" filter. Applies to: Name, Calculation, E-mail, URL, Number, Selection, Sequence, Telephone, Text fields.
        /// </summary>
        [StringValue("Equal To")]
        [FilterValues(1)]
        EqualTo,
        /// <summary>
        /// "Not Equal To" filter. Applies to: Name, Calculation, E-mail, URL, Number, Selection, Sequence, Telephone, Text fields.
        /// </summary>
        [StringValue("Not Equal To")]
        [FilterValues(1)]
        NotEqualTo,
        /// <summary>
        /// "Less Than" filter. Applies to: Calculation, Number, Sequence number fields.
        /// </summary>
        [StringValue("Less Than")]
        [FilterValues(1)]
        LessThan,
        /// <summary>
        /// "Greater Than" filter. Applies to: Calculation, Number, Sequence number fields.
        /// </summary>
        [StringValue("Greater Than")]
        [FilterValues(1)]
        GreaterThan,
        /// <summary>
        /// "Between" filter. Applies to: Name, Calculation, Date, E-mail, URL, Number, Sequence number, Time fields.
        /// </summary>
        [StringValue("Between")]
        [FilterValues(2)]
        Between,
        /// <summary>
        /// "True" filter. Applies to: Checkbox fields.
        /// </summary>
        [StringValue("TRUE")]
        [FilterValues(0)]
        True,
        /// <summary>
        /// "False" filter. Applies to: Checkbox fields.
        /// </summary>
        [StringValue("FALSE")]
        [FilterValues(0)]
        False,
        /// <summary>
        /// "True" filter. Applies to: Checkbox fields.
        /// </summary>
        [StringValue("Checked")]
        [FilterValues(0)]
        Checked,
        /// <summary>
        /// "False" filter. Applies to: Checkbox fields.
        /// </summary>
        [StringValue("Not Checked")]
        [FilterValues(0)]
        NotChecked,
        /// <summary>
        /// "True" filter. Applies to: Checkbox fields.
        /// </summary>
        [StringValue("Checked")]
        [FilterValues(0)]
        Yes,
        /// <summary>
        /// "False" filter. Applies to: Checkbox fields.
        /// </summary>
        [StringValue("Not Checked")]
        [FilterValues(0)]
        No,
        /// <summary>
        /// "Before" filter. Applies to: Date, Time fields.
        /// </summary>
        [StringValue("Before")]
        [FilterValues(1)]
        Before,
        /// <summary>
        /// "After" filter. Applies to: Date, Time fields.
        /// </summary>
        [StringValue("After")]
        [FilterValues(1)]
        After,
        /// <summary>
        /// "Blank" filter. Applies to Name fields, Date fields, E-mail fields, Telephone fields, URL fields, Time fields.
        /// </summary>
        [StringValue("Blank")] // Note: not "Is Blank" as stated in old documentation
        [FilterValues(0)]
        Blank,
        /// <summary>
        /// "Shared" filter. Applies to: N/A when used with fields.
        /// </summary>
        [StringValue("Shared")]
        [FilterValues(0)]
        Shared,
        /// <summary>
        /// "Local" filter. Applies to: N/A when used with fields.
        /// </summary>
        [StringValue("Local")]
        [FilterValues(0)]
        Local,
        /// <summary>
        /// "True" filter. Applies to: Checkbox fields.
        /// </summary>
        [StringValue("1")]
        [FilterValues(0)]
        One,
        /// <summary>
        /// "False" filter. Applies to: Checkbox fields.
        /// </summary>
        [StringValue("0")]
        [FilterValues(0)]
        Zero
    }
}

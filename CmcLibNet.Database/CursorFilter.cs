using System;
using System.ComponentModel;
using System.Reflection;
using System.Runtime.InteropServices;

namespace Vovin.CmcLibNet.Database
{
    #region Enumerations
    /// <summary>
    /// Filter Types that a Commence ViewFilter request can use, i.e. the filters you can apply to Commence.
    /// </summary>
    [ComVisible(true)]
    [Guid("55B52ADD-2B54-489D-AF40-800F750E2EE1")]
    public enum FilterType 
    {
        /// <summary>
        /// Filter on fieldvalue.
        /// </summary>
        [Description("Field")]
        Field = 0,
        /// <summary>
        /// Filter on connection to specific item.
        /// </summary>
        [Description("Connection To Item")]
        ConnectionToItem = 1,
        /// <summary>
        /// Filter on connected item to connected item.
        /// </summary>
        [Description("Connection To Category To Item")]
        ConnectionToCategoryToItem =2,
        /// <summary>
        /// Filter on fieldvalue in connected category.
        /// </summary>
        [Description("Connection To Category Field")]
        ConnectionToCategoryField = 3
    }
    /// <summary>
    /// Valid qualifiers used in a Commence ViewFilter request.
    /// We use this enum to check against the values passed in.
    /// This enum is not exposed to COM because enums cannot be strings,
    /// we just trick ourselves here.
    /// </summary>
    public enum FilterQualifier
    {
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
        /// "On" filter. Applies to: Date fields.
        /// </summary>
        On,
        /// <summary>
        /// "At" filter. Applies to: Time fields.
        /// </summary>
        At,
        /// <summary>
        /// "After" filter. Applies to: Date, Time fields.
        /// </summary>
        After,
        /// <summary>
        /// "Blank" filter. Applies to Name fields, Date fields, E-mail fields, Telephone fields, URL fields, Time fields.
        /// </summary>
        Blank,
        /// <summary>
        /// "Contains" filter. Applies to: Name, E-mail, URL, Telephone, Text fields.
        /// </summary>
        Contains,
        /// <summary>
        /// "Does Not Contain" filter. Applies to: Name, E-mail, URL, Telephone, Text fields.
        /// </summary>
        [Description("Doesn't Contain")]
        DoesNotContain,
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
    #endregion

    /// <summary>
    /// CursorFilter is a an abstract base class for the various Commence filters you can create.
    /// Use CursorFilters.Create to get a derived filter object that exposes properties only applicable to that filtertype.
    /// This shields you from having to use the horrible DDE syntax the ViewFilter request takes.
    /// You can still set filters the old-fashioned way via CommenceCursor.SetFilter.
    /// </summary>
    [ComVisible(true)]
    [Guid("691B0432-73F1-4508-B44E-A8AACB26F1FB")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(ICursorFilter))]
    /* a problem we will now run into is,
     * that interfaces of derived classes are not fully exposed to COM
     * if early-binding is used, we're fine,
     * but when late binding is used,
     * the caller does not have access to the derived interface;
     * it will only see the ICursorFilter interface.
     * This is why we use the 'new' keyword in the derived interfaces.
     * */
    public abstract class CursorFilter : ICursorFilter
    {
        /// <summary>
        /// The clause number is a 1-based int specifying the filter order. There can be up to 8 filters.
        /// </summary>
        private readonly int _clauseNumber;

        /// <summary>
        /// Constructor checks and sets the clausenumber.
        /// </summary>
        /// <param name="clauseNumber">Clause number, should be between 1 and 8.</param>
        protected internal CursorFilter(int clauseNumber)
        {
            if (clauseNumber >= 1 && clauseNumber <= 8)
            {
                _clauseNumber = clauseNumber;
            }
            else
            {
                throw new ArgumentOutOfRangeException("Clause number must be between 1 and 8");
            }
        }

        #region Properties
        /// <inheritdoc />
        public virtual bool OrFilter { get; set; }
        /// <inheritdoc />
        public virtual bool Except { get; set; }

        #endregion

        #region Methods

        /// <summary>
        /// (Read-only) Returns ClauseNumber for use in derived classes.
        /// Should be a value between 1 and 8.
        /// </summary>
        protected internal int ClauseNumber
        {
            get { return _clauseNumber; }
        }

        /// <inheritdoc />
        public virtual string GetViewFilterString()
        {
            return this.ToString();
        }
        #endregion
    }
}

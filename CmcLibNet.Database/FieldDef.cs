using System.ComponentModel;
using System.Runtime.InteropServices;

namespace Vovin.CmcLibNet.Database
{
    #region Enumerations
    /// <summary>
    /// Commence field types.
    /// </summary>
    [ComVisible(true)]
    [GuidAttribute("FC59E8E7-B996-45B5-ABDD-5B582E7B2B9D")]
    public enum CommenceFieldType
    {
        /// <summary>
        /// Text field.
        /// </summary>
        [Description("Text")]
        Text = 0,
        /// <summary>
        /// Number field.
        /// </summary>
        [Description("Number")]
        Number = 1,
        /// <summary>
        /// Date field.
        /// </summary>
        [Description("Date")]
        Date = 2,
        /// <summary>
        /// Telephone field
        /// </summary>
        [Description("Telephone")]
        Telephone = 3,
        /// <summary>
        /// Check Box field
        /// </summary>
        [Description("Check Box")]
        Checkbox = 7,
        /// <summary>
        /// Name field (= primary key) .
        /// </summary>
        [Description("Name")]
        Name= 11,
        /// <summary>
        /// Data File field (= filepath).
        /// </summary>
        [Description("Data File")]
        Datafile = 12,
        /// <summary>
        /// Image field.
        /// </summary>
        [Description("Image")]
        Image = 13,
        /// <summary>
        /// Time field.
        /// </summary>
        [Description("Time")]
        Time = 14,
        /// <summary>
        /// Excel cell. (OBSOLETE)
        /// </summary>
        [Description("Excel Cell")]
        ExcelCell = 15,
        /// <summary>
        /// Calculation field.
        /// </summary>
        [Description("Calculation")]
        Calculation = 20,
        /// <summary>
        /// Sequence number field.
        /// </summary>
        [Description("Sequence")]
        Sequence = 21,
        /// <summary>
        /// Selection field.
        /// </summary>
        [Description("Selection")]
        Selection = 22,
        /// <summary>
        /// E-mail address field.
        /// </summary>
        [Description("E-Mail Address")]
        Email = 23,
        /// <summary>
        /// Internet address field.
        /// </summary>
        [Description("Internet Address")]
        URL = 24
    }
    #endregion
    /// <summary>
    /// Interface for the field definition.
    /// </summary>
    [ComVisible(true)]
    [GuidAttribute("D4434784-03C4-4eea-8660-BB03182B4D1C")]
    public interface IFieldDef
    {
        /// <summary>
        /// Field type.
        /// </summary>
        CommenceFieldType Type { get; }
        /// <summary>
        /// Field type description.
        /// </summary>
        string TypeDescription { get; }
        /// <summary>
        /// Indicates if field is shared.
        /// </summary>
        bool Shared { get; }
        /// <summary>
        /// Indicates if field is mandatory.
        /// </summary>
        bool Mandatory { get; }
        /// <summary>
        /// Indicates if field is a recurring date field.
        /// </summary>
        bool Recurring { get; }
        /// <summary>
        /// Indicates if field is a combobox.
        /// </summary>
        bool Combobox { get; }
        /// <summary>
        /// Maximum number of characters field can hold.
        /// </summary>
        int MaxChars { get; }
        /// <summary>
        /// Default fielvalue (if any).
        /// </summary>
        string DefaultString { get; }
    }

    /// <summary>
    /// Holds information on the field definition.
    /// </summary>
    [ComVisible(true)]
    [GuidAttribute("1F0AEAE1-84CD-44e7-8504-08EE94493439")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IFieldDef))]
    public class FieldDef : IFieldDef
    {
        internal FieldDef() { }
        /// <inheritdoc />
        public bool Shared { get; internal set; }
        /// <inheritdoc />
        public bool Mandatory { get; internal set; }
        /// <inheritdoc />
        public bool Recurring { get; internal set; }
        /// <inheritdoc />
        public bool Combobox { get; internal set; }
        /// <inheritdoc />
        public int MaxChars { get; internal set; }
        /// <inheritdoc />
        public string DefaultString { get; internal set; }
        /// <inheritdoc />
        public CommenceFieldType Type { get; internal set; }
        /// <inheritdoc />
        public string TypeDescription { get; internal set; }
    }
}

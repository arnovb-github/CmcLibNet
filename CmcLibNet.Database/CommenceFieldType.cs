using System.ComponentModel;
using System.Runtime.InteropServices;

namespace Vovin.CmcLibNet.Database
{
    /// <summary>
    /// Commence field types.
    /// </summary>
    [ComVisible(true)]
    [Guid("FC59E8E7-B996-45B5-ABDD-5B582E7B2B9D")]
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
}
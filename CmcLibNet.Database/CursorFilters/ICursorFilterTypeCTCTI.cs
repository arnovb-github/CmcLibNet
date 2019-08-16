using System.Runtime.InteropServices;

namespace Vovin.CmcLibNet.Database
{
    /// <summary>
    /// interface for CursorFilterTypeCTCTI.
    /// </summary>
    [ComVisible(true)]
    [Guid("884531AA-EC45-4a0a-939C-AD065EA9EF7E")]
    public interface ICursorFilterTypeCTCTI : ICursorFilter
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
        string Connection { get;  set; }
        /// <summary>
        /// Connected category name; case-sensitive!
        /// </summary>
        string Category { get;  set; }
        /// <summary>
        /// Connection 2 name; case-sensitive!
        /// </summary>
        string Connection2 { get;  set; }
        /// <summary>
        /// Connected category 2 name; case-sensitive!
        /// </summary>
        string Category2 { get;  set; }
        /// <summary>
        /// Connected itemname. If the filter is on a Commence category that is clarified, it is recommended you provide both the clarify separator and the clarify value.
        /// However, it is possible to use the fully qualified itemname here, including all the right the padding(!).
        /// </summary>
        string Item { get;  set; }
        /// <summary>
        /// Clarify separator. If the filter is on a Commence category that is clarified,
        /// you must provide both the clarify separator and the clarify value
        /// </summary>
        string ClarifySeparator { get;  set; }
        /// <summary>
        /// Clarify value. If the filter is on a Commence category that is clarified,
        /// you must provide both the clarify separator and the clarify value
        /// </summary>
        string ClarifyValue { get;  set; }
    }

}

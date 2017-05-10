using System.Runtime.InteropServices;

namespace Vovin.CmcLibNet.Database
{
    /// <summary>
    /// Exposes members of CursorFilter, a base class for all filters.
    /// </summary>
    [ComVisible(true)]
    [GuidAttribute("E655C043-066E-45eb-BEBF-ECF77CF190FF")]
    public interface ICursorFilter
    {
        /// <summary>
        /// Except flag. Set to true to create a Except (NOT) filter.
        /// </summary>
        bool Except { get; set; }
        /// <summary>
        /// Defines AND/OR filter logic. Set to true to make filter an OR filter. Defaults to 'AND'.
        /// <para>The relationship applies to the next filter you set (just like in the Commence UI) so do not set to <c>true</c> unless you define an additional filter, otherwise your will get all items.</para>
        /// <para>Keep in mind the special grouping Commence uses for filters: [[1 and/or2] and/or [3 and/or 4]] and/or [[5 and/or 6] and/or [7 and/or 8]].</para>
        /// </summary>
        bool OrFilter { get; set; }
        /// <summary>
        /// Convenience method mainly for COM Interop; returns ViewFilter string as you would use it in DDE syntax.
        /// </summary>
        /// <remarks>If you use CmcLibNet through COM Interop, for instance in Commence Form Script, you may have no way of inspecting the filter in an IDE. This method returns the string representation of the filter.</remarks>
        /// <returns>ViewFilter string.</returns>
        string GetViewFilterString();
    }
}

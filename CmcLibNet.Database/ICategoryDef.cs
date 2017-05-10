using System.Runtime.InteropServices;

namespace Vovin.CmcLibNet.Database
{
    /// <summary>
    /// Exposes the category definition.
    /// </summary>
    [ComVisible(true)]
    [GuidAttribute("4237710B-435C-477e-908F-9CE970E4A0A0")]
    public interface ICategoryDef
    {
        /// <summary>
        /// Number of items a category can hold.
        /// <remarks>Up until at least Commence RM 6.1, Commence has a hardcoded limit of 500.000 items per category, no matter what this property will report.</remarks>
        /// </summary>
        int MaxItems { get; }
        /// <summary>
        /// Indicates whether the category is shared.
        /// </summary>
        bool Shared { get; }
        /// <summary>
        /// Indicates if the category allows duplicate items.
        /// </summary>
        bool Duplicates { get; }
        /// <summary>
        /// Indicates if the category is defines to use a clarifier.
        /// </summary>
        bool Clarified { get; }
        /// <summary>
        /// Returns the clarify separator (if any).
        /// </summary>
        string ClarifySeparator { get; }
        /// <summary>
        /// Returns the clarify field (if any).
        /// </summary>
        string ClarifyField { get; }
        /// <summary>
        /// Returns the category ID, -1 on error.  The ID is a sequence number Commence uses to identify categories. No ID can be obtained if the category contains no items.
        /// <remarks>
        /// The category ID returned by this call is always a snapshot value and it is only valid for the current database even when it is part of a workgroup.
        /// <para>Example: A 'Person' category in one client may have a different id on another client, even when they are in the same workgroup.</para>
        /// <para>Category IDs are reused in Commence. Examples:</para>
        /// <para>
        /// If you delete category X and recreate it with the same name, it may get another id than it had before.
        /// If you delete category X and create another category Y, it may get the id category X had.
        /// </para>
        /// <para>
        /// Do not build persistent key-value pairs of categories and ids but always use whatever Commence provides.
        /// As part of the synchronization process (in a workgroup) it will provide several files which will contain the current lookup tables.
        /// </para>
        /// <para>There is an edge situation in which the ID will not be obtainable: if a category is intentionally empty and 'protected' by a Commence Agent that automatically deletes new items as soon as they are created, this property will always return -1.</para>
        /// <para>An intentionally empty category is not as exotic as it may seem. An example of such a category is one that is one that only contains Item Detail Forms that perform specialized scripting, or a category that contains highly customized Report Viewer reports.</para>
        /// <para>When the ID cannot be obtained by CategoryID, the only way to get the ID is to manually look it up in Help | System Information | Category Information.</para>
        /// </remarks>
        /// </summary>
        int CategoryID { get; }
    }
}

using System.Runtime.InteropServices;

namespace Vovin.CmcLibNet.Database.Metadata
{
    /// <summary>
    /// Exposes properties of the currently active Commence view.
    /// </summary>
    [ComVisible(true)]
    [Guid("1B602809-2599-4dac-8BDB-71D0931E909F")]
    public interface IActiveViewInfo
    {
        /// <summary>
        /// Name of the view.
        /// </summary>
        string Name { get; }
        /// <summary>
        /// Type of the view.
        /// </summary>
        string Type { get; }
        /// <summary>
        /// Underlying view category.
        /// </summary>
        string Category { get; }
        /// <summary>
        /// Current itemname (if view is a detail form).
        /// </summary>
        string Item { get; }
        /// <summary>
        /// Currently active fieldname (if view is a detail form).
        /// </summary>
        string Field { get; }
    }
}

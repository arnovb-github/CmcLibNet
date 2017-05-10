using System.Runtime.InteropServices;

namespace Vovin.CmcLibNet.Database
{
    /// <summary>
    /// Exposes properties of the currently active Commence view.
    /// </summary>
    [ComVisible(true)]
    [GuidAttribute("1B602809-2599-4dac-8BDB-71D0931E909F")]
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

    /// <summary>
    /// Exposes properties of the currently active Commence view.
    /// </summary>
    [ComVisible(true)]
    [GuidAttribute("BE123F26-7226-41ea-9DC4-9B6B20C9DC2C")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IActiveViewInfo))]
    public class ActiveViewInfo : IActiveViewInfo
    {
        internal ActiveViewInfo() { }
        /// <inheritdoc />
        public string Name { get; internal set; }
        /// <inheritdoc />
        public string Type { get; internal set; }
        /// <inheritdoc />
        public string Category { get; internal set; }
        /// <inheritdoc />
        public string Item { get; internal set; }
        /// <inheritdoc />
        public string Field { get; internal set; }
    }
}

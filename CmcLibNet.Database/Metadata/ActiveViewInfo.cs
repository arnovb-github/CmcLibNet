using System.Runtime.InteropServices;

namespace Vovin.CmcLibNet.Database.Metadata
{

    /// <summary>
    /// Exposes properties of the currently active Commence view.
    /// </summary>
    [ComVisible(true)]
    [Guid("BE123F26-7226-41ea-9DC4-9B6B20C9DC2C")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IActiveViewInfo))]
    public class ActiveViewInfo : IActiveViewInfo
    {
        internal ActiveViewInfo() { } // prevent .Net consumers newing this up directly
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
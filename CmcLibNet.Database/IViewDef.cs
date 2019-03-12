using System.Runtime.InteropServices;

namespace Vovin.CmcLibNet.Database
{
    /// <summary>
    /// Exposes member of the view definition.
    /// </summary>
    [ComVisible(true)]
    [Guid("94D7005B-1E3F-45d1-B566-9A0F55BE3A3C")]
    public interface IViewDef
    {
        /// <summary>
        /// View name.
        /// </summary>
        string Name { get; }
        /// <summary>
        /// View Type.
        /// </summary>
        string TypeDescription { get; }
        /// <summary>
        /// Underlying Commence category.
        /// </summary>
        string Category { get;}
        /// <summary>
        /// Filename (e.g. for Report Views).
        /// </summary>
        string FileName { get; }
    }
}

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

    /// <summary>
    /// Holds information on the view definition.
    /// </summary>
    [ComVisible(true)]
    [Guid("7326FF13-98BC-466c-B45F-44F1656B605C")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IViewDef))]
    public class ViewDef : IViewDef
    {
        
        internal ViewDef() { }
        /// <summary>
        /// ViewType as enum, is a little more practical in some circumstances
        /// </summary>
        internal CommenceViewType Type { get; set; } //enum isn't public so internal.
        /// <inheritdoc />
        public string Name { get; internal set; }
        /// <inheritdoc />
        public string TypeDescription { get; internal set; }
        /// <inheritdoc />
        public string Category { get; internal set; }
        /// <inheritdoc />
        public string FileName { get; internal set; }
    }
}

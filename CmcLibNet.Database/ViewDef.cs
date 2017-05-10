using System.ComponentModel;
using System.Runtime.InteropServices;

namespace Vovin.CmcLibNet.Database
{
    #region Enumerations
    /// <summary>
    /// Enum for the different viewtypes Commence has.
    /// The Description matches the view types as returned by Commence.
    /// </summary>
    internal enum CommenceViewType
    {
        [Description("Report")]
        Report = 0,
        [Description("Calendar")]
        Calendar = 1,
        [Description("Book")]
        Book = 2,
        [Description("Report Viewer")]
        ReportViewer = 3,
        [Description("Grid")]
        Grid = 4,
        [Description("Add Item")] // this is the string Commence returns!
        ItemDetail = 5,
        [Description("Multi-View")]
        MultiView = 6,
        [Description("Document")]
        Document = 7
    }
    #endregion

    /// <summary>
    /// Exposes member of the view definition.
    /// </summary>
    [ComVisible(true)]
    [GuidAttribute("94D7005B-1E3F-45d1-B566-9A0F55BE3A3C")]
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
    [GuidAttribute("7326FF13-98BC-466c-B45F-44F1656B605C")]
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

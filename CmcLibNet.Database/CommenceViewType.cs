using System.ComponentModel;
using Vovin.CmcLibNet.Attributes;

namespace Vovin.CmcLibNet.Database
{
    /// <summary>
    /// Enum for the different viewtypes Commence has.
    /// The Description matches the view types as returned by Commence.
    /// </summary>
    internal enum CommenceViewType
    {
        [Description("Report")]
        [StringValue("Report")]
        [CursorCreatable(true)]
        Report = 0,
        [Description("Calendar")]
        [StringValue("Calendar")]
        [CursorCreatable(false)]
        Calendar = 1,
        [Description("Book")]
        [StringValue("Book")]
        [CursorCreatable(true)]
        Book = 2,
        [Description("Report Viewer")]
        [StringValue("Report Viewer")]
        [CursorCreatable(false)]
        ReportViewer = 3,
        [Description("Grid")]
        [StringValue("Grid")]
        [CursorCreatable(true)]
        Grid = 4,
        [Description("Add Item")] // yes, this is the string that Commence returns!
        [StringValue("Add Item")] // yes, this is the string that Commence returns!
        [CursorCreatable(false)]
        ItemDetail = 5,
        [Description("Multi-View")]
        [StringValue("Multi-View")]
        [CursorCreatable(false)]
        MultiView = 6,
        [Description("Document")]
        [StringValue("Document")]
        [CursorCreatable(false)]
        Document = 7,
        [Description("Gantt Chart")]
        [StringValue("Gantt Chart")]
        [CursorCreatable(false)]
        Gantt = 8
    }
}
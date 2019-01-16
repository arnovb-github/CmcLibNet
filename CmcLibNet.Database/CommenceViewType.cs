using System.ComponentModel;

namespace Vovin.CmcLibNet.Database
{
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
}

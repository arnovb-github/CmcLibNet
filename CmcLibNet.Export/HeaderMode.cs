using System.Runtime.InteropServices;

namespace Vovin.CmcLibNet.Export
{
    /// <summary>
    /// Enum specifying what headers (or nodenames) to use.
    /// </summary>
    [ComVisible(true)]
    [Guid("4793B385-DCE9-4A87-B557-20423EC0F1BB")]
    public enum HeaderMode
    {
        /// <summary>
        /// Use fieldnames as headers.
        /// </summary>
        Fieldname = 0,
        /// <summary>
        /// Use columnnames as headers. Only applies when exporting views.
        /// If columns have the same label, a sequence number is added. If no columlabel is defined, the underlying fieldname is used.
        /// </summary>
        Columnlabel = 1, // remember the columnname can be empty, Commence then defaults to the fieldname.
        /// <summary>
        /// Use custom headers. The number of supplied headers must be equal to the number of columns to export and they must be unique, regardless of whether SkipConnectedColumns is used.
        /// </summary>
        CustomLabel = 2 // must be correct number AND unique.
    }
}

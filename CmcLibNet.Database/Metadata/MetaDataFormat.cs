using System.Runtime.InteropServices;

namespace Vovin.CmcLibNet.Database.Metadata
{
    /// <summary>
    /// Formats that schema information can be written to.
    /// </summary>
    [ComVisible(true)]
    [Guid("6A199C16-F590-4DCB-9245-47DC9295BA45")]
    public enum MetaDataFormat
    {
        /// <summary>
        /// JSON
        /// </summary>
        Json = 0,
        /// <summary>
        /// XML
        /// </summary>
        Xml = 1,
    }
}
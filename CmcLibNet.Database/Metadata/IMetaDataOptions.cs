using System.Runtime.InteropServices;

namespace Vovin.CmcLibNet.Database.Metadata
{
    /// <summary>
    /// Options for generating metadata file.
    /// </summary>
    [ComVisible(true)]
    [Guid("5F7A4247-240D-4560-9E3A-F14524B0AC49")]
    public interface IMetaDataOptions
    {
        /// <summary>
        /// <see cref="MetaDataFormat"/>
        /// </summary>
        MetaDataFormat Format { get; set; }
        /// <summary>
        /// Include form script.
        /// </summary>
        /// <remarks>Will be included as CData in XML.</remarks>
        bool IncludeFormScript { get; set; }
        /// <summary>
        /// Include form xml.
        /// </summary>
        /// <remarks>Will be included as CData in XML.</remarks>
        bool IncludeFormXml { get; set; }
    }
}
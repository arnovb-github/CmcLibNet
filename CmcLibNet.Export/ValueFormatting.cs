namespace Vovin.CmcLibNet.Export
{
    /// <summary>
    /// Enum for data output formats.
    /// </summary>
    internal enum ValueFormatting
    {
        /// <summary>
        /// No formatting, use data as-is. This means formatted to whatever formatting Commence inherits from the system.
        /// </summary>
        None = 0,
        /// <summary>
        /// Return canonical format as defined by Commence. Unlike Commence, CmcLibNet returns connected data in the format as well. <seealso cref="CmcOptionFlags.Canonical"/>
        /// </summary>
        Canonical = 1,
        /// <summary>
        /// Return data compliant with ISO 8601 format. http://www.iso.org/iso/iso8601
        /// </summary>
        XSD_ISO8601 = 2,
    }
}

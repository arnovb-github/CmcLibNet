using System;
using System.Runtime.InteropServices;

namespace Vovin.CmcLibNet.Database.Metadata
{
    /// <summary>
    /// Options for generating metadata file.
    /// </summary>
    [ComVisible(true)]
    [Guid("F3213CD4-3289-4FEB-9905-6A45E6189600")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IMetaDataOptions))]
    public class MetaDataOptions : IMetaDataOptions
    {
        /// <inheritdoc />
        public MetaDataFormat Format { get; set; }
        /// <inheritdoc />
        public bool IncludeFormXml { get; set; }
        /// <inheritdoc />
        public bool IncludeFormScript { get; set; }
    }
}

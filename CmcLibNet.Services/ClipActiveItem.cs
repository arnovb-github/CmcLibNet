using System;
using System.Runtime.InteropServices;

namespace Vovin.CmcLibNet.Services
{
    /// <summary>
    /// Specifies how active item fields will be copied to clipboard.
    /// </summary>
    [ComVisible(true)]
    [Guid("A10A1A10-A862-4E44-9C6B-89C0588C6082")]
    public enum ClipActiveItem
    {
        /// <summary>
        /// Just the values.
        /// </summary>
        ValuesOnly = 0,
        /// <summary>
        /// Fieldnames and values.
        /// </summary>
        IncludeFieldName = 1,
        /// <summary>
        /// Columnames and values.
        /// </summary>
        IncludeColumnLabel = 2,
        /// <summary>
        /// Fieldnames, columnames and values.
        /// </summary>
        IncludeAll = 3
    }
}

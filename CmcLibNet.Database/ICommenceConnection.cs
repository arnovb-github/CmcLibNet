using System;
using System.Runtime.InteropServices;

namespace Vovin.CmcLibNet.Database
{
    /// <summary>
    /// Describes a Commence connection.
    /// </summary>
    [ComVisible(true)]
    [Guid("7D0CE94D-F112-4BA4-A02F-04D0C8B3241A")]
    public interface ICommenceConnection
    {
        /// <summary>
        /// Name of connection.
        /// </summary>
        [DispId(1)]
        string Name { get; set; }
        /// <summary>
        /// Connected category name.
        /// </summary>
        [DispId(2)]
        string ToCategory { get; set; }

        /// <summary>
        /// Fully qualified name
        /// </summary>
        /// <exception cref="NullReferenceException">Name or ToCategory not specfied.</exception>
        [DispId(0)]
        string FullName { get; }
    }
}

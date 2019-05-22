using System.Runtime.InteropServices;

namespace Vovin.CmcLibNet.Database
{
    /// <summary>
    /// Distinguishes GetRowSetBy* parameters
    /// </summary>
    [ComVisible(true)]
    [Guid("E5C36BD5-184F-4C66-9BDE-B1CFB65E6C5C")]
    internal enum RowSetIdentifier
    {
        RowId = 0,
        Thid = 1
    }
}

using System;
using System.Runtime.InteropServices;

namespace Vovin.CmcLibNet.Database
{
    /// <summary>
    /// Cursor's rowpointer enum
    /// </summary>
    [ComVisible(true)]
    [Guid("A2EA7D1C-85CE-4bfa-8E7F-764C01F1546C")]
    public enum CmcCursorBookmark
    {
        /// <summary>
        /// Move rows starting from first row in cursor.
        /// </summary>
        Beginning = 0,
        /// <summary>
        /// Move rows starting from current row in cursor.
        /// </summary>
        Current = 1,
        /// <summary>
        /// Move rows starting from last row in cursor.
        /// </summary>
        End = 2
    }
}

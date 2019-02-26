using System;
using System.Runtime.InteropServices;

namespace Vovin.CmcLibNet.Export
{
    /// <summary>
    /// Enum specifying the update options for exporting to Excel
    /// </summary>
    [ComVisible(true)]
    [Guid("29B32C40-C6C8-461D-AA56-426E31F814B4")]
    public enum ExcelUpdateOptions
    {
        /// <summary>
        /// Create a new workbook. Overwrites existing workbook.
        /// </summary>
        CreateNewWorkbook = 0,
        /// <summary>
        /// Create new sheet in workbook.
        /// </summary>
        CreateNewWorksheet = 1,
        /// <summary>
        /// (Default) Replaces the workheet. Overwrites existing worksheet.
        /// </summary>
        ReplaceWorksheet = 2,
        /// <summary>
        /// Append data to the worksheet.
        /// </summary>
        AppendToWorksheet = 3

    }
}

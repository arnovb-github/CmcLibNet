using System;
using System.Runtime.InteropServices;

namespace Vovin.CmcLibNet.Database
{
    /// <summary>
    /// Link modes used with View linking. Only used in the GetViewToFile method.
    /// </summary>
    [ComVisible(true)]
    [Guid("1E6E5FBB-A9CB-4838-84E2-B0691F7CB935")]
    public enum CmcLinkMode
    {
        /// <summary>
        /// No linkmode.
        /// </summary>
        None = 0,
        /// <summary>
        /// Treat view as if filtered on selected item.
        /// </summary>
        SelectedItem = 1,
        /// <summary>
        /// Treat view as if filtered on selected date
        /// </summary>
        SelectedDate = 2,
        /// <summary>
        /// Treat view as if filtered on between data A and date B
        /// </summary>
        SelectedDateRange = 3
    }
}
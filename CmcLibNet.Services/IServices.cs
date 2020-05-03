using System;
using System.Runtime.InteropServices;

namespace Vovin.CmcLibNet.Services
{
    /// <summary>
    /// Interface for Services.
    /// </summary>
    [ComVisible(true)]
    [Guid("D719FD36-8828-426A-97E3-F9EADCF45A91")]
    public interface IServices
    {
        /// <summary>
        /// Copies the values of the active item to clipboard. If no view is active, or the viewtype cannot be queried (like Document, Report Viewer), nothing happens.
        /// </summary>
        /// <param name="flag">Formatting options for columns.</param>
        /// <returns><c>true</c> if data was copied to clipboard.</returns>
        bool CopyActiveItemToClipboard(ClipActiveItem flag = ClipActiveItem.ValuesOnly);
        /// <summary>
        /// Does not do anything, just in there for consistency.
        /// </summary>
        void Close();
    }
}

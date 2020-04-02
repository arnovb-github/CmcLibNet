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
        /// Close any references to Commence. The object should be disposed after this.
        /// </summary>
        /// <remarks>When used from within a Commence Form Script, failing to call the <c>Close</c> method will leave the commence.exe process running in the background when the user closes Commence. IMPORTANT: this also happens when an unhandled exception (a 'script error') occurs. The Commence process then has to be closed manually from the Windows Task Manager. Be careful to implement proper error handling.
        /// <para>When the assembly is called from a.NET application, there is rarely a need to call this method, unless you want to explicitly release COM references.</para>
        /// <para>Technical details: calling this method tells the assembly to release all COM handles (called 'RCW' for 'runtime callable wrapper') to Commence that are open. This is needed because when the object reference to this assembly is set to Nothing (in VB), the .NET assembly may not be notified and will think they are still in use. Garbage Collection will therefore not release them, and the commence.exe process will not be terminated.</para>
        /// </remarks>
        void Close();
    }
}

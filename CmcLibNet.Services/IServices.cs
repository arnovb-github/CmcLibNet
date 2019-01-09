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
        /// Gets the places where a particular field is used. For instance, in what Views, what Detail Forms, etc.
        /// <para>Does NOT return an exhaustive list. Do NOT rely on this for finding all references to a field. Some Commence parts are simply not exposed to programming.
        /// For instance, this function will not return whether a field is used in a filter, or agent.</para>
        /// </summary>
        /// <param name="categoryName">Commence categoryname.</param>
        /// <param name="fieldName">Commence fieldname.</param>
        /// <returns>JSON string summarizing where a field is used.</returns>
        /// <remarks>This call takes a bit of time because it queries a lot of database proprties and files!</remarks>
        string FieldUsage(string categoryName, string fieldName);
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

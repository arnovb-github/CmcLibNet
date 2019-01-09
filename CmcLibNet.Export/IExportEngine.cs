using System.Runtime.InteropServices;

namespace Vovin.CmcLibNet.Export
{
    /// <summary>
    /// interface for ExportEngine.
    /// </summary>
    [ComVisible(true)]
    [Guid("F7E7295C-7151-41BC-BACE-4CB789EDA5B1")]
    public interface IExportEngine : IExportEngineEvents
    {
        /// <summary>
        /// Controls the export settings.
        /// </summary>
        IExportSettings Settings { get; }
        /// <summary>
        /// Export a view.
        /// </summary>
        /// <param name="viewName">Commence view name (case-sensitive). Pass an empty string to use the active view, if any. Note that not all view types can be exported.</param>
        /// <param name="fileName">Fully qualified pathname to export file. Will be overwritten if exists.</param>
        /// <param name="settings">ExportSettings.</param>
        /// <remarks>Some export formats (such as Excel) do not require a filename. In those cases pass <c>null</c> or an empty string as filename.</remarks>
        void ExportView(string viewName, string fileName, IExportSettings settings = null);
        /// <summary>
        /// Export a category.
        /// </summary>
        /// <param name="categoryName">Commence category name.</param>
        /// <param name="fileName">Fully qualified pathname to export file. Will be overwritten if exists.</param>
        /// <param name="settings">ExportSettings.</param>
        /// <remarks>Some export formats (such as Excel) do not require a filename. In those cases pass <c>null</c> or an empty string as filename.</remarks>
        void ExportCategory(string categoryName, string fileName, IExportSettings settings = null);
        /// <summary>
        /// Close any references to Commence. The object should be disposed after this.
        /// </summary>
        /// <remarks>When used from within a Commence Form Script, failing to call the <c>Close</c> method will leave the commence.exe process running in the background when the user closes Commence. IMPORTANT: this also happens when an unhandled exception (a 'script error') occurs. The Commence process then has to be closed manually from the Windows Task Manager. Be careful to implement proper error handling.
        /// <para>When the assembly is called from a.NET application, there is rarely a need to call this method, unless you want to explicitly release COM references and/or release memory. It can be useful in some cases, because Commence may complain about running out of memory before the Garbage Collector has a chance to kick in.</para>
        /// <para>Technical details: calling this method tells the assembly to release all COM handles (called 'RCW' for 'runtime callable wrapper') to Commence that are open. This is needed because when the object reference to this assembly is set to Nothing (in VB), the .NET assembly may not be notified and will think they are still in use. Garbage Collection will therefore not release them, and the commence.exe process will not be terminated.</para>
        /// </remarks>
        void Close();
    }
}

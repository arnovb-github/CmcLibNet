using System.Collections.Generic;
using System.Runtime.InteropServices;
using Vovin.CmcLibNet.Database;

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
        /// <param name="fileName">Fully qualified filename.</param>
        /// <param name="settings"><see cref="IExportSettings"/></param>
        void ExportView(string viewName, string fileName, IExportSettings settings = null);
        /// <summary>
        /// Export a category.
        /// </summary>
        /// <param name="categoryName">Commence category name.</param>
        /// <param name="fileName">Fully qualified filename.</param>
        /// <param name="settings"><see cref="IExportSettings"/></param>
        void ExportCategory(string categoryName, string fileName, IExportSettings settings = null);
        /// <summary>
        /// Export a filtered category.
        /// </summary>
        /// <param name="categoryName">Commence category name.</param>
        /// <param name="filters">Filter(s) to apply to category</param>
        /// <param name="fileName">Fully qualified filename.</param>
        /// <param name="settings"><see cref="IExportSettings"/></param>
        [ComVisible(false)]
        void ExportCategory(string categoryName, IEnumerable<ICursorFilter> filters, string fileName, IExportSettings settings = null);
        /// <summary>
        /// Not needed, does nothing.
        /// </summary>
        void Close();
    }
}
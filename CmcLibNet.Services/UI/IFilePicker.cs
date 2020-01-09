using System.Runtime.InteropServices;

namespace Vovin.CmcLibNet.Services.UI
{
    /// <summary>
    /// Interface for FilePicker
    /// </summary>
    [ComVisible(true)]
    [Guid("EB19C890-5EED-4388-80E1-2F0DDA490902")]
    public interface IFilePicker
    {
        /// <summary>
        /// Gets or sets the initial directory displayed by the file dialog box.
        /// </summary>
        string InitialDirectory { get; set; }
        /// <summary>
        /// Gets or sets the index of the filter currently selected in the file dialog box.
        /// The default value is 1.
        /// </summary>
        int FilterIndex { get; set; }
        /// <summary>
        /// Gets or sets the current file name filter string, which determines the choices
        /// that appear in the "Save as file type" or "Files of type" box in the dialog box.
        /// </summary>
        string Filter { get; set; }
        /// <summary>
        /// Gets or sets the file dialog box title.
        /// </summary>
        string Title { get; set; }
        /// <summary>
        /// Gets or sets the default file name extension.
        /// </summary>
        string DefaultExt { get; set; }

        /// <summary>
        /// Returns an OpenFileDialog instance.
        /// </summary>
        /// <returns>Dialog window.</returns>
        string ShowDialog();
    }
}
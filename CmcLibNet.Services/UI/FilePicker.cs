using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace Vovin.CmcLibNet.Services.UI
{
    /// <summary>
    /// From VBScript, it is notoriously hard to get an OpenFileDialog.
    /// This class is intended primarily for use from Commence scripts.
    /// All it does is expose an OpenFileDialog
    /// </summary>
    [ComVisible(true)]
    [ProgId("CmcLibNet.Services.FilePicker")]
    [Guid("1E09B0AA-48C0-4F7D-BA5B-0986CEAA315E")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IFilePicker))]
    public class FilePicker : IFilePicker
    {
        /// <inheritdoc />
        public string InitialDirectory { get; set; }
        /// <inheritdoc />
        public int FilterIndex { get; set; }
        /// <inheritdoc />
        public string Filter { get; set; }
        /// <inheritdoc />
        public string Title { get; set; }
        /// <inheritdoc />
        public string DefaultExt { get; set; }

        /// <inheritdoc />
        public string ShowDialog()
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Title = Title;
                openFileDialog.InitialDirectory = ValidatePath(InitialDirectory);
                openFileDialog.Filter = Filter;
                openFileDialog.FilterIndex = FilterIndex == 0 ? 1: FilterIndex ;
                openFileDialog.DefaultExt = DefaultExt;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    return openFileDialog.FileName;
                }
            }
            return string.Empty;
        }

        private string ValidatePath(string path)
        {
            if (string.IsNullOrEmpty(path)) return string.Empty;
            if (System.IO.Directory.Exists(path))
                return System.IO.Path.GetFullPath(path);
            else
                return string.Empty;
        }
    }
}
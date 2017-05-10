using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;

namespace Vovin.CmcLibNet.Services
{
    /// <summary>
    /// Specifies how active item fields will be copied to clipboard.
    /// </summary>
    [ComVisible(true),
    GuidAttribute("A10A1A10-A862-4E44-9C6B-89C0588C6082")]
    public enum ClipActiveItem
    {
        /// <summary>
        /// Just the values.
        /// </summary>
        ValuesOnly = 0,
        /// <summary>
        /// Fieldnames and values.
        /// </summary>
        IncludeFieldName = 1,
        /// <summary>
        /// Columnames and values.
        /// </summary>
        IncludeColumnLabel = 2,
        /// <summary>
        /// Fieldnames, columnames and values.
        /// </summary>
        IncludeAll = 3
    }

    /// <summary>
    /// Class that points to services
    /// </summary>
    [ComVisible(true)]
    [GuidAttribute("BFE25F8A-9F4C-4AD9-8662-84D13B12B4D9")]
    [ProgIdAttribute("CmcLibNet.Services")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IServices))]
    public class Services : IServices
    {
        private CommenceApp _app = null;
        /// <summary>
        /// Public constructor.
        /// </summary>
        public Services()
        {
            _app = new CommenceApp();
        }

        /// <inheritdoc />
        [STAThreadAttribute] // needed because we will be putting data on the clipboard
        public bool CopyActiveItemToClipboard(ClipActiveItem flag = ClipActiveItem.ValuesOnly)
        {
            List<Field> activeItemValues = null;
            ActiveItem ai = new ActiveItem();
            activeItemValues = ai.GetValues();
            if (activeItemValues == null) { return false; }
            // okay, we got something. Let's put the data on the clipboard
            StringBuilder sb = new StringBuilder();
            foreach (Field f in activeItemValues)
            {
                switch (flag) {
                    case ClipActiveItem.ValuesOnly:
                        sb.AppendLine(f.Value);
                        break;
                    case ClipActiveItem.IncludeColumnLabel:
                        sb.AppendLine(f.Label + "\t" + f.Value);
                        break;
                    case ClipActiveItem.IncludeFieldName:
                        sb.AppendLine(f.Name + "\t" + f.Value);
                        break;
                    case ClipActiveItem.IncludeAll:
                        sb.AppendLine(f.Label + "\t" + f.Name + "\t" + f.Value);
                        break;
                    default:
                        sb.AppendLine(f.Label + "\t" + f.Name + "\t" + f.Value);
                        break;
                }
            }
            System.Windows.Forms.Clipboard.SetText(sb.ToString());
            return true;
        }
        /// <inheritdoc />
        public string FieldUsage(string categoryName, string fieldName)
        {
            throw new NotImplementedException("Not yet implemented");
            // views

            // view filters
            // not possible

            // items detail forms

            // report viewer views

            // agents
            // not possible
        }
        /// <inheritdoc />
        public void Close()
        {
            if (_app != null)
            {
                _app.Close();
            }
        }
    }
}

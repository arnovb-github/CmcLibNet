using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;

namespace Vovin.CmcLibNet.Services
{
    /// <summary>
    /// Class that points to services
    /// </summary>
    [ComVisible(true)]
    [Guid("BFE25F8A-9F4C-4AD9-8662-84D13B12B4D9")]
    [ProgId("CmcLibNet.Services")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IServices))]
    public class Services : IServices
    {
        /// <inheritdoc />
        [STAThread] // needed because we will be putting data on the clipboard
        public bool CopyActiveItemToClipboard(ClipActiveItem flag = ClipActiveItem.ValuesOnly)
        {
            List<Field> activeItemValues;
            using (ActiveItem ai = new ActiveItem())
            { 
                activeItemValues = ai.GetValues();
                if (activeItemValues == null) { return false; }
                // okay, we got something. Let's put the data on the clipboard
                StringBuilder sb = new StringBuilder();
                foreach (Field f in activeItemValues)
                {
                    switch (flag)
                    {
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
                if (sb?.Length > 0)
                {
                    System.Windows.Forms.Clipboard.SetText(sb.ToString());
                }
            }

            return true;
        }

        /// <inheritdoc />
        public void Close() { }
    }
}
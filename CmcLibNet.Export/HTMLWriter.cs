using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace Vovin.CmcLibNet.Export
{
    internal class HTMLWriter : BaseWriter
    {
        bool disposed = false;
        StreamWriter _sw = null;
        int _rowcounter = 0;

        #region Contructors
        internal HTMLWriter(Database.ICommenceCursor cursor, IExportSettings settings)
            : base(cursor, settings){}

        ~HTMLWriter()
        {
            Dispose(false);
        }
        #endregion

        protected internal override void WriteOut(string fileName)
        {
            _sw = new StreamWriter(fileName);
            _sw.WriteLine("<!DOCTYPE html><html>");
            _sw.WriteLine("<head>");
            if (base._settings.CSSFile != null)
            {
                _sw.WriteLine("<link rel=\"stylesheet\" type=\"text/css\" href=\"" + base._settings.CSSFile + "\">");
            }
            _sw.WriteLine("<title>" + HtmlEncode(base._dataSourceName) + "</title>");
            _sw.WriteLine("</head>");
            _sw.WriteLine("<body>");
            _sw.WriteLine("<table class=\"cmclibnet-export\">");
            // set table headers
            if (_settings.HeadersOnFirstRow)
            {
                _sw.WriteLine("<thead class=\"cmclibnet-header\"><tr>" + this.GetTableHeaders() + "</tr></thead>"); // meh bullshit
            }
            // start body
            _sw.WriteLine("<tbody>");
            base.ReadData();
        }

        protected internal override void ProcessDataRows(object sender, DataProgressChangedArgs e)
        {
            StringBuilder sb = new StringBuilder();
            
            foreach (List<CommenceValue> row in e.Values)
            {
                _rowcounter++;
                int colcounter = 0;
                sb.Append("<tr class=\"cmclibnet-item\">");
                foreach (CommenceValue v in row)
                {
                    colcounter++;
                    // we can have either no values, or a direct value, or connected values
                    if (v.IsEmpty)
                    {
                        sb.Append("<td class=\"cmclibnet-value\" id=\"r" + _rowcounter + "c" + colcounter + "\"></td>");
                    }
                    else
                    {
                        string s = (v.DirectFieldValue != null) ? v.DirectFieldValue : string.Join(base._settings.TextDelimiterConnections, v.ConnectedFieldValues);
                        sb.Append("<td class=\"cmclibnet-value\" id=\"r" + _rowcounter + "c" + colcounter + "\">" + HtmlEncode(s) + "</td>");
                    }
                } // foreach
                sb.Append("</tr>");
            } // foreach
            _sw.WriteLine(sb.ToString());
        }

        protected internal override void DataReadComplete(object sender, DataReadCompleteArgs e)
        {
            _sw.WriteLine("</tbody></table></body></html>");
            _sw.Flush();
            _sw.Close();
        }

        private string GetTableHeaders()
        {
            StringBuilder sb = new StringBuilder();
            List<string> headers = null;
            if (base._settings.SkipConnectedItems)
            {
                // how can we find the headers belonging to non-connected columns?
                switch (base._settings.HeaderMode)
                {
                    case HeaderMode.Columnlabel:
                        headers = base.ColumnDefinitions.Where(o => !o.IsConnection).Select(o => o.ColumnLabel).ToList();
                        break;
                    case HeaderMode.Fieldname:
                        headers = base.ColumnDefinitions.Where(o => !o.IsConnection).Select(o => o.FieldName).ToList();
                        break;
                    case HeaderMode.CustomLabel:
                        headers = base.ColumnDefinitions.Where(o => !o.IsConnection).Select(o => o.CustomColumnLabel).ToList();
                        break;
                } // switch
            } // if
            else
            {
                headers = base.ExportHeaders;
            }
            foreach (string s in headers)
            {
                sb.Append("<th>" + HtmlEncode(s) + "</th>");
            }
            return sb.ToString();
        }

        /// <summary>
        /// HTML-encodes a string and returns the encoded string.
        /// </summary>
        /// <param name="text">The text string to encode. </param>
        /// <returns>The HTML-encoded text.</returns>
        private static string HtmlEncode(string text)
        {
            if (String.IsNullOrEmpty(text))
                return null;

            StringBuilder sb = new StringBuilder(text.Length);

            int len = text.Length;
            for (int i = 0; i < len; i++)
            {
                switch (text[i])
                {

                    case '<':
                        sb.Append("&lt;");
                        break;
                    case '>':
                        sb.Append("&gt;");
                        break;
                    case '"':
                        sb.Append("&quot;");
                        break;
                    case '&':
                        sb.Append("&amp;");
                        break;
                    default:
                        if (text[i] > 159)
                        {
                            // decimal numeric entity
                            sb.Append("&#");
                            sb.Append(((int)text[i]).ToString(System.Globalization.CultureInfo.InvariantCulture));
                            sb.Append(";");
                        }
                        else
                            sb.Append(text[i]);
                        break;
                }
            }
            return sb.ToString();
        }

        /// <summary>
        /// Dispose method.
        /// </summary>
        /// <param name="disposing">disposing.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposed)
                return;

            if (disposing)
            {
                // Free any other managed objects here.
                //
                if (_sw != null)
                {
                    _sw.Close();
                }
            }

            // Free any unmanaged objects here.
            //
            disposed = true;

            // Call the base class implementation.
            base.Dispose(disposing);
        }
    }
}

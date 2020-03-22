using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;

namespace Vovin.CmcLibNet.Export
{
    internal class TextWriter : BaseWriter
    {
        StreamWriter _sw = null;
        bool disposed = false;

        #region Constructors
        internal TextWriter(Database.ICommenceCursor cursor, IExportSettings settings)
            : base(cursor, settings){}

        ~TextWriter()
        {
            Dispose(false);
        }
        #endregion

        #region Methods
        protected internal override void WriteOut(string fileName)
        {
            if (base.IsFileLocked(new FileInfo(fileName)))
            {
                throw new IOException("File '" + fileName + "' in use.");
            }
            _sw = new StreamWriter(fileName);
            // headers
            if (base._settings.HeadersOnFirstRow)
            {
                List<string> headers = null;
                if (base._settings.SkipConnectedItems)
                {
                    switch (base._settings.HeaderMode)
                    {
                        case HeaderMode.Columnlabel:
                            // following line means: "from columndefinitions collection take all the ones that are not marked as connection and give us some property as a list."
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
                List<string> formattedHeaderList = headers.Select(o => base._settings.TextQualifier + o + base._settings.TextQualifier).ToList();
                _sw.WriteLine(String.Join(base._settings.TextDelimiter, formattedHeaderList));
            }
            base.ReadCommenceData();
        }

        protected internal override void HandleProcessedDataRows(object sender, CursorDataReadProgressChangedArgs e)
        {
            foreach (List<CommenceValue> row in e.RowValues)
            {
                List<string> rowvalues = new List<string>();
                foreach (CommenceValue v in row)
                {
                    if (v.ColumnDefinition.IsConnection) // connection
                    {
                        if (!base._settings.SkipConnectedItems)
                        {
                            if (v.ConnectedFieldValues == null)
                            {
                                rowvalues.Add(base._settings.TextQualifier + String.Join(base._settings.TextDelimiterConnections, string.Empty) + base._settings.TextQualifier);
                            }
                            else
                            {
                                // we concatenate connected values that were split here.
                                // that is not a very good idea.
                                // a much better idea is to not split at all.
                                // not splitting was implemented.
                                // we can leave in the string.Join, we pass just one value to it.
                                rowvalues.Add(base._settings.TextQualifier + String.Join(base._settings.TextDelimiterConnections, v.ConnectedFieldValues) + base._settings.TextQualifier);
                            } // if
                        } // if
                    } //if
                    else
                    {
                        rowvalues.Add(base._settings.TextQualifier + v.DirectFieldValue + base._settings.TextQualifier);
                    } // else
                } // foreach
                _sw.WriteLine(String.Join(base._settings.TextDelimiter, rowvalues));
            } //foreach
            BubbleUpProgressEvent(e);
        }

        protected internal override void HandleDataReadComplete(object sender, ExportCompleteArgs e)
        {
            _sw.Flush();
            _sw.Close();
            base.BubbleUpCompletedEvent(e);
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
        #endregion
    }
}

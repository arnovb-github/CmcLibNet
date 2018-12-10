using System;
using System.Runtime.InteropServices;

namespace Vovin.CmcLibNet.Export
{
    /* Does the same as any export writer except that it just emits the data to an event.
     * It consumes events raised by the datareader and reformats them, then throws an event with the reformattted data.
     * This allows consumers to do what they want with the data without writing them to a file first.
     * Currently only JSON format is supported. BSON and XML may be added in the future.
     */
    internal class EventWriter : BaseWriter
    {

        #region Contructors
        internal EventWriter(Database.ICommenceCursor cursor, IExportSettings settings)
            : base(cursor, settings) { }

        ~EventWriter()
        {
            Dispose(false);
        }
        #endregion

        #region Methods
        protected internal override void WriteOut(string fileName)
        {
            base.ReadData();
        }

        protected internal override void ProcessDataRows(object sender, DataProgressChangedArgs e)
        {
            // construct data, create eventargs, raise event
            JSONCreator ja = new JSONCreator(this);
            ja.AddRowValues(e.Values);
            CommenceRowsReadArgs args = new CommenceRowsReadArgs(base.CurrentRow, base.TotalRows, ja.ToJObject().ToString(Newtonsoft.Json.Formatting.Indented,null));
            OnCommenceRowsRead(args);
        }

        protected internal override void DataReadComplete(object sender, DataReadCompleteArgs e)
        {
            return;
        }
        #endregion
    }
}

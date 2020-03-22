using Newtonsoft.Json.Linq;
using System.Collections.Generic;
using System.Linq;

namespace Vovin.CmcLibNet.Export
{
    /* Does the same as any export writer except that it just emits the data to an event.
     * It consumes events raised by the datareader and reformats them, then throws an event with the reformattted data.
     * This allows consumers to do what they want with the data without writing them to a file first.
     * Currently only JSON format is supported. BSON and XML may be added in the future.
     */
    internal class EventWriter : BaseWriter
    {

        #region Constructors
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
            base.ReadCommenceData();
        }

        protected internal override void HandleProcessedDataRows(object sender, CursorDataReadProgressChangedArgs e)
        {
            // construct data, create eventargs, raise event
            JsonCreator jc = new JsonCreator(this);
            List<JObject> list = jc.SerializeRowValues(e.RowValues);
            var jsonString = "[" + string.Join(",", list.Select(o => o.ToString())) + "]";
            // do custom bubbling up
            ExportProgressChangedArgs args = new ExportProgressChangedArgs(
                e.RowsProcessed,
                e.RowsTotal,
                jsonString);
            base.OnExportProgressChanged(args);
        }

        protected internal override void HandleDataReadComplete(object sender, ExportCompleteArgs e)
        {
            base.BubbleUpCompletedEvent(e);
        }
        #endregion
    }
}

namespace Vovin.CmcLibNet.Export
{
    /// <summary>
    /// Interface for exposing events.
    /// </summary>
    public interface IExportEngineEvents
    {
        /* 
        * Put events in a separate interface so as to not pollute the COM interface.
        * If we put the event in the COM-exposed interface,
        * COM clients will see a add_EventX and remove_EventX method.
        * For COM clients, there is a separate interface implementation for events.
        */

        /// <summary>
        /// ExportProgressChanged event.
        /// </summary>
        event ExportProgressChangedHandler ExportProgressChanged; // event is invisible to COM, there is a separate interface for COM.
        /// <summary>
        /// CommenceRowsRead event used in conjunction with Event export format.
        /// </summary>
        event CommenceRowsReadHandler CommenceRowsRead;
    }
}

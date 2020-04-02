namespace Vovin.CmcLibNet.Export
{
    /// <summary>
    /// Interface for exposing events.
    /// </summary>
    public interface IExportEngineEvents
    {
        /* 
        * Put events in a separate interface so as to not pollute the COM interface.
        * If we were to put the event in the COM-exposed interface,
        * COM clients will see a add_EventX and remove_EventX method.
        * For COM clients, there is a separate interface implementation for events.
        */

        /// <summary>
        /// ExportProgressChanged event raised for every batch of rows read.
        /// </summary>
        event ExportProgressChangedHandler ExportProgressChanged;

        /// <summary>
        /// ExportCompleted event.
        /// </summary>
        event ExportCompletedHandler ExportCompleted;
        
    }
}

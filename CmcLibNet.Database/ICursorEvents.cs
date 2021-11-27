using System;

namespace Vovin.CmcLibNet.Database
{
    /// <summary>
    /// Interface for cursor events not exposed to COM.
    /// </summary>
    public interface ICursorEvents
    {
        /* 
        * Put events in a separate interface so as to not pollute the COM interface.
        * If we were to put the event in the COM-exposed interface,
        * COM clients would see a add_EventX and remove_EventX method.
        */

        /// <summary>
        /// Event for keeping track of the cursor read.
        /// </summary>
        event EventHandler<CursorRowsReadArgs> RowsRead;
    }
}

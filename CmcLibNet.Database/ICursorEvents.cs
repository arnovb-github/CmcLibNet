using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Vovin.CmcLibNet.Database
{
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
        public event EventHandler<CursorRowsReadArgs> RowsRead;
    }
}

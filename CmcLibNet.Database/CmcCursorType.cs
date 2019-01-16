using System;
using System.Runtime.InteropServices;

namespace Vovin.CmcLibNet.Database
{
    /// <summary>
    /// Cursor flags enum; determines what type of cursor to create.
    /// </summary>
    [ComVisible(true)]
    [Guid("7E21D7C0-B0FE-431e-8E74-99FAF8CAED12")]
    public enum CmcCursorType
    {
        /// <summary>
        /// Use a category.
        /// </summary>
        Category = 0,
        /// <summary>
        /// Use a view.
        /// </summary>
        View = 1,
        /// <summary>
        /// Use the fields as defined for the Palm Pilot Address Book.
        /// </summary>
        PilotAb = 2,
        /// <summary>
        /// Use the fields as defined for the Palm Pilot Memo.
        /// </summary>
        PilotMemo = 3,
        /// <summary>
        /// Use the fields as defined for the Palm Pilot To-Do list.
        /// </summary>
        PilotToDo = 5,
        /// <summary>
        /// Use the fields as defined for the Palm Pilot Calendar.
        /// </summary>
        PilotAppt = 6,
        /// <summary>
        /// Use the fields as defined for the Microsoft Outlook Address book.
        /// </summary>
        OutlookAb = 7,
        /// <summary>
        /// Use the fields as defined for the Microsoft Outlook Calendar.
        /// </summary>
        OutlookAppt = 8,
        /// <summary>
        /// Use the fields as defined for the E-mail Log category.
        /// </summary>
        EmailLog = 9,
        /// <summary>
        /// Use the fields as defined for the Microsoft Outlook Tasks.
        /// </summary>
        OutlookTask = 10,
        /// <summary>
        /// Based cursor on the view used when Send Letter command was called.
        /// </summary>
        Merge = 11,
    }
}

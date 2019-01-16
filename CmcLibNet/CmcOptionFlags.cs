using System;
using System.Runtime.InteropServices;

namespace Vovin.CmcLibNet
{
    /// <summary>
    /// Commence option flags.
    /// </summary>
    [ComVisible(true)]
    [Guid("1382A6E0-19E5-41c1-9BE7-10545F538DBE")]
    [FlagsAttribute]
    public enum CmcOptionFlags
    {
        //If bitwise operators are making your head spin,
        //.NET 4.0 has introduced the method HasFlag which can be used as follows: .HasFlag(enum.value)
        /// <summary>
        /// Default flag.
        /// </summary>
        Default = 0x0,
        /// <summary>
        /// Indicates fieldnames should be used, columnnames ignored.
        /// </summary>
        Fieldname = 0x0001,
        /// <summary>
        /// Used to request all fields from a category.
        /// </summary>
        All = 0x0002,
        /// <summary>
        /// Mark item as shared when adding items.
        /// </summary>
        Shared = 0x0004,
        /// <summary>
        /// Changes from 3Com Palm Pilot.
        /// </summary>
        [Obsolete]
        PalmPilot = 0x0008,
        /// <summary>
        /// Make Commence return data in canonical (i.e. consistent) format.
        /// <remarks>
        /// <list type="table">
        /// <listheader><term>Datatype</term><description>Format notes</description></listheader>
        /// <item><term>Date</term><description>yyyymmdd</description></item>
        /// <item><term>Time</term><description>hh:mm military time, 24 hour clock.</description></item>
        /// <item><term>Number</term><description>123456.78, no thousand separator, period for decimal delimiter. Note that when a field is defined as 'Show as currency' in the Commence UI, numerical values are prepended with a '$' sign.</description></item>
        /// <item><term>CheckBox</term><description>TRUE or FALSE.</description></item>
        /// </list>
        /// </remarks>
        /// </summary>
        Canonical = 0x0010,
        /// <summary>
        /// Allows for the Agent subsystem to distinguish between Internet/Intranet database operations.
        /// </summary>
        Internet = 0x0020,
        /// <summary>
        /// (Undocumented by Commence) Make Commence return THIDs instead of Name field values.
        /// Use *RowSet.GetRowID() on a cursor defined with thids flag to get a row's THID
        /// </summary>
        UseThids = 0x0100,
        /// <summary>
        /// (Undocumented by Commence) Unknown.
        /// </summary>
        IgnoreSyncCondition= 0x0200
    }
}
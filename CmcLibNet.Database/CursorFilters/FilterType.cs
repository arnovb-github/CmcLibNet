using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace Vovin.CmcLibNet.Database
{
    /// <summary>
    /// Filter Types that a Commence ViewFilter request can use, i.e. the filters you can apply to Commence.
    /// </summary>
    [ComVisible(true)]
    [Guid("55B52ADD-2B54-489D-AF40-800F750E2EE1")]
    public enum FilterType 
    {
        /// <summary>
        /// Filter on fieldvalue.
        /// </summary>
        [Description("Field (F)")]
        Field = 0,
        /// <summary>
        /// Filter on connection to specific item.
        /// </summary>
        [Description("Connection To Item (CTI)")]
        ConnectionToItem = 1,
        /// <summary>
        /// Filter on connected item to connected item.
        /// </summary>
        [Description("Connection To Category To Item (CTCTI")]
        ConnectionToCategoryToItem =2,
        /// <summary>
        /// Filter on fieldvalue in connected category.
        /// </summary>
        [Description("Connection To Category Field (CTCF)")]
        ConnectionToCategoryField = 3
    }
}

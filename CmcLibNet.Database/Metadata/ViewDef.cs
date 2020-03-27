using System;
using System.Runtime.InteropServices;

namespace Vovin.CmcLibNet.Database.Metadata
{
    /// <summary>
    /// Holds information on the view definition.
    /// </summary>
    [ComVisible(true)]
    [Guid("7326FF13-98BC-466c-B45F-44F1656B605C")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IViewDef))]
    [Serializable]
    public class ViewDef : IViewDef
    {
        internal ViewDef() { } // required for XML serialization
        /// <summary>
        /// Using ViewType enum is a little more practical in some circumstances
        /// </summary>
        internal  CommenceViewType ViewType { get; set; }
        /// <inheritdoc />
        public string Name { get; internal set; }
        /// <inheritdoc />
        public string Type { get; internal set; }
        /// <inheritdoc />
        public string Category { get; internal set; }
        /// <inheritdoc />
        public string FileName { get; internal set; }
    }
}
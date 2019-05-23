using System;
using System.Runtime.InteropServices;

namespace Vovin.CmcLibNet.Database.Metadata
{

    /// <summary>
    /// Holds information on the database definition.
    /// </summary>
    [ComVisible(true)]
    [Guid("7ABBCC06-26AE-4a34-B07D-873F9B0A9C26")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IDBDef))]
    [Serializable]
    public class DBDef : IDBDef
    {
        internal DBDef() { }
        /// <inheritdoc />
        public string Name { get; internal set; }
        /// <inheritdoc />
        public string Path { get; internal set; }
        /// <inheritdoc />
        public bool Attached { get; internal set; }
        /// <inheritdoc />
        public bool Connected { get; internal set; }
        /// <inheritdoc />
        public bool IsServer { get; internal set; }
        /// <inheritdoc />
        public bool IsClient { get; internal set; }
        /// <inheritdoc />
        public string Username { get; internal set; }
        /// <inheritdoc />
        public string Spoolpath { get; internal set; }
    }
}

using System;
using System.Collections.Generic;

namespace Vovin.CmcLibNet.Database.Metadata
{
    /// <summary>
    /// Schema information
    /// </summary>
    /// <remarks>Exposes concrete objects and not interfaces because of XML serialization requirement.</remarks>
    [Serializable]
    public class DatabaseSchema : IDatabaseSchema
    {
        internal readonly IDBDef _definition;
        internal DatabaseSchema(IDBDef definition)
        {
            _definition = definition;
        }
        /// <summary>
        /// Empty constructor required for XML serialization.
        /// </summary>
        internal DatabaseSchema() { }

        /// <inheritdoc />
        public string Name
        {
            get
            {
                return _definition.Name;
            }
            set { } // empty setter needed for serialization
        }
        /// <inheritdoc />
        public string Path
        {
            get
            {
                return _definition.Path;
            }
            set { }
        }
        /// <inheritdoc />
        public bool Attached
        {
            get
            {
                return _definition.Attached;
            }
            set { }
        }
        /// <inheritdoc />
        public bool Connected
        {
            get
            {
                return _definition.Connected;
            }
            set { }
        }
        /// <inheritdoc />
        public bool IsServer
        {
            get
            {
                return _definition.IsServer;
            }
            set { }
        }
        /// <inheritdoc />
        public bool IsClient
        {
            get
            {
                return _definition.IsClient;
            }
            set { }
        }
        /// <inheritdoc />
        public string Username
        {
            get
            {
                return _definition.Username;
            }
            set { }
        }
        /// <inheritdoc />
        public string Spoolpath
        {
            get
            {
                return _definition.Spoolpath;
            }
            set { }
        }
        /// <inheritdoc />
        public List<CommenceCategoryMetaData> Categories { get; internal set; }
        /// <inheritdoc />
        public long Size { get; set; }
    }
}

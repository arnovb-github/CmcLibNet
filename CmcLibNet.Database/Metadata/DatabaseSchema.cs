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
        private readonly IDBDef _definition;
        internal DatabaseSchema(IDBDef definition)
        {
            _definition = definition;
        }
        /// <summary>
        /// Empty public constructor required for XML serialization.
        /// </summary>
        internal DatabaseSchema() { }
        /// <inheritdoc />
        public string Name => _definition.Name;
        /// <inheritdoc />
        public string Path => _definition.Path;
        /// <inheritdoc />
        public bool Attached => _definition.Attached;
        /// <inheritdoc />
        public bool Connected => _definition.Connected;
        /// <inheritdoc />
        public bool IsServer => _definition.IsServer;
        /// <inheritdoc />
        public bool IsClient => _definition.IsClient;
        /// <inheritdoc />
        public string Username => _definition.Username;
        /// <inheritdoc />
        public string Spoolpath => _definition.Spoolpath;
        /// <inheritdoc />
        public List<CommenceCategoryMetaData> Categories { get; internal set; }
        /// <inheritdoc />
        public long Size { get; set; }
    }
}

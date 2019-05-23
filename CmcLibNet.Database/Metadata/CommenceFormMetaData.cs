using System;

namespace Vovin.CmcLibNet.Database.Metadata
{
    /// <summary>
    /// Commence Item detail form metadata.
    /// </summary>
    [Serializable]
    public class CommenceFormMetaData : ICommenceFormMetaData
    {
        internal CommenceFormMetaData() { } // required for XML serialization

        internal CommenceFormMetaData(string name)
        {
            Name = name;
        }

        /// <inheritdoc />
        public string Name { get; set; }
        /// <inheritdoc />
        public string Script { get; set; }
        /// <inheritdoc />
        public string Xml { get; set; }
        /// <inheritdoc />
        public string Category { get; set; }
        /// <inheritdoc />
        public string Path { get; set; }
    }
}

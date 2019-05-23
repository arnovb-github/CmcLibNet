using System;

namespace Vovin.CmcLibNet.Database.Metadata
{

    /// <summary>
    /// Commence view information.
    /// </summary>
    [Serializable]
    public class CommenceViewMetaData : ICommenceViewMetaData
    {
        private readonly IViewDef _definition;

        /// <summary>
        /// Empty public constructor required for XML serialization.
        /// </summary>
        internal CommenceViewMetaData() { }

        internal CommenceViewMetaData(string name, IViewDef definition)
        {
            _definition = definition;
        }
        /// <inheritdoc />
        public string Type => _definition.Type;
        /// <inheritdoc />
        public string Category => _definition.Category;
        /// <inheritdoc />
        public string FileName => _definition.FileName;
        /// <inheritdoc />
        public string Name => _definition.Name;
    }
}

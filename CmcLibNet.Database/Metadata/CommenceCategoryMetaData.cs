using System;
using System.Collections.Generic;

namespace Vovin.CmcLibNet.Database.Metadata
{
    /// <summary>
    /// Commence category metadata.
    /// </summary>
    [Serializable]
    public class CommenceCategoryMetaData : ICommenceCategoryMetaData
    {
        private readonly ICategoryDef _definition;

        internal CommenceCategoryMetaData() { } // required for XML serialization
        internal CommenceCategoryMetaData(string name, ICategoryDef definition)
        {
            Name = name;
            _definition = definition;
        }
        /// <inheritdoc />
        public string Name { get; set; }
        /// <inheritdoc />
        public int MaxItems => _definition.MaxItems;
        /// <inheritdoc />
        public bool Shared => _definition.Shared;
        /// <inheritdoc />
        public bool Duplicates => _definition.Duplicates;
        /// <inheritdoc />
        public bool Clarified => _definition.Clarified;
        /// <inheritdoc />
        public string ClarifySeparator => _definition.ClarifySeparator;
        /// <inheritdoc />
        public string ClarifyField => _definition.ClarifyField;
        /// <inheritdoc />
        public int CategoryID => _definition.CategoryID;
        /// <inheritdoc />
        public int Items { get; set; }
        /// <inheritdoc />
        public List<CommenceViewMetaData> Views { get; set; }
        /// <inheritdoc />
        public List<CommenceFieldMetaData> Fields { get; set; }
        /// <inheritdoc />
        public List<CommenceFormMetaData> Forms { get; set; }
        /// <inheritdoc />
        public List<CommenceConnection> Connections { get; set; }
    }
}

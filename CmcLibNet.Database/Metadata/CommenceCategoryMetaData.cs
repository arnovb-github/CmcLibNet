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
        public int MaxItems
        {
            get
            {
                return _definition.MaxItems;
            }
            set { }
        }

        /// <inheritdoc />
        public bool Shared
        {
            get
            {
                return _definition.Shared;
            }
            set { }
        }

        /// <inheritdoc />
        public bool Duplicates
        {
            get
            {
                return _definition.Duplicates;
            }
            set { }
        }

        /// <inheritdoc />
        public bool Clarified
        {
            get
            {
                return _definition.Clarified;
            }
            set { }
        }

        /// <inheritdoc />
        public string ClarifySeparator
        {
            get
            {
                return _definition.ClarifySeparator;
            }
            set { }
        }

        /// <inheritdoc />
        public string ClarifyField
        {
            get
            {
                return _definition.ClarifyField;
            }
            set { }
        }

        /// <inheritdoc />
        public int CategoryID
        {
            get
            {
                return _definition.CategoryID;
            }
            set { }
        }

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

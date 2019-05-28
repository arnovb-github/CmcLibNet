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
        /// Empty constructor required for XML serialization.
        /// </summary>
        internal CommenceViewMetaData() { }

        internal CommenceViewMetaData(string name, IViewDef definition)
        {
            _definition = definition;
        }
        /// <inheritdoc />
        public string Type
        {
            get
            {
                return _definition.Type;
            }
            set { }
        }

        /// <inheritdoc />
        public string Category
        {
            get
            {
                return _definition.Category;
            }
            set { }
        }

        /// <inheritdoc />
        public string FileName
        {
            get
            {
                return _definition.FileName;
            }
            set { }
        }

        /// <inheritdoc />
        public string Name
        {
            get
            {
                return _definition.Name;
            }
            set { }
        }
    }
}

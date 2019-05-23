using System;

namespace Vovin.CmcLibNet.Database.Metadata
{
    /// <summary>
    /// Commence field metadata.
    /// </summary>
    [Serializable]
    public class CommenceFieldMetaData :  ICommenceFieldMetaData
    {
        private readonly ICommenceFieldDefinition _definition;
        
        internal CommenceFieldMetaData() { }
        internal CommenceFieldMetaData(string name, ICommenceFieldDefinition definition)
        {
            Name = name;
            _definition = definition;
        }
        /// <inheritdoc />
        public CommenceFieldType Type => _definition.Type;
        /// <inheritdoc />
        public bool Shared => _definition.Shared;
        /// <inheritdoc />
        public bool Mandatory => _definition.Mandatory;
        /// <inheritdoc />
        public bool Recurring => _definition.Recurring;
        /// <inheritdoc />
        public bool Combobox => _definition.Combobox;
        /// <inheritdoc />
        public int MaxChars => _definition.MaxChars;
        /// <inheritdoc />
        public string DefaultString => _definition.DefaultString;
        /// <inheritdoc />
        public string Name { get; set; }
    }
}

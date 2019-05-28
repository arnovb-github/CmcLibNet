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
        public bool Shared
        {
            get
            {
                return _definition.Shared;
            }
            set { }
        }

        /// <inheritdoc />
        public bool Mandatory
        {
            get
            {
                return _definition.Mandatory;
            }
            set { }
        }

        /// <inheritdoc />
        public bool Recurring
        {
            get
            {
                return _definition.Recurring;
            }
            set { }
        }

        /// <inheritdoc />
        public bool Combobox
        {
            get
            {
                return _definition.Combobox;
            }
            set { }
        }

        /// <inheritdoc />
        public int MaxChars
        {
            get
            {
                return _definition.MaxChars;
            }
            set { }
        }

        /// <inheritdoc />
        public string DefaultString
        {
            get
            {
                return _definition.DefaultString;
            }
            set { }
        }

        /// <inheritdoc />
        public string Name { get; set; }
    }
}

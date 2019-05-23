using System.Runtime.InteropServices;

namespace Vovin.CmcLibNet.Database.Metadata
{
    /// <summary>
    /// Holds information on the field definition.
    /// </summary>
    [ComVisible(true)]
    [Guid("1F0AEAE1-84CD-44e7-8504-08EE94493439")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(ICommenceFieldDefinition))]
    public class CommenceFieldDefinition : ICommenceFieldDefinition
    {
        internal CommenceFieldDefinition() { }
        /// <inheritdoc />
        public bool Shared { get; internal set; }
        /// <inheritdoc />
        public bool Mandatory { get; internal set; }
        /// <inheritdoc />
        public bool Recurring { get; internal set; }
        /// <inheritdoc />
        public bool Combobox { get; internal set; }
        /// <inheritdoc />
        public int MaxChars { get; internal set; }
        /// <inheritdoc />
        public string DefaultString { get; internal set; }
        /// <inheritdoc />
        public CommenceFieldType Type { get; internal set; } = CommenceFieldType.Text;
        ///// <inheritdoc />
        //public string TypeDescription { get; internal set; }
    }
}

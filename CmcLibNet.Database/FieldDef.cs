using System.Runtime.InteropServices;

namespace Vovin.CmcLibNet.Database
{
    /// <summary>
    /// Interface for the field definition.
    /// </summary>
    [ComVisible(true)]
    [Guid("D4434784-03C4-4eea-8660-BB03182B4D1C")]
    public interface IFieldDef
    {
        /// <summary>
        /// Field type.
        /// </summary>
        CommenceFieldType Type { get; }
        /// <summary>
        /// Field type description.
        /// </summary>
        string TypeDescription { get; }
        /// <summary>
        /// Indicates if field is shared.
        /// </summary>
        bool Shared { get; }
        /// <summary>
        /// Indicates if field is mandatory.
        /// </summary>
        bool Mandatory { get; }
        /// <summary>
        /// Indicates if field is a recurring date field.
        /// </summary>
        bool Recurring { get; }
        /// <summary>
        /// Indicates if field is a combobox.
        /// </summary>
        bool Combobox { get; }
        /// <summary>
        /// Maximum number of characters field can hold.
        /// </summary>
        int MaxChars { get; }
        /// <summary>
        /// Default fielvalue (if any).
        /// </summary>
        string DefaultString { get; }
    }

    /// <summary>
    /// Holds information on the field definition.
    /// </summary>
    [ComVisible(true)]
    [Guid("1F0AEAE1-84CD-44e7-8504-08EE94493439")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IFieldDef))]
    public class FieldDef : IFieldDef
    {
        internal FieldDef() { }
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
        public CommenceFieldType Type { get; internal set; }
        /// <inheritdoc />
        public string TypeDescription { get; internal set; }
    }
}

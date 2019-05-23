namespace Vovin.CmcLibNet.Database.Metadata
{
    /// <summary>
    /// Commence field metadata.
    /// </summary>
    public interface ICommenceFieldMetaData : ICommenceFieldDefinition
    {
        /// <summary>
        /// Field name.
        /// </summary>
        string Name { get; set; }
    }
}
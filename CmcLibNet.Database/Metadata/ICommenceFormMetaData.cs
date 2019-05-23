namespace Vovin.CmcLibNet.Database.Metadata
{
    /// <summary>
    /// Commence Item detail form metadata.
    /// </summary>
    public interface ICommenceFormMetaData
    {
        /// <summary>
        /// Categoryname.
        /// </summary>
        string Category { get; }
        /// <summary>
        /// Form name.
        /// </summary>
        string Name { get; }
        /// <summary>
        /// Form filepath.
        /// </summary>
        string Path { get; }
        /// <summary>
        /// Form script.
        /// </summary>
        string Script { get; }
        /// <summary>
        /// Form XML.
        /// </summary>
        string Xml { get; }
    }
}
namespace Vovin.CmcLibNet.Database
{
    /// <summary>
    /// Describes a Commence connection.
    /// </summary>
    public interface ICommenceConnection
    {
        /// <summary>
        /// Name of connection.
        /// </summary>
        string Name { get; set; }
        /// <summary>
        /// Connected category name.
        /// </summary>
        string ToCategory { get; set; }
    }
}

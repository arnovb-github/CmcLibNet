namespace Vovin.CmcLibNet.Database
{
    /// <summary>
    /// Describes a Commence connection.
    /// </summary>
    public class CommenceConnection : ICommenceConnection
    {
        /// <inheritdoc />
        public string Name { get; set; }
        /// <inheritdoc />
        public string ToCategory { get; set; }
    }
}

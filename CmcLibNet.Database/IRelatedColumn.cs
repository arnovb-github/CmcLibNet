namespace Vovin.CmcLibNet.Database
{
    /// <summary>
    /// Interface for RelatedColumn
    /// </summary>
    internal interface IRelatedColumn
    {
        /// <summary>
        /// Commence Connection name (case-sensitive!).
        /// </summary>
        string Connection { get; set; }
        /// <summary>
        /// Connected Commence category name.
        /// </summary>
        string Category { get; set; }
        /// <summary>
        /// Connected Commence field name.
        /// </summary>
        string Field { get; set; }
        /// <summary>
        /// Columntype.
        /// </summary>
        RelatedColumnType ColumnType { get; set; }
        /// <summary>
        /// Item delimiter. This is dependent on the way a related column is defined 
        /// AND dependent on the type of view if the cursor was created on a view. Sigh.
        /// </summary>
        string Delimiter { get; set; }
    }
}
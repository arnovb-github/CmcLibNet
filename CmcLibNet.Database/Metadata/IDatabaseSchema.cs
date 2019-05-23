using System.Collections.Generic;

namespace Vovin.CmcLibNet.Database.Metadata
{
    /// <summary>
    /// Schema information
    /// </summary>
    public interface IDatabaseSchema : IDBDef
    {
        /// <summary>
        /// Categories.
        /// </summary>
        List<CommenceCategoryMetaData> Categories { get; }
        /// <summary>
        /// Database size (combined size of database directory)
        /// </summary>
        long Size { get; }
    }
}
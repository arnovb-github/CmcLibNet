using System.Collections.Generic;

namespace Vovin.CmcLibNet.Database.Metadata
{
    /// <summary>
    /// Commence category metadata.
    /// </summary>
    public interface ICommenceCategoryMetaData : ICategoryDef
    {
        /// <summary>
        /// Category name.
        /// </summary>
        string Name { get; }
        /// <summary>
        /// Number of items.
        /// </summary>
        int Items { get; }
        /// <summary>
        /// Fields.
        /// </summary>
        List<CommenceFieldMetaData> Fields { get; }
        /// <summary>
        /// Views.
        /// </summary>
        List<CommenceViewMetaData> Views { get; }
        /// <summary>
        /// Forms.
        /// </summary>
        List<CommenceFormMetaData> Forms { get; }
        /// <summary>
        /// Connections.
        /// </summary>
        List<CommenceConnection> Connections { get; }
    }
}
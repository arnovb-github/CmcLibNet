using System.Collections.Generic;
using System.Data;
using Vovin.CmcLibNet.Database;

namespace Vovin.CmcLibNet.Export.Complex
{
    /// <summary>
    /// Used to define cursors in ComplexWriter
    /// </summary>
    internal class CursorDescriptor
    {
        // if we make this class serializable,
        // we could assign the serialized data to a custom property
        // in the dataset
        internal CursorDescriptor(string categoryOrView)
        {
            CategoryOrView = categoryOrView;
        }
        internal string CategoryOrView { get; }
        internal IList<string> Fields { get; set; } = new List<string>(); // all fields will be defined as direct fields
        internal IList<ICursorFilterTypeCTCF> Filters { get; set; } = new List<ICursorFilterTypeCTCF>(); // the filtertype is very specific
        internal int MaxFieldSize { get; set; } = CommenceLimits.DefaultMaxFieldSize;
        internal CmcCursorType CursorType { get; set; } = CmcCursorType.Category;
        /// <summary>
        /// Maps cursor columns on sql fieldnames.
        /// the int represents the Commence columnindex
        /// </summary>
        internal IDictionary<int, SqlMap> SqlColumnMappings { get; set; } = new Dictionary<int, SqlMap>();

        /// <summary>
        /// Creates a mapping between the Commence column in a cursor and the corresponding
        /// table and field in the DataSet
        /// </summary>
        /// <param name="dt"></param>
        internal void CreateSqlMapping(DataTable dt)
        {
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                SqlColumnMappings.Add(i, new SqlMap(dt.TableName, dt.Columns[i].ColumnName, !dt.Columns[i].AllowDBNull));
            }
        }
        internal bool IsTableWithConnectedThids { get; set; } // link tables are processed differently
    }
}

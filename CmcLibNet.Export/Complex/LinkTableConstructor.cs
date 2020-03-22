using System.Text;

namespace Vovin.CmcLibNet.Export.Complex
{
    // class to assist in creating consistent columnnames for SQLite link table
    // this just captures an internal convention in this assembly
    internal static class LinkTableConstructor
    {
        private static StringBuilder sb = new StringBuilder();

        /// <summary>
        /// Link table name.
        /// </summary>
        /// <param name="primaryTableName">Name of the parent Commence category.</param>
        /// <param name="connectionName">Name of the Commence connection from the parent to the connected category.</param>
        /// <param name="toCategory">Name of the connected Commence category.</param>
        /// <returns>Link table name.</returns>
        internal static string TableName(string primaryTableName, string connectionName, string toCategory)
        {
            sb.Clear();
            sb.Append(primaryTableName);
            sb.Append(connectionName);
            sb.Append(toCategory);
            return sb.ToString();
        }

        /// <summary>
        /// Foreign key of parent table.
        /// </summary>
        /// <param name="connectionName">Name of connection from parent Commence category to connected category.</param>
        /// <param name="toCategory">Name of connected Commence category.</param>
        /// <param name="postFix">A postfix.</param>
        /// <returns>Columnname for foreign key of connected table.</returns>
        internal static string ForeignKeyOfConnectedTable(string connectionName, string toCategory, string postFix)
        {
            sb.Clear();
            sb.Append(connectionName);
            sb.Append(toCategory);
            sb.Append(postFix);
            return sb.ToString();
        }

        /// <summary>
        /// Foreign key of connected table.
        /// </summary>
        /// <param name="primaryTableName">Name of connected Commence category.</param>
        /// <param name="postFix">A postfix.</param>
        /// <returns>Columnname for foreign key of parent table.</returns>
        internal static string ForeignKeyOfPrimaryTable(string primaryTableName, string postFix)
        {
            sb.Clear();
            sb.Append(primaryTableName);
            sb.Append(postFix);
            return sb.ToString();
        }
        
    }
}

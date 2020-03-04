using System.Collections.Generic;

namespace Vovin.CmcLibNet.Export
{
    /// <summary>
    /// TableDef is a rudimentary table description, mimicking Commence tables.
    /// It is used to construct a true DataSet, or to more efficiently do DDE requests.
    /// </summary>
    internal class TableDef
    {
        internal TableDef(string name, string category, bool primary = false)
        {
            ColumnDefinitions = new List<ColumnDefinition>();
            Name = name;
            Category = category;
            Primary = primary;
        }
        internal bool Primary { get; set; }
        internal string Name { get; set; }
        internal string Category { get; set; }
        internal List<ColumnDefinition> ColumnDefinitions { get; } = null;
    }
}

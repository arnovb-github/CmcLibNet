using System.Collections.Generic;

namespace Vovin.CmcLibNet.Export
{
    /// <summary>
    /// TableDef is a rudimentary table description, mimicking Commence tables.
    /// It is used to construct a true DataSet, or to more efficiently do DDE requests.
    /// </summary>
    internal class TableDef
    {
        List<ColumnDefinition> _columnDefinitions = null;

        internal TableDef(string name, string category, bool primary = false)
        {
            _columnDefinitions = new List<ColumnDefinition>();
            this.Name = name;
            this.Category = category;
            this.Primary = primary;
        }
        internal bool Primary { get; set; }
        internal string Name { get; set; }
        internal string Category { get; set; } 
        internal List<ColumnDefinition> ColumnDefinitions
        {
            get
            {
                return _columnDefinitions;
            }
        }
    }
}

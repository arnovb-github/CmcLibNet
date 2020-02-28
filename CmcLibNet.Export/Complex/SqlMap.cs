namespace Vovin.CmcLibNet.Export.Complex
{
    internal struct SqlMap
    {
        internal SqlMap(string tableName, string columnName, bool isRequired)
        {
            TableName = tableName;
            ColumnName = columnName;
            IsRequired = isRequired;
        }
        internal string TableName { get; }
        internal string ColumnName { get; }
        internal bool IsRequired { get; } // for checking NOT NULL
    }
}
using Vovin.CmcLibNet.Database.Metadata;

namespace Vovin.CmcLibNet.Export.Complex
{
    internal class ChildTableQuery
    {
        public ChildTableQuery(string commandText, CommenceConnection commenceConnection)
        {
            CommandText = commandText;
            Connection = commenceConnection;
        }
        public string CommandText { get; }
        public CommenceConnection Connection { get; }
    }
}

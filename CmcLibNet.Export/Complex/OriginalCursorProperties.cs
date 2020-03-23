using Vovin.CmcLibNet.Database;

namespace Vovin.CmcLibNet.Export.Complex
{
    // used to pass some metadata on the original export cursor around
    internal class OriginalCursorProperties
    {
        public OriginalCursorProperties(ICommenceCursor cursor)
        {
            Name = string.IsNullOrEmpty(cursor.View)
                ? cursor.Category
                : cursor.View;
            Category = cursor.Category;
            Type = string.IsNullOrEmpty(cursor.View)
                ? CmcCursorType.Category.ToString()
                : CmcCursorType.View.ToString();
        }
        public string Name { get; set; } // van be View name or Category name
        public string Category { get; set; } // underlying category
        public string Type { get; set; }
    }
}

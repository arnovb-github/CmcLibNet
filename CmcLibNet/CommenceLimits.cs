namespace Vovin.CmcLibNet
{
    /// <summary>
    /// Internal Commence limits
    /// </summary>
    internal static class CommenceLimits
    {
        /// <summary>
        /// Maximum number of items per category.
        /// </summary>
        internal static int MaxItems => 500000; // different from what Commence says (750.000)! Up to at least RM 7.1, this is the limit.
        /// <summary>
        /// Maximum number of fields per category, including connections.
        /// </summary>
        internal static int MaxFieldsPerCategory => 250;
        /// <summary>
        /// Maximum length of a THID.
        /// </summary>
        internal static int ThidLength => 21;
        /// <summary>
        /// Maximum number of characters a large text field can contain.
        /// </summary>
        internal static int MaxTextFieldCapacity => 30000;
        /// <summary>
        /// Maximum number of characters a name field can contain.
        /// </summary>
        internal static int MaxNameFieldCapacity => 50;
        /// <summary>
        /// Maximum length of a connection name.
        /// </summary>
        internal static int MaxConnectionNameLength => 16;
        /// <summary>
        /// Maximum length of a field name.
        /// </summary>
        internal static int MaxFieldNameLength = 20;
        /// <summary>
        /// Maximum number of fields that a view can display, including connections.
        /// </summary>
        internal static int MaxFieldsPerView => 255;
        /// <summary>
        /// Maximum number of filters that can be defined on a cursor
        /// </summary>
        internal static int MaxFilters => 8;
        /// <summary>
        /// Default maximum number of characters returned for connections in a cursor
        /// </summary>
        internal static int DefaultMaxFieldSize => 93750;
    }
}

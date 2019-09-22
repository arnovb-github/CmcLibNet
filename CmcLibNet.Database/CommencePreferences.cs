namespace Vovin.CmcLibNet.Database
{
    /// <summary>
    /// Class exposing some of Commence's preferences settings
    /// </summary>
    internal class Preferences
    {

        /// <summary>
        /// Gets (-Me-) item.
        /// </summary>
        /// <remarks>Does NOT set this preference in Commence.</remarks>
        internal string Me { get; set; } = "(-Me-) item not set.";

        /// <summary>
        /// Gets/Sets (-Me-) category.
        /// </summary>
        /// <remarks>Does NOT set this preference in Commence.</remarks>
        internal string MeCategory { get; set; } = "(-Me-) category not set.";

        /// <summary>
        /// Gets/Sets Letter Log path
        /// </summary>
        /// <remarks>Does NOT set this preference in Commence.</remarks>
        internal string LetterLogDir { get; set; } = "Letter Log Directory not set.";

        /// <summary>
        /// Gets/Sets External Data files path.
        /// </summary>
        /// <remarks>Does NOT set this preference in Commence.</remarks>
        internal string ExternalDir { get; set; } = "Spool directory not set or not applicable.";
    }
}

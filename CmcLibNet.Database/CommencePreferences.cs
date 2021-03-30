namespace Vovin.CmcLibNet.Database
{
    /// <summary>
    /// Class exposing some of Commence's preferences settings
    /// </summary>
    public class Preferences
    {
        /// <summary>
        /// Gets (-Me-) item.
        /// </summary>
        /// <remarks>Does NOT set this preference in Commence.</remarks>
        public string Me { get; internal set; } = "(-Me-) item not set.";

        /// <summary>
        /// Gets/Sets (-Me-) category.
        /// </summary>
        /// <remarks>Does NOT set this preference in Commence.</remarks>
        public string MeCategory { get; internal set; } = "(-Me-) category not set.";

        /// <summary>
        /// Gets/Sets Letter Log path
        /// </summary>
        /// <remarks>Does NOT set this preference in Commence.</remarks>
        public string LetterLogDir { get; internal set; } = "Letter Log Directory not set.";

        /// <summary>
        /// Gets/Sets External Data files path.
        /// </summary>
        /// <remarks>Does NOT set this preference in Commence.</remarks>
        public string ExternalDir { get; internal set; } = "Spool directory not set or not applicable.";
    }
}
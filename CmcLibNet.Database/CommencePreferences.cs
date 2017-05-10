namespace Vovin.CmcLibNet.Database
{
    /// <summary>
    /// Class exposing some of Commence's preferences settings
    /// </summary>
    internal class Preferences
    {
        string _me = "(-Me-) item not set.";
        string _mecat = "(-Me-) category not set.";
        string _letterlogdir = "Letter Log Directory not set.";
        string _externaldir = "Spool directory not set or not applicable.";

        /// <summary>
        /// Gets (-Me-) item.
        /// </summary>
        /// <remarks>Does NOT set this preference in Commence.</remarks>
        internal string Me
        {
            get
            {
                return _me;
            }
            set
            {
                _me = value;
            }

        }

        /// <summary>
        /// Gets/Sets (-Me-) category.
        /// </summary>
        /// <remarks>Does NOT set this preference in Commence.</remarks>
        internal string MeCategory
        {
            get
            {
                return _mecat;
            }
            set
            {
                _mecat = value;
            }
        }

        /// <summary>
        /// Gets/Sets Letter Log path
        /// </summary>
        /// <remarks>Does NOT set this preference in Commence.</remarks>
        internal string LetterLogDir
        {
            get
            {
                return _letterlogdir;
            }
            set
            {
                _letterlogdir = value;
            }
        }
        /// <summary>
        /// Gets/Sets External Data files path.
        /// </summary>
        /// <remarks>Does NOT set this preference in Commence.</remarks>
        internal string ExternalDir
        {
            get
            {
                return _externaldir;
            }
            set
            {
                _externaldir = value;
            }
        }
    }
}

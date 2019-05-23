using System.Runtime.InteropServices;

namespace Vovin.CmcLibNet.Database.Metadata
{
    /// <summary>
    /// Exposes members of the database definition.
    /// </summary>
    [ComVisible(true)]
    [Guid("574A9962-86A3-4489-8001-67AC08464EEB")]
    public interface IDBDef
    {
        /// <summary>
        /// Database name.
        /// </summary>
        string Name { get;  }
        /// <summary>
        /// Database path.
        /// </summary>
        string Path { get;  }
        /// <summary>
        /// Indicates if database is attached, whatever that may mean.
        /// </summary>
        bool Attached { get;  }
        /// <summary>
        /// Indicates if field is connected, whatever that may mean.
        /// </summary>
        bool Connected { get;  }
        /// <summary>
        /// Indicates if database is a server in a workgroup.
        /// </summary>
        bool IsServer { get; }
        /// <summary>
        /// Indicates if database is a client in a workgroup.
        /// </summary>
        bool IsClient { get;  }
        /// <summary>
        /// Login name of the user (only applies to password-protected databases).
        /// </summary>
        string Username { get;  }
        /// <summary>
        /// Path where synchronization packets exchanged. WARNING: Commence does not properly return this value if client is set to sync by any means other than Shared LAN. If it syncs via FTP, you will still erronously get a Shared LAN location if one was ever set.
        /// </summary>
        string Spoolpath { get;  }
    }
}

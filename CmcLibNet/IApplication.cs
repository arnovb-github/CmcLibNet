using System.Runtime.InteropServices;

namespace Vovin.CmcLibNet
{
    /// <summary>
    /// Application interface.
    /// </summary>
    [ComVisible(true)]
    [Guid("74AB03F3-FF88-447A-8EB6-6CB89240D69C")]
    public interface IApplication
    {
        /// <summary>
        /// Returns assembly version.
        /// </summary>
        string Version { get; }
        /// <summary>
        /// Returns information on the assembly. .Net only.
        /// </summary>
        [ComVisible(false)]
        AssemblyInfo AssemblyInfo { get; }
        /// <summary>
        /// Returns a reference to the Commence application
        /// </summary>
        ICommenceApp CommenceApp { get; }
        /// <summary>
        /// Returns a reference to the Commence database.
        /// </summary>
        Database.ICommenceDatabase Database { get; }
        /// <summary>
        /// Returns a reference to the Export engine.
        /// </summary>
        Export.IExportEngine Export { get; }
        /// <summary>
        /// Returns a reference to the Services utility.
        /// </summary>
        Services.IServices Services { get; }
    }
}

using System.Reflection;
using System.Runtime.InteropServices;
using Vovin.CmcLibNet.Database;
using Vovin.CmcLibNet.Export;
using Vovin.CmcLibNet.Services;

namespace Vovin.CmcLibNet
{
    /// <summary>
    /// References the assembly itself.
    /// </summary>
    [ComVisible(true)]
    [Guid("D49ABB93-8B54-4DDB-AD0E-C531D993C414")]
    [ProgId("CmcLibNet.Application")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IApplication))]
    public class Application : IApplication
    {
        /// <inheritdoc />
        public string Version => Assembly.GetExecutingAssembly().GetName().Version.ToString();
        /// <inheritdoc />
        [ComVisible(false)]
        public AssemblyInfo AssemblyInfo => new AssemblyInfo();
        /// <inheritdoc />
        public ICommenceApp CommenceApp => new CommenceApp();
        /// <inheritdoc />
        public ICommenceDatabase Database => new CommenceDatabase();
        /// <inheritdoc />
        public IExportEngine Export => new ExportEngine();
        /// <inheritdoc />
        public IServices Services => new Services.Services();
    }
}

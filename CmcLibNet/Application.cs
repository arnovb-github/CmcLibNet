using System.Reflection;
using System.Runtime.InteropServices;
using Vovin.CmcLibNet.Database;
using Vovin.CmcLibNet.Export;
using Vovin.CmcLibNet.Services;

namespace Vovin.CmcLibNet
{
    /// <summary>
    /// CmcLibNet is a .NET assembly that wraps the Commence API. 
    /// The primary goal for developing this library was to make it easier to communicate with Commence from Powershell.
    /// Some convenience methods can only be used from .NET applications, but all functionality as defined in the Commence API is available from COM.
    /// </summary>
    /// <remarks>
    /// <para>
    /// COM clients can use it by calling it with ProgId <c>'CmcLibNet.Application'</c>.
    /// </para>
    /// <para>COM applications such as Commence Item Detail Form scripts that use so-called 'late binding' can call the assembly thus:</para>
    /// <para>VBscript:</para>
    /// <code language="vbscript">Dim obj : Set obj = CreateObject("CmcLibNet.Application")</code>
    /// </remarks>
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

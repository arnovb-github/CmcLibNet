using System.Runtime.InteropServices;

namespace Vovin.CmcLibNet
{
    /// <summary>
    /// Interface for CommenceApp.
    /// </summary>
    [ComVisible(true)]
    [Guid("3C377558-ADDF-47a6-A7E3-0CB4735534DD")]
    public interface ICommenceApp
    {
        /// <summary>
        /// Path to commence.exe.
        /// </summary>
        string ExePath { get; }
        /// <summary>
        /// Name of currently open Commence database.
        /// </summary>
        string Name { get; }
        /// <summary>
        /// Path of currently open Commence database.
        /// </summary>
        string Path { get; }
        /// <summary>
        /// Registered user of current Commence database.
        /// </summary>
        string RegisteredUser { get; }
        /// <summary>
        /// Version of active Commence application.
        /// </summary>
        string Version { get; }
        /// <summary>
        /// Extended version of active Commence application.
        /// </summary>
        string VersionExt { get; }
        /// <summary>
        /// Close any references to Commence. Objects implementing this interface should be disposed after this. Important: object derived from this interface will also have their COM-references closed.
        /// </summary>
        /// <remarks>When used from within a Commence Form Script, failing to call the <c>Close</c> method will leave the commence.exe process running in the background when the user closes Commence. IMPORTANT: this also happens when an unhandled exception (a 'script error') occurs. The Commence process then has to be closed manually from the Windows Task Manager. Be careful to implement proper error handling.
        /// <para>When the assembly is called from a.NET application, there is rarely a need to call this method, unless you want to explicitly release COM references and/or release memory. It can be useful in some cases, because Commence may complain about running out of memory before the Garbage Collector has a chance to kick in.</para>
        /// <para>Technical details: calling this method tells the assembly to release all COM handles (called 'RCW' for 'runtime callable wrapper') to Commence that are open. This is needed because when the object reference to this assembly is set to Nothing (in VB), the .NET assembly may not be notified and will think they are still in use. Garbage Collection will therefore not release them, and the commence.exe process will not be terminated.</para>
        /// </remarks>
        void Close();
    }
}
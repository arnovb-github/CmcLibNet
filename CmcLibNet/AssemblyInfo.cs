using System.Diagnostics;
using System.Reflection;

namespace Vovin.CmcLibNet
{
    /// <summary>
    /// Contains general assembly information.
    /// </summary>
    public class AssemblyInfo
    {
        /// <summary>
        /// Returns assembly version.
        /// </summary>
        public string Version => Assembly.GetExecutingAssembly().GetName().Version.ToString();
        /// <summary>
        /// Returns targeted framework version.
        /// </summary>
        public string TargetedFramework => GetTargetFrameworkVersion();
        /// <summary>
        /// Returns CLR version.
        /// </summary>
        public string ClrVersion => System.Environment.Version.ToString();
        /// <summary>
        /// Returns the ImageRuntimeVersion.
        /// </summary>
        public string ImageRuntimeVersion => Assembly.GetExecutingAssembly().ImageRuntimeVersion;
        /// <summary>
        /// Returns the file version.
        /// </summary>
        public string FileVersion => FileVersionInfo.GetVersionInfo(Assembly.GetExecutingAssembly().Location).FileVersion;
        /// <summary>
        /// Returns the product version.
        /// </summary>
        public string ProductVersion => FileVersionInfo.GetVersionInfo(Assembly.GetExecutingAssembly().Location).ProductVersion;
        /// <summary>
        /// Returns the location of the assembly
        /// </summary>
        public string Path => Assembly.GetExecutingAssembly().Location;

        private static string GetTargetFrameworkVersion()
        {
            AssemblyName[] references = Assembly.GetExecutingAssembly().GetReferencedAssemblies();
            foreach (AssemblyName reference in references)
            {
                if (reference.Name == "mscorlib")
                {
                    return reference.Version.ToString();
                }
            }
            return string.Empty;
        }
    }
}

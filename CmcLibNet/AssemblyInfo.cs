using System.Reflection;

namespace Vovin.CmcLibNet
{
    /// <summary>
    /// Contains general assembly information.
    /// </summary>
    public class AssemblyInfo
    {
        /// <summary>
        /// Returns version.
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

using System;

namespace Vovin.CmcLibNet.Database
{
    /// <summary>
    /// Used to indicate the number of values that a FilterQualifier relies on.
    /// </summary>
    /// <remarks>This information can be used when validating an <see cref="ICursorFilter"/></remarks>
    /// <example>A <see cref="FilterQualifier"/> <code>Checked</code> filter relies on 0 values.
    /// <para>A <see cref="FilterQualifier"/> <code>Contains</code> filter relies on 1 value.</para>
    /// <para>A <see cref="FilterQualifier"/> <code>Between</code> filter relies on 2 values.</para>
    /// </example>
    [AttributeUsage(AttributeTargets.All)]
    public class FilterValuesAttribute : Attribute
    {
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="n">Number of dependencies</param>
        public FilterValuesAttribute(int n)
        {
            Number = n;
        }
        /// <summary>
        /// Number of dependencies
        /// </summary>
        public int Number { get; private set; }
    }
}

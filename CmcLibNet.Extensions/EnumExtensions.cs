using System;
using System.ComponentModel;
using System.Reflection;

namespace Vovin.CmcLibNet.Extensions
{
    /// <summary>
    /// Enum extension methods
    /// </summary>
    public static class EnumExtensions
    {
        /// <summary>
        /// Extension method for enums that returns description attributes for enum value.
        /// </summary>
        /// <param name="value">Enum value.</param>
        /// <returns>String representation of enum member, or its description (if defined).</returns>
        public static string GetEnumDescription(this Enum value)
        {
            FieldInfo fi = value.GetType().GetField(value.ToString());

            DescriptionAttribute[] attributes =
                (DescriptionAttribute[])fi.GetCustomAttributes(typeof(DescriptionAttribute), false);

            if (attributes != null && attributes.Length > 0)
                return attributes[0].Description;
            else
                return value.ToString();
        }
    }
}

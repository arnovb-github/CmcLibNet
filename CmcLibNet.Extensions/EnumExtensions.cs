using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Reflection;
using Vovin.CmcLibNet.Attributes;

namespace Vovin.CmcLibNet.Extensions
{
    /// <summary>
    /// Enum extension methods
    /// </summary>
    internal static class EnumExtensions
    {
        /// <summary>
        /// Extension method for enums that returns description attributes for enum value.
        /// </summary>
        /// <param name="value">Enum value.</param>
        /// <returns>String representation of enum member, or its description (if defined).</returns>
        internal static string GetEnumDescription(this Enum value)
        {
            // TODO get rid of this extension method and use StringAttribute consistently

            FieldInfo fi = value.GetType().GetField(value.ToString());

            DescriptionAttribute[] attributes =
                (DescriptionAttribute[])fi.GetCustomAttributes(typeof(DescriptionAttribute), false);

            if (attributes != null && attributes.Length > 0)
                return attributes[0].Description;
            else
                return value.ToString();
        }

        /// <summary>
        /// Will get the string value for a given enums value, this will
        /// only work if you assign the StringValue attribute to
        /// the items in your enum.
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        internal static string GetStringValue(this Enum value)
        {
            // Get the type
            Type type = value.GetType();

            // Get fieldinfo for this type
            FieldInfo fieldInfo = type.GetField(value.ToString());

            // Get the stringvalue attributes
            StringValueAttribute[] attribs = fieldInfo.GetCustomAttributes(
                typeof(StringValueAttribute), false) as StringValueAttribute[];

            // Return the first if there was a match.
            return attribs.Length > 0 ? attribs[0].StringValue : null;
        }

        /// <summary>
        /// Returns flags from an enum.
        /// </summary>
        /// <param name="e"></param>
        /// <returns></returns>
        public static IEnumerable<Enum> GetFlags(this Enum e)
        {
            return Enum.GetValues(e.GetType()).Cast<Enum>().Where(e.HasFlag);
        }

        /// <summary>
        /// Returns elements from end of an array
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="collection"></param>
        /// <param name="n"></param>
        /// <returns></returns>
        public static IEnumerable<T> TakeLast<T>(this IEnumerable<T> collection, int n)
        {
            if (collection == null)
                throw new ArgumentNullException(nameof(collection));
            if (n < 0)
                throw new ArgumentOutOfRangeException(nameof(n), $"{nameof(n)} must be 0 or greater");

            LinkedList<T> temp = new LinkedList<T>();

            foreach (var value in collection)
            {
                temp.AddLast(value);
                if (temp.Count > n)
                    temp.RemoveFirst();
            }

            return temp;
        }
    }
}
